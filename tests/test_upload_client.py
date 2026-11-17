"""Tests for courier/quickxor.py and the upload client in courier/send.py: chunked PUT against a local fake upload
server (a 120 MB sparse file, one injected 503 resumed through nextExpectedRanges, Content-Range, no Authorization),
and the whole `send.py list` / `send.py upload` flow against a local git remote with a stand-in courier.
Run: python3 -m unittest discover tests"""
import base64
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import unittest.mock
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(HERE / "courier"))
os.environ.setdefault("TF_UPLOAD_RETRY_SLEEP", "0.01")
import send  # noqa: E402
from box import new_keypair, open_box, seal  # noqa: E402
from quickxor import QuickXor, quickxor_bytes, quickxor_file  # noqa: E402

send.RETRY_SLEEP = 0.01


def b64_20(**at):
    """20 state bytes with the given byte positions set -> base64"""
    raw = bytearray(20)
    for k, v in at.items():
        raw[int(k[1:])] = v
    return base64.b64encode(bytes(raw)).decode()


def slow_quickxor(data: bytes) -> str:
    """The definition, one byte at a time: byte i XORed in at bit 11*i (mod 160, wrapping), then the length."""
    state = 0
    for i, b in enumerate(data):
        sh = b << ((11 * i) % 160)
        state ^= (sh | (sh >> 160)) & ((1 << 160) - 1)
    state ^= len(data) << 96
    return base64.b64encode(state.to_bytes(20, "little")).decode()


class QuickXorTests(unittest.TestCase):
    def test_known_vectors(self):
        self.assertEqual(quickxor_bytes(b""), "AAAAAAAAAAAAAAAAAAAAAAAAAAA=")
        # one byte 0x01: bit 0 set; length 1 lands in byte 12
        self.assertEqual(quickxor_bytes(b"\x01"), b64_20(b0=1, b12=1))
        self.assertEqual(quickxor_bytes(b"a"), b64_20(b0=0x61, b12=1))
        # second byte goes in at bit 11 = byte 1, bit 3; length 2
        self.assertEqual(quickxor_bytes(b"\x01\x01"), b64_20(b0=1, b1=8, b12=2))
        # 0xFF at bit 0 of an empty-ish stream, then 0xFF at bit 11: bits 0-7 and 11-18
        self.assertEqual(quickxor_bytes(b"\xff\xff"), b64_20(b0=0xFF, b1=0xF8, b2=0x07, b12=2))
        # the 15th byte sits at bit 154 and wraps around the 160-bit state
        wrap = quickxor_bytes(b"\x00" * 14 + b"\xff")
        raw = base64.b64decode(wrap)
        self.assertEqual((raw[19], raw[0], raw[12]), (0xFC, 0x03, 15))

    def test_matches_the_definition_across_row_boundaries_and_chunkings(self):
        for n in (0, 1, 2, 159, 160, 161, 320, 321, 479, 1000, 4097):
            data = os.urandom(n)
            want = slow_quickxor(data)
            self.assertEqual(quickxor_bytes(data), want, n)
            h, i = QuickXor(), 0
            while i < n:                                  # odd-sized pieces
                h.update(data[i:i + 7 + i % 301])
                i += 7 + i % 301
            self.assertEqual(h.b64(), want, n)

    def test_file(self):
        p = Path(tempfile.mkdtemp()) / "f.bin"
        data = os.urandom(50_000)
        p.write_bytes(data)
        self.assertEqual(quickxor_file(p, chunk=4096), (quickxor_bytes(data), 50_000))


class Upload:
    """Fake pre-authenticated upload endpoint: one session, strict contiguous ranges."""
    received = 0
    sha = None
    puts = []            # (Content-Range, has Authorization, body length)
    gets = 0
    fail_at = None       # offset whose chunk is stored but answered 503, once
    total = 0
    refuse = False       # answer every PUT with 400
    expired = False      # the session is gone: GET says 404
    deletes = []         # (has Authorization) for each DELETE (= session cancelled)


class UH(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_DELETE(self):
        Upload.deletes.append("Authorization" in self.headers)
        self.send_response(204)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        Upload.gets += 1
        if Upload.expired:
            return self._json(404, {"error": "gone"})
        self._json(200, {"nextExpectedRanges": [f"{Upload.received}-"]})

    def do_PUT(self):
        n = int(self.headers["Content-Length"])
        rng = self.headers["Content-Range"]
        Upload.puts.append((rng, "Authorization" in self.headers, n))
        if Upload.refuse:
            self.rfile.read(n)
            return self._json(400, {"error": "refused"})
        a, rest = rng.split(" ")[1].split("-")
        b, total = rest.split("/")
        a, b, total = int(a), int(b), int(total)
        left, body = n, []
        while left:
            piece = self.rfile.read(min(left, 1 << 20))
            body.append(piece)
            left -= len(piece)
        if a != Upload.received or b - a + 1 != n or total != Upload.total:
            return self._json(416, {"error": "range"})
        if b < total - 1 and (b - a + 1) % (320 * 1024):
            return self._json(400, {"error": "chunk not a multiple of 320 KiB"})
        for piece in body:
            Upload.sha.update(piece)
        Upload.received = b + 1
        if Upload.fail_at == a:
            Upload.fail_at = None
            return self._json(503, {"error": "busy"})
        if Upload.received == total:
            return self._json(201, {"id": "ITEM-1", "size": total})
        self._json(202, {"nextExpectedRanges": [f"{Upload.received}-"]})


class ChunkedPutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.srv = ThreadingHTTPServer(("127.0.0.1", 0), UH)
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.url = f"http://127.0.0.1:{cls.srv.server_port}/upload/secret-session-id"

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()

    def reset(self, total, fail_at=None):
        Upload.received, Upload.sha, Upload.puts, Upload.gets = 0, hashlib.sha256(), [], 0
        Upload.fail_at, Upload.total = fail_at, total
        Upload.refuse, Upload.expired, Upload.deletes = False, False, []

    def test_cancel_session_sends_a_delete_without_a_token_and_survives_a_dead_server(self):
        self.reset(1)
        self.assertTrue(send.cancel_session(self.url))
        self.assertEqual(Upload.deletes, [False])
        self.assertFalse(send.cancel_session("http://127.0.0.1:9/upload/secret-session-id"))

    def test_put_all_cancels_the_session_of_a_file_that_fails_and_of_every_unfinished_one_on_interrupt(self):
        keys = Path(tempfile.mkdtemp())
        files = []
        for i in range(3):
            q = keys / f"f{i}.pdf"
            q.write_bytes(b"x" * 10)
            files.append({"index": i, "local": str(q), "name": q.name, "bytes": 10, "item_id": "", "error": ""})
        sessions = {i: {"upload_url": self.url} for i in range(3)}
        with unittest.mock.patch.object(send, "FETCH_KEYS", keys):
            self.reset(10)
            Upload.refuse = True                                  # file 0 is refused -> its session is cancelled
            state = {"files": files}
            with unittest.mock.patch.object(send, "put_file", side_effect=[RuntimeError("SharePoint refused a chunk (400)"),
                                                                          KeyboardInterrupt]):
                with self.assertRaises(KeyboardInterrupt):        # SIGTERM / Ctrl-C on file 1: files 1 and 2 cancelled
                    send._put_all("c1", state, sessions, 327680)
            self.assertEqual(len(Upload.deletes), 3)
            self.assertIn("refused a chunk", files[0]["error"])
            self.assertIn("interrupted", files[1]["error"])
            self.assertIn("interrupted", files[2]["error"])
            self.assertTrue((keys / "up-c1.state.json").exists())   # the state survives for upload-resume
            self.assertNotIn("secret-session-id", (keys / "up-c1.state.json").read_text())

    def test_120_mb_sparse_file_with_a_503_resumed_from_next_expected_ranges(self):
        total, chunk = 120 * 1024 * 1024 + 12345, 10 * 1024 * 1024
        p = Path(tempfile.mkdtemp()) / "big.pdf"
        with p.open("wb") as f:                    # sparse: holes read as zeros; marks make misordering visible
            for off in (0, chunk - 1, chunk, 57 * chunk + 5, total - 1):
                f.seek(off)
                f.write(b"\x07")
            f.truncate(total)
        want_sha, want_qx, size = send.hash_file(p)
        self.assertEqual(size, total)
        self.reset(total, fail_at=3 * chunk)
        item = send.put_file(self.url, p, total, chunk)
        self.assertEqual(item["id"], "ITEM-1")
        self.assertEqual((Upload.received, Upload.sha.hexdigest()), (total, want_sha))   # every byte, in order
        self.assertGreaterEqual(Upload.gets, 1)                       # resumed by asking the session
        ranges = [r for r, _, _ in Upload.puts]
        self.assertEqual(ranges[0], f"bytes 0-{chunk - 1}/{total}")
        self.assertEqual(ranges[-1], f"bytes {12 * chunk}-{total - 1}/{total}")
        self.assertEqual(len(ranges), 13)                             # the 503'd chunk was NOT sent twice
        self.assertFalse(any(auth for _, auth, _ in Upload.puts))     # pre-authenticated URL: no token
        self.assertEqual(sum(n for _, _, n in Upload.puts), total)

    def test_small_file_is_one_chunk_and_a_resent_chunk_resumes_in_place(self):
        p = Path(tempfile.mkdtemp()) / "s.zip"
        p.write_bytes(os.urandom(1000))
        self.reset(1000)
        self.assertEqual(send.put_file(self.url, p, 1000, 10 * 1024 * 1024)["id"], "ITEM-1")
        self.assertEqual(Upload.puts[0][0], "bytes 0-999/1000")
        q = Path(tempfile.mkdtemp()) / "m.zip"
        q.write_bytes(os.urandom(3 * 327680 + 5))
        self.reset(3 * 327680 + 5, fail_at=0)        # first chunk stored, reply 503
        send.put_file(self.url, q, 3 * 327680 + 5, 327680)
        self.assertEqual(Upload.sha.hexdigest(), hashlib.sha256(q.read_bytes()).hexdigest())

    def test_gives_up_after_five_retries_without_leaking_the_url(self):
        p = Path(tempfile.mkdtemp()) / "x.pdf"
        p.write_bytes(b"x" * 10)
        old = send.RETRY_SLEEP
        dead = "http://127.0.0.1:9/upload/secret-session-id"   # nothing listens
        with self.assertRaises(RuntimeError) as cm:
            send.put_file(dead, p, 10, 327680)
        self.assertNotIn("secret-session-id", str(cm.exception))
        self.assertIn("5 retries", str(cm.exception))
        send.RETRY_SLEEP = old

    def test_409_is_a_name_conflict(self):
        class C(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def do_PUT(self):
                self.rfile.read(int(self.headers["Content-Length"]))
                self.send_response(409)
                self.send_header("Content-Length", "0")
                self.end_headers()
        srv = ThreadingHTTPServer(("127.0.0.1", 0), C)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        p = Path(tempfile.mkdtemp()) / "x.pdf"
        p.write_bytes(b"x")
        with self.assertRaises(RuntimeError) as cm:
            send.put_file(f"http://127.0.0.1:{srv.server_port}/u", p, 1, 327680)
        self.assertIn("name conflict", str(cm.exception))
        srv.shutdown()
        srv.server_close()


def run(*cmd, cwd=None, env=None):
    return subprocess.run(cmd, cwd=cwd, env=env, capture_output=True, text=True)


class FlowTests(unittest.TestCase):
    """send.py list / upload / upload-wait through a local bare git remote, with a thread playing the courier."""
    BRANCH = "scout-upload-session"

    def setUp(self):
        tmp = Path(tempfile.mkdtemp())
        self.remote, self.box, self.courier = tmp / "remote.git", tmp / "box", tmp / "courier"
        run("git", "init", "-q", "--bare", str(self.remote))
        run("git", "clone", "-q", str(self.remote), str(self.box))
        self.cpriv, self.cpub = new_keypair()
        for clone in (self.box,):
            run("git", "-C", str(clone), "checkout", "-q", "-b", self.BRANCH)
            run("git", "-C", str(clone), "config", "user.email", "t@example.invalid")
            run("git", "-C", str(clone), "config", "user.name", "t")
        (self.box / "courier.pub").write_text(self.cpub)
        run("git", "-C", str(self.box), "add", "courier.pub")
        run("git", "-C", str(self.box), "commit", "-q", "-m", "init")
        run("git", "-C", str(self.box), "push", "-q", "origin", f"HEAD:{self.BRANCH}")
        run("git", "clone", "-q", "-b", self.BRANCH, str(self.remote), str(self.courier))
        run("git", "-C", str(self.courier), "config", "user.email", "c@example.invalid")
        run("git", "-C", str(self.courier), "config", "user.name", "c")
        self.keys = tmp / "keys"
        self.env = {**os.environ, "TF_COURIER_KEYS": str(self.keys), "TF_COURIER_POLL": "0.2",
                    "TF_UPLOAD_RETRY_SLEEP": "0.01"}
        self.stop, self.seen, self.errors = False, [], []
        self.srv = ThreadingHTTPServer(("127.0.0.1", 0), UH)
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.url = f"http://127.0.0.1:{self.srv.server_port}/upload/secret-session-id"
        self.t = threading.Thread(target=self.courier_loop, daemon=True)
        self.t.start()
        self.verify_ok = True
        self.hold, self.office = False, False
        Upload.refuse, Upload.expired, Upload.deletes = False, False, []

    def tearDown(self):
        self.stop = True
        self.t.join(5)
        self.srv.shutdown()
        self.srv.server_close()

    def courier_loop(self):
        c = str(self.courier)
        while not self.stop:
            try:
                run("git", "-C", c, "pull", "-q", "--rebase", "origin", self.BRANCH)
                changed = False
                for kind_dir in ("fetch", "upload"):
                    for f in sorted((self.courier / kind_dir).glob("*.tfc")):
                        req = json.loads(open_box(f.read_bytes(), self.cpriv))
                        self.seen.append(req)
                        self.answer(req)
                        run("git", "-C", c, "rm", "-q", f"{kind_dir}/{f.name}")
                        changed = True
                if changed:
                    run("git", "-C", c, "add", "-A")
                    run("git", "-C", c, "commit", "-q", "-m", "courier")
                    run("git", "-C", c, "push", "-q", "origin", f"HEAD:{self.BRANCH}")
            except Exception as e:  # noqa: BLE001
                self.errors.append(repr(e))
            time.sleep(0.1)

    def answer(self, req):
        res = self.courier / "results"
        res.mkdir(exist_ok=True)
        (self.courier / "pickup").mkdir(exist_ok=True)
        cid = req["id"]
        if req["kind"] == "list":
            m = {"id": cid, "drive_id": req["drive_id"], "root_id": "ROOT",
                 "entries": [{"path": "a.pdf", "name": "a.pdf", "id": "F1", "size": 3, "is_folder": False,
                              "quickxor": "q", "sha256": "", "modified": "t"}]}
            (self.courier / "pickup" / f"{cid}.tfc").write_bytes(seal(json.dumps(m).encode(), req["reply_pub"]))
            (res / f"fetch-{cid}.json").write_text(json.dumps({"id": cid, "status": "ok", "kind": "list", "count": 1}))
        elif req["kind"] == "session" and self.hold:
            (res / f"upload-{cid}.json").write_text(json.dumps({"id": cid, "stage": "session", "status": "failed", "files": [
                {"index": 0, "bytes": req["files"][0]["bytes"], "sha256": req["files"][0]["sha256"], "ready": False,
                 "error": "held by an unfinished upload (empty placeholder) — not replaced; retry later",
                 "existing_bytes": 0, "existing_item_id": "HOLD"}]}))
        elif req["kind"] == "session":
            total = req["files"][0]["bytes"]
            Upload.received, Upload.sha, Upload.puts, Upload.gets = 0, hashlib.sha256(), [], 0
            Upload.fail_at, Upload.total = None, total
            m = {"id": cid, "drive_id": req["drive_id"], "folder_id": req["folder_id"],
                 "files": [{**req["files"][0], "upload_url": self.url, "expires": "x"}]}
            (self.courier / "pickup" / f"up-{cid}.tfc").write_bytes(seal(json.dumps(m).encode(), req["reply_pub"]))
            (res / f"upload-{cid}.json").write_text(json.dumps({"id": cid, "stage": "session", "status": "ok", "files": [
                {"index": 0, "bytes": total, "sha256": req["files"][0]["sha256"], "ready": True}]}))
        else:
            f = req["files"][0]
            ok = self.verify_ok and Upload.sha.hexdigest() == f["sha256"]
            row = {"index": 0, "bytes": f["bytes"], "sha256": f["sha256"], "item_id": f["item_id"], "verified": ok}
            if not ok:
                row["error"] = "content mismatch"
            if self.office:
                row["office_rewritten"] = True
            (res / f"upload-{cid}.json").write_text(json.dumps({"id": cid, "stage": "verify",
                                                                "status": "ok" if ok else "failed", "files": [row]}))

    def send(self, *args):
        return run(sys.executable, str(HERE / "courier" / "send.py"), *args, env=self.env)

    def test_list_writes_the_listing_deletes_the_pickup_and_the_key(self):
        out = self.box.parent / "listing.json"
        r = self.send("list", "--repo", str(self.box), "--drive", "D", "--path", "Jobs/Alpha", "--depth", "3",
                      "--branch", self.BRANCH, "--out", str(out), "--timeout", "60")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual(json.loads(out.read_text())["entries"][0]["id"], "F1")
        self.assertEqual((self.seen[0]["kind"], self.seen[0]["depth"], self.seen[0]["path"]),
                         ("list", 3, "Jobs/Alpha"))
        time.sleep(0.5)
        run("git", "-C", str(self.courier), "pull", "-q", "--rebase", "origin", self.BRANCH)
        self.assertEqual(list((self.courier / "pickup").glob("*.tfc")), [])        # pickup removed from git
        self.assertEqual(list(self.keys.glob("*.key")), [])                        # key thrown away
        again = self.send("list", "--repo", str(self.box), "--drive", "D", "--item", "X", "--branch", self.BRANCH,
                          "--out", str(out))
        self.assertEqual(again.returncode, 1)                                      # never overwrites
        self.assertEqual(self.errors, [])

    def test_upload_end_to_end_two_requests_bytes_never_in_git(self):
        src = self.box.parent / "Plan Set.pdf"
        src.write_bytes(os.urandom(700_000))
        summ = self.box.parent / "sum.json"
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "FOLD", "--branch", self.BRANCH,
                      "--chunk-mb", "0.625", "--timeout", "60", "--summary", str(summ),
                      f"{src}=Job - BC - Plan Set.pdf")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertTrue(r.stdout.startswith("upload "))                              # id printed first
        self.assertIn("✓ 0  700000 bytes", r.stdout)
        self.assertNotIn("secret-session-id", r.stdout + r.stderr)
        s = json.loads(summ.read_text())
        f = s["files"][0]
        self.assertEqual((f["name"], f["bytes"], f["item_id"], f["verified"]), ("Job - BC - Plan Set.pdf", 700000, "ITEM-1", True))
        self.assertEqual(f["sha256"], hashlib.sha256(src.read_bytes()).hexdigest())
        self.assertEqual(f["quickxor"], quickxor_bytes(src.read_bytes()))
        kinds = [q["kind"] for q in self.seen]
        self.assertEqual(kinds, ["session", "verify"])
        self.assertEqual(self.seen[1]["files"][0]["quickxor"], f["quickxor"])
        self.assertEqual([n for _, _, n in Upload.puts], [655360, 44640])           # 0.625 MiB = 2 x 320 KiB chunks
        self.assertFalse(any(a for _, a, _ in Upload.puts))
        raw = subprocess.run(["git", "-C", str(self.courier), "log", "--all", "-p", "--binary"],
                             capture_output=True).stdout
        self.assertNotIn(b"secret-session-id", raw)
        self.assertEqual(list(self.keys.glob("*")), [])
        self.assertEqual(self.errors, [])

    def test_failed_verify_exits_1_and_missing_result_exits_2_then_upload_wait_finishes(self):
        src = self.box.parent / "p.pdf"
        src.write_bytes(os.urandom(1000))
        self.verify_ok = False
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "FOLD", "--branch", self.BRANCH,
                      "--timeout", "60", str(src))
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("✗ 0", r.stdout)
        # no courier answer in time -> exit 2
        self.stop = True
        self.t.join(5)
        r2 = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "FOLD", "--branch", self.BRANCH,
                       "--timeout", "1", str(src))
        self.assertEqual(r2.returncode, 2, r2.stdout + r2.stderr)
        cid = r.stdout.split()[1]
        w = self.send("upload-wait", "--repo", str(self.box), "--branch", self.BRANCH, "--timeout", "1", cid)
        self.assertEqual(w.returncode, 1)                                            # finds the earlier (failed) verify

    def test_default_keys_dir_is_persistent_and_private_and_the_env_still_wins(self):
        code = "import sys; sys.path.insert(0, %r); import send; print(send._keys_dir())" % str(HERE / "courier")
        home = Path(tempfile.mkdtemp())
        env = {k: v for k, v in os.environ.items() if k != "TF_COURIER_KEYS"}
        r = run(sys.executable, "-c", code, env={**env, "HOME": str(home)})
        want = home / ".local" / "state" / "tf-courier-keys"
        self.assertEqual(r.stdout.strip(), str(want), r.stderr)
        self.assertEqual(want.stat().st_mode & 0o777, 0o700)
        r = run(sys.executable, "-c", code, env={**env, "HOME": str(home), "TF_COURIER_KEYS": str(self.keys)})
        self.assertEqual(r.stdout.strip(), str(self.keys))

    def test_a_file_that_cannot_be_sent_has_its_session_cancelled_and_key_and_state_go_at_the_end(self):
        src = self.box.parent / "p.pdf"
        src.write_bytes(os.urandom(1000))
        Upload.refuse = True
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "FOLD", "--branch", self.BRANCH,
                      "--timeout", "60", str(src))
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertEqual(Upload.deletes, [False])                        # DELETE to the session URL, no token
        self.assertNotIn("secret-session-id", r.stdout + r.stderr)
        self.assertEqual(list(self.keys.glob("*")), [])

    def test_summary_flags_a_held_placeholder_and_an_office_rewrite(self):
        src = self.box.parent / "p.pdf"
        src.write_bytes(os.urandom(1000))
        self.hold = True
        summ = self.box.parent / "held.json"
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "FOLD", "--branch", self.BRANCH,
                      "--timeout", "60", "--summary", str(summ), str(src))
        self.assertEqual(r.returncode, 1)
        row = json.loads(summ.read_text())["files"][0]
        self.assertTrue(row["held_placeholder"])
        self.assertEqual((row["existing_bytes"], row["existing_item_id"]), (0, "HOLD"))
        self.assertIn("held by an unfinished upload", row["error"])
        self.hold, self.office = False, True
        summ2 = self.box.parent / "office.json"
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "FOLD", "--branch", self.BRANCH,
                      "--timeout", "60", "--summary", str(summ2), f"{src}=Book.xlsx")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        row = json.loads(summ2.read_text())["files"][0]
        self.assertTrue(row["verified"] and row["office_rewritten"] and not row["held_placeholder"])

    def test_staging_names_the_action_would_refuse_are_stopped_before_anything_is_pushed(self):
        before = run("git", "-C", str(self.remote), "rev-parse", self.BRANCH).stdout
        for args in (("mkdir", "--parent", "P", "--name", "Bad#Name"), ("mkdir", "--parent", "P", "--name", "a/b"),
                     ("move", "--item", "I", "--dest", "D2", "--name", "what?")):
            r = self.send("staging", args[0], "--repo", str(self.box), "--drive", "D", "--branch", self.BRANCH, *args[1:])
            self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
            self.assertIn("not allowed", r.stdout + r.stderr)
        self.assertEqual(run("git", "-C", str(self.remote), "rev-parse", self.BRANCH).stdout, before)
        self.assertEqual(self.seen, [])

    # ---- upload-resume: the box died after the courier made the sessions
    def stage_crash(self, data, received=0, remove_pickup=False):
        """What a killed `upload` leaves: key + state on disk, the sealed pickup (and session result) in the repo."""
        cid, src = "20991231-010101-abcdef", self.box.parent / "Big.pdf"
        src.write_bytes(data)
        sha, qx, size = send.hash_file(src)
        self.keys.mkdir(mode=0o700, parents=True, exist_ok=True)
        priv, pub = new_keypair()
        (self.keys / f"up-{cid}.key").write_text(priv)
        f = {"index": 0, "local": str(src), "name": "Big.pdf", "bytes": size, "sha256": sha, "quickxor": qx,
             "item_id": "", "verified": False, "error": ""}
        (self.keys / f"up-{cid}.state.json").write_text(json.dumps({"files": [f]}))
        Upload.received, Upload.sha, Upload.puts, Upload.gets = received, hashlib.sha256(data[:received]), [], 0
        Upload.fail_at, Upload.total = None, size
        m = {"id": cid, "drive_id": "D", "folder_id": "FOLD", "files": [{**f, "upload_url": self.url, "expires": "x"}]}
        (self.box / "pickup").mkdir(exist_ok=True)
        (self.box / "results").mkdir(exist_ok=True)
        (self.box / "pickup" / f"up-{cid}.tfc").write_bytes(seal(json.dumps(m).encode(), pub))
        (self.box / "results" / f"upload-{cid}.json").write_text(json.dumps({"id": cid, "stage": "session", "status": "ok",
            "files": [{"index": 0, "bytes": size, "sha256": sha, "ready": True}]}))
        b = str(self.box)
        run("git", "-C", b, "add", "-A")
        run("git", "-C", b, "commit", "-q", "-m", "courier made sessions")
        if remove_pickup:
            run("git", "-C", b, "rm", "-q", f"pickup/up-{cid}.tfc")
            run("git", "-C", b, "commit", "-q", "-m", "picked up by an earlier run")
        run("git", "-C", b, "push", "-q", "origin", f"HEAD:{self.BRANCH}")
        return cid, src

    def resume(self, cid, *extra):
        return self.send("upload-resume", "--repo", str(self.box), "--branch", self.BRANCH, "--timeout", "30",
                         "--chunk-mb", "0.3125", *extra, cid)

    def test_resume_finishes_from_next_expected_ranges_and_verifies(self):
        data = os.urandom(700_000)
        cid, src = self.stage_crash(data, received=327_680)
        r = self.resume(cid)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual(Upload.puts[0][0], "bytes 327680-655359/700000")     # continued, not restarted
        self.assertEqual(Upload.sha.hexdigest(), hashlib.sha256(data).hexdigest())
        self.assertEqual([q["kind"] for q in self.seen], ["verify"])
        self.assertNotIn("secret-session-id", r.stdout + r.stderr)
        self.assertEqual(list(self.keys.glob("*")), [])
        self.assertEqual(self.errors, [])

    def test_resume_reads_a_pickup_that_an_earlier_run_already_removed_from_the_tip(self):
        data = os.urandom(400_000)
        cid, src = self.stage_crash(data, remove_pickup=True)
        r = self.resume(cid)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual(Upload.sha.hexdigest(), hashlib.sha256(data).hexdigest())

    def test_resume_reports_an_expired_session_and_cancels_it(self):
        cid, src = self.stage_crash(os.urandom(1000))
        Upload.expired = True
        r = self.resume(cid)
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("expired", r.stdout)
        self.assertEqual((Upload.puts, Upload.deletes), ([], [False]))
        self.assertEqual(self.seen, [])                                          # nothing sent: no verify request
        self.assertEqual(list(self.keys.glob("*")), [])

    def test_resume_without_saved_state_says_so(self):
        r = self.resume("20991231-020202-000000")
        self.assertEqual(r.returncode, 1)
        self.assertIn("no saved state", r.stdout + r.stderr)

    def test_wrong_checked_out_branch_is_refused(self):
        run("git", "-C", str(self.box), "checkout", "-q", "-b", "other")
        src = self.box.parent / "p.pdf"
        src.write_bytes(b"x")
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "F", "--branch", self.BRANCH, str(src))
        self.assertEqual(r.returncode, 1)
        self.assertIn("checked out", r.stderr + r.stdout)


class GitHelperTests(unittest.TestCase):
    """pull --autostash, push's last-pull-error, and reading a pickup from history."""

    def setUp(self):
        tmp = Path(tempfile.mkdtemp())
        self.remote, self.a, self.b = tmp / "r.git", tmp / "a", tmp / "b"
        run("git", "init", "-q", "--bare", "-b", "main", str(self.remote))
        for c in (self.a, self.b):
            run("git", "clone", "-q", str(self.remote), str(c))
            run("git", "-C", str(c), "config", "user.email", "t@example.invalid")
            run("git", "-C", str(c), "config", "user.name", "t")
        (self.a / "f.txt").write_text("one\n")
        run("git", "-C", str(self.a), "add", "-A")
        run("git", "-C", str(self.a), "commit", "-q", "-m", "init")
        run("git", "-C", str(self.a), "push", "-q", "origin", "HEAD:main")
        run("git", "-C", str(self.b), "pull", "-q", "origin", "main")

    def test_pull_with_a_dirty_tree_is_stashed_not_refused(self):
        (self.a / "g.txt").write_text("two\n")
        run("git", "-C", str(self.a), "add", "-A")
        run("git", "-C", str(self.a), "commit", "-q", "-m", "more")
        run("git", "-C", str(self.a), "push", "-q", "origin", "HEAD:main")
        (self.b / "f.txt").write_text("edited locally\n")                     # a dirty tracked file
        self.assertEqual(send.pull(self.b, "main").returncode, 0)
        self.assertTrue((self.b / "g.txt").exists())                           # the pull went through
        self.assertEqual((self.b / "f.txt").read_text(), "edited locally\n")   # and the edit is back

    def test_push_gives_up_with_the_last_pull_error(self):
        hook = self.remote / "hooks" / "pre-receive"
        hook.write_text("#!/bin/sh\nexit 1\n")
        hook.chmod(0o755)
        (self.b / "h.txt").write_text("x\n")
        run("git", "-C", str(self.b), "add", "-A")
        with unittest.mock.patch.object(send.time, "sleep"):
            with self.assertRaises(SystemExit) as cm:
                send.push(self.b, "m", branch="no-such-branch")                 # the pull of that branch fails too
        self.assertIn("push failed", str(cm.exception))
        self.assertIn("last pull error", str(cm.exception))

    def test_pickup_blob_falls_back_to_the_history_when_the_tip_no_longer_has_it(self):
        (self.a / "pickup").mkdir()
        (self.a / "pickup" / "x.tfc").write_bytes(b"sealed")
        run("git", "-C", str(self.a), "add", "-A")
        run("git", "-C", str(self.a), "commit", "-q", "-m", "pickup")
        run("git", "-C", str(self.a), "rm", "-q", "pickup/x.tfc")
        run("git", "-C", str(self.a), "commit", "-q", "-m", "picked up")
        run("git", "-C", str(self.a), "push", "-q", "origin", "HEAD:main")
        self.assertEqual(send._pickup_blob(self.b, "main", "pickup/x.tfc"), b"sealed")
        self.assertIsNone(send._pickup_blob(self.b, "main", "pickup/never.tfc"))


if __name__ == "__main__":
    unittest.main()
