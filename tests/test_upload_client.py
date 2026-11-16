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

    def do_GET(self):
        Upload.gets += 1
        self._json(200, {"nextExpectedRanges": [f"{Upload.received}-"]})

    def do_PUT(self):
        n = int(self.headers["Content-Length"])
        rng = self.headers["Content-Range"]
        Upload.puts.append((rng, "Authorization" in self.headers, n))
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

    def test_wrong_checked_out_branch_is_refused(self):
        run("git", "-C", str(self.box), "checkout", "-q", "-b", "other")
        src = self.box.parent / "p.pdf"
        src.write_bytes(b"x")
        r = self.send("upload", "--repo", str(self.box), "--drive", "D", "--folder", "F", "--branch", self.BRANCH, str(src))
        self.assertEqual(r.returncode, 1)
        self.assertIn("checked out", r.stderr + r.stdout)


if __name__ == "__main__":
    unittest.main()
