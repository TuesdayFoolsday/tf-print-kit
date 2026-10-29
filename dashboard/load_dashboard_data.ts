/**
 * Office Script: load "Dashboard Data.txt" (written nightly by build_dashboard.py) into the Dashboard workbook.
 *
 * Power Automate passes the file's content as `data`. It may arrive as the JSON text itself or wrapped as
 * {"$content-type": "...", "$content": "<base64>"} — both are handled.
 *
 * For every table in data.tables: the table's formula columns (e.g. in_period, the GC stats) are read from its first
 * row and kept; all rows are replaced with the new data; formula columns get their formula back on every row.
 * A table always keeps at least one row. Column formats are re-applied from the first row.
 * Returns a one-line summary (rows per table), shown in the flow's run history.
 */

interface TableData {
  columns: string[];
  rows: (string | number | boolean | null)[][];
}
interface DashboardData {
  stamp: string;
  about: string[];
  tables: { [name: string]: TableData };
}
interface Wrapped {
  "$content-type"?: string;
  "$content"?: string;
}

function main(workbook: ExcelScript.Workbook, data: string): string {
  const parsed = parseInput(data);
  const summary: string[] = [];

  const names = Object.keys(parsed.tables);
  for (let t = 0; t < names.length; t++) {
    const name = names[t];
    const spec = parsed.tables[name];
    const table = workbook.getTable(name);
    if (!table) {
      throw new Error("Workbook has no table named " + name);
    }
    const headers: string[] = table.getHeaderRowRange().getValues()[0].map((h) => String(h));
    const firstRow = table.getRangeBetweenHeaderAndTotal().getRow(0);
    const firstFormulas = firstRow.getFormulas()[0];
    const firstFormats = firstRow.getNumberFormats()[0];

    const keep: string[] = [];
    const source: number[] = [];
    for (let i = 0; i < headers.length; i++) {
      const f = String(firstFormulas[i]);
      const isFormula = f.length > 0 && f.charAt(0) === "=";
      const idx = spec.columns.indexOf(headers[i]);
      if (!isFormula && idx < 0) {
        throw new Error(name + ": column '" + headers[i] + "' has no data and no formula");
      }
      keep.push(isFormula && idx < 0 ? f : "");
      source.push(idx);
    }

    const dataRows = spec.rows.length > 0 ? spec.rows : [spec.columns.map(() => null)];
    const rows: (string | number | boolean)[][] = dataRows.map((r) =>
      headers.map((h, i) => {
        if (keep[i] !== "") {
          return keep[i];
        }
        const v = r[source[i]];
        return v === null || v === undefined ? "" : v;
      })
    );

    // Replace rows: keep row 0 (so the table never empties), overwrite it, drop the rest, append the remainder.
    const count = table.getRowCount();
    if (count > 1) {
      table.deleteRowsAt(1, count - 1);
    }
    table.getRangeBetweenHeaderAndTotal().getRow(0).setValues([rows[0]]);
    if (rows.length > 1) {
      table.addRows(-1, rows.slice(1));
    }

    // Re-apply each column's number format to every row.
    for (let i = 0; i < headers.length; i++) {
      table.getColumn(headers[i]).getRangeBetweenHeaderAndTotal().setNumberFormat(String(firstFormats[i]));
    }
    summary.push(name + " " + rows.length);
  }

  const dash = workbook.getWorksheet("Dashboard");
  if (dash) {
    dash.getRange("A2").setValue(parsed.stamp);
  }
  const about = workbook.getWorksheet("About");
  if (about && parsed.about) {
    about.getRange("A1:A300").clear(ExcelScript.ClearApplyTo.contents);
    if (parsed.about.length > 0) {
      about.getRange("A1:A" + parsed.about.length).setValues(parsed.about.map((line) => [line]));
    }
  }
  return "Loaded: " + summary.join(" · ");
}

function parseInput(data: string): DashboardData {
  let text = data.trim();
  if (text.indexOf('"$content"') >= 0 && text.indexOf('"tables"') < 0) {
    const wrapped = JSON.parse(text) as Wrapped;
    text = utf8Decode(base64Decode(wrapped["$content"] || ""));
  }
  const parsed = JSON.parse(text) as DashboardData;
  if (!parsed.tables) {
    throw new Error("Data has no 'tables' — wrong file?");
  }
  return parsed;
}

function base64Decode(b64: string): number[] {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const clean = b64.replace(/[^A-Za-z0-9+/]/g, "");
  const bytes: number[] = [];
  let buffer = 0;
  let bits = 0;
  for (let i = 0; i < clean.length; i++) {
    buffer = (buffer << 6) | chars.indexOf(clean.charAt(i));
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      bytes.push((buffer >> bits) & 0xff);
    }
  }
  return bytes;
}

function utf8Decode(bytes: number[]): string {
  let out = "";
  let i = 0;
  while (i < bytes.length) {
    const b = bytes[i];
    let cp = 0;
    if (b < 0x80) {
      cp = b; i += 1;
    } else if (b < 0xe0) {
      cp = ((b & 0x1f) << 6) | (bytes[i + 1] & 0x3f); i += 2;
    } else if (b < 0xf0) {
      cp = ((b & 0x0f) << 12) | ((bytes[i + 1] & 0x3f) << 6) | (bytes[i + 2] & 0x3f); i += 3;
    } else {
      cp = ((b & 0x07) << 18) | ((bytes[i + 1] & 0x3f) << 12) | ((bytes[i + 2] & 0x3f) << 6) | (bytes[i + 3] & 0x3f); i += 4;
    }
    out += String.fromCodePoint(cp);
  }
  return out;
}
