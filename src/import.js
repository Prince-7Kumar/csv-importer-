require("dotenv").config();

const pool = require("./db");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");


const EXCEL_FILE = path.resolve(
  process.env.EXCEL_FILE || "files/Call_details_100000_dummy_records.xlsx",
);

// Values for columns the Excel file does not provide, e.g.
// IMPORT_DEFAULTS={"branch_id":"BR_0007"}
function readImportDefaults() {
  const value = process.env.IMPORT_DEFAULTS || '{"branch_id":"BR_0007"}';

  try {
    return JSON.parse(value);
  } catch (error) {
    throw new Error(`IMPORT_DEFAULTS in .env is not valid JSON: ${value}`);
  }
}

const pad = (value) => String(value).padStart(2, "0");

// Turns date/time cells into the exact text Excel shows, e.g.
// "2026-09-01 08:00:00" or "0:04:15". (xlsx's cellDates option shifts
// times by a few seconds on IST machines, so it is not used.)
function convertDateCells(worksheet) {
  for (const [address, cell] of Object.entries(worksheet)) {
    if (
      address.startsWith("!") ||
      cell.t !== "n" ||
      !cell.z ||
      !XLSX.SSF.is_date(cell.z)
    ) {
      continue;
    }

    const date = XLSX.SSF.parse_date_code(cell.v);
    const time = `${pad(date.H)}:${pad(date.M)}:${pad(date.S)}`;
    let text;

    if (/\[h/i.test(cell.z)) {
      // Elapsed time such as [h]:mm:ss (can exceed 24 hours).
      const seconds = Math.round(cell.v * 86400);
      text = `${Math.floor(seconds / 3600)}:${pad(Math.floor(seconds / 60) % 60)}:${pad(seconds % 60)}`;
    } else if (cell.v < 1) {
      // Time of day only, e.g. a duration typed as 0:04:15.
      text = `${date.H}:${pad(date.M)}:${pad(date.S)}`;
    } else {
      text = `${date.y}-${pad(date.m)}-${pad(date.d)} ${time}`;
    }

    worksheet[address] = { t: "s", v: text };
  }
}

// Rows keep their original Excel headers. Matching headers to table
// columns is done by public.import_metadata_rows (see sql/import_metadata.sql).
function readExcelRows(worksheet) {
  convertDateCells(worksheet);

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    defval: null,
    raw: true,
    blankrows: false,
  });

  // __rowNum__ is the 0-based sheet row; send it so database errors
  // can point at the real Excel row.
  return rows.map((row) => ({ ...row, __row: row.__rowNum__ + 1 }));
}

async function importWithPostgres(rows, defaults) {
  const client = await pool.connect();
  const dryRun =
    String(process.env.DRY_RUN ?? "true").toLowerCase() !== "false";

  try {
    await client.query("BEGIN");

    // Send every record in one database request.
    const result = await client.query(
      `SELECT public.import_metadata_rows($1::jsonb, $2::jsonb) AS summary`,
      [JSON.stringify(rows), JSON.stringify(defaults)],
    );

    const summary = result.rows[0].summary;

    if (dryRun) {
      await client.query("ROLLBACK");
      console.log("\nDry run completed. Nothing saved.");
    } else {
      await client.query("COMMIT");
      console.log("\nImport committed successfully.");
    }

    return { ...summary, dry_run: dryRun };
  } catch (error) {
    try {
      await client.query("ROLLBACK");
    } catch (rollbackError) {
      console.error("Rollback failed:", rollbackError.message);
    }

    throw error;
  } finally {
    client.release();
  }
}

async function main() {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(`Excel file not found: ${EXCEL_FILE}`);
  }

  const startedAt = Date.now();
  const defaults = readImportDefaults();

  // cellNF keeps each cell's number format, needed to spot date cells.
  // raw keeps CSV values as text instead of guessing dates/numbers.
  const workbook = XLSX.readFile(EXCEL_FILE, {
    cellNF: true,
    raw: true,
  });

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("The workbook contains no worksheets.");
  }

  const rows = readExcelRows(workbook.Sheets[sheetName]);

  if (rows.length === 0) {
    console.log("No data rows found.");
    return;
  }

  console.log(`Workbook: ${EXCEL_FILE}`);
  console.log(`Worksheet: ${sheetName}`);
  console.log(`Rows read: ${rows.length.toLocaleString()}`);

  const {
    mapped_columns: mappedColumns,
    unmapped_headers: unmappedHeaders,
    ...summary
  } = await importWithPostgres(rows, defaults);

  console.log("\nExcel headers stored:");
  console.table(mappedColumns);

  if (unmappedHeaders.length > 0) {
    console.warn("\nUnmapped Excel headers (ignored):");
    console.table(unmappedHeaders.map((header) => ({ header })));
  }

  console.table(summary);

  console.log(
    `\nElapsed time: ${((Date.now() - startedAt) / 1000).toFixed(2)} seconds`,
  );
  console.log(
    `Inserted: ${summary.rows_inserted}; ` +
      `existing DB duplicates: ${summary.existing_database_duplicates}; ` +
      `repeated Excel IDs: ${summary.repeated_excel_call_ids}`,
  );
}

main()
  .catch((error) => {
    console.error("\nImport failed:", error.message);
    process.exitCode = 1;
  })
  .finally(async () => {
    await pool.end();
  });
