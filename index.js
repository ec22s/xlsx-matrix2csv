'use strict';
import { read, utils } from "xlsx";
import { readFile, writeFile } from 'fs/promises';

const config = {};

// main
loadConfig().then(() =>
  buildCsvArray2D().then((csv) => {
    writeCsvFile(csv)
  })
);

async function buildCsvArray2D() {
  const { file, sheetIndex, range } = config.xlsx;
  try {
    const { SheetNames, Sheets } = read(await readFile(file), {
      // cellDates: true,
      // codepage: 932,
      dense: true,
    });
    const originalArray = utils.sheet_to_json(
      Sheets[SheetNames[sheetIndex]],
      { header: 1, range: range.join(":") }
    );
    const matrix = [];
    for (const row of originalArray) {
      const tmpRow = [];
      if (Array.isArray(row)) {
        for (const col of row) {
          tmpRow.push(col != null ? col : "");
        }
      }
      matrix.push(tmpRow);
    }
    return matrix[0].map((headerKey, col) => matrix.slice(1)
      .map((bodyRow) => {
        const val = bodyRow[col];
        if (val != null && val !== "") return [headerKey, val]
      })
      .filter((elem) => elem != null)
    ).flat();
  } catch (err) {
    console.error(`[ERROR] failed to load ${file}`, err);
  }
}

async function loadConfig() {
  const args = (process.argv || []).slice(2);
  if (!args.length) {
    console.error('[ERROR] config JSON not provided.');
    process.exit(1);
  }
  if (args.length > 1) {
    console.warn('[WARN] two or more argument(s) found but ignored.');
  }
  try {
    Object.assign(config, JSON.parse(await readFile(args[0])));
  } catch (err) {
    console.error("[ERROR] failed to load config JSON.\n", err);
    process.exit(1);
  }
}

function outputFilePath(basePath, timestamp, ext) {
  return basePath
    + (timestamp ? `-${(new Date).getTime()}` : '')
    + `.${ext}`;
}

async function writeCsvFile(csvArray2D) {
  const { csv: csvConfig, json: jsonConfig } = config;
  const {
    basePath, overWrite, timestamp, delimCol, delimRow, extraData
  } = csvConfig;
  const csvFile = outputFilePath(basePath, timestamp, 'csv');
  const data = csvArray2D.map((row) =>
    extraData.map((extraRow) => row.concat(extraRow))
  ).flat();
  if (jsonConfig.output) {
    const { basePath, indent, overWrite, timestamp } = jsonConfig;
    const jsonFile = outputFilePath(basePath, timestamp, 'json');
    await writeFile(
      jsonFile, JSON.stringify(data, null, indent), { overwrite: overWrite }
    );
  }
  const csvString = data.map((row) =>
    row.map((col) => `"${col}"`).join(delimCol)
  ).join(delimRow);
  await writeFile(csvFile, csvString, { overwrite: overWrite });
}
