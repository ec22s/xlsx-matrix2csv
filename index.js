'use strict';
import { read } from "xlsx";
import { readFile, writeFile } from 'fs/promises';

const config = {};

// main
loadConfig().then(() =>
  buildCsvArray2D().then((csv) =>
    writeCsvFile(csv)
  )
);

async function buildCsvArray2D() {
  const { file, sheetIndex, rowRange, colRange } = config.xlsx;
  try {
    const { SheetNames, Sheets } = read(await readFile(file), {
      // cellDates: true,
      // codepage: 932,
      dense: true,
    });
    const data = Sheets[SheetNames[sheetIndex]]
      .filter(($) => Array.isArray($))
      .filter((_, i) => (i + 1) >= rowRange[0] && i < rowRange[1])
      .map((row) => row
        .filter((_, j) => (j + 1) >= colRange[0] && j < colRange[1])
        .map((col) => col.v)
      );
    console.log('[WIP]', data, config.xlsx);
  } catch (err) {
    console.error(`[ERROR] failed to load ${file}`, err);
  }

  return [ // stub
    [
      1, 2, 3
    ],
    [
      4, 5, 6
    ]
  ];
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

async function writeCsvFile(csvArray2D) {
  const {
    addTimestamp, basePath, overWrite, delimCol, delimRow, extraData
  } = config.csv;
  const outputFile = basePath
    + (addTimestamp ? `-${(new Date).getTime()}` : '')
    + '.csv';
  const data = csvArray2D.map((row) =>
    extraData.map((extraRow) => row.concat(extraRow))
  )
    .flat()
    .map((row) =>
      row.map((col) => `"${col}"`).join(delimCol)
    ).join(delimRow);
  console.log('[CHECK]', outputFile, `\n${data}`);
  await writeFile(outputFile, data, { overwrite: overWrite });
}
