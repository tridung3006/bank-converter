const express = require("express");
const multer = require("multer");
const path = require("node:path");
const fs = require("node:fs");
const os = require("node:os");
const XLSX = require("xlsx");
const { createWorker } = require("tesseract.js");
const sharp = require("sharp");

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
    files: 20,
  },
});

const PORT = process.env.PORT || 5177;

const OUTPUT_HEADERS = [
  "Ngày",
  "Sổ nhật ký",
  "Diễn giải",
  "Chi tiết bút toán/Đối tác",
  "Chi tiết bút toán/Bộ phận",
  "Chi tiết bút toán/Tài khoản phân tích",
  "Chi tiết bút toán/Tài khoản",
  "Chi tiết bút toán/Nợ",
  "Chi tiết bút toán/Có",
];

const DEFAULT_BANK_ACCOUNT = "112";
const IMAGE_BANK_ACCOUNT = "112181";
const BIDV_BANK_ACCOUNT = "112191";
const VTTT_BANK_ACCOUNT = "112120";
const TECH_BANK_ACCOUNTS = {
  TECH: "112",
  TECH014: "112111",
  TECH588: "112114",
};
const POS_EXPENSE_ACCOUNT = "641521";
const VAT_INPUT_ACCOUNT = "133111";
const POS_FEE_DESCRIPTION = "Phí cà thẻ POS";
const POS_FEE_VAT_DESCRIPTION = "Thuế phí cà thẻ POS";
const TECH_FEE_INTEREST_DESCRIPTION = "Phí, lãi techcombank";
const TECH_TAX_DESCRIPTION = "Thuế";
const EXCEL_COUNTERPART_ACCOUNTS = {
  debitWhenBankCredit: "331111",
  creditWhenBankDebit: "131111",
  posDebitWhenBankCredit: POS_EXPENSE_ACCOUNT,
};
const IMAGE_COUNTERPART_ACCOUNTS = {
  debitWhenBankCredit: "641231",
  creditWhenBankDebit: "113112",
};
const JOURNAL_BANK_CREDIT = "Ngân hàng - Báo Có";
const JOURNAL_BANK_DEBIT = "Ngân hàng - Báo Nợ";

app.use(express.static(path.join(__dirname, "public")));

app.post("/convert", upload.array("files"), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: "Chua chon file Excel." });
    }

    const rows = [];
    const summaries = [];

    for (const file of req.files) {
      const result = await convertUploadedFile(file);
      rows.push(...result.rows);
      summaries.push(result.summary);
    }

    if (rows.length === 0) {
      return res.status(422).json({
        error: "Khong tim thay sheet co du lieu hop le. App ho tro BIDV, TECH, TECH014, TECH588, VTTT, hoac sheet co header Ngay GD / No-Co / So tien / Noi dung thanh toan.",
        summaries,
      });
    }

    const outputBook = XLSX.utils.book_new();
    const outputSheet = XLSX.utils.aoa_to_sheet([OUTPUT_HEADERS, ...rows]);
    outputSheet["!cols"] = [
      { wch: 14 },
      { wch: 22 },
      { wch: 64 },
      { wch: 28 },
      { wch: 26 },
      { wch: 30 },
      { wch: 24 },
      { wch: 18 },
      { wch: 18 },
    ];

    XLSX.utils.book_append_sheet(outputBook, outputSheet, "ERP_Import");
    const outputBuffer = XLSX.write(outputBook, {
      type: "buffer",
      bookType: "xlsx",
      compression: true,
    });

    const fileName = `ERP_Import_${timestampForFileName()}.xlsx`;
    const savedPath = saveOutputToDownloads(fileName, outputBuffer);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.setHeader("X-Conversion-Summary", encodeURIComponent(JSON.stringify(summaries)));
    res.setHeader("X-Saved-Path", encodeURIComponent(savedPath));
    return res.send(outputBuffer);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: error.message || "Convert file that bai." });
  }
});

app.listen(PORT, () => {
  console.log(`Bank converter app is running at http://localhost:${PORT}`);
});

async function convertUploadedFile(file) {
  if (isImageUpload(file)) {
    return convertImageFile(file);
  }

  const workbook = XLSX.read(file.buffer, {
    type: "buffer",
    cellDates: true,
    raw: false,
  });

  return convertWorkbook(workbook, file.originalname);
}

function isImageUpload(file) {
  const name = file.originalname.toLowerCase();
  return file.mimetype.startsWith("image/") || /\.(png|jpe?g|webp|bmp)$/i.test(name);
}

async function convertImageFile(file) {
  const gridConverted = await convertDbCrImageGrid(file.buffer);
  const converted = gridConverted.rows.length > 0
    ? gridConverted
    : convertDbCrText(await recognizeImageText(file.buffer), IMAGE_BANK_ACCOUNT, IMAGE_COUNTERPART_ACCOUNTS);

  return {
    rows: converted.rows,
    summary: {
      fileName: file.originalname,
      sheets: [
        {
          sheetName: "OCR_IMAGE",
          outputRows: converted.rows.length,
          mode: converted.mode || "IMAGE_OCR_DB_CR",
          detectedRows: converted.detectedRows,
        },
      ],
      outputRows: converted.rows.length,
    },
  };
}

async function convertDbCrImageGrid(buffer) {
  const grid = await detectDbCrImageGrid(buffer);

  if (!grid) {
    return { rows: [], detectedRows: 0, mode: "IMAGE_GRID_OCR_DB_CR" };
  }

  const worker = await createWorker("eng");
  const parsedRows = [];

  try {
    for (let index = 0; index < grid.rowLines.length - 1; index += 1) {
      const top = grid.rowLines[index];
      const bottom = grid.rowLines[index + 1];
      const rowHeight = bottom - top;

      if (rowHeight < 14) {
        continue;
      }

      const [dateText, dbCrText, amountText, descriptionText] = await Promise.all([
        recognizeCell(worker, buffer, grid.cols.date[0], top, grid.cols.date[1] - grid.cols.date[0], rowHeight, "date"),
        recognizeCell(worker, buffer, grid.cols.dbCr[0], top, grid.cols.dbCr[1] - grid.cols.dbCr[0], rowHeight, "dbcr"),
        recognizeCell(worker, buffer, grid.cols.amount[0], top, grid.cols.amount[1] - grid.cols.amount[0], rowHeight, "amount"),
        recognizeCell(worker, buffer, grid.cols.description[0], top, grid.cols.description[1] - grid.cols.description[0], rowHeight, "description"),
      ]);

      const date = parseDate(dateText);
      const dbCr = normalizeDbCr(dbCrText);
      const amount = Math.abs(toNumber(amountText));
      const description = cleanText(descriptionText);

      if (!date || !dbCr || amount === 0 || !description) {
        continue;
      }

      parsedRows.push({ date, dbCr, amount, description });
    }
  } finally {
    await worker.terminate();
  }

  const converted = convertParsedDbCrRows(parsedRows, IMAGE_BANK_ACCOUNT, IMAGE_COUNTERPART_ACCOUNTS);
  return {
    rows: converted,
    detectedRows: parsedRows.length,
    mode: "IMAGE_GRID_OCR_DB_CR",
  };
}

async function detectDbCrImageGrid(buffer) {
  const { data, info } = await sharp(buffer).grayscale().raw().toBuffer({ resolveWithObject: true });
  const verticals = [];
  const horizontalGroups = [];

  for (let x = 0; x < info.width; x += 1) {
    let dark = 0;
    for (let y = 0; y < info.height; y += 1) {
      if (data[y * info.width + x] < 80) dark += 1;
    }
    if (dark > info.height * 0.35) verticals.push(x);
  }

  let activeGroup = null;
  for (let y = 0; y < info.height; y += 1) {
    let dark = 0;
    for (let x = 0; x < info.width; x += 1) {
      if (data[y * info.width + x] < 120) dark += 1;
    }

    if (dark > info.width * 0.55) {
      if (!activeGroup || y - activeGroup.end > 2) {
        activeGroup = { start: y, end: y };
        horizontalGroups.push(activeGroup);
      } else {
        activeGroup.end = y;
      }
    }
  }

  const verticalLines = groupLineCenters(verticals).filter((x, index, arr) => index === 0 || x - arr[index - 1] > 35);
  const rowLines = horizontalGroups
    .map((group) => group.end)
    .filter((y, index, arr) => index === 0 || y - arr[index - 1] > 12);

  const descriptionRight = verticalLines.find((x) => x > info.width * 0.78 && x < info.width * 0.9);

  if (verticalLines.length < 6 || rowLines.length < 3 || !descriptionRight) {
    return null;
  }

  return {
    cols: {
      date: [verticalLines[1], verticalLines[2]],
      dbCr: [verticalLines[3], verticalLines[4]],
      amount: [verticalLines[4], verticalLines[5]],
      description: [verticalLines[5], descriptionRight],
    },
    rowLines,
  };
}

function groupLineCenters(values) {
  const groups = [];

  for (const value of values) {
    const last = groups[groups.length - 1];
    if (!last || value - last[last.length - 1] > 2) {
      groups.push([value]);
    } else {
      last.push(value);
    }
  }

  return groups.map((group) => Math.round(group.reduce((sum, value) => sum + value, 0) / group.length));
}

async function recognizeCell(worker, buffer, left, top, width, height, kind) {
  const scale = kind === "description" ? 4 : 6;
  const threshold = kind === "description" ? 178 : 185;
  const processed = await sharp(buffer)
    .extract({
      left: Math.max(Math.round(left + 2), 0),
      top: Math.max(Math.round(top + 2), 0),
      width: Math.max(Math.round(width - 4), 1),
      height: Math.max(Math.round(height - 4), 1),
    })
    .resize({ width: Math.max(Math.round((width - 4) * scale), 240), withoutEnlargement: false })
    .grayscale()
    .normalize()
    .sharpen()
    .threshold(threshold)
    .png()
    .toBuffer();

  const result = await worker.recognize(processed);
  return result.data.text.replace(/\s+/g, " ").trim();
}

async function recognizeImageText(buffer) {
  const processed = await preprocessImageForOcr(buffer);
  const worker = await createWorker("eng");
  try {
    const result = await worker.recognize(processed);
    return result.data.text || "";
  } finally {
    await worker.terminate();
  }
}

async function preprocessImageForOcr(buffer) {
  const metadata = await sharp(buffer).metadata();
  return sharp(buffer)
    .resize({ width: Math.max((metadata.width || 0) * 4, 3200), withoutEnlargement: false })
    .grayscale()
    .normalize()
    .sharpen()
    .threshold(185)
    .png()
    .toBuffer();
}

function convertWorkbook(workbook, fileName) {
  const rows = [];
  const supportedSheets = [];

  for (const sheetName of workbook.SheetNames) {
    const normalizedName = sheetName.trim().toUpperCase();
    const sheet = workbook.Sheets[sheetName];

    if (normalizedName === "BIDV") {
      const converted = convertBIDV(sheet);
      rows.push(...converted);
      supportedSheets.push({ sheetName, outputRows: converted.length });
    }

    if (Object.hasOwn(TECH_BANK_ACCOUNTS, normalizedName)) {
      const converted = convertTECH(sheet, TECH_BANK_ACCOUNTS[normalizedName]);
      rows.push(...converted);
      supportedSheets.push({ sheetName, outputRows: converted.length });
      continue;
    }

    if (normalizedName === "VTTT") {
      const converted = convertVTTT(sheet);
      rows.push(...converted);
      supportedSheets.push({ sheetName, outputRows: converted.length });
      continue;
    }

    const headerBased = convertHeaderBasedDbCr(sheet);
    if (headerBased.rows.length > 0) {
      rows.push(...headerBased.rows);
      supportedSheets.push({
        sheetName,
        outputRows: headerBased.rows.length,
        mode: "HEADER_DB_CR",
        headerRow: headerBased.headerRow,
      });
    }
  }

  return {
    rows,
    summary: {
      fileName,
      sheets: supportedSheets,
      outputRows: rows.length,
    },
  };
}

function convertDbCrText(text, bankAccount = DEFAULT_BANK_ACCOUNT, counterpartAccounts = EXCEL_COUNTERPART_ACCOUNTS) {
  const parsedRows = parseDbCrOcrRows(text);
  const output = convertParsedDbCrRows(parsedRows, bankAccount, counterpartAccounts);

  return {
    rows: output,
    detectedRows: parsedRows.length,
    mode: "IMAGE_OCR_DB_CR",
  };
}

function convertParsedDbCrRows(parsedRows, bankAccount = DEFAULT_BANK_ACCOUNT, counterpartAccounts = EXCEL_COUNTERPART_ACCOUNTS) {
  const output = [];

  for (const row of parsedRows) {
    if (row.dbCr === "D") {
      output.push(...createVoucherRows(row.date, row.description, row.amount, bankAccount, "credit", counterpartAccounts));
    } else if (row.dbCr === "C") {
      output.push(...createVoucherRows(row.date, row.description, row.amount, bankAccount, "debit", counterpartAccounts));
    }
  }

  return output;
}

function parseDbCrOcrRows(text) {
  const rows = [];
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/[|]/g, " ").replace(/\s+/g, " ").trim())
    .filter(Boolean);

  for (const line of lines) {
    const normalizedLine = line
      .replace(/\b0(?=\s+[\d,]+\s+)/g, "D")
      .replace(/\bO(?=\s+[\d,]+\s+)/g, "D");

    const match = normalizedLine.match(
      /(?:\S+\s+)?(\d{1,2}\/\d{1,2}\/\d{4})\s+\d{1,2}\/\d{1,2}\/\d{4}\s+([DC])\s+([\d.,]+)\s+(.+)$/i,
    );

    if (match) {
      rows.push({
        date: parseDate(match[1]),
        dbCr: match[2].toUpperCase(),
        amount: Math.abs(toNumber(match[3])),
        description: cleanText(match[4]),
      });
      continue;
    }

    if (rows.length > 0 && !looksLikeHeaderOrFooter(line)) {
      rows[rows.length - 1].description = cleanText(`${rows[rows.length - 1].description} ${line}`);
    }
  }

  return rows.filter((row) => row.date && row.amount > 0 && row.description);
}

function normalizeDbCr(value) {
  const text = normalizeText(value).toUpperCase();

  if (/\bC\b/.test(text) || text === "€") {
    return "C";
  }

  if (/\bD\b/.test(text) || /\b0\b/.test(text) || /\bO\b/.test(text)) {
    return "D";
  }

  if (text.includes("C")) {
    return "C";
  }

  if (text.includes("D") || text.includes("0") || text.includes("O")) {
    return "D";
  }

  return "";
}


function looksLikeHeaderOrFooter(line) {
  const label = normalizeText(line);
  return (
    label.includes("tran code") ||
    label.includes("tran date") ||
    label.includes("value date") ||
    label.includes("payment detail") ||
    label.includes("transfer detail") ||
    label.includes("amount")
  );
}

function convertHeaderBasedDbCr(sheet) {
  const output = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const header = findDbCrHeader(sheet, range);

  if (!header) {
    return { rows: output, headerRow: null };
  }

  for (let rowIndex = header.rowIndex + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const transactionDate = parseDate(cellValue(sheet, rowIndex, header.dateCol));
    const dbCr = normalizeText(cellValue(sheet, rowIndex, header.dbCrCol)).toUpperCase();
    const amount = Math.abs(toNumber(cellValue(sheet, rowIndex, header.amountCol)));
    const description = cleanText(cellValue(sheet, rowIndex, header.paymentDetailCol));

    if (!transactionDate || amount === 0 || (dbCr !== "D" && dbCr !== "C")) {
      continue;
    }

    if (dbCr === "D") {
      output.push(...createVoucherRows(transactionDate, description, amount, DEFAULT_BANK_ACCOUNT, "credit"));
    } else {
      output.push(...createVoucherRows(transactionDate, description, amount, DEFAULT_BANK_ACCOUNT, "debit"));
    }
  }

  return { rows: output, headerRow: header.rowIndex + 1 };
}

function findDbCrHeader(sheet, range) {
  const maxHeaderRow = Math.min(range.e.r, 80);

  for (let rowIndex = range.s.r; rowIndex <= maxHeaderRow; rowIndex += 1) {
    const columns = {};

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const label = normalizeHeader(cellValue(sheet, rowIndex, columnIndex));

      if (!Number.isInteger(columns.dateCol) && (label.includes("ngay gd") || label.includes("tran date"))) {
        columns.dateCol = columnIndex;
      }

      if (!Number.isInteger(columns.dbCrCol) && (label.includes("no/co") || label.includes("db/cr") || label.includes("db cr"))) {
        columns.dbCrCol = columnIndex;
      }

      if (!Number.isInteger(columns.amountCol) && (label.includes("so tien") || label.includes("amount"))) {
        columns.amountCol = columnIndex;
      }

      if (!Number.isInteger(columns.paymentDetailCol) && (label.includes("noi dung thanh toan") || label.includes("payment detail"))) {
        columns.paymentDetailCol = columnIndex;
      }
    }

    if (
      Number.isInteger(columns.dateCol) &&
      Number.isInteger(columns.dbCrCol) &&
      Number.isInteger(columns.amountCol) &&
      Number.isInteger(columns.paymentDetailCol)
    ) {
      return { rowIndex, ...columns };
    }
  }

  return null;
}

function convertBIDV(sheet) {
  const output = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const firstDataRow = 12; // Excel row 13, zero-based.
  let posGrossTotal = 0;
  let posTransactionDate = "";

  for (let rowIndex = firstDataRow; rowIndex <= range.e.r; rowIndex += 1) {
    const transactionDate = parseDate(cellValue(sheet, rowIndex, 2));
    const debitAmount = toNumber(cellValue(sheet, rowIndex, 3));
    const creditAmount = toNumber(cellValue(sheet, rowIndex, 4));
    const description = cleanText(cellValue(sheet, rowIndex, 8));

    if (!transactionDate || (debitAmount === 0 && creditAmount === 0)) {
      continue;
    }

    if (debitAmount !== 0) {
      if (hasPOS(description)) {
        posGrossTotal += Math.abs(debitAmount);
        if (!posTransactionDate) {
          posTransactionDate = transactionDate;
        }
      } else {
        output.push(...createVoucherRows(transactionDate, description, Math.abs(debitAmount), BIDV_BANK_ACCOUNT, "credit"));
      }
    } else {
      output.push(...createVoucherRows(transactionDate, description, Math.abs(creditAmount), BIDV_BANK_ACCOUNT, "debit"));
    }
  }

  if (posGrossTotal > 0) {
    output.push(...createBIDVPosVoucherRows(posTransactionDate, posGrossTotal));
  }

  return output;
}

function convertTECH(sheet, bankAccount) {
  const output = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const firstDataRow = 21; // Excel row 22, zero-based.
  let posFeeTotal = 0;
  let posTransactionDate = "";
  let feeInterestTotal = 0;
  let taxTotal = 0;
  let feeTaxTransactionDate = "";

  for (let rowIndex = firstDataRow; rowIndex <= range.e.r; rowIndex += 1) {
    const transactionDate = parseDate(cellValue(sheet, rowIndex, 1));
    const description = cleanText(cellValue(sheet, rowIndex, 6));
    const debitAmount = Math.abs(toNumber(cellValue(sheet, rowIndex, 7)));
    const creditAmount = Math.abs(toNumber(cellValue(sheet, rowIndex, 8)));
    const feeInterestAmount = Math.abs(toNumber(cellValue(sheet, rowIndex, 9)));
    const taxAmount = Math.abs(toNumber(cellValue(sheet, rowIndex, 10)));

    if (!transactionDate) {
      continue;
    }

    if (feeInterestAmount !== 0 || taxAmount !== 0) {
      if (!feeTaxTransactionDate) {
        feeTaxTransactionDate = transactionDate;
      }
      feeInterestTotal += feeInterestAmount;
      taxTotal += taxAmount;
    }

    if (debitAmount !== 0 || creditAmount !== 0) {
      if (debitAmount !== 0) {
        if (hasPOS(description)) {
          posFeeTotal += debitAmount;
          if (!posTransactionDate) {
            posTransactionDate = transactionDate;
          }
        } else {
          output.push(...createVoucherRows(transactionDate, description, debitAmount, bankAccount, "credit"));
        }
      } else {
        output.push(...createVoucherRows(transactionDate, description, creditAmount, bankAccount, "debit"));
      }
    }
  }

  if (posFeeTotal > 0) {
    output.push(...createVoucherRows(posTransactionDate, POS_FEE_DESCRIPTION, posFeeTotal, bankAccount, "credit"));
  }

  if (feeInterestTotal > 0) {
    output.push(...createDebitBankCreditRows(feeTaxTransactionDate, TECH_FEE_INTEREST_DESCRIPTION, feeInterestTotal, POS_EXPENSE_ACCOUNT, bankAccount));
  }

  if (taxTotal > 0) {
    output.push(...createDebitBankCreditRows(feeTaxTransactionDate, TECH_TAX_DESCRIPTION, taxTotal, VAT_INPUT_ACCOUNT, bankAccount));
  }

  return output;
}

function convertVTTT(sheet) {
  const output = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const FIRST_DATA_EXCEL_ROW = 9;
  const firstDataRow = FIRST_DATA_EXCEL_ROW - 1;

  for (let rowIndex = firstDataRow; rowIndex <= Math.max(range.e.r, firstDataRow); rowIndex += 1) {
    const transactionDate = parseDate(cellValue(sheet, rowIndex, 0));
    const description = cleanText(cellValue(sheet, rowIndex, 1));
    const debitAmount = Math.abs(toNumber(cellValue(sheet, rowIndex, 3)));
    const creditAmount = Math.abs(toNumber(cellValue(sheet, rowIndex, 4)));

    if (!transactionDate || (debitAmount === 0 && creditAmount === 0)) {
      continue;
    }

    if (debitAmount !== 0) {
      output.push(...createVoucherRows(transactionDate, description, debitAmount, VTTT_BANK_ACCOUNT, "credit"));
    } else {
      output.push(...createVoucherRows(transactionDate, description, creditAmount, VTTT_BANK_ACCOUNT, "debit"));
    }
  }

  return output;
}

function createVoucherRows(
  transactionDate,
  description,
  amount,
  bankAccount,
  bankSide,
  counterpartAccounts = EXCEL_COUNTERPART_ACCOUNTS,
) {
  if (bankSide === "credit") {
    const debitCounterpart = hasPOS(description) && counterpartAccounts.posDebitWhenBankCredit
      ? counterpartAccounts.posDebitWhenBankCredit
      : counterpartAccounts.debitWhenBankCredit;

    return [
      [transactionDate, JOURNAL_BANK_DEBIT, description, "", "", "", bankAccount, "", amount],
      ["", "", "", "", "", "", debitCounterpart, amount, ""],
    ];
  }

  return [
    [transactionDate, JOURNAL_BANK_CREDIT, description, "", "", "", bankAccount, amount, ""],
    ["", "", "", "", "", "", counterpartAccounts.creditWhenBankDebit, "", amount],
  ];
}

function createBIDVPosVoucherRows(transactionDate, grossAmount) {
  const netAmount = Math.round(grossAmount / 1.1);
  const vatAmount = grossAmount - netAmount;

  return [
    [transactionDate, JOURNAL_BANK_DEBIT, POS_FEE_DESCRIPTION, "", "", "", BIDV_BANK_ACCOUNT, "", netAmount],
    ["", "", "", "", "", "", POS_EXPENSE_ACCOUNT, netAmount, ""],
    [transactionDate, JOURNAL_BANK_DEBIT, POS_FEE_VAT_DESCRIPTION, "", "", "", VAT_INPUT_ACCOUNT, vatAmount, ""],
    ["", "", "", "", "", "", BIDV_BANK_ACCOUNT, "", vatAmount],
  ];
}

function createDebitBankCreditRows(transactionDate, description, amount, debitAccount, bankAccount) {
  return [
    [transactionDate, JOURNAL_BANK_DEBIT, description, "", "", "", debitAccount, amount, ""],
    ["", "", "", "", "", "", bankAccount, "", amount],
  ];
}

function hasPOS(description) {
  return /\bPOS\b/i.test(String(description || ""));
}

function cellValue(sheet, rowIndex, columnIndex) {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  const cell = sheet[address];

  if (!cell) {
    return "";
  }

  return cell.v ?? cell.w ?? "";
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const text = String(value ?? "")
    .trim()
    .replace(/\s/g, "")
    .replace(/,/g, "");

  if (text === "" || text === "-") {
    return 0;
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeHeader(value) {
  return normalizeText(value).replace(/\s+/g, " ");
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .replace(/[().]/g, "")
    .replace(/\s*\/\s*/g, "/")
    .trim()
    .toLowerCase();
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDate(value);
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsedExcelDate = XLSX.SSF.parse_date_code(value);
    if (parsedExcelDate) {
      return `${parsedExcelDate.y}-${String(parsedExcelDate.m).padStart(2, "0")}-${String(parsedExcelDate.d).padStart(2, "0")}`;
    }
  }

  const text = String(value ?? "").trim();
  if (!text) {
    return "";
  }

  const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2].padStart(2, "0")}-${isoMatch[3].padStart(2, "0")}`;
  }

  const dmyMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmyMatch) {
    return `${dmyMatch[3]}-${dmyMatch[2].padStart(2, "0")}-${dmyMatch[1].padStart(2, "0")}`;
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return formatDate(parsed);
  }

  return "";
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function timestampForFileName() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return `${y}${m}${d}_${hh}${mm}${ss}`;
}

function saveOutputToDownloads(fileName, outputBuffer) {
  const downloadsDir = path.join(os.homedir(), "Downloads");
  fs.mkdirSync(downloadsDir, { recursive: true });
  const outputPath = path.join(downloadsDir, fileName);
  fs.writeFileSync(outputPath, outputBuffer);
  return outputPath;
}
