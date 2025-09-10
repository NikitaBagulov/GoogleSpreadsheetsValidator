// === КОНСТАНТЫ ===
const SCHEMA_FIELDS = [
  "DOI", "Year", "InputData", "FeatureExtractionEngineering", "PreprocessingTransformation",
  "Dataset", "DatasetTimeSpan", "DatasetQuality", "MLProblem", "ScientificTask",
  "MLTechnique", "Comment", "Shortcomings", "Benchmarks"
];

const SUBPOINTS = {
  "InputData": ["Name:", "InstrumentSource:", "Type:", "InputFeatures:"],
  "Dataset": ["Format:", "MarkupLabeling:"],
  "DatasetTimeSpan": ["Training:", "Validation:", "Test:"],
  "DatasetQuality": ["Sufficiency:", "Balance:", "Representativeness:"],
  "MLProblem": ["MLTaskClass:", "MLTaskSubclass:"],
  "MLTechnique": ["ModelType:", "Architecture:", "TrainingDetails:"]
};

const ALLOWED_SHEETS = ["Bagulov", "Popov"]; // ← ЗАМЕНИТЕ НА СВОИ ЛИСТЫ

const DATE_RANGE_REGEX = /^\d{4}\.\d{2}\.\d{2}–\d{4}\.\d{2}\.\d{2}$/;

const COLORS = {
  VALID: "#d4edda",    // ✅ светло-зеленый
  INVALID: "#f8d7da",  // ❌ светло-красный
  PROCESSING: "#fff3cd" // 🟡 светло-желтый (индикация)
};

// === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

function isValidDateRange(text) {
  return typeof text === "string" && DATE_RANGE_REGEX.test(text.trim());
}

function normalizeText(text) {
  return (text || "").toString().trim();
}

function createFlexibleRegex(label) {
  const base = label.replace(":", "").trim();
  const escaped = base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`^(?:${escaped}\\s*:\\s*)`, "i");
}

function validateSubpointsDetailed(actual, subpointList) {
  const parts = actual.split(/[,;]/).map(s => s.trim());
  const details = {};
  let allValid = true;

  for (let sub of subpointList) {
    const baseName = sub.replace(":", "").trim();
    const regexPattern = `(?:^|\\s)${baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*:\\s*`;
    const regex = new RegExp(regexPattern, "i");

    let found = false;
    for (let part of parts) {
      if (regex.test(part)) {
        found = true;
        break;
      }
    }

    details[sub] = found;
    if (!found) allValid = false;
  }

  return { isValid: allValid, details: details };
}

function validateDateSpan(actual) {
  const result = validateDateSpanDetailed(actual);
  return result.isValid;
}

function validateDateSpanDetailed(actual) {
  const parts = actual.split(/[,;]/).map(s => s.trim());
  const details = {};
  let allValid = true;

  for (let label of SUBPOINTS["DatasetTimeSpan"]) {
    let foundValid = false;

    for (let part of parts) {
      const regex = createFlexibleRegex(label);
      if (regex.test(part)) {
        const datePart = part.replace(regex, "").trim();
        const dateMatch = datePart.match(/^(\d{4}\.\d{2}\.\d{2}–\d{4}\.\d{2}\.\d{2})/);
        if (dateMatch && dateMatch[1]) {
          foundValid = true;
          break;
        }
      }
    }

    details[label] = foundValid;
    if (!foundValid) allValid = false;
  }

  return { isValid: allValid, details: details };
}

function generateDetailedReport(headerValues, subHeaderValues, headerErrors, subHeaderDetails) {
  let report = "";

  for (let i = 0; i < SCHEMA_FIELDS.length; i++) {
    const field = SCHEMA_FIELDS[i];
    const hStatus = headerErrors[i] ? "❌" : "✅";
    report += `${field}: ${hStatus}\n`;

    if (subHeaderDetails && subHeaderDetails[field]) {
      const { isValid, subDetails } = subHeaderDetails[field];

      if (SUBPOINTS[field]) {
        for (let sub of SUBPOINTS[field]) {
          const subStatus = subDetails[sub] ? "✅" : "❌";
          report += `  ${sub} ${subStatus}\n`;
        }
      }
    }
  }

  return report.trim();
}

// === ОСНОВНАЯ ФУНКЦИЯ ВАЛИДАЦИИ БЛОКА ===

function validateBlock(sheet, startRow) {
  const headerRow = startRow;
  const subHeaderRow = startRow + 1;

  if (subHeaderRow > sheet.getLastRow()) return;

  // 🟡 Визуальная индикация "проверяется"
  const range = sheet.getRange(headerRow, 1, 2, 14);
  range.setBackground(COLORS.PROCESSING);

  // 📥 Получаем все значения за один вызов
  const headerValues = sheet.getRange(headerRow, 1, 1, SCHEMA_FIELDS.length).getValues()[0];
  const subHeaderValues = sheet.getRange(subHeaderRow, 1, 1, SCHEMA_FIELDS.length).getValues()[0];

  let totalErrors = 0;
  const headerColors = [];
  const subHeaderColors = [];
  const headerErrors = [];
  let subHeaderDetails = {}; // ← храним детали для отчёта

  // 🔍 Проверка заголовков (строка 1)
  for (let i = 0; i < SCHEMA_FIELDS.length; i++) {
    const actual = normalizeText(headerValues[i]);
    const expected = SCHEMA_FIELDS[i];
    const isValid = actual === expected;
    headerErrors.push(!isValid);
    const color = isValid ? COLORS.VALID : COLORS.INVALID;
    headerColors.push(color);
    if (!isValid) totalErrors++;
  }

  // 🔍 Проверка подзаголовков (строка 2)
  for (let i = 0; i < SCHEMA_FIELDS.length; i++) {
    const fieldName = SCHEMA_FIELDS[i];
    const actual = normalizeText(subHeaderValues[i]);
    let isValid = true;
    let subDetails = {};

    if (fieldName === "DatasetTimeSpan") {
    const result = validateDateSpanDetailed(actual);
    isValid = result.isValid;
    subDetails = result.details;
  } else if (SUBPOINTS[fieldName]) {
    const result = validateSubpointsDetailed(actual, SUBPOINTS[fieldName]);
    isValid = result.isValid;
    subDetails = result.details;
  }

    subHeaderDetails[fieldName] = { isValid, subDetails };
    const color = isValid ? COLORS.VALID : COLORS.INVALID;
    subHeaderColors.push(color);
    if (!isValid) totalErrors++;
  }

  // 🎨 Применяем цвета массово
  sheet.getRange(headerRow, 1, 1, SCHEMA_FIELDS.length).setBackgrounds([headerColors]);
  sheet.getRange(subHeaderRow, 1, 1, SCHEMA_FIELDS.length).setBackgrounds([subHeaderColors]);

  // 📊 Краткий статус в O1, O3, O5...
  const resultCell = sheet.getRange(headerRow, 15);
  const newResult = totalErrors === 0 ? "✅" : "❌";
  const currentResult = resultCell.getValue();
  const resultColor = totalErrors === 0 ? COLORS.VALID : COLORS.INVALID;

  if (currentResult !== newResult) {
    resultCell.setValue(newResult)
      .setBackground(resultColor)
      .setHorizontalAlignment("center")
      .setFontWeight("bold");
  } else {
    resultCell.setBackground(resultColor);
  }

  // 📋 Развёрнутый отчёт в O2, O4, O6...
  const detailCell = sheet.getRange(subHeaderRow, 15);
  const detailReport = generateDetailedReport(headerValues, subHeaderValues, headerErrors, subHeaderDetails);

  detailCell
    .setValue(detailReport)
    .setBackground(resultColor)
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment("top");

  // 📏 Подстраиваем высоту строки под отчёт
  sheet.setRowHeight(subHeaderRow, 150);
}

// === ТРИГГЕРЫ ===

function onEdit(e) {
  if (!e?.range) return;
  const sheet = e.source.getActiveSheet();
  if (!ALLOWED_SHEETS.includes(sheet.getName())) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (col < 1 || col > 14) return; // Только A-N

  const blockStartRow = row % 2 === 1 ? row : row - 1;
  if (blockStartRow >= 1) {
    validateBlock(sheet, blockStartRow);
  }
}

function onChange(e) {
  if (!e?.source) return;
  const sheet = e.source.getActiveSheet();
  if (!ALLOWED_SHEETS.includes(sheet.getName())) return;

  if (["INSERT_ROW", "INSERT_GRID", "EDIT"].includes(e.changeType)) {
    validateAllBlocksOnSheet(sheet);
  }
}

function validateAllBlocksOnSheet(sheet) {
  const lastRow = sheet.getLastRow();
  for (let startRow = 1; startRow <= lastRow; startRow += 2) {
    sheet.getRange(startRow, 1, 2, 14).setBackground(COLORS.PROCESSING);
    Utilities.sleep(100); // визуальная задержка
    validateBlock(sheet, startRow);
  }
}

function validateAllBlocks() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!ALLOWED_SHEETS.includes(sheet.getName())) {
    console.log("⚠️ Валидация отключена для листа '" + sheet.getName() + "'");
    return;
  }
  validateAllBlocksOnSheet(sheet);
  console.log("✅ Полная валидация завершена на листе: " + sheet.getName());
}