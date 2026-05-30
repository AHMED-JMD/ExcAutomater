import * as XLSX from "xlsx";

// Target network tabs to create in the output workbook.
export const NETWORK_FILE_NAMES = ["2G", "3G", "4G"];
// Final output filename requested by the user.
export const OUTPUT_FILE_NAME = "2, 3, 4 G cell_naming.xlsx";

// Allowed upload extensions.
const EXCEL_EXTENSIONS = ["xlsx", "xls"];
// Replacement value used for `Site Name` in each network tab.
const NETWORK_SITE_NAME_REPLACEMENT = {
  "2G": "G",
  "3G": "W",
  "4G": "L",
};

// Final output columns and order in each network sheet.
const OUTPUT_COLUMNS = [
  "Cell Name",
  "Region",
  "Site Name",
  "Site Id",
  "Longitude",
  "Latitude",
  "Azimuth",
  "Tower Type",
  "Tower Height",
];

// 4G naming suffix sets.
// - Base set for normal bandwidth rows.
// - Extended set when `L800 BW` contains `30 MHz`.
const FOUR_G_SUFFIXES_BASE = [
  "A1",
  "A2",
  "A3",
  "B1",
  "B2",
  "B3",
  "C1",
  "C2",
  "C3",
];
const FOUR_G_SUFFIXES_EXTENDED = [...FOUR_G_SUFFIXES_BASE, "B4", "B5", "B6"];

// Normalizes header names so we can match columns even if formatting differs
// (spaces, case, symbols).
function normalizeHeaderName(value = "") {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

// Finds the actual column name from the Excel sheet by checking a list of
// expected aliases (for example: "Site Code" or "SiteCode").
function resolveColumnName(columns, expectedNames) {
  const normalizedColumns = new Map(
    columns.map((columnName) => [normalizeHeaderName(columnName), columnName]),
  );

  for (const expectedName of expectedNames) {
    const matchedColumn = normalizedColumns.get(
      normalizeHeaderName(expectedName),
    );
    if (matchedColumn) {
      return matchedColumn;
    }
  }

  return "";
}

function buildColumnMap(columns) {
  // Maps source columns to logical fields, allowing minor naming variations.
  return {
    region: resolveColumnName(columns, ["Region"]),
    siteCode: resolveColumnName(columns, ["Site Code", "SiteCode"]),
    siteName: resolveColumnName(columns, ["Site Name", "SiteName"]),
    longitude: resolveColumnName(columns, ["Longitude", "LON", "Long", "Lon"]),
    latitude: resolveColumnName(columns, ["Latitude", "LAT", "Lat"]),
    azimuth: resolveColumnName(columns, ["Azimuth"]),
    towerType: resolveColumnName(columns, ["Tower Type", "TowerType"]),
    towerHeight: resolveColumnName(columns, [
      "Tower Height",
      "Tower Hight",
      "Tower height",
      "Tower hight",
    ]),
    l800bw: resolveColumnName(columns, ["L800 BW", "L800BW", "L800 Bw"]),
  };
}

function getMissingColumns(columnMap) {
  // Required input columns for rule-based generation.
  const required = [
    ["region", "Region"],
    ["siteCode", "Site Code"],
    ["siteName", "Site Name"],
    ["longitude", "Longitude/LON"],
    ["latitude", "Latitude/LAT"],
    ["azimuth", "Azimuth"],
    ["towerType", "Tower Type"],
    ["towerHeight", "Tower Height"],
    ["l800bw", "L800 BW"],
  ];

  return required
    .filter(([key]) => !columnMap[key])
    .map(([, readableName]) => readableName);
}

// Extracts the last 4-digit sequence from a site name (used as `Site Id`).
function extractSiteIdFromSiteName(siteName = "") {
  const matched = String(siteName).match(/(\d{4})(?!.*\d)/);
  return matched ? matched[1] : "";
}

// Builds the base name per network by replacing `GUL` with `G/W/L`.
// Example: `BNL-HillatAlbeer-GUL-3373` -> `BNL-HillatAlbeer-L-3373` for 4G.
function buildCellBaseName(siteName, networkName) {
  const replacementValue = NETWORK_SITE_NAME_REPLACEMENT[networkName] || "";
  return String(siteName).replace(/GUL/gi, replacementValue);
}

// 4G rule selector based on `L800 BW` content.
function hasThirtyMhzL800(value) {
  return String(value).toLowerCase().includes("30 mhz");
}

// Builds one output row.
// `cellName` is unique per duplicated row (with suffix).
// `siteName` is the modified base value without the duplication suffix.
function createOutputRow(sourceRow, columnMap, cellName, siteName) {
  const normalizedSiteName = String(siteName ?? "");

  return {
    "Cell Name": cellName,
    Region: sourceRow[columnMap.region] ?? "",
    "Site Name": normalizedSiteName,
    "Site Id": extractSiteIdFromSiteName(normalizedSiteName),
    Longitude: sourceRow[columnMap.longitude] ?? "",
    Latitude: sourceRow[columnMap.latitude] ?? "",
    Azimuth: sourceRow[columnMap.azimuth] ?? "",
    "Tower Type": sourceRow[columnMap.towerType] ?? "",
    "Tower Height": sourceRow[columnMap.towerHeight] ?? "",
  };
}

// Keeps only one row for each unique Region + Site Code combination.
// If duplicates exist, the first occurrence is kept.
function getUniqueRowsByRegionAndSiteCode(rows, columnMap) {
  const uniqueRows = [];
  const seen = new Set();

  rows.forEach((row) => {
    const region = String(row[columnMap.region] ?? "").trim();
    const siteCode = String(row[columnMap.siteCode] ?? "").trim();
    const key = `${region}::${siteCode}`;

    if (!seen.has(key)) {
      seen.add(key);
      uniqueRows.push(row);
    }
  });

  return uniqueRows;
}

// Returns extension without dot (e.g. "xlsx").
export function getFileExtension(fileName = "") {
  const parts = fileName.split(".");

  if (parts.length <= 1) {
    return "";
  }

  return parts.pop().toLowerCase();
}

// Validates the selected upload is an Excel file.
export function isExcelFile(file) {
  if (!file) {
    return false;
  }

  const extension = getFileExtension(file.name);
  return EXCEL_EXTENSIONS.includes(extension);
}

// Human-readable file size for UI.
export function formatFileSize(bytes = 0) {
  if (!bytes || Number.isNaN(bytes)) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let unitIndex = 0;
  let size = bytes;

  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }

  return `${size.toFixed(size >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

// Main transformation pipeline:
// 1) Read uploaded workbook.
// 2) Extract first sheet rows.
// 3) Validate required columns.
// 4) Deduplicate by Region + Site Code.
// 5) Create 2G/3G/4G sheets with Site Name replacement rules.
// 6) Return one downloadable workbook file.
export async function createWorkbookWithNetworkTabs(sourceFile) {
  const buffer = await sourceFile.arrayBuffer();
  const sourceWorkbook = XLSX.read(buffer, { type: "array" });
  const [firstSheetName] = sourceWorkbook.SheetNames;

  if (!firstSheetName) {
    throw new Error("No worksheet found in the uploaded file.");
  }

  const sourceSheet = sourceWorkbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sourceSheet, { raw: true, defval: "" });

  if (rows.length === 0) {
    throw new Error("The uploaded worksheet is empty.");
  }

  const columns = Object.keys(rows[0]);
  const columnMap = buildColumnMap(columns);
  const missingColumns = getMissingColumns(columnMap);

  if (missingColumns.length > 0) {
    throw new Error(`Missing required columns: ${missingColumns.join(", ")}.`);
  }

  const uniqueRows = getUniqueRowsByRegionAndSiteCode(rows, columnMap);

  const outputWorkbook = XLSX.utils.book_new();

  NETWORK_FILE_NAMES.forEach((networkName) => {
    const networkRows = [];

    uniqueRows.forEach((row) => {
      const siteName = String(row[columnMap.siteName] ?? "");
      const baseCellName = buildCellBaseName(siteName, networkName);

      if (networkName === "2G") {
        for (let counter = 1; counter <= 7; counter += 1) {
          networkRows.push(
            createOutputRow(
              row,
              columnMap,
              `${baseCellName}-${counter}`,
              baseCellName,
            ),
          );
        }

        return;
      }

      if (networkName === "3G") {
        for (let counter = 1; counter <= 6; counter += 1) {
          networkRows.push(
            createOutputRow(
              row,
              columnMap,
              `${baseCellName}-${counter}`,
              baseCellName,
            ),
          );
        }

        return;
      }

      const suffixes = hasThirtyMhzL800(row[columnMap.l800bw])
        ? FOUR_G_SUFFIXES_EXTENDED
        : FOUR_G_SUFFIXES_BASE;

      suffixes.forEach((suffix) => {
        networkRows.push(
          createOutputRow(
            row,
            columnMap,
            `${baseCellName}-${suffix}`,
            baseCellName,
          ),
        );
      });
    });

    const networkSheet = XLSX.utils.json_to_sheet(networkRows, {
      header: OUTPUT_COLUMNS,
    });
    XLSX.utils.book_append_sheet(outputWorkbook, networkSheet, networkName);
  });

  const outputArray = XLSX.write(outputWorkbook, {
    bookType: "xlsx",
    type: "array",
  });

  const blob = new Blob([outputArray], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return {
    id: "2-3-4-g-workbook",
    label: "2G, 3G, 4G",
    fileName: OUTPUT_FILE_NAME,
    size: blob.size,
    url: URL.createObjectURL(blob),
  };
}

// Releases generated object URLs to avoid browser memory leaks.
export function cleanupDownloadUrls(downloads = []) {
  downloads.forEach(({ url }) => {
    if (url) {
      URL.revokeObjectURL(url);
    }
  });
}
