import * as XLSX from "xlsx";

// Target network tabs to create in the output workbook.
export const NETWORK_FILE_NAMES = ["2G", "3G", "4G"];
// Final output filename requested by the user.
export const OUTPUT_FILE_NAME = "2, 3, 4 G cell_naming.xlsx";

// Allowed upload extensions.
const EXCEL_EXTENSIONS = ["xlsx", "xls", "csv"];
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
  "M.Tilt",
  "TRX Number",
  "Tower Type",
  "Tower Height",
];

// Number of sectors per site. Per-sector source columns (such as `Azimuth`)
// hold one value per sector, and generated cells cycle through them in order.
const SECTOR_COUNT = 3;

// 2G naming suffixes: 1 through 7, skipping 4.
// Positions 0-2 are the first TRX group of sectors 1-3, positions 3-5 are the
// second TRX group of the same sectors.
const TWO_G_SUFFIXES = ["1", "2", "3", "5", "6", "7"];

// 3G naming suffixes: 1 through 6.
// Positions 0-2 are the first carrier of sectors 1-3, positions 3-5 the second.
const THREE_G_SUFFIXES = ["1", "2", "3", "4", "5", "6"];

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

// 4G suffix groups tied to a band column. When that band is not deployed on a
// site, its suffixes are dropped from that site's 4G rows.
const FOUR_G_SUFFIXES_L1800 = ["A1", "A2", "A3"];
const FOUR_G_SUFFIXES_L2100 = ["C1", "C2", "C3"];

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
    mTilt: resolveColumnName(columns, ["M.Tilt", "MTilt", "M Tilt", "M.tilt"]),
    l800bw: resolveColumnName(columns, ["L800 BW", "L800BW", "L800 Bw"]),
    l1800bw: resolveColumnName(columns, ["L1800 BW", "L1800BW", "L1800 Bw"]),
    l2100bw: resolveColumnName(columns, ["L2100 BW", "L2100BW", "L2100 Bw"]),
    twoGConfig: resolveColumnName(columns, ["2G Config", "2G Configuration"]),
    threeGConfig: resolveColumnName(columns, ["3G Config", "3G Configuration"]),
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
    ["mTilt", "M.Tilt"],
    ["towerType", "Tower Type"],
    ["towerHeight", "Tower Height"],
    ["l800bw", "L800 BW"],
    ["l1800bw", "L1800 BW"],
    ["l2100bw", "L2100 BW"],
    ["twoGConfig", "2G Config"],
    ["threeGConfig", "3G Config"],
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

// A band is treated as not deployed when its bandwidth cell is blank or `-`.
function isBandMissing(value) {
  const normalized = String(value ?? "").trim();
  return normalized === "" || normalized === "-";
}

// Picks the 4G suffixes for one site:
// - `L800 BW` of `30 MHz` adds B4/B5/B6.
// - Missing `L1800 BW` drops A1/A2/A3.
// - Missing `L2100 BW` drops C1/C2/C3.
function getFourGSuffixes(sourceRow, columnMap) {
  const baseSuffixes = hasThirtyMhzL800(sourceRow[columnMap.l800bw])
    ? FOUR_G_SUFFIXES_EXTENDED
    : FOUR_G_SUFFIXES_BASE;

  const excluded = new Set();

  if (isBandMissing(sourceRow[columnMap.l1800bw])) {
    FOUR_G_SUFFIXES_L1800.forEach((suffix) => excluded.add(suffix));
  }

  if (isBandMissing(sourceRow[columnMap.l2100bw])) {
    FOUR_G_SUFFIXES_L2100.forEach((suffix) => excluded.add(suffix));
  }

  return baseSuffixes.filter((suffix) => !excluded.has(suffix));
}

// Per-sector source values hold one entry per sector, separated by `/` or `,`
// (for example `115/185/255`). Returns null when the value is not a full set,
// so the caller can fall back to the raw cell value.
function splitSectorValues(value) {
  const parts = String(value ?? "")
    .split(/[/,]/)
    .map((part) => part.trim());

  return parts.length === SECTOR_COUNT ? parts : null;
}

// Picks the sector value matching a cell's position in its suffix list.
// Cells cycle through the sectors: positions 0,1,2 map to sectors 1,2,3, then
// position 3 starts over at sector 1.
function getSectorValue(value, sectorIndex) {
  const parts = splitSectorValues(value);

  if (!parts) {
    return value ?? "";
  }

  return parts[sectorIndex % SECTOR_COUNT];
}

// Parses a config cell into groups of per-sector values.
// Handles the shapes used in the source sheets:
//   `S5/5/5 0/0/0` and `6/6/6 0/0/0` -> [["5","5","5"], ["0","0","0"]]
//   `S555/000`                       -> [["5","5","5"], ["0","0","0"]]
//   `S222` and `S2/2/2`              -> [["2","2","2"]]
// Returns an empty list when the cell holds no usable config (blank or `-`).
function parseConfigGroups(value) {
  const cleaned = String(value ?? "")
    .trim()
    .replace(/^s\s*/i, "")
    .trim();

  if (cleaned === "" || cleaned === "-") {
    return [];
  }

  const parts = cleaned.split(/[\s/]+/).filter(Boolean);

  // Compact form: each part already holds one value per sector (`555`).
  if (parts.every((part) => /^\d{3}$/.test(part))) {
    return parts.map((part) => part.split(""));
  }

  // Expanded form: one value per part, grouped in sector-sized chunks.
  const groups = [];

  for (let index = 0; index < parts.length; index += SECTOR_COUNT) {
    const group = parts.slice(index, index + SECTOR_COUNT);

    if (group.length === SECTOR_COUNT) {
      groups.push(group);
    }
  }

  return groups;
}

// Builds the 2G cell list for one site from `2G Config`.
// Cells 1-3 take their TRX count from the first config group and cells 5-7 from
// the second, so a second group of zeros drops 5, 6 and 7.
function getTwoGCells(sourceRow, columnMap) {
  const groups = parseConfigGroups(sourceRow[columnMap.twoGConfig]);

  return TWO_G_SUFFIXES.map((suffix, index) => {
    const sectorIndex = index % SECTOR_COUNT;
    const group = groups[Math.floor(index / SECTOR_COUNT)] ?? [];

    return { suffix, sectorIndex, trxNumber: group[sectorIndex] ?? "" };
  }).filter((cell) => Number(cell.trxNumber) > 0);
}

// Builds the 3G cell list for one site from `3G Config`.
// The config holds the carrier count per sector, so `2/2/2` yields cells 1-6
// and `1/1/1` only cells 1-3. A sector with 0 carriers gets no cell at all.
function getThreeGCells(sourceRow, columnMap) {
  const [carriersPerSector = []] = parseConfigGroups(
    sourceRow[columnMap.threeGConfig],
  );

  return THREE_G_SUFFIXES.map((suffix, index) => {
    const sectorIndex = index % SECTOR_COUNT;
    const carrierNumber = Math.floor(index / SECTOR_COUNT) + 1;
    const sectorCarriers = carriersPerSector[sectorIndex] ?? "";

    return {
      suffix,
      sectorIndex,
      trxNumber: sectorCarriers,
      hasCarrier: Number(sectorCarriers) >= carrierNumber,
    };
  }).filter((cell) => cell.hasCarrier);
}

// Builds the 4G cell list for one site. 4G has no config column, so `TRX
// Number` stays empty and only the band rules apply.
function getFourGCells(sourceRow, columnMap) {
  return getFourGSuffixes(sourceRow, columnMap).map((suffix, index) => ({
    suffix,
    sectorIndex: index % SECTOR_COUNT,
    trxNumber: "",
  }));
}

// Returns the cells to generate for one site on one network.
function getNetworkCells(networkName, sourceRow, columnMap) {
  if (networkName === "2G") {
    return getTwoGCells(sourceRow, columnMap);
  }

  if (networkName === "3G") {
    return getThreeGCells(sourceRow, columnMap);
  }

  return getFourGCells(sourceRow, columnMap);
}

// Builds one output row.
// `cellName` is unique per duplicated row (with suffix).
// `siteName` is the modified base value without the duplication suffix.
// `sectorIndex` is the cell's position in its network suffix list, used to pick
// the matching per-sector values.
function createOutputRow(
  sourceRow,
  columnMap,
  cellName,
  siteName,
  sectorIndex,
  trxNumber,
) {
  const normalizedSiteName = String(siteName ?? "");

  return {
    "Cell Name": cellName,
    Region: sourceRow[columnMap.region] ?? "",
    "Site Name": normalizedSiteName,
    "Site Id": extractSiteIdFromSiteName(normalizedSiteName),
    Longitude: sourceRow[columnMap.longitude] ?? "",
    Latitude: sourceRow[columnMap.latitude] ?? "",
    Azimuth: getSectorValue(sourceRow[columnMap.azimuth], sectorIndex),
    "M.Tilt": getSectorValue(sourceRow[columnMap.mTilt], sectorIndex),
    "TRX Number": trxNumber ?? "",
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

      const cells = getNetworkCells(networkName, row, columnMap);

      cells.forEach(({ suffix, sectorIndex, trxNumber }) => {
        networkRows.push(
          createOutputRow(
            row,
            columnMap,
            `${baseCellName}-${suffix}`,
            baseCellName,
            sectorIndex,
            trxNumber,
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
