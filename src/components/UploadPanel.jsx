import { formatFileSize } from "../utils/excelUtils";

// Describes a sheet in the picker: row count, or why it cannot be used.
function describeSheet(sheet) {
  if (sheet.rowCount === 0) {
    return "empty";
  }

  if (!sheet.isUsable) {
    return `${sheet.rowCount} rows — missing ${sheet.missingColumns.length} columns`;
  }

  return `${sheet.rowCount} rows`;
}

export function UploadPanel({
  selectedFile,
  onFileSelect,
  sheets,
  selectedSheet,
  onSheetSelect,
  isReadingSheets,
  onGenerate,
  onReset,
  isProcessing,
  errorMessage,
  hasDownloads,
}) {
  function handleFileChange(event) {
    const [file] = event.target.files || [];
    onFileSelect(file || null);
    event.target.value = "";
  }

  const activeSheet = sheets.find((sheet) => sheet.name === selectedSheet);
  const canGenerate = Boolean(selectedFile) && !isProcessing && !isReadingSheets;

  return (
    <section className="card upload-card">
      <header>
        <h2>Upload Excel file</h2>
        <p> (.xlsx or .xls file.)</p>
      </header>

      <label className="file-picker" htmlFor="excel-file-input">
        <span>{selectedFile ? selectedFile.name : "Choose an Excel file"}</span>
        <small>
          {selectedFile
            ? formatFileSize(selectedFile.size)
            : "No file selected"}
        </small>
      </label>

      <input
        id="excel-file-input"
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileChange}
        disabled={isProcessing}
      />

      {isReadingSheets && <p className="status">Reading worksheets...</p>}

      {sheets.length > 0 && (
        <div className="sheet-picker">
          <label htmlFor="excel-sheet-input">Worksheet</label>
          <select
            id="excel-sheet-input"
            value={selectedSheet}
            onChange={(event) => onSheetSelect(event.target.value)}
            disabled={isProcessing}
          >
            {sheets.map((sheet) => (
              <option key={sheet.name} value={sheet.name}>
                {sheet.name} ({describeSheet(sheet)})
              </option>
            ))}
          </select>
          {activeSheet && !activeSheet.isUsable && (
            <small className="sheet-hint">
              This sheet cannot be converted. Missing:{" "}
              {activeSheet.missingColumns.join(", ")}.
            </small>
          )}
        </div>
      )}

      {errorMessage && <p className="status error">{errorMessage}</p>}

      <div className="actions">
        <button
          type="button"
          className="btn btn-primary"
          disabled={!canGenerate}
          onClick={() => onGenerate(selectedFile)}
        >
          {isProcessing ? "Generating..." : "Generate workbook"}
        </button>

        <button
          type="button"
          className="btn btn-secondary"
          disabled={isProcessing || (!selectedFile && !hasDownloads)}
          onClick={onReset}
        >
          Reset
        </button>
      </div>
    </section>
  );
}
