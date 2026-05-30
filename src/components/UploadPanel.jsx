import { formatFileSize } from "../utils/excelUtils";

export function UploadPanel({
  selectedFile,
  onFileSelect,
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

      {errorMessage && <p className="status error">{errorMessage}</p>}

      <div className="actions">
        <button
          type="button"
          className="btn btn-primary"
          disabled={!selectedFile || isProcessing}
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
