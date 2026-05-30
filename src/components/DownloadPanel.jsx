import { formatFileSize } from "../utils/excelUtils";

export function DownloadPanel({ downloads, isProcessing }) {
  return (
    <section className="card downloads-card">
      <header>
        <h2>Generated workbook</h2>
      </header>

      {isProcessing && <p className="status info">Processing your file...</p>}

      {!isProcessing && downloads.length === 0 && (
        <p className="status muted">
          Your generated workbook will appear here after processing.
        </p>
      )}

      {downloads.length > 0 && (
        <ul className="download-list">
          {downloads.map((file) => (
            <li key={file.id} className="download-item">
              <div>
                <strong>{file.fileName}</strong>
                <small>{formatFileSize(file.size)}</small>
              </div>
              <a
                className="btn btn-download"
                href={file.url}
                download={file.fileName}
              >
                Download
              </a>
            </li>
          ))}
        </ul>
      )}
    </section>
  );
}
