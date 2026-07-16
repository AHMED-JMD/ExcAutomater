import { DownloadPanel } from "./components/DownloadPanel";
import { UploadPanel } from "./components/UploadPanel";
import { useExcelGenerator } from "./hooks/useExcelGenerator";
import "./App.css";

function App() {
  const {
    selectedFile,
    setSelectedFile,
    downloads,
    errorMessage,
    isProcessing,
    generateFiles,
    resetAll,
  } = useExcelGenerator();

  return (
    <main className="page">
      <header className="hero">
        <span className="badge">2g, 3g, 4g Excel Automation</span>
        <h1>2g, 3g, 4g cell naming Automater</h1>
        <p>create one workbook with 2G, 3G, and 4G sheets.</p>
      </header>

      <section className="grid">
        <UploadPanel
          selectedFile={selectedFile}
          onFileSelect={setSelectedFile}
          onGenerate={generateFiles}
          onReset={resetAll}
          isProcessing={isProcessing}
          errorMessage={errorMessage}
          hasDownloads={downloads.length > 0}
        />

        <DownloadPanel downloads={downloads} isProcessing={isProcessing} />
      </section>
    </main>
  );
}

export default App;
