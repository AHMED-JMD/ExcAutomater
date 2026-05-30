import { useEffect, useState } from "react";
import {
  cleanupDownloadUrls,
  createWorkbookWithNetworkTabs,
  isExcelFile,
} from "../utils/excelUtils";

const PROCESSING_DELAY_MS = 900;

function wait(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

export function useExcelGenerator() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [downloads, setDownloads] = useState([]);
  const [errorMessage, setErrorMessage] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);

  function clearDownloads() {
    setDownloads((currentDownloads) => {
      cleanupDownloadUrls(currentDownloads);
      return [];
    });
  }

  function resetAll() {
    setSelectedFile(null);
    setErrorMessage("");
    setIsProcessing(false);
    clearDownloads();
  }

  async function generateFiles(file) {
    if (!file) {
      setErrorMessage("Please choose an Excel file first.");
      return;
    }

    if (!isExcelFile(file)) {
      setErrorMessage("Only Excel files are supported (.xlsx or .xls).");
      return;
    }

    clearDownloads();
    setErrorMessage("");
    setIsProcessing(true);

    try {
      await wait(PROCESSING_DELAY_MS);
      const generatedFile = await createWorkbookWithNetworkTabs(file);
      setDownloads([generatedFile]);
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "Something went wrong while generating files.",
      );
    } finally {
      setIsProcessing(false);
    }
  }

  useEffect(() => {
    return () => {
      cleanupDownloadUrls(downloads);
    };
  }, [downloads]);

  return {
    selectedFile,
    setSelectedFile,
    downloads,
    errorMessage,
    isProcessing,
    generateFiles,
    resetAll,
  };
}
