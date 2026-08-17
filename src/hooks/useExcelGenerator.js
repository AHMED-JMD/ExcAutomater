import { useEffect, useState } from "react";
import {
  cleanupDownloadUrls,
  createWorkbookWithNetworkTabs,
  isExcelFile,
  listWorkbookSheets,
} from "../utils/excelUtils";

const PROCESSING_DELAY_MS = 900;

function wait(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

export function useExcelGenerator() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [isReadingSheets, setIsReadingSheets] = useState(false);
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
    setSheets([]);
    setSelectedSheet("");
    setErrorMessage("");
    setIsProcessing(false);
    clearDownloads();
  }

  // Reads the workbook as soon as a file is chosen so the sheet list is ready
  // before the user presses generate. Preselects the first sheet that has all
  // the required columns, which is the right one in most workbooks.
  async function selectFile(file) {
    clearDownloads();
    setErrorMessage("");
    setSelectedFile(file);
    setSheets([]);
    setSelectedSheet("");

    if (!file) {
      return;
    }

    if (!isExcelFile(file)) {
      setErrorMessage("Only Excel files are supported (.xlsx or .xls).");
      return;
    }

    setIsReadingSheets(true);

    try {
      const workbookSheets = await listWorkbookSheets(file);
      const firstUsableSheet = workbookSheets.find((sheet) => sheet.isUsable);

      setSheets(workbookSheets);
      setSelectedSheet((firstUsableSheet ?? workbookSheets[0])?.name ?? "");

      if (!firstUsableSheet) {
        setErrorMessage(
          "No worksheet in this file has the required columns. Pick a sheet to see what it is missing.",
        );
      }
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "Could not read the worksheets in this file.",
      );
    } finally {
      setIsReadingSheets(false);
    }
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
      const generatedFile = await createWorkbookWithNetworkTabs(
        file,
        selectedSheet,
      );
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
    selectFile,
    sheets,
    selectedSheet,
    setSelectedSheet,
    isReadingSheets,
    downloads,
    errorMessage,
    isProcessing,
    generateFiles,
    resetAll,
  };
}
