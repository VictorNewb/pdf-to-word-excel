
import * as pdfjsLib from
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

// DOM elements
const pdfInput = document.getElementById("pdfInput");
const dropZone = document.getElementById("dropZone");
const fileInfo = document.getElementById("fileInfo");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");
const convertBtn = document.getElementById("convertBtn");
const removeFile = document.getElementById("removeFile");
const progressSection = document.getElementById("progressSection");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");
const progressPercent = document.getElementById("progressPercent");
const successMessage = document.getElementById("successMessage");
const errorMessage = document.getElementById("errorMessage");
const downloadBtn = document.getElementById("downloadBtn");

let selectedFile = null;
let selectedFormat = "word";
let downloadBlob = null;
let downloadName = "";

// Helpers
function formatBytes(bytes) {
  return (bytes / (1024 * 1024)).toFixed(2) + " MB";
}

function setProgress(percent, message) {
  const value = Math.min(100, Math.max(0, percent));

  progressBar.style.width = value + "%";
  progressPercent.textContent = Math.round(value) + "%";
  progressText.textContent = message;
}

function showError(message) {
  errorMessage.textContent = message;
  errorMessage.classList.remove("hidden");
}

function resetMessages() {
  errorMessage.classList.add("hidden");
  successMessage.classList.add("hidden");
  progressSection.classList.add("hidden");
  downloadBlob = null;
}

function setFile(file) {
  resetMessages();

  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".pdf")) {
    showError("Please select a valid PDF file.");
    return;
  }

  if (file.size > 50 * 1024 * 1024) {
    showError("File is too large. Maximum allowed size is 50 MB.");
    return;
  }

  if (file.size === 0) {
    showError("The selected file is empty.");
    return;
  }

  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = formatBytes(file.size);

  fileInfo.classList.remove("hidden");
  dropZone.classList.add("hidden");
  convertBtn.disabled = false;
}

function clearFile() {
  selectedFile = null;
  pdfInput.value = "";

  fileInfo.classList.add("hidden");
  dropZone.classList.remove("hidden");
  convertBtn.disabled = true;

  resetMessages();
}

// File selection
pdfInput.addEventListener("change", event => {
  setFile(event.target.files[0]);
});

removeFile.addEventListener("click", clearFile);

// Drag and drop
["dragenter", "dragover"].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});

dropZone.addEventListener("drop", event => {
  const file = event.dataTransfer.files[0];
  setFile(file);
});

// Output format selection
document.querySelectorAll(".format-option").forEach(button => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".format-option").forEach(option => {
      option.classList.remove("active");
    });

    button.classList.add("active");
    selectedFormat = button.dataset.format;

    resetMessages();
  });
});

// Extract PDF text and organize it into lines and cells
async function extractPDF(file) {
  const buffer = await file.arrayBuffer();

  // Check the PDF file signature
  const signature = new TextDecoder().decode(buffer.slice(0, 5));

  if (signature !== "%PDF-") {
    throw new Error("This file does not contain a valid PDF document.");
  }

  const pdf = await pdfjsLib.getDocument({
    data: new Uint8Array(buffer)
  }).promise;

  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    setProgress(
      10 + (pageNumber / pdf.numPages) * 55,
      `Extracting page ${pageNumber} of ${pdf.numPages}...`
    );

    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();

    const items = content.items
      .filter(item => item.str && item.str.trim())
      .map(item => ({
        text: item.str.trim(),
        x: item.transform[4],
        y: item.transform[5],
        width: item.width || 0
      }));

    // Sort from top to bottom, then left to right
    items.sort((a, b) => {
      const verticalDifference = b.y - a.y;

      if (Math.abs(verticalDifference) > 3) {
        return verticalDifference;
      }

      return a.x - b.x;
    });

    const lines = [];

    // Group text items belonging to the same visual line
    for (const item of items) {
      let line = lines.find(
        existing => Math.abs(existing.y - item.y) <= 3
      );

      if (!line) {
        line = {
          y: item.y,
          items: []
        };

        lines.push(line);
      }

      line.items.push(item);
    }

    lines.sort((a, b) => b.y - a.y);

    const rows = lines.map(line => {
      line.items.sort((a, b) => a.x - b.x);

      const cells = [];
      let currentText = "";
      let previousEnd = null;

      for (const item of line.items) {
        const gap = previousEnd === null
          ? 0
          : item.x - previousEnd;

        // A large horizontal gap may indicate another column
        if (gap > 18 && currentText.trim()) {
          cells.push(currentText.trim());
          currentText = "";
        } else if (currentText) {
          currentText += " ";
        }

        currentText += item.text;
        previousEnd = item.x + item.width;
      }

      if (currentText.trim()) {
        cells.push(currentText.trim());
      }

      return cells;
    }).filter(row => row.length > 0);

    pages.push(rows);
  }

  if (!pages.some(page => page.length > 0)) {
    throw new Error(
      "No selectable text was found. This may be a scanned PDF. OCR support is not included in this version."
    );
  }

  return pages;
}

// Convert extracted content to Word
async function convertToWord(pages) {
  if (!window.docx) {
    throw new Error("Word generation library failed to load. Check your internet connection.");
  }

  const {
    Document,
    Paragraph,
    TextRun,
    Packer,
    PageBreak
  } = window.docx;

  const children = [];

  pages.forEach((page, pageIndex) => {
    if (pageIndex > 0) {
      children.push(
        new Paragraph({
          children: [new PageBreak()]
        })
      );
    }

    page.forEach(row => {
      const text = row.join("    ");

      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text,
              size: 22
            })
          ],
          spacing: {
            after: 140,
            line: 300
          }
        })
      );
    });
  });

  const document = new Document({
    creator: "PDFConvert",
    title: "Converted PDF Document",
    description: "Document converted using PDFConvert",
    sections: [{
      properties: {},
      children
    }]
  });

  return await Packer.toBlob(document);
}

// Convert extracted content to Excel
async function convertToExcel(pages) {
  if (!window.ExcelJS) {
    throw new Error("Excel generation library failed to load. Check your internet connection.");
  }

  const workbook = new ExcelJS.Workbook();

  workbook.creator = "PDFConvert";
  workbook.created = new Date();

  pages.forEach((page, pageIndex) => {
    const worksheet = workbook.addWorksheet(
      `Page ${pageIndex + 1}`
    );

    worksheet.columns = Array.from(
      { length: Math.max(1, ...page.map(row => row.length)) },
      (_, index) => ({
        key: `column${index + 1}`,
        width: 28
      })
    );

    page.forEach(row => {
      worksheet.addRow(row);
    });

    // Format cells
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.alignment = {
          vertical: "middle",
          wrapText: true
        };

        cell.font = {
          name: "Arial",
          size: 10
        };
      });
    });

    worksheet.views = [
      { state: "frozen", ySplit: 0 }
    ];
  });

  const buffer = await workbook.xlsx.writeBuffer();

  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

// Download generated file
function downloadFile(blob, name) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = name;

  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Main conversion process
convertBtn.addEventListener("click", async () => {
  if (!selectedFile) return;

  if (!window.docx || !window.ExcelJS) {
    showError("Conversion libraries are unavailable. Check your internet connection and reload the page.");
    return;
  }

  resetMessages();

  convertBtn.disabled = true;
  progressSection.classList.remove("hidden");

  try {
    setProgress(5, "Opening PDF document...");

    const pages = await extractPDF(selectedFile);

    setProgress(72, "Generating your document...");

    if (selectedFormat === "word") {
      downloadBlob = await convertToWord(pages);

      downloadName =
        selectedFile.name.replace(/\.pdf$/i, "") + ".docx";
    } else {
      downloadBlob = await convertToExcel(pages);

      downloadName =
        selectedFile.name.replace(/\.pdf$/i, "") + ".xlsx";
    }

    setProgress(100, "Conversion completed successfully!");

    successMessage.classList.remove("hidden");
  } catch (error) {
    console.error("Conversion error:", error);

    showError(
      error.message || "An unexpected error occurred during conversion."
    );

    progressSection.classList.add("hidden");
  } finally {
    convertBtn.disabled = !selectedFile;
  }
});

// Download button
downloadBtn.addEventListener("click", () => {
  if (downloadBlob) {
    downloadFile(downloadBlob, downloadName);
  }
});
