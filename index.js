const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");
const { getPrinters, print } = require("pdf-to-printer");
const { PDFDocument } = require("pdf-lib");
const PDFDocumentKit = require("pdfkit");
const sharp = require("sharp");

const app = express();
const PORT = process.env.PORT || 8080;

const IMG_EXTENSIONS = [
  ".jpg",
  ".jpeg",
  ".png",
  ".tiff",
  ".tif",
  ".bmp",
  ".gif",
  ".webp",
];

const DOC_EXTENSIONS = [
  ".docx",
  ".doc",
  ".odt",
  ".ott",
  ".rtf",
  ".txt",
  ".xlsx",
  ".xls",
  ".ods",
  ".pptx",
  ".ppt",
  ".odp",
];

const ALLOWED_EXTS = new Set([".pdf", ...IMG_EXTENSIONS, ...DOC_EXTENSIONS]);

app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`);
  next();
});

const upload = multer({
  dest: path.join(__dirname, "uploads"),
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ALLOWED_EXTS.has(ext)) {
      cb(null, true);
    } else {
      const err = new Error(
        `Unsupported file type: ${ext}. Allowed: PDF, Images, Office Docs.`,
      );
      err.status = 400;
      cb(err);
    }
  },
});

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const parseBool = (val) => {
  if (typeof val === "boolean") return val;
  if (!val) return undefined;
  const lower = String(val).toLowerCase();
  return lower === "true" || lower === "1" || lower === "on";
};

const parsePageRange = (rangeStr, maxPages) => {
  const pages = new Set();
  if (!rangeStr) return pages;
  const parts = rangeStr.split(",");
  for (const part of parts) {
    if (part.includes("-")) {
      const [start, end] = part.split("-").map((n) => parseInt(n.trim(), 10));
      if (!isNaN(start) && !isNaN(end)) {
        for (let i = start; i <= end; i++) {
          if (i <= maxPages) pages.add(i);
        }
      }
    } else {
      const p = parseInt(part.trim(), 10);
      if (!isNaN(p) && p <= maxPages) pages.add(p);
    }
  }
  return pages;
};

/**
 * Converts Office/Text docs to PDF using LibreOffice CLI
 */
function convertDocWithLibreOffice(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    let command = "soffice";
    if (process.platform === "win32") {
      const possiblePaths = [
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
      ];
      const foundPath = possiblePaths.find((p) => fs.existsSync(p));
      if (foundPath) command = `"${foundPath}"`;
    }

    const outputDir = path.dirname(outputPath);

    const cmdString = `${command} --headless --convert-to pdf --outdir "${outputDir}" "${inputPath}"`;

    exec(cmdString, (error, stdout, stderr) => {
      if (error) {
        return reject(
          new Error(
            `LibreOffice conversion failed. Is it installed? Details: ${stderr || error.message}`,
          ),
        );
      }

      const basename = path.basename(inputPath, path.extname(inputPath));
      const generatedFile = path.join(outputDir, `${basename}.pdf`);

      try {
        if (fs.existsSync(generatedFile)) {
          fs.renameSync(generatedFile, outputPath);
          resolve(outputPath);
        } else {
          reject(
            new Error("LibreOffice finished but output PDF was not found."),
          );
        }
      } catch (err) {
        reject(err);
      }
    });
  });
}

/**
 * Converts Images to PDF using Sharp & PDFKit
 */
async function convertImageToPdf(inputPath, outputPath) {
  const imageBuffer = await sharp(inputPath)
    .rotate()
    .toFormat("png")
    .toBuffer();
  const metadata = await sharp(imageBuffer).metadata();

  const A4_WIDTH = 595.28;
  const A4_HEIGHT = 841.89;

  let pdfWidth = A4_WIDTH;
  let pdfHeight = A4_HEIGHT;
  let layout = "portrait";

  if (metadata.width > metadata.height) {
    layout = "landscape";
    pdfWidth = A4_HEIGHT;
    pdfHeight = A4_WIDTH;
  }

  const doc = new PDFDocumentKit({ layout, size: "A4", margin: 0 });
  const writeStream = fs.createWriteStream(outputPath);
  doc.pipe(writeStream);

  doc.image(imageBuffer, 0, 0, {
    fit: [pdfWidth, pdfHeight],
    align: "center",
    valign: "center",
  });

  doc.end();

  return new Promise((resolve, reject) => {
    writeStream.on("finish", () => resolve(outputPath));
    writeStream.on("error", reject);
  });
}

/**
 * Main Conversion Coordinator
 * Returns path to a PDF file (either original or converted)
 */
async function ensurePdf(fileObj) {
  const ext = path.extname(fileObj.originalname).toLowerCase();

  if (ext === ".pdf") {
    return fileObj.path;
  }

  const correctExtPath = fileObj.path + ext;
  fs.renameSync(fileObj.path, correctExtPath);

  const outputPdfPath = fileObj.path + ".converted.pdf";

  if (DOC_EXTENSIONS.includes(ext)) {
    await convertDocWithLibreOffice(correctExtPath, outputPdfPath);

    fs.unlink(correctExtPath, () => {});
    return outputPdfPath;
  }

  if (IMG_EXTENSIONS.includes(ext)) {
    await convertImageToPdf(correctExtPath, outputPdfPath);
    fs.unlink(correctExtPath, () => {});
    return outputPdfPath;
  }

  throw new Error(`Unhandled file extension in processing: ${ext}`);
}

app.get("/printers", async (req, res) => {
  try {
    const printers = await getPrinters();
    res.status(200).json({ success: true, printers });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post("/print", upload.array("files"), async (req, res, next) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "No files uploaded." });
    }

    const params = { ...req.query, ...req.body };

    const baseOptions = {};
    if (params.printer) baseOptions.printer = params.printer;
    if (params.paperSize) baseOptions.paperSize = params.paperSize;
    if (params.subset) baseOptions.subset = params.subset;
    if (params.copies) baseOptions.copies = Number(params.copies);
    if (parseBool(params.landscape)) baseOptions.orientation = "landscape";
    if (parseBool(params.portrait)) baseOptions.orientation = "portrait";
    if (params.scale) baseOptions.scale = params.scale;
    if (params.side) baseOptions.side = params.side;

    const results = [];
    const errors = [];

    await Promise.all(
      req.files.map(async (file) => {
        let processPath = null;
        try {
          processPath = await ensurePdf(file);

          if (params.monorange) {
            const fileBuffer = fs.readFileSync(processPath);
            const pdfDoc = await PDFDocument.load(fileBuffer);
            const totalPages = pdfDoc.getPageCount();

            const monoSet = parsePageRange(params.monorange, totalPages);
            const colorSet = new Set();

            for (let i = 1; i <= totalPages; i++) {
              if (!monoSet.has(i)) colorSet.add(i);
            }

            const toPageString = (set) =>
              Array.from(set)
                .sort((a, b) => a - b)
                .join(",");
            const jobs = [];

            if (monoSet.size > 0) {
              jobs.push(
                print(processPath, {
                  ...baseOptions,
                  monochrome: true,
                  pages: toPageString(monoSet),
                }),
              );
            }

            if (colorSet.size > 0) {
              jobs.push(
                print(processPath, {
                  ...baseOptions,
                  monochrome: false,
                  pages: toPageString(colorSet),
                }),
              );
            }

            await Promise.all(jobs);
            results.push({
              file: file.originalname,
              status: "converted & split printed",
              details: `Mono: [${toPageString(monoSet)}], Color: [${toPageString(colorSet)}]`,
            });
          } else {
            const standardOptions = { ...baseOptions };
            if (params.pages) standardOptions.pages = String(params.pages);
            if (parseBool(params.monochrome)) standardOptions.monochrome = true;

            await print(processPath, standardOptions);
            results.push({
              file: file.originalname,
              status: "converted & sent to printer",
            });
          }
        } catch (err) {
          console.error(`Failed to process ${file.originalname}:`, err);
          errors.push({ file: file.originalname, error: err.message });
        } finally {
          if (processPath && fs.existsSync(processPath)) {
            fs.unlink(processPath, (unlinkErr) => {
              if (unlinkErr)
                console.error(
                  `Error deleting temp file ${processPath}:`,
                  unlinkErr,
                );
            });
          }
        }
      }),
    );

    if (errors.length > 0 && results.length === 0) {
      return res.status(500).json({ success: false, errors });
    } else if (errors.length > 0) {
      return res
        .status(207)
        .json({ success: true, message: "Partial success", results, errors });
    }

    res.status(200).json({
      success: true,
      message: `Successfully processed ${results.length} file(s).`,
      results,
    });
  } catch (err) {
    return next(err);
  }
});

if (!fs.existsSync(path.join(__dirname, "uploads"))) {
  fs.mkdirSync(path.join(__dirname, "uploads"));
}

app.use((err, req, res, next) => {
  console.error(`[${err.status || 500}] Error:`, err.message);
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({
    success: false,
    error: err.message || "An internal server error occurred.",
  });
});

app.listen(PORT, () => {
  console.log(`Print Server running on http://localhost:${PORT}`);
});
