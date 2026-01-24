const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { getPrinters, print } = require("pdf-to-printer");

const app = express();
const PORT = process.env.PORT || 8080;

const upload = multer({
  dest: path.join(__dirname, "uploads"),
  fileFilter: (req, file, cb) => {
    if (file.mimetype === "application/pdf") {
      cb(null, true);
    } else {
      cb(new Error("Only PDF files are allowed"));
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

/**
 * GET /printers
 * Returns a list of available printers on the system.
 */
app.get("/printers", async (req, res) => {
  try {
    const printers = await getPrinters();
    res.status(200).json({ success: true, printers });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * POST /print
 * Accepts 'multipart/form-data'.
 * Field 'files': One or more PDF files.
 * Params (Body or URL): printer, copies, paperSize, monochrome, orientation, scale, etc.
 */
app.post("/print", upload.array("files"), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res
      .status(400)
      .json({ success: false, message: "No PDF files uploaded." });
  }

  const params = { ...req.query, ...req.body };

  const printOptions = {};

  if (params.printer) printOptions.printer = params.printer;
  if (params.paperSize) printOptions.paperSize = params.paperSize;
  if (params.subset) printOptions.subset = params.subset;

  if (params.copies) printOptions.copies = Number(params.copies);
  if (params.pages) printOptions.pages = String(params.pages);

  if (parseBool(params.monochrome)) printOptions.monochrome = true;
  if (parseBool(params.landscape)) printOptions.orientation = "landscape";
  if (parseBool(params.portrait)) printOptions.orientation = "portrait";
  if (params.scale) printOptions.scale = params.scale;
  if (params.side) printOptions.side = params.side;

  const results = [];
  const errors = [];

  await Promise.all(
    req.files.map(async (file) => {
      const filePath = file.path;
      try {
        await print(filePath, printOptions);
        results.push({ file: file.originalname, status: "sent to printer" });
      } catch (err) {
        console.error(`Failed to print ${file.originalname}:`, err);
        errors.push({ file: file.originalname, error: err.message });
      } finally {
        fs.unlink(filePath, (unlinkErr) => {
          if (unlinkErr) console.error("Error deleting temp file:", unlinkErr);
        });
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
    message: `Successfully sent ${results.length} job(s) to printer.`,
    options: printOptions,
    results,
  });
});

if (!fs.existsSync(path.join(__dirname, "uploads"))) {
  fs.mkdirSync(path.join(__dirname, "uploads"));
}

app.listen(PORT, () => {
  console.log(`Print Server running on http://localhost:${PORT}`);
});
