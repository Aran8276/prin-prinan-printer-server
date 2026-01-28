const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { exec, spawn } = require("child_process");
const { getPrinters, print } = require("pdf-to-printer");
const { PDFDocument } = require("pdf-lib");
const PDFDocumentKit = require("pdfkit");
const sharp = require("sharp");
const axios = require("axios");
const EventEmitter = require("events");

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

const activePrintJobs = new Map();

class PrintMonitor extends EventEmitter {
  constructor(interval = 1000) {
    super();
    this.interval = interval;
    this.psProcess = null;
    this.buffer = "";
  }

  start() {
    if (this.psProcess) return;

    const psScript = `
        $ErrorActionPreference = "SilentlyContinue"
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $previousState = @{}
        
        while ($true) {
            try {
                $jobs = Get-WmiObject Win32_PrintJob
                $currentIds = @()

                if ($jobs) {
                    if ($jobs -isnot [array]) { $jobs = @($jobs) }
                    foreach ($job in $jobs) {
                        $id = $job.JobId
                        $currentIds += $id
                        $jobData = @{
                            id = $id
                            document = $job.Document
                            printer = $job.Name
                            pagesPrinted = $job.PagesPrinted
                            status = if ($job.JobStatus) { $job.JobStatus } else { "Spooling" }
                        }

                        if (-not $previousState.ContainsKey($id)) {
                            $eventData = $jobData.Clone()
                            $eventData.event = "add"
                            Write-Output (ConvertTo-Json $eventData -Compress)
                        }
                        $previousState[$id] = $jobData
                    }
                }

                $prevKeys = @($previousState.Keys)
                foreach ($key in $prevKeys) {
                    if ($currentIds -notcontains $key) {
                        $removedJob = $previousState[$key]
                        $removedJob.event = "remove"
                        $removedJob.status = "Completed"
                        Write-Output (ConvertTo-Json $removedJob -Compress)
                        $previousState.Remove($key)
                    }
                }
            } catch {}
            Start-Sleep -Milliseconds ${this.interval}
        }
    `;

    this.psProcess = spawn("powershell.exe", ["-Command", "-"], {
      stdio: ["pipe", "pipe", "ignore"],
    });

    this.psProcess.stdin.write(psScript);
    this.psProcess.stdin.end();
    this.psProcess.stdout.setEncoding("utf8");

    this.psProcess.stdout.on("data", (data) => {
      this.buffer += data;
      const lines = this.buffer.split(/\r?\n/);
      this.buffer = lines.pop();

      for (const line of lines) {
        if (!line.trim()) continue;
        try {
          const job = JSON.parse(line);
          this.emit(job.event, job);
        } catch (e) {
          /* ignore */
        }
      }
    });
  }
}

const monitor = new PrintMonitor();
monitor.start();

monitor.on("add", (job) => {
  console.log(`[SPOOLER ADD] Job ${job.id}: ${job.document}`);

  for (const [filename, data] of activePrintJobs.entries()) {
    if (job.document.includes(filename)) {
      data.windowsJobId = job.id;

      activePrintJobs.set(String(job.id), { ...data, filename });

      console.log(`>> MATCHED Job ${job.id} to Laravel ID ${data.jobDetailId}`);

      sendWebhook(data.webhookUrl, {
        job_detail_id: data.jobDetailId,
        status: "running",
        message: `Spooling on ${job.printer}`,
      });
      break;
    }
  }
});

monitor.on("remove", (job) => {
  console.log(`[SPOOLER REMOVE] Job ${job.id} finished.`);

  const lookupKey = String(job.id);
  const data = activePrintJobs.get(lookupKey);

  if (data) {
    console.log(`>> REPORTING COMPLETION for Laravel ID ${data.jobDetailId}`);
    sendWebhook(data.webhookUrl, {
      job_detail_id: data.jobDetailId,
      status: "completed",
      message: "Print job finished successfully",
      pages_printed: job.pagesPrinted,
    });

    activePrintJobs.delete(lookupKey);
    if (data.filename) activePrintJobs.delete(data.filename);

    const filePath = path.join(__dirname, "uploads", data.filename);
    if (fs.existsSync(filePath)) {
      fs.unlink(filePath, () => {});
    }
  } else {
    console.log(
      `>> UNMATCHED Job ${job.id} removed (orphaned or external job)`,
    );
  }
});

async function sendWebhook(url, payload) {
  if (!url) return;
  try {
    await axios.post(url, payload);
    console.log(
      `[WEBHOOK SENT] ${payload.status} for Job ${payload.job_detail_id} to ${url}`,
    );
  } catch (err) {
    console.error(`[WEBHOOK FAILED] ${err.message}`);
  }
}

app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`);
  next();
});

const upload = multer({
  dest: path.join(__dirname, "uploads"),
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ALLOWED_EXTS.has(ext)) cb(null, true);
    else cb(new Error(`Unsupported file type: ${ext}`), false);
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
      if (error)
        return reject(
          new Error(`LibreOffice failed: ${stderr || error.message}`),
        );
      const basename = path.basename(inputPath, path.extname(inputPath));
      const generatedFile = path.join(outputDir, `${basename}.pdf`);

      setTimeout(() => {
        if (fs.existsSync(generatedFile)) {
          fs.renameSync(generatedFile, outputPath);
          resolve(outputPath);
        } else reject(new Error("LibreOffice output not found."));
      }, 500);
    });
  });
}

async function convertImageToPdf(inputPath, outputPath) {
  const imageBuffer = await sharp(inputPath)
    .rotate()
    .toFormat("png")
    .toBuffer();
  const metadata = await sharp(imageBuffer).metadata();
  const doc = new PDFDocumentKit({
    layout: metadata.width > metadata.height ? "landscape" : "portrait",
    size: "A4",
    margin: 0,
  });
  const writeStream = fs.createWriteStream(outputPath);
  doc.pipe(writeStream);
  doc.image(imageBuffer, 0, 0, {
    fit: [doc.page.width, doc.page.height],
    align: "center",
    valign: "center",
  });
  doc.end();
  return new Promise((resolve, reject) => {
    writeStream.on("finish", () => resolve(outputPath));
    writeStream.on("error", reject);
  });
}

async function ensurePdf(fileObj) {
  const ext = path.extname(fileObj.originalname).toLowerCase();
  if (ext === ".pdf") return fileObj.path;
  const correctExtPath = fileObj.path + ext;
  fs.renameSync(fileObj.path, correctExtPath);
  const outputPdfPath = fileObj.path + ".converted.pdf";
  if (DOC_EXTENSIONS.includes(ext)) {
    await convertDocWithLibreOffice(correctExtPath, outputPdfPath);
    if (fs.existsSync(correctExtPath)) fs.unlink(correctExtPath, () => {});
    return outputPdfPath;
  }
  if (IMG_EXTENSIONS.includes(ext)) {
    await convertImageToPdf(correctExtPath, outputPdfPath);
    if (fs.existsSync(correctExtPath)) fs.unlink(correctExtPath, () => {});
    return outputPdfPath;
  }
  return fileObj.path;
}

app.get("/printers", async (req, res) => {
  try {
    const printers = await getPrinters();
    res.status(200).json({ success: true, printers });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post("/count-pages", upload.single("file"), async (req, res, next) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, message: "No file uploaded." });
    }

    let processPath = null;
    try {
      processPath = await ensurePdf(req.file);

      const pdfBuffer = fs.readFileSync(processPath);

      const pdfDoc = await PDFDocument.load(pdfBuffer);
      const pages = pdfDoc.getPageCount();

      res.status(200).json({ success: true, pages });
    } catch (err) {
      res.status(500).json({ success: false, error: err.message });
    } finally {
      if (processPath && fs.existsSync(processPath)) {
        fs.unlink(processPath, () => {});
      }
    }
  } catch (err) {
    next(err);
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
    const webhookUrl = params.webhook_url || null;
    const jobDetailId = params.job_detail_id || null;

    console.log("Incoming Print Request:", {
      jobDetailId,
      webhookUrl,
      files: req.files.map((f) => f.filename),
    });

    const baseOptions = {};
    if (params.printer) baseOptions.printer = params.printer;
    if (params.paperSize) baseOptions.paperSize = params.paperSize;
    if (params.copies) baseOptions.copies = Number(params.copies);
    if (parseBool(params.landscape)) baseOptions.orientation = "landscape";
    if (parseBool(params.portrait)) baseOptions.orientation = "portrait";
    if (params.scale) baseOptions.scale = params.scale;
    if (params.side) baseOptions.side = params.side;
    if (params.pages) baseOptions.pages = String(params.pages);
    if (parseBool(params.monochrome)) baseOptions.monochrome = true;

    const results = [];

    for (const file of req.files) {
      let processPath = null;
      try {
        processPath = await ensurePdf(file);

        const actualFilename = path.basename(processPath);

        if (jobDetailId && webhookUrl) {
          console.log(
            `>> REGISTERING WATCH: ${actualFilename} -> ${jobDetailId}`,
          );
          activePrintJobs.set(actualFilename, {
            jobDetailId,
            webhookUrl,
            windowsJobId: null,
            filename: actualFilename,
          });
        }

        await print(processPath, baseOptions);

        results.push({
          file: file.originalname,
          status: "sent to spooler",
          internal_name: actualFilename,
        });
      } catch (err) {
        console.error(`Processing Error: ${err.message}`);
        if (webhookUrl && jobDetailId) {
          sendWebhook(webhookUrl, {
            job_detail_id: jobDetailId,
            status: "failed",
            message: err.message,
          });
        }
      }
    }

    res.status(200).json({ success: true, results });
  } catch (err) {
    next(err);
  }
});

if (!fs.existsSync(path.join(__dirname, "uploads"))) {
  fs.mkdirSync(path.join(__dirname, "uploads"));
}

app.listen(PORT, () => {
  console.log(`Print Server running on http://localhost:${PORT}`);
});
