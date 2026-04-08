const express = require("express");
const cors = require("cors");
const multer = require("multer");
const fs = require("fs/promises");
const fsSync = require("fs");
const path = require("path");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const JSZip = require("jszip");

const app = express();
const PORT = process.env.PORT || 5000;
const UPLOADS_DIR = path.join(__dirname, "uploads");

// Ensure temporary upload directory exists before multer tries to use it.
if (!fsSync.existsSync(UPLOADS_DIR)) {
  fsSync.mkdirSync(UPLOADS_DIR, { recursive: true });
}

// Store uploaded files in a temporary folder.
const upload = multer({
  dest: UPLOADS_DIR,
  limits: {
    // Keep upload size reasonable to avoid memory pressure.
    fileSize: 20 * 1024 * 1024,
  },
});

app.use(cors());
app.use(express.json());

/**
 * Adds footer text to every page in a PDF.
 */
async function processPdf(fileBuffer, footerParts, includePageNumbers) {
  const pdfDoc = await PDFDocument.load(fileBuffer);
  const pages = pdfDoc.getPages();
  const font = await pdfDoc.embedFont(StandardFonts.TimesRoman);
  const fontSize = 10;
  const margin = 24;

  pages.forEach((page, index) => {
    const { width } = page.getSize();
    const yPosition = 20;
    const leftText = `Name: ${footerParts.name}`;
    const centerText = `Class: ${footerParts.className}`;
    const rightBaseText = `Roll No: ${footerParts.rollNo}`;
    const pagePart = includePageNumbers ? ` | ${index + 1}/${pages.length}` : "";
    const rightText = `${rightBaseText}${pagePart}`;
    const centerTextWidth = font.widthOfTextAtSize(centerText, fontSize);
    const rightTextWidth = font.widthOfTextAtSize(rightText, fontSize);

    // Draw three footer blocks across full width: left, center, right.
    page.drawText(leftText, {
      x: margin,
      y: yPosition,
      size: fontSize,
      font,
      color: rgb(0.3, 0.3, 0.3),
    });

    page.drawText(centerText, {
      x: (width - centerTextWidth) / 2,
      y: yPosition,
      size: fontSize,
      font,
      color: rgb(0.3, 0.3, 0.3),
    });

    page.drawText(rightText, {
      x: Math.max(margin, width - rightTextWidth - margin),
      y: yPosition,
      size: fontSize,
      font,
      color: rgb(0.3, 0.3, 0.3),
    });
  });

  return pdfDoc.save();
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function makeDocxRunXml(text, extraInnerXml = "") {
  return `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:t${text.includes(" ") ? ' xml:space="preserve"' : ""}>${text}</w:t>${extraInnerXml}</w:r>`;
}

/**
 * Adds/updates a footer in the existing DOCX package while preserving original content.
 */
async function processDocx(fileBuffer, footerParts, includePageNumbers) {
  const zip = await JSZip.loadAsync(fileBuffer);

  const documentXmlPath = "word/document.xml";
  const relsXmlPath = "word/_rels/document.xml.rels";
  const contentTypesPath = "[Content_Types].xml";
  const footerFileName = "footer-custom.xml";
  const footerXmlPath = `word/${footerFileName}`;

  const documentXmlFile = zip.file(documentXmlPath);
  const relsXmlFile = zip.file(relsXmlPath);
  const contentTypesFile = zip.file(contentTypesPath);

  if (!documentXmlFile || !relsXmlFile || !contentTypesFile) {
    throw new Error("Invalid DOCX structure.");
  }

  let documentXml = await documentXmlFile.async("string");
  let relsXml = await relsXmlFile.async("string");
  let contentTypesXml = await contentTypesFile.async("string");

  const leftText = escapeXml(`Name: ${footerParts.name}`);
  const centerText = escapeXml(`Class: ${footerParts.className}`);
  const rightText = escapeXml(`Roll No: ${footerParts.rollNo}`);
  const pageNumberXml = includePageNumbers
    ? `
      ${makeDocxRunXml(" | ")}
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
      ${makeDocxRunXml("1")}
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
      ${makeDocxRunXml("/")}
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
      ${makeDocxRunXml("1")}
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i w:val="0"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>`
    : "";

  // Footer XML that Word can attach to document sections.
  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:tabs>
        <w:tab w:val="center" w:pos="4680"/>
        <w:tab w:val="right" w:pos="9360"/>
      </w:tabs>
    </w:pPr>
    ${makeDocxRunXml(leftText)}
    <w:r><w:tab/></w:r>
    ${makeDocxRunXml(centerText)}
    <w:r><w:tab/></w:r>
    ${makeDocxRunXml(rightText)}${pageNumberXml}
  </w:p>
</w:ftr>`;

  // 1) Add or update footer file in package.
  zip.file(footerXmlPath, footerXml);

  // 2) Ensure relationship exists from document.xml to footer XML.
  const existingFooterRel = relsXml.match(
    /<Relationship[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/footer"[^>]*Target="footer-custom\.xml"[^>]*Id="([^"]+)"[^>]*\/>|<Relationship[^>]*Id="([^"]+)"[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/footer"[^>]*Target="footer-custom\.xml"[^>]*\/>/i
  );

  let footerRelId = existingFooterRel?.[1] || existingFooterRel?.[2] || "rIdFooterCustom";
  if (!existingFooterRel) {
    relsXml = relsXml.replace(
      "</Relationships>",
      `<Relationship Id="${footerRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${footerFileName}"/></Relationships>`
    );
  }

  // 3) Ensure each section has a default footer reference.
  const addOrReplaceFooterReference = (sectPrXml) => {
    if (/<w:footerReference[^>]*w:type="default"[^>]*\/>/.test(sectPrXml)) {
      return sectPrXml.replace(
        /<w:footerReference[^>]*w:type="default"[^>]*\/>/,
        `<w:footerReference w:type="default" r:id="${footerRelId}"/>`
      );
    }
    return sectPrXml.replace(
      "</w:sectPr>",
      `<w:footerReference w:type="default" r:id="${footerRelId}"/></w:sectPr>`
    );
  };

  if (/<w:sectPr[\s\S]*?<\/w:sectPr>/.test(documentXml)) {
    documentXml = documentXml.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, (match) =>
      addOrReplaceFooterReference(match)
    );
  } else {
    documentXml = documentXml.replace(
      "</w:body>",
      `<w:sectPr><w:footerReference w:type="default" r:id="${footerRelId}"/></w:sectPr></w:body>`
    );
  }

  // 4) Ensure content type entry exists for footer XML.
  if (!contentTypesXml.includes('/word/footer-custom.xml')) {
    contentTypesXml = contentTypesXml.replace(
      "</Types>",
      '<Override PartName="/word/footer-custom.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
    );
  }

  zip.file(documentXmlPath, documentXml);
  zip.file(relsXmlPath, relsXml);
  zip.file(contentTypesPath, contentTypesXml);

  return zip.generateAsync({ type: "nodebuffer" });
}

app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

const MAX_BATCH = 10;
const PDF_MIME = "application/pdf";
const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function isAllowedMime(mimetype) {
  return mimetype === PDF_MIME || mimetype === DOCX_MIME;
}

/**
 * Reads one uploaded temp file, applies footer, returns output buffer + download name.
 */
async function processOneUploadedFile(uploadedFile, footerParts, includePageNumbers) {
  const fileBuffer = await fs.readFile(uploadedFile.path);
  const originalName = uploadedFile.originalname;
  const baseName = path.parse(originalName).name;

  if (uploadedFile.mimetype === PDF_MIME) {
    const outputPdf = await processPdf(fileBuffer, footerParts, includePageNumbers);
    return { buffer: Buffer.from(outputPdf), fileName: `${baseName}-with-footer.pdf` };
  }

  if (uploadedFile.mimetype === DOCX_MIME) {
    const outputDocx = await processDocx(fileBuffer, footerParts, includePageNumbers);
    return { buffer: outputDocx, fileName: `${baseName}-with-footer.docx` };
  }

  throw new Error(`Unsupported type for "${originalName}". Use PDF or DOCX only.`);
}

app.post("/api/process-document", upload.array("documents", MAX_BATCH), async (req, res) => {
  const uploadedFiles = req.files || [];

  try {
    if (!uploadedFiles.length) {
      return res.status(400).json({ error: "Please upload at least one PDF or DOCX file." });
    }

    const name = String(req.body.name || "").trim();
    const className = String(req.body.className || "").trim();
    const rollNo = String(req.body.rollNo || "").trim();
    const includePageNumbers = String(req.body.includePageNumbers || "false").toLowerCase() === "true";

    if (!name || !className || !rollNo) {
      for (const f of uploadedFiles) {
        await fs.unlink(f.path).catch(() => {});
      }
      return res.status(400).json({ error: "Name, Class, and Roll No are all required." });
    }

    const footerParts = { name, className, rollNo };

    // Validate all MIME types before processing (beginner-friendly: fail fast with a clear message).
    for (const f of uploadedFiles) {
      if (!isAllowedMime(f.mimetype)) {
        for (const file of uploadedFiles) {
          await fs.unlink(file.path).catch(() => {});
        }
        return res.status(400).json({
          error: `Invalid file type: "${f.originalname}". Only .pdf and .docx are allowed.`,
        });
      }
    }

    const outputs = [];
    for (const f of uploadedFiles) {
      try {
        outputs.push(await processOneUploadedFile(f, footerParts, includePageNumbers));
      } finally {
        await fs.unlink(f.path).catch(() => {});
      }
    }

    // One file → return that file directly (same as before). Several files → ZIP bundle.
    if (outputs.length === 1) {
      const { buffer, fileName } = outputs[0];
      const isPdf = fileName.toLowerCase().endsWith(".pdf");
      res.setHeader("Content-Type", isPdf ? PDF_MIME : DOCX_MIME);
      res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
      return res.send(buffer);
    }

    const zip = new JSZip();
    const usedNames = new Set();
    for (const { buffer, fileName } of outputs) {
      let uniqueName = fileName;
      let suffix = 1;
      const parsed = path.parse(fileName);
      while (usedNames.has(uniqueName)) {
        uniqueName = `${parsed.name}-${suffix}${parsed.ext}`;
        suffix += 1;
      }
      usedNames.add(uniqueName);
      zip.file(uniqueName, buffer);
    }

    const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });
    const zipFileName = `documents-with-footer-${Date.now()}.zip`;
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${zipFileName}"`);
    return res.send(zipBuffer);
  } catch (error) {
    console.error("Processing error:", error);
    for (const f of uploadedFiles) {
      await fs.unlink(f.path).catch(() => {});
    }
    return res.status(500).json({
      error: "Something went wrong while processing your documents.",
    });
  }
});

app.listen(PORT, () => {
  console.log(`Backend server running at http://localhost:${PORT}`);
});
