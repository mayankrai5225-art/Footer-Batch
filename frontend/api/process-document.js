import Busboy from "busboy";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import JSZip from "jszip";
import path from "path";

const MAX_BATCH = 10;
const PDF_MIME = "application/pdf";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function isAllowedMime(mimetype) {
  return mimetype === PDF_MIME || mimetype === DOCX_MIME;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function makeDocxRunPropertiesXml() {
  return '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:i w:val="0"/></w:rPr>';
}

function makeDocxRunXml(text, extraInnerXml = "") {
  return `<w:r>${makeDocxRunPropertiesXml()}<w:t${text.includes(" ") ? ' xml:space="preserve"' : ""}>${text}</w:t>${extraInnerXml}</w:r>`;
}

async function processPdf(fileBuffer, footerParts, includePageNumbers) {
  const pdfDoc = await PDFDocument.load(fileBuffer);
  const pages = pdfDoc.getPages();
  const font = await pdfDoc.embedFont(StandardFonts.TimesRoman);
  const fontSize = 12;
  const footerHeight = 76;

  pages.forEach((page, index) => {
    const { width, height } = page.getSize();
    const yPosition = 24;
    const leftText = footerParts.name;
    const centerText = footerParts.className;
    const rightBaseText = footerParts.rollNo;
    const pagePart = includePageNumbers ? ` | ${index + 1}/${pages.length}` : "";
    const rightText = `${rightBaseText}${pagePart}`;
    const footerBandWidth = width * 0.68;
    const footerBandStart = (width - footerBandWidth) / 2;
    const footerBandEnd = footerBandStart + footerBandWidth;
    const centerTextWidth = font.widthOfTextAtSize(centerText, fontSize);
    const rightTextWidth = font.widthOfTextAtSize(rightText, fontSize);

    page.drawRectangle({
      x: 0,
      y: 0,
      width,
      height: Math.min(footerHeight, height),
      color: rgb(1, 1, 1),
    });

    page.drawText(leftText, {
      x: footerBandStart,
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
      x: Math.max(footerBandStart, footerBandEnd - rightTextWidth),
      y: yPosition,
      size: fontSize,
      font,
      color: rgb(0.3, 0.3, 0.3),
    });
  });

  return pdfDoc.save();
}

async function processDocx(fileBuffer, footerParts, includePageNumbers) {
  const zip = await JSZip.loadAsync(fileBuffer);

  const documentXmlPath = "word/document.xml";
  const relsXmlPath = "word/_rels/document.xml.rels";
  const contentTypesPath = "[Content_Types].xml";
  const footerFileName = "footer-custom.xml";
  const footerXmlPath = `word/${footerFileName}`;
  const footerRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

  const documentXmlFile = zip.file(documentXmlPath);
  const relsXmlFile = zip.file(relsXmlPath);
  const contentTypesFile = zip.file(contentTypesPath);

  if (!documentXmlFile || !relsXmlFile || !contentTypesFile) {
    throw new Error("Invalid DOCX structure.");
  }

  let documentXml = await documentXmlFile.async("string");
  let relsXml = await relsXmlFile.async("string");
  let contentTypesXml = await contentTypesFile.async("string");

  const leftText = escapeXml(footerParts.name);
  const centerText = escapeXml(footerParts.className);
  const rightText = escapeXml(footerParts.rollNo);
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

  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:tabs>
        <w:tab w:val="center" w:pos="3600"/>
        <w:tab w:val="right" w:pos="7200"/>
      </w:tabs>
    </w:pPr>
    ${makeDocxRunXml(leftText)}
    <w:r><w:tab/></w:r>
    ${makeDocxRunXml(centerText)}
    <w:r><w:tab/></w:r>
    ${makeDocxRunXml(rightText)}${pageNumberXml}
  </w:p>
</w:ftr>`;

  zip.file(footerXmlPath, footerXml);

  const existingFooterRel = relsXml.match(
    new RegExp(
      `<Relationship[^>]*Type="${footerRelationshipType}"[^>]*Target="${footerFileName}"[^>]*Id="([^"]+)"[^>]*\/>|<Relationship[^>]*Id="([^"]+)"[^>]*Type="${footerRelationshipType}"[^>]*Target="${footerFileName}"[^>]*\/>`,
      "i"
    )
  );

  const footerRelId = existingFooterRel?.[1] || existingFooterRel?.[2] || "rIdFooterCustom";
  relsXml = relsXml.replace(
    /<Relationship[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/footer"[^>]*\/?>/gi,
    ""
  );
  relsXml = relsXml.replace(
    "</Relationships>",
    `<Relationship Id="${footerRelId}" Type="${footerRelationshipType}" Target="${footerFileName}"/></Relationships>`
  );

  const addOrReplaceFooterReference = (sectPrXml) => {
    const withoutExisting = sectPrXml.replace(/<w:footerReference[^>]*\/>/g, "");

    return withoutExisting.replace(
      "</w:sectPr>",
      `<w:footerReference w:type="default" r:id="${footerRelId}"/><w:footerReference w:type="first" r:id="${footerRelId}"/><w:footerReference w:type="even" r:id="${footerRelId}"/></w:sectPr>`
    );
  };

  if (/<w:sectPr[\s\S]*?<\/w:sectPr>/.test(documentXml)) {
    documentXml = documentXml.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, (match) =>
      addOrReplaceFooterReference(match)
    );
  } else {
    documentXml = documentXml.replace(
      "</w:body>",
      `<w:sectPr><w:footerReference w:type="default" r:id="${footerRelId}"/><w:footerReference w:type="first" r:id="${footerRelId}"/><w:footerReference w:type="even" r:id="${footerRelId}"/></w:sectPr></w:body>`
    );
  }

  contentTypesXml = contentTypesXml.replace(
    /<Override PartName="\/word\/footer[^\"]*\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.footer\+xml"\/>/gi,
    ""
  );
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

function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const busboy = Busboy({ headers: req.headers });
    const fields = {};
    const files = [];

    busboy.on("field", (name, value) => {
      fields[name] = value;
    });

    busboy.on("file", (name, file, info) => {
      const chunks = [];

      file.on("data", (chunk) => {
        chunks.push(chunk);
      });

      file.on("end", () => {
        files.push({
          fieldname: name,
          filename: info.filename,
          mimetype: info.mimeType,
          buffer: Buffer.concat(chunks),
        });
      });
    });

    busboy.on("error", reject);
    busboy.on("finish", () => resolve({ fields, files }));
    req.pipe(busboy);
  });
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed." });
  }

  try {
    const { fields, files } = await parseMultipart(req);

    if (!files.length) {
      return res.status(400).json({ error: "Please upload at least one PDF or DOCX file." });
    }

    if (files.length > MAX_BATCH) {
      return res.status(400).json({ error: `Add at most ${MAX_BATCH} files.` });
    }

    const includePageNumbers = String(fields.includePageNumbers || "false").toLowerCase() === "true";
    const footerParts = {
      name: String(fields.name || "").trim(),
      className: String(fields.className || "").trim(),
      rollNo: String(fields.rollNo || "").trim(),
    };

    if (!footerParts.name || !footerParts.className || !footerParts.rollNo) {
      return res.status(400).json({ error: "Name, Class, and Roll No are all required." });
    }

    for (const file of files) {
      if (!isAllowedMime(file.mimetype)) {
        return res.status(400).json({
          error: `Invalid file type: "${file.filename}". Only .pdf and .docx are allowed.`,
        });
      }
    }

    const outputs = [];
    for (const file of files) {
      const baseName = file.filename.replace(/\.[^.]+$/, "") || "document";

      if (file.mimetype === PDF_MIME) {
        const outputPdf = await processPdf(file.buffer, footerParts, includePageNumbers);
        outputs.push({ buffer: Buffer.from(outputPdf), fileName: `${baseName}-with-footer.pdf` });
      } else {
        const outputDocx = await processDocx(file.buffer, footerParts, includePageNumbers);
        outputs.push({ buffer: outputDocx, fileName: `${baseName}-with-footer.docx` });
      }
    }

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
    return res.status(500).json({ error: "Something went wrong while processing your documents." });
  }
}

export const config = {
  maxDuration: 60,
};