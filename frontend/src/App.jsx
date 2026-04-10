import { useCallback, useMemo, useState } from "react";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import JSZip from "jszip";

const ALLOWED_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
];
const MAX_FILES = 10;
const PDF_MIME = "application/pdf";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function formatFooterValue(value, fallback) {
  const trimmed = value.trim();
  return trimmed ? trimmed : fallback;
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
  const footerHeight = 60;

  pages.forEach((page, index) => {
    const { width, height } = page.getSize();
    const yPosition = 18;
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

    // Properly balanced white rectangle height (60) to cover old footers without deleting main text
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

  return zip.generateAsync({ type: "blob" });
}

function App() {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [studentName, setStudentName] = useState("");
  const [className, setClassName] = useState("");
  const [rollNo, setRollNo] = useState("");
  const [includePageNumbers, setIncludePageNumbers] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");

  const [processedResults, setProcessedResults] = useState([]);
  const [renamingId, setRenamingId] = useState(null);
  const [renamingValue, setRenamingValue] = useState("");

  const acceptedExtensionsLabel = useMemo(() => ".pdf, .docx", []);
  const footerPreview = useMemo(
    () => ({
      name: formatFooterValue(studentName, "Your name"),
      className: formatFooterValue(className, "Your class"),
      rollNo: formatFooterValue(rollNo, "Your roll no"),
      pageNumbers: includePageNumbers,
    }),
    [studentName, className, rollNo, includePageNumbers]
  );

  /** Keep only PDF/DOCX and cap at MAX_FILES; merge with existing list without duplicates by name+size. */
  const addFiles = useCallback((fileList) => {
    const incoming = Array.from(fileList || []);
    const valid = incoming.filter((f) => ALLOWED_TYPES.includes(f.type));

    let nextMsg = "";
    if (valid.length < incoming.length) {
      nextMsg = "Some files were skipped (only .pdf and .docx are allowed).";
    }

    setSelectedFiles((prev) => {
      const key = (f) => `${f.name}-${f.size}`;
      const seen = new Set(prev.map(key));
      const merged = [...prev];
      for (const f of valid) {
        if (merged.length >= MAX_FILES) {
          if (!nextMsg) nextMsg = `At most ${MAX_FILES} files. Extra files were not added.`;
          break;
        }
        const k = key(f);
        if (!seen.has(k)) {
          seen.add(k);
          merged.push(f);
        }
      }
      return merged;
    });
    setErrorMessage(nextMsg);
  }, []);

  function handleFileInputChange(event) {
    addFiles(event.target.files);
    event.target.value = "";
  }

  function handleDragOver(event) {
    event.preventDefault();
    setIsDragging(true);
  }

  function handleDragLeave(event) {
    event.preventDefault();
    setIsDragging(false);
  }

  function handleDrop(event) {
    event.preventDefault();
    setIsDragging(false);
    addFiles(event.dataTransfer.files);
  }

  function removeFile(index) {
    setSelectedFiles((prev) => prev.filter((_, i) => i !== index));
    setErrorMessage("");
  }

  function clearAllFiles() {
    setSelectedFiles([]);
    setProcessedResults([]);
    setErrorMessage("");
    setRenamingId(null);
  }

  function downloadBlob(blob, fileName) {
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => window.URL.revokeObjectURL(url), 100);
  }

  function startRenaming(id, currentName) {
    setRenamingId(id);
    setRenamingValue(currentName);
  }

  function saveRename(id) {
    if (!renamingValue.trim()) {
      setRenamingId(null);
      return;
    }
    
    setProcessedResults((prev) =>
      prev.map((res) => (res.id === id ? { ...res, fileName: renamingValue.trim() } : res))
    );
    setRenamingId(null);
  }

  async function handleSubmit(event) {
    event.preventDefault();
    setErrorMessage("");
    setProcessedResults([]);
    setRenamingId(null);

    if (!selectedFiles.length) {
      setErrorMessage(`Add at least one file (up to ${MAX_FILES}).`);
      return;
    }

    const trimmedName = studentName.trim();
    const trimmedClass = className.trim();
    const trimmedRoll = rollNo.trim();

    if (!trimmedName || !trimmedClass || !trimmedRoll) {
      setErrorMessage("Please enter Name, Class, and Roll No.");
      return;
    }

    const footerParts = {
      name: trimmedName,
      className: trimmedClass,
      rollNo: trimmedRoll,
    };

    try {
      setIsLoading(true);

      const outputs = [];
      
      // Process files in parallel for maximum speed
      await Promise.all(selectedFiles.map(async (file) => {
        const fileBuffer = await file.arrayBuffer();
        const baseName = file.name.replace(/\.[^.]+$/, "") || "document";

        if (file.type === PDF_MIME) {
          const processedPdfBytes = await processPdf(fileBuffer, footerParts, includePageNumbers);
          outputs.push({ 
            id: Math.random().toString(36).substr(2, 9),
            blob: new Blob([processedPdfBytes], { type: PDF_MIME }), 
            fileName: `${baseName}-with-footer.pdf` 
          });
        } else {
          const processedDocxBlob = await processDocx(fileBuffer, footerParts, includePageNumbers);
          outputs.push({ 
            id: Math.random().toString(36).substr(2, 9),
            blob: processedDocxBlob, 
            fileName: `${baseName}-with-footer.docx` 
          });
        }
      }));

      setProcessedResults(outputs);

      if (outputs.length === 1) {
        downloadBlob(outputs[0].blob, outputs[0].fileName);
      } else {
        const zip = new JSZip();
        const usedNames = new Set();

        for (const { blob, fileName } of outputs) {
          let uniqueName = fileName;
          let suffix = 1;
          const dotIndex = fileName.lastIndexOf(".");
          const namePart = fileName.substring(0, dotIndex);
          const extPart = fileName.substring(dotIndex);

          while (usedNames.has(uniqueName)) {
            uniqueName = `${namePart}-${suffix}${extPart}`;
            suffix += 1;
          }
          usedNames.add(uniqueName);
          zip.file(uniqueName, blob);
        }

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipFileName = `documents-with-footer-${Date.now()}.zip`;
        downloadBlob(zipBlob, zipFileName);
      }
    } catch (error) {
      console.error("Processing error:", error);
      setErrorMessage("Something went wrong while processing your documents.");
    } finally {
      setIsLoading(false);
    }
  }

  async function handleDownloadZip() {
    if (processedResults.length === 0) return;
    
    if (processedResults.length === 1) {
      downloadBlob(processedResults[0].blob, processedResults[0].fileName);
      return;
    }

    const zip = new JSZip();
    const usedNames = new Set();

    for (const { blob, fileName } of processedResults) {
      let uniqueName = fileName;
      let suffix = 1;
      const dotIndex = fileName.lastIndexOf(".");
      const namePart = fileName.substring(0, dotIndex);
      const extPart = fileName.substring(dotIndex);

      while (usedNames.has(uniqueName)) {
        uniqueName = `${namePart}-${suffix}${extPart}`;
        suffix += 1;
      }
      usedNames.add(uniqueName);
      zip.file(uniqueName, blob);
    }

    const zipBlob = await zip.generateAsync({ type: "blob" });
    const zipFileName = `documents-with-footer-${Date.now()}.zip`;
    downloadBlob(zipBlob, zipFileName);
  }

  return (
    <main className="page">
      <div className="layout">
        <header className="header">
          <p className="eyebrow">Footer batch</p>
          <h1>Add footer to documents</h1>
          <p className="lead">
            Up to {MAX_FILES} PDF or Word files. Same footer details on every file.
          </p>
        </header>

        <div className="shell">
          <form onSubmit={handleSubmit} className="form">
            <section className="panel panel-upload">
              <div className="panel-head">
                <h2>Files</h2>
                <span className="count">
                  {selectedFiles.length}/{MAX_FILES}
                </span>
              </div>

              <label
                className={`dropzone ${isDragging ? "drag-active" : ""}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                <input
                  type="file"
                  accept={acceptedExtensionsLabel}
                  multiple
                  onChange={handleFileInputChange}
                  className="file-input"
                  disabled={selectedFiles.length >= MAX_FILES}
                />
                <span className="dropzone-title">Drop files here or click to choose</span>
                <span className="dropzone-subtitle">
                  PDF & DOCX only · max {MAX_FILES} files
                </span>
              </label>

              {selectedFiles.length > 0 ? (
                <ul className="file-list">
                  {selectedFiles.map((file, index) => (
                    <li key={`${file.name}-${file.size}-${index}`} className="file-row">
                      <span className="file-name" title={file.name}>
                        {file.name}
                      </span>
                      <span className="file-meta">{(file.size / 1024).toFixed(1)} KB</span>
                      <button
                        type="button"
                        className="btn-ghost"
                        onClick={() => removeFile(index)}
                        aria-label={`Remove ${file.name}`}
                      >
                        Remove
                      </button>
                    </li>
                  ))}
                </ul>
              ) : (
                <p className="empty-hint">No files yet. Add up to {MAX_FILES}.</p>
              )}

              {selectedFiles.length > 0 ? (
                <button type="button" className="btn-text" onClick={clearAllFiles}>
                  Clear all
                </button>
              ) : null}
            </section>

            <section className="panel panel-details">
              <div className="panel-head">
                <h2>Footer details</h2>
              </div>

              <div className="fields-grid">
                <label className="field">
                  <span>Name</span>
                  <input
                    type="text"
                    placeholder="Full name"
                    value={studentName}
                    onChange={(e) => setStudentName(e.target.value)}
                    maxLength={80}
                  />
                </label>
                <label className="field">
                  <span>Class</span>
                  <input
                    type="text"
                    placeholder="e.g. 10-A"
                    value={className}
                    onChange={(e) => setClassName(e.target.value)}
                    maxLength={40}
                  />
                </label>
                <label className="field field-full">
                  <span>Roll No</span>
                  <input
                    type="text"
                    placeholder="Roll number"
                    value={rollNo}
                    onChange={(e) => setRollNo(e.target.value)}
                    maxLength={40}
                  />
                </label>
              </div>

              <label className="checkbox-row">
                <input
                  type="checkbox"
                  checked={includePageNumbers}
                  onChange={(e) => setIncludePageNumbers(e.target.checked)}
                />
                <span>Include page numbers (right side of footer)</span>
              </label>

              {errorMessage ? <p className="error-text">{errorMessage}</p> : null}

              <button type="submit" disabled={isLoading || selectedFiles.length === 0} className="submit-button">
                {isLoading ? "Processing…" : selectedFiles.length > 1 ? "Process & download ZIP" : "Process & download"}
              </button>
            </section>

            {processedResults.length > 0 && (
              <section className="panel panel-results animate-fade-in">
                <div className="panel-head">
                  <h2>Processed results</h2>
                  <span className="processed-count">{processedResults.length} files ready</span>
                </div>
                
                <ul className="file-list">
                  {processedResults.map((result) => (
                    <li key={result.id} className="file-row">
                      {renamingId === result.id ? (
                        <div className="rename-container">
                          <input
                            type="text"
                            className="rename-input"
                            value={renamingValue}
                            onChange={(e) => setRenamingValue(e.target.value)}
                            onKeyDown={(e) => e.key === "Enter" && saveRename(result.id)}
                            autoFocus
                          />
                          <button
                            type="button"
                            className="btn-save"
                            onClick={() => saveRename(result.id)}
                          >
                            Save
                          </button>
                        </div>
                      ) : (
                        <div className="file-info-group">
                          <span className="file-name" title={result.fileName}>
                            {result.fileName}
                          </span>
                          <button
                            type="button"
                            className="btn-icon-soft"
                            onClick={() => startRenaming(result.id, result.fileName)}
                            aria-label="Rename file"
                          >
                            ✏️
                          </button>
                        </div>
                      )}
                      <button
                        type="button"
                        className="btn-download"
                        onClick={() => downloadBlob(result.blob, result.fileName)}
                      >
                        Download
                      </button>
                    </li>
                  ))}
                </ul>

                {processedResults.length > 1 && (
                  <div className="results-actions">
                    <button type="button" className="btn-primary-small" onClick={handleDownloadZip}>
                      Download All as ZIP
                    </button>
                  </div>
                )}
                
                <button type="button" className="btn-text-dim" onClick={() => setProcessedResults([])}>
                  Close results
                </button>
              </section>
            )}

            <section className="panel panel-preview">
              <div className="panel-head">
                <h2>Live footer preview</h2>
                <span className="count">Updates as you type</span>
              </div>

              <div className="preview-card">
                <div className="preview-page">
                  <div className="preview-body">
                    <div className="preview-text-block">
                      <span className="preview-label">Document preview</span>
                      <p className="preview-copy">
                        Your selected files will receive the footer shown below.
                      </p>
                    </div>

                    <div className="preview-footer">
                      <div className="preview-footer-item">
                        <strong>{footerPreview.name}</strong>
                      </div>
                      <div className="preview-footer-item preview-footer-center">
                        <strong>{footerPreview.className}</strong>
                      </div>
                      <div className="preview-footer-item preview-footer-right">
                        <strong>
                           {footerPreview.rollNo}
                           {footerPreview.pageNumbers ? " | Page 1/5" : ""}
                        </strong>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </section>
          </form>
        </div>
      </div>
    </main>
  );
}

export default App;
