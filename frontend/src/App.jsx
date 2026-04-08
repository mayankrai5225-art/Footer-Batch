import { useCallback, useMemo, useState } from "react";

const API_BASE_URL = "http://localhost:5000";
const ALLOWED_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
];
const MAX_FILES = 10;

function App() {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [studentName, setStudentName] = useState("");
  const [className, setClassName] = useState("");
  const [rollNo, setRollNo] = useState("");
  const [includePageNumbers, setIncludePageNumbers] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");

  const acceptedExtensionsLabel = useMemo(() => ".pdf, .docx", []);

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
    setErrorMessage("");
  }

  async function handleSubmit(event) {
    event.preventDefault();
    setErrorMessage("");

    if (!selectedFiles.length) {
      setErrorMessage(`Add at least one file (up to ${MAX_FILES}).`);
      return;
    }

    if (!studentName.trim() || !className.trim() || !rollNo.trim()) {
      setErrorMessage("Please enter Name, Class, and Roll No.");
      return;
    }

    const formData = new FormData();
    for (const file of selectedFiles) {
      formData.append("documents", file);
    }
    formData.append("name", studentName.trim());
    formData.append("className", className.trim());
    formData.append("rollNo", rollNo.trim());
    formData.append("includePageNumbers", String(includePageNumbers));

    try {
      setIsLoading(true);

      const response = await fetch(`${API_BASE_URL}/api/process-document`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let backendError = "Failed to process documents.";
        try {
          const errorPayload = await response.json();
          backendError = errorPayload.error || backendError;
        } catch {
          /* use fallback */
        }
        throw new Error(backendError);
      }

      const processedBlob = await response.blob();
      const contentDisposition = response.headers.get("Content-Disposition") || "";
      const fileNameMatch = contentDisposition.match(/filename="(.+)"/);
      const defaultName =
        selectedFiles.length > 1
          ? `documents-with-footer.zip`
          : `processed-${selectedFiles[0].name}`;
      const downloadedFileName = fileNameMatch?.[1] || defaultName;

      const downloadUrl = window.URL.createObjectURL(processedBlob);
      const anchor = document.createElement("a");
      anchor.href = downloadUrl;
      anchor.download = downloadedFileName;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.URL.revokeObjectURL(downloadUrl);
    } catch (error) {
      setErrorMessage(error.message || "Unexpected error occurred.");
    } finally {
      setIsLoading(false);
    }
  }

  return (
    <main className="page">
      <div className="layout">
        <header className="header">
          <p className="eyebrow">Footer batch</p>
          <h1>Add footer to documents</h1>
          <p className="lead">
            Up to {MAX_FILES} PDF or Word files. Same footer (Name · Class · Roll No) on every file.
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

              <p className="footer-preview">
                Layout: <strong>Name</strong> (left) · <strong>Class</strong> (center) ·{" "}
                <strong>Roll No</strong> (right) — Times New Roman, not italic
              </p>

              {errorMessage ? <p className="error-text">{errorMessage}</p> : null}

              <button type="submit" disabled={isLoading || selectedFiles.length === 0} className="submit-button">
                {isLoading ? "Processing…" : selectedFiles.length > 1 ? "Process & download ZIP" : "Process & download"}
              </button>
            </section>
          </form>
        </div>
      </div>
    </main>
  );
}

export default App;
