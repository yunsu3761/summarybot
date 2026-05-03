import { useState, useRef, useCallback } from 'react';
import { uploadPdfs, scanFolder, uploadPatentExcel, deleteDocument, DocumentInfo } from '../api';

interface DocumentInputProps {
  documents: DocumentInfo[];
  onDocumentsChange: (docs: DocumentInfo[]) => void;
  summaryDone?: boolean;   // 요약 완료 상태 여부
}

export default function DocumentInput({ documents, onDocumentsChange, summaryDone = false }: DocumentInputProps) {
  const [activeTab, setActiveTab] = useState(0);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ type: string; text: string } | null>(null);
  const [folderPath, setFolderPath] = useState('./input_pdfs');
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);

  const showMsg = (type: string, text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage(null), 5000);
  };

  // --- PDF Upload ---
  const handlePdfUpload = useCallback(async (files: FileList) => {
    setLoading(true);
    const shouldReset = summaryDone;   // 요약 완료 후 새 업로드 → 초기화
    try {
      const res = await uploadPdfs(files, shouldReset);
      if (res.documents) {
        // reset이면 새 문서만, 아니면 기존에 추가
        onDocumentsChange(shouldReset ? res.documents : [...documents, ...res.documents]);
        const msg = shouldReset
          ? `기존 문서 초기화 후 ${res.uploaded}개 PDF 업로드 완료`
          : `${res.uploaded}개 PDF 업로드 완료`;
        showMsg('success', msg);
      }
    } catch {
      showMsg('error', 'PDF 업로드 실패');
    }
    setLoading(false);
  }, [documents, onDocumentsChange, summaryDone]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files.length > 0) {
      handlePdfUpload(e.dataTransfer.files);
    }
  }, [handlePdfUpload]);

  // --- Folder Scan ---
  const handleScanFolder = async () => {
    if (!folderPath.trim()) return;
    setLoading(true);
    const shouldReset = summaryDone;
    try {
      const res = await scanFolder(folderPath.trim(), shouldReset);
      if (res.error) {
        showMsg('error', res.error);
      } else {
        onDocumentsChange(shouldReset ? (res.documents || []) : [...documents, ...(res.documents || [])]);
        const msg = shouldReset
          ? `기존 문서 초기화 후 ${res.scanned}개 PDF 스캔 완료`
          : `${res.scanned}개 PDF 스캔 완료`;
        showMsg('success', msg);
      }
    } catch {
      showMsg('error', '폴더 스캔 실패');
    }
    setLoading(false);
  };

  // --- Patent Excel ---
  const handlePatentExcel = async (file: File) => {
    setLoading(true);
    const shouldReset = summaryDone;
    try {
      const res = await uploadPatentExcel(file, shouldReset);
      if (res.error) {
        showMsg('error', res.error);
      } else {
        onDocumentsChange(shouldReset ? (res.documents || []) : [...documents, ...(res.documents || [])]);
        const msg = shouldReset
          ? `기존 문서 초기화 후 ${res.processed}개 특허 URL 처리 완료`
          : `${res.processed}개 특허 URL 처리 완료`;
        showMsg('success', msg);
        if (res.errors?.length) {
          showMsg('warning', `${res.errors.length}개 URL 처리 실패`);
        }
      }
    } catch {
      showMsg('error', '엑셀 업로드 실패');
    }
    setLoading(false);
  };

  // --- Remove document ---
  const handleRemove = async (idx: number) => {
    try {
      await deleteDocument(idx);
      const newDocs = documents.filter((_, i) => i !== idx);
      onDocumentsChange(newDocs);
    } catch { /* ignore */ }
  };

  const tabs = ['📄 PDF 업로드', '📂 폴더 스캔', '📋 특허 엑셀'];

  return (
    <div className="card fade-in">
      <div className="card-title">
        <span className="icon">📁</span>
        요약 대상 문서 입력
      </div>

      {message && (
        <div className={`status-msg ${message.type}`}>
          {message.type === 'success' ? '✅' : message.type === 'error' ? '❌' : '⚠️'} {message.text}
        </div>
      )}

      <div className="tab-bar">
        {tabs.map((tab, i) => (
          <button
            key={i}
            className={`tab-btn ${activeTab === i ? 'active' : ''}`}
            onClick={() => setActiveTab(i)}
          >
            {tab}
          </button>
        ))}
      </div>

      {/* Tab 0: PDF Upload */}
      {activeTab === 0 && (
        <div>
          <div
            className={`drop-zone ${dragOver ? 'drag-over' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
          >
            <span className="icon">📄</span>
            <div className="title">{loading ? '업로드 중...' : 'PDF 파일을 드래그하거나 클릭하세요'}</div>
            <div className="subtitle">여러 파일을 한번에 업로드할 수 있습니다</div>
            {loading && <div className="spinner" style={{ marginTop: '0.75rem' }} />}
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf"
            multiple
            hidden
            onChange={(e) => e.target.files && handlePdfUpload(e.target.files)}
          />
        </div>
      )}

      {/* Tab 1: Folder Scan */}
      {activeTab === 1 && (
        <div>
          <div className="input-group">
            <label>📂 PDF 폴더 경로 (로컬)</label>
            <div style={{ display: 'flex', gap: '0.5rem' }}>
              <input
                className="input-field"
                value={folderPath}
                onChange={(e) => setFolderPath(e.target.value)}
                placeholder="/path/to/pdf/folder"
              />
              <button
                className="btn btn-primary"
                onClick={handleScanFolder}
                disabled={loading || !folderPath.trim()}
              >
                {loading ? <span className="spinner" /> : '스캔'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Tab 2: Patent Excel */}
      {activeTab === 2 && (
        <div>
          <div
            className="drop-zone"
            onClick={() => excelInputRef.current?.click()}
          >
            <span className="icon">📋</span>
            <div className="title">{loading ? '처리 중...' : '특허 URL 엑셀 파일을 선택하세요'}</div>
            <div className="subtitle">URL, 링크, WIPS 등의 컬럼이 포함된 .xlsx / .csv 파일</div>
            {loading && <div className="spinner" style={{ marginTop: '0.75rem' }} />}
          </div>
          <input
            ref={excelInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            hidden
            onChange={(e) => e.target.files?.[0] && handlePatentExcel(e.target.files[0])}
          />
        </div>
      )}

      {/* Document List */}
      {documents.length > 0 && (
        <>
          <div className="metrics-row" style={{ marginTop: '1.5rem' }}>
            <div className="metric-card">
              <div className="metric-value">{documents.length}</div>
              <div className="metric-label">총 문서</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">{documents.filter(d => d.doc_type === 'paper').length}</div>
              <div className="metric-label">논문</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">{documents.filter(d => d.doc_type === 'patent').length}</div>
              <div className="metric-label">특허</div>
            </div>
          </div>

          <ul className="doc-list">
            {documents.map((doc, i) => (
              <li key={i} className="doc-item">
                <span className="doc-icon">{doc.doc_type === 'patent' ? '📜' : '🔬'}</span>
                <div className="doc-info">
                  <div className="doc-name">{doc.filename}</div>
                  <div className="doc-meta">
                    {doc.title?.substring(0, 60)}{doc.title && doc.title.length > 60 ? '...' : ''}
                    {doc.year ? ` · ${doc.year}` : ''}
                  </div>
                </div>
                <button className="doc-remove" onClick={() => handleRemove(i)} title="제거">✕</button>
              </li>
            ))}
          </ul>
        </>
      )}
    </div>
  );
}
