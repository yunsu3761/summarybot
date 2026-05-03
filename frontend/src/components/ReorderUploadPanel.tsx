import { useRef, useState, useEffect } from 'react';
import { analyzeExcelHeaders } from '../api';

interface ReorderUploadPanelProps {
  pptFile: File | null;
  excelFile: File | null;
  onPptChange: (file: File | null) => void;
  onExcelChange: (file: File | null) => void;
  oldOrderCol: string;
  newOrderCol: string;
  setOldOrderCol: (v: string) => void;
  setNewOrderCol: (v: string) => void;
}

export default function ReorderUploadPanel({
  pptFile,
  excelFile,
  onPptChange,
  onExcelChange,
  oldOrderCol,
  newOrderCol,
  setOldOrderCol,
  setNewOrderCol,
}: ReorderUploadPanelProps) {
  const pptRef = useRef<HTMLInputElement>(null);
  const excelRef = useRef<HTMLInputElement>(null);
  const [pptDragOver, setPptDragOver] = useState(false);
  const [excelDragOver, setExcelDragOver] = useState(false);

  const [headers, setHeaders] = useState<string[]>([]);
  const [analyzing, setAnalyzing] = useState(false);

  useEffect(() => {
    if (excelFile) {
      setAnalyzing(true);
      analyzeExcelHeaders(excelFile).then((res) => {
        if (res.headers) {
          setHeaders(res.headers);
          if (res.recommended_old_col) setOldOrderCol(res.recommended_old_col);
          if (res.recommended_new_col) setNewOrderCol(res.recommended_new_col);
        }
        setAnalyzing(false);
      }).catch(() => setAnalyzing(false));
    } else {
      setHeaders([]);
      setOldOrderCol('');
      setNewOrderCol('');
    }
  }, [excelFile, setOldOrderCol, setNewOrderCol]);

  return (
    <div className="section-block fade-in">
      <div className="section-label">
        <span className="section-number">2</span>
        파일 업로드
      </div>

      <div className="reorder-upload-grid">
        {/* PPT 업로드 */}
        <div className="reorder-upload-item">
          <div className="reorder-upload-label">📊 후보기술 요약본 (PPT)</div>
          {pptFile ? (
            <div className="file-selected-card">
              <div className="file-selected-icon">📊</div>
              <div className="file-selected-info">
                <div className="file-selected-name">{pptFile.name}</div>
                <div className="file-selected-size">
                  {(pptFile.size / 1024 / 1024).toFixed(1)} MB
                </div>
              </div>
              <button
                className="file-selected-remove"
                onClick={() => onPptChange(null)}
                title="제거"
              >
                ✕
              </button>
            </div>
          ) : (
            <div
              className={`drop-zone drop-zone-sm ${pptDragOver ? 'drag-over' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setPptDragOver(true); }}
              onDragLeave={() => setPptDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setPptDragOver(false);
                if (e.dataTransfer.files[0]) onPptChange(e.dataTransfer.files[0]);
              }}
              onClick={() => pptRef.current?.click()}
            >
              <span className="icon" style={{ fontSize: '2rem' }}>📄</span>
              <div className="title">PPT 파일 선택</div>
              <div className="subtitle">.pptx 파일</div>
            </div>
          )}
          <input
            ref={pptRef}
            type="file"
            accept=".pptx"
            hidden
            onChange={(e) => e.target.files?.[0] && onPptChange(e.target.files[0])}
          />
        </div>

        {/* 엑셀 업로드 */}
        <div className="reorder-upload-item">
          <div className="reorder-upload-label">📋 순번 엑셀</div>
          {excelFile ? (
            <div className="file-selected-card">
              <div className="file-selected-icon">📋</div>
              <div className="file-selected-info">
                <div className="file-selected-name">{excelFile.name}</div>
                <div className="file-selected-size">
                  {(excelFile.size / 1024).toFixed(0)} KB
                </div>
              </div>
              <button
                className="file-selected-remove"
                onClick={() => onExcelChange(null)}
                title="제거"
              >
                ✕
              </button>
            </div>
          ) : (
            <div
              className={`drop-zone drop-zone-sm ${excelDragOver ? 'drag-over' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setExcelDragOver(true); }}
              onDragLeave={() => setExcelDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setExcelDragOver(false);
                if (e.dataTransfer.files[0]) onExcelChange(e.dataTransfer.files[0]);
              }}
              onClick={() => excelRef.current?.click()}
            >
              <span className="icon" style={{ fontSize: '2rem' }}>📋</span>
              <div className="title">엑셀 파일 선택</div>
              <div className="subtitle">.xlsx / .csv 파일</div>
            </div>
          )}
          <input
            ref={excelRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            hidden
            onChange={(e) => e.target.files?.[0] && onExcelChange(e.target.files[0])}
          />
        </div>

        {/* 컬럼 선택기 */}
        {excelFile && (
          <div className="reorder-upload-item" style={{ gridColumn: '1 / -1', marginTop: '1rem' }}>
            <div className="reorder-upload-label">엑셀 컬럼 매핑 설정</div>
            <div className="file-selected-card" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', background: 'var(--bg-glass)', border: '1px solid var(--border)' }}>
              {analyzing ? (
                <div style={{ padding: '1rem', textAlign: 'center', color: 'var(--text-muted)' }}>
                  <span className="spinner" /> 엑셀 파일 분석 중...
                </div>
              ) : (
                <div style={{ display: 'flex', gap: '1rem', width: '100%', flexWrap: 'wrap' }}>
                  <div style={{ flex: 1, minWidth: '200px' }}>
                    <label style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', display: 'block', marginBottom: '0.3rem' }}>
                      변경 전 (기존) <b>PPT 번호</b> 역할을 할 컬럼
                    </label>
                    <select 
                      value={oldOrderCol} 
                      onChange={(e) => setOldOrderCol(e.target.value)}
                      style={{ width: '100%', padding: '0.5rem', borderRadius: 'var(--radius-sm)', border: '1px solid var(--border)' }}
                    >
                      <option value="">(자동 찾기)</option>
                      {headers.map(h => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: 1, minWidth: '200px' }}>
                    <label style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', display: 'block', marginBottom: '0.3rem' }}>
                      변경 할 (새로운) <b>개요서 번호</b> 역할을 할 컬럼
                    </label>
                    <select 
                      value={newOrderCol} 
                      onChange={(e) => setNewOrderCol(e.target.value)}
                      style={{ width: '100%', padding: '0.5rem', borderRadius: 'var(--radius-sm)', border: '1px solid var(--border)' }}
                    >
                      <option value="">(자동 찾기)</option>
                      {headers.map(h => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
