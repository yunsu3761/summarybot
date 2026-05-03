import { useState, useRef, useEffect } from 'react';
import { uploadOrderExcel, analyzeExcelHeaders, MatchResult, OrderItem } from '../api';

interface ExcelOptionSelectorProps {
  isExcelProvided: boolean | null;
  onOptionChange: (v: boolean) => void;
  onMatchResults: (results: MatchResult[], orderData: OrderItem[]) => void;
  matchResults: MatchResult[];
  orderData: OrderItem[];
}

export default function ExcelOptionSelector({
  isExcelProvided,
  onOptionChange,
  onMatchResults,
  matchResults,
  orderData,
}: ExcelOptionSelectorProps) {
  const [loading, setLoading] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [message, setMessage] = useState<{ type: string; text: string } | null>(null);
  
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [indexCol, setIndexCol] = useState('');
  
  const fileRef = useRef<HTMLInputElement>(null);

  const showMsg = (type: string, text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage(null), 5000);
  };

  useEffect(() => {
    if (selectedFile) {
      setAnalyzing(true);
      analyzeExcelHeaders(selectedFile)
        .then((res) => {
          if (res.headers) {
            setHeaders(res.headers);
            // Use recommended_old_col (which matches 번호, No.) or recommended_new_col
            if (res.recommended_old_col) setIndexCol(res.recommended_old_col);
            else if (res.recommended_new_col) setIndexCol(res.recommended_new_col);
          }
          setAnalyzing(false);
        })
        .catch(() => setAnalyzing(false));
    } else {
      setHeaders([]);
      setIndexCol('');
    }
  }, [selectedFile]);

  const handleUpload = async () => {
    if (!selectedFile) return;
    setLoading(true);
    try {
      const res = await uploadOrderExcel(selectedFile, indexCol);
      if (res.error) {
        showMsg('error', res.error);
      } else {
        onMatchResults(res.match_results || [], res.order_data || []);
        showMsg('success', `기준 리스트 ${res.ref_count}개 항목 로드, 매칭 완료`);
        setSelectedFile(null); // Clear selected file after successful upload to show only results
      }
    } catch {
      showMsg('error', '엑셀 업로드 실패');
    }
    setLoading(false);
  };

  const handleFileChange = (file: File) => {
    setSelectedFile(file);
    onMatchResults([], []); // Reset previous matching
  };

  return (
    <div className="section-block fade-in">
      <div className="section-label">
        <span className="section-number">3</span>
        순번을 참고할 후보기술 리스트(Excel)가 있습니까?
      </div>

      <div className="option-toggle-group">
        <button
          className={`option-toggle-btn ${isExcelProvided === true ? 'active yes' : ''}`}
          onClick={() => onOptionChange(true)}
        >
          ✅ 예, 엑셀이 있습니다
        </button>
        <button
          className={`option-toggle-btn ${isExcelProvided === false ? 'active no' : ''}`}
          onClick={() => onOptionChange(false)}
        >
          ❌ 아니오, 자동 생성
        </button>
      </div>

      {message && (
        <div className={`status-msg ${message.type}`} style={{ marginTop: '1rem' }}>
          {message.type === 'success' ? '✅' : '❌'} {message.text}
        </div>
      )}

      {/* 예: 엑셀 업로드 영역 (매칭 전) */}
      {isExcelProvided === true && orderData.length === 0 && (
        <div className="fade-in" style={{ marginTop: '1.25rem' }}>
          {!selectedFile ? (
            <div
              className="drop-zone"
              onClick={() => fileRef.current?.click()}
              style={{ padding: '2rem 1.5rem' }}
            >
              <span className="icon">📊</span>
              <div className="title">후보기술 순서 엑셀 파일을 선택하세요</div>
              <div className="subtitle">
                후보기술 번호, 기술명, DOI, 특허번호 등이 포함된 .xlsx / .csv 파일
              </div>
            </div>
          ) : (
            <div className="file-selected-card" style={{ display: 'flex', flexDirection: 'column', gap: '1rem', padding: '1.5rem' }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                  <span style={{ fontSize: '1.5rem' }}>📋</span>
                  <div>
                    <div style={{ fontWeight: 600 }}>{selectedFile.name}</div>
                    <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                      {(selectedFile.size / 1024).toFixed(0)} KB
                    </div>
                  </div>
                </div>
                <button
                  className="btn btn-secondary btn-sm"
                  onClick={() => setSelectedFile(null)}
                  disabled={loading}
                >
                  취소
                </button>
              </div>

              {analyzing ? (
                <div style={{ textAlign: 'center', padding: '1rem', color: 'var(--text-muted)', fontSize: '0.9rem' }}>
                  <span className="spinner" /> 엑셀 구조 분석 중...
                </div>
              ) : (
                <div style={{ background: 'var(--bg-glass)', border: '1px solid var(--border)', borderRadius: 'var(--radius-sm)', padding: '1rem' }}>
                  <label style={{ fontSize: '0.9rem', color: 'var(--text-secondary)', display: 'block', marginBottom: '0.5rem' }}>
                    어떤 순서로 개요서 번호를 설정할 것인지 컬럼을 선택하세요:
                  </label>
                  <select
                    value={indexCol}
                    onChange={(e) => setIndexCol(e.target.value)}
                    style={{ width: '100%', padding: '0.6rem', borderRadius: 'var(--radius-sm)', border: '1px solid var(--border)', marginBottom: '1rem' }}
                  >
                    <option value="">(자동 찾기)</option>
                    {headers.map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                  
                  <button
                    className="btn btn-primary btn-block"
                    onClick={handleUpload}
                    disabled={loading}
                  >
                    {loading ? <><span className="spinner" /> 처리 중...</> : '✅ 분석 및 매칭 시작'}
                  </button>
                </div>
              )}
            </div>
          )}
          
          <input
            ref={fileRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            hidden
            onChange={(e) => e.target.files?.[0] && handleFileChange(e.target.files[0])}
          />
        </div>
      )}

      {/* 예: 엑셀 업로드 완료 (매칭 결과 표시) */}
      {isExcelProvided === true && orderData.length > 0 && (
        <div className="fade-in" style={{ marginTop: '1.25rem' }}>
           <div className="status-msg success">
            ✅ {orderData.length}개 항목 로드 완료 |
            매칭 {matchResults.filter(r => r.match_type !== 'unmatched').length}건 |
            미매칭 {matchResults.filter(r => r.match_type === 'unmatched').length}건
          </div>
          <button 
            className="btn btn-secondary btn-sm" 
            style={{ marginTop: '0.75rem' }}
            onClick={() => onMatchResults([], [])}
          >
            🔄 엑셀 파일 다시 선택하기
          </button>
        </div>
      )}

      {/* 아니오: 안내 메시지 */}
      {isExcelProvided === false && (
        <div className="fade-in" style={{ marginTop: '1.25rem' }}>
          <div className="status-msg info">
            ℹ️ 요약 완료 후, 모든 기술 개요의 메타정보를 포함한 엑셀 파일(briefs_summary.xlsx)이 자동 생성됩니다
          </div>
        </div>
      )}
    </div>
  );
}
