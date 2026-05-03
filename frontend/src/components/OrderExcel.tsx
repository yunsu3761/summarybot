import { useState, useRef } from 'react';
import { uploadOrderExcel, MatchResult, OrderItem } from '../api';

interface OrderExcelProps {
  documentCount: number;
  matchResults: MatchResult[];
  orderData: OrderItem[];
  onMatchResults: (results: MatchResult[], orderData: OrderItem[]) => void;
}

export default function OrderExcel({
  documentCount,
  matchResults,
  orderData,
  onMatchResults,
}: OrderExcelProps) {
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ type: string; text: string } | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  const showMsg = (type: string, text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage(null), 5000);
  };

  const handleUpload = async (file: File) => {
    setLoading(true);
    try {
      const res = await uploadOrderExcel(file);
      if (res.error) {
        showMsg('error', res.error);
      } else {
        onMatchResults(res.match_results || [], res.order_data || []);
        showMsg('success', `기준 리스트 ${res.ref_count}개 항목 로드, 매칭 완료`);
      }
    } catch {
      showMsg('error', '엑셀 업로드 실패');
    }
    setLoading(false);
  };

  const getBadgeClass = (matchType: string, score: number) => {
    if (matchType.startsWith('exact')) return 'badge-exact';
    if (matchType === 'fuzzy' && score >= 80) return 'badge-fuzzy';
    if (matchType === 'fuzzy') return 'badge-review';
    return 'badge-unmatched';
  };

  const getBadgeLabel = (matchType: string, score: number) => {
    if (matchType.startsWith('exact')) return '정확';
    if (matchType === 'fuzzy' && score >= 80) return '자동';
    if (matchType === 'fuzzy') return '검토';
    if (matchType === 'manual') return '수동';
    return '미매칭';
  };

  return (
    <div className="card fade-in">
      <div className="card-title">
        <span className="icon">🔗</span>
        문서 순서 엑셀 업로드 · 후보기술 매칭
      </div>

      {documentCount === 0 && (
        <div className="status-msg info">
          ℹ️ Step 1에서 먼저 문서를 입력해주세요
        </div>
      )}

      {message && (
        <div className={`status-msg ${message.type}`}>
          {message.type === 'success' ? '✅' : '❌'} {message.text}
        </div>
      )}

      <div
        className="drop-zone"
        onClick={() => fileRef.current?.click()}
        style={{ marginBottom: '1.5rem' }}
      >
        <span className="icon">📊</span>
        <div className="title">{loading ? '처리 중...' : '문서 순서 엑셀 파일을 선택하세요'}</div>
        <div className="subtitle">
          후보기술 번호, 기술명, DOI, 특허번호 등이 포함된 .xlsx / .csv 파일
        </div>
        {loading && <div className="spinner" style={{ marginTop: '0.75rem' }} />}
      </div>
      <input
        ref={fileRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        hidden
        onChange={(e) => e.target.files?.[0] && handleUpload(e.target.files[0])}
      />

      {/* Order Data Preview */}
      {orderData.length > 0 && (
        <div style={{ marginBottom: '1.5rem' }}>
          <h4 style={{ fontSize: '0.95rem', fontWeight: 600, marginBottom: '0.5rem', color: 'var(--text-secondary)' }}>
            📋 기준 후보기술 리스트 ({orderData.length}개)
          </h4>
          <div style={{ maxHeight: '200px', overflow: 'auto', borderRadius: 'var(--radius-sm)', border: '1px solid var(--border)' }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th>번호</th>
                  <th>기술명</th>
                  <th>분류</th>
                </tr>
              </thead>
              <tbody>
                {orderData.map((item, i) => (
                  <tr key={i}>
                    <td style={{ fontWeight: 600 }}>{item.index}</td>
                    <td>{item.title}</td>
                    <td>{item.category || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Match Results */}
      {matchResults.length > 0 && (
        <>
          <h4 style={{ fontSize: '0.95rem', fontWeight: 600, marginBottom: '0.5rem', color: 'var(--text-secondary)' }}>
            🔗 매칭 결과
          </h4>

          <div className="metrics-row">
            <div className="metric-card">
              <div className="metric-value">
                {matchResults.filter(r => r.match_type.startsWith('exact')).length}
              </div>
              <div className="metric-label">정확 매칭</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">
                {matchResults.filter(r => r.match_type === 'fuzzy' && r.score >= 80).length}
              </div>
              <div className="metric-label">자동 매칭</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">
                {matchResults.filter(r => r.needs_review).length}
              </div>
              <div className="metric-label">검토 필요</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">
                {matchResults.filter(r => r.match_type === 'unmatched').length}
              </div>
              <div className="metric-label">미매칭</div>
            </div>
          </div>

          <div style={{ overflowX: 'auto', borderRadius: 'var(--radius-sm)', border: '1px solid var(--border)' }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th>파일명</th>
                  <th>제목</th>
                  <th>매칭 번호</th>
                  <th>매칭 기술명</th>
                  <th>일치도</th>
                  <th>상태</th>
                </tr>
              </thead>
              <tbody>
                {matchResults.map((mr, i) => (
                  <tr key={i}>
                    <td style={{ fontWeight: 500, maxWidth: '180px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {mr.filename}
                    </td>
                    <td style={{ maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {mr.title}
                    </td>
                    <td style={{ fontWeight: 600 }}>{mr.ref_index ?? '-'}</td>
                    <td>{mr.ref_title || '-'}</td>
                    <td>{mr.score}%</td>
                    <td>
                      <span className={`badge ${getBadgeClass(mr.match_type, mr.score)}`}>
                        {getBadgeLabel(mr.match_type, mr.score)}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}
