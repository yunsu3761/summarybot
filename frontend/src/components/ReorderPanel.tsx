import { useState } from 'react';
import { reorderPpt, getDownloadUrl } from '../api';

interface ReorderPanelProps {
  hasOutput: boolean;
  hasOrderData: boolean;
}

export default function ReorderPanel({ hasOutput, hasOrderData }: ReorderPanelProps) {
  const [loading, setLoading] = useState(false);
  const [reorderResult, setReorderResult] = useState<{
    ok: boolean;
    output_file?: string;
  } | null>(null);
  const [error, setError] = useState('');

  const handleReorder = async () => {
    setLoading(true);
    setError('');
    try {
      const res = await reorderPpt();
      if (res.error) {
        setError(res.error);
      } else {
        setReorderResult({ ok: true, output_file: res.output_file });
      }
    } catch {
      setError('재정렬 실패');
    }
    setLoading(false);
  };

  return (
    <div className="card fade-in">
      <div className="card-title">
        <span className="icon">🔄</span>
        후보기술 재정렬 · 다운로드
      </div>

      {!hasOutput && (
        <div className="status-msg info">
          ℹ️ Step 3에서 먼저 요약을 실행해주세요
        </div>
      )}

      {error && (
        <div className="status-msg error">❌ {error}</div>
      )}

      {reorderResult?.ok && (
        <div className="status-msg success">
          ✅ PPT 재정렬 완료!
        </div>
      )}

      {/* Downloads */}
      {hasOutput && (
        <div className="download-section">
          <div className="download-card">
            <div className="file-icon">📊</div>
            <div className="file-name">output_briefs.pptx</div>
            <a
              href={getDownloadUrl('output_briefs.pptx')}
              className="btn btn-primary"
              download
            >
              📥 요약 PPT 다운로드
            </a>
          </div>

          {reorderResult?.output_file && (
            <div className="download-card">
              <div className="file-icon">🔄</div>
              <div className="file-name">{reorderResult.output_file}</div>
              <a
                href={getDownloadUrl(reorderResult.output_file)}
                className="btn btn-success"
                download
              >
                📥 재정렬 PPT 다운로드
              </a>
            </div>
          )}
        </div>
      )}

      {/* Reorder button */}
      <div style={{ marginTop: '1.5rem' }}>
        <button
          className="btn btn-success btn-lg btn-block"
          onClick={handleReorder}
          disabled={loading || !hasOutput || !hasOrderData}
          title={!hasOrderData ? 'Step 2에서 순서 엑셀을 먼저 업로드하세요' : ''}
        >
          {loading ? (
            <>
              <span className="spinner" /> 재정렬 중...
            </>
          ) : (
            '🔄 후보기술 재정렬'
          )}
        </button>
        {!hasOrderData && hasOutput && (
          <p style={{ fontSize: '0.82rem', color: 'var(--text-muted)', marginTop: '0.5rem', textAlign: 'center' }}>
            Step 2에서 순서 엑셀을 업로드하면 재정렬이 가능합니다
          </p>
        )}
      </div>
    </div>
  );
}
