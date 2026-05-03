import { useState, useEffect, useCallback } from 'react';
import { startSummarize, getJobStatus, JobStatus } from '../api';

interface SummarizePanelProps {
  documentCount: number;
  apiKeySet: boolean;
  model: string;
  templatePath: string;
  outputDir: string;
  onComplete: () => void;
}

export default function SummarizePanel({
  documentCount,
  apiKeySet,
  model,
  templatePath,
  outputDir,
  onComplete,
}: SummarizePanelProps) {
  const [jobId, setJobId] = useState<string | null>(null);
  const [status, setStatus] = useState<JobStatus | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  // Poll for job status
  const pollStatus = useCallback(async (id: string) => {
    try {
      const s = await getJobStatus(id);
      setStatus(s);
      if (s.status === 'done') {
        setLoading(false);
        onComplete();
      } else if (s.status === 'error') {
        setLoading(false);
        setError('요약 중 오류 발생');
      }
    } catch {
      setLoading(false);
      setError('서버 연결 실패');
    }
  }, [onComplete]);

  useEffect(() => {
    if (!jobId || !loading) return;
    const interval = setInterval(() => pollStatus(jobId), 2000);
    return () => clearInterval(interval);
  }, [jobId, loading, pollStatus]);

  const handleStart = async () => {
    if (!apiKeySet) {
      setError('.env 파일에 API Key가 설정되지 않았습니다');
      return;
    }
    if (documentCount === 0) {
      setError('문서가 없습니다. Step 1에서 문서를 입력하세요');
      return;
    }

    setError('');
    setLoading(true);
    setStatus(null);

    try {
      const res = await startSummarize({
        model,
        template_path: templatePath,
        output_dir: outputDir,
      });
      if (res.error) {
        setError(res.error);
        setLoading(false);
      } else {
        setJobId(res.job_id);
      }
    } catch {
      setError('요약 시작 실패');
      setLoading(false);
    }
  };

  const progressPct = status ? Math.round((status.progress / Math.max(status.total, 1)) * 100) : 0;

  return (
    <div className="card fade-in">
      <div className="card-title">
        <span className="icon">📝</span>
        문서 요약 · PPT 생성
      </div>

      {documentCount === 0 && (
        <div className="status-msg info">
          ℹ️ Step 1에서 먼저 문서를 입력해주세요
        </div>
      )}

      {!apiKeySet && (
        <div className="status-msg warning">
          ⚠️ .env 파일에 API Key를 설정해주세요 (OPENAI_API_KEY 또는 GEMINI_API_KEY)
        </div>
      )}

      {error && (
        <div className="status-msg error">❌ {error}</div>
      )}

      {/* Summary info */}
      <div className="metrics-row">
        <div className="metric-card">
          <div className="metric-value">{documentCount}</div>
          <div className="metric-label">요약 대상 문서</div>
        </div>
        <div className="metric-card">
          <div className="metric-value" style={{ fontSize: '1.2rem' }}>{model}</div>
          <div className="metric-label">LLM 모델</div>
        </div>
        <div className="metric-card">
          <div className="metric-value">{status?.completed_count ?? '-'}</div>
          <div className="metric-label">완료</div>
        </div>
        <div className="metric-card">
          <div className="metric-value">{status?.error_count ?? '-'}</div>
          <div className="metric-label">오류</div>
        </div>
      </div>

      {/* Progress */}
      {status && (
        <div className="progress-container">
          <div className="progress-bar-bg">
            <div
              className="progress-bar-fill"
              style={{ width: `${progressPct}%` }}
            />
          </div>
          <div className="progress-text">
            <span>{progressPct}% ({status.progress}/{status.total})</span>
            {status.current_file && status.status === 'running' && (
              <span className="current-file">📄 {status.current_file}</span>
            )}
            {status.status === 'done' && <span style={{ color: 'var(--success)' }}>✅ 완료!</span>}
          </div>
        </div>
      )}

      {/* Error details */}
      {status && status.errors.length > 0 && (
        <div style={{ marginTop: '1rem' }}>
          <h4 style={{ fontSize: '0.85rem', color: 'var(--warning)', marginBottom: '0.5rem' }}>
            ⚠️ 오류 목록 ({status.errors.length}건)
          </h4>
          <div style={{
            maxHeight: '150px',
            overflow: 'auto',
            background: 'var(--bg-glass)',
            borderRadius: 'var(--radius-sm)',
            padding: '0.5rem 0.75rem',
            fontSize: '0.82rem',
          }}>
            {status.errors.map((e, i) => (
              <div key={i} style={{ marginBottom: '0.3rem', color: 'var(--text-secondary)' }}>
                <strong>{e.filename}:</strong> {e.error}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Start button */}
      <div style={{ marginTop: '1.5rem' }}>
        <button
          className="btn btn-primary btn-lg btn-block"
          onClick={handleStart}
          disabled={loading || documentCount === 0 || !apiKeySet}
        >
          {loading ? (
            <>
              <span className="spinner" /> 요약 진행 중...
            </>
          ) : status?.status === 'done' ? (
            '🔄 다시 요약하기'
          ) : (
            '🚀 요약 및 PPT 생성 시작'
          )}
        </button>
      </div>
    </div>
  );
}
