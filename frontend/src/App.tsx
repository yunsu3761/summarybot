import { useState, useCallback, useEffect } from 'react';
import WorkTypeSelector from './components/WorkTypeSelector';
import DocumentInput from './components/DocumentInput';
import ExcelOptionSelector from './components/ExcelOptionSelector';
import ReorderUploadPanel from './components/ReorderUploadPanel';
import SummarizePanel from './components/SummarizePanel';
import {
  getSettings, updateSettings, clearDocuments, resetSession,
  DocumentInfo, MatchResult, OrderItem,
  startSummarize, getJobStatus, JobStatus,
  reorderExistingPpt, generateMetaExcel,
  getDownloadUrl,
} from './api';

type WorkType = 'summary' | 'reorder' | null;

export default function App() {
  // ── Global settings ──
  const [apiKeySet, setApiKeySet] = useState(false);
  const [geminiKeySet, setGeminiKeySet] = useState(false);
  const [model, setModel] = useState('gpt-5');
  const [templatePath, setTemplatePath] = useState('');
  const [outputDir] = useState('output');

  useEffect(() => {
    getSettings().then((s) => {
      setApiKeySet(s.api_key_set);
      setGeminiKeySet(s.gemini_key_set);
      if (s.model) setModel(s.model);
      if (s.template_path) setTemplatePath(s.template_path);
    }).catch(() => {});
  }, []);

  // ── Workflow state ──
  const [workType, setWorkType] = useState<WorkType>(null);

  // Case A: Summary
  const [documents, setDocuments] = useState<DocumentInfo[]>([]);
  const [isExcelProvided, setIsExcelProvided] = useState<boolean | null>(null);
  const [matchResults, setMatchResults] = useState<MatchResult[]>([]);
  const [orderData, setOrderData] = useState<OrderItem[]>([]);

  // Case A: Summarize execution
  const [jobId, setJobId] = useState<string | null>(null);
  const [jobStatus, setJobStatus] = useState<JobStatus | null>(null);
  const [summaryLoading, setSummaryLoading] = useState(false);
  const [summaryDone, setSummaryDone] = useState(false);
  const [summaryError, setSummaryError] = useState('');

  // Case B: Reorder
  const [reorderPptFile, setReorderPptFile] = useState<File | null>(null);
  const [reorderExcelFile, setReorderExcelFile] = useState<File | null>(null);
  const [reorderLoading, setReorderLoading] = useState(false);
  const [reorderResult, setReorderResult] = useState<{
    ok: boolean;
    output_file?: string;
  } | null>(null);
  const [reorderError, setReorderError] = useState('');
  const [reorderOldCol, setReorderOldCol] = useState('');
  const [reorderNewCol, setReorderNewCol] = useState('');

  // Meta excel generation
  const [metaExcelResult, setMetaExcelResult] = useState<{
    ok: boolean;
    output_file?: string;
  } | null>(null);

  // ── Handlers ──
  const handleModelChange = useCallback(async (m: string) => {
    setModel(m);
    await updateSettings({ model: m });
  }, []);

  const handleDocumentsChange = useCallback((docs: DocumentInfo[]) => {
    setDocuments(docs);
  }, []);

  const handleMatchResults = useCallback((results: MatchResult[], order: OrderItem[]) => {
    setMatchResults(results);
    setOrderData(order);
  }, []);

  // ── Poll job status ──
  useEffect(() => {
    if (!jobId || !summaryLoading) return;
    const interval = setInterval(async () => {
      try {
        const s = await getJobStatus(jobId);
        setJobStatus(s);
        if (s.status === 'done') {
          setSummaryLoading(false);
          setSummaryDone(true);
          // 요약 완료 시 항상 메타 엑셀 자동 생성
          try {
            const res = await generateMetaExcel();
            if (res.ok) {
              setMetaExcelResult({ ok: true, output_file: res.output_file });
            }
          } catch { /* ignore */ }
        } else if (s.status === 'error') {
          setSummaryLoading(false);
          setSummaryError('요약 중 오류 발생');
        }
      } catch {
        setSummaryLoading(false);
        setSummaryError('서버 연결 실패');
      }
    }, 2000);
    return () => clearInterval(interval);
  }, [jobId, summaryLoading, isExcelProvided]);

  // ── Execute ──
  const handleExecuteSummary = async () => {
    const keyOk = apiKeySet || (model.startsWith('gemini') && geminiKeySet);
    if (!keyOk) {
      setSummaryError('.env 파일에 API Key가 설정되지 않았습니다');
      return;
    }
    if (documents.length === 0) {
      setSummaryError('문서가 없습니다. PDF를 먼저 업로드하세요');
      return;
    }
    setSummaryError('');
    setSummaryLoading(true);
    setJobStatus(null);
    setSummaryDone(false);
    setMetaExcelResult(null);

    try {
      const res = await startSummarize({
        model,
        template_path: templatePath,
        output_dir: outputDir,
      });
      if (res.error) {
        setSummaryError(res.error);
        setSummaryLoading(false);
      } else {
        setJobId(res.job_id);
      }
    } catch {
      setSummaryError('요약 시작 실패');
      setSummaryLoading(false);
    }
  };

  const handleExecuteReorder = async () => {
    if (!reorderPptFile || !reorderExcelFile) return;
    setReorderLoading(true);
    setReorderError('');
    setReorderResult(null);
    try {
      const res = await reorderExistingPpt(reorderPptFile, reorderExcelFile, reorderOldCol, reorderNewCol);
      if (res.error) {
        setReorderError(res.error);
      } else {
        setReorderResult({ ok: true, output_file: res.output_file });
      }
    } catch {
      setReorderError('재정렬 실패');
    }
    setReorderLoading(false);
  };

  // ── Validation ──
  const isSummaryReady = (() => {
    if (documents.length === 0) return false;
    if (isExcelProvided === null) return false;
    if (isExcelProvided === true && orderData.length === 0) return false;
    const keyOk = apiKeySet || (model.startsWith('gemini') && geminiKeySet);
    if (!keyOk) return false;
    return true;
  })();

  const isReorderReady = reorderPptFile !== null && reorderExcelFile !== null;

  // ── Reset ──
  const handleReset = async () => {
    await resetSession();
    await clearDocuments();
    setWorkType(null);
    setDocuments([]);
    setIsExcelProvided(null);
    setMatchResults([]);
    setOrderData([]);
    setJobId(null);
    setJobStatus(null);
    setSummaryLoading(false);
    setSummaryDone(false);
    setSummaryError('');
    setReorderPptFile(null);
    setReorderExcelFile(null);
    setReorderLoading(false);
    setReorderResult(null);
    setReorderError('');
    setMetaExcelResult(null);
  };

  // ── Progress ──
  const progressPct = jobStatus
    ? Math.round((jobStatus.progress / Math.max(jobStatus.total, 1)) * 100)
    : 0;

  return (
    <div className="app-container">
      {/* Header */}
      <header className="app-header">
        <h1>📄 후보기술 개요서 자동 요약 봇</h1>
        <p>논문/특허 → LLM 요약 → PPT 개요서 자동 생성 · 후보기술 재정렬</p>
      </header>

      {/* Settings Bar */}
      <div className="settings-bar">
        <div className="settings-group">
          <label>🔑 API Key</label>
          <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
            <span style={{
              padding: '0.3rem 0.7rem',
              borderRadius: 'var(--radius-sm)',
              fontSize: '0.82rem',
              background: apiKeySet ? 'rgba(76, 175, 80, 0.15)' : 'rgba(244, 67, 54, 0.15)',
              color: apiKeySet ? 'var(--success)' : 'var(--error)',
              border: `1px solid ${apiKeySet ? 'var(--success)' : 'var(--error)'}`,
            }}>
              OpenAI: {apiKeySet ? '✅ 설정됨' : '❌ 미설정'}
            </span>
            <span style={{
              padding: '0.3rem 0.7rem',
              borderRadius: 'var(--radius-sm)',
              fontSize: '0.82rem',
              background: geminiKeySet ? 'rgba(76, 175, 80, 0.15)' : 'rgba(150, 150, 150, 0.1)',
              color: geminiKeySet ? 'var(--success)' : 'var(--text-muted)',
              border: `1px solid ${geminiKeySet ? 'var(--success)' : 'var(--border)'}`,
            }}>
              Gemini: {geminiKeySet ? '✅ 설정됨' : '➖ 미설정'}
            </span>
          </div>
        </div>
        <div className="settings-group">
          <label>🤖 모델</label>
          <select value={model} onChange={(e) => handleModelChange(e.target.value)}>
            <optgroup label="OpenAI">
              <option value="gpt-5">gpt-5</option>
              <option value="gpt-4o">gpt-4o</option>
              <option value="gpt-4o-mini">gpt-4o-mini</option>
            </optgroup>
            <optgroup label="Google Gemini">
              <option value="gemini-2.5-flash-preview-05-20">gemini-2.5-flash (최신)</option>
              <option value="gemini-2.5-flash-preview-04-17">gemini-2.5-flash (04-17)</option>
              <option value="gemini-2.5-pro-preview-05-06">gemini-2.5-pro</option>
              <option value="gemini-2.0-flash">gemini-2.0-flash</option>
            </optgroup>
          </select>
        </div>
        <div className="settings-group">
          <label>📄 템플릿</label>
          <input
            type="text"
            value={templatePath}
            onChange={(e) => setTemplatePath(e.target.value)}
            style={{ width: '300px', fontSize: '0.82rem' }}
          />
        </div>
        <div style={{ marginLeft: 'auto' }}>
          <button className="btn btn-secondary" onClick={handleReset}>
            🔄 초기화
          </button>
        </div>
      </div>

      {/* ─── Step 1: Work Type Selection ─── */}
      <WorkTypeSelector workType={workType} onSelect={setWorkType} />

      {/* ─── Case A: Summary Workflow ─── */}
      {workType === 'summary' && (
        <>
          {/* Step 2: PDF Upload */}
          <div className="section-block fade-in">
            <div className="section-label">
              <span className="section-number">2</span>
              요약 대상 문서 입력
            </div>
            <DocumentInput
              documents={documents}
              onDocumentsChange={handleDocumentsChange}
              summaryDone={summaryDone}
            />
          </div>

          {/* Step 3: Excel Option (only if documents exist) */}
          {documents.length > 0 && (
            <ExcelOptionSelector
              isExcelProvided={isExcelProvided}
              onOptionChange={setIsExcelProvided}
              onMatchResults={handleMatchResults}
              matchResults={matchResults}
              orderData={orderData}
            />
          )}

          {/* Step 4: Summary settings & progress */}
          {documents.length > 0 && isExcelProvided !== null && (
            <div className="section-block fade-in">
              <div className="section-label">
                <span className="section-number">4</span>
                요약 실행
              </div>

              <div className="metrics-row">
                <div className="metric-card">
                  <div className="metric-value">{documents.length}</div>
                  <div className="metric-label">요약 대상 문서</div>
                </div>
                <div className="metric-card">
                  <div className="metric-value" style={{ fontSize: '1.2rem' }}>{model}</div>
                  <div className="metric-label">LLM 모델</div>
                </div>
                {jobStatus && (
                  <>
                    <div className="metric-card">
                      <div className="metric-value">{jobStatus.completed_count}</div>
                      <div className="metric-label">완료</div>
                    </div>
                    <div className="metric-card">
                      <div className="metric-value">{jobStatus.error_count}</div>
                      <div className="metric-label">오류</div>
                    </div>
                  </>
                )}
              </div>

              {summaryError && (
                <div className="status-msg error">❌ {summaryError}</div>
              )}

              {/* Progress */}
              {jobStatus && (
                <div className="progress-container">
                  <div className="progress-bar-bg">
                    <div
                      className="progress-bar-fill"
                      style={{ width: `${progressPct}%` }}
                    />
                  </div>
                  <div className="progress-text">
                    <span>{progressPct}% ({jobStatus.progress}/{jobStatus.total})</span>
                    {jobStatus.current_file && jobStatus.status === 'running' && (
                      <span className="current-file">📄 {jobStatus.current_file}</span>
                    )}
                    {jobStatus.status === 'done' && <span style={{ color: 'var(--success)' }}>✅ 완료!</span>}
                  </div>
                </div>
              )}

              {/* Error details */}
              {jobStatus && jobStatus.errors.length > 0 && (
                <div style={{ marginTop: '1rem' }}>
                  <h4 style={{ fontSize: '0.85rem', color: 'var(--warning)', marginBottom: '0.5rem' }}>
                    ⚠️ 오류 목록 ({jobStatus.errors.length}건)
                  </h4>
                  <div style={{
                    maxHeight: '150px',
                    overflow: 'auto',
                    background: 'var(--bg-glass)',
                    borderRadius: 'var(--radius-sm)',
                    padding: '0.5rem 0.75rem',
                    fontSize: '0.82rem',
                  }}>
                    {jobStatus.errors.map((e, i) => (
                      <div key={i} style={{ marginBottom: '0.3rem', color: 'var(--text-secondary)' }}>
                        <strong>{e.filename}:</strong> {e.error}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Downloads (after completion) */}
              {summaryDone && (
                <div className="download-section fade-in">
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
                  {metaExcelResult?.ok && metaExcelResult.output_file && (
                    <div className="download-card">
                      <div className="file-icon">📋</div>
                      <div className="file-name">{metaExcelResult.output_file}</div>
                      <a
                        href={getDownloadUrl(metaExcelResult.output_file)}
                        className="btn btn-success"
                        download
                      >
                        📥 메타정보 엑셀 다운로드
                      </a>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {/* Execute Button */}
          {documents.length > 0 && isExcelProvided !== null && (
            <div className="execute-bar fade-in">
              <button
                className="btn btn-primary btn-lg btn-execute"
                onClick={handleExecuteSummary}
                disabled={!isSummaryReady || summaryLoading}
              >
                {summaryLoading ? (
                  <><span className="spinner" /> 요약 진행 중...</>
                ) : summaryDone ? (
                  '🔄 다시 요약하기'
                ) : (
                  '🚀 요약 및 PPT 생성 시작'
                )}
              </button>
              {!isSummaryReady && !summaryLoading && (
                <p className="execute-hint">
                  {!(apiKeySet || (model.startsWith('gemini') && geminiKeySet))
                    ? '⚠️ .env 파일에 API Key를 설정하세요'
                    : isExcelProvided === true && orderData.length === 0
                    ? '⚠️ 엑셀 파일을 업로드하세요'
                    : ''}
                </p>
              )}
            </div>
          )}
        </>
      )}

      {/* ─── Case B: Reorder Workflow ─── */}
      {workType === 'reorder' && (
        <>
          <ReorderUploadPanel
            pptFile={reorderPptFile}
            excelFile={reorderExcelFile}
            onPptChange={setReorderPptFile}
            onExcelChange={setReorderExcelFile}
            oldOrderCol={reorderOldCol}
            newOrderCol={reorderNewCol}
            setOldOrderCol={setReorderOldCol}
            setNewOrderCol={setReorderNewCol}
          />

          {reorderError && (
            <div className="status-msg error" style={{ marginTop: '1rem' }}>
              ❌ {reorderError}
            </div>
          )}

          {reorderResult?.ok && (
            <div className="section-block fade-in">
              <div className="status-msg success">
                ✅ PPT 재정렬 완료!
              </div>
              <div className="download-section">
                <div className="download-card">
                  <div className="file-icon">🔄</div>
                  <div className="file-name">{reorderResult.output_file}</div>
                  <a
                    href={getDownloadUrl(reorderResult.output_file || '')}
                    className="btn btn-success"
                    download
                  >
                    📥 재정렬된 PPT 다운로드
                  </a>
                </div>
              </div>
            </div>
          )}

          {/* Execute Button */}
          <div className="execute-bar fade-in">
            <button
              className="btn btn-success btn-lg btn-execute"
              onClick={handleExecuteReorder}
              disabled={!isReorderReady || reorderLoading}
            >
              {reorderLoading ? (
                <><span className="spinner" /> 재정렬 중...</>
              ) : reorderResult?.ok ? (
                '🔄 다시 재정렬하기'
              ) : (
                '🔄 PPT 재정렬 실행'
              )}
            </button>
            {!isReorderReady && (
              <p className="execute-hint">
                {!reorderPptFile ? '⚠️ PPT 파일을 업로드하세요' : '⚠️ 엑셀 파일을 업로드하세요'}
              </p>
            )}
          </div>
        </>
      )}

      {/* Footer */}
      <div style={{
        textAlign: 'center',
        color: 'var(--text-muted)',
        fontSize: '0.8rem',
        marginTop: '3rem',
        padding: '1rem 0',
        borderTop: '1px solid var(--border)',
      }}>
        후보기술 개요서 자동 요약 봇 v2.0 — React + Flask
      </div>
    </div>
  );
}
