// API utility functions for communicating with Flask backend

const API_BASE = '/api';

export interface DocumentInfo {
  filename: string;
  doc_type: string;
  title: string;
  doi: string;
  patent_no: string;
  institution: string;
  year: string;
  source_type: string;
  source_path: string;
  has_text: boolean;
  error?: string;
}

export interface MatchResult {
  filename: string;
  doc_type: string;
  title: string;
  ref_index: number | null;
  ref_title: string;
  score: number;
  match_type: string;
  needs_review: boolean;
  status: string;
}

export interface OrderItem {
  index: number;
  title: string;
  category: string;
}

export interface JobStatus {
  id: string;
  status: 'running' | 'done' | 'error';
  progress: number;
  total: number;
  current_file: string;
  completed_count: number;
  error_count: number;
  errors: { filename: string; error: string }[];
  output_file: string | null;
}

// ---------- Settings ----------
export async function getSettings() {
  const res = await fetch(`${API_BASE}/settings`);
  return res.json();
}

export async function updateSettings(data: {
  api_key?: string;
  model?: string;
  template_path?: string;
  output_dir?: string;
}) {
  const res = await fetch(`${API_BASE}/settings`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return res.json();
}

// ---------- Documents ----------
export async function uploadPdfs(files: FileList, reset = false): Promise<{
  uploaded: number;
  documents: DocumentInfo[];
  total: number;
}> {
  const formData = new FormData();
  for (let i = 0; i < files.length; i++) {
    formData.append('files', files[i]);
  }
  if (reset) formData.append('reset', 'true');
  const res = await fetch(`${API_BASE}/upload-pdfs`, {
    method: 'POST',
    body: formData,
  });
  return res.json();
}

export async function scanFolder(folderPath: string, reset = false) {
  const res = await fetch(`${API_BASE}/scan-folder`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ folder_path: folderPath, reset }),
  });
  return res.json();
}

export async function uploadPatentExcel(file: File, reset = false) {
  const formData = new FormData();
  formData.append('file', file);
  if (reset) formData.append('reset', 'true');
  const res = await fetch(`${API_BASE}/upload-patent-excel`, {
    method: 'POST',
    body: formData,
  });
  return res.json();
}

export async function getDocuments(): Promise<{
  documents: DocumentInfo[];
  total: number;
}> {
  const res = await fetch(`${API_BASE}/documents`);
  return res.json();
}

export async function clearDocuments() {
  const res = await fetch(`${API_BASE}/documents`, { method: 'DELETE' });
  return res.json();
}

export async function deleteDocument(index: number) {
  const res = await fetch(`${API_BASE}/documents/${index}`, {
    method: 'DELETE',
  });
  return res.json();
}

// ---------- Order Excel & Matching ----------
export async function uploadOrderExcel(file: File, indexCol?: string): Promise<{
  ref_count: number;
  match_results: MatchResult[];
  order_data: OrderItem[];
  error?: string;
}> {
  const formData = new FormData();
  formData.append('file', file);
  if (indexCol) formData.append('index_col', indexCol);
  
  const res = await fetch(`${API_BASE}/upload-order-excel`, {
    method: 'POST',
    body: formData,
  });
  return res.json();
}

// ---------- Summarize ----------
export async function startSummarize(params: {
  api_key?: string;
  model?: string;
  template_path?: string;
  output_dir?: string;
}): Promise<{ job_id: string; status: string; error?: string }> {
  const res = await fetch(`${API_BASE}/summarize`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(params),
  });
  return res.json();
}

export async function getJobStatus(jobId: string): Promise<JobStatus> {
  const res = await fetch(`${API_BASE}/status/${jobId}`);
  return res.json();
}

// ---------- Reorder (기존 세션 기반) ----------
export async function reorderPpt(): Promise<{
  ok?: boolean;
  output_file?: string;
  error?: string;
}> {
  const res = await fetch(`${API_BASE}/reorder`, { method: 'POST' });
  return res.json();
}

// ---------- Reorder Existing (Case B: PPT + Excel 직접 업로드) ----------
export async function reorderExistingPpt(
  pptFile: File,
  excelFile: File,
  oldOrderCol?: string,
  newOrderCol?: string
): Promise<{
  ok?: boolean;
  output_file?: string;
  slide_count?: number;
  excel_count?: number;
  error?: string;
}> {
  const formData = new FormData();
  formData.append('ppt_file', pptFile);
  formData.append('excel_file', excelFile);
  if (oldOrderCol) formData.append('old_order_col', oldOrderCol);
  if (newOrderCol) formData.append('new_order_col', newOrderCol);
  
  const res = await fetch(`${API_BASE}/reorder-existing`, {
    method: 'POST',
    body: formData,
  });
  return res.json();
}

export async function analyzeExcelHeaders(file: File): Promise<{
  headers?: string[];
  recommended_old_col?: string;
  recommended_new_col?: string;
  error?: string;
}> {
  const formData = new FormData();
  formData.append('file', file);
  const res = await fetch(`${API_BASE}/analyze-excel-headers`, {
    method: 'POST',
    body: formData,
  });
  return res.json();
}

// ---------- Generate Meta Excel (Case A: 엑셀 없음) ----------
export async function generateMetaExcel(): Promise<{
  ok?: boolean;
  output_file?: string;
  record_count?: number;
  error?: string;
}> {
  const res = await fetch(`${API_BASE}/generate-meta-excel`, { method: 'POST' });
  return res.json();
}

// ---------- Download ----------
export function getDownloadUrl(filename: string): string {
  return `${API_BASE}/download/${encodeURIComponent(filename)}`;
}

// ---------- Reset ----------
export async function resetSession() {
  const res = await fetch(`${API_BASE}/reset`, { method: 'POST' });
  return res.json();
}
