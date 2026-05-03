import React from 'react';

interface WorkTypeSelectorProps {
  workType: 'summary' | 'reorder' | null;
  onSelect: (type: 'summary' | 'reorder') => void;
}

const WORK_TYPES = [
  {
    key: 'summary' as const,
    icon: '📝',
    title: '대상 PDF 요약',
    desc: '논문/특허 PDF를 업로드하면 LLM이 핵심 내용을 요약하고 PPT 개요서를 자동 생성합니다',
    tags: ['PDF 입력', 'LLM 요약', 'PPT 생성'],
  },
  {
    key: 'reorder' as const,
    icon: '🔄',
    title: '기존 요약본 재정렬',
    desc: '이미 생성된 PPT 요약본의 슬라이드 순서를 엑셀 기준으로 재배치합니다',
    tags: ['PPT 입력', '엑셀 순번', '재정렬'],
  },
];

export default function WorkTypeSelector({ workType, onSelect }: WorkTypeSelectorProps) {
  return (
    <div className="section-block fade-in">
      <div className="section-label">
        <span className="section-number">1</span>
        작업 유형을 선택하세요
      </div>
      <div className="work-type-grid">
        {WORK_TYPES.map((wt) => (
          <div
            key={wt.key}
            className={`work-type-card ${workType === wt.key ? 'selected' : ''}`}
            onClick={() => onSelect(wt.key)}
          >
            <div className="work-type-check">
              {workType === wt.key ? '✓' : ''}
            </div>
            <div className="work-type-icon">{wt.icon}</div>
            <div className="work-type-title">{wt.title}</div>
            <div className="work-type-desc">{wt.desc}</div>
            <div className="work-type-tags">
              {wt.tags.map((tag, i) => (
                <span key={i} className="work-type-tag">{tag}</span>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
