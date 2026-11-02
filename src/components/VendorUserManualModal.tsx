import React, { useState } from 'react';
import {
  X,
  Printer,
  Copy,
  Download,
  CheckCircle2,
  AlertTriangle,
  Lock,
  Mail,
  FileText,
  Receipt,
  CreditCard,
  Percent,
  FileSpreadsheet,
  Building2,
  FileSignature,
  Stamp,
  ShieldAlert,
  ArrowRight,
  HelpCircle,
  ExternalLink,
  ChevronDown,
  ChevronUp,
  Layers,
  Sparkles,
  Info
} from 'lucide-react';
import * as XLSX from 'xlsx';

interface VendorUserManualModalProps {
  key?: React.Key;
  isOpen: boolean;
  onClose: () => void;
  vendorName?: string;
}

export default function VendorUserManualModal({
  isOpen,
  onClose,
  vendorName
}: VendorUserManualModalProps) {
  const [activeTab, setActiveTab] = useState<'all' | 'step1' | 'step2' | 'step3' | 'step4' | 'step5' | 'step6'>('all');
  const [copySuccess, setCopySuccess] = useState(false);
  const [openFaq, setOpenFaq] = useState<number | null>(null);

  if (!isOpen) return null;

  // Print function
  const handlePrint = () => {
    window.print();
  };

  // Download sample Excel template for vendors
  const handleDownloadExcelTemplate = () => {
    const templateHeaders = [
      '품번',
      '품목명',
      '규격',
      '단위',
      '제조사', // 메이커 필수 강조!
      '협가',
      '네고율',
      '네고방식',
      '단중',
      '비고'
    ];

    const sampleRows = [
      ['VALVE-001', '볼밸브 (10K)', '50A', 'EA', '삼양', 25000, 10, 'percent', 0, 'KS 규격품 (메이커 필히 기재)'],
      ['VALVE-002', '글로브밸브 (10K)', '65A', 'EA', '동양', 48000, 8.5, 'percent', 0, '나사식 (제조사 기재 필수)'],
      ['PIPE-001', 'STS304 파이프', '100A Sch10S', 'M', '세아제강', 0, 4200, 'sts_pipe', 3.82, 'STS 단중단가 적용 (원/kg)'],
      ['FITTING-001', '90도 엘보 (ELBOW)', '50A', 'EA', '태광', 8500, 12, 'percent', 0, '용접식']
    ];

    const worksheet = XLSX.utils.aoa_to_sheet([templateHeaders, ...sampleRows]);
    
    // Column widths
    worksheet['!cols'] = [
      { wch: 14 }, // 품번
      { wch: 22 }, // 품목명
      { wch: 18 }, // 규격
      { wch: 8 },  // 단위
      { wch: 16 }, // 제조사 (Maker)
      { wch: 14 }, // 협가
      { wch: 12 }, // 네고율
      { wch: 14 }, // 네고방식
      { wch: 10 }, // 단중
      { wch: 28 }  // 비고
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '단가표업로드양식');
    XLSX.writeFile(workbook, `명신기공_협력사_단가표_표준양식(메이커필수).xlsx`);
  };

  // Copy distribution summary for KakaoTalk/SMS/Email
  const handleCopyDistributionText = () => {
    const text = `[(주)명신기공 협력사(납품업체) 전용 포탈 사용 안내]

안녕하십니까, (주)명신기공 협력사 담당자님.
당사와의 투명하고 신속한 단가 협의 및 발주 업무를 위해 전용 포탈 사용방법을 안내해 드립니다.

1. 최초 로그인 및 비밀번호 변경 (필수/보안)
 - 최초 임시 비밀번호: 1234
 - 🚨 미변경 시 타 업체가 귀사 단가/서류를 볼 수 있으므로 최초 접속 즉시 우측 상단 [비밀번호 변경]을 통해 새 암호로 변경해 주십시오.

2. 발주서 수신용 공식 이메일 등록
 - [업체정보 수정]에서 이메일을 실제 '발주서(PO)를 수신할 담당자/부서 이메일'로 정확하게 기재해 주십시오.

3. 필수 증빙서류 3종 첨부
 - 상단 증빙서류 바에서 [사업자등록증], [통장사본], [전자어음약정확인서]를 첨부해 주십시오. (PDF 또는 사진)

4. 품목별 단가표 및 요율(네고율) 작성법
 - 품목별 협가(기본가격) 입력 후 네고율(%)을 입력하면 납품단가가 자동 계산됩니다. (STS 파이프는 단중단가 방식 지원)
 - 🚨 [매우 중요] 품목별 '제조사(메이커)'는 필히 기재하셔야 합니다! 제조사 미기재 시 규격 검증이 불가하여 발주 채택에서 누락될 수 있습니다.

5. 전자 계약서 상호 날인 회신
 - 명신기공 직인 계약서 다운로드 -> 거래처 도장 날인 후 스캔하여 업로드 회신해 주십시오. (회신 전까지 알림 발생)

문의: (주)명신기공 원자재관리팀 (TEL: 02-421-1392 / FAX: 02-422-4171)`;

    navigator.clipboard.writeText(text);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2500);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-2 sm:p-4 backdrop-blur-sm overflow-y-auto">
      <div 
        id="printable-vendor-manual"
        className="w-full max-w-5xl bg-white rounded-3xl shadow-2xl flex flex-col max-h-[94vh] overflow-hidden border border-slate-200"
      >
        {/* Top Header (Modal Header) - Hidden in Print */}
        <div className="bg-slate-900 text-white px-6 py-4.5 flex items-center justify-between border-b border-slate-800 shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-indigo-600 flex items-center justify-center shadow-lg shadow-indigo-500/30 text-white">
              <FileSpreadsheet className="h-5 w-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-black uppercase tracking-widest text-indigo-400 bg-indigo-950/80 px-2 py-0.5 rounded-full border border-indigo-800/60">
                  협력사 배포용 매뉴얼
                </span>
                <span className="text-xs text-slate-400">Ver 2.0 (실무자용)</span>
              </div>
              <h2 className="text-lg sm:text-xl font-black tracking-tight text-white flex items-center gap-2 mt-0.5">
                (주)명신기공 협력사 단가관리 시스템 사용설명서
              </h2>
            </div>
          </div>

          <div className="flex items-center gap-2 print:hidden">
            <button
              onClick={handlePrint}
              className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold transition-all border border-slate-700 active:scale-95 shadow-sm"
              title="A4 규격 인쇄 또는 PDF로 저장"
            >
              <Printer className="h-3.5 w-3.5 text-indigo-400" />
              <span>인쇄 / PDF 저장</span>
            </button>
            <button
              onClick={handleCopyDistributionText}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold transition-all border active:scale-95 shadow-sm ${
                copySuccess 
                  ? 'bg-emerald-600 border-emerald-500 text-white' 
                  : 'bg-indigo-600 hover:bg-indigo-500 border-indigo-500 text-white'
              }`}
              title="거래처 발송용 요약 안내문 복사"
            >
              {copySuccess ? <CheckCircle2 className="h-3.5 w-3.5" /> : <Copy className="h-3.5 w-3.5" />}
              <span>{copySuccess ? '안내문 복사됨!' : '카톡/문자 안내문 복사'}</span>
            </button>
            <button
              onClick={onClose}
              className="p-1.5 hover:bg-slate-800 rounded-full text-slate-400 hover:text-white transition-colors ml-1"
            >
              <X className="h-5 w-5" />
            </button>
          </div>
        </div>

        {/* Action & Quick Guide Ribbon (Print: Hidden) */}
        <div className="bg-slate-50 border-b border-slate-200 px-6 py-3 shrink-0 flex flex-wrap items-center justify-between gap-3 print:hidden">
          <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar py-0.5">
            {[
              { id: 'all', label: '전체 펼쳐보기 (완독)' },
              { id: 'step1', label: '1. 비밀번호 변경 (필수)' },
              { id: 'step2', label: '2. 발주 이메일 등록' },
              { id: 'step3', label: '3. 증빙서류 3종' },
              { id: 'step4', label: '4. 메이커 필수 입력' },
              { id: 'step5', label: '5. 단가 & 요율 계산법' },
              { id: 'step6', label: '6. 전자계약서 날인' }
            ].map((tab, idx) => (
              <button
                key={`manual-tab-${tab.id}-${idx}`}
                onClick={() => setActiveTab(tab.id as any)}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${
                  activeTab === tab.id
                    ? 'bg-indigo-600 text-white shadow-sm shadow-indigo-200'
                    : 'bg-white text-slate-600 hover:bg-slate-200/70 border border-slate-200/80'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleDownloadExcelTemplate}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-xl text-xs font-black transition-all shadow-xs"
              title="단가표 엑셀 업로드용 샘플 양식 다운로드"
            >
              <Download className="h-3.5 w-3.5 text-emerald-600" />
              <span>표준 엑셀 양식 다운로드</span>
            </button>
          </div>
        </div>

        {/* Scrollable Document Body */}
        <div className="p-6 overflow-y-auto space-y-8 text-slate-700 leading-relaxed custom-scrollbar">

          {/* OFFICIAL HEADER (Visible in Print & Screen) */}
          <div className="border-b-2 border-slate-900 pb-4 flex flex-col sm:flex-row sm:items-end justify-between gap-3">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-sm font-black text-indigo-700 tracking-wider">(주)명신기공 구매관리팀 공문</span>
                <span className="text-xs text-slate-400">| 협력사 배포용 업무 매뉴얼</span>
              </div>
              <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight mt-1">
                협력사 단가관리 및 전자계약 포탈 사용설명서
              </h1>
              <p className="text-xs text-slate-500 mt-1">
                본 문서는 (주)명신기공과 거래하는 모든 협력사의 신속하고 안전한 단가 등록, 증빙서류 관리, 발주서 수신을 돕기 위해 제작되었습니다.
              </p>
            </div>
            <div className="text-left sm:text-right shrink-0 bg-slate-50 p-3 rounded-xl border border-slate-200 text-xs font-medium">
              <div>대상: <strong>(주)명신기공 등록 전체 거래처</strong></div>
              <div>시행일자: <strong>상시 적용</strong></div>
              <div className="text-indigo-600 font-bold mt-0.5">※ 최초 수령 후 즉시 정독 요망</div>
            </div>
          </div>

          {/* 🚨 3대 핵심 필수 수칙 하이라이트 배너 */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* Warning 1 */}
            <div className="p-4.5 rounded-2xl bg-rose-50/80 border-2 border-rose-300 shadow-sm relative overflow-hidden">
              <div className="flex items-center gap-2.5 mb-2">
                <div className="w-8 h-8 rounded-xl bg-rose-600 text-white flex items-center justify-center font-black shrink-0">
                  <Lock className="h-4 w-4" />
                </div>
                <div>
                  <span className="text-[10px] font-black uppercase tracking-wider text-rose-700 bg-rose-100 px-1.5 py-0.5 rounded">필수 1. 보안</span>
                  <h4 className="text-sm font-black text-rose-950">최초 비밀번호 즉시 변경</h4>
                </div>
              </div>
              <p className="text-xs text-rose-900 font-semibold leading-relaxed">
                최초 임시 비밀번호는 <strong className="underline decoration-rose-400 decoration-2 font-mono text-rose-950 text-[13px]">1234</strong>입니다.
                미변경 시 타 업체가 귀사 단가를 열람할 수 있으므로 <strong className="text-rose-950">접속 즉시 변경</strong>하십시오.
              </p>
            </div>

            {/* Warning 2 */}
            <div className="p-4.5 rounded-2xl bg-amber-50/80 border-2 border-amber-300 shadow-sm relative overflow-hidden">
              <div className="flex items-center gap-2.5 mb-2">
                <div className="w-8 h-8 rounded-xl bg-amber-600 text-white flex items-center justify-center font-black shrink-0">
                  <Mail className="h-4 w-4" />
                </div>
                <div>
                  <span className="text-[10px] font-black uppercase tracking-wider text-amber-800 bg-amber-100 px-1.5 py-0.5 rounded">필수 2. 발주</span>
                  <h4 className="text-sm font-black text-amber-950">발주서 수신용 이메일 등록</h4>
                </div>
              </div>
              <p className="text-xs text-amber-900 font-semibold leading-relaxed">
                공식 이메일란에는 <strong className="underline decoration-amber-400 decoration-2 text-amber-950">실제 발주서(PO)를 수신할 담당자/부서 이메일</strong>을 정확히 입력하셔야 발주 누락이 발생하지 않습니다.
              </p>
            </div>

            {/* Warning 3 */}
            <div className="p-4.5 rounded-2xl bg-indigo-50/80 border-2 border-indigo-300 shadow-sm relative overflow-hidden">
              <div className="flex items-center gap-2.5 mb-2">
                <div className="w-8 h-8 rounded-xl bg-indigo-600 text-white flex items-center justify-center font-black shrink-0">
                  <Stamp className="h-4 w-4" />
                </div>
                <div>
                  <span className="text-[10px] font-black uppercase tracking-wider text-indigo-700 bg-indigo-100 px-1.5 py-0.5 rounded">필수 3. 품질</span>
                  <h4 className="text-sm font-black text-indigo-950">품목별 메이커(제조사) 필수</h4>
                </div>
              </div>
              <p className="text-xs text-indigo-900 font-semibold leading-relaxed">
                단가표 및 엑셀 작성 시 <strong className="underline decoration-indigo-400 decoration-2 text-indigo-950">제조사(Maker)명은 필히 기재</strong>해야 합니다. 메이커 미기재 시 규격 미달로 발주 선정에서 제외됩니다.
              </p>
            </div>
          </div>

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* STEP 1. 최초 로그인 및 비밀번호 변경 (보안 필수) */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          {(activeTab === 'all' || activeTab === 'step1') && (
            <section className="p-6 rounded-3xl bg-white border-2 border-slate-200/90 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-slate-900 text-white font-black text-sm flex items-center justify-center">
                    1
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase tracking-wider text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-full">
                      STEP 1. 접속 및 보안
                    </span>
                    <h3 className="text-lg font-black text-slate-900">
                      최초 접속 및 비밀번호 즉시 변경 (최우선 필수)
                    </h3>
                  </div>
                </div>
                <span className="text-xs font-black text-rose-600 bg-rose-50 border border-rose-200 px-2.5 py-1 rounded-full flex items-center gap-1">
                  <AlertTriangle className="h-3.5 w-3.5" />
                  미변경 시 보안 경고
                </span>
              </div>

              <div className="space-y-3 text-xs leading-relaxed">
                <p>
                  당사에서 발송해 드린 거래처 전용 접속 링크로 접속하시거나, 메인 화면 좌측 <strong>[거래처 목록]</strong>에서 귀사 명칭을 클릭합니다.
                </p>

                {/* Visual Diagram: Password Flow */}
                <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                  <div className="text-[11px] font-black text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                    <Layers className="h-3.5 w-3.5 text-indigo-600" />
                    [시각 도식] 비밀번호 변경 3단계 프로세스
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-center">
                    {/* Box 1 */}
                    <div className="p-3 bg-white rounded-xl border border-slate-200 shadow-xs flex flex-col items-center justify-center">
                      <div className="w-7 h-7 rounded-full bg-slate-100 text-slate-700 font-bold flex items-center justify-center text-xs mb-1.5">
                        1
                      </div>
                      <span className="text-[11px] font-black text-slate-800">최초 임시 암호 입력</span>
                      <div className="mt-2 px-3 py-1 bg-rose-50 border border-rose-200 rounded-lg text-rose-700 font-mono font-black text-sm">
                        1234
                      </div>
                      <span className="text-[10px] text-slate-400 mt-1">시스템 기본 부여 비밀번호</span>
                    </div>

                    {/* Arrow / Box 2 */}
                    <div className="p-3 bg-white rounded-xl border-2 border-indigo-200 shadow-xs flex flex-col items-center justify-center relative">
                      <div className="w-7 h-7 rounded-full bg-indigo-100 text-indigo-700 font-bold flex items-center justify-center text-xs mb-1.5">
                        2
                      </div>
                      <span className="text-[11px] font-black text-indigo-900">우측 상단 [암호변경] 클릭</span>
                      <div className="mt-2 px-3 py-1 bg-indigo-50 border border-indigo-200 rounded-lg text-indigo-700 font-bold text-xs flex items-center gap-1">
                        <Lock className="h-3 w-3" />
                        <span>비밀번호 변경 창 열림</span>
                      </div>
                      <span className="text-[10px] text-slate-500 mt-1">기존 암호: 1234 입력</span>
                    </div>

                    {/* Box 3 */}
                    <div className="p-3 bg-emerald-50 rounded-xl border-2 border-emerald-300 shadow-xs flex flex-col items-center justify-center">
                      <div className="w-7 h-7 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center text-xs mb-1.5">
                        3
                      </div>
                      <span className="text-[11px] font-black text-emerald-950">새 고유 암호 설정 완료</span>
                      <div className="mt-2 px-3 py-1 bg-white border border-emerald-300 rounded-lg text-emerald-700 font-mono font-black text-sm">
                        ••••••••
                      </div>
                      <span className="text-[10px] text-emerald-800 font-bold mt-1">✓ 타업체 열람 차단 및 보안 완성</span>
                    </div>
                  </div>
                </div>

                <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl text-rose-900 text-xs font-medium space-y-1">
                  <div className="font-black flex items-center gap-1.5 text-rose-800">
                    <ShieldAlert className="h-4 w-4 text-rose-600 shrink-0" />
                    왜 비밀번호 변경이 필수인가요?
                  </div>
                  <p className="leading-relaxed pl-5">
                    명신기공 시스템은 다수의 협력사가 함께 이용하므로, 초기 비밀번호인 <strong className="font-mono">1234</strong>를 그대로 방치할 경우 
                    <strong> 타 업체 담당자가 귀사 명칭을 누르고 단가표, 할인율, 계약서를 열람할 수 있는 심각한 영업비밀 유출 위험</strong>이 있습니다. 
                    따라서 <strong>최초 접속 시 즉시 귀사 전담자만 아는 4자리 이상 신규 비밀번호로 변경</strong>해 주셔야 합니다.
                  </p>
                </div>
              </div>
            </section>
          )}

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* STEP 2. 발주서 수신용 공식 이메일 등록 */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          {(activeTab === 'all' || activeTab === 'step2') && (
            <section className="p-6 rounded-3xl bg-white border-2 border-slate-200/90 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-slate-900 text-white font-black text-sm flex items-center justify-center">
                    2
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase tracking-wider text-amber-600 bg-amber-50 px-2 py-0.5 rounded-full">
                      STEP 2. 기본 정보 점검
                    </span>
                    <h3 className="text-lg font-black text-slate-900">
                      발주서(PO) 수신용 공식 이메일 및 기본 정보 등록
                    </h3>
                  </div>
                </div>
                <span className="text-xs font-black text-amber-700 bg-amber-50 border border-amber-200 px-2.5 py-1 rounded-full">
                  발주 누락 방지
                </span>
              </div>

              <div className="space-y-3 text-xs leading-relaxed">
                <p>
                  단가표 화면 상단 우측의 <strong>[업체정보 수정]</strong> 버튼을 누르면 귀사의 사업자 기본 정보 입력창이 나타납니다.
                </p>

                {/* Form UI Representation Card */}
                <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                  <div className="text-[11px] font-black text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                    <Building2 className="h-3.5 w-3.5 text-amber-600" />
                    [화면 목업] 업체 정보 등록 필수 항목
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-left">
                    <div className="p-3 bg-white rounded-xl border border-slate-200">
                      <span className="text-[10px] font-bold text-slate-400 block mb-1">상호명 및 대표자</span>
                      <div className="font-bold text-slate-800 text-xs">정식 법인/상호명 & 대표자 성함 입력</div>
                    </div>

                    <div className="p-3 bg-amber-50/70 rounded-xl border-2 border-amber-300">
                      <div className="flex items-center justify-between mb-1">
                        <span className="text-[10px] font-black text-amber-800">공식 이메일 (🚨 최우선 필수)</span>
                        <span className="text-[9px] bg-amber-200 text-amber-900 px-1.5 py-0.2 rounded font-black">발주서 수신용</span>
                      </div>
                      <div className="font-mono font-black text-slate-900 text-xs">
                        order@company.com / 담당자 이메일
                      </div>
                      <p className="text-[10px] text-amber-800 mt-1 leading-tight">
                        ※ (주)명신기공 전자 발주서(PO)가 자동 전송되는 주소이므로 반드시 매일 수신 확인 가능한 공식 메일로 입력!
                      </p>
                    </div>

                    <div className="p-3 bg-white rounded-xl border border-slate-200">
                      <span className="text-[10px] font-bold text-slate-400 block mb-1">전화번호 / 팩스번호</span>
                      <div className="font-mono font-bold text-slate-800 text-xs">02-000-0000 / 02-000-0001</div>
                    </div>

                    <div className="p-3 bg-white rounded-xl border border-slate-200">
                      <span className="text-[10px] font-bold text-slate-400 block mb-1">사업자등록번호</span>
                      <div className="font-mono font-bold text-slate-800 text-xs">000-00-00000 (세금계산서 발행용)</div>
                    </div>
                  </div>
                </div>

                <p className="text-slate-600 text-[11.5px]">
                  💡 <strong>팁:</strong> 비고란(메모)에는 '배송지 기준 (공장/현장 직납 여부)', '납기 소요일수 (예: 주문 후 3일 이내)', '최소 주문 수량(MOQ)' 등의 특이사항을 적어주시면 구매팀에서 발주 시 적극 반영합니다.
                </p>
              </div>
            </section>
          )}

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* STEP 3. 필수 증빙서류 3종 첨부 */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          {(activeTab === 'all' || activeTab === 'step3') && (
            <section className="p-6 rounded-3xl bg-white border-2 border-slate-200/90 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-slate-900 text-white font-black text-sm flex items-center justify-center">
                    3
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase tracking-wider text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-full">
                      STEP 3. 거래 증빙
                    </span>
                    <h3 className="text-lg font-black text-slate-900">
                      필수 증빙서류 3종 첨부 (사업자등록증, 통장사본, 전자어음약정서)
                    </h3>
                  </div>
                </div>
                <span className="text-xs font-black text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-full">
                  정산/결제 필수
                </span>
              </div>

              <div className="space-y-3 text-xs leading-relaxed">
                <p>
                  거래 개시 및 대금 결제 계좌 검증을 위해 상단 바의 <strong>[필수 증빙서류]</strong> 섹션에서 다음 3종 서류를 반드시 업로드해 주십시오.
                  PDF 파일뿐만 아니라 <strong>스마트폰 카메라로 촬영한 선명한 사진(JPG, PNG)</strong>도 그대로 업로드하실 수 있습니다.
                </p>

                {/* 3 Document Cards Visualization */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  {/* Doc 1 */}
                  <div className="p-4 rounded-2xl bg-slate-50 border-2 border-slate-200 hover:border-emerald-400 transition-all text-center space-y-2">
                    <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 mx-auto flex items-center justify-center">
                      <Receipt className="h-5 w-5" />
                    </div>
                    <div>
                      <h5 className="font-black text-slate-800 text-xs">1. 사업자등록증 사본</h5>
                      <p className="text-[10.5px] text-slate-500 mt-0.5">최신 갱신된 정식 사업자등록증</p>
                    </div>
                    <div className="pt-1 text-[10px] font-bold text-emerald-600 bg-emerald-50 py-1 rounded-lg border border-emerald-100">
                      PDF, JPG, PNG 첨부
                    </div>
                  </div>

                  {/* Doc 2 */}
                  <div className="p-4 rounded-2xl bg-slate-50 border-2 border-slate-200 hover:border-emerald-400 transition-all text-center space-y-2">
                    <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 mx-auto flex items-center justify-center">
                      <CreditCard className="h-5 w-5" />
                    </div>
                    <div>
                      <h5 className="font-black text-slate-800 text-xs">2. 결제 계좌 통장사본</h5>
                      <p className="text-[10.5px] text-slate-500 mt-0.5">물품대금 입금용 법인/대표자 계좌</p>
                    </div>
                    <div className="pt-1 text-[10px] font-bold text-emerald-600 bg-emerald-50 py-1 rounded-lg border border-emerald-100">
                      계좌번호/예금주 확인용
                    </div>
                  </div>

                  {/* Doc 3 */}
                  <div className="p-4 rounded-2xl bg-slate-50 border-2 border-slate-200 hover:border-emerald-400 transition-all text-center space-y-2">
                    <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 mx-auto flex items-center justify-center">
                      <FileText className="h-5 w-5" />
                    </div>
                    <div>
                      <h5 className="font-black text-slate-800 text-xs">3. 전자어음약정확인서</h5>
                      <p className="text-[10.5px] text-slate-500 mt-0.5">주거래은행 전자어음 약정 체결 증빙</p>
                    </div>
                    <div className="pt-1 text-[10px] font-bold text-emerald-600 bg-emerald-50 py-1 rounded-lg border border-emerald-100">
                      어음 거래 시 필수 증서
                    </div>
                  </div>
                </div>

                <div className="p-3 bg-emerald-50/60 rounded-xl border border-emerald-200 text-emerald-900 text-xs flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <CheckCircle2 className="h-4 w-4 text-emerald-600 shrink-0" />
                    <span>서류 등록 후 <strong>[미리보기]</strong> 버튼을 누르면 언제든 정상 등록 여부를 즉시 열람 및 확인할 수 있습니다.</span>
                  </div>
                </div>
              </div>
            </section>
          )}

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* STEP 4. 품목별 메이커(제조사) 필수 입력 강력 강조! */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          {(activeTab === 'all' || activeTab === 'step4') && (
            <section className="p-6 rounded-3xl bg-indigo-50/50 border-2 border-indigo-400 shadow-md space-y-4">
              <div className="flex items-center justify-between border-b border-indigo-200 pb-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-indigo-700 text-white font-black text-sm flex items-center justify-center shadow-sm">
                    4
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase tracking-wider text-indigo-700 bg-indigo-100 px-2 py-0.5 rounded-full">
                      STEP 4. 핵심 품질 규격
                    </span>
                    <h3 className="text-lg font-black text-indigo-950">
                      품목별 메이커(제조사) 필수 작성 (🚨 필독 강력 강조)
                    </h3>
                  </div>
                </div>
                <span className="text-xs font-black text-white bg-indigo-600 px-3 py-1 rounded-full shadow-sm animate-pulse">
                  발주 선정 필수 조건
                </span>
              </div>

              <div className="space-y-4 text-xs leading-relaxed text-slate-800">
                <div className="p-4 bg-white rounded-2xl border-2 border-indigo-300 shadow-sm space-y-2">
                  <h4 className="text-sm font-black text-indigo-900 flex items-center gap-2">
                    <AlertTriangle className="h-4 w-4 text-amber-500 shrink-0" />
                    메이커(제조사)는 왜 필히 작성해야 하나요?
                  </h4>
                  <p className="leading-relaxed text-slate-600">
                    (주)명신기공은 고객사 납품 품질 및 안전 검증을 위해 <strong>'단가 통합 비교 매트릭스'</strong>에서 전 업체의 동일 품명/규격 품목을 제조사별로 교차 비교합니다.
                    <br />
                    만약 단가표에 <strong className="text-rose-600 underline">제조사(Maker)가 공란이거나 누락되어 있으면 규격 인증을 통과하지 못해 발주 선정 대상에서 자동으로 제외</strong>되오니,
                    모든 품목의 제조사를 빠짐없이 기재해 주시기 바랍니다.
                  </p>
                </div>

                {/* Good vs Bad Example Visual Table */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {/* Bad Example */}
                  <div className="p-4 rounded-2xl bg-rose-50/70 border border-rose-200 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-black text-rose-800 flex items-center gap-1">
                        <X className="h-4 w-4 text-rose-600" />
                        잘못된 작성 예시 (발주 탈락)
                      </span>
                      <span className="text-[9px] bg-rose-200 text-rose-900 px-2 py-0.5 rounded font-black">불가</span>
                    </div>
                    <div className="bg-white p-3 rounded-xl border border-rose-200 text-[11px] font-mono space-y-1">
                      <div className="flex justify-between"><span>품명:</span> <strong className="text-slate-800">볼밸브 10K</strong></div>
                      <div className="flex justify-between"><span>규격:</span> <strong className="text-slate-800">50A</strong></div>
                      <div className="flex justify-between text-rose-600 font-black bg-rose-50 px-1 py-0.5 rounded">
                        <span>제조사(Maker):</span> <span>[ 공란 / 미입력 ] ❌</span>
                      </div>
                      <div className="flex justify-between text-slate-400"><span>단가:</span> <span>₩25,000</span></div>
                    </div>
                    <p className="text-[10px] text-rose-700 leading-tight">
                      ❌ 제조사 불명확으로 규격 확인 불가 ➡️ 비교표에서 비신뢰 품목으로 분류
                    </p>
                  </div>

                  {/* Good Example */}
                  <div className="p-4 rounded-2xl bg-emerald-50/70 border-2 border-emerald-400 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-black text-emerald-800 flex items-center gap-1">
                        <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                        올바른 작성 예시 (발주 적극 채택)
                      </span>
                      <span className="text-[9px] bg-emerald-600 text-white px-2 py-0.5 rounded font-black">추천</span>
                    </div>
                    <div className="bg-white p-3 rounded-xl border border-emerald-200 text-[11px] font-mono space-y-1">
                      <div className="flex justify-between"><span>품명:</span> <strong className="text-slate-800">볼밸브 10K</strong></div>
                      <div className="flex justify-between"><span>규격:</span> <strong className="text-slate-800">50A</strong></div>
                      <div className="flex justify-between text-emerald-700 font-black bg-emerald-50 px-1 py-0.5 rounded">
                        <span>제조사(Maker):</span> <span>삼양 (또는 동양, 효성 등) ⭕</span>
                      </div>
                      <div className="flex justify-between text-slate-600 font-bold"><span>단가:</span> <span>₩25,000</span></div>
                    </div>
                    <p className="text-[10px] text-emerald-800 font-bold leading-tight">
                      ⭕ 정확한 브랜드/제조사 기재 ➡️ 신속한 구매 심사 및 발주서 발행!
                    </p>
                  </div>
                </div>

                <div className="text-[11px] text-indigo-900 bg-white p-3 rounded-xl border border-indigo-200 flex items-center gap-2">
                  <Info className="h-4 w-4 text-indigo-600 shrink-0" />
                  <span>
                    제조사가 여러 개인 경우 대표 제조사를 기재하거나 비고란에 '삼양 또는 동양 동등 이상품'으로 명시해 주셔도 됩니다.
                  </span>
                </div>
              </div>
            </section>
          )}

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* STEP 5. 취급 품목 단가표 작성 및 요율(네고율) 계산 방법 */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          {(activeTab === 'all' || activeTab === 'step5') && (
            <section className="p-6 rounded-3xl bg-white border-2 border-slate-200/90 shadow-sm space-y-5">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-slate-900 text-white font-black text-sm flex items-center justify-center">
                    5
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase tracking-wider text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-full">
                      STEP 5. 단가 & 요율
                    </span>
                    <h3 className="text-lg font-black text-slate-900">
                      취급 품목 단가표 작성 및 요율(네고율) 아주 쉬운 계산법
                    </h3>
                  </div>
                </div>
                <span className="text-xs font-black text-indigo-600 bg-indigo-50 border border-indigo-200 px-2.5 py-1 rounded-full">
                  자동 연산 시스템
                </span>
              </div>

              <div className="space-y-4 text-xs leading-relaxed">
                <p>
                  (주)명신기공 시스템은 <strong>기본 협가(정가 또는 원가)</strong>를 입력하고 <strong>요율(네고율)</strong>을 입력하면
                  최종 납품단가가 실시간으로 자동 계산되어 화면에 즉시 표시됩니다.
                </p>

                {/* 2 Calculation Modes Comparison Cards */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* Mode 1: % 할인 방식 */}
                  <div className="p-5 rounded-2xl bg-indigo-50/50 border-2 border-indigo-200 flex flex-col justify-between space-y-3">
                    <div>
                      <div className="flex items-center justify-between mb-2">
                        <div className="flex items-center gap-2">
                          <span className="w-6 h-6 rounded-lg bg-indigo-600 text-white font-black text-xs flex items-center justify-center">
                            %
                          </span>
                          <h4 className="text-sm font-black text-indigo-950">방식 A. 일반 % 할인율 (Percent)</h4>
                        </div>
                        <span className="text-[10px] bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded font-bold">기본 방식</span>
                      </div>
                      <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                        밸브류, 피팅류, 일반 자재 등 대부분의 제품에 적용되는 가장 대중적인 정률 할인 방식입니다.
                      </p>

                      {/* Formula Card */}
                      <div className="bg-white p-3.5 rounded-xl border border-indigo-100 font-mono space-y-1.5 shadow-xs">
                        <div className="text-[10.5px] text-indigo-700 font-bold">공식: 최종단가 = 협가 × (1 - 네고율%)</div>
                        <div className="text-xs text-slate-700 pt-1 border-t border-slate-100 space-y-1">
                          <div className="flex justify-between">
                            <span>1. 기본 협가(원가):</span>
                            <strong>₩10,000</strong>
                          </div>
                          <div className="flex justify-between text-indigo-600 font-bold">
                            <span>2. 네고율(할인율):</span>
                            <span>10 % (₩1,000 할인)</span>
                          </div>
                          <div className="flex justify-between text-emerald-600 font-black text-sm pt-1 border-t border-slate-150">
                            <span>3. 최종 납품단가:</span>
                            <span>₩9,000</span>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="text-[10.5px] text-indigo-800 bg-white/80 p-2.5 rounded-lg border border-indigo-100">
                      💡 <strong>적용법:</strong> 네고율 입력창에 숫자 <code className="font-bold text-indigo-600">10</code>만 입력하면 납품단가가 ₩9,000으로 즉시 연산됩니다.
                    </div>
                  </div>

                  {/* Mode 2: STS 파이프 방식 */}
                  <div className="p-5 rounded-2xl bg-emerald-50/50 border-2 border-emerald-200 flex flex-col justify-between space-y-3">
                    <div>
                      <div className="flex items-center justify-between mb-2">
                        <div className="flex items-center gap-2">
                          <span className="w-6 h-6 rounded-lg bg-emerald-600 text-white font-black text-xs flex items-center justify-center">
                            STS
                          </span>
                          <h4 className="text-sm font-black text-emerald-950">방식 B. STS 파이프 방식 (KG단가)</h4>
                        </div>
                        <span className="text-[10px] bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-bold">배관/원자재</span>
                      </div>
                      <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                        스테인리스 파이프나 철강재처럼 미터당 단중(kg)에 kg당 단가를 곱하여 산출하는 방식입니다.
                      </p>

                      {/* Formula Card */}
                      <div className="bg-white p-3.5 rounded-xl border border-emerald-100 font-mono space-y-1.5 shadow-xs">
                        <div className="text-[10.5px] text-emerald-700 font-bold">공식: 최종단가 = 단중(kg) × 단중당단가(원/kg)</div>
                        <div className="text-xs text-slate-700 pt-1 border-t border-slate-100 space-y-1">
                          <div className="flex justify-between">
                            <span>1. 규격별 단중:</span>
                            <strong>3.50 kg</strong>
                          </div>
                          <div className="flex justify-between text-indigo-600 font-bold">
                            <span>2. STS 단중단가:</span>
                            <span>₩4,000 / kg</span>
                          </div>
                          <div className="flex justify-between text-emerald-600 font-black text-sm pt-1 border-t border-slate-150">
                            <span>3. 최종 납품단가:</span>
                            <span>3.5 × 4,000 = ₩14,000</span>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="text-[10.5px] text-emerald-900 bg-white/80 p-2.5 rounded-lg border border-emerald-100">
                      💡 <strong>전환법:</strong> 네고율 옆의 <strong className="px-1.5 py-0.5 bg-slate-200 rounded font-bold">[%]</strong> 버튼을 클릭하면 <strong className="px-1.5 py-0.5 bg-indigo-100 text-indigo-700 rounded font-bold">[STS]</strong> 모드로 즉시 변경됩니다.
                    </div>
                  </div>
                </div>

                {/* 10-Won Rounding & Registration Method */}
                <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                  <div className="flex items-center gap-2 font-black text-slate-800 text-xs">
                    <CheckCircle2 className="h-4 w-4 text-indigo-600" />
                    <span>단가표 등록하는 2가지 편리한 방법</span>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                    <div className="p-3 bg-white rounded-xl border border-slate-200 space-y-1">
                      <strong className="text-slate-800 block text-xs">방법 1. 화면에서 직접 입력 (1~20건 추천)</strong>
                      <p className="text-slate-500 text-[11px] leading-relaxed">
                        화면 우측 상단의 <strong>[+ 품목 추가]</strong> 버튼을 눌러 품명, 규격, 단위, 제조사, 협가, 요율을 기재하고 즉시 저장합니다.
                      </p>
                    </div>

                    <div className="p-3 bg-white rounded-xl border border-slate-200 space-y-1">
                      <strong className="text-slate-800 block text-xs">방법 2. 엑셀 일괄 업로드 (수십~수백 건 추천)</strong>
                      <p className="text-slate-500 text-[11px] leading-relaxed">
                        상단의 <strong>[표준 엑셀 양식 다운로드]</strong>를 클릭하여 품목 데이터를 기재한 뒤, 
                        <strong>[엑셀 업로드]</strong> 버튼을 통해 드래그 앤 드롭하면 1초 만에 전체 등록됩니다.
                      </p>
                    </div>
                  </div>

                  <div className="text-[10.5px] text-slate-500 pt-1 border-t border-slate-200">
                    ※ <strong>10원 단위 반올림 옵션:</strong> 단가 옆의 작은 체크박스를 선택하면, 1원 단위의 끝전을 10원 단위로 깔끔하게 반올림하여 회계 및 정산의 오차를 방지합니다.
                  </div>
                </div>
              </div>
            </section>
          )}

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* STEP 6. 전자 계약서 직인/도장 날인 회신 및 체결 */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          {(activeTab === 'all' || activeTab === 'step6') && (
            <section className="p-6 rounded-3xl bg-white border-2 border-slate-200/90 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-slate-900 text-white font-black text-sm flex items-center justify-center">
                    6
                  </div>
                  <div>
                    <span className="text-[10px] font-black uppercase tracking-wider text-rose-600 bg-rose-50 px-2 py-0.5 rounded-full">
                      STEP 6. 계약 체결
                    </span>
                    <h3 className="text-lg font-black text-slate-900">
                      전자 계약서 직인/도장 날인 회신 (알림 자동 해제 절차)
                    </h3>
                  </div>
                </div>
                <span className="text-xs font-black text-rose-700 bg-rose-50 border border-rose-200 px-2.5 py-1 rounded-full">
                  쌍방 날인 체결
                </span>
              </div>

              <div className="space-y-3 text-xs leading-relaxed">
                <p>
                  (주)명신기공은 협력사와의 거래 신뢰도를 높이기 위해 <strong>전자 계약서 상호 직인/도장 날인 체결</strong>을 지원합니다.
                </p>

                {/* 3-Step Contract Flow */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-center">
                  <div className="p-4 rounded-2xl bg-indigo-50 border border-indigo-200 space-y-2">
                    <span className="text-[10px] font-black text-indigo-600 bg-white px-2 py-0.5 rounded-full border border-indigo-100">
                      단계 1. 당사 날인
                    </span>
                    <div className="w-8 h-8 rounded-full bg-indigo-600 text-white flex items-center justify-center mx-auto text-xs font-black">
                      <Stamp className="h-4 w-4" />
                    </div>
                    <h5 className="font-bold text-indigo-950 text-xs">명신기공 직인 계약서 등록</h5>
                    <p className="text-[10.5px] text-slate-500">
                      명신기공 대표직인이 날인된 계약서 스캔본이 시스템에 등록됩니다.
                    </p>
                  </div>

                  <div className="p-4 rounded-2xl bg-rose-50 border-2 border-rose-300 space-y-2 relative">
                    <span className="text-[10px] font-black text-rose-700 bg-white px-2 py-0.5 rounded-full border border-rose-200 animate-pulse">
                      단계 2. 거래처 회신 (알림)
                    </span>
                    <div className="w-8 h-8 rounded-full bg-rose-600 text-white flex items-center justify-center mx-auto text-xs font-black">
                      <FileSignature className="h-4 w-4" />
                    </div>
                    <h5 className="font-bold text-rose-950 text-xs">거래처 도장 날인 후 업로드</h5>
                    <p className="text-[10.5px] text-rose-800 font-medium">
                      계약서를 다운로드하여 귀사 인감 날인 후 스캔하여 업로드합니다.
                    </p>
                  </div>

                  <div className="p-4 rounded-2xl bg-emerald-50 border border-emerald-200 space-y-2">
                    <span className="text-[10px] font-black text-emerald-700 bg-white px-2 py-0.5 rounded-full border border-emerald-100">
                      단계 3. 체결 완료
                    </span>
                    <div className="w-8 h-8 rounded-full bg-emerald-600 text-white flex items-center justify-center mx-auto text-xs font-black">
                      <CheckCircle2 className="h-4 w-4" />
                    </div>
                    <h5 className="font-bold text-emerald-950 text-xs">상호 날인 체결 확정</h5>
                    <p className="text-[10.5px] text-emerald-800 font-medium">
                      회신본이 등록되는 즉시 알림이 해제되고 정식 계약 체결로 완료됩니다.
                    </p>
                  </div>
                </div>

                <div className="p-3 bg-amber-50 rounded-xl border border-amber-200 text-amber-900 text-xs flex items-center gap-2">
                  <AlertTriangle className="h-4 w-4 text-amber-600 shrink-0" />
                  <span>
                    <strong>알림 해제 안내:</strong> 명신기공 날인 계약서가 등록된 후 거래처에서 회신본을 올릴 때까지 시스템에 
                    <strong className="text-rose-700"> '계약서 미회신 알림'</strong>이 지속 발생하오니 신속한 회신을 부탁드립니다.
                  </span>
                </div>
              </div>
            </section>
          )}

          {/* ══════════════════════════════════════════════════════════════════════ */}
          {/* FAQ & CONTACT SUPPORT */}
          {/* ══════════════════════════════════════════════════════════════════════ */}
          <section className="p-6 rounded-3xl bg-slate-50 border border-slate-200 space-y-4">
            <div className="flex items-center gap-2 border-b border-slate-200 pb-2">
              <HelpCircle className="h-5 w-5 text-indigo-600" />
              <h3 className="text-base font-black text-slate-900">자주 묻는 질문 (FAQ)</h3>
            </div>

            <div className="space-y-2 text-xs">
              {[
                {
                  q: '비밀번호를 분실했을 때는 어떻게 해야 하나요?',
                  a: '(주)명신기공 원자재관리팀으로 유선 연락(TEL: 02-421-1392) 주시면 본인 확인 후 초기 비밀번호(1234)로 즉시 재설정해 드립니다. 재설정 후에는 반드시 다시 비밀번호를 변경해 주십시오.'
                },
                {
                  q: '단가를 수정했는데 화면에 "변경요청(승인대기)"라고 뜨고 바로 적용되지 않아요.',
                  a: '협력사 포탈에서 단가나 네고율을 수정하시면 회계 및 품질 보증을 위해 당사 구매담당자의 승인 절차를 거치게 됩니다. 관리자 확인 후 즉시 최종 확정 반영됩니다.'
                },
                {
                  q: '엑셀 파일 업로드 시 파일 오류가 발생합니다.',
                  a: '본 매뉴얼 상단의 [표준 엑셀 양식 다운로드] 버튼을 눌러 양식을 내려받으신 후, 1행의 열 명칭(품번, 품명, 규격, 단위, 제조사, 협가, 네고율, 네고방식 등)을 변경하지 마시고 데이터를 기입하여 업로드해 주십시오.'
                },
                {
                  q: '단가표 파일(PDF 또는 이미지)도 등록해야 하나요?',
                  a: '자체 제작하신 카탈로그나 공식 단가표(PDF/Excel/이미지)가 있으신 경우 [업체정보 수정]의 단가표 파일 첨부란에 함께 올려주시면 구매팀에서 상세 스펙 확인 시 큰 도움이 됩니다.'
                }
              ].map((faq, idx) => (
                <div 
                  key={`manual-faq-${idx}`}
                  className="bg-white rounded-xl border border-slate-200 overflow-hidden transition-all"
                >
                  <button
                    onClick={() => setOpenFaq(openFaq === idx ? null : idx)}
                    className="w-full px-4 py-3 text-left font-bold text-slate-800 flex items-center justify-between hover:bg-slate-50 transition-colors"
                  >
                    <span className="flex items-center gap-2">
                      <span className="text-indigo-600 font-black">Q.</span>
                      <span>{faq.q}</span>
                    </span>
                    {openFaq === idx ? <ChevronUp className="h-4 w-4 text-slate-400" /> : <ChevronDown className="h-4 w-4 text-slate-400" />}
                  </button>
                  {openFaq === idx && (
                    <div className="px-4 pb-3 pt-1 text-slate-600 border-t border-slate-100 bg-slate-50/50 leading-relaxed">
                      <span className="text-emerald-600 font-black mr-1.5">A.</span>
                      {faq.a}
                    </div>
                  )}
                </div>
              ))}
            </div>

            {/* Official Support Footer Card */}
            <div className="pt-4 border-t border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-slate-500">
              <div>
                <strong className="text-slate-800">(주)명신기공 원자재관리 및 구매사업부</strong>
                <div className="text-[11px] text-slate-400 mt-0.5">시스템 문의, 비밀번호 초기화 및 품목 규격 상담</div>
              </div>
              <div className="flex items-center gap-3 font-mono font-bold text-slate-700">
                <span>TEL: 02-421-1392</span>
                <span>|</span>
                <span>FAX: 02-422-4171</span>
              </div>
            </div>
          </section>

          {/* Bottom Close Button (Screen only) */}
          <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-3 print:hidden">
            <div className="text-xs text-slate-400">
              * 본 설명서는 상단 [인쇄 / PDF 저장]을 통해 종이 문서 또는 PDF 파일로 간편 배포 가능합니다.
            </div>
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <button
                type="button"
                onClick={handleCopyDistributionText}
                className="flex-1 sm:flex-initial px-4 py-2.5 rounded-xl border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-xs transition-colors flex items-center justify-center gap-1.5"
              >
                <Copy className="h-3.5 w-3.5" />
                <span>문자/카톡 발송문 복사</span>
              </button>
              <button
                type="button"
                onClick={onClose}
                className="flex-1 sm:flex-initial px-6 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-extrabold text-xs transition-colors shadow-sm"
              >
                확인 및 닫기
              </button>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}
