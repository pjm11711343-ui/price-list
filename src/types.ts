export interface Vendor {
  id: string;
  name: string;
  phone?: string;
  fax?: string;
  representative?: string;
  businessNumber?: string;
  email: string;
  password?: string;
  priceTableUrl?: string;
  priceTableFileType?: 'image' | 'pdf' | 'excel' | 'unknown';
  priceTableFileName?: string;
  notes?: string;
  useRounding?: boolean;
  roundingMethod?: 'none' | 'round' | 'floor';
  masterCustomFlag?: boolean;
  handlingDetails?: Record<string, string>;
  categories?: string[];
  order?: number;
  deleted?: boolean;
  deletedAt?: any;
  createdAt?: any;
  updatedAt?: any;
  notice?: string;
  noticeFileUrl?: string;
  noticeFileType?: 'image' | 'pdf' | 'unknown';
  noticeFileName?: string;
  noticeUpdatedAt?: any;
  // 필수 증빙 서류 (사업자등록증, 통장사본, 전자어음약정확인서)
  businessCertUrl?: string;
  businessCertFileName?: string;
  businessCertFileType?: 'image' | 'pdf' | 'unknown';
  businessCertUpdatedAt?: any;
  bankbookUrl?: string;
  bankbookFileName?: string;
  bankbookFileType?: 'image' | 'pdf' | 'unknown';
  bankbookUpdatedAt?: any;
  promissoryNoteUrl?: string;
  promissoryNoteFileName?: string;
  promissoryNoteFileType?: 'image' | 'pdf' | 'unknown';
  promissoryNoteUpdatedAt?: any;
  // 계약서 전자 체결 및 관리 (명신기공 날인 -> 거래처 날인 회신)
  contractAdminUrl?: string; // (주)명신기공 날인 계약서
  contractAdminFileName?: string;
  contractAdminFileType?: 'image' | 'pdf' | 'unknown';
  contractAdminUpdatedAt?: any;
  contractVendorUrl?: string; // 거래처 직인/인감 날인 계약서
  contractVendorFileName?: string;
  contractVendorFileType?: 'image' | 'pdf' | 'unknown';
  contractVendorUpdatedAt?: any;
  contractStatus?: 'none' | 'pending_vendor' | 'completed'; // 미등록 / 거래처 날인 회신 대기(알림 발생) / 체결 완료
  contractNote?: string; // 계약 비고사항 / 계약기간 등
}

export interface ComparisonRow {
  id: string;
  category: string;
  vendor1: string;
  vendor2: string;
  vendor3: string;
  vendor4: string;
  order: number;
}

export interface PriceItem {
  id: string;
  vendorId: string;
  itemCode?: string; // 품번
  itemName: string;  // 품명
  category?: string; // 카테고리 (밸브류, 피팅류, STS파이프 등)
  spec?: string;
  unit?: string;
  costPrice: number; // 현단가 (협가 또는 KG단가)
  negoRate: number;  // 네고율 (%)
  negoType?: 'percent' | 'sts_pipe'; // 네고 방식 (STS PIPE 추가)
  weight?: number; // 단중 (kg/m 등)
  unitPrice: number; // 구매단가 (최종 계산가)
  baseUnitPrice?: number; // 기준 구매단가 (변동률 계산용)
  remarks?: string;
  maker?: string;
  order?: number;
  useRounding?: boolean;
  hasPendingUpdate?: boolean; // 승인 대기 중 여부
  priceHistory?: { price: number; date: any }[]; // 최근 3개 가격 변동 이력
  targetPrice?: number; // 목표가 / 알림 임계값
  updatedAt?: any;
}

export interface PendingPriceUpdate {
  id: string;
  vendorId: string;
  priceItemId: string;
  itemName: string;
  oldData: Partial<PriceItem>;
  newData: Partial<PriceItem>;
  status: 'pending' | 'approved' | 'rejected';
  requestedBy: string; // 'vendor' or name
  requestedAt: any;
  approvedAt?: any;
}

export type FluctuationPeriodType = '1m' | '3m' | '6m' | '1y' | 'all' | 'custom';
export type FluctuationSortMode = 'absAmount' | 'rate' | 'increase' | 'decrease';

export interface PriceFluctuationItem {
  key: string; // unique item key (${itemName}_${spec})
  itemName: string;
  spec: string;
  unit: string;
  category: string;
  primaryVendorId: string;
  primaryVendorName: string;
  allVendorNames: string[];
  startPrice: number;
  endPrice: number;
  priceDiff: number; // endPrice - startPrice
  absDiff: number; // Math.abs(priceDiff)
  percentDiff: number; // ((endPrice - startPrice) / startPrice) * 100
  direction: 'up' | 'down' | 'same';
  minPrice: number;
  maxPrice: number;
  lastUpdated?: any;
  changeCount: number;
  historyTimeline: { price: number; date: any; label?: string }[];
}
