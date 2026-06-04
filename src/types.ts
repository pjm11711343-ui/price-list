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
