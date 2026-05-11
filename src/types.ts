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
  createdAt?: any;
  updatedAt?: any;
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
  category?: string; // 카테고리 (밸브류, 피팅류 등)
  spec?: string;
  unit?: string;
  costPrice: number; // 현단가 (협가)
  negoRate: number;  // 네고율 (%)
  unitPrice: number; // 구매단가 (최종 계산가)
  remarks?: string;
  maker?: string;
  order?: number;
  useRounding?: boolean;
}
