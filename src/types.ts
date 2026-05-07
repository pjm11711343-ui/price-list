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
  itemName: string;
  spec?: string;
  unit?: string;
  costPrice: number; // 협가
  negoRate: number;  // 네고율 (%)
  unitPrice: number; // 단가 (최종 계산가)
  remarks?: string;
  maker?: string;
  order?: number;
}
