import React, { useState, useEffect, useMemo } from 'react';
import { 
  Plus, 
  Search, 
  Lock, 
  Unlock, 
  ChevronRight, 
  ArrowLeft, 
  Building2, 
  Phone, 
  Mail, 
  User, 
  LayoutGrid, 
  List,
  Edit2,
  Trash2,
  X,
  FileText,
  Save,
  Loader2,
  Upload,
  Download,
  Table,
  ChevronUp,
  ChevronDown,
  Link as LinkIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { QRCodeCanvas } from 'qrcode.react';
import { 
  collection, 
  query, 
  onSnapshot, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc, 
  orderBy, 
  serverTimestamp,
  getDocs,
  getDocFromServer
} from 'firebase/firestore';
import { 
  ref as storageRef, 
  uploadBytes, 
  getDownloadURL,
  deleteObject
} from 'firebase/storage';
import { db, auth, storage } from './lib/firebase';
import { Vendor, PriceItem, ComparisonRow } from './types';
import { ALL_VENDORS } from './data/seedData';

// Error Handling Types
enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
    },
    operationType,
    path
  }
  console.error('Firestore Error Detailed: ', JSON.stringify(errInfo));
  return errInfo;
}

// Components
const Modal = ({ children, onClose, title }: { children: React.ReactNode; onClose: () => void; title: string; key?: string }) => (
  <motion.div 
    initial={{ opacity: 0 }} 
    animate={{ opacity: 1 }} 
    exit={{ opacity: 0 }}
    className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 backdrop-blur-sm"
  >
    <motion.div 
      initial={{ scale: 0.95, opacity: 0 }}
      animate={{ scale: 1, opacity: 1 }}
      exit={{ scale: 0.95, opacity: 0 }}
      className="w-full max-w-2xl overflow-hidden rounded-2xl bg-white shadow-2xl"
    >
      <div className="flex items-center justify-between border-b border-slate-100 bg-slate-50 px-6 py-4">
        <h3 className="text-xl font-bold text-slate-800">{title}</h3>
        <button onClick={onClose} className="rounded-full p-1 hover:bg-slate-200">
          <X className="h-6 w-6 text-slate-500" />
        </button>
      </div>
      <div className="p-6 overflow-y-auto max-h-[80vh]">{children}</div>
    </motion.div>
  </motion.div>
);

export default function App() {
  const [vendors, setVendors] = useState<Vendor[]>([]);
  const [selectedVendor, setSelectedVendor] = useState<Vendor | null>(null);
  const [isAddingVendor, setIsAddingVendor] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [isVerified, setIsVerified] = useState<string | null>(null); // Tracks which vendor ID is verified
  const [loading, setLoading] = useState(true);
  const [viewMode, setViewMode] = useState<'detail' | 'matrix'>('detail');

  // Price table state
  const [priceItems, setPriceItems] = useState<PriceItem[]>([]);
  const [isAddingItem, setIsAddingItem] = useState(false);
  const [isEditingVendorInfo, setIsEditingVendorInfo] = useState(false);
  const [isChangingPassword, setIsChangingPassword] = useState(false);
  const [errorInfo, setErrorInfo] = useState<FirestoreErrorInfo | null>(null);
  const [isSeedModalOpen, setIsSeedModalOpen] = useState(false);
  const [bulkFile, setBulkFile] = useState<File | null>(null);
  const [bulkItemFile, setBulkItemFile] = useState<File | null>(null);
  const [isBulkItemUploadOpen, setIsBulkItemUploadOpen] = useState(false);
  const [editingPriceId, setEditingPriceId] = useState<string | null>(null);
  const [priceTableFile, setPriceTableFile] = useState<File | null>(null);
  const [excelPreviewData, setExcelPreviewData] = useState<any[] | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [seedStatus, setSeedStatus] = useState<{ current: number; total: number; isDone: boolean } | null>(null);
  const [isAdminMode, setIsAdminMode] = useState(false);
  const [isDeepLinkMode, setIsDeepLinkMode] = useState(false);
  const [adminPasswordInput, setAdminPasswordInput] = useState('');
  const [showAdminLogin, setShowAdminLogin] = useState(false);
  const [copySuccess, setCopySuccess] = useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const [adminNotification, setAdminNotification] = useState(false);
  const [isNotesExpanded, setIsNotesExpanded] = useState(true);
  const [isMatrixExpanded, setIsMatrixExpanded] = useState(false);
  const [comparisonRows, setComparisonRows] = useState<ComparisonRow[]>([]);
  const [isAddingComparisonRow, setIsAddingComparisonRow] = useState(false);
  const [usdToKrw, setUsdToKrw] = useState<number>(1350);
  const [lastRateUpdate, setLastRateUpdate] = useState<string>("");
  const [savingMatrixId, setSavingMatrixId] = useState<string | null>(null);

  const canManageItems = isAdminMode || (selectedVendor && isVerified === selectedVendor.id);

  const calculatePrice = (cost: number, rate: number) => {
    let price = cost * (1 - (rate / 100));
    if (selectedVendor?.useRounding) {
      // 5원 이하 절사, 5원 이상 반올림 (사실상 5원 미만 절사, 5원 이상 반올림이 표준이나
      // 요청대로 5원 경계에서의 처리를 구현. 보통 5원 이상이면 올리는 게 표준 반올림임)
      const lastDigit = price % 10;
      if (lastDigit >= 5) {
        price = Math.ceil(price / 10) * 10;
      } else {
        price = Math.floor(price / 10) * 10;
      }
    }
    return price;
  };

  // Fetch Exchange Rate
  useEffect(() => {
    const fetchRate = async () => {
      try {
        const response = await fetch('https://open.er-api.com/v6/latest/USD');
        const data = await response.json();
        if (data.rates && data.rates.KRW) {
          setUsdToKrw(data.rates.KRW);
          setLastRateUpdate(new Date().toLocaleTimeString());
        }
      } catch (error) {
        console.error("Failed to fetch exchange rate:", error);
      }
    };
    fetchRate();
    // Refresh every 30 minutes
    const interval = setInterval(fetchRate, 30 * 60 * 1000);
    return () => clearInterval(interval);
  }, []);

  // Sync selectedVendor when vendors list updates
  useEffect(() => {
    if (vendors.length === 0) return;

    // Check deep link first
    const params = new URLSearchParams(window.location.search);
    const urlVendorId = params.get('v');

    if (urlVendorId) {
      const vendor = vendors.find(v => v.id === urlVendorId);
      if (vendor) {
        setSelectedVendor(vendor);
        setIsDeepLinkMode(true);
      }
    } else if (selectedVendor) {
      // Sync manual selection with latest data from Firestore
      const updated = vendors.find(v => v.id === selectedVendor.id);
      if (!updated) {
        // If vendor was deleted from the list, deselect it
        setSelectedVendor(null);
        setIsVerified(null);
        setViewMode('detail');
      } else if (JSON.stringify(updated) !== JSON.stringify(selectedVendor)) {
        setSelectedVendor(updated);
      }
    }
  }, [vendors, selectedVendor?.id]);

  const copyVendorLink = (vendorId: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    const url = new URL(window.location.origin);
    url.searchParams.set('v', vendorId);
    navigator.clipboard.writeText(url.toString());
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  const handleAdminLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (adminPasswordInput === 'admin1234') { // Default admin password
      setIsAdminMode(true);
      setShowAdminLogin(false);
      setAdminPasswordInput('');
      setAdminNotification(true);
      setTimeout(() => setAdminNotification(false), 5000); // Auto-dismiss after 5s
    } else {
      alert('관리자 비밀번호가 틀렸습니다.');
    }
  };

  const deleteVendor = async (id: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    
    if (!isAdminMode) {
      alert('관리자 모드에서만 삭제가 가능합니다.');
      return;
    }
    
    const vendorToDelete = vendors.find(v => v.id === id);
    const vendorName = vendorToDelete ? vendorToDelete.name : "이 업체";

    if (!window.confirm(`[🚨 주의] '${vendorName}' 업체를 정말로 삭제하시겠습니까?\n이 작업은 되돌릴 수 없으며 모든 단가 정보가 함께 영구 삭제됩니다.`)) {
      return;
    }

    try {
      setLoading(true);
      console.log(`Starting deletion for vendor: ${id}`);
      
      // 1. Delete associated prices
      const pricesRef = collection(db, 'vendors', id, 'prices');
      const pricesSnap = await getDocs(pricesRef);
      console.log(`Found ${pricesSnap.size} prices to delete`);
      
      for (const priceDoc of pricesSnap.docs) {
        await deleteDoc(doc(db, 'vendors', id, 'prices', priceDoc.id));
      }
      console.log("Sub-collection 'prices' deleted.");

      // 2. Delete the vendor document
      await deleteDoc(doc(db, 'vendors', id));
      console.log("Vendor document deleted.");
      
      // Clear state if we deleted the currently viewed vendor
      if (selectedVendor?.id === id) {
        setSelectedVendor(null);
        setIsVerified(null);
        setViewMode('detail');
      }
      
      alert('업체가 성공적으로 삭제되었습니다.');
    } catch (error) {
      console.error("Delete Error details:", error);
      let message = '업체 삭제 중 오류가 발생했습니다.';
      if (error instanceof Error) {
        message += `\n내용: ${error.message}`;
        if (error.message.includes('permission')) {
          message += '\n(권한 부족: Firestore 규칙을 확인하세요)';
        }
      }
      alert(message);
    } finally {
      setLoading(false);
    }
  };

  // Connection Test
  useEffect(() => {
    const testConnection = async () => {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
        console.log("Firebase Connection Active");
      } catch (error) {
        if (error instanceof Error && error.message.includes('offline')) {
          console.error("Firebase is offline. Check configuration.");
        }
      }
    };
    testConnection();
  }, []);

  const handleBulkUpload = async () => {
    if (!bulkFile) return;
    setIsSeedModalOpen(false);
    setLoading(true);
    
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        setSeedStatus({ current: 0, total: data.length, isDone: false });
        
        let count = 0;
        for (const row of data) {
          const vendorData = {
            name: row['업체명'] || row['Name'] || '',
            representative: row['대표자'] || row['Representative'] || '',
            businessNumber: row['사업자번호'] || row['Business Number'] || '',
            phone: row['전화번호'] || row['Phone'] || '',
            fax: row['팩스번호'] || row['Fax'] || '',
            email: row['이메일'] || row['Email'] || '',
            password: String(row['비밀번호'] || row['Password'] || '1234'),
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp()
          };

          if (vendorData.name) {
            await addDoc(collection(db, "vendors"), vendorData);
            count++;
            setSeedStatus(prev => prev ? { ...prev, current: count } : null);
          }
        }
        setSeedStatus(prev => prev ? { ...prev, isDone: true } : null);
        setBulkFile(null);
      } catch (error) {
        handleFirestoreError(error, OperationType.WRITE, "vendors");
        setSeedStatus(null);
        alert("엑셀 파싱 중 오류가 발생했습니다. 규격을 확인해주세요.");
      } finally {
        setLoading(false);
      }
    };
    reader.readAsBinaryString(bulkFile);
  };

  const downloadPriceTemplate = () => {
    const templateData = [
      {
        '카테고리': '배관재',
        '하위 카테고리': '강관',
        '품명': '무계목 강관',
        '규격': '100A SCH40',
        '단위': 'M',
        '단가': 15000,
        '비고': '신규 모델'
      },
      {
        '카테고리': '배관재',
        '하위 카테고리': '밸브',
        '품명': '게이트 밸브',
        '규격': '50A',
        '단위': 'EA',
        '단가': 25000,
        '비고': '내식성 강화'
      }
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "단가표_양식");
    XLSX.writeFile(wb, "ERP_단가표_업로드_양식.xlsx");
  };

  const handleBulkItemUpload = async () => {
    if (!bulkItemFile || !selectedVendor) return;
    
    setIsBulkItemUploadOpen(false);
    setLoading(true);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        let count = 0;
        for (const row of data) {
          const costPrice = Number(row['협가'] || row['단가'] || row['Price'] || 0);
          const negoRate = Number(row['네고율'] || 0);
          const unitPrice = calculatePrice(costPrice, negoRate);

          const itemData = {
            vendorId: selectedVendor.id,
            itemName: String(row['품목명'] || row['품명'] || row['Name'] || ''),
            spec: String(row['규격'] || ''),
            unit: String(row['단위'] || 'EA'),
            costPrice,
            negoRate,
            unitPrice,
            remarks: String(row['비고'] || ''),
            maker: String(row['메이커'] || row['제조사'] || ''),
            isConfirmed: false,
            order: count,
            updatedAt: serverTimestamp()
          };

          if (itemData.itemName) {
            await addDoc(collection(db, "vendors", selectedVendor.id, "prices"), itemData);
            count++;
          }
        }
        alert(`${count}개의 품목이 성공적으로 등록되었습니다.`);
        setBulkItemFile(null);
      } catch (error) {
        console.error("Bulk item upload error:", error);
        alert("업로드 중 오류가 발생했습니다. 양식을 확인해주세요.");
      } finally {
        setLoading(false);
      }
    };
    reader.readAsBinaryString(bulkItemFile);
  };

  // Fetch Comparison Rows
  useEffect(() => {
    const q = query(collection(db, 'comparison_matrix'), orderBy('order', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as ComparisonRow));
      setComparisonRows(data);
    });
    return unsubscribe;
  }, []);

  // Fetch Vendors
  useEffect(() => {
    const q = query(collection(db, 'vendors'), orderBy('name', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Vendor));
      setVendors(data);
      setLoading(false);
      setErrorInfo(null);
    }, (error) => {
      const info = handleFirestoreError(error, OperationType.LIST, "vendors");
      setErrorInfo(info);
      setLoading(false);
    });
    return unsubscribe;
  }, []);

  // Fetch Prices for selected vendor
  useEffect(() => {
    if (selectedVendor) {
      const q = query(
        collection(db, 'vendors', selectedVendor.id, 'prices'), 
        orderBy('order', 'asc')
      );
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as PriceItem));
        setPriceItems(data);
      });
      return unsubscribe;
    } else {
      setPriceItems([]);
    }
  }, [selectedVendor]);
  
  // Excel preview fetch
  useEffect(() => {
    if (selectedVendor?.priceTableUrl && selectedVendor?.priceTableFileType === 'excel' && isVerified === selectedVendor.id) {
      fetch(selectedVendor.priceTableUrl)
        .then(res => res.arrayBuffer())
        .then(buffer => {
          const workbook = XLSX.read(buffer, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          setExcelPreviewData(data as any[]);
        })
        .catch(err => {
          console.error("Excel preview fetch error:", err);
          setExcelPreviewData(null);
        });
    } else {
      setExcelPreviewData(null);
    }
  }, [selectedVendor?.priceTableUrl, selectedVendor?.priceTableFileType, isVerified, selectedVendor?.id]);

  const baseSortedVendors = useMemo(() => {
    const stripPrefix = (name: string) => {
      return name.replace(/^(\(주\)|주식회사)\s?/, '');
    };

    return [...vendors].sort((a, b) => {
      const nameA = stripPrefix(a.name);
      const nameB = stripPrefix(b.name);
      return nameA.localeCompare(nameB, 'ko');
    });
  }, [vendors]);

  const sortedVendors = useMemo(() => {
    return baseSortedVendors
      .filter(v => 
        v.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        v.representative?.toLowerCase().includes(searchTerm.toLowerCase())
      );
  }, [baseSortedVendors, searchTerm]);

  const handleAddVendor = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const password = formData.get('password') as string;
    
    setIsUploading(true);
    let priceTableUrl = '';
    let priceTableFileType: Vendor['priceTableFileType'] = 'unknown';
    let priceTableFileName = '';

    if (priceTableFile) {
      try {
        // File size limit check (e.g., 10MB)
        if (priceTableFile.size > 10 * 1024 * 1024) {
          throw new Error("파일 크기가 너무 큽니다. (최대 10MB)");
        }

        const fileExt = priceTableFile.name.split('.').pop()?.toLowerCase();
        const fileName = `${Date.now()}_${priceTableFile.name}`;
        const sRef = storageRef(storage, `price_tables/${fileName}`);
        const snapshot = await uploadBytes(sRef, priceTableFile);
        priceTableUrl = await getDownloadURL(snapshot.ref);
        priceTableFileName = priceTableFile.name;
        
        if (['jpg', 'jpeg', 'png', 'gif', 'webp'].includes(fileExt || '')) {
          priceTableFileType = 'image';
        } else if (fileExt === 'pdf') {
          priceTableFileType = 'pdf';
        } else if (['xlsx', 'xls', 'csv'].includes(fileExt || '')) {
          priceTableFileType = 'excel';
        }
      } catch (error: any) {
        console.error("Error uploading price table (Add):", error);
        setIsUploading(false);
        let msg = `단가표 파일 업로드에 실패했습니다: ${error.message || "서버 오류"}`;
        if (error.code === 'storage/retry-limit-exceeded' || error.code === 'storage/unauthorized' || error.message?.includes('retry-limit-exceeded')) {
          msg = "Firebase Storage 업로드에 실패했습니다.\n\n해결 방법:\n1. Firebase 콘솔의 'Storage' 메뉴에서 '시작하기'를 눌러 서비스를 활성화했는지 확인해 주세요.\n2. Storage 보안 규칙이 업로드를 허용하는지 확인해 주세요. (예: allow read, write: if true;)\n3. 네트워크 상태를 확인하고 다시 시도해 주세요.";
        }
        alert(msg);
        return; // Stop execution if upload fails
      }
    }

    const newVendor = {
      name: formData.get('name') as string,
      phone: formData.get('phone') as string,
      fax: formData.get('fax') as string,
      representative: formData.get('representative') as string,
      businessNumber: formData.get('businessNumber') as string,
      email: formData.get('email') as string,
      password: password,
      notes: formData.get('notes') as string,
      useRounding: formData.get('useRounding') === 'on',
      masterCustomFlag: formData.get('masterCustomFlag') === 'on',
      priceTableUrl,
      priceTableFileType,
      priceTableFileName,
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp(),
    };

    try {
      await addDoc(collection(db, 'vendors'), newVendor);
      setIsAddingVendor(false);
      setPriceTableFile(null);
    } catch (error) {
      console.error("Error adding vendor:", error);
      alert("업체 추가에 실패했습니다.");
    } finally {
      setIsUploading(false);
    }
  };

  const handleUpdateVendorInfo = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedVendor) return;

    const formData = new FormData(e.currentTarget);
    setIsUploading(true);

    let priceTableUrl = selectedVendor.priceTableUrl || '';
    let priceTableFileType = selectedVendor.priceTableFileType || 'unknown';
    let priceTableFileName = selectedVendor.priceTableFileName || '';

    if (priceTableFile) {
      try {
        // File size limit check (e.g., 10MB)
        if (priceTableFile.size > 10 * 1024 * 1024) {
          throw new Error("파일 크기가 너무 큽니다. (최대 10MB)");
        }

        const fileExt = priceTableFile.name.split('.').pop()?.toLowerCase();
        const fileName = `${Date.now()}_${priceTableFile.name}`;
        const sRef = storageRef(storage, `price_tables/${fileName}`);
        const snapshot = await uploadBytes(sRef, priceTableFile);
        priceTableUrl = await getDownloadURL(snapshot.ref);
        priceTableFileName = priceTableFile.name;
        
        if (['jpg', 'jpeg', 'png', 'gif', 'webp'].includes(fileExt || '')) {
          priceTableFileType = 'image';
        } else if (fileExt === 'pdf') {
          priceTableFileType = 'pdf';
        } else if (['xlsx', 'xls', 'csv'].includes(fileExt || '')) {
          priceTableFileType = 'excel';
        }
      } catch (error: any) {
        console.error("Error uploading price table (Update):", error);
        setIsUploading(false);
        let msg = `단가표 파일 업로드에 실패했습니다: ${error.message || "서버 오류"}`;
        if (error.code === 'storage/retry-limit-exceeded' || error.code === 'storage/unauthorized' || error.message?.includes('retry-limit-exceeded')) {
          msg = "Firebase Storage 업로드에 실패했습니다.\n\n해결 방법:\n1. Firebase 콘솔의 'Storage' 메뉴에서 '시작하기'를 눌러 서비스를 활성화했는지 확인해 주세요.\n2. Storage 보안 규칙이 업로드를 허용하는지 확인해 주세요. (예: allow read, write: if true;)\n3. 네트워크 상태를 확인하고 다시 시도해 주세요.";
        }
        alert(msg);
        return; // Stop execution if upload fails
      }
    }

    const updatedData = {
      name: formData.get('name') as string,
      representative: formData.get('representative') as string,
      phone: formData.get('phone') as string,
      fax: formData.get('fax') as string,
      businessNumber: formData.get('businessNumber') as string,
      email: formData.get('email') as string,
      notes: formData.get('notes') as string,
      useRounding: formData.get('useRounding') === 'on',
      masterCustomFlag: formData.get('masterCustomFlag') === 'on',
      priceTableUrl,
      priceTableFileType,
      priceTableFileName,
      updatedAt: serverTimestamp()
    };

    try {
      await updateDoc(doc(db, 'vendors', selectedVendor.id), updatedData);
      setSelectedVendor(prev => prev ? { ...prev, ...updatedData } : null);
      setIsEditingVendorInfo(false);
      setPriceTableFile(null);
      alert('업체 정보가 수정되었습니다.');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `vendors/${selectedVendor.id}`);
      alert('정보 수정 중 오류가 발생했습니다.');
    } finally {
      setIsUploading(false);
    }
  };

  const handleUpdatePassword = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedVendor) return;

    const formData = new FormData(e.currentTarget);
    const newPassword = formData.get('newPassword') as string;

    if (!newPassword || newPassword.length < 4) {
      alert('비밀번호는 최소 4자 이상이어야 합니다.');
      return;
    }

    try {
      await updateDoc(doc(db, 'vendors', selectedVendor.id), {
        password: newPassword,
        updatedAt: serverTimestamp()
      });
      // Immediately update local state
      setSelectedVendor(prev => prev ? { ...prev, password: newPassword } : null);
      setIsChangingPassword(false);
      alert('비밀번호가 성공적으로 변경되었습니다.');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `vendors/${selectedVendor.id}`);
      alert('비밀번호 변경 중 오류가 발생했습니다.');
    }
  };


  const handleSelectVendor = (vendor: Vendor) => {
    setSelectedVendor(vendor);
    setPasswordInput('');
  };

  const verifyPassword = (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedVendor && passwordInput === selectedVendor.password) {
      setIsVerified(selectedVendor.id);
    } else {
      alert('비밀번호가 일치하지 않습니다.');
    }
  };

  const handleAddComparisonRow = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const newRow = {
      category: formData.get('category') as string,
      vendor1: formData.get('vendor1') as string,
      vendor2: formData.get('vendor2') as string,
      vendor3: formData.get('vendor3') as string,
      vendor4: formData.get('vendor4') as string,
      order: comparisonRows.length,
    };

    try {
      await addDoc(collection(db, 'comparison_matrix'), newRow);
      setIsAddingComparisonRow(false);
    } catch (error) {
      console.error("Error adding comparison row:", error);
    }
  };

  const updateComparisonRow = async (id: string, field: string, value: string) => {
    try {
      setSavingMatrixId(id);
      await updateDoc(doc(db, 'comparison_matrix', id), {
        [field]: value,
        updatedAt: serverTimestamp()
      });
      // Keeping the success indicator for a bit
      setTimeout(() => setSavingMatrixId(null), 1500);
    } catch (error) {
      console.error("Error updating comparison row:", error);
      setSavingMatrixId(null);
    }
  };

  // Debounced wrapper for real-time saving
  const debouncedUpdateRefs = React.useRef<{[key: string]: NodeJS.Timeout}>({});
  const handleMatrixChange = (id: string, field: string, value: string) => {
    const key = `${id}-${field}`;
    if (debouncedUpdateRefs.current[key]) {
      clearTimeout(debouncedUpdateRefs.current[key]);
    }
    
    debouncedUpdateRefs.current[key] = setTimeout(() => {
      updateComparisonRow(id, field, value);
      delete debouncedUpdateRefs.current[key];
    }, 800); // Save after 800ms of inactivity
  };

  const moveComparisonRow = async (rowIndex: number, direction: 'up' | 'down') => {
    const newIndex = direction === 'up' ? rowIndex - 1 : rowIndex + 1;
    if (newIndex < 0 || newIndex >= comparisonRows.length) return;

    const row1 = comparisonRows[rowIndex];
    const row2 = comparisonRows[newIndex];

    try {
      await updateDoc(doc(db, 'comparison_matrix', row1.id), { order: newIndex });
      await updateDoc(doc(db, 'comparison_matrix', row2.id), { order: rowIndex });
    } catch (error) {
      console.error("Error moving row:", error);
    }
  };

  const deleteComparisonRow = async (id: string) => {
    if (window.confirm('항목을 삭제하시겠습니까?')) {
      try {
        await deleteDoc(doc(db, 'comparison_matrix', id));
      } catch (error) {
        console.error("Error deleting comparison row:", error);
      }
    }
  };

  const handleAddItem = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedVendor) return;
    const formData = new FormData(e.currentTarget);
    const costPrice = Number(formData.get('costPrice') || 0);
    const negoRate = Number(formData.get('negoRate') || 0);
    const unitPrice = calculatePrice(costPrice, negoRate);

    const newItem = {
      vendorId: selectedVendor.id,
      itemName: formData.get('itemName') as string,
      spec: formData.get('spec') as string,
      unit: formData.get('unit') as string,
      costPrice,
      negoRate,
      unitPrice,
      remarks: formData.get('remarks') as string,
      maker: formData.get('maker') as string,
      isConfirmed: false,
      order: priceItems.length + 1,
    };

    try {
      await addDoc(collection(db, 'vendors', selectedVendor.id, 'prices'), newItem);
      setIsAddingItem(false);
    } catch (error) {
      console.error("Error adding price item:", error);
    }
  };

  const deletePriceItem = async (itemId: string) => {
    if (!selectedVendor) return;
    if (window.confirm('항목을 삭제하시겠습니까?')) {
      try {
        await deleteDoc(doc(db, 'vendors', selectedVendor.id, 'prices', itemId));
      } catch (error) {
        console.error("Error deleting price item:", error);
      }
    }
  };

  const handleUpdatePriceItem = async (itemId: string, updates: Partial<PriceItem>) => {
    if (!selectedVendor) return;
    
    // Recalculate unit price if cost or rate changed
    let finalUnitPrice = updates.unitPrice;
    if (updates.costPrice !== undefined || updates.negoRate !== undefined) {
      const item = priceItems.find(i => i.id === itemId);
      if (item) {
        const cost = updates.costPrice ?? item.costPrice;
        const rate = updates.negoRate ?? item.negoRate;
        finalUnitPrice = calculatePrice(cost, rate);
      }
    }

    try {
      await updateDoc(doc(db, 'vendors', selectedVendor.id, 'prices', itemId), {
        ...updates,
        ...(finalUnitPrice !== undefined ? { unitPrice: finalUnitPrice } : {}),
        updatedAt: serverTimestamp()
      });
      setEditingPriceId(null);
    } catch (error) {
      console.error("Error updating price item:", error);
      alert("수정 중 오류가 발생했습니다.");
    }
  };

  const togglePriceItemConfirmation = async (itemId: string, currentState: boolean) => {
    if (!selectedVendor) return;
    try {
      await updateDoc(doc(db, 'vendors', selectedVendor.id, 'prices', itemId), {
        isConfirmed: !currentState,
        updatedAt: serverTimestamp()
      });
    } catch (error) {
      console.error("Error toggling item confirmation:", error);
      handleFirestoreError(error, OperationType.UPDATE, `vendors/${selectedVendor.id}/prices/${itemId}`);
      alert("확인 상태 변경 중 오류가 발생했습니다.");
    }
  };

  const exportToExcel = () => {
    if (!selectedVendor || priceItems.length === 0) {
      alert('내보낼 데이터가 없습니다.');
      return;
    }

    const data = priceItems.map((item, index) => ({
      '번호': index + 1,
      '품목': item.itemName,
      '규격': item.spec || '',
      '단위': item.unit || '',
      '제조사': item.maker || '',
      '협가': item.costPrice || 0,
      '네고율(%)': item.negoRate || 0,
      '단가(최종)': item.unitPrice,
      '비고': item.remarks || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '단가표');
    
    XLSX.writeFile(workbook, `${selectedVendor.name}_단가표_${new Date().toLocaleDateString()}.xlsx`);
  };

  const importFromExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !selectedVendor) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        setLoading(true);
        let importCount = 0;

        for (const row of data) {
          const costPrice = Number(row['협가'] || row['Negotiated Price'] || 0);
          const negoRate = Number(row['네고율'] || row['Nego Rate'] || 0);
          const unitPrice = calculatePrice(costPrice, negoRate);

          const newItem = {
            vendorId: selectedVendor.id,
            itemName: row['품목명'] || row['품목'] || '',
            spec: row['규격'] || '',
            unit: row['단위'] || '',
            maker: row['메이커'] || row['제조사'] || '',
            costPrice,
            negoRate,
            unitPrice,
            remarks: row['비고'] || '',
            order: priceItems.length + importCount + 1,
          };

          if (newItem.itemName) {
            await addDoc(collection(db, 'vendors', selectedVendor.id, 'prices'), newItem);
            importCount++;
          }
        }

        alert(`${importCount}개의 항목을 성공적으로 불러왔습니다.`);
      } catch (error) {
        console.error("Error importing excel:", error);
        alert('엑셀 파일을 읽는 중 오류가 발생했습니다. 규격을 확인해주세요.');
      } finally {
        setLoading(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  if (loading) {
    return (
      <div className="flex h-screen w-full flex-col items-center justify-center bg-slate-900 gap-4">
        <Loader2 className="h-10 w-10 animate-spin text-indigo-500" />
        <p className="text-slate-500 font-medium animate-pulse">시스템 초기화 중...</p>
      </div>
    );
  }

  return (
    <div className="flex h-screen w-full overflow-hidden bg-slate-50 text-slate-900 font-sans relative">
      {/* Admin Notification Banner */}
      <AnimatePresence>
        {adminNotification && (
          <motion.div
            initial={{ y: -100, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: -100, opacity: 0 }}
            className="fixed top-6 left-1/2 -translate-x-1/2 z-[100] w-full max-w-sm"
          >
            <div className="mx-4 bg-slate-900 text-white rounded-2xl shadow-2xl shadow-indigo-500/20 border border-slate-800 p-4 flex items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <div className="bg-indigo-600 p-2 rounded-xl">
                  <Unlock className="h-4 w-4" />
                </div>
                <div>
                  <p className="text-sm font-bold">관리자 인증 완료</p>
                  <p className="text-[10px] text-slate-400 font-medium">관리자 모드로 전환되었습니다.</p>
                </div>
              </div>
              <button 
                onClick={() => setAdminNotification(false)}
                className="p-2 hover:bg-slate-800 rounded-lg transition-colors text-slate-500 hover:text-white"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Sidebar: Supplier Management */}
      {(!isDeepLinkMode || isAdminMode) && (
        <aside className="w-80 bg-slate-900 flex flex-col border-r border-slate-800 shrink-0">
        <div className="p-6 border-b border-slate-800">
          <div className="flex items-center justify-between mb-6">
            <h2 className="text-white font-bold text-lg">업체 리스트</h2>
            <div className="flex gap-2">
              {isAdminMode && isDeepLinkMode && (
                <button 
                  onClick={() => {
                    setIsDeepLinkMode(false);
                    const url = new URL(window.location.origin);
                    window.history.replaceState({}, '', url.toString());
                  }}
                  className="bg-slate-700 hover:bg-slate-600 text-white p-1.5 rounded transition-colors"
                  title="전체 리스트로 복구"
                >
                  <ArrowLeft className="h-4 w-4" />
                </button>
              )}
              <button 
                onClick={() => setIsAddingVendor(true)}
                className="bg-indigo-600 hover:bg-indigo-500 text-white p-1.5 rounded text-xs flex items-center gap-1 transition-colors"
              >
                <Plus className="h-4 w-4" />
                업체 추가
              </button>
            </div>
          </div>
          <div className="relative">
            <input 
              type="text" 
              placeholder="업체명 검색..." 
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full bg-slate-800 border border-slate-700 rounded px-3 py-2 text-sm text-slate-300 focus:outline-none focus:ring-1 focus:ring-indigo-500"
            />
          </div>
        </div>

        <nav className="flex-1 overflow-y-auto py-2">
          {errorInfo && (
            <div className="mx-6 mb-4 p-3 bg-red-50 border border-red-100 rounded-lg">
              <p className="text-[10px] text-red-600 font-bold uppercase mb-1">데이터 연결 오류</p>
              <p className="text-xs text-red-500 leading-tight">{errorInfo.error}</p>
              <button 
                onClick={() => window.location.reload()}
                className="mt-2 text-[10px] text-red-600 font-bold underline"
              >
                새로고침
              </button>
            </div>
          )}

          <div className="px-6 py-2 flex items-center justify-between">
            <span className="text-[10px] uppercase tracking-wider text-slate-500 font-bold">가나다 순 정렬</span>
            <div className="flex gap-2">
              <button 
                onClick={() => setViewMode('matrix')}
                className={`text-[10px] font-bold px-2 py-1 rounded transition-colors ${
                  viewMode === 'matrix' ? 'bg-indigo-600 text-white' : 'text-indigo-400 hover:text-indigo-300 underline'
                }`}
              >
                비교 매트릭스
              </button>
              <button 
                onClick={() => setIsSeedModalOpen(true)}
                className="text-[10px] font-bold text-indigo-400 hover:text-indigo-300 underline"
              >
                업무 일괄 등록
              </button>
            </div>
          </div>

          {!loading && vendors.length === 0 && !errorInfo && (
            <div className="mx-6 mt-4 p-4 border-2 border-dashed border-slate-200 rounded-xl bg-slate-50 text-center">
              <Building2 className="h-8 w-8 text-slate-300 mx-auto mb-2" />
              <p className="text-xs text-slate-500 mb-3">등록된 업체가 없습니다.</p>
              <button 
                onClick={() => setIsSeedModalOpen(true)}
                className="w-full py-2 bg-indigo-500 text-white text-xs font-bold rounded-lg hover:bg-indigo-600 transition-colors"
              >
                이미지 속 업체 일괄 등록하기
              </button>
            </div>
          )}

          <div className="mt-2">
            <AnimatePresence mode="popLayout">
              {sortedVendors.map((vendor) => (
                <motion.div
                  layout
                  key={vendor.id}
                  className={`flex w-full items-center justify-between px-6 py-3 transition-colors border-l-4 group ${
                    selectedVendor?.id === vendor.id 
                    ? 'bg-indigo-900/30 text-white border-indigo-500' 
                    : 'text-slate-400 hover:bg-slate-800 hover:text-white border-transparent'
                  }`}
                >
                  <button 
                    onClick={() => {
                      handleSelectVendor(vendor);
                      setViewMode('detail');
                    }}
                    className="flex-1 flex items-center truncate text-left"
                  >
                    <span className={`w-2 h-2 rounded-full mr-3 ${
                      selectedVendor?.id === vendor.id ? 'bg-indigo-500' : 'bg-slate-600'
                    }`}></span>
                    <span className="truncate">{vendor.name}</span>
                  </button>
                  <div className="flex items-center gap-1 shrink-0 relative z-10">
                    <button 
                      onClick={(e) => copyVendorLink(vendor.id, e)}
                      className="p-2 hover:bg-slate-700 active:bg-slate-600 rounded-lg transition-all text-slate-500 hover:text-indigo-400"
                      title="링크 복사"
                    >
                      <LinkIcon className="h-4 w-4" />
                    </button>
                    {isAdminMode && (
                      <button 
                        onClick={(e) => deleteVendor(vendor.id, e)}
                        className="p-2 hover:bg-red-500/20 active:bg-red-500/30 rounded-lg transition-all text-red-400 hover:text-red-300"
                        title="업체 삭제"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    )}
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleSelectVendor(vendor); }}
                      className="p-2 hover:bg-slate-700 active:bg-slate-600 rounded-lg transition-all"
                    >
                      {isVerified === vendor.id ? (
                        <Unlock className="h-4 w-4 text-indigo-400 opacity-80" />
                      ) : (
                        <Lock className="h-4 w-4 text-slate-500" />
                      )}
                    </button>
                  </div>
                </motion.div>
              ))}
            </AnimatePresence>
          </div>
        </nav>

        <div className="p-6 border-t border-slate-800 mt-auto bg-slate-900/50">
          <div className="flex items-center justify-between mb-2">
            <button 
              onClick={() => isAdminMode ? setIsAdminMode(false) : setShowAdminLogin(true)}
              className={`text-[10px] font-bold px-2 py-1 rounded transition-colors ${
                isAdminMode ? 'bg-amber-500/20 text-amber-500 hover:bg-amber-500/30' : 'text-slate-500 hover:text-white hover:bg-slate-800'
              }`}
            >
              {isAdminMode ? '관리자 모드 해제' : '관리자 로그인'}
            </button>
          </div>
          <p className="text-slate-600 text-[10px] text-center">© 2024 ERP Cost System</p>
        </div>
      </aside>
      )}

      {/* Main Content */}
      <main className="flex-1 flex flex-col relative overflow-hidden">
        {viewMode === 'matrix' ? (
          <div className="flex-1 flex flex-col h-full bg-white">
            <header className="px-8 py-6 border-b border-slate-200 shrink-0">
              <div className="flex items-center justify-between">
                <div>
                  <h1 className="text-2xl font-black text-slate-900 leading-tight">Vendor Handling Matrix</h1>
                  <p className="text-xs text-slate-500">업체별 취급 품목 및 단가 통합 비교</p>
                </div>
                <button 
                  onClick={() => setViewMode('detail')}
                  className="px-4 py-2 bg-slate-100 text-slate-600 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                >
                  상세 보기로 돌아가기
                </button>
              </div>
            </header>
            <div className="flex-1 overflow-auto p-8">
              <div className="inline-block min-w-full align-middle border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                <table className="min-w-full divide-y divide-slate-200">
                  <thead className="bg-slate-900">
                    <tr>
                      <th className="px-6 py-4 text-xs font-black text-white uppercase tracking-wider text-left border-r border-slate-800 w-[300px]">
                        구분 품명 (Category)
                      </th>
                      <th className="px-6 py-4 text-xs font-black text-slate-300 uppercase tracking-wider">업체명1</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-300 uppercase tracking-wider">업체명2</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-300 uppercase tracking-wider">업체명3</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-300 uppercase tracking-wider">업체명4</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-slate-200">
                    {comparisonRows.map((row) => (
                      <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                        <td className="px-6 py-4 text-sm font-bold text-slate-900 border-r border-slate-100 bg-slate-50/30">
                          {row.category}
                        </td>
                        <td className="px-6 py-4 text-sm text-slate-600 text-center">{row.vendor1 || '-'}</td>
                        <td className="px-6 py-4 text-sm text-slate-600 text-center">{row.vendor2 || '-'}</td>
                        <td className="px-6 py-4 text-sm text-slate-600 text-center">{row.vendor3 || '-'}</td>
                        <td className="px-6 py-4 text-sm text-slate-600 text-center">{row.vendor4 || '-'}</td>
                      </tr>
                    ))}
                    {comparisonRows.length === 0 && (
                      <tr>
                        <td colSpan={5} className="px-6 py-20 text-center text-slate-400">
                           비교 매트릭스 데이터가 없습니다. 업체 상세 페이지에서 구분을 추가해 주세요.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : selectedVendor ? (
          <>
            {/* Header Section */}
            <header className="bg-white border-b border-slate-200 px-8 py-6 shrink-0">
              <div className="flex items-start justify-between mb-6">
                <div className="flex items-center gap-3">
                  <div>
                    <h1 className="text-3xl font-bold text-slate-900 mb-1">{selectedVendor.name}</h1>
                    <p className="text-sm text-slate-500 flex items-center gap-2">
                      <Save className="h-3 w-3" />
                      최종 업데이트: {selectedVendor.updatedAt && typeof selectedVendor.updatedAt.toDate === 'function' 
                        ? selectedVendor.updatedAt.toDate().toLocaleString() 
                        : '방금 전'}
                    </p>
                  </div>
                    {(isVerified === selectedVendor.id || isAdminMode) && (
                      <div className="flex gap-2">
                        <button 
                          onClick={() => setIsEditingVendorInfo(true)}
                          className="p-2 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-indigo-600 transition-colors"
                          title="업체 정보 수정"
                        >
                          <LayoutGrid className="h-5 w-5" />
                        </button>
                        <button 
                          onClick={() => setIsChangingPassword(true)}
                          className="p-2 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-amber-600 transition-colors"
                          title="비밀번호 변경"
                        >
                          <Lock className="h-5 w-5" />
                        </button>
                        {isAdminMode && (
                          <button 
                            onClick={() => deleteVendor(selectedVendor.id)}
                            className="p-2 hover:bg-red-50 rounded-lg text-red-500 hover:text-red-700 transition-colors"
                            title="업체 삭제"
                          >
                            <Trash2 className="h-5 w-5" />
                          </button>
                        )}
                      </div>
                    )}
                </div>
                <div className="flex gap-3">
                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    onChange={importFromExcel} 
                    className="hidden" 
                    accept=".xlsx, .xls" 
                  />
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isVerified !== selectedVendor.id}
                    className="px-4 py-2 bg-white border border-slate-300 text-slate-700 rounded-lg text-sm font-medium hover:bg-slate-50 disabled:opacity-50 transition-colors flex items-center gap-2 shadow-sm"
                  >
                    <Upload className="h-4 w-4" />
                    불러오기
                  </button>
                  <button 
                    onClick={exportToExcel}
                    className="px-4 py-2 bg-white border border-slate-300 text-slate-700 rounded-lg text-sm font-medium hover:bg-slate-50 transition-colors flex items-center gap-2 shadow-sm"
                  >
                    <Download className="h-4 w-4" />
                    내보내기
                  </button>
                  {isAdminMode && (
                    <button 
                      onClick={() => setIsBulkItemUploadOpen(true)}
                      className="px-4 py-2 bg-slate-900 text-white rounded-lg text-sm font-bold hover:bg-slate-800 transition-all flex items-center gap-2 shadow-lg"
                    >
                      <Table className="h-4 w-4" />
                      단가표 양식 업로드
                    </button>
                  )}
                  <button 
                    onClick={() => setIsAddingItem(true)}
                    disabled={isVerified !== selectedVendor.id}
                    className="px-6 py-2 bg-indigo-600 text-white rounded-lg text-sm font-bold hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-md shadow-indigo-100"
                  >
                    단가 추가/수정
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 py-4 px-6 bg-slate-50 rounded-xl border border-slate-100">
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                    <User className="h-3 w-3" /> 대표자
                  </p>
                  <p className="text-sm font-semibold text-slate-700">{selectedVendor.representative || '-'}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                    <Phone className="h-3 w-3" /> 회사 전화
                  </p>
                  <p className="text-sm font-semibold text-slate-700">{selectedVendor.phone || '-'}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                    <FileText className="h-3 w-3" /> 팩스 번호
                  </p>
                  <p className="text-sm font-semibold text-slate-700">{selectedVendor.fax || '-'}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                    <LayoutGrid className="h-3 w-3" /> 사업자 번호
                  </p>
                  <p className="text-sm font-semibold text-slate-700">{selectedVendor.businessNumber || '-'}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                    <Mail className="h-3 w-3" /> 이메일
                  </p>
                  <p className="text-sm font-semibold text-indigo-600 truncate">{selectedVendor.email || '-'}</p>
                </div>
                {(isAdminMode || (selectedVendor && isVerified === selectedVendor.id)) && (
                  <div className="col-span-full grid grid-cols-1 lg:grid-cols-3 gap-4 border-t border-slate-200 mt-2 pt-4">
                    <div className="lg:col-span-2 space-y-2">
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                        <LinkIcon className="h-3 w-3" /> 업체 전용 접속 링크 (Admin Only)
                      </p>
                      <div className="flex items-center gap-4 bg-indigo-50/30 border border-indigo-100 rounded-xl p-3">
                        <div className="bg-white p-2 rounded-lg shadow-sm">
                          <QRCodeCanvas 
                            value={`${window.location.origin}/?v=${selectedVendor.id}`} 
                            size={64}
                            level="H"
                            includeMargin={false}
                          />
                        </div>
                        <div className="flex-1 space-y-2">
                          <code className="text-[10px] text-indigo-600 font-mono break-all block leading-tight">
                            {`${window.location.origin}/?v=${selectedVendor.id}`}
                          </code>
                          <button 
                            onClick={() => copyVendorLink(selectedVendor.id)}
                            className="w-full px-3 py-1.5 bg-white border border-indigo-200 text-indigo-600 text-[10px] font-bold rounded shadow-sm hover:bg-indigo-50 transition-colors flex items-center justify-center gap-2"
                          >
                            <LinkIcon className="h-3 w-3" /> 링크 복사
                          </button>
                        </div>
                      </div>
                    </div>
                    <div className="space-y-2">
                      <p className="text-[10px] font-bold text-amber-600 uppercase tracking-wider flex items-center gap-1">
                        <Lock className="h-3 w-3" /> 현재 비밀번호
                      </p>
                      <div className="bg-amber-50 border border-amber-100 rounded-lg p-2.5 text-sm font-black text-amber-700 flex items-center justify-between">
                        <span>{selectedVendor.password}</span>
                        <Unlock className="h-3 w-3 opacity-30" />
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </header>

            {/* Content Area */}
            <div className="flex-1 p-8 overflow-auto relative">
              <div className="space-y-8">
                {/* Notes Section - Collapsible */}
                {selectedVendor.notes && (
                  <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
                    <button 
                      onClick={() => setIsNotesExpanded(!isNotesExpanded)}
                      className="w-full flex items-center justify-between p-4 bg-slate-50 border-b border-slate-100 hover:bg-slate-100 transition-colors"
                    >
                      <div className="flex items-center gap-2">
                        <FileText className="h-4 w-4 text-amber-500" />
                        <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest">
                          업체 특이사항 및 메모
                        </h3>
                      </div>
                      <ChevronRight className={`h-4 w-4 text-slate-400 transition-transform ${isNotesExpanded ? 'rotate-90' : ''}`} />
                    </button>
                    {isNotesExpanded && (
                      <div className="p-4 text-sm text-slate-700 whitespace-pre-wrap leading-relaxed border-t border-slate-50">
                        {selectedVendor.notes}
                      </div>
                    )}
                  </div>
                )}

                {/* Handled Items Comparison Matrix - Collapsible */}
                <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
                  <div 
                    onClick={() => setIsMatrixExpanded(!isMatrixExpanded)}
                    className="w-full flex items-center justify-between p-4 bg-slate-900 border-b border-slate-800 hover:bg-slate-800 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center gap-2 text-white">
                      <Table className="h-4 w-4 text-indigo-400" />
                      <h3 className="text-sm font-black uppercase tracking-widest">취급 품목 비교 (구분)</h3>
                    </div>
                    <div className="flex items-center gap-3">
                      {canManageItems && isMatrixExpanded && (
                        <button 
                          onClick={(e) => {
                            e.stopPropagation();
                            setIsAddingComparisonRow(true);
                          }}
                          className="p-1 px-3 bg-white/10 hover:bg-white/20 text-white rounded text-[10px] font-bold transition-all flex items-center gap-1"
                        >
                          <Plus className="h-3 w-3" />
                          구분 추가
                        </button>
                      )}
                      <ChevronRight className={`h-4 w-4 text-slate-400 transition-transform ${isMatrixExpanded ? 'rotate-90' : ''}`} />
                    </div>
                  </div>
                  
                  {isMatrixExpanded && (
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse min-w-[800px]">
                        <thead className="bg-slate-50 text-slate-500 text-[10px] font-bold uppercase tracking-wider">
                          <tr>
                            <th className="px-6 py-3 border-b border-slate-200 w-1/4">구분품명</th>
                            <th className="px-6 py-3 border-b border-slate-200">업체명1</th>
                            <th className="px-6 py-3 border-b border-slate-200">업체명2</th>
                            <th className="px-6 py-3 border-b border-slate-200">업체명3</th>
                            <th className="px-6 py-3 border-b border-slate-200">업체명4</th>
                            {canManageItems && <th className="px-6 py-3 border-b border-slate-200 w-16"></th>}
                          </tr>
                        </thead>
                        <tbody className="text-slate-600 text-sm divide-y divide-slate-100">
                          {comparisonRows.map((row, index) => (
                            <tr key={row.id} className="hover:bg-slate-50/50 transition-colors group">
                              <td className="px-6 py-3 font-bold text-slate-900">
                                {canManageItems ? (
                                  <input 
                                    defaultValue={row.category} 
                                    onChange={(e) => handleMatrixChange(row.id, 'category', e.target.value)}
                                    className="w-full bg-transparent border-none p-0 focus:ring-0 font-bold"
                                  />
                                ) : row.category}
                              </td>
                              <td className="px-6 py-3">
                                {canManageItems ? (
                                  <input 
                                    defaultValue={row.vendor1} 
                                    onChange={(e) => handleMatrixChange(row.id, 'vendor1', e.target.value)}
                                    className="w-full bg-transparent border-none p-0 focus:ring-0 text-slate-600"
                                  />
                                ) : row.vendor1}
                              </td>
                              <td className="px-6 py-3">
                                {canManageItems ? (
                                  <input 
                                    defaultValue={row.vendor2} 
                                    onChange={(e) => handleMatrixChange(row.id, 'vendor2', e.target.value)}
                                    className="w-full bg-transparent border-none p-0 focus:ring-0 text-slate-600"
                                  />
                                ) : row.vendor2}
                              </td>
                              <td className="px-6 py-3">
                                {canManageItems ? (
                                  <input 
                                    defaultValue={row.vendor3} 
                                    onChange={(e) => handleMatrixChange(row.id, 'vendor3', e.target.value)}
                                    className="w-full bg-transparent border-none p-0 focus:ring-0 text-slate-600"
                                  />
                                ) : row.vendor3}
                              </td>
                              <td className="px-6 py-3">
                                {canManageItems ? (
                                  <input 
                                    defaultValue={row.vendor4} 
                                    onChange={(e) => handleMatrixChange(row.id, 'vendor4', e.target.value)}
                                    className="w-full bg-transparent border-none p-0 focus:ring-0 text-slate-600"
                                  />
                                ) : row.vendor4}
                              </td>
                              {canManageItems && (
                                <td className="px-6 py-3 text-right">
                                  <div className="flex items-center justify-end gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                    {savingMatrixId === row.id && (
                                      <span className="text-[10px] text-emerald-500 font-bold mr-2 animate-pulse whitespace-nowrap">저장 중...</span>
                                    )}
                                    <button 
                                      onClick={() => moveComparisonRow(index, 'up')}
                                      disabled={index === 0}
                                      className="p-1 hover:text-indigo-500 disabled:opacity-30"
                                    >
                                      <ChevronUp className="h-3 w-3" />
                                    </button>
                                    <button 
                                      onClick={() => moveComparisonRow(index, 'down')}
                                      disabled={index === comparisonRows.length - 1}
                                      className="p-1 hover:text-indigo-500 disabled:opacity-30"
                                    >
                                      <ChevronDown className="h-3 w-3" />
                                    </button>
                                    <button 
                                      onClick={() => deleteComparisonRow(row.id)}
                                      className="p-1 hover:text-red-500 ml-1"
                                    >
                                      <Trash2 className="h-3 w-3" />
                                    </button>
                                  </div>
                                </td>
                              )}
                            </tr>
                          ))}
                          {comparisonRows.length === 0 && (
                            <tr>
                              <td colSpan={canManageItems ? 6 : 5} className="px-6 py-10 text-center text-slate-400 text-xs italic">
                                등록된 비교 품목 구분이 없습니다. {canManageItems && "내용을 입력하여 추가해 보세요."}
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                {/* Price Table File Preview Section */}
                {selectedVendor.priceTableUrl && isVerified === selectedVendor.id && (
                  <div className="mb-0">
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest flex items-center gap-2">
                        <FileText className="h-4 w-4 text-indigo-500" />
                        첨부 단가표 (원본 파일)
                      </h3>
                      <a 
                        href={selectedVendor.priceTableUrl} 
                        target="_blank" 
                        rel="noopener noreferrer"
                        className="text-[10px] font-bold text-indigo-600 hover:underline flex items-center gap-1"
                      >
                        <Download className="h-3 w-3" /> 새 창에서 파일 열기
                      </a>
                    </div>
                    
                    <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                      {selectedVendor.priceTableFileType === 'image' ? (
                        <div className="p-4 bg-slate-50 flex justify-center">
                          <img 
                            src={selectedVendor.priceTableUrl} 
                            alt="Price Table" 
                            className="max-w-full h-auto rounded shadow-lg border border-slate-200" 
                            referrerPolicy="no-referrer" 
                          />
                        </div>
                      ) : selectedVendor.priceTableFileType === 'pdf' ? (
                        <iframe src={selectedVendor.priceTableUrl} className="w-full h-[800px] border-none" title="PDF Price Table" />
                      ) : selectedVendor.priceTableFileType === 'excel' ? (
                        <div className="overflow-auto max-h-[800px] p-4 bg-slate-50">
                          {excelPreviewData ? (
                            <div className="inline-block min-w-full align-middle border border-slate-200 rounded-xl overflow-hidden bg-white shadow-sm font-sans">
                              <table className="min-w-full divide-y divide-slate-200 text-xs">
                                <tbody className="divide-y divide-slate-200">
                                  {excelPreviewData.map((row: any[], i) => (
                                    <tr key={i} className={i === 0 ? "bg-slate-100 font-bold" : "hover:bg-slate-50 transition-colors"}>
                                      {row.map((cell: any, jValue) => (
                                        <td key={jValue} className="px-4 py-3 whitespace-nowrap border-r border-slate-100 last:border-r-0 text-slate-700">
                                          {cell !== null && cell !== undefined ? String(cell) : ''}
                                        </td>
                                      ))}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : (
                            <div className="flex flex-col items-center justify-center py-20 text-slate-400 gap-4">
                              <Loader2 className="h-8 w-8 animate-spin text-indigo-400" />
                              <p className="text-xs font-bold uppercase tracking-widest text-slate-300">엑셀 데이터 파싱 중...</p>
                            </div>
                          )}
                        </div>
                      ) : (
                        <div className="p-20 text-center text-slate-400">
                          <Building2 className="h-12 w-12 mx-auto mb-4 opacity-20" />
                          <p className="text-sm font-medium">미리보기를 지원하지 않는 형식입니다.</p>
                          <p className="text-xs">{selectedVendor.priceTableFileName || '알 수 없는 파일'}</p>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm flex flex-col">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-slate-50 text-slate-500 text-xs font-bold uppercase tracking-wider">
                      <tr>
                        <th className="px-6 py-4 border-b border-slate-200">번호</th>
                        <th className="px-6 py-4 border-b border-slate-200">품목</th>
                        <th className="px-6 py-4 border-b border-slate-200">규격</th>
                        <th className="px-6 py-4 border-b border-slate-200">단위</th>
                        <th className="px-6 py-4 border-b border-slate-200 text-right">협가</th>
                        <th className="px-6 py-4 border-b border-slate-200 text-right">네고율</th>
                        <th className="px-6 py-4 border-b border-slate-200 text-right">단가 (최종)</th>
                        <th className="px-6 py-4 border-b border-slate-200">제조사</th>
                        <th className="px-6 py-4 border-b border-slate-200">비고</th>
                        <th className="px-6 py-4 border-b border-slate-200">확인</th>
                        <th className="px-6 py-4 border-b border-slate-200 text-right">관리</th>
                      </tr>
                    </thead>
                    <tbody className="text-slate-600 text-sm divide-y divide-slate-100">
                      <AnimatePresence mode="popLayout">
                        {priceItems.map((item, idx) => (
                          <motion.tr 
                            layout
                            key={item.id}
                            initial={{ opacity: 0 }}
                            animate={{ opacity: 1 }}
                            className={`group hover:bg-slate-50/50 transition-colors ${item.isConfirmed ? 'bg-emerald-50/50' : ''}`}
                          >
                            <td className="px-6 py-4">
                              <div className="flex items-center gap-2">
                                {String(idx + 1).padStart(2, '0')}
                                {item.isConfirmed && (
                                  <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse" />
                                )}
                              </div>
                            </td>
                            <td className="px-6 py-4 font-medium text-slate-800">
                              {editingPriceId === item.id ? (
                                <input 
                                  autoFocus
                                  defaultValue={item.itemName} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { itemName: e.target.value })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { itemName: (e.target as HTMLInputElement).value })}
                                  className="w-full bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none"
                                />
                              ) : (
                                <span onDoubleClick={() => canManageItems && setEditingPriceId(item.id)}>{item.itemName}</span>
                              )}
                            </td>
                            <td className="px-6 py-4">
                              {editingPriceId === item.id ? (
                                <input 
                                  defaultValue={item.spec} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { spec: e.target.value })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { spec: (e.target as HTMLInputElement).value })}
                                  className="w-full bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none"
                                />
                              ) : item.spec || '-'}
                            </td>
                            <td className="px-6 py-4">
                              {editingPriceId === item.id ? (
                                <input 
                                  defaultValue={item.unit} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { unit: e.target.value })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { unit: (e.target as HTMLInputElement).value })}
                                  className="w-20 bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none"
                                />
                              ) : item.unit || '-'}
                            </td>
                            <td className="px-6 py-4 text-right">
                              {editingPriceId === item.id ? (
                                <input 
                                  type="number"
                                  defaultValue={item.costPrice} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { costPrice: Number(e.target.value) })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { costPrice: Number((e.target as HTMLInputElement).value) })}
                                  className="w-24 bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none text-right"
                                />
                              ) : <span className="text-slate-400">{item.costPrice?.toLocaleString() || '0'}원</span>}
                            </td>
                            <td className="px-6 py-4 text-right font-bold text-indigo-500">
                              {editingPriceId === item.id ? (
                                <input 
                                  type="number"
                                  step="any"
                                  defaultValue={item.negoRate} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { negoRate: Number(e.target.value) })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { negoRate: Number((e.target as HTMLInputElement).value) })}
                                  className="w-16 bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none text-right"
                                />
                              ) : <span>{item.negoRate || 0}%</span>}
                            </td>
                            <td className="px-6 py-4 text-right font-bold text-slate-900 bg-slate-50/50">
                              {item.unitPrice.toLocaleString()}원
                            </td>
                            <td className="px-6 py-4 text-xs font-bold text-indigo-600">
                              {editingPriceId === item.id ? (
                                <input 
                                  defaultValue={item.maker} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { maker: e.target.value })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { maker: (e.target as HTMLInputElement).value })}
                                  className="w-full bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none"
                                />
                              ) : item.maker || '-'}
                            </td>
                            <td className="px-6 py-4 text-xs text-slate-400">
                              {editingPriceId === item.id ? (
                                <input 
                                  defaultValue={item.remarks} 
                                  onBlur={(e) => handleUpdatePriceItem(item.id, { remarks: e.target.value })}
                                  onKeyDown={(e) => e.key === 'Enter' && handleUpdatePriceItem(item.id, { remarks: (e.target as HTMLInputElement).value })}
                                  className="w-full bg-white border border-indigo-200 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-indigo-400 outline-none"
                                />
                              ) : item.remarks || '-'}
                            </td>
                            <td className="px-6 py-4">
                              <input 
                                type="checkbox" 
                                checked={item.isConfirmed || false} 
                                onChange={() => togglePriceItemConfirmation(item.id, item.isConfirmed || false)}
                                className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                              />
                            </td>
                            <td className="px-6 py-4 text-right">
                              {canManageItems && (
                                <div className="flex items-center justify-end gap-1">
                                  <button 
                                    onClick={() => setEditingPriceId(editingPriceId === item.id ? null : item.id)}
                                    className={`p-1.5 rounded transition-all ${editingPriceId === item.id ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:text-indigo-500 group-hover:opacity-100 opacity-0'}`}
                                  >
                                    <Edit2 className="h-4 w-4" />
                                  </button>
                                  <button 
                                    onClick={() => deletePriceItem(item.id)}
                                    className="opacity-0 group-hover:opacity-100 p-1.5 text-slate-300 hover:text-red-500 transition-all"
                                  >
                                    <Trash2 className="h-4 w-4" />
                                  </button>
                                </div>
                              )}
                            </td>
                          </motion.tr>
                        ))}
                      </AnimatePresence>
                      {priceItems.length === 0 && (
                        <tr>
                          <td colSpan={11} className="px-6 py-20 text-center">
                            <p className="text-slate-400">등록된 단가 정보가 없습니다.</p>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Password Overlay */}
                <AnimatePresence>
                  {isVerified !== selectedVendor.id && (
                    <motion.div 
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="absolute inset-0 bg-slate-900/60 backdrop-blur-[4px] flex items-center justify-center m-8 rounded-xl z-20"
                    >
                      <motion.div 
                        initial={{ scale: 0.95, opacity: 0 }}
                        animate={{ scale: 1, opacity: 1 }}
                        className="bg-white p-8 rounded-2xl shadow-2xl w-[400px] text-center border border-slate-200"
                      >
                        <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-full flex items-center justify-center mx-auto mb-6">
                          <Lock className="h-8 w-8" />
                        </div>
                        <h3 className="text-xl font-bold text-slate-800 mb-2">업체 접속 인증</h3>
                        <p className="text-slate-500 text-sm mb-6">
                          '{selectedVendor.name}' 전용 단가표입니다.<br />접근을 위해 비밀번호를 입력해주세요.
                        </p>
                        <form onSubmit={verifyPassword} className="space-y-4">
                          <input 
                            autoFocus
                            type="password" 
                            placeholder="비밀번호 입력" 
                            value={passwordInput}
                            onChange={(e) => setPasswordInput(e.target.value)}
                            className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg text-center tracking-[0.5em] focus:outline-none focus:ring-2 focus:ring-indigo-500 transition-all"
                          />
                          <button className="w-full py-3 bg-indigo-600 text-white font-bold rounded-lg hover:bg-indigo-700 shadow-lg shadow-indigo-200 transition-all active:scale-95">
                            접속하기
                          </button>
                        </form>
                      </motion.div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </div>
          </div>

          {/* Summary Footer */}
            <footer className="h-12 px-8 flex items-center justify-between text-[11px] text-slate-400 bg-white border-t border-slate-200 shrink-0">
              <div className="flex gap-4">
                <span>총 품목 수: {priceItems.length}개</span>
                <span>환율 기준: {usdToKrw.toLocaleString('ko-KR', { maximumFractionDigits: 2 })}원 (USD) {lastRateUpdate && <span className="opacity-60 ml-1">({lastRateUpdate} 기준)</span>}</span>
              </div>
              <div>© 2024 ERP Cost Management System</div>
            </footer>
          </>
        ) : (
          <div className="flex-1 flex flex-col items-center justify-center p-20 text-center">
            <Building2 className="h-16 w-16 text-slate-200 mb-6" />
            <h2 className="text-xl font-bold text-slate-900 mb-2">업체를 선택해 주세요</h2>
            <p className="text-slate-500 text-sm max-w-xs mx-auto">
              왼쪽 리스트에서 관리할 업체를 선택하시면 해당 업체의 전용 단가표를 조회할 수 있습니다.
            </p>
          </div>
        )}
      </main>

      {/* Modals */}
      <AnimatePresence>
        {isAddingVendor && (
          <Modal key="modal-add-vendor" title="신규 업체 마스터 등록" onClose={() => setIsAddingVendor(false)}>
            <form onSubmit={handleAddVendor} className="space-y-8">
              <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">거래처 정식 명칭 *</label>
                  <input name="name" required placeholder="예) (주)미래정밀" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all placeholder:text-slate-300" />
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">업체 메모 (비고)</label>
                  <textarea name="notes" rows={4} placeholder="업체와 관련된 상세 정보, 특이사항, 연락 기록 등을 자유롭게 입력하세요." className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-medium text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all"></textarea>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">대표자 성함</label>
                  <input name="representative" placeholder="성함 입력" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all placeholder:text-slate-300" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">사업자 등록번호</label>
                  <input name="businessNumber" placeholder="000-00-00000" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all placeholder:text-slate-300" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">회사 전화번호</label>
                  <input name="phone" placeholder="02-000-0000" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all placeholder:text-slate-300" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">팩스 번호</label>
                  <input name="fax" placeholder="02-000-0000" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all placeholder:text-slate-300" />
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">공식 이메일 *</label>
                  <input name="email" type="email" required placeholder="contact@company.com" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all placeholder:text-slate-300" />
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">단가표 파일 (Excel, PDF, Image)</label>
                  <div className="relative group/upload">
                    <input 
                      type="file" 
                      accept=".xlsx,.xls,.csv,.pdf,image/*" 
                      onChange={(e) => setPriceTableFile(e.target.files?.[0] || null)}
                      className="hidden" 
                      id="price-table-upload"
                    />
                    <label 
                      htmlFor="price-table-upload" 
                      className="flex flex-col items-center justify-center gap-2 w-full p-8 border-2 border-dashed border-slate-200 rounded-2xl bg-slate-50 hover:bg-white hover:border-indigo-400 transition-all cursor-pointer group"
                    >
                      {priceTableFile ? (
                        <>
                          <FileText className="h-8 w-8 text-indigo-500" />
                          <span className="text-sm font-bold text-slate-700">{priceTableFile.name}</span>
                          <span className="text-[10px] text-slate-400">클릭하여 파일 변경</span>
                        </>
                      ) : (
                        <>
                          <Upload className="h-8 w-8 text-slate-300 group-hover:text-indigo-400 transition-colors" />
                          <span className="text-sm font-bold text-slate-500 group-hover:text-indigo-600 transition-colors">파일 선택 또는 드래그</span>
                          <span className="text-[10px] text-slate-400 uppercase tracking-widest font-black">XLSX, PDF, JPG/PNG 지원</span>
                        </>
                      )}
                    </label>
                  </div>
                </div>
                <div className="col-span-full py-4 border-y border-slate-100 bg-indigo-50/20 px-4 -mx-4">
                   <label className="flex items-center gap-2 text-[10px] font-black uppercase tracking-[0.2em] text-indigo-600 mb-2">
                    <Lock className="h-3 w-3" />
                    접속 보안 비밀번호 설정 *
                   </label>
                   <input name="password" required type="password" placeholder="업체 전용 비밀번호 4-20자" className="w-full rounded-2xl border-2 border-indigo-100 bg-white p-5 text-center text-xl font-black tracking-[1em] text-indigo-700 focus:border-indigo-500 focus:outline-none transition-all" />
                   <p className="mt-2 text-[10px] text-indigo-400 font-bold text-center">이 비밀번호는 업체의 단가표를 조회할 때 사용됩니다.</p>
                </div>
                <div className="col-span-full grid grid-cols-1 md:grid-cols-2 gap-4">
                  <label className="flex items-center gap-3 cursor-pointer p-4 bg-slate-50 rounded-2xl border-2 border-slate-100 hover:bg-white hover:border-indigo-200 transition-all">
                    <input type="checkbox" name="useRounding" className="w-5 h-5 rounded-lg border-2 border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                    <div>
                      <p className="text-sm font-bold text-slate-800">단가 10원 단위 반올림/절사</p>
                      <p className="text-[10px] text-slate-500 font-medium tracking-tight">5원 이하 절사, 5원 이상 반올림</p>
                    </div>
                  </label>
                  <label className="flex items-center gap-3 cursor-pointer p-4 bg-indigo-50/50 rounded-2xl border-2 border-indigo-100 hover:bg-white hover:border-indigo-400 transition-all">
                    <input type="checkbox" name="masterCustomFlag" className="w-5 h-5 rounded-lg border-2 border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                    <div>
                      <p className="text-sm font-bold text-slate-800">마스터 커스텀 설정 적용</p>
                      <p className="text-[10px] text-indigo-400 font-medium tracking-tight">비고란 옆 체크박스 활성화</p>
                    </div>
                  </label>
                </div>
              </div>
              <button 
                type="submit" 
                disabled={isUploading}
                className="w-full rounded-3xl bg-indigo-600 py-5 text-lg font-black text-white shadow-2xl shadow-indigo-200 transition-all hover:bg-indigo-700 hover:scale-[1.02] active:scale-100 disabled:opacity-70 disabled:cursor-not-allowed flex items-center justify-center gap-2"
              >
                {isUploading && <Loader2 className="h-5 w-5 animate-spin" />}
                {isUploading ? '업로드 및 데이터 생성 중...' : '업체 마스터 데이터 생성'}
              </button>
            </form>
          </Modal>
        )}

        {isAddingComparisonRow && (
          <Modal key="modal-add-comparison" title="취급 품목 구분 추가" onClose={() => setIsAddingComparisonRow(false)}>
            <form onSubmit={handleAddComparisonRow} className="space-y-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">구분 품명 *</label>
                <input name="category" required placeholder="예) 강관, 밸브, 피팅 등" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">업체명 1</label>
                  <input name="vendor1" placeholder="업체명 입력" className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-3 text-sm focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">업체명 2</label>
                  <input name="vendor2" placeholder="업체명 입력" className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-3 text-sm focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">업체명 3</label>
                  <input name="vendor3" placeholder="업체명 입력" className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-3 text-sm focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">업체명 4</label>
                  <input name="vendor4" placeholder="업체명 입력" className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-3 text-sm focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
              </div>
              <div className="pt-4">
                <button type="submit" className="w-full rounded-3xl bg-indigo-600 py-4 text-white font-black hover:bg-slate-900 transition-all shadow-xl shadow-indigo-100">
                  구분 항목 저장
                </button>
              </div>
            </form>
          </Modal>
        )}

        {isAddingItem && (
          <Modal key="modal-add-item" title="신규 품목 단가 등록" onClose={() => setIsAddingItem(false)}>
            <form onSubmit={handleAddItem} className="space-y-8">
              <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                <div className="col-span-full space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">품목명 (Item Name) *</label>
                  <input name="itemName" required placeholder="예) 스텐레스 앵글" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-black text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">규격 (Spec)</label>
                  <input name="spec" placeholder="100x100x5T" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">단위 (Unit)</label>
                  <input name="unit" placeholder="kg, ea, m 등" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">협가 (Negotiated Price) *</label>
                   <div className="relative">
                    <span className="absolute left-4 top-1/2 -translate-y-1/2 text-lg font-bold text-slate-300">₩</span>
                    <input name="costPrice" type="number" required placeholder="0" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 pl-10 font-black text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-indigo-600">네고율 (Nego Rate %) *</label>
                  <div className="relative">
                    <span className="absolute right-6 top-1/2 -translate-y-1/2 text-xl font-black text-indigo-300">%</span>
                    <input name="negoRate" type="number" step="any" defaultValue="0" className="w-full rounded-2xl border-2 border-indigo-100 bg-indigo-50/30 p-4 pr-12 text-xl font-black text-indigo-700 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                  </div>
                </div>
                <div className="col-span-full grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">제조사 (Maker)</label>
                    <input name="maker" placeholder="메이커명 입력" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">구매 조건 및 비고</label>
                    <input name="remarks" placeholder="항목별 특이사항 (예: 최소 주문 500개)" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-medium text-slate-700 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                  </div>
                </div>
              </div>
              <button type="submit" className="w-full rounded-3xl bg-slate-900 py-5 text-xl font-black text-white shadow-2xl transition-all hover:bg-slate-800 hover:scale-[1.02]">
                단가 데이터 업데이트
              </button>
            </form>
          </Modal>
        )}

        {isEditingVendorInfo && selectedVendor && (
          <Modal key="modal-edit-vendor" title="업체 정보 수정" onClose={() => setIsEditingVendorInfo(false)}>
            <form onSubmit={handleUpdateVendorInfo} className="space-y-8">
              <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                <div className="space-y-2 col-span-full text-center py-4 bg-slate-50 rounded-2xl border-2 border-dashed border-slate-100 flex flex-col items-center justify-center">
                   <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">업체 고유 링크</p>
                   <code className="text-xs text-indigo-500 font-mono bg-white px-3 py-1 rounded-full border border-indigo-50">{window.location.origin}/?v={selectedVendor.id}</code>
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">거래처 정식 명칭 *</label>
                  <input name="name" required defaultValue={selectedVendor.name} className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">업체 메모 (비고)</label>
                  <textarea name="notes" defaultValue={selectedVendor.notes} rows={4} placeholder="업체와 관련된 상세 정보, 특이사항, 연락 기록 등을 자유롭게 입력하세요." className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-medium text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all"></textarea>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">대표자명</label>
                  <input name="representative" defaultValue={selectedVendor.representative} className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">단가표 파일 (변경 시 선택)</label>
                  <div className="relative group/upload">
                    <input 
                      type="file" 
                      accept=".xlsx,.xls,.csv,.pdf,image/*" 
                      onChange={(e) => setPriceTableFile(e.target.files?.[0] || null)}
                      className="hidden" 
                      id="price-table-edit-upload"
                    />
                    <label 
                      htmlFor="price-table-edit-upload" 
                      className="flex flex-col items-center justify-center gap-2 w-full p-8 border-2 border-dashed border-slate-200 rounded-2xl bg-slate-50 hover:bg-white hover:border-indigo-400 transition-all cursor-pointer group"
                    >
                      {priceTableFile ? (
                        <>
                          <FileText className="h-8 w-8 text-indigo-500" />
                          <span className="text-sm font-bold text-slate-700">{priceTableFile.name}</span>
                          <span className="text-[10px] text-slate-400">변경될 파일</span>
                        </>
                      ) : selectedVendor.priceTableUrl ? (
                        <>
                          <div className="flex items-center gap-2 text-indigo-600 font-bold mb-1">
                            <LinkIcon className="h-4 w-4" />
                            <span className="text-sm">현재 파일: {selectedVendor.priceTableFileName || '파일'}</span>
                          </div>
                          <Upload className="h-8 w-8 text-slate-300 group-hover:text-indigo-400 transition-colors" />
                          <span className="text-[10px] text-slate-400 uppercase tracking-widest font-black">클릭하여 새 파일로 교체</span>
                        </>
                      ) : (
                        <>
                          <Upload className="h-8 w-8 text-slate-300 group-hover:text-indigo-400 transition-colors" />
                          <span className="text-sm font-bold text-slate-500 group-hover:text-indigo-600 transition-colors">파일 선택 또는 드래그</span>
                        </>
                      )}
                    </label>
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">사업자 등록번호</label>
                  <input name="businessNumber" defaultValue={selectedVendor.businessNumber} className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">회사 전화번호</label>
                  <input name="phone" defaultValue={selectedVendor.phone} className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">팩스 번호</label>
                  <input name="fax" defaultValue={selectedVendor.fax} className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="space-y-2 col-span-full">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">공식 이메일 *</label>
                  <input name="email" type="email" required defaultValue={selectedVendor.email} className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                </div>
                <div className="col-span-full grid grid-cols-1 md:grid-cols-2 gap-4">
                  <label className="flex items-center gap-3 cursor-pointer p-4 bg-slate-50 rounded-2xl border-2 border-slate-100 hover:bg-white hover:border-indigo-200 transition-all">
                    <input type="checkbox" name="useRounding" defaultChecked={selectedVendor.useRounding} className="w-5 h-5 rounded-lg border-2 border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                    <div>
                      <p className="text-sm font-bold text-slate-800">단가 10원 단위 반올림/절사</p>
                      <p className="text-[10px] text-slate-500 font-medium tracking-tight">5원 이하 절사, 5원 이상 반올림</p>
                    </div>
                  </label>
                  <label className="flex items-center gap-3 cursor-pointer p-4 bg-indigo-50/50 rounded-2xl border-2 border-indigo-100 hover:bg-white hover:border-indigo-400 transition-all">
                    <input type="checkbox" name="masterCustomFlag" defaultChecked={selectedVendor.masterCustomFlag} className="w-5 h-5 rounded-lg border-2 border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                    <div>
                      <p className="text-sm font-bold text-slate-800">마스터 커스텀 설정 적용</p>
                      <p className="text-[10px] text-indigo-400 font-medium tracking-tight">비고란 옆 체크박스 활성화</p>
                    </div>
                  </label>
                </div>
              </div>
              <button 
                type="submit" 
                disabled={isUploading}
                className="w-full py-5 bg-indigo-600 text-white font-black uppercase tracking-[0.1em] rounded-2xl hover:bg-slate-900 disabled:opacity-70 disabled:cursor-not-allowed transition-all shadow-xl shadow-indigo-100 flex items-center justify-center gap-2"
              >
                {isUploading && <Loader2 className="h-5 w-5 animate-spin" />}
                {isUploading ? '업로드 및 수정 중...' : '수정 사항 저장하기'}
              </button>
            </form>
          </Modal>
        )}

        {isChangingPassword && selectedVendor && (
          <Modal key="modal-change-password" title="접속 비밀번호 변경" onClose={() => setIsChangingPassword(false)}>
            <form onSubmit={handleUpdatePassword} className="space-y-6">
              <div className="text-center py-4">
                <div className="w-16 h-16 bg-amber-50 text-amber-600 rounded-full flex items-center justify-center mx-auto mb-4">
                  <Lock className="h-8 w-8" />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">보안 설정 변경</h3>
                <p className="text-slate-500 text-sm">업체 관리용 비밀번호를 새로 설정합니다.</p>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">새 비밀번호 (최소 4자)</label>
                <input 
                  name="newPassword"
                  type="password" 
                  required
                  placeholder="새로운 비밀번호 입력"
                  className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-5 text-center text-2xl font-black tracking-[0.5em] text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" 
                />
              </div>
              <div className="flex gap-3">
                <button 
                  type="button"
                  onClick={() => setIsChangingPassword(false)}
                  className="flex-1 py-4 bg-slate-100 text-slate-600 font-bold rounded-2xl hover:bg-slate-200 transition-all"
                >
                  취소
                </button>
                <button 
                  type="submit" 
                  className="flex-1 py-4 bg-indigo-600 text-white font-bold rounded-2xl hover:bg-slate-900 transition-all shadow-xl shadow-indigo-100"
                >
                  비밀번호 저장
                </button>
              </div>
            </form>
          </Modal>
        )}

        {isBulkItemUploadOpen && (
          <Modal key="modal-bulk-items" title="단가표 일괄 등록 (Excel)" onClose={() => setIsBulkItemUploadOpen(false)}>
            <div className="space-y-6">
              <div className="text-center py-4">
                <div className="w-16 h-16 bg-slate-100 text-slate-900 rounded-full flex items-center justify-center mx-auto mb-6">
                  <FileText className="h-8 w-8" />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">업체 단가표 일괄 등록</h3>
                <p className="text-slate-500 text-sm mb-4">
                  표준 엑셀 양식을 다운로드하여 작성 후 업로드해주세요.<br/>
                  등록된 품목은 현재 업체의 단가 목록에 추가됩니다.
                </p>
                
                <button 
                  onClick={downloadPriceTemplate}
                  className="inline-flex items-center gap-2 px-4 py-2 bg-indigo-50 text-indigo-700 rounded-full text-xs font-bold hover:bg-indigo-100 transition-all"
                >
                  <Download className="h-3 w-3" />
                  표준 엑셀 양식(Template) 다운로드
                </button>
              </div>

              <div 
                className="border-2 border-dashed border-slate-200 rounded-3xl p-10 text-center bg-slate-50 hover:border-indigo-400 hover:bg-white transition-all cursor-pointer group"
                onClick={() => document.getElementById('price-items-bulk-upload')?.click()}
              >
                <input 
                  id="price-items-bulk-upload"
                  type="file" 
                  className="hidden" 
                  accept=".xlsx, .xls"
                  onChange={(e) => setBulkItemFile(e.target.files?.[0] || null)}
                />
                <div className="w-12 h-12 bg-white rounded-2xl shadow-sm border border-slate-100 flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
                  <Upload className="h-6 w-6 text-indigo-500" />
                </div>
                <p className="text-sm text-slate-600 font-bold mb-1">
                  {bulkItemFile ? bulkItemFile.name : "엑셀 파일을 선택하거나 드래그하세요"}
                </p>
                <p className="text-[10px] text-slate-400 font-medium uppercase tracking-widest">.XLSX, .XLS 지원</p>
                {bulkItemFile && (
                  <div className="mt-4 inline-flex items-center gap-1.5 px-3 py-1 bg-emerald-50 text-emerald-600 rounded-full text-[10px] font-black uppercase">
                    <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse" />
                    Ready to upload
                  </div>
                )}
              </div>

              <div className="flex gap-4 pt-4">
                <button 
                  onClick={() => setIsBulkItemUploadOpen(false)}
                  className="flex-1 py-4 bg-slate-100 text-slate-600 font-bold rounded-2xl hover:bg-slate-200 transition-all"
                >
                  취소
                </button>
                <button 
                  onClick={handleBulkItemUpload}
                  disabled={!bulkItemFile}
                  className="flex-1 py-4 bg-indigo-600 text-white font-black rounded-2xl hover:bg-slate-900 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-xl shadow-indigo-200"
                >
                  업로드 및 데이터 생성
                </button>
              </div>
            </div>
          </Modal>
        )}

        {isSeedModalOpen && (
          <Modal key="modal-seed-confirm" title="업무 일괄 등록 (Excel)" onClose={() => setIsSeedModalOpen(false)}>
            <div className="space-y-6">
              <div className="text-center py-4">
                <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-full flex items-center justify-center mx-auto mb-6">
                  <Upload className="h-8 w-8" />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">업체 대량 등록</h3>
                <p className="text-slate-500 text-sm">
                  엑셀 파일을 업로드하여 여러 업체를 한 번에 등록할 수 있습니다.<br/>
                  (컬럼명: 업체명, 대표자, 이메일, 전화번호, 팩스번호, 사업자번호, 비밀번호, 취급품목)
                </p>
              </div>

              <div 
                className="border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center bg-slate-50 hover:border-indigo-400 transition-colors cursor-pointer"
                onClick={() => document.getElementById('bulk-excel-upload')?.click()}
              >
                <input 
                  id="bulk-excel-upload"
                  type="file" 
                  className="hidden" 
                  accept=".xlsx, .xls"
                  onChange={(e) => setBulkFile(e.target.files?.[0] || null)}
                />
                <Download className="h-8 w-8 text-slate-300 mx-auto mb-4" />
                <p className="text-sm text-slate-600 font-bold">
                  {bulkFile ? bulkFile.name : "엑셀 파일을 선택하거나 여기로 드래그하세요"}
                </p>
                {bulkFile && (
                  <p className="text-[10px] text-indigo-500 mt-2 font-bold uppercase">파일 준비됨</p>
                )}
              </div>

              <div className="flex gap-3">
                <button 
                  onClick={() => setIsSeedModalOpen(false)}
                  className="flex-1 py-4 bg-slate-100 text-slate-600 font-bold rounded-2xl hover:bg-slate-200 transition-all"
                >
                  취소
                </button>
                <button 
                  onClick={handleBulkUpload}
                  disabled={!bulkFile}
                  className="flex-1 py-4 bg-indigo-600 text-white font-bold rounded-2xl hover:bg-slate-900 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-xl shadow-indigo-100"
                >
                  데이터 등록 시작
                </button>
              </div>
            </div>
          </Modal>
        )}

        {seedStatus && (
          <Modal key="modal-seed-status" title={seedStatus.isDone ? "등록 완료" : "데이터 등록 중..."} onClose={() => {
            if (seedStatus.isDone) setSeedStatus(null);
          }}>
            <div className="py-8 text-center">
              {!seedStatus.isDone ? (
                <>
                  <Loader2 className="h-12 w-12 text-indigo-500 animate-spin mx-auto mb-6" />
                  <div className="mb-4 h-2 w-full bg-slate-100 rounded-full overflow-hidden">
                    <motion.div 
                      className="h-full bg-indigo-500"
                      initial={{ width: 0 }}
                      animate={{ width: `${(seedStatus.current / seedStatus.total) * 100}%` }}
                    />
                  </div>
                  <p className="text-lg font-bold text-slate-800 mb-2">업체 데이터를 저장하고 있습니다.</p>
                  <p className="text-slate-400 text-sm">{seedStatus.current} / {seedStatus.total} 진행 중</p>
                </>
              ) : (
                <>
                  <div className="w-16 h-16 bg-green-50 text-green-600 rounded-full flex items-center justify-center mx-auto mb-6">
                    <Save className="h-8 w-8" />
                  </div>
                  <h3 className="text-xl font-bold text-slate-800 mb-2">모든 업체 등록 완료!</h3>
                  <p className="text-slate-500 text-sm mb-8">이미지 속 {seedStatus.total}개의 업체가 성공적으로 저장되었습니다.</p>
                  <button 
                    onClick={() => setSeedStatus(null)}
                    className="w-full py-4 bg-slate-900 text-white font-bold rounded-xl hover:bg-slate-800 transition-all"
                  >
                    확인
                  </button>
                </>
              )}
            </div>
          </Modal>
        )}

        {showAdminLogin && (
          <Modal key="modal-admin-login" title="관리자 시스템 로그인" onClose={() => setShowAdminLogin(false)}>
            <form onSubmit={handleAdminLogin} className="space-y-6">
              <div className="text-center py-4">
                <div className="w-16 h-16 bg-amber-50 text-amber-600 rounded-full flex items-center justify-center mx-auto mb-4">
                  <Lock className="h-8 w-8" />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">관리자 인증</h3>
                <p className="text-slate-500 text-sm">마스터 관리자 비밀번호를 입력해주세요.</p>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">비밀번호</label>
                <input 
                  type="password" 
                  value={adminPasswordInput}
                  onChange={(e) => setAdminPasswordInput(e.target.value)}
                  placeholder="Password..."
                  autoFocus
                  className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-amber-500 focus:bg-white focus:outline-none transition-all" 
                />
              </div>
              <button 
                type="submit" 
                className="w-full py-5 bg-amber-600 text-white font-black uppercase tracking-[0.1em] rounded-2xl hover:bg-slate-900 transition-all shadow-xl shadow-amber-100"
              >
                관리자 권한 확인
              </button>
            </form>
          </Modal>
        )}

        {/* Global Toast */}
        <AnimatePresence>
          {copySuccess && (
            <motion.div 
              key="toast-copy-success"
              initial={{ opacity: 0, y: 50 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 50 }}
              className="fixed bottom-8 left-1/2 -translate-x-1/2 z-[100] bg-slate-900 text-white px-6 py-3 rounded-full border border-slate-700 shadow-2xl flex items-center gap-3"
            >
              <LinkIcon className="h-4 w-4 text-indigo-400" />
              <span className="text-sm font-bold">링크가 클립보드에 복사되었습니다!</span>
            </motion.div>
          )}
        </AnimatePresence>
      </AnimatePresence>
    </div>
  );
}
