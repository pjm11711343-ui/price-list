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
  Info,
  Link as LinkIcon,
  ArrowUpDown,
  History,
  RotateCcw,
  ArrowLeftRight,
  GripVertical,
  BarChart3,
  Users
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip as RechartsTooltip, 
  ResponsiveContainer, 
  Cell 
} from 'recharts';
import { motion, AnimatePresence, Reorder } from 'motion/react';
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
  getDoc,
  setDoc,
  getDocFromServer,
  writeBatch
} from 'firebase/firestore';
import { 
  ref as storageRef, 
  uploadBytes, 
  getDownloadURL,
  deleteObject
} from 'firebase/storage';
import { db, auth, storage } from './lib/firebase';
import { Vendor, PriceItem, ComparisonRow, PendingPriceUpdate } from './types';
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
const Modal = ({ children, onClose, title, maxWidth = "max-w-2xl" }: { children: React.ReactNode; onClose: () => void; title: string; key?: string; maxWidth?: string }) => (
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
      className={`w-full ${maxWidth} overflow-hidden rounded-3xl bg-white shadow-2xl relative flex flex-col max-h-[90vh]`}
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
  const [vendorSearchTerm, setVendorSearchTerm] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [isVerified, setIsVerified] = useState<string | null>(null); // Tracks which vendor ID is verified
  const [loading, setLoading] = useState(true);
  const [viewMode, setViewMode] = useState<'detail' | 'matrix'>('detail');

  // Price table state
  const [priceItems, setPriceItems] = useState<PriceItem[]>([]);
  const [isAddingItem, setIsAddingItem] = useState(false);
  const [isViewingPriceTable, setIsViewingPriceTable] = useState(false);
  const [isEditingVendorInfo, setIsEditingVendorInfo] = useState(false);
  const [isChangingPassword, setIsChangingPassword] = useState(false);
  const [isViewingDeletedVendors, setIsViewingDeletedVendors] = useState(false);
  const [vendorSortMode, setVendorSortMode] = useState<'name' | 'manual'>('name');
  const [priceSortField, setPriceSortField] = useState<'itemName' | 'spec' | 'order'>('order');
  const [priceSortOrder, setPriceSortOrder] = useState<'asc' | 'desc'>('asc');
  const [errorInfo, setErrorInfo] = useState<FirestoreErrorInfo | null>(null);
  const [isSeedModalOpen, setIsSeedModalOpen] = useState(false);
  const [bulkFile, setBulkFile] = useState<File | null>(null);
  const [bulkItemFile, setBulkItemFile] = useState<File | null>(null);
  const [isBulkItemUploadOpen, setIsBulkItemUploadOpen] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [editingPriceId, setEditingPriceId] = useState<string | null>(null);
  const [priceTableFile, setPriceTableFile] = useState<File | null>(null);
  const [excelPreviewData, setExcelPreviewData] = useState<any[] | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [seedStatus, setSeedStatus] = useState<{ current: number; total: number; isDone: boolean } | null>(null);
  const [isAdminMode, setIsAdminMode] = useState(false);
  const [showVendorPasswords, setShowVendorPasswords] = useState<{[key: string]: boolean}>({});
  const [isDeleteModalOpen, setIsDeleteModalOpen] = useState(false);
  const [vendorToDelete, setVendorToDelete] = useState<{id: string, name: string} | null>(null);
  const [deletePasswordInput, setDeletePasswordInput] = useState('');
  const [addItemNegoType, setAddItemNegoType] = useState<'percent' | 'sts_pipe'>('percent');
  const [selectedPriceIds, setSelectedPriceIds] = useState<Set<string>>(new Set());
  const [isBulkAdjustModalOpen, setIsBulkAdjustModalOpen] = useState(false);
  const [bulkAdjustValue, setBulkAdjustValue] = useState<number>(0);
  const [bulkAdjustType, setBulkAdjustType] = useState<'percent'>('percent');
  const [isDeepLinkMode, setIsDeepLinkMode] = useState(false);
  const [isGuestMode, setIsGuestMode] = useState(false);
  const [newAdminPassword, setNewAdminPassword] = useState('');
  const [isChangingAdminPassword, setIsChangingAdminPassword] = useState(false);
  const [adminPassword, setAdminPassword] = useState('admin1234');
  const [adminPasswordInput, setAdminPasswordInput] = useState('');
  const [showAdminLogin, setShowAdminLogin] = useState(false);
  const [copySuccess, setCopySuccess] = useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // Fetch admin password from Firestore
  useEffect(() => {
    const fetchAdminSettings = async () => {
      try {
        const adminRef = doc(db, 'settings', 'admin');
        const adminSnap = await getDoc(adminRef);
        
        if (adminSnap.exists()) {
          setAdminPassword(adminSnap.data().password);
        } else {
          try {
            // Initialize if not exists
            await setDoc(adminRef, { password: 'admin1234' });
            setAdminPassword('admin1234');
          } catch (writeError) {
            handleFirestoreError(writeError, OperationType.WRITE, 'settings/admin');
          }
        }
      } catch (error) {
        console.error("Error fetching admin settings:", error);
        // Only throw if it's a permission error we want to debug
        if (error instanceof Error && error.message.includes('permissions')) {
           handleFirestoreError(error, OperationType.GET, 'settings/admin');
        }
      }
    };
    fetchAdminSettings();
  }, []);

  const handleUpdateAdminPassword = async () => {
    if (!newAdminPassword) return;
    try {
      setLoading(true);
      await updateDoc(doc(db, 'settings', 'admin'), {
        password: newAdminPassword
      });
      setAdminPassword(newAdminPassword);
      setIsChangingAdminPassword(false);
      setNewAdminPassword('');
      alert('관리자 비밀번호가 성공적으로 변경되었습니다.');
    } catch (error) {
      console.error("Error updating admin password:", error);
      alert('비밀번호 변경 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const copyGuestLink = () => {
    let baseOrigin = window.location.origin;
    if (baseOrigin.includes('ais-dev-')) {
      baseOrigin = baseOrigin.replace('ais-dev-', 'ais-pre-');
    }
    const url = new URL(baseOrigin);
    url.searchParams.set('g', '1');
    navigator.clipboard.writeText(url.toString());
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  const [adminNotification, setAdminNotification] = useState(false);
  const [isNotesExpanded, setIsNotesExpanded] = useState(true);
  const [isMatrixExpanded, setIsMatrixExpanded] = useState(false);
  const [allPrices, setAllPrices] = useState<PriceItem[]>([]);
  const [isMatrixLoading, setIsMatrixLoading] = useState(false);
  const [matrixCategory, setMatrixCategory] = useState<string>('전체');

  // Fetch all prices from all vendors for Matrix View
  useEffect(() => {
    if (viewMode !== 'matrix') return;

    const fetchAllPrices = async () => {
      setIsMatrixLoading(true);
      try {
        const pricesPromises = vendors.map(async (vendor) => {
          const q = query(collection(db, 'vendors', vendor.id, 'prices'), orderBy('itemName', 'asc'));
          const snapshot = await getDocs(q);
          return snapshot.docs.map(doc => ({ 
            id: doc.id, 
            ...doc.data(),
            vendorId: vendor.id,
            vendorName: vendor.name 
          } as any as PriceItem));
        });

        const allResults = await Promise.all(pricesPromises);
        setAllPrices(allResults.flat());
      } catch (error) {
        console.error("Error fetching all prices for matrix:", error);
      } finally {
        setIsMatrixLoading(false);
      }
    };

    fetchAllPrices();
  }, [viewMode, vendors.length]);

  const [matrixVendorIds, setMatrixVendorIds] = useState<string[]>([]);
  const [showOnlySelectedVendors, setShowOnlySelectedVendors] = useState(false);

  // Filter vendors for matrix based on selection, preserving click order
  const filteredVendors = useMemo(() => {
    if (matrixVendorIds.length > 0) {
      // Map based on matrixVendorIds order
      return matrixVendorIds
        .map(id => vendors.find(v => v.id === id))
        .filter((v): v is Vendor => !!v);
    }
    return vendors;
  }, [vendors, matrixVendorIds]);

  // Matrix Processing: Group items by Name + Spec
  const matrixData = useMemo(() => {
    const grouped: Record<string, { 
      itemName: string; 
      spec: string; 
      unit: string; 
      category: string; 
      prices: Record<string, { unitPrice: number; negoRate: number; negoType: 'percent' | 'sts_pipe'; hasPendingUpdate?: boolean }> 
    }> = {};
    
    // Filter allPrices based on current search and category
    const filtered = allPrices.filter(item => {
      const matchesSearch = searchTerm === '' || 
        item.itemName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (item.spec && item.spec.toLowerCase().includes(searchTerm.toLowerCase()));
      
      const matchesCategory = matrixCategory === '전체' || item.category === matrixCategory;
      
      // If we are filtering vendors, only show rows that have prices in at least one of the selected vendors
      // But we can keep all filtered rows for now, the columns will be filtered.
      // Actually, it's better to only show items that appear in at least one of the SELECTED vendors if filtering is active.
      const matchesVendorSelection = matrixVendorIds.length === 0 || matrixVendorIds.includes(item.vendorId);
      
      return matchesSearch && matchesCategory && matchesVendorSelection;
    });

    filtered.forEach(item => {
      const key = `${item.itemName}_${item.spec || ''}`;
      if (!grouped[key]) {
        grouped[key] = {
          itemName: item.itemName,
          spec: item.spec || '-',
          unit: item.unit || '-',
          category: item.category || '기타',
          prices: {}
        };
      }
      grouped[key].prices[item.vendorId] = {
        unitPrice: item.unitPrice,
        negoRate: item.negoRate,
        negoType: item.negoType || 'percent',
        hasPendingUpdate: pendingUpdates.some(u => u.priceItemId === item.id && u.status === 'pending')
      };
    });

    return Object.values(grouped).sort((a, b) => a.itemName.localeCompare(b.itemName));
  }, [allPrices, searchTerm, matrixCategory, matrixVendorIds]);

  const matrixCategories = useMemo(() => {
    return ['전체', ...Array.from(new Set(allPrices.map(p => p.category).filter(c => c && c !== '전체')))];
  }, [allPrices]);
  const [usdToKrw, setUsdToKrw] = useState<number>(1350);
  const [lastRateUpdate, setLastRateUpdate] = useState<string>("");
  const [savingMatrixId, setSavingMatrixId] = useState<string | null>(null);
  const [pendingUpdates, setPendingUpdates] = useState<PendingPriceUpdate[]>([]);
  const [isApprovalsModalOpen, setIsApprovalsModalOpen] = useState(false);

  const [activeCategory, setActiveCategory] = useState('전체');
  const [selectedMatrixRow, setSelectedMatrixRow] = useState<any>(null);

  const chartData = useMemo(() => {
    if (!selectedMatrixRow) return [];
    return vendors.map(v => ({
      name: v.name,
      price: selectedMatrixRow.prices[v.id] || 0
    })).filter(d => d.price > 0).sort((a, b) => a.price - b.price);
  }, [selectedMatrixRow, vendors]);

  const [categories, setCategories] = useState<string[]>(['밸브류', '피팅류', '파이프', 'STS파이프', '프랜지', '기타']);
  const [bulkNegoValue, setBulkNegoValue] = useState(5);
  const [columnWidths, setColumnWidths] = useState({
    itemCode: 100,
    category: 100,
    itemName: 250,
    maker: 120,
    spec: 150,
    unit: 80,
    costPrice: 120,
    negoRate: 100,
    weight: 90,
    discountAmount: 120,
    unitPrice: 120,
    change: 100,
    lastUpdated: 150
  });
  const [resizing, setResizing] = useState<keyof typeof columnWidths | null>(null);

  useEffect(() => {
    if (selectedVendor?.categories && selectedVendor.categories.length > 0) {
      // Ensure categories are unique and don't contain "전체" which is reserved for the default tab
      const uniqueCats = Array.from(new Set(selectedVendor.categories.filter(c => c && c.trim() !== '' && c.trim() !== '전체')));
      setCategories(uniqueCats);
    } else {
      setCategories(['밸브류', '피팅류', '파이프', 'STS파이프', '프랜지', '기타']);
    }
  }, [selectedVendor?.id, selectedVendor?.categories]);

  useEffect(() => {
    const q = query(collection(db, 'pending_updates'), orderBy('requestedAt', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const updates = snapshot.docs.map(doc => ({ 
        id: doc.id, 
        ...doc.data() 
      } as PendingPriceUpdate));
      setPendingUpdates(updates);
    });
    
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!isAdminMode && !isGuestMode && viewMode === 'matrix') {
      setViewMode('detail');
    }
  }, [isAdminMode, isGuestMode, viewMode]);

  const updateVendorCategories = async (newCategories: string[]) => {
    if (!selectedVendor) return;
    try {
      setLoading(true);
      const uniqueCats = Array.from(new Set(newCategories.filter(c => c && c.trim() !== '' && c.trim() !== '전체')));
      const docRef = doc(db, 'vendors', selectedVendor.id);
      const updateData = {
        categories: uniqueCats,
        updatedAt: serverTimestamp()
      };
      await updateDoc(docRef, updateData);
      
      setCategories(uniqueCats);
      setSelectedVendor(prev => prev ? { ...prev, categories: uniqueCats } : null);
      return true;
    } catch (error: any) {
      console.error("Error updating categories:", error);
      alert(`카테고리 업데이트에 실패했습니다. (${error.message || '권한 부족'})`);
      return false;
    } finally {
      setLoading(false);
    }
  };

  const startResize = (e: React.MouseEvent, column: keyof typeof columnWidths) => {
    e.preventDefault();
    setResizing(column);
  };

  useEffect(() => {
    if (!resizing) return;

    const handleMouseMove = (e: MouseEvent) => {
      setColumnWidths(prev => {
        const newWidths = { ...prev };
        // Use a simple movement calculation
        newWidths[resizing] = Math.max(50, prev[resizing] + e.movementX);
        return newWidths;
      });
    };

    const handleMouseUp = () => setResizing(null);

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizing]);

  const canManageItems = !isGuestMode && (isAdminMode || (selectedVendor && (isVerified === selectedVendor.id || isDeepLinkMode)));
  const isActuallyAuthorized = !isGuestMode && (isAdminMode || (selectedVendor && isVerified === selectedVendor.id));

  const toggleSelectItem = (id: string) => {
    setSelectedPriceIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (selectedPriceIds.size === priceItems.length && priceItems.length > 0) {
      setSelectedPriceIds(new Set());
    } else {
      setSelectedPriceIds(new Set(priceItems.map(item => item.id)));
    }
  };

  const handleBulkDelete = async () => {
    if (!selectedVendor || selectedPriceIds.size === 0) return;
    if (!confirm(`${selectedPriceIds.size}개의 항목을 삭제하시겠습니까?`)) return;

    try {
      setLoading(true);
      const ids = Array.from(selectedPriceIds);
      const chunks = [];
      for (let i = 0; i < ids.length; i += 400) {
        chunks.push(ids.slice(i, i + 400));
      }

      for (const chunk of chunks) {
        const batch = writeBatch(db);
        chunk.forEach(id => {
          batch.delete(doc(db, 'vendors', selectedVendor.id, 'prices', id));
        });
        await batch.commit();
      }
      
      setSelectedPriceIds(new Set());
      alert('선택한 항목들이 성공적으로 삭제되었습니다.');
    } catch (error) {
      console.error("Error bulk deleting items:", error);
      alert('삭제 중 오류가 발생했습니다. 권한을 확인해주세요.');
    } finally {
      setLoading(false);
    }
  };

  const handleBulkPriceAdjust = async () => {
    if (!selectedVendor || selectedPriceIds.size === 0) return;
    
    try {
      if (isDeepLinkMode && !isAdminMode) {
        const targets = Array.from(selectedPriceIds)
          .map(id => priceItems.find(p => p.id === id))
          .filter(p => p && !p.hasPendingUpdate) as PriceItem[];
          
        if (targets.length === 0) {
          alert('변경 가능한 품목이 없습니다. (이미 승인 대기 중이거나 대상 없음)');
          return;
        }

        if (!confirm(`${targets.length}개 품목의 네고율을 현재 대비 ${bulkAdjustValue}% 가감 조정하여 승인 요청하시겠습니까?`)) {
          return;
        }

        setLoading(true);
        const chunkSize = 400;
        let processedCount = 0;

        for (let i = 0; i < targets.length; i += chunkSize) {
          const chunk = targets.slice(i, i + chunkSize);
          const batch = writeBatch(db);

          for (const item of chunk) {
            let newNegoRate = item.negoRate || 0;
            newNegoRate = Number(newNegoRate) + Number(bulkAdjustValue);
            const newUnitPrice = calculatePrice(item.costPrice, newNegoRate, item.useRounding, item.negoType || 'percent', item.weight);

            const pendingRef = doc(collection(db, 'pending_updates'));
            batch.set(pendingRef, {
              vendorId: selectedVendor.id,
              priceItemId: item.id,
              itemName: item.itemName,
              spec: item.spec || '',
              oldData: { negoRate: item.negoRate, unitPrice: item.unitPrice },
              newData: { negoRate: newNegoRate, unitPrice: newUnitPrice },
              status: 'pending',
              requestedBy: '업체(링크)',
              requestedAt: serverTimestamp()
            });

            const itemRef = doc(db, 'vendors', selectedVendor.id, 'prices', item.id);
            batch.update(itemRef, {
              hasPendingUpdate: true
            });
            processedCount++;
          }
          await batch.commit();
        }
        alert(`${processedCount}개 품목에 대한 승인 요청이 전송되었습니다.`);
      } else {
        setLoading(true);
        const targets = Array.from(selectedPriceIds)
          .map(id => priceItems.find(p => p.id === id))
          .filter(Boolean) as PriceItem[];

        const chunkSize = 400;
        for (let i = 0; i < targets.length; i += chunkSize) {
          const chunk = targets.slice(i, i + chunkSize);
          const batch = writeBatch(db);
          
          chunk.forEach(item => {
            let newNegoRate = item.negoRate || 0;
            newNegoRate = Number(newNegoRate) + Number(bulkAdjustValue);
            const newUnitPrice = calculatePrice(item.costPrice, newNegoRate, item.useRounding, item.negoType || 'percent', item.weight);

            batch.update(doc(db, 'vendors', selectedVendor.id, 'prices', item.id), {
              negoRate: newNegoRate,
              unitPrice: newUnitPrice,
              updatedAt: serverTimestamp()
            });
          });
          await batch.commit();
        }
        alert('가격이 일괄 조정되었습니다.');
      }
      setSelectedPriceIds(new Set());
      setIsBulkAdjustModalOpen(false);
      setBulkAdjustValue(0);
    } catch (error) {
      console.error("Error adjusting prices:", error);
      alert('가격 조정 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const handleBulkCategoryUpdate = async (newCategory: string) => {
    if (!selectedVendor || selectedPriceIds.size === 0 || !newCategory) return;
    
    try {
      setLoading(true);
      const targets = Array.from(selectedPriceIds);
      const chunkSize = 400;
      for (let i = 0; i < targets.length; i += chunkSize) {
        const chunk = targets.slice(i, i + chunkSize);
        const batch = writeBatch(db);
        const vId = selectedVendor.id;
        chunk.forEach((id: string) => {
          batch.update(doc(db, 'vendors', vId, 'prices', id), {
            category: newCategory,
            updatedAt: serverTimestamp()
          });
        });
        await batch.commit();
      }
      alert(`선택한 ${targets.length}개 품목의 카테고리가 '${newCategory}'(으)로 변경되었습니다.`);
      setSelectedPriceIds(new Set());
    } catch (error) {
      console.error("Error bulk updating categories:", error);
      alert('카테고리 변경 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const handleInlinePriceUpdate = async (id: string, field: string, value: any) => {
    if (!selectedVendor) return;
    
    try {
      const item = priceItems.find(p => p.id === id);
      if (!item) return;

      const updates: any = {};

      if (field === 'costPrice') {
        const newCost = Number(value);
        const newUnitPrice = calculatePrice(newCost, item.negoRate, item.useRounding, item.negoType || 'percent', item.weight);
        updates.costPrice = newCost;
        updates.unitPrice = newUnitPrice;
      } else if (field === 'weight') {
        const newWeight = Number(value);
        const newUnitPrice = calculatePrice(item.costPrice, item.negoRate, item.useRounding, item.negoType || 'percent', newWeight);
        updates.weight = newWeight;
        updates.unitPrice = newUnitPrice;
      } else if (field === 'negoRate') {
        const newNegoRate = Number(value);
        const newUnitPrice = calculatePrice(item.costPrice, newNegoRate, item.useRounding, item.negoType || 'percent', item.weight);
        updates.negoRate = newNegoRate;
        updates.unitPrice = newUnitPrice;
      } else if (field === 'negoType') {
        const newType = value as 'percent' | 'sts_pipe';
        const newUnitPrice = calculatePrice(item.costPrice, item.negoRate, item.useRounding, newType, item.weight);
        updates.negoType = newType;
        updates.unitPrice = newUnitPrice;
      } else if (field === 'useRounding') {
        const newUseRounding = value as boolean;
        const newUnitPrice = calculatePrice(item.costPrice, item.negoRate, newUseRounding, item.negoType || 'percent', item.weight);
        updates.useRounding = newUseRounding;
        updates.unitPrice = newUnitPrice;
      } else {
        updates[field] = value;
      }

      if (isDeepLinkMode && !isAdminMode) {
        if (item.hasPendingUpdate) {
          alert('이미 승인 대기 중인 변경 요청이 있습니다. 기존 요청이 처리된 후 다시 시도해 주세요.');
          return;
        }

        const changesString = Object.entries(updates)
          .map(([key, val]) => {
            const oldVal = item[key as keyof PriceItem];
            return `${key}: ${oldVal?.toLocaleString()} -> ${val?.toLocaleString()}`;
          })
          .join('\n');

        if (!confirm(`변경 사항을 요청하시겠습니까?\n\n[변경 내용]\n${changesString}`)) {
          return;
        }

        // Create pending update
        await addDoc(collection(db, 'pending_updates'), {
          vendorId: selectedVendor.id,
          priceItemId: id,
          itemName: item.itemName,
          spec: item.spec || '',
          oldData: Object.keys(updates).reduce((acc: any, key) => {
            acc[key] = item[key as keyof PriceItem];
            return acc;
          }, {}),
          newData: updates,
          status: 'pending',
          requestedBy: '업체(링크)',
          requestedAt: serverTimestamp()
        });
        
        // Mark item as having pending update
        await updateDoc(doc(db, 'vendors', selectedVendor.id, 'prices', id), {
          hasPendingUpdate: true
        });
        
        alert('단가 변경 요청이 전송되었습니다. 관리자 승인 후 반영됩니다.');
      } else {
        await updateDoc(doc(db, 'vendors', selectedVendor.id, 'prices', id), {
          ...updates,
          updatedAt: serverTimestamp()
        });
      }
    } catch (error) {
      console.error("Error updating price inline:", error);
      alert('변경 사항 반영 중 오류가 발생했습니다. 권한을 확인해 주세요.');
    }
  };

  const handleApproveUpdate = async (update: PendingPriceUpdate) => {
    try {
      const batch = writeBatch(db);
      
      // 1. Apply to price item
      batch.update(doc(db, 'vendors', update.vendorId, 'prices', update.priceItemId), {
        ...update.newData,
        updatedAt: serverTimestamp(),
        hasPendingUpdate: false
      });
      
      // 2. Mark update as approved
      batch.update(doc(db, 'pending_updates', update.id), {
        status: 'approved',
        approvedAt: serverTimestamp()
      });
      
      await batch.commit();
    } catch (error) {
      console.error("Approval error:", error);
      alert("승인 처리 중 오류가 발생했습니다.");
    }
  };

  const handleBulkApproveUpdates = async () => {
    const pendings = pendingUpdates.filter(u => u.status === 'pending');
    if (pendings.length === 0) return;
    if (!confirm(`${pendings.length}개의 모든 요청을 일괄 승인하시겠습니까?`)) return;

    try {
      setLoading(true);
      
      const chunks = [];
      for (let i = 0; i < pendings.length; i += 400) {
        chunks.push(pendings.slice(i, i + 400));
      }

      for (const chunk of chunks) {
        const batch = writeBatch(db);
        chunk.forEach(update => {
          // 1. Apply to price item
          batch.update(doc(db, 'vendors', update.vendorId, 'prices', update.priceItemId), {
            ...update.newData,
            updatedAt: serverTimestamp(),
            hasPendingUpdate: false
          });
          
          // 2. Mark update as approved
          batch.update(doc(db, 'pending_updates', update.id), {
            status: 'approved',
            approvedAt: serverTimestamp()
          });
        });
        await batch.commit();
      }
      
      alert('모든 요청이 일괄 승인되었습니다.');
    } catch (error) {
      console.error("Bulk approval error:", error);
      alert("일괄 승인 중 오류가 발생했습니다.");
    } finally {
      setLoading(false);
    }
  };

  const handleRejectUpdate = async (updateId: string) => {
    try {
      const update = pendingUpdates.find(u => u.id === updateId);
      if (!update) return;

      const batch = writeBatch(db);
      
      // 1. Mark as rejected
      batch.update(doc(db, 'pending_updates', updateId), {
        status: 'rejected',
        approvedAt: serverTimestamp()
      });

      // 2. Clear flag on price item
      batch.update(doc(db, 'vendors', update.vendorId, 'prices', update.priceItemId), {
        hasPendingUpdate: false
      });

      await batch.commit();
    } catch (error) {
      console.error("Rejection error:", error);
      alert("거절 처리 중 오류가 발생했습니다.");
    }
  };

  const handleBackupAllData = async () => {
    try {
      setLoading(true);
      const allVendorsSnapshot = await getDocs(collection(db, 'vendors'));
      const workbook = XLSX.utils.book_new();

      for (const vendorDoc of allVendorsSnapshot.docs) {
        const vendorData = vendorDoc.data();
        const pricesSnapshot = await getDocs(collection(db, 'vendors', vendorDoc.id, 'prices'));
        
        const prices = pricesSnapshot.docs.map(d => {
          const data = d.data();
          return {
            '품번': data.itemCode || '',
            '카테고리': data.category || '',
            '품명': data.itemName || '',
            '규격': data.spec || '',
            '단위': data.unit || '',
            '협가': data.costPrice || 0,
            '네고율': data.negoRate || 0,
            '구매단가': data.unitPrice || 0,
            '메이커': data.maker || '',
            '최근변경일': data.updatedAt?.toDate ? data.updatedAt.toDate().toLocaleString('ko-KR') : '-'
          };
        });

        if (prices.length > 0) {
          const worksheet = XLSX.utils.json_to_sheet(prices);
          XLSX.utils.book_append_sheet(workbook, worksheet, vendorData.name.substring(0, 31).replace(/[\\?*\/\[\]]/g, ''));
        }
      }

      XLSX.writeFile(workbook, `단가표_전체백업_${new Date().toLocaleDateString('ko-KR')}.xlsx`);
      alert('전체 데이터 백업이 완료되었습니다.');
    } catch (error) {
      console.error("Backup error:", error);
      alert('백업 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const handleBackupCurrentVendor = async () => {
    if (!selectedVendor || priceItems.length === 0) return;
    
    try {
      setLoading(true);
      const workbook = XLSX.utils.book_new();
      const exportData = priceItems.map(item => ({
        '품번': item.itemCode || '',
        '카테고리': item.category || '',
        '품명': item.itemName || '',
        '규격': item.spec || '',
        '단위': item.unit || '',
        '협가': item.costPrice || 0,
        '네고율': item.negoRate || 0,
        '구매단가': item.unitPrice || 0,
        '메이커': item.maker || '',
        '최근변경일': item.updatedAt?.toDate ? item.updatedAt.toDate().toLocaleString('ko-KR') : '-'
      }));

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      XLSX.utils.book_append_sheet(workbook, worksheet, '단가표');
      XLSX.writeFile(workbook, `${selectedVendor.name}_단가표_백업_${new Date().toLocaleDateString('ko-KR')}.xlsx`);
    } catch (error) {
      console.error("Vendor backup error:", error);
      alert('백업 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const calculatePrice = (cost: number, rate: number, itemRoundingOverride?: boolean, type: 'percent' | 'sts_pipe' = 'percent', weight?: number) => {
    let price = 0;
    
    if (type === 'sts_pipe') {
      // STS PIPE: 단중(costPrice) * KG단가(rate)
      price = cost * rate;
    } else {
      // Default: Percentage discount
      price = cost * (1 - (rate / 100));
    }

    const method = selectedVendor?.roundingMethod || 'none';
    
    // Use item override if provided, otherwise fallback to vendor setting
    const shouldRound = itemRoundingOverride !== undefined ? itemRoundingOverride : (method !== 'none');
    
    if (shouldRound) {
      price = Math.round(price / 10) * 10;
    }
    
    return price;
  };

  const getRoundedValue = (val: number, itemRoundingOverride?: boolean) => {
    const method = selectedVendor?.roundingMethod || 'none';
    const shouldRound = itemRoundingOverride !== undefined ? itemRoundingOverride : (method !== 'none');
    
    if (shouldRound) {
      return Math.round(val / 10) * 10;
    }
    return Math.round(val);
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
    const urlGuestMode = params.get('g');

    if (urlGuestMode === '1') {
      setIsGuestMode(true);
    } else {
      setIsGuestMode(false);
    }

    if (urlVendorId) {
      const vendor = vendors.find(v => v.id === urlVendorId);
      if (vendor) {
        setSelectedVendor(vendor);
        setIsDeepLinkMode(true);
        setViewMode('detail'); // Ensure we are in detail view for deep links
        
        // Clear param after selection to avoid sticky state if desired, 
        // but keeping it is better for refreshing the page.
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
    let baseOrigin = window.location.origin;
    
    // AI Studio environment: transform 'ais-dev-' to 'ais-pre-' to generate a publicly accessible link
    if (baseOrigin.includes('ais-dev-')) {
      baseOrigin = baseOrigin.replace('ais-dev-', 'ais-pre-');
    }
    
    const url = new URL(baseOrigin);
    url.searchParams.set('v', vendorId);
    navigator.clipboard.writeText(url.toString());
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  const handleAdminLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (adminPasswordInput === adminPassword) { 
      setIsAdminMode(true);
      setShowAdminLogin(false);
      setAdminPasswordInput('');
      setAdminNotification(true);
      setTimeout(() => setAdminNotification(false), 5000); // Auto-dismiss after 5s
    } else {
      alert('관리자 비밀번호가 틀렸습니다.');
    }
  };

  const deleteVendor = (id: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    
    if (!isAdminMode) {
      setShowAdminLogin(true);
      return;
    }
    
    const vendor = vendors.find(v => v.id === id);
    if (!vendor) return;

    setVendorToDelete({ id: vendor.id, name: vendor.name });
    setIsDeleteModalOpen(true);
  };

  const handleConfirmDelete = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!vendorToDelete) return;

    // Verify admin password for deletion
    if (deletePasswordInput !== 'admin1234') {
      alert('비밀번호가 일치하지 않습니다.');
      return;
    }

    try {
      setLoading(true);
      const id = vendorToDelete.id;
      
      // Soft delete: Mark as deleted instead of removing from DB
      await updateDoc(doc(db, 'vendors', id), {
        deleted: true,
        deletedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      
      if (selectedVendor?.id === id) {
        setSelectedVendor(null);
        setIsVerified(null);
        setViewMode('detail');
      }
      
      alert('업체가 성공적으로 삭제되었습니다.');
      setIsDeleteModalOpen(false);
      setVendorToDelete(null);
      setDeletePasswordInput('');
    } catch (error) {
      console.error("Delete Error:", error);
      alert('삭제 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const handleRestoreVendor = async (id: string) => {
    try {
      setLoading(true);
      await updateDoc(doc(db, 'vendors', id), {
        deleted: false,
        deletedAt: null,
        updatedAt: serverTimestamp()
      });
      alert('업체가 성공적으로 복원되었습니다.');
    } catch (error) {
      console.error("Restore Error:", error);
      alert('복원 중 오류가 발생했습니다.');
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
            order: count,
            deleted: false,
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
        '카테고리': '밸브',
        '품명': 'KS 10K 플랜지형 볼밸브',
        '규격': '50A',
        '품번': 'VB-KS10K-50',
        '단위': 'EA',
        '네고치': 5,
        '네고방식': '%',
        '협가': 45000,
        '제조사': 'A-Maker',
        '비고': '신규 모델'
      },
      {
        '카테고리': '피팅',
        '품명': '90도 엘보',
        '규격': '100A SCH40',
        '품번': 'FT-EL90-100',
        '단위': 'EA',
        '네고치': 0,
        '네고방식': '%',
        '협가': 12000,
        '제조사': 'B-Maker',
        '비고': ''
      },
      {
        '카테고리': 'STS파이프',
        '품명': 'STS304 PIPE 20A',
        '규격': '20A SCH10',
        '품번': 'PP-STS-20A',
        '단위': 'M',
        '네고치': 4500,
        '네고방식': 'STS',
        '단중': 1.62,
        '협가': 1.62,
        '제조사': 'MS-Metal',
        '비고': 'STS PIPE (협가=단중, 네고치=KG단가)'
      }
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "단가표_양식");
    XLSX.writeFile(wb, `${selectedVendor?.name || '업체'}_단가표_양식.xlsx`);
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

        if (data.length === 0) {
          alert("파일에 데이터가 없습니다.");
          setLoading(false);
          return;
        }

        // Fetch existing items to identify duplicates for update
        const existingItemsSnap = await getDocs(collection(db, "vendors", selectedVendor.id, "prices"));
        const existingItemsMap = new Map();
        existingItemsSnap.docs.forEach(doc => {
          const d = doc.data();
          const key = `${d.itemName}_${d.spec || ''}_${d.category || ''}`;
          existingItemsMap.set(key, doc.id);
        });

        let createdCount = 0;
        let updatedCount = 0;
        
        // Chunk processing for batches
        const chunks = [];
        for (let i = 0; i < data.length; i += 400) {
          chunks.push(data.slice(i, i + 400));
        }

        for (const chunk of chunks) {
          const batch = writeBatch(db);
          chunk.forEach((row, index) => {
            const costPrice = Number(row['협가'] || row['구매단가(입력)'] || row['현단가'] || row['단가'] || row['Price'] || 0);
            const negoRate = Number(row['네고율'] || row['네고율(%)'] || row['네고치'] || row['Discount'] || 0);

            let negoType: 'percent' | 'sts_pipe' = 'percent';
            const rawType = String(row['네고방식'] || row['Nego Type'] || '').trim();
            if (rawType === 'STS' || rawType === 'sts_pipe') {
              negoType = 'sts_pipe';
            }

            const weight = Number(row['단중'] || row['Weight'] || 0);

            const itemRounding = row['반올림여부'] !== undefined ? 
                               (row['반올림여부'] === 'Y' || row['반올림여부'] === true) : 
                               (selectedVendor.roundingMethod !== 'none');
            
            const unitPrice = calculatePrice(costPrice, negoRate, itemRounding, negoType, weight);

            const itemName = String(row['품목명'] || row['품명'] || row['Name'] || '').trim();
            const spec = String(row['규격'] || row['Spec'] || '').trim();
            const category = String(row['카테고리'] || row['분류'] || row['Category'] || '').trim();
            
            if (!itemName) return; // Skip empty rows

            const itemData: any = {
              vendorId: selectedVendor.id,
              itemName,
              itemCode: String(row['품번'] || row['코드'] || row['Code'] || ''),
              category,
              spec,
              unit: String(row['단위'] || row['Unit'] || 'EA'),
              costPrice,
              negoRate,
              negoType,
              weight,
              unitPrice,
              remarks: String(row['비고'] || row['Remarks'] || ''),
              maker: String(row['메이커'] || row['제조사'] || row['Maker'] || ''),
              useRounding: itemRounding,
              updatedAt: serverTimestamp()
            };

            const key = `${itemName}_${spec}_${category}`;
            const existingId = existingItemsMap.get(key);

            if (existingId) {
              const itemRef = doc(db, "vendors", selectedVendor.id, "prices", existingId);
              batch.update(itemRef, itemData);
              updatedCount++;
            } else {
              const itemRef = doc(collection(db, "vendors", selectedVendor.id, "prices"));
              batch.set(itemRef, { ...itemData, order: existingItemsMap.size + createdCount, createdAt: serverTimestamp() });
              createdCount++;
            }
          });
          await batch.commit();
        }

        alert(`처리 완료: ${createdCount}개 신규 등록, ${updatedCount}개 정보 업데이트`);
        setBulkItemFile(null);
      } catch (error) {
        console.error("Bulk item upload error:", error);
        alert("업로드 중 오류가 발생했습니다. 파일 형식을 확인해주세요.");
      } finally {
        setLoading(false);
      }
    };
    reader.readAsBinaryString(bulkItemFile);
  };

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
      const q = query(collection(db, 'vendors', selectedVendor.id, 'prices'));
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const data = snapshot.docs.map(doc => ({ 
          id: doc.id, 
          vendorId: selectedVendor.id,
          ...doc.data() 
        } as PriceItem));
        setPriceItems(data);
      }, (error) => {
        handleFirestoreError(error, OperationType.LIST, `vendors/${selectedVendor.id}/prices`);
      });
      return unsubscribe;
    } else {
      setPriceItems([]);
    }
  }, [selectedVendor]);
  
  // Excel preview fetch
  useEffect(() => {
    if (selectedVendor?.priceTableUrl && selectedVendor?.priceTableFileType === 'excel' && (isVerified === selectedVendor.id || isAdminMode)) {
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
    const activeVendors = vendors.filter(v => !v.deleted);
    
    const cleanName = (name: string) => {
      // Remove (주), 주식회사, ㈜ prefixes and suffixes for cleaner sorting
      return name.replace(/^(\(주\)|주식회사|㈜|주\))/, '')
                 .replace(/(\(주\)|주식회사|㈜|주\))$/, '')
                 .trim();
    };

    if (vendorSortMode === 'name') {
      return [...activeVendors].sort((a, b) => {
        const nameA = cleanName(a.name);
        const nameB = cleanName(b.name);
        return nameA.localeCompare(nameB, 'ko');
      });
    }

    // Default: Sort by order field
    return [...activeVendors].sort((a, b) => {
      const orderA = a.order ?? 0;
      const orderB = b.order ?? 0;
      if (orderA !== orderB) return orderA - orderB;
      
      const timeA = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
      const timeB = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
      return timeB - timeA;
    });
  }, [vendors, vendorSortMode]);

  const sortedPriceItems = useMemo(() => {
    let items = priceItems.map(item => ({
      ...item,
      hasPendingUpdate: item.hasPendingUpdate || pendingUpdates.some(u => u.priceItemId === item.id && u.status === 'pending')
    }));
    
    // Sort by selected field
    items.sort((a, b) => {
      let result = 0;
      if (priceSortField === 'itemName') {
        result = a.itemName.localeCompare(b.itemName, 'ko');
      } else if (priceSortField === 'spec') {
        result = (a.spec || '').localeCompare(b.spec || '', undefined, { numeric: true, sensitivity: 'base' });
      } else {
        // Default order sorting
        const orderA = a.order ?? 0;
        const orderB = b.order ?? 0;
        result = orderA - orderB;
      }
      
      // Secondary sort for stability
      if (result === 0) {
        if (priceSortField !== 'itemName') result = a.itemName.localeCompare(b.itemName, 'ko');
        if (result === 0 && priceSortField !== 'spec') result = (a.spec || '').localeCompare(b.spec || '', undefined, { numeric: true, sensitivity: 'base' });
      }
      
      return priceSortOrder === 'asc' ? result : -result;
    });
    
    return items;
  }, [priceItems, priceSortField, priceSortOrder, pendingUpdates]);

  const handlePriceSort = (field: 'itemName' | 'spec' | 'order') => {
    if (priceSortField === field) {
      setPriceSortOrder(priceSortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setPriceSortField(field);
      setPriceSortOrder('asc');
    }
  };

  const sortedVendors = useMemo(() => {
    return baseSortedVendors
      .filter(v => 
        v.name.toLowerCase().includes(vendorSearchTerm.toLowerCase()) ||
        v.representative?.toLowerCase().includes(vendorSearchTerm.toLowerCase())
      );
  }, [baseSortedVendors, vendorSearchTerm]);

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
      roundingMethod: formData.get('useRounding') === 'on' ? 'round' : 'none',
      masterCustomFlag: formData.get('masterCustomFlag') === 'on',
      categories: ['밸브류', '피팅류', '파이프', 'STS파이프', '프랜지', '기타'],
      priceTableUrl,
      priceTableFileType,
      priceTableFileName,
      order: vendors.length,
      deleted: false,
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
      roundingMethod: formData.get('useRounding') === 'on' ? 'round' : 'none',
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
    const currentPasswordInput = formData.get('currentPassword') as string;
    const newPassword = formData.get('newPassword') as string;

    if (currentPasswordInput !== selectedVendor.password) {
      alert('기존 비밀번호가 일치하지 않습니다.');
      return;
    }

    if (!newPassword || newPassword.length < 4) {
      alert('새 비밀번호는 최소 4자 이상이어야 합니다.');
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
    if (isAdminMode || (selectedVendor && passwordInput === selectedVendor.password)) {
      setIsVerified(selectedVendor?.id || null);
    } else {
      alert('비밀번호가 일치하지 않습니다.');
    }
  };

  const handleAddItem = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedVendor) return;
    const formData = new FormData(e.currentTarget);
    const costPrice = Number(String(formData.get('costPrice') || '0').replace(/,/g, ''));
    const negoRate = Number(String(formData.get('negoRate') || '0').replace(/,/g, ''));
    const negoType = (formData.get('negoType') as 'percent' | 'sts_pipe') || 'percent';
    const weight = Number(String(formData.get('weight') || '0').replace(/,/g, ''));
    const useRounding = formData.get('itemRounding') === 'on';
    const unitPrice = calculatePrice(costPrice, negoRate, useRounding, negoType, weight);

    const newItem = {
      vendorId: selectedVendor.id,
      itemName: formData.get('itemName') as string,
      itemCode: (formData.get('itemCode') as string) || '',
      category: (formData.get('category') as string) || '',
      spec: formData.get('spec') as string,
      unit: formData.get('unit') as string,
      costPrice,
      negoRate,
      negoType,
      unitPrice,
      baseUnitPrice: unitPrice,
      weight,
      useRounding,
      remarks: formData.get('remarks') as string,
      maker: formData.get('maker') as string,
      order: priceItems.length + 1,
      updatedAt: serverTimestamp()
    };

    try {
      await addDoc(collection(db, 'vendors', selectedVendor.id, 'prices'), newItem);
      setIsAddingItem(false);
    } catch (error) {
      console.error("Error adding price item:", error);
      alert('품목 추가 중 오류가 발생했습니다.');
    }
  };

  const deletePriceItem = async (itemId: string) => {
    if (!selectedVendor) return;
    if (window.confirm('항목을 삭제하시겠습니까?')) {
      try {
        setLoading(true);
        await deleteDoc(doc(db, 'vendors', selectedVendor.id, 'prices', itemId));
      } catch (error) {
        console.error("Error deleting price item:", error);
        alert('삭제 중 오류가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    }
  };

  const handleUpdatePriceItem = async (itemId: string, updates: Partial<PriceItem>) => {
    if (!selectedVendor) return;
    
    // Recalculate unit price if cost, rate, type, or weight changed
    let finalUnitPrice = updates.unitPrice;
    if (updates.costPrice !== undefined || updates.negoRate !== undefined || updates.negoType !== undefined || updates.weight !== undefined) {
      const item = priceItems.find(i => i.id === itemId);
      if (item) {
        const cost = updates.costPrice ?? item.costPrice;
        const rate = updates.negoRate ?? item.negoRate;
        const type = updates.negoType ?? item.negoType ?? 'percent';
        const useRounding = updates.useRounding ?? item.useRounding;
        const weight = updates.weight ?? item.weight;
        finalUnitPrice = calculatePrice(cost, rate, useRounding, type, weight);
      }
    }

    try {
      if (isDeepLinkMode && !isAdminMode) {
        const item = priceItems.find(i => i.id === itemId);
        if (item?.hasPendingUpdate) {
          alert('이미 승인 대기 중인 변경 요청이 있습니다. 기존 요청이 처리된 후 다시 시도해 주세요.');
          return;
        }

        await addDoc(collection(db, 'pending_updates'), {
          vendorId: selectedVendor.id,
          priceItemId: itemId,
          itemName: item?.itemName || 'Unknown Item',
          spec: item?.spec || '',
          oldData: item ? { costPrice: item.costPrice, negoRate: item.negoRate, negoType: item.negoType, weight: item.weight } : {},
          newData: {
            ...updates,
            ...(finalUnitPrice !== undefined ? { unitPrice: finalUnitPrice } : {})
          },
          status: 'pending',
          requestedBy: '업체(링크)',
          requestedAt: serverTimestamp()
        });

        await updateDoc(doc(db, 'vendors', selectedVendor.id, 'prices', itemId), {
          hasPendingUpdate: true
        });

        alert('상세 수정 요청이 전송되었습니다. 관리자 승인 후 반영됩니다.');
        setEditingPriceId(null);
      } else {
        await updateDoc(doc(db, 'vendors', selectedVendor.id, 'prices', itemId), {
          ...updates,
          ...(finalUnitPrice !== undefined ? { unitPrice: finalUnitPrice } : {}),
          updatedAt: serverTimestamp()
        });
        setEditingPriceId(null);
      }
    } catch (error) {
      console.error("Error updating price item:", error);
      alert("수정 중 오류가 발생했습니다.");
    }
  };

  const exportToExcel = () => {
    if (!selectedVendor || priceItems.length === 0) {
      alert('내보낼 데이터가 없습니다.');
      return;
    }

    const data = priceItems.map((item, index) => ({
      '번호': index + 1,
      '품번': item.itemCode || '',
      '품명': item.itemName,
      '규격': item.spec || '',
      '단위': item.unit || '',
      '네고치': item.negoRate || 0,
      '네고방식': item.negoType === 'sts_pipe' ? 'STS' : '%',
      '단중': item.weight || 0,
      '구매단가': item.unitPrice || 0,
      '협가': item.costPrice || 0,
      '제조사': item.maker || '',
      '비고': item.remarks || '',
      '최근변경일': item.updatedAt?.toDate ? item.updatedAt.toDate().toLocaleString('ko-KR') : '-'
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
          const costPrice = Number(row['협가'] || row['구매단가(입력)'] || row['현단가'] || row['단가'] || row['Negotiated Price'] || 0);
          const negoRate = Number(row['네고율'] || row['네고율(%)'] || row['Nego Rate'] || 0);
          
          let negoType: 'percent' | 'sts_pipe' = 'percent';
          const rawType = String(row['네고방식'] || row['Nego Type'] || '').trim().toUpperCase();
          if (rawType === 'STS') {
            negoType = 'sts_pipe';
          }

          const weight = Number(row['단중'] || row['Weight'] || 0);
          const unitPrice = calculatePrice(costPrice, negoRate, undefined, negoType, weight);

          const newItem = {
            vendorId: selectedVendor.id,
            itemCode: row['품번'] || row['품목코드'] || row['품번'] || row['Item Code'] || '',
            itemName: row['품목명'] || row['품목'] || '',
            spec: row['규격'] || '',
            unit: row['단위'] || '',
            maker: row['메이커'] || row['제조사'] || '',
            costPrice,
            negoRate,
            negoType,
            weight,
            unitPrice,
            baseUnitPrice: unitPrice,
            remarks: row['비고'] || '',
            order: priceItems.length + importCount + 1,
            updatedAt: serverTimestamp()
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
      <div className="flex h-screen w-full flex-col items-center justify-center bg-white gap-4">
        <Loader2 className="h-10 w-10 animate-spin text-indigo-500" />
        <p className="text-slate-500 font-medium animate-pulse">시스템 초기화 중...</p>
      </div>
    );
  }

  return (
    <div className="flex h-screen w-full flex-col bg-[#F8F9FA] text-[#333] font-sans selection:bg-indigo-100 selection:text-indigo-900 overflow-hidden">
      {/* 1. TOP HEADER */}
      <header className="h-14 bg-white border-b border-slate-200 flex items-center justify-between px-4 shrink-0 shadow-sm z-30">
        <div className="flex items-center gap-8">
          <div className="flex items-center gap-2">
            <span className="text-lg font-bold text-slate-800 tracking-tighter">(주)명신기공 <span className="font-medium text-slate-500">단가관리</span></span>
          </div>
          <nav className="flex items-center gap-6">
            {!isDeepLinkMode && !isGuestMode ? (
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => isAdminMode ? setIsAdminMode(false) : setShowAdminLogin(true)}
                  className={`flex items-center gap-1.5 text-sm font-bold transition-colors ${isAdminMode ? 'text-indigo-600' : 'text-slate-600 hover:text-indigo-600'}`}
                >
                  <User className="h-4 w-4" />
                  {isAdminMode ? '관리모드 해제' : '관리자'}
                </button>
                {isAdminMode && (
                  <div className="flex items-center gap-1 ml-2 border-l border-slate-200 pl-3">
                    <button 
                      onClick={() => setIsChangingAdminPassword(true)}
                      className="text-[11px] font-bold text-slate-400 hover:text-slate-600 bg-slate-100 px-2 py-1 rounded"
                    >
                      암호변경
                    </button>
                    <button 
                      onClick={copyGuestLink}
                      className={`flex items-center gap-1 text-[11px] font-bold px-2 py-1 rounded transition-all ${copySuccess ? 'bg-emerald-500 text-white' : 'bg-indigo-50 text-indigo-600 hover:bg-indigo-100'}`}
                    >
                      <LinkIcon className="h-3 w-3" />
                      {copySuccess ? '복사됨!' : '게스트링크'}
                    </button>
                  </div>
                )}
              </div>
            ) : isGuestMode ? (
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-1.5 bg-emerald-50 text-emerald-600 px-3 py-1 rounded-full border border-emerald-100 font-bold">
                  <Users className="h-3.5 w-3.5" />
                  <span className="text-xs">게스트 모드</span>
                </div>
              </div>
            ) : selectedVendor && (
              <div className="flex items-center gap-2">
                <div className="h-4 w-[1px] bg-slate-200 mr-2" />
                <span className="text-sm font-bold text-slate-400">전용 포탈</span>
                <span className="text-sm font-black text-indigo-600">{selectedVendor.name}</span>
              </div>
            )}

            {!isDeepLinkMode && (
              <div className="flex items-center gap-6 border-l border-slate-100 pl-6 h-6">
                <button 
                  onClick={() => {
                    setViewMode('detail');
                    setSelectedVendor(null);
                    const url = new URL(window.location.href);
                    url.searchParams.delete('v');
                    window.history.replaceState({}, '', url);
                  }}
                  className={`text-sm font-medium transition-colors ${viewMode === 'detail' ? 'text-indigo-600 font-bold underline underline-offset-8' : 'text-slate-600 hover:text-indigo-600'}`}
                >
                  거래처별 단가표
                </button>
                {(isAdminMode || isGuestMode) && (
                  <button 
                    onClick={() => {
                      setViewMode('matrix');
                      setSelectedVendor(null);
                    }}
                    className={`text-sm font-medium transition-colors ${viewMode === 'matrix' ? 'text-indigo-600 font-bold underline underline-offset-8' : 'text-slate-600 hover:text-indigo-600'}`}
                  >
                    단가 통합 비교
                  </button>
                )}
              </div>
            )}

            {isGuestMode && (
              <button 
                onClick={() => {
                   const url = new URL(window.location.href);
                   url.searchParams.delete('g');
                   url.searchParams.delete('v');
                   window.location.href = url.toString();
                }}
                className="text-xs font-bold text-slate-400 hover:text-slate-600 underline underline-offset-2 ml-auto"
              >
                일반 모드로 전환
              </button>
            )}
          </nav>
        </div>
        <div className="flex items-center gap-4">
          {isAdminMode && pendingUpdates.filter(u => u.status === 'pending').length > 0 && (
            <button 
              onClick={() => setIsApprovalsModalOpen(true)}
              className="relative flex items-center gap-2 px-3 py-1.5 bg-amber-50 text-amber-700 rounded-full border border-amber-200 hover:bg-amber-100 transition-all shadow-sm group"
            >
              <Info className="h-4 w-4" />
              <span className="text-xs font-bold">승인 대기 {pendingUpdates.filter(u => u.status === 'pending').length}</span>
              <span className="absolute -top-1 -right-1 flex h-2 w-2">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2 w-2 bg-amber-500"></span>
              </span>
            </button>
          )}
          <div className="relative group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400 group-focus-within:text-indigo-500 transition-colors" />
            <input 
              type="text" 
              placeholder="품번·품명 검색..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-64 bg-slate-100 border-none rounded-full pl-10 pr-4 py-2 text-sm focus:ring-2 focus:ring-indigo-500/20 focus:bg-white transition-all outline-none"
            />
          </div>
        </div>
      </header>

      <div className="flex flex-1 overflow-hidden">
        {/* 2. SIDEBAR */}
        {!isDeepLinkMode && (
          <aside className="w-64 bg-white border-r border-slate-200 flex flex-col shrink-0 overflow-hidden">
            <div className="p-4 border-b border-slate-50 bg-[#FBFBFC]" id="sidebar-header">
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  거래처
                  <span className="text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded-full font-medium">
                    {sortedVendors.length}
                  </span>
                </h2>
                <div className="flex items-center gap-1">
                  {isAdminMode && (
                    <div className="flex items-center gap-1 mr-1 border-r border-slate-200 pr-1">
                      <button 
                        onClick={() => setIsViewingDeletedVendors(true)}
                        className="p-1.5 hover:bg-slate-200 rounded-md transition-colors text-slate-400 hover:text-amber-600"
                        title="삭제된 업체 복구"
                      >
                        <History className="h-4 w-4" />
                      </button>
                      <button 
                        onClick={handleBackupAllData}
                        className="p-1.5 hover:bg-slate-200 rounded-md transition-colors text-slate-400 hover:text-indigo-600"
                        title="전체 데이터 엑셀 백업"
                      >
                        <Download className="h-4 w-4" />
                      </button>
                      <button 
                        onClick={() => setIsAddingVendor(true)}
                        className="p-1.5 hover:bg-slate-200 rounded-md transition-colors text-slate-600"
                        id="add-vendor-btn"
                      >
                        <Plus className="h-4 w-4" />
                      </button>
                    </div>
                  )}
                  <select 
                    value={vendorSortMode}
                    onChange={(e) => setVendorSortMode(e.target.value as 'name' | 'manual')}
                    className="text-[10px] font-bold bg-white border border-slate-200 rounded px-1 py-0.5 outline-none focus:border-indigo-500"
                  >
                    <option value="name">가나다순</option>
                    <option value="manual">사용자 지정</option>
                  </select>
                </div>
              </div>
              <div className="relative group">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-400 group-focus-within:text-indigo-500 transition-colors" />
                <input 
                  type="text" 
                  placeholder="거래처 검색..."
                  value={vendorSearchTerm}
                  onChange={(e) => setVendorSearchTerm(e.target.value)}
                  className="w-full bg-white border border-slate-200 rounded-lg pl-9 pr-3 py-1.5 text-xs focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all outline-none"
                />
                {vendorSearchTerm && (
                  <button 
                    onClick={() => setVendorSearchTerm('')}
                    className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 hover:bg-slate-100 rounded-full"
                  >
                    <X className="h-3 w-3 text-slate-400" />
                  </button>
                )}
              </div>
            </div>
            <div className="flex-1 overflow-y-auto px-2 py-4 space-y-0.5" id="vendor-list-container">
              {vendorSortMode === 'manual' && !vendorSearchTerm && isAdminMode ? (
                <Reorder.Group 
                  axis="y" 
                  values={sortedVendors} 
                  onReorder={async (newOrder) => {
                    // Update order in Firestore
                    const batch = writeBatch(db);
                    newOrder.forEach((vendor, index) => {
                      const vRef = doc(db, 'vendors', vendor.id);
                      batch.update(vRef, { order: index });
                    });
                    try {
                      await batch.commit();
                    } catch (e) {
                      console.error("Error saving new vendor order:", e);
                    }
                  }}
                  className="space-y-0.5"
                >
                  {sortedVendors.map((vendor, index) => (
                    <Reorder.Item
                      key={`reorder-${vendor.id}`}
                      value={vendor}
                      id={`vendor-btn-${vendor.id}`}
                      onClick={() => {
                        if (viewMode === 'matrix') {
                          setMatrixVendorIds(prev => {
                            if (prev.includes(vendor.id)) return prev.filter(id => id !== vendor.id);
                            return [...prev, vendor.id];
                          });
                        } else {
                          setSelectedVendor(vendor);
                          setViewMode('detail');
                        }
                      }}
                      className={`w-full cursor-pointer text-left px-4 py-2.5 rounded-lg text-sm font-medium transition-all flex items-center justify-between group ${
                        (viewMode === 'matrix' ? matrixVendorIds.includes(vendor.id) : selectedVendor?.id === vendor.id)
                        ? 'bg-indigo-50 text-indigo-700' 
                        : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <GripVertical className="h-3.5 w-3.5 text-slate-300 group-hover:text-slate-400 shrink-0" />
                        {viewMode === 'matrix' && (
                          <div className={`w-4 h-4 rounded border flex items-center justify-center transition-all ${
                            matrixVendorIds.includes(vendor.id) 
                              ? 'bg-indigo-600 border-indigo-600 shadow-sm' 
                              : 'bg-white border-slate-300'
                          }`}>
                            {matrixVendorIds.includes(vendor.id) && (
                              <svg className="w-2.5 h-2.5 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={4}>
                                <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                              </svg>
                            )}
                          </div>
                        )}
                        <span className="text-[10px] font-mono text-slate-400 w-4 tabular-nums shrink-0">{index + 1}</span>
                        <span className="truncate">{vendor.name}</span>
                      </div>
                      <div className="flex items-center gap-2">
                        {isAdminMode && (
                          <button
                            onClick={(e) => deleteVendor(vendor.id, e)}
                            className="opacity-0 group-hover:opacity-100 p-1 hover:bg-rose-100 hover:text-rose-600 rounded transition-all"
                            title="업체 삭제"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </button>
                        )}
                        {selectedVendor?.id === vendor.id && (
                          <div className="w-1.5 h-1.5 rounded-full bg-indigo-600" />
                        )}
                      </div>
                    </Reorder.Item>
                  ))}
                </Reorder.Group>
              ) : (
                sortedVendors.map((vendor, index) => (
                  <div
                    key={`vendor-list-${vendor.id}`}
                    id={`vendor-btn-${vendor.id}`}
                    onClick={() => {
                      if (viewMode === 'matrix') {
                        setMatrixVendorIds(prev => {
                          if (prev.includes(vendor.id)) return prev.filter(id => id !== vendor.id);
                          return [...prev, vendor.id];
                        });
                      } else {
                        setSelectedVendor(vendor);
                        setViewMode('detail');
                      }
                    }}
                    className={`w-full cursor-pointer text-left px-4 py-2.5 rounded-lg text-sm font-medium transition-all flex items-center justify-between group ${
                      (viewMode === 'matrix' ? matrixVendorIds.includes(vendor.id) : selectedVendor?.id === vendor.id)
                      ? 'bg-indigo-50 text-indigo-700' 
                      : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
                    }`}
                  >
                    <div className="flex items-center gap-2 truncate">
                      {viewMode === 'matrix' && (
                        <div className={`w-3.5 h-3.5 rounded border border-slate-300 flex items-center justify-center transition-all ${matrixVendorIds.includes(vendor.id) ? 'bg-indigo-600 border-indigo-600' : 'bg-white'}`}>
                          {matrixVendorIds.includes(vendor.id) && <div className="w-1.5 h-1.5 bg-white rounded-full scale-50" />}
                        </div>
                      )}
                      <span className="text-[10px] font-mono text-slate-400 w-4 tabular-nums shrink-0">{index + 1}</span>
                      <span className="truncate">{vendor.name}</span>
                    </div>
                    <div className="flex items-center gap-2">
                      {isAdminMode && (
                        <button
                          onClick={(e) => deleteVendor(vendor.id, e)}
                          className="opacity-0 group-hover:opacity-100 p-1 hover:bg-rose-100 hover:text-rose-600 rounded transition-all"
                          title="업체 삭제"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      )}
                      {selectedVendor?.id === vendor.id && (
                        <div className="w-1.5 h-1.5 rounded-full bg-indigo-600" />
                      )}
                    </div>
                  </div>
                ))
              )}
            </div>
          </aside>
        )}

        {/* 3. MAIN CONTENT */}
        <main className="flex-1 flex flex-col overflow-hidden bg-[#F8F9FA]" id="main-content-view">
          {viewMode === 'matrix' ? (
            <div className="flex-1 flex flex-col overflow-hidden bg-white">
                <header className="px-8 py-4 border-b border-slate-100 shrink-0 bg-white">
                   <div className="flex flex-col gap-4">
                     <div className="flex items-center justify-between">
                       <div>
                         <h1 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
                           <ArrowLeftRight className="h-5 w-5 text-indigo-600" />
                           단가 통합 비교 매트릭스
                         </h1>
                         <p className="text-[11px] text-slate-400 font-medium tracking-tight mt-1">업체별 동일 품명/규격 상품의 단가를 한눈에 비교합니다.</p>
                       </div>
                       <div className="flex items-center gap-2">
                          <span className="text-[10px] font-bold text-slate-400 bg-slate-50 px-2 py-1 rounded border border-slate-100">
                             총 {matrixData.length}개 품목</span>{matrixVendorIds.length > 0 && <div className="flex items-center gap-2 ml-3"><button onClick={() => { setMatrixVendorIds([]); }} className="text-[10px] font-bold text-slate-400 hover:text-red-500 bg-white px-2 py-0.5 rounded border border-slate-100 cursor-pointer">선택 해제({matrixVendorIds.length})</button></div>}<span>
                          </span>
                       </div>
                     </div>
                     
                     <div className="flex items-center gap-2 overflow-x-auto pb-1 no-scrollbar">
                        {matrixCategories.map(cat => (
                          <button
                            key={`matrix-cat-${cat}`}
                            onClick={() => setMatrixCategory(cat)}
                            className={`px-3 py-1 rounded-full text-xs font-bold transition-all whitespace-nowrap ${
                              matrixCategory === cat
                                ? 'bg-indigo-600 text-white shadow-md'
                                : 'bg-slate-50 text-slate-500 hover:bg-slate-100'
                            }`}
                          >
                            {cat}
                          </button>
                        ))}
                     </div>
                   </div>
                </header>
                
                <div className="flex-1 overflow-auto p-6 bg-slate-50/30">
                   {isMatrixLoading ? (
                     <div className="h-full flex flex-col items-center justify-center gap-4 text-slate-400">
                        <Loader2 className="h-8 w-8 animate-spin" />
                        <p className="text-sm font-bold">전체 업체 단가 정보를 불러오는 중...</p>
                     </div>
                   ) : matrixData.length === 0 ? (
                     <div className="h-full flex flex-col items-center justify-center gap-4 text-slate-300">
                        <Search className="h-12 w-12 opacity-20" />
                        <p className="text-sm font-bold opacity-40">비교할 데이터가 없습니다. (검색어 또는 카테고리를 확인해주세요)</p>
                     </div>
                   ) : (
                     <div className="border border-slate-200 rounded-2xl overflow-hidden shadow-xl bg-white">
                        <div className="overflow-x-auto">
                          <table className="w-full text-[11px] border-collapse font-medium min-w-[2000px]">
                            <thead className="bg-slate-900 text-white sticky top-0 z-20">
                              <tr className="h-12">
                                <th className="px-6 text-left font-bold border-r border-slate-800 sticky left-0 bg-slate-900 z-30" style={{ width: '250px' }}>구분 품명</th>
                                <th className="px-4 text-left font-bold border-r border-slate-800" style={{ width: '150px' }}>규격</th>
                                <th className="px-3 text-center font-bold border-r border-slate-800" style={{ width: '80px' }}>단위</th>
                                {filteredVendors.map(vendor => (
                                  <th key={`matrix-header-th-${vendor.id}`} className="px-4 text-center font-bold border-r border-slate-800 bg-slate-800/50 min-w-[120px]">
                                    {vendor.name}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {matrixData.map((row, idx) => {
                                  // Find min price for highlighting
                                  const priceEntries = Object.values(row.prices) as any[];
                                  const priceValues = priceEntries.map(p => p.unitPrice);
                                  const minPrice = priceValues.length > 0 ? Math.min(...priceValues) : null;
                                
                                return (
                                  <tr key={`matrix-row-${idx}`} className="hover:bg-indigo-50/30 transition-colors h-10 group">
                                    <td className="px-6 font-bold text-slate-800 border-r border-slate-100 sticky left-0 bg-white group-hover:bg-indigo-50/50 z-10 transition-colors shadow-[2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                                      <div className="flex items-center justify-between">
                                        <div className="flex flex-col truncate">
                                          <span className="text-[10px] text-indigo-400 font-bold uppercase tracking-widest">{row.category}</span>
                                          <span className="truncate">{row.itemName}</span>
                                        </div>
                                        <button 
                                          onClick={() => setSelectedMatrixRow(row)}
                                          className="p-1.5 text-slate-300 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                                          title="단가 비교 차트 보기"
                                        >
                                          <BarChart3 className="h-3.5 w-3.5" />
                                        </button>
                                      </div>
                                    </td>
                                    <td className="px-4 text-slate-500 font-bold border-r border-slate-100 truncate">{row.spec}</td>
                                    <td className="px-3 text-center text-slate-400 border-r border-slate-100">{row.unit}</td>
                                    {filteredVendors.map(vendor => {
                                      const entry = row.prices[vendor.id];
                                      const priceValue = entry?.unitPrice;
                                      const isMin = priceValue && minPrice && priceValue === minPrice && priceValues.length > 1;
                                      
                                      return (
                                        <td 
                                          key={`matrix-body-price-${row.itemName}-${row.spec || 'ns'}-${vendor.id}`} 
                                          className={`px-4 text-center border-r border-slate-100 font-mono transition-all ${
                                            isMin ? 'bg-emerald-50 text-emerald-700 font-black' : 'text-slate-600'
                                          }`}
                                        >
                                          {entry ? (
                                            <div className="flex flex-col items-center">
                                              <div className="flex items-center gap-1 leading-tight">
                                                <span>₩{entry.unitPrice.toLocaleString()}</span>
                                                {entry.hasPendingUpdate && (
                                                  <div className="w-2 h-2 rounded-full bg-amber-500 animate-pulse shrink-0" title="승인 대기 중인 변경 요청이 있습니다." />
                                                )}
                                              </div>
                                              <span className="text-[8px] font-black text-indigo-400 opacity-70">
                                                 {entry.negoType === 'sts_pipe' ? `STS ${entry.negoRate}` : `${entry.negoRate}%`}
                                              </span>
                                            </div>
                                          ) : (
                                            <span className="text-slate-200">데이터 없음</span>
                                          )}
                                        </td>
                                      );
                                    })}
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                     </div>
                   )}
                </div>
            </div>
          ) : selectedVendor ? (
            <>
              {/* VENDOR INFO HEADER */}
              <div className="bg-white border-b border-slate-200 px-6 py-4 shrink-0 relative overflow-hidden">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 relative z-10">
                  <div className="flex flex-col gap-3">
                    <div className="flex items-center gap-3">
                      <div className="bg-indigo-600 text-white p-2 rounded-xl shadow-lg shadow-indigo-100 shrink-0">
                        <Building2 className="h-5 w-5" />
                      </div>
                      <div className="flex items-center gap-3 flex-wrap">
                        <h1 className="text-xl font-black text-slate-900 tracking-tight">{selectedVendor.name}</h1>
                        <button 
                          onClick={(e) => copyVendorLink(selectedVendor.id, e)}
                          className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[9px] font-black uppercase tracking-wider transition-all border ${
                            copySuccess 
                            ? 'bg-emerald-50 border-emerald-200 text-emerald-600' 
                            : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-white hover:border-indigo-400 hover:text-indigo-600'
                          }`}
                        >
                          <LinkIcon className="h-2.5 w-2.5" />
                          {copySuccess ? '복사됨!' : '링크 복사'}
                        </button>
                        <div className="text-[10px] text-slate-400 font-bold bg-slate-50 px-2 py-0.5 rounded border border-slate-100">
                          V-ID: {selectedVendor.id.slice(0, 8)}
                        </div>
                        {isAdminMode && (
                          <div className="inline-flex items-center gap-2">
                            {showVendorPasswords[selectedVendor.id] ? (
                              <span className="text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-100 font-mono text-[10px] flex items-center gap-1.5">
                                <Lock className="h-2.5 w-2.5" />
                                <span className="font-black">{selectedVendor.password}</span>
                                <button 
                                  onClick={() => setShowVendorPasswords(prev => ({ ...prev, [selectedVendor.id]: false }))}
                                  className="text-[9px] text-slate-400 hover:text-slate-600 underline"
                                >
                                  숨기기
                                </button>
                              </span>
                            ) : (
                              <button 
                                onClick={() => setShowVendorPasswords(prev => ({ ...prev, [selectedVendor.id]: true }))}
                                className="text-[9px] bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded border border-indigo-100 hover:bg-indigo-100 transition-colors flex items-center gap-1"
                              >
                                <Unlock className="h-2.5 w-2.5" />
                                비밀번호 보기
                              </button>
                            )}
                          </div>
                        )}
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 border-t border-slate-50 pt-2">
                      {[
                        { label: '대표자', value: selectedVendor.representative },
                        { label: '전화', value: selectedVendor.phone },
                        { label: '팩스', value: selectedVendor.fax },
                        { label: '사업자번호', value: selectedVendor.businessNumber },
                        { label: '이메일', value: selectedVendor.email }
                      ].map((item, idx) => (
                        <div key={`vendor-info-${item.label}`} className="flex items-center gap-1.5">
                          <span className="text-[9px] font-black text-slate-400/80 uppercase tracking-tighter">{item.label}</span>
                          <span className={`text-[11px] font-bold ${item.label === '이메일' ? 'text-indigo-600' : 'text-slate-600'}`}>
                            {item.value || '-'}
                          </span>
                          {idx < 4 && <div className="h-2 w-[1px] bg-slate-200 ml-1.5" />}
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="flex items-center gap-3">
                    {selectedVendor.priceTableUrl && (
                      <button 
                        onClick={() => setIsViewingPriceTable(true)}
                        className="flex items-center gap-2 px-5 py-2.5 bg-indigo-50 text-indigo-600 rounded-xl text-xs font-bold hover:bg-indigo-100 transition-all active:scale-95 border border-indigo-100"
                      >
                        <FileText className="h-3.5 w-3.5" />
                        단가표 원본 보기
                      </button>
                    )}
                    {(isAdminMode || (selectedVendor && isVerified === selectedVendor.id)) && (
                      <div className="flex items-center gap-3">
                        <button 
                          onClick={() => setIsEditingVendorInfo(true)}
                          className="flex items-center gap-2 px-5 py-2.5 bg-slate-900 text-white rounded-xl text-xs font-bold hover:bg-slate-800 transition-all shadow-lg shadow-slate-200 active:scale-95"
                        >
                          <Edit2 className="h-3.5 w-3.5" />
                          거래처 정보 수정
                        </button>
                        <button 
                          onClick={() => setIsChangingPassword(true)}
                          className="flex items-center gap-2 px-5 py-2.5 bg-white text-slate-700 rounded-xl text-xs font-bold hover:bg-slate-50 transition-all border border-slate-200 active:scale-95 shadow-sm"
                        >
                          <Lock className="h-3.5 w-3.5" />
                          비밀번호 변경
                        </button>
                        {isAdminMode && (
                          <>
                            <button 
                              onClick={handleBackupCurrentVendor}
                              className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 text-white rounded-xl text-xs font-bold hover:bg-indigo-700 transition-all active:scale-95 shadow-lg shadow-indigo-100"
                            >
                              <Download className="h-3.5 w-3.5" />
                              단가표 백업 (Excel)
                            </button>
                            <button 
                              onClick={(e) => deleteVendor(selectedVendor.id, e)}
                              className="flex items-center gap-2 px-5 py-2.5 bg-rose-50 text-rose-600 rounded-xl text-xs font-bold hover:bg-rose-100 transition-all active:scale-95 border border-rose-100"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                              업체 삭제
                            </button>
                          </>
                        )}
                      </div>
                    )}
                  </div>
                </div>
                
                {/* Background Accent */}
                <div className="absolute top-0 right-0 w-64 h-64 bg-indigo-50/30 rounded-full blur-3xl -mr-32 -mt-32 pointer-events-none" />
              </div>

              {/* TOOLBAR */}
              {canManageItems && (
                <div className="h-11 border-b border-slate-200 flex items-center justify-between px-6 shrink-0 bg-white z-20 shadow-sm shadow-slate-100">
                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-3">
                    <span className="text-[10px] font-bold text-indigo-600 whitespace-nowrap">선택 {selectedPriceIds.size}개</span>
                    <div className="h-3 w-[1px] bg-slate-200" />
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-medium text-slate-500 whitespace-nowrap">단가 가감</span>
                      <div className="relative flex items-center">
                        <input 
                          type="number"
                          value={bulkNegoValue}
                          onChange={(e) => setBulkNegoValue(Number(e.target.value))}
                          className="w-14 h-7 bg-slate-50 border border-slate-200 rounded px-2 pr-4 text-[11px] font-bold text-center outline-none focus:ring-1 focus:ring-indigo-500 transition-all hover:bg-slate-100"
                        />
                        <span className="absolute right-1 text-[9px] text-slate-400 font-bold">%</span>
                      </div>
                    </div>
                    <div className="h-3 w-[1px] bg-slate-200" />
                    <div className="flex items-center gap-2">
                       <span className="text-[10px] font-medium text-slate-500 whitespace-nowrap">카테고리 이동</span>
                       <select 
                         onChange={(e) => {
                           const val = e.target.value;
                           if (val && selectedPriceIds.size > 0) {
                             if (confirm(`${selectedPriceIds.size}개 품목의 카테고리를 '${val}'(으)로 변경하시겠습니까?`)) {
                               handleBulkCategoryUpdate(val);
                             }
                             e.target.value = "";
                           }
                         }}
                         className="h-7 bg-slate-50 border border-slate-200 rounded px-2 text-[10px] font-bold text-slate-600 outline-none focus:ring-1 focus:ring-indigo-500 transition-all hover:bg-slate-100 cursor-pointer"
                       >
                         <option value="">카테고리 선택</option>
                         {categories.filter(c => c !== '전체').map(cat => (
                           <option key={`bulk-cat-opt-${cat}`} value={cat}>{cat}</option>
                         ))}
                       </select>
                    </div>
                    <button 
                      onClick={async () => {
                        if (selectedPriceIds.size === 0) return;
                        
                        // Link mode confirmation
                        if (isDeepLinkMode && !isAdminMode) {
                          if (!confirm(`선택한 ${selectedPriceIds.size}개 품목에 대해 네고율 ${bulkNegoValue}%를 일괄 승인 요청하시겠습니까?`)) {
                            return;
                          }
                        }

                        try {
                          setLoading(true);
                          let skipped = 0;
                          let pendingCount = 0;
                          
                          const targets = Array.from(selectedPriceIds)
                            .map(id => priceItems.find(p => p.id === id))
                            .filter(Boolean) as PriceItem[];

                          const chunkSize = 400;
                          for (let i = 0; i < targets.length; i += chunkSize) {
                            const chunk = targets.slice(i, i + chunkSize);
                            const batch = writeBatch(db);

                            for (const item of chunk) {
                              if (item.negoType === 'sts_pipe') {
                                skipped++;
                                continue;
                              }
                              const newUnitPrice = calculatePrice(item.costPrice, bulkNegoValue, item.useRounding, item.negoType || 'percent', item.weight);
                              
                              if (isDeepLinkMode && !isAdminMode) {
                                if (!item.hasPendingUpdate) {
                                  const pendingRef = doc(collection(db, 'pending_updates'));
                                  batch.set(pendingRef, {
                                    vendorId: selectedVendor.id,
                                    priceItemId: item.id,
                                    itemName: item.itemName,
                                    spec: item.spec || '',
                                    oldData: { negoRate: item.negoRate, unitPrice: item.unitPrice },
                                    newData: { negoRate: bulkNegoValue, unitPrice: newUnitPrice },
                                    status: 'pending',
                                    requestedBy: '업체(링크)',
                                    requestedAt: serverTimestamp()
                                  });
                                  
                                  const itemRef = doc(db, 'vendors', selectedVendor.id, 'prices', item.id);
                                  batch.update(itemRef, { hasPendingUpdate: true });
                                  pendingCount++;
                                }
                              } else {
                                batch.update(doc(db, 'vendors', selectedVendor.id, 'prices', item.id), {
                                  negoRate: bulkNegoValue,
                                  unitPrice: newUnitPrice,
                                  updatedAt: serverTimestamp()
                                });
                              }
                            }
                            await batch.commit();
                          }

                          if (isDeepLinkMode && !isAdminMode) {
                            alert(`${pendingCount}개 품목에 대한 승인 요청이 전송되었습니다.`);
                          } else {
                            alert(`선택항목 적용 완료${skipped > 0 ? ` (STS 품목 ${skipped}개 제외)` : ''}`);
                          }
                          
                          setSelectedPriceIds(new Set());
                        } catch (error) {
                          console.error("Selection update error:", error);
                          alert('업데이트 중 오류가 발생했습니다.');
                        } finally {
                          setLoading(false);
                        }
                      }}
                      className="h-8 px-4 bg-[#A3D169] hover:bg-[#8FBC4A] text-[#3D5620] text-xs font-bold rounded shadow-sm transition-all"
                    >
                      선택 적용
                    </button>
                    
                    {selectedPriceIds.size > 0 && isActuallyAuthorized && (
                      <button 
                        onClick={handleBulkDelete}
                        className="h-8 px-4 bg-rose-50 text-rose-600 border border-rose-100 hover:bg-rose-100 text-xs font-bold rounded shadow-sm transition-all flex items-center gap-1.5"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                        선택 삭제
                      </button>
                    )}
                    <button 
                      onClick={async () => {
                        const targetItems = activeCategory === '전체' 
                          ? priceItems 
                          : priceItems.filter(i => i.category === activeCategory);
                        
                        if (targetItems.length === 0) {
                          alert('적용할 대상 품목이 없습니다.');
                          return;
                        }

                        if (!confirm(`${activeCategory === '전체' ? '전체' : `'${activeCategory}'`} ${targetItems.length}개 품목 중 % 네고 항목에 일괄 적용하시겠습니까?`)) return;
                        
                        try {
                          setLoading(true);
                          const eligibleItems = targetItems.filter(i => i.negoType !== 'sts_pipe');
                          
                          if (eligibleItems.length === 0) {
                            alert('적용 가능한 % 네고 항목이 없습니다.');
                            setLoading(false);
                            return;
                          }

                          const chunkSize = 400;
                          let pendingCount = 0;
                          
                          for (let i = 0; i < eligibleItems.length; i += chunkSize) {
                            const chunk = eligibleItems.slice(i, i + chunkSize);
                            const batch = writeBatch(db);

                            for (const item of chunk) {
                              if (isDeepLinkMode && !isAdminMode) {
                                if (!item.hasPendingUpdate) {
                                  const newUnitPrice = calculatePrice(item.costPrice, bulkNegoValue, item.useRounding, item.negoType || 'percent', item.weight);
                                  const pendingRef = doc(collection(db, 'pending_updates'));
                                  batch.set(pendingRef, {
                                    vendorId: selectedVendor.id,
                                    priceItemId: item.id,
                                    itemName: item.itemName,
                                    spec: item.spec || '',
                                    oldData: { negoRate: item.negoRate, unitPrice: item.unitPrice },
                                    newData: { negoRate: bulkNegoValue, unitPrice: newUnitPrice },
                                    status: 'pending',
                                    requestedBy: '업체(링크)',
                                    requestedAt: serverTimestamp()
                                  });
                                  const itemRef = doc(db, 'vendors', selectedVendor.id, 'prices', item.id);
                                  batch.update(itemRef, { hasPendingUpdate: true });
                                  pendingCount++;
                                }
                              } else {
                                const newUnitPrice = calculatePrice(item.costPrice, bulkNegoValue, item.useRounding, item.negoType || 'percent', item.weight);
                                batch.update(doc(db, 'vendors', selectedVendor.id, 'prices', item.id), {
                                  negoRate: bulkNegoValue,
                                  unitPrice: newUnitPrice,
                                  updatedAt: serverTimestamp()
                                });
                              }
                            }
                            await batch.commit();
                          }

                          if (isDeepLinkMode && !isAdminMode) {
                            alert(`${pendingCount}개 품목에 대한 승인 요청이 전송되었습니다.`);
                          } else {
                            alert(`${activeCategory === '전체' ? '전체' : `'${activeCategory}'`} ${eligibleItems.length}개 품목 업데이트 완료`);
                          }
                        } catch (error) {
                          console.error("Full update error:", error);
                          alert('일괄 업데이트 중 오류가 발생했습니다.');
                        } finally {
                          setLoading(false);
                        }
                      }}
                      className="h-8 px-4 border border-slate-200 hover:bg-slate-50 text-slate-600 text-xs font-bold rounded transition-all"
                    >
                      {activeCategory === '전체' ? '전체 적용' : `'${activeCategory}' 적용`}
                    </button>
                    {isActuallyAuthorized && (
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => setIsBulkItemUploadOpen(true)}
                          className="h-8 px-4 border border-indigo-200 text-indigo-600 bg-indigo-50/50 text-xs font-bold rounded hover:bg-indigo-100 transition-all flex items-center gap-1.5"
                        >
                          <Table className="h-3.5 w-3.5" />
                          엑셀 일괄 등록
                        </button>
                        <button 
                          onClick={() => setIsAddingItem(true)}
                          className="h-8 px-4 bg-slate-900 text-white text-xs font-bold rounded hover:bg-slate-800 transition-all"
                        >
                          품목 추가
                        </button>
                      </div>
                    )}
                  </div>
                  
                  <div className="h-4 w-[1px] bg-slate-200" />
                  
                  <div className="flex items-center gap-3 text-slate-400">
                    <span className="text-[10px] font-medium leading-none whitespace-nowrap tracking-tight">Ctrl+V 붙여넣기 · ↑ ↓ 0.5% 조정 · Enter 다음행</span>
                  </div>
                </div>
                
                <div className="flex items-center gap-1 text-[10px] font-bold text-slate-400">
                  <span className="text-indigo-600 font-black">{priceItems.length}</span>건 표시 중
                </div>
              </div>
              )}

              {/* CATEGORY TABS */}
              <div className="flex items-center gap-1 px-6 py-2 bg-[#F8F9FA] border-b border-slate-100 shrink-0 overflow-x-auto no-scrollbar">
                <div className="flex items-center gap-2 mr-3 shrink-0 py-1 border-r border-slate-200 pr-3">
                  <div className="w-1 h-3.5 bg-indigo-500 rounded-full" />
                  <span className="text-[11px] font-black text-slate-800 tracking-tight">취급품목</span>
                </div>
                {['전체', ...categories.filter(c => c !== '전체')].map(cat => (
                  <div key={`cat-tab-${cat}`} className="relative group/cat">
                    <button
                      onClick={() => setActiveCategory(cat)}
                      className={`px-4 py-1.5 rounded-full text-xs font-bold transition-all flex items-center gap-1.5 ${
                        activeCategory === cat 
                        ? 'bg-slate-900 text-white shadow-md' 
                        : 'text-slate-500 hover:text-slate-900 hover:bg-slate-200/50'
                      }`}
                    >
                      {cat}
                      <span className={`text-[10px] ${activeCategory === cat ? 'text-slate-400' : 'text-slate-300'}`}>
                        {cat === '전체' ? priceItems.length : priceItems.filter(i => i.category === cat).length}
                      </span>
                    </button>
                    {cat !== '전체' && canManageItems && activeCategory === cat && (
                      <button 
                        onClick={async (e) => {
                          e.stopPropagation();
                          if (confirm(`'${cat}' 카테고리를 삭제하시겠습니까?\n해당 카테고리의 품목들을 먼저 확인해주세요.`)) {
                            const newCats = categories.filter(c => c !== cat);
                            await updateVendorCategories(newCats);
                            setActiveCategory('전체');
                          }
                        }}
                        className="absolute -top-1 -right-1 w-4 h-4 bg-rose-500 text-white rounded-full flex items-center justify-center opacity-0 group-hover/cat:opacity-100 transition-opacity scale-75"
                      >
                        <X className="h-2 w-2" />
                      </button>
                    )}
                  </div>
                ))}
                {canManageItems && (
                  <button 
                    onClick={async () => {
                      const newCat = prompt('새 카테고리명을 입력하세요:');
                      if (newCat) {
                        const trimmed = newCat.trim();
                        if (trimmed && !categories.includes(trimmed)) {
                          await updateVendorCategories([...categories, trimmed]);
                        } else if (categories.includes(trimmed)) {
                          alert('이미 존재하는 카테고리명입니다.');
                        }
                      }
                    }}
                    className="p-1.5 text-slate-300 hover:text-indigo-600 transition-colors"
                    title="카테고리 추가"
                  >
                    <Plus className="h-4 w-4" />
                  </button>
                )}
              </div>

              {/* TABLE CONTAINER */}
              <div className="flex-1 overflow-auto bg-white relative">
                {/* Password Overlay */}
                <AnimatePresence>
                  {isVerified !== selectedVendor.id && !isAdminMode && !isGuestMode && (
                    <motion.div 
                      key="password-overlay"
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="absolute inset-0 bg-white/60 backdrop-blur-md z-30 flex items-center justify-center p-8"
                    >
                      <motion.div 
                        initial={{ scale: 0.95, opacity: 0 }}
                        animate={{ scale: 1, opacity: 1 }}
                        className="bg-white border border-slate-200 p-8 rounded-2xl shadow-2xl w-full max-w-sm text-center"
                      >
                        <Lock className="h-10 w-10 text-indigo-500 mx-auto mb-6" />
                        <h3 className="text-lg font-black text-slate-900 mb-2">업체 보안 인증</h3>
                        <p className="text-xs text-slate-500 font-medium mb-8 leading-relaxed">
                          접근 권한이 없습니다.<br />업체 전용 비밀번호를 입력해 주세요.
                        </p>
                        <form onSubmit={verifyPassword} className="space-y-4">
                          <input 
                            autoFocus
                            type="password" 
                            placeholder="비밀번호" 
                            value={passwordInput}
                            onChange={(e) => setPasswordInput(e.target.value)}
                            className="w-full h-12 bg-slate-50 border border-slate-200 rounded-xl px-4 text-center text-xl font-bold tracking-[0.5em] focus:ring-2 focus:ring-indigo-500/20 focus:outline-none transition-all outline-none"
                          />
                          <button className="w-full h-12 bg-indigo-600 text-white font-bold rounded-xl hover:bg-indigo-700 transition-all active:scale-95 shadow-lg shadow-indigo-100">
                            연결하기
                          </button>
                        </form>
                      </motion.div>
                    </motion.div>
                  )}
                </AnimatePresence>

                <table className="w-full border-collapse table-fixed min-w-[1200px]">
                  <thead className="sticky top-0 bg-white z-10 border-b border-slate-200">
                    <tr className="text-[10px] font-bold text-slate-400 uppercase tracking-tight h-8">
                      <th className="w-10 px-4 text-center relative border-r border-slate-50">
                        <input 
                          type="checkbox" 
                          checked={selectedPriceIds.size === priceItems.length && priceItems.length > 0}
                          onChange={toggleSelectAll}
                          className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 ring-offset-0 focus:ring-0 cursor-pointer"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th cursor-pointer hover:bg-slate-50 transition-colors" style={{ width: columnWidths.itemCode }}>
                        품번
                        <div 
                          onMouseDown={(e) => startResize(e, 'itemCode')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th" style={{ width: columnWidths.category }}>
                        카테고리
                        <div 
                          onMouseDown={(e) => startResize(e, 'category')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th cursor-pointer hover:bg-slate-50 transition-colors" style={{ width: columnWidths.itemName }} onClick={() => handlePriceSort('itemName')}>
                        <div className="flex items-center gap-1">
                          품명
                          {priceSortField === 'itemName' && (
                            <ArrowUpDown className={`h-3 w-3 ${priceSortOrder === 'asc' ? 'text-indigo-500' : 'text-rose-500'}`} />
                          )}
                        </div>
                        <div 
                          onMouseDown={(e) => startResize(e, 'itemName')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th cursor-pointer hover:bg-slate-50 transition-colors" style={{ width: columnWidths.spec }} onClick={() => handlePriceSort('spec')}>
                        <div className="flex items-center gap-1">
                          규격
                          {priceSortField === 'spec' && (
                            <ArrowUpDown className={`h-3 w-3 ${priceSortOrder === 'asc' ? 'text-indigo-500' : 'text-rose-500'}`} />
                          )}
                        </div>
                        <div 
                          onMouseDown={(e) => startResize(e, 'spec')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th" style={{ width: columnWidths.unit }}>
                        단위
                        <div 
                          onMouseDown={(e) => startResize(e, 'unit')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-right font-semibold bg-indigo-50/50 text-indigo-700 relative border-r border-slate-50 group/th" style={{ width: columnWidths.unitPrice }}>
                        <div className="flex items-center justify-end gap-1.5">
                           <span>구매단가</span>
                           <div className="group/hint relative">
                              <Info className="h-3 w-3 text-indigo-300" />
                              <div className="absolute bottom-full right-0 mb-2 w-32 p-2 bg-slate-800 text-white text-[9px] rounded-lg opacity-0 group-hover/hint:opacity-100 pointer-events-none transition-opacity z-50">
                                체크 시 10원 단위 반올림 적용
                              </div>
                           </div>
                        </div>
                        <div 
                          onMouseDown={(e) => startResize(e, 'unitPrice')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-center font-semibold bg-slate-50/80 relative border-r border-slate-50 group/th" style={{ width: columnWidths.negoRate }}>
                        네고율 / KG단가
                        <div 
                          onMouseDown={(e) => startResize(e, 'negoRate')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-center font-semibold bg-slate-50/80 relative border-r border-slate-50 group/th" style={{ width: columnWidths.weight }}>
                        단중 (참조)
                        <div 
                          onMouseDown={(e) => startResize(e, 'weight')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-right font-semibold relative border-r border-slate-50 group/th" style={{ width: columnWidths.costPrice }}>
                        협가 / 단중
                        <div 
                          onMouseDown={(e) => startResize(e, 'costPrice')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-right font-semibold bg-slate-50/80 relative border-r border-slate-50 group/th" style={{ width: columnWidths.discountAmount }}>
                        할인액
                        <div 
                          onMouseDown={(e) => startResize(e, 'discountAmount')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-right font-semibold relative border-r border-slate-50 group/th" style={{ width: columnWidths.change }}>
                        변동
                        <div 
                          onMouseDown={(e) => startResize(e, 'change')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th" style={{ width: columnWidths.maker }}>
                        메이커
                        <div 
                          onMouseDown={(e) => startResize(e, 'maker')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      <th className="px-4 text-left font-semibold relative border-r border-slate-50 group/th" style={{ width: columnWidths.lastUpdated }}>
                        최근변경일
                        <div 
                          onMouseDown={(e) => startResize(e, 'lastUpdated')}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-indigo-400 transition-colors bg-transparent z-10"
                        />
                      </th>
                      {canManageItems && (
                        <th className="w-12 px-4 text-center font-semibold">관리</th>
                      )}
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-slate-100 text-[12px]">
                    {sortedPriceItems.filter(item => {
                      const matchesSearch = item.itemName.toLowerCase().includes(searchTerm.toLowerCase()) || 
                                           (item.itemCode && item.itemCode.toLowerCase().includes(searchTerm.toLowerCase()));
                      const matchesCategory = activeCategory === '전체' || item.category === activeCategory;
                      return matchesSearch && matchesCategory;
                    }).map((item, idx) => {
                      const discount = item.negoType === 'sts_pipe' ? 0 : item.costPrice * (item.negoRate / 100);
                      return (
                        <tr 
                          key={`price-row-${item.id}`} 
                          className={`hover:bg-indigo-50/20 group transition-colors h-[34px] ${selectedPriceIds.has(item.id) ? 'bg-indigo-50/10' : ''}`}
                        >
                          <td className="px-4 text-center">
                            <input 
                              type="checkbox" 
                              checked={selectedPriceIds.has(item.id)}
                              onChange={() => toggleSelectItem(item.id)}
                              className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 cursor-pointer"
                            />
                          </td>
                          <td className="px-4 text-slate-400 font-mono text-[10px] leading-none">{item.itemCode || `DP-${String(idx+1).padStart(3, '0')}`}</td>
                          <td className="px-1 text-center">
                            <select 
                              value={item.category || ''} 
                              onChange={(e) => handleInlinePriceUpdate(item.id, 'category', e.target.value)}
                              className="w-full bg-transparent border-none text-[10px] font-bold text-slate-500 outline-none focus:ring-0 appearance-none cursor-pointer hover:text-indigo-600 transition-colors"
                            >
                              <option value="">미지정</option>
                              {categories.filter(c => c !== '전체').map(cat => (
                                <option key={`inline-cat-opt-${item.id}-${cat}`} value={cat}>{cat}</option>
                              ))}
                            </select>
                          </td>
                          <td className="px-4 font-bold text-slate-800 truncate">
                            <div className="flex items-center gap-2">
                              {item.itemName}
                              {item.hasPendingUpdate && (
                                <div className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-amber-50 rounded border border-amber-200 shrink-0" title="승인 대기 중인 변경 요청이 있습니다.">
                                  <span className="flex h-1.5 w-1.5 rounded-full bg-amber-500 animate-pulse"></span>
                                  <span className="text-[9px] font-black text-amber-600 uppercase tracking-tighter">변경요청</span>
                                </div>
                              )}
                            </div>
                          </td>
                          <td className="px-4 text-slate-500 font-medium truncate text-xs">{item.spec}</td>
                          <td className="px-4 text-slate-400 font-medium text-xs">{item.unit}</td>
                          <td className="px-4 text-right font-mono font-black text-indigo-700 bg-indigo-50/10">
                            <div className="flex items-center justify-end gap-2">
                              <input 
                                type="checkbox"
                                checked={item.useRounding !== undefined ? item.useRounding : (selectedVendor?.roundingMethod !== 'none')}
                                onChange={(e) => handleInlinePriceUpdate(item.id, 'useRounding', e.target.checked)}
                                className="w-3 h-3 rounded border-slate-300 text-indigo-400 focus:ring-0 cursor-pointer"
                                title="10원 단위 반올림 적용 여부"
                              />
                              <span>{getRoundedValue(item.unitPrice, item.useRounding).toLocaleString()}</span>
                            </div>
                          </td>
                          <td className="px-4 text-center bg-slate-50/10 font-bold text-indigo-600">
                             <div className="flex items-center justify-center gap-0.5">
                                <input 
                                  type="number"
                                  value={item.negoRate}
                                  onChange={(e) => handleInlinePriceUpdate(item.id, 'negoRate', Number(e.target.value))}
                                  className={`w-12 bg-transparent border-none text-center p-0 outline-none focus:ring-0 appearance-none font-bold text-xs ${item.negoType === 'sts_pipe' ? 'text-indigo-800' : ''}`}
                                  placeholder={item.negoType === 'sts_pipe' ? "단가" : "%"}
                                />
                                <button 
                                  onClick={() => {
                                    const nextType = item.negoType === 'percent' ? 'sts_pipe' : 'percent';
                                    handleInlinePriceUpdate(item.id, 'negoType', nextType);
                                  }}
                                  className="text-[10px] font-medium opacity-50 hover:opacity-100 hover:text-indigo-600 transition-all cursor-pointer bg-slate-100 px-1 rounded shrink-0"
                                  title="네고 방식 변경 (% <-> STS)"
                                >
                                  {item.negoType === 'sts_pipe' ? 'STS' : '%'}
                                </button>
                             </div>
                          </td>
                          <td className="px-4 text-center bg-slate-50/10 font-bold text-slate-400">
                            <input 
                              type="text"
                              inputMode="decimal"
                              value={(item.weight || 0).toLocaleString(undefined, { maximumFractionDigits: 4 })}
                              onChange={(e) => {
                                const val = e.target.value.replace(/,/g, '');
                                if (val === '' || !isNaN(Number(val))) {
                                  handleInlinePriceUpdate(item.id, 'weight', val === '' ? 0 : Number(val));
                                }
                              }}
                              className="w-16 bg-transparent border-none text-center p-0 outline-none focus:ring-0 rounded appearance-none font-medium text-[10px] opacity-40 hover:opacity-100"
                              title="참조용 단중 데이터"
                            />
                          </td>
                          <td className={`px-4 text-right font-mono font-medium h-full ${item.negoType === 'sts_pipe' ? 'text-indigo-600 font-bold' : 'text-slate-600'}`}>
                            <input 
                              type="text"
                              inputMode="decimal"
                              value={(item.costPrice || 0).toLocaleString()}
                              onChange={(e) => {
                                const val = e.target.value.replace(/,/g, '');
                                if (val === '' || !isNaN(Number(val))) {
                                  handleInlinePriceUpdate(item.id, 'costPrice', val === '' ? 0 : Number(val));
                                }
                              }}
                              className={`w-full bg-transparent border-none text-right p-0 outline-none focus:ring-0 appearance-none font-bold ${item.negoType === 'sts_pipe' ? 'text-indigo-600' : 'text-slate-600'}`}
                            />
                          </td>
                          <td className="px-4 text-right font-mono font-bold text-emerald-600/80 bg-slate-50/10 text-xs">
                            {item.negoType === 'sts_pipe' ? '-' : `▼${getRoundedValue(discount, item.useRounding).toLocaleString()}`}
                          </td>
                          <td className="px-4 text-right">
                             {(() => {
                               const base = getRoundedValue(item.baseUnitPrice ?? item.unitPrice, item.useRounding);
                               const current = getRoundedValue(item.unitPrice, item.useRounding);
                               const diff = current - base;
                               const percent = (base && base !== 0) ? (diff / base) * 100 : 0;
                               
                               if (Math.abs(percent) < 0.01) return <span className="text-[10px] font-bold text-slate-300">-</span>;
                               
                               return (
                                 <span className={`text-[10px] font-bold ${percent > 0 ? 'text-rose-500' : 'text-indigo-500'}`}>
                                    {percent > 0 ? '+' : ''}{percent.toFixed(1)}%
                                 </span>
                               );
                             })()}
                          </td>
                          <td className="px-4 text-slate-500 font-medium truncate">{item.maker || '-'}</td>
                          <td className="px-4 text-slate-400 text-[10px] tabular-nums truncate">
                            {item.updatedAt?.toDate ? item.updatedAt.toDate().toLocaleString('ko-KR', {
                              year: '2-digit',
                              month: '2-digit',
                              day: '2-digit',
                              hour: '2-digit',
                              minute: '2-digit'
                            }) : '-'}
                          </td>
                          {canManageItems && (
                            <td className="px-4 py-2">
                              <div className="flex items-center justify-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button 
                                  onClick={() => setEditingPriceId(item.id)}
                                  className="p-1.5 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all"
                                  title="상세 수정"
                                >
                                  <Edit2 className="h-3.5 w-3.5" />
                                </button>
                                <button 
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    deletePriceItem(item.id);
                                  }}
                                  className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                                  title="품목 삭제"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            </td>
                          )}
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
              
              {/* TABLE FOOTER SUMMARY */}
              <div className="h-12 border-t border-slate-200 bg-white flex items-center justify-between px-6 shrink-0 z-20">
                <div className="flex items-center gap-8 text-[11px] font-bold">
                  <span className="text-slate-400">선택 {selectedPriceIds.size}개</span>
                  <div className="flex items-center gap-2">
                    <span className="text-slate-500 font-medium">구매단가 합계</span>
                    <span className="text-indigo-700 font-mono text-[13px] font-black">₩{priceItems.reduce((acc, i) => acc + getRoundedValue(i.unitPrice || 0, i.useRounding), 0).toLocaleString()}</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-slate-500 font-medium">협가 합계</span>
                    <span className="text-slate-900 font-mono">₩{priceItems.reduce((acc, i) => acc + (i.costPrice || 0), 0).toLocaleString()}</span>
                  </div>
                  <div className="px-3 py-1 bg-[#F1F8E9] text-[#4A6332] rounded-md flex items-center gap-2">
                    <span>절감</span>
                    <span className="font-mono font-black">₩{(priceItems.reduce((acc, i) => acc + (i.costPrice || 0), 0) - priceItems.reduce((acc, i) => acc + getRoundedValue(i.unitPrice || 0, i.useRounding), 0)).toLocaleString()}</span>
                  </div>
                </div>
                <div className="text-[10px] font-black text-slate-300 uppercase tracking-[0.2em] italic">
                   MS SYSTEM DASHBOARD v1.2
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
              <Building2 className="h-20 w-20 text-slate-200 mb-6" />
              <h2 className="text-xl font-bold text-slate-900 mb-2">업체를 선택해 주세요</h2>
              <p className="text-slate-500 text-sm max-w-xs mx-auto">
                왼쪽 리스트에서 관리할 업체를 선택하시면 해당 업체의 전용 단가표를 조회할 수 있습니다.
              </p>
            </div>
          )}
        </main>
      </div>

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
                      <p className="text-[10px] text-slate-500 font-medium tracking-tight">5원 이상 올림, 4원 이하 절사</p>
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

        {selectedMatrixRow && (
          <Modal 
            key="modal-price-chart" 
            title={`${selectedMatrixRow.itemName} 단가 비교 차트`} 
            onClose={() => setSelectedMatrixRow(null)}
            maxWidth="max-w-4xl"
          >
            <div className="p-4 bg-white rounded-xl">
              <div className="mb-8 flex flex-col gap-2">
                <div className="flex items-center gap-2">
                  <span className="text-[10px] font-black bg-indigo-100 text-indigo-700 px-2 py-1 rounded uppercase tracking-[0.2em]">{selectedMatrixRow.category}</span>
                  <span className="text-xl font-black text-slate-800 tracking-tight">{selectedMatrixRow.itemName}</span>
                </div>
                <p className="text-xs font-bold text-slate-400">규격: {selectedMatrixRow.spec} | 단위: {selectedMatrixRow.unit}</p>
              </div>
              
              <div className="h-[400px] w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData} margin={{ top: 20, right: 30, left: 40, bottom: 80 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis 
                      dataKey="name" 
                      angle={-45} 
                      textAnchor="end" 
                      interval={0} 
                      height={80} 
                      tick={{ fontSize: 10, fontWeight: 900, fill: '#64748b' }}
                      axisLine={{ stroke: '#f1f5f9' }}
                    />
                    <YAxis 
                      tick={{ fontSize: 10, fontWeight: 900, fill: '#64748b' }}
                      tickFormatter={(value) => `₩${value.toLocaleString()}`}
                      axisLine={{ stroke: '#f1f5f9' }}
                    />
                    <RechartsTooltip 
                      cursor={{ fill: '#f8fafc' }}
                      content={({ active, payload }) => {
                        if (active && payload && payload.length) {
                          return (
                            <div className="bg-slate-900 text-white p-4 shadow-2xl rounded-2xl border border-slate-800 backdrop-blur-sm bg-opacity-95">
                              <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 border-b border-slate-800 pb-2">{payload[0].payload.name}</p>
                              <div className="flex items-baseline gap-1">
                                <span className="text-lg font-black text-indigo-400">₩</span>
                                <span className="text-xl font-black">{Number(payload[0].value).toLocaleString()}</span>
                              </div>
                            </div>
                          );
                        }
                        return null;
                      }}
                    />
                    <Bar dataKey="price" radius={[8, 8, 0, 0]} barSize={50} animationDuration={1500} animationEasing="ease-out">
                      {chartData.map((entry: any, index: number) => (
                        <Cell 
                          key={`cell-${index}`} 
                          fill={index === 0 ? '#10b981' : '#6366f1'} 
                          fillOpacity={0.9}
                        />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="mt-10 flex justify-end">
                <button 
                  onClick={() => setSelectedMatrixRow(null)}
                  className="px-8 py-3 bg-slate-900 text-white text-xs font-black uppercase tracking-widest rounded-2xl hover:bg-slate-800 hover:shadow-xl transition-all"
                >
                  차트 닫기
                </button>
              </div>
            </div>
          </Modal>
        )}

        {isViewingPriceTable && selectedVendor && (
          <Modal 
            key="modal-price-table-view" 
            title={`${selectedVendor.name} - 단가표 원본`} 
            onClose={() => setIsViewingPriceTable(false)}
            maxWidth="max-w-6xl"
          >
            <div className="flex flex-col gap-4">
              <div className="flex items-center justify-between bg-slate-50 p-4 rounded-xl border border-slate-100">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-white rounded-lg shadow-sm">
                    {selectedVendor.priceTableFileType === 'excel' ? <Table className="h-5 w-5 text-emerald-500" /> : 
                     selectedVendor.priceTableFileType === 'pdf' ? <FileText className="h-5 w-5 text-rose-500" /> : 
                     <Building2 className="h-5 w-5 text-indigo-500" />}
                  </div>
                  <div>
                    <p className="text-sm font-bold text-slate-800">{selectedVendor.priceTableFileName || '단가표 파일'}</p>
                    <p className="text-[10px] text-slate-400 font-medium uppercase tracking-wider">{selectedVendor.priceTableFileType} FORMAT</p>
                  </div>
                </div>
                <a 
                  href={selectedVendor.priceTableUrl} 
                  target="_blank" 
                  rel="noreferrer"
                  className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-600 hover:bg-slate-50 hover:border-indigo-300 hover:text-indigo-600 transition-all shadow-sm"
                >
                  <Download className="h-3.5 w-3.5" />
                  파일 다운로드
                </a>
              </div>

              <div className="bg-slate-50 rounded-2xl border-2 border-slate-100 p-2 min-h-[400px] flex items-center justify-center overflow-auto">
                {selectedVendor.priceTableFileType === 'image' ? (
                  <img 
                    src={selectedVendor.priceTableUrl} 
                    alt="Price Table" 
                    className="max-w-full h-auto rounded-lg shadow-xl"
                  />
                ) : selectedVendor.priceTableFileType === 'pdf' ? (
                  <iframe 
                    src={`${selectedVendor.priceTableUrl}#toolbar=0`} 
                    className="w-full h-[700px] rounded-lg border border-slate-200 bg-white"
                    title="PDF Viewer"
                  />
                ) : selectedVendor.priceTableFileType === 'excel' ? (
                  excelPreviewData ? (
                    <div className="w-full h-[600px] bg-white rounded-lg shadow-inner overflow-auto border border-slate-200">
                      <table className="w-full text-xs text-left border-collapse">
                        <thead className="sticky top-0 bg-slate-800 text-white z-10">
                          <tr>
                            {excelPreviewData[0] && Array.isArray(excelPreviewData[0]) && excelPreviewData[0].map((cell: any, i: number) => (
                              <th key={i} className="px-3 py-2 font-bold border-r border-slate-700">{cell || `Col ${i+1}`}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {excelPreviewData.slice(1).map((row, i) => (
                            <tr key={i} className={i % 2 === 0 ? 'bg-white' : 'bg-slate-50'}>
                              {Array.isArray(row) ? row.map((cell: any, j: number) => (
                                <td key={j} className="px-3 py-2 border-r border-slate-100 whitespace-nowrap">{cell}</td>
                              )) : (
                                <td colSpan={100} className="px-3 py-2 text-slate-400 italic">데이터를 불러올 수 없습니다.</td>
                              )}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center gap-4 py-20">
                      <Loader2 className="h-8 w-8 animate-spin text-indigo-500" />
                      <p className="text-sm font-bold text-slate-400 italic">엑셀 파일을 분석하고 있습니다...</p>
                    </div>
                  )
                ) : (
                  <div className="flex flex-col items-center gap-4 py-20">
                    <FileText className="h-12 w-12 text-slate-200" />
                    <p className="text-sm font-bold text-slate-400 italic">미리보기를 지원하지 않는 파일 형식입니다.</p>
                    <a 
                      href={selectedVendor.priceTableUrl} 
                      target="_blank" 
                      rel="noreferrer"
                      className="text-xs text-indigo-600 font-bold underline"
                    >
                      직접 다운로드하여 확인하기
                    </a>
                  </div>
                )}
              </div>
            </div>
          </Modal>
        )}

        {isAddingItem && (
          <Modal key="modal-add-item" title="신규 품목 단가 등록" onClose={() => setIsAddingItem(false)}>
            <form onSubmit={handleAddItem} className="space-y-8">
              <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">카테고리 (Category)</label>
                    <select 
                      name="category"
                      className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all outline-none appearance-none"
                    >
                      <option value="">카테고리 선택</option>
                      {categories.filter(c => c !== '전체').map(cat => (
                        <option key={`opt-add-item-${cat}`} value={cat}>{cat}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">품번 (Item Code)</label>
                    <input name="itemCode" placeholder="예) MS-VL-001" className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" />
                  </div>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">품명 (Item Name) *</label>
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
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">
                    {addItemNegoType === 'sts_pipe' ? '단중 (Weight: KG/M) *' : '협가 (Negotiated Price) *'}
                  </label>
                   <div className="relative">
                    <span className="absolute left-4 top-1/2 -translate-y-1/2 text-lg font-bold text-slate-300">
                      {addItemNegoType === 'sts_pipe' ? 'kg' : '₩'}
                    </span>
                    <input 
                      name="costPrice" 
                      type="text" 
                      inputMode="decimal"
                      required 
                      placeholder="0" 
                      onChange={(e) => {
                        const val = e.target.value.replace(/,/g, '');
                        if (val === '' || !isNaN(Number(val))) {
                          e.target.value = val === '' ? '' : Number(val).toLocaleString();
                        } else {
                          e.target.value = e.target.value.replace(/[^0-9.]/g, '');
                        }
                      }}
                      className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 pl-10 font-black text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" 
                    />
                  </div>
                </div>
                <div className="space-y-4">
                  <div className="space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-indigo-600">
                        {addItemNegoType === 'sts_pipe' ? 'KG 단가 (Price per KG) *' : '네고 방식 & 율/금액/방식 *'}
                      </label>
                      <select 
                        name="negoType" 
                        value={addItemNegoType}
                        onChange={(e) => setAddItemNegoType(e.target.value as any)}
                        className="text-[10px] font-black bg-indigo-100 text-indigo-700 border-none rounded-lg px-2 py-1 outline-none cursor-pointer"
                      >
                        <option value="percent">할인율 (%)</option>
                        <option value="sts_pipe">STS PIPE (단중식)</option>
                      </select>
                    </div>
                    <div className="relative">
                      {addItemNegoType === 'sts_pipe' && <span className="absolute left-4 top-1/2 -translate-y-1/2 text-lg font-bold text-indigo-200">₩</span>}
                    <input 
                      name="negoRate" 
                      type="text" 
                      inputMode="decimal"
                      defaultValue="0" 
                      onChange={(e) => {
                        const val = e.target.value.replace(/,/g, '');
                        if (val === '' || !isNaN(Number(val))) {
                          e.target.value = val === '' ? '' : Number(val).toLocaleString();
                        } else {
                          e.target.value = e.target.value.replace(/[^0-9.]/g, '');
                        }
                      }}
                      className={`w-full rounded-2xl border-2 border-indigo-100 bg-indigo-50/30 p-4 font-black text-indigo-700 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all ${addItemNegoType === 'sts_pipe' ? 'pl-10' : ''}`} 
                    />
                    </div>
                  </div>
                  {addItemNegoType === 'sts_pipe' && (
                    <p className="text-[9px] text-indigo-400 font-medium">※ STS PIPE: 단중(협가란에 입력) × KG단가(네고란에 입력)로 구매단가가 자동계산됩니다.</p>
                  )}
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
                <div className="col-span-full">
                  <label className="flex items-center gap-3 cursor-pointer p-4 bg-slate-50 rounded-2xl border-2 border-slate-100 hover:bg-white hover:border-indigo-200 transition-all">
                    <input type="checkbox" name="itemRounding" defaultChecked className="w-5 h-5 rounded-lg border-2 border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                    <div>
                      <p className="text-sm font-bold text-slate-800">단가 10원 단위 반올림 적용</p>
                      <p className="text-[10px] text-slate-500 font-medium tracking-tight">현재 품목에 대해 10원 단위 반올림(5UP 4DOWN)을 적용합니다.</p>
                    </div>
                  </label>
                </div>
              </div>
              <button type="submit" className="w-full rounded-3xl bg-slate-900 py-5 text-xl font-black text-white shadow-2xl transition-all hover:bg-slate-800 hover:scale-[1.02]">
                단가 데이터 업데이트
              </button>
            </form>
          </Modal>
        )}

        {isDeleteModalOpen && vendorToDelete && (
          <Modal key="modal-delete-vendor-confirm" title="업체 삭제 최종 확인" onClose={() => setIsDeleteModalOpen(false)}>
            <form onSubmit={handleConfirmDelete} className="space-y-6">
              <div className="bg-rose-50 border border-rose-100 p-6 rounded-2xl">
                <div className="flex items-center gap-3 text-rose-600 mb-4">
                  <Trash2 className="h-6 w-6" />
                  <h3 className="text-lg font-black tracking-tight">이 작업은 취소할 수 없습니다.</h3>
                </div>
                <p className="text-sm text-rose-700 leading-relaxed font-bold">
                  ' <span className="underline decoration-rose-300 underline-offset-4">{vendorToDelete.name}</span> ' 업체와 관련된 모든 데이터(단가표, 설정 등)가 <span className="bg-rose-100 px-1">영구적으로 삭제</span>됩니다.
                </p>
              </div>
              
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">관리자 비밀번호를 한 번 더 입력해주세요</label>
                <input 
                  type="password" 
                  value={deletePasswordInput}
                  onChange={(e) => setDeletePasswordInput(e.target.value)}
                  placeholder="관리자 비밀번호"
                  autoFocus
                  required
                  className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-rose-500 focus:bg-white focus:outline-none transition-all" 
                />
              </div>

              <div className="flex gap-4">
                <button 
                  type="button"
                  onClick={() => setIsDeleteModalOpen(false)}
                  className="flex-1 py-4 bg-slate-100 text-slate-600 font-bold rounded-2xl hover:bg-slate-200 transition-all"
                >
                  취소
                </button>
                <button 
                  type="submit" 
                  className="flex-1 py-4 bg-rose-600 text-white font-black rounded-2xl hover:bg-rose-700 transition-all shadow-xl shadow-rose-100"
                >
                  업체 데이터 영구 삭제
                </button>
              </div>
            </form>
          </Modal>
        )}

        {isViewingDeletedVendors && (
          <Modal 
            key="modal-view-deleted-vendors" 
            title="삭제된 거래처 복구" 
            onClose={() => setIsViewingDeletedVendors(false)}
            maxWidth="max-w-xl"
          >
            <div className="space-y-4">
              <div className="bg-amber-50 border border-amber-100 p-4 rounded-xl mb-4">
                <p className="text-xs text-amber-800 font-medium leading-relaxed">
                  거래처를 삭제(소프트 삭제)한 경우 이곳에서 확인할 수 있습니다.<br/>
                  복구 버튼을 누르면 다시 목록에 표시되며, 기존 모든 데이터가 유지됩니다.
                </p>
              </div>

              {vendors.filter(v => v.deleted).length === 0 ? (
                <div className="py-20 text-center bg-slate-50 rounded-2xl border-2 border-dashed border-slate-100">
                  <History className="h-10 w-10 text-slate-300 mx-auto mb-3" />
                  <p className="text-slate-400 font-bold">최근 삭제된 거래처가 없습니다.</p>
                </div>
              ) : (
                <div className="max-h-[500px] overflow-y-auto pr-2 space-y-2">
                  {vendors.filter(v => v.deleted).sort((a,b) => (b.deletedAt?.toMillis?.() || 0) - (a.deletedAt?.toMillis?.() || 0)).map(vendor => (
                    <div key={vendor.id} className="flex items-center justify-between p-4 bg-white border border-slate-100 rounded-xl hover:border-indigo-200 transition-all group">
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-400 group-hover:bg-indigo-50 group-hover:text-indigo-600 transition-colors">
                          <Building2 className="h-5 w-5" />
                        </div>
                        <div>
                          <p className="font-bold text-slate-800 text-sm">{vendor.name}</p>
                          <p className="text-[10px] text-slate-400 font-medium">
                            삭제일: {vendor.deletedAt?.toDate ? vendor.deletedAt.toDate().toLocaleString('ko-KR') : '알 수 없음'}
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => handleRestoreVendor(vendor.id)}
                        className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg text-xs font-bold hover:bg-slate-900 transition-all shadow-lg shadow-indigo-100"
                      >
                        <RotateCcw className="h-3.5 w-3.5" />
                        복구하기
                      </button>
                    </div>
                  ))}
                </div>
              )}

              <button 
                onClick={() => setIsViewingDeletedVendors(false)}
                className="w-full py-4 mt-4 bg-slate-100 text-slate-600 font-bold rounded-xl hover:bg-slate-200 transition-all"
              >
                닫기
              </button>
            </div>
          </Modal>
        )}

        {isEditingVendorInfo && selectedVendor && (
          <Modal key="modal-edit-vendor" title="업체 정보 수정" onClose={() => setIsEditingVendorInfo(false)}>
            <form onSubmit={handleUpdateVendorInfo} className="space-y-8">
              <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                <div className="space-y-2 col-span-full text-center py-4 bg-slate-50 rounded-2xl border-2 border-dashed border-slate-100 flex flex-col items-center justify-center">
                   <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">업체 고유 링크</p>
                   <code className="text-xs text-indigo-500 font-mono bg-white px-3 py-1 rounded-full border border-indigo-50">{window.location.origin.replace('ais-dev-', 'ais-pre-')}/?v={selectedVendor.id}</code>
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
                    <input type="checkbox" name="useRounding" defaultChecked={selectedVendor.roundingMethod === 'round'} className="w-5 h-5 rounded-lg border-2 border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                    <div>
                      <p className="text-sm font-bold text-slate-800">단가 10원 단위 반올림/절사</p>
                      <p className="text-[10px] text-slate-500 font-medium tracking-tight">5원 이상 올림, 4원 이하 절사</p>
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
              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">기존 비밀번호 확인</label>
                  <input 
                    name="currentPassword"
                    type="password" 
                    required
                    placeholder="기존 비밀번호 입력"
                    className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 text-center text-xl font-black text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">새 비밀번호 (최소 4자)</label>
                  <input 
                    name="newPassword"
                    type="password" 
                    required
                    placeholder="새로운 비밀번호 입력"
                    className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 text-center text-xl font-black text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all" 
                  />
                </div>
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
                className={`border-2 border-dashed rounded-3xl p-10 text-center transition-all cursor-pointer group ${
                  isDragging ? 'border-indigo-500 bg-indigo-50 scale-[1.02]' : 'border-slate-200 bg-slate-50 hover:border-indigo-400 hover:bg-white'
                }`}
                onClick={() => document.getElementById('price-items-bulk-upload')?.click()}
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDragging(true);
                }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsDragging(false);
                  const file = e.dataTransfer.files?.[0];
                  if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
                    setBulkItemFile(file);
                  }
                }}
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
                  (최대 300개 등록 가능, 컬럼명: 업체명, 대표자, 이메일, 전화번호, 팩스번호, 사업자번호, 비밀번호, 취급품목)
                </p>
              </div>

              <div 
                className={`border-2 border-dashed rounded-2xl p-8 text-center transition-all cursor-pointer ${
                  isDragging ? 'border-indigo-500 bg-indigo-50 scale-[1.02]' : 'border-slate-200 bg-slate-50 hover:border-indigo-400'
                }`}
                onClick={() => document.getElementById('bulk-excel-upload')?.click()}
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDragging(true);
                }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsDragging(false);
                  const file = e.dataTransfer.files?.[0];
                  if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
                    setBulkFile(file);
                  }
                }}
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

        {/* ADMIN PASSWORD CHANGE MODAL */}
        <AnimatePresence>
          {isChangingAdminPassword && (
            <Modal key="modal-change-admin-password" title="관리자 비밀번호 변경" onClose={() => setIsChangingAdminPassword(false)}>
              <div className="space-y-4">
                <div className="bg-amber-50 border border-amber-200 p-3 rounded-xl flex items-start gap-3">
                  <Info className="h-5 w-5 text-amber-500 shrink-0 mt-0.5" />
                  <p className="text-xs text-amber-700 leading-relaxed font-medium">
                    새로운 비밀번호로 변경됩니다. 변경 즉시 전체 시스템에 적용되니 신중히 입력해주세요.
                  </p>
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase tracking-wider">새 비밀번호</label>
                  <input 
                    type="password"
                    value={newAdminPassword}
                    onChange={(e) => setNewAdminPassword(e.target.value)}
                    placeholder="새로운 관리자 비밀번호 입력"
                    className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all font-bold placeholder:font-medium"
                  />
                </div>
                <div className="flex gap-3 pt-2">
                  <button 
                    onClick={() => setIsChangingAdminPassword(false)}
                    className="flex-1 px-4 py-3 bg-slate-100 text-slate-600 font-bold rounded-2xl hover:bg-slate-200 transition-all"
                  >
                    취소
                  </button>
                  <button 
                    onClick={handleUpdateAdminPassword}
                    className="flex-[2] px-4 py-3 bg-indigo-600 text-white font-bold rounded-2xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200"
                  >
                    비밀번호 변경하기
                  </button>
                </div>
              </div>
            </Modal>
          )}
        </AnimatePresence>

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

        {/* Bulk Price Adjustment Modal */}
        {isBulkAdjustModalOpen && (
          <Modal key="modal-bulk-adjust" title="가격 일괄 조정" onClose={() => setIsBulkAdjustModalOpen(false)}>
            <div className="space-y-6 py-4">
              <div className="p-4 bg-indigo-50 rounded-2xl border border-indigo-100 mb-6">
                <p className="text-sm font-bold text-indigo-900 mb-1">{selectedPriceIds.size}개의 항목이 선택되었습니다.</p>
                <p className="text-xs text-indigo-600">선택한 모든 항목의 가격을 일괄적으로 조정합니다.</p>
              </div>

              <div className="space-y-4">
                <label className="block">
                  <span className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400 mb-2 block">조정 방식: 네고율 추가 (%)</span>
                </label>

                <label className="block">
                  <span className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400 mb-2 block">
                    네고율 변동치
                  </span>
                  <div className="relative">
                    <input 
                      type="text"
                      inputMode="decimal"
                      value={bulkAdjustValue.toLocaleString()}
                      onChange={(e) => {
                        const val = e.target.value.replace(/,/g, '');
                        if (val === '' || !isNaN(Number(val))) {
                          setBulkAdjustValue(val === '' ? 0 : Number(val));
                        }
                      }}
                      placeholder="예: 5 (5% 추가 네고)"
                      autoFocus
                      className="w-full rounded-2xl border-2 border-slate-100 bg-slate-50 p-4 font-bold text-slate-800 focus:border-indigo-500 focus:bg-white focus:outline-none transition-all pr-12 text-lg"
                    />
                    <span className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 font-bold">
                      %
                    </span>
                  </div>
                  <p className="mt-2 text-[10px] text-slate-500 font-medium">
                    기존 네고율에 입력한 수치를 더합니다. (예: 5 입력 시 기존 10% {"->"} 15%)
                  </p>
                </label>
              </div>

              <div className="flex gap-3 pt-6 border-t border-slate-100">
                <button 
                  onClick={() => setIsBulkAdjustModalOpen(false)}
                  className="flex-1 rounded-2xl border-2 border-slate-100 p-4 font-black transition-all hover:bg-slate-50 text-slate-400"
                >
                  취소
                </button>
                <button 
                  onClick={handleBulkPriceAdjust}
                  className="flex-1 rounded-2xl bg-indigo-600 p-4 font-black text-white shadow-xl shadow-indigo-200 transition-all hover:bg-indigo-700 active:scale-95"
                >
                  일괄 적용하기
                </button>
              </div>
            </div>
          </Modal>
        )}

        {/* Editing Price Item Modal */}
        {editingPriceId && (
          <Modal 
            key="modal-edit-price" 
            title="품목 정보 수정" 
            onClose={() => setEditingPriceId(null)}
          >
            {(() => {
              const item = priceItems.find(p => p.id === editingPriceId);
              if (!item) return <p className="text-slate-400 p-8 text-center">항목을 찾을 수 없습니다.</p>;
              
              return (
                <form 
                  onSubmit={(e) => {
                    e.preventDefault();
                    const formData = new FormData(e.currentTarget);
                    handleUpdatePriceItem(editingPriceId, {
                      itemName: formData.get('itemName') as string,
                      itemCode: formData.get('itemCode') as string,
                      category: formData.get('category') as string,
                      spec: formData.get('spec') as string,
                      unit: formData.get('unit') as string,
                      costPrice: Number(String(formData.get('costPrice') || '0').replace(/,/g, '')),
                      negoRate: Number(String(formData.get('negoRate') || '0').replace(/,/g, '')),
                      negoType: formData.get('negoType') as 'percent' | 'sts_pipe',
                      weight: Number(String(formData.get('weight') || '0').replace(/,/g, '')),
                      useRounding: formData.get('useRounding') === 'on',
                      maker: formData.get('maker') as string,
                      remarks: formData.get('remarks') as string,
                    });
                  }}
                  className="space-y-6"
                >
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2 col-span-full">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">품목명 *</label>
                      <input name="itemName" defaultValue={item.itemName} required className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-bold focus:border-indigo-500 outline-none" />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">품번</label>
                      <input name="itemCode" defaultValue={item.itemCode} className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-bold focus:border-indigo-500 outline-none" />
                    </div>
                    <div className="space-y-2">
                       <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">카테고리</label>
                       <select name="category" defaultValue={item.category} className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-bold focus:border-indigo-500 outline-none appearance-none">
                         <option value="">선택 안함</option>
                         {categories.filter(c => c !== '전체').map(cat => <option key={`opt-edit-item-${cat}`} value={cat}>{cat}</option>)}
                       </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">규격</label>
                      <input name="spec" defaultValue={item.spec} className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-bold focus:border-indigo-500 outline-none" />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">단위</label>
                      <input name="unit" defaultValue={item.unit} className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-bold focus:border-indigo-500 outline-none" />
                    </div>
                    
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">협가 / 단중 *</label>
                      <input 
                        name="costPrice" 
                        type="text" 
                        inputMode="decimal"
                        defaultValue={item.costPrice?.toLocaleString()} 
                        required 
                        onChange={(e) => {
                          const val = e.target.value.replace(/,/g, '');
                          if (val === '' || !isNaN(Number(val))) {
                            e.target.value = val === '' ? '' : Number(val).toLocaleString();
                          } else {
                            e.target.value = e.target.value.replace(/[^0-9.]/g, '');
                          }
                        }}
                        className="w-full rounded-xl border-2 border-indigo-50 bg-indigo-50/10 p-4 font-bold text-indigo-700 focus:border-indigo-500 outline-none" 
                      />
                      <p className="text-[9px] text-slate-400">※ STS 방식일 경우 단중을, 아닐 경우 협가(원)를 입력하세요.</p>
                    </div>
                    
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-indigo-600">네고 상세 방식 & 값 *</label>
                      <div className="flex gap-2">
                        <select name="negoType" defaultValue={item.negoType || 'percent'} className="rounded-xl border-2 border-indigo-100 bg-indigo-50/30 p-4 font-bold text-indigo-700 outline-none">
                          <option value="percent">네고율 (%)</option>
                          <option value="sts_pipe">STS (KG단가)</option>
                        </select>
                        <input 
                          name="negoRate" 
                          type="text" 
                          inputMode="decimal"
                          defaultValue={item.negoRate?.toLocaleString()} 
                          required 
                          onChange={(e) => {
                            const val = e.target.value.replace(/,/g, '');
                            if (val === '' || !isNaN(Number(val))) {
                              e.target.value = val === '' ? '' : Number(val).toLocaleString();
                            } else {
                              e.target.value = e.target.value.replace(/[^0-9.]/g, '');
                            }
                          }}
                          className="flex-1 rounded-xl border-2 border-indigo-100 bg-indigo-50/30 p-4 font-bold text-indigo-700 focus:border-indigo-500 outline-none" 
                        />
                      </div>
                    </div>

                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">참조 단중 (KG/M)</label>
                      <input 
                        name="weight" 
                        type="text" 
                        inputMode="decimal"
                        defaultValue={item.weight?.toLocaleString()} 
                        onChange={(e) => {
                          const val = e.target.value.replace(/,/g, '');
                          if (val === '' || !isNaN(Number(val))) {
                            e.target.value = val === '' ? '' : Number(val).toLocaleString();
                          } else {
                            e.target.value = e.target.value.replace(/[^0-9.]/g, '');
                          }
                        }}
                        className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-medium focus:border-indigo-500 outline-none" 
                      />
                    </div>

                    <div className="space-y-2 flex items-center pt-6">
                      <label className="flex items-center gap-3 cursor-pointer">
                        <input type="checkbox" name="useRounding" defaultChecked={item.useRounding !== undefined ? item.useRounding : (selectedVendor?.roundingMethod !== 'none')} className="w-5 h-5 rounded border-slate-300 text-indigo-600" />
                        <span className="text-sm font-bold text-slate-700">10원 단위 반올림 적용</span>
                      </label>
                    </div>

                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">제조사</label>
                      <input name="maker" defaultValue={item.maker} className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-medium focus:border-indigo-500 outline-none" />
                    </div>
                    <div className="space-y-2 col-span-full">
                      <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">비고</label>
                      <textarea name="remarks" defaultValue={item.remarks} rows={2} className="w-full rounded-xl border-2 border-slate-100 bg-slate-50 p-4 font-medium focus:border-indigo-500 outline-none"></textarea>
                    </div>
                  </div>
                  
                  <div className="flex gap-3 pt-6 border-t border-slate-100">
                    <button type="button" onClick={() => setEditingPriceId(null)} className="flex-1 p-4 rounded-xl border-2 border-slate-100 font-bold text-slate-400 hover:bg-slate-50">취소</button>
                    <button type="submit" className="flex-1 p-4 rounded-xl bg-slate-900 text-white font-bold hover:bg-slate-800 shadow-xl shadow-slate-200">수정 내역 저장</button>
                  </div>
                </form>
              );
            })()}
          </Modal>
        )}

        {/* Approvals Modal */}
        {isApprovalsModalOpen && (
          <Modal 
            key="modal-approvals-queue"
            title="상세 수정 승인 대기 목록" 
            onClose={() => setIsApprovalsModalOpen(false)}
            maxWidth="max-w-4xl"
          >
            <div className="space-y-4">
              {pendingUpdates.filter(u => u.status === 'pending').length === 0 ? (
                <div className="p-12 text-center">
                  <p className="text-slate-400 font-medium text-sm">승인 대기 중인 항목이 없습니다.</p>
                </div>
              ) : (
                <div className="space-y-4">
                  <div className="flex items-center justify-between bg-slate-50 border border-slate-200 p-4 rounded-2xl">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-xl bg-indigo-100 flex items-center justify-center">
                        <FileText className="h-5 w-5 text-indigo-600" />
                      </div>
                      <div>
                        <h4 className="text-sm font-black text-slate-800">전체 승인 대기 내역</h4>
                        <p className="text-[10px] text-slate-500 font-bold">총 {pendingUpdates.filter(u => u.status === 'pending').length}건의 요청이 있습니다.</p>
                      </div>
                    </div>
                    <button 
                      onClick={handleBulkApproveUpdates}
                      className="px-6 py-2.5 bg-indigo-600 text-white text-xs font-black rounded-xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 flex items-center gap-2"
                    >
                      <Save className="h-3.5 w-3.5" />
                      일괄 승인 적용
                    </button>
                  </div>

                  <div className="space-y-3">
                    {pendingUpdates.filter(u => u.status === 'pending').map((update) => {
                      const vendor = vendors.find(v => v.id === update.vendorId);
                      return (
                        <div key={update.id} className="p-5 border border-slate-100 rounded-2xl bg-white shadow-sm hover:shadow-md transition-all">
                          <div className="flex items-start justify-between mb-4">
                            <div>
                              <div className="flex items-center gap-2 mb-1">
                                <span className="text-[10px] font-black bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded-full uppercase tracking-wider">
                                  {vendor?.name || '알 수 없는 업체'}
                                </span>
                                <span className="text-[10px] text-slate-400">
                                  {new Date(update.requestedAt?.toDate?.() || Date.now()).toLocaleString()}
                                </span>
                              </div>
                              <h4 className="text-sm font-black text-slate-800">{update.itemName}</h4>
                              {update.spec && (
                                <p className="text-[10px] text-slate-500 font-bold mt-0.5">{update.spec}</p>
                              )}
                            </div>
                            <div className="flex gap-2">
                              <button 
                                onClick={() => handleRejectUpdate(update.id)}
                                className="px-4 py-2 bg-rose-50 text-rose-600 text-xs font-bold rounded-xl hover:bg-rose-100 transition-colors"
                              >
                                거절
                              </button>
                              <button 
                                onClick={() => handleApproveUpdate(update)}
                                className="px-4 py-2 bg-indigo-600 text-white text-xs font-bold rounded-xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100"
                              >
                                승인 적용
                              </button>
                            </div>
                          </div>
                          
                          <div className="grid grid-cols-2 gap-4 bg-slate-50/50 p-4 rounded-xl border border-dashed border-slate-200">
                            <div className="space-y-2">
                              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">변경 전 (기존)</span>
                              <div className="space-y-1">
                                {Object.entries(update.oldData).map(([key, val]) => (
                                  <div key={key} className="flex items-center justify-between text-[11px] font-mono text-slate-500">
                                    <span className="opacity-60">{key}:</span>
                                    <span className="font-black text-slate-600">{val?.toLocaleString()}</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                            <div className="space-y-2 border-l border-slate-200 pl-4">
                              <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-widest">변경 후 (요청)</span>
                              <div className="space-y-1">
                                {Object.entries(update.newData).map(([key, val]) => (
                                  <div key={key} className="flex items-center justify-between text-[11px] font-mono text-indigo-600">
                                    <span className="opacity-60">{key}:</span>
                                    <span className="font-black">{val?.toLocaleString()}</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          </Modal>
        )}

        {/* Global Toast */}
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
    </div>
  );
}
