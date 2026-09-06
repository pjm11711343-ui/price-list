import React, { useState, useMemo } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  TrendingUp, 
  TrendingDown, 
  Calendar, 
  ChevronDown, 
  ChevronUp, 
  ArrowRight, 
  Filter, 
  BarChart3, 
  Copy, 
  Check, 
  AlertCircle,
  Clock,
  Sparkles,
  Layers,
  ArrowUpDown,
  Building2,
  ExternalLink
} from 'lucide-react';
import { PriceItem, PriceFluctuationItem, FluctuationPeriodType, FluctuationSortMode } from '../types';

interface TopFluctuationReportProps {
  allPrices: PriceItem[];
  matrixCategory: string;
  searchTerm?: string;
  onSelectRow?: (itemName: string, spec: string) => void;
}

// Robust timestamp parser supporting Date, Firestore Timestamp, ISO string, epoch ms
function parseTimestamp(val: any): Date | null {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'object') {
    if (typeof val.toDate === 'function') {
      try {
        const d = val.toDate();
        if (d instanceof Date && !isNaN(d.getTime())) return d;
      } catch {
        // continue
      }
    }
    if (typeof val.seconds === 'number') {
      return new Date(val.seconds * 1000 + (val.nanoseconds ? val.nanoseconds / 1000000 : 0));
    }
  }
  if (typeof val === 'string' || typeof val === 'number') {
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function formatDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

export const TopFluctuationReport: React.FC<TopFluctuationReportProps> = ({
  allPrices,
  matrixCategory,
  searchTerm = '',
  onSelectRow
}) => {
  const [isExpanded, setIsExpanded] = useState<boolean>(true);
  const [periodType, setPeriodType] = useState<FluctuationPeriodType>('3m');
  const [sortMode, setSortMode] = useState<FluctuationSortMode>('absAmount');
  const [useCategoryFilter, setUseCategoryFilter] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);

  // Custom date range state (defaulting to last 90 days)
  const [customStartDate, setCustomStartDate] = useState<string>(() => {
    const d = new Date();
    d.setDate(d.getDate() - 90);
    return formatDate(d);
  });
  const [customEndDate, setCustomEndDate] = useState<string>(() => {
    return formatDate(new Date());
  });

  // Calculate effective start and end dates based on periodType
  const { startDate, endDate, periodLabel } = useMemo(() => {
    const end = new Date();
    end.setHours(23, 59, 59, 999);

    let start = new Date();
    start.setHours(0, 0, 0, 0);

    let label = '';

    switch (periodType) {
      case '1m':
        start.setDate(start.getDate() - 30);
        label = '최근 1개월 (30일)';
        break;
      case '3m':
        start.setDate(start.getDate() - 90);
        label = '최근 3개월 (90일)';
        break;
      case '6m':
        start.setDate(start.getDate() - 180);
        label = '최근 6개월 (180일)';
        break;
      case '1y':
        start.setDate(start.getDate() - 365);
        label = '최근 1년 (365일)';
        break;
      case 'all':
        start = new Date(2020, 0, 1);
        label = '전체 기간';
        break;
      case 'custom':
        if (customStartDate) {
          const s = new Date(customStartDate);
          if (!isNaN(s.getTime())) {
            s.setHours(0, 0, 0, 0);
            start = s;
          }
        }
        if (customEndDate) {
          const e = new Date(customEndDate);
          if (!isNaN(e.getTime())) {
            e.setHours(23, 59, 59, 999);
            end.setTime(e.getTime());
          }
        }
        label = `${formatDate(start)} ~ ${formatDate(end)}`;
        break;
    }

    return { startDate: start, endDate: end, periodLabel: label };
  }, [periodType, customStartDate, customEndDate]);

  // Compute fluctuation metrics for all items and extract Top 5
  const { topItems, allFluctuatingItems, stats } = useMemo(() => {
    if (!allPrices || allPrices.length === 0) {
      return {
        topItems: [] as PriceFluctuationItem[],
        allFluctuatingItems: [] as PriceFluctuationItem[],
        stats: { totalItems: 0, changedCount: 0, avgDiff: 0, maxIncreaseItem: null, maxDecreaseItem: null }
      };
    }

    // 1. Group all price records by product key (${itemName}_${spec})
    // Each product key will track candidate fluctuations across its vendor offerings
    const productMap = new Map<string, {
      itemName: string;
      spec: string;
      unit: string;
      category: string;
      candidates: {
        vendorId: string;
        vendorName: string;
        startPrice: number;
        endPrice: number;
        priceDiff: number;
        absDiff: number;
        percentDiff: number;
        direction: 'up' | 'down' | 'same';
        minPrice: number;
        maxPrice: number;
        lastUpdated?: Date | null;
        historyTimeline: { price: number; date: any; label?: string }[];
        hasHistory: boolean;
      }[];
      allVendors: Set<string>;
    }>();

    const startEpoch = startDate.getTime();
    const endEpoch = endDate.getTime();

    // Process all prices
    allPrices.forEach((priceItem) => {
      // Check search filter if active
      if (searchTerm.trim().length > 0) {
        const term = searchTerm.toLowerCase();
        const matches = (priceItem.itemName || '').toLowerCase().includes(term) ||
          (priceItem.spec || '').toLowerCase().includes(term) ||
          (priceItem.vendorName || '').toLowerCase().includes(term);
        if (!matches) return;
      }

      // Check category filter if enabled
      if (useCategoryFilter && matrixCategory && matrixCategory !== '전체') {
        if (priceItem.category !== matrixCategory) return;
      }

      const pKey = `${(priceItem.itemName || '').trim()}_${(priceItem.spec || '').trim()}`;
      if (!productMap.has(pKey)) {
        productMap.set(pKey, {
          itemName: priceItem.itemName || '미지정 품목',
          spec: priceItem.spec || '-',
          unit: priceItem.unit || 'EA',
          category: priceItem.category || '기타',
          candidates: [],
          allVendors: new Set()
        });
      }

      const prod = productMap.get(pKey)!;
      if (priceItem.vendorName) {
        prod.allVendors.add(priceItem.vendorName);
      }

      // Build timeline points for this vendor item
      const points: { price: number; date: Date; source: string }[] = [];

      // 1) Current unitPrice
      const currentPrice = priceItem.unitPrice || 0;
      const updatedDate = parseTimestamp(priceItem.updatedAt) || parseTimestamp(priceItem.createdAt) || new Date();
      if (currentPrice > 0) {
        points.push({ price: currentPrice, date: updatedDate, source: 'current' });
      }

      // 2) priceHistory array: [{ price, date }]
      const histList = priceItem.priceHistory || [];
      histList.forEach((h: any, hIdx: number) => {
        const hPrice = typeof h.price === 'number' ? h.price : Number(h.price);
        const hDate = parseTimestamp(h.date);
        if (!isNaN(hPrice) && hPrice > 0 && hDate) {
          points.push({ price: hPrice, date: hDate, source: `history-${hIdx}` });
        }
      });

      // 3) baseUnitPrice
      if (priceItem.baseUnitPrice && priceItem.baseUnitPrice > 0 && priceItem.baseUnitPrice !== currentPrice) {
        const createdDate = parseTimestamp(priceItem.createdAt) || new Date(updatedDate.getTime() - 120 * 24 * 60 * 60 * 1000);
        points.push({ price: priceItem.baseUnitPrice, date: createdDate, source: 'base' });
      }

      // Deduplicate points with same price and date within 60 seconds
      points.sort((a, b) => a.date.getTime() - b.date.getTime());
      const uniquePoints: { price: number; date: Date; source: string }[] = [];
      points.forEach(pt => {
        if (uniquePoints.length === 0) {
          uniquePoints.push(pt);
        } else {
          const last = uniquePoints[uniquePoints.length - 1];
          const timeDiff = Math.abs(pt.date.getTime() - last.date.getTime());
          if (timeDiff > 60000 || pt.price !== last.price) {
            uniquePoints.push(pt);
          }
        }
      });

      // Now calculate price fluctuation for this vendor item in [startDate, endDate]
      let startP = 0;
      let endP = 0;
      let minP = Infinity;
      let maxP = -Infinity;
      let changeDetected = false;

      if (periodType === 'all') {
        if (uniquePoints.length >= 2) {
          startP = uniquePoints[0].price;
          endP = uniquePoints[uniquePoints.length - 1].price;
          changeDetected = startP !== endP;
        } else if (priceItem.baseUnitPrice && priceItem.baseUnitPrice !== currentPrice && priceItem.baseUnitPrice > 0) {
          startP = priceItem.baseUnitPrice;
          endP = currentPrice;
          changeDetected = true;
        }
      } else {
        // Date range bounded
        // Points that occurred before or at startDate
        const beforeOrAtStart = uniquePoints.filter(p => p.date.getTime() <= startEpoch);
        // Points that occurred strictly within (startDate, endDate]
        const withinPeriod = uniquePoints.filter(p => p.date.getTime() >= startEpoch && p.date.getTime() <= endEpoch);
        // Points on or before endDate
        const beforeOrAtEnd = uniquePoints.filter(p => p.date.getTime() <= endEpoch);

        if (beforeOrAtStart.length > 0) {
          startP = beforeOrAtStart[beforeOrAtStart.length - 1].price;
        } else if (withinPeriod.length > 0) {
          startP = withinPeriod[0].price;
        }

        if (beforeOrAtEnd.length > 0) {
          endP = beforeOrAtEnd[beforeOrAtEnd.length - 1].price;
        } else if (withinPeriod.length > 0) {
          endP = withinPeriod[withinPeriod.length - 1].price;
        }

        // Did any price change occur within this period?
        if (withinPeriod.length > 0) {
          // If there are recorded updates in this period
          if (startP > 0 && endP > 0 && startP !== endP) {
            changeDetected = true;
          } else if (withinPeriod.some(p => p.price !== startP)) {
            changeDetected = true;
          }
        }

        // If no explicit history records fell strictly in the dates, but updatedAt was in period and baseUnitPrice differs
        if (!changeDetected && priceItem.baseUnitPrice && priceItem.baseUnitPrice !== currentPrice && priceItem.baseUnitPrice > 0) {
          const itemUpdateEpoch = updatedDate.getTime();
          if (itemUpdateEpoch >= startEpoch && itemUpdateEpoch <= endEpoch) {
            startP = priceItem.baseUnitPrice;
            endP = currentPrice;
            changeDetected = true;
          }
        }
      }

      // Calculate min and max
      uniquePoints.forEach(p => {
        if (p.price < minP) minP = p.price;
        if (p.price > maxP) maxP = p.price;
      });
      if (minP === Infinity) minP = currentPrice;
      if (maxP === -Infinity) maxP = currentPrice;

      // If change was detected and prices are valid
      if (changeDetected && startP > 0 && endP > 0 && startP !== endP) {
        const pDiff = endP - startP;
        const absD = Math.abs(pDiff);
        const pctD = (pDiff / startP) * 100;

        prod.candidates.push({
          vendorId: priceItem.vendorId,
          vendorName: priceItem.vendorName || '기본 거래처',
          startPrice: startP,
          endPrice: endP,
          priceDiff: pDiff,
          absDiff: absD,
          percentDiff: pctD,
          direction: pDiff > 0 ? 'up' : 'down',
          minPrice: minP,
          maxPrice: maxP,
          lastUpdated: updatedDate,
          historyTimeline: uniquePoints.map(p => ({ price: p.price, date: p.date })),
          hasHistory: uniquePoints.length > 1
        });
      }
    });

    // 2. Aggregate candidates per product: choose the vendor with the highest fluctuation
    const list: PriceFluctuationItem[] = [];

    productMap.forEach((prod, key) => {
      if (prod.candidates.length === 0) return;

      // Pick top candidate by absDiff
      prod.candidates.sort((a, b) => b.absDiff - a.absDiff);
      const topCand = prod.candidates[0];

      list.push({
        key,
        itemName: prod.itemName,
        spec: prod.spec,
        unit: prod.unit,
        category: prod.category,
        primaryVendorId: topCand.vendorId,
        primaryVendorName: topCand.vendorName,
        allVendorNames: Array.from(prod.allVendors),
        startPrice: topCand.startPrice,
        endPrice: topCand.endPrice,
        priceDiff: topCand.priceDiff,
        absDiff: topCand.absDiff,
        percentDiff: topCand.percentDiff,
        direction: topCand.direction,
        minPrice: topCand.minPrice,
        maxPrice: topCand.maxPrice,
        lastUpdated: topCand.lastUpdated,
        changeCount: prod.candidates.length,
        historyTimeline: topCand.historyTimeline
      });
    });

    // 3. Sort according to sortMode
    const sorted = [...list].sort((a, b) => {
      switch (sortMode) {
        case 'absAmount':
          return b.absDiff - a.absDiff;
        case 'rate':
          return Math.abs(b.percentDiff) - Math.abs(a.percentDiff);
        case 'increase':
          // Sort positive changes descending, then negative
          return b.priceDiff - a.priceDiff;
        case 'decrease':
          // Sort negative changes ascending (largest drops first)
          return a.priceDiff - b.priceDiff;
        default:
          return b.absDiff - a.absDiff;
      }
    });

    // 4. Calculate summary stats
    const totalItems = productMap.size;
    const changedCount = sorted.length;
    const totalAbsDiff = sorted.reduce((sum, item) => sum + item.absDiff, 0);
    const avgDiff = changedCount > 0 ? Math.round(totalAbsDiff / changedCount) : 0;

    let maxInc: PriceFluctuationItem | null = null;
    let maxDec: PriceFluctuationItem | null = null;

    sorted.forEach(item => {
      if (item.priceDiff > 0) {
        if (!maxInc || item.priceDiff > maxInc.priceDiff) maxInc = item;
      } else if (item.priceDiff < 0) {
        if (!maxDec || item.priceDiff < maxDec.priceDiff) maxDec = item;
      }
    });

    return {
      topItems: sorted.slice(0, 5),
      allFluctuatingItems: sorted,
      stats: {
        totalItems,
        changedCount,
        avgDiff,
        maxIncreaseItem: maxInc,
        maxDecreaseItem: maxDec
      }
    };
  }, [allPrices, searchTerm, matrixCategory, useCategoryFilter, periodType, startDate, endDate, sortMode]);

  // Copy report summary to clipboard
  const handleCopyReport = () => {
    if (topItems.length === 0) return;

    let text = `[명신기공 단가 통합 비교 - 주요 변동 품목 리포트 TOP 5]\n`;
    text += `조회 기간: ${periodLabel} (${formatDate(startDate)} ~ ${formatDate(endDate)})\n`;
    text += `카테고리: ${useCategoryFilter ? matrixCategory : '전체'}\n`;
    text += `정렬 기준: ${
      sortMode === 'absAmount' ? '변동 금액순' :
      sortMode === 'rate' ? '변동률순' :
      sortMode === 'increase' ? '단가 급등순(인상)' : '단가 급락순(인하)'
    }\n`;
    text += `--------------------------------------------------\n`;

    topItems.forEach((item, idx) => {
      const sign = item.priceDiff > 0 ? '+' : '';
      text += `${idx + 1}위: [${item.category}] ${item.itemName} (${item.spec}) - ${item.primaryVendorName}\n`;
      text += `    단가: ${item.startPrice.toLocaleString()}원 -> ${item.endPrice.toLocaleString()}원 (${sign}${item.priceDiff.toLocaleString()}원 / ${sign}${item.percentDiff.toFixed(1)}%)\n`;
    });

    text += `--------------------------------------------------\n`;
    text += `* 변동 품목 수: ${stats.changedCount}개 / 평균 변동폭: ${stats.avgDiff.toLocaleString()}원`;

    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <section 
      id="top-fluctuation-report-container"
      className="bg-white border-b border-slate-200/80 shadow-xs transition-all select-none"
    >
      {/* 1. COMPACT / EXPANDABLE HEADER */}
      <div className="px-6 py-3.5 flex flex-wrap items-center justify-between gap-3 bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-xl bg-rose-500/20 border border-rose-400/30 flex items-center justify-center text-rose-400 shadow-inner">
            <TrendingUp className="h-4 w-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-black tracking-tight text-white flex items-center gap-1.5">
                주요 변동 품목 리포트
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-rose-500/30 text-rose-300 border border-rose-400/40 font-black tracking-wider">
                  TOP 5
                </span>
              </h2>
              <span className="text-[11px] text-slate-300 font-medium hidden sm:inline">
                | {periodLabel}
              </span>
            </div>
            <p className="text-[10.5px] text-slate-400 font-normal">
              선택한 기간 동안 단가 변동폭(인상/인하)이 가장 큰 상위 5개 품목을 실시간 추출합니다.
            </p>
          </div>
        </div>

        {/* Right Header Controls */}
        <div className="flex items-center gap-2">
          {/* Copy Report Button */}
          {topItems.length > 0 && (
            <button
              onClick={handleCopyReport}
              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-slate-200 text-[11px] font-bold transition-all border border-white/10"
              title="리포트 요약 클립보드 복사"
            >
              {copied ? (
                <>
                  <Check className="h-3 w-3 text-emerald-400" />
                  <span className="text-emerald-300">복사완료</span>
                </>
              ) : (
                <>
                  <Copy className="h-3 w-3 text-slate-300" />
                  <span>요약 복사</span>
                </>
              )}
            </button>
          )}

          {/* Expand/Collapse Toggle */}
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-indigo-600/60 hover:bg-indigo-600 text-white text-[11px] font-bold transition-all border border-indigo-400/30 shadow-xs"
          >
            <span>{isExpanded ? '접기' : '리포트 펼치기'}</span>
            {isExpanded ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
          </button>
        </div>
      </div>

      {/* 2. REPORT BODY CONTENT */}
      <AnimatePresence>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2 }}
            className="overflow-hidden bg-slate-50/50"
          >
            {/* FILTER & PERIOD SELECTOR TOOLBAR */}
            <div className="px-6 py-3 border-b border-slate-200/60 bg-white flex flex-wrap items-center justify-between gap-3 text-xs">
              {/* Period Selector Buttons */}
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-[11px] font-black text-slate-500 mr-1 flex items-center gap-1">
                  <Calendar className="h-3.5 w-3.5 text-indigo-600" />
                  기간 선택:
                </span>
                {(['1m', '3m', '6m', '1y', 'all', 'custom'] as FluctuationPeriodType[]).map((type) => {
                  const labelMap: Record<FluctuationPeriodType, string> = {
                    '1m': '최근 1개월',
                    '3m': '최근 3개월',
                    '6m': '최근 6개월',
                    '1y': '최근 1년',
                    'all': '전체 기간',
                    'custom': '직접 지정'
                  };
                  const active = periodType === type;
                  return (
                    <button
                      key={`period-btn-${type}`}
                      onClick={() => setPeriodType(type)}
                      className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all ${
                        active
                          ? 'bg-slate-900 text-white shadow-xs'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      {labelMap[type]}
                    </button>
                  );
                })}

                {/* Custom Date Pickers */}
                {periodType === 'custom' && (
                  <div className="flex items-center gap-1.5 ml-2 bg-slate-50 px-2 py-0.5 rounded-lg border border-slate-200">
                    <input
                      type="date"
                      value={customStartDate}
                      onChange={(e) => setCustomStartDate(e.target.value)}
                      className="bg-transparent text-[11px] font-bold text-slate-700 focus:outline-none"
                    />
                    <span className="text-slate-400 font-bold">~</span>
                    <input
                      type="date"
                      value={customEndDate}
                      onChange={(e) => setCustomEndDate(e.target.value)}
                      className="bg-transparent text-[11px] font-bold text-slate-700 focus:outline-none"
                    />
                  </div>
                )}
              </div>

              {/* Sort Mode Selector */}
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-[11px] font-black text-slate-500 mr-1 flex items-center gap-1">
                  <ArrowUpDown className="h-3.5 w-3.5 text-indigo-600" />
                  정렬 기준:
                </span>
                <button
                  onClick={() => setSortMode('absAmount')}
                  className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all ${
                    sortMode === 'absAmount'
                      ? 'bg-indigo-600 text-white shadow-xs'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}
                  title="인상/인하 구분 없이 단가 변동 금액이 가장 큰 순서"
                >
                  변동폭 순(|Δ₩|)
                </button>
                <button
                  onClick={() => setSortMode('rate')}
                  className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all ${
                    sortMode === 'rate'
                      ? 'bg-indigo-600 text-white shadow-xs'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}
                  title="변동 비율(%)이 가장 큰 순서"
                >
                  변동률 순(%)
                </button>
                <button
                  onClick={() => setSortMode('increase')}
                  className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all flex items-center gap-1 ${
                    sortMode === 'increase'
                      ? 'bg-rose-600 text-white shadow-xs'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}
                  title="단가가 가장 많이 오른 급등 품목 순"
                >
                  <TrendingUp className="h-3 w-3 text-rose-400" />
                  급등 순(▲)
                </button>
                <button
                  onClick={() => setSortMode('decrease')}
                  className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all flex items-center gap-1 ${
                    sortMode === 'decrease'
                      ? 'bg-blue-600 text-white shadow-xs'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}
                  title="단가가 가장 많이 내린 급락 품목 순"
                >
                  <TrendingDown className="h-3 w-3 text-blue-300" />
                  급락 순(▼)
                </button>

                {/* Category Scope Toggle */}
                {matrixCategory && matrixCategory !== '전체' && (
                  <button
                    onClick={() => setUseCategoryFilter(!useCategoryFilter)}
                    className={`ml-2 px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all border ${
                      useCategoryFilter
                        ? 'bg-amber-50 border-amber-300 text-amber-800'
                        : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'
                    }`}
                  >
                    현재 카테고리만({matrixCategory})
                  </button>
                )}
              </div>
            </div>

            {/* KEY STATS SUMMARY STRIP */}
            <div className="px-6 py-2.5 bg-slate-100/70 border-b border-slate-200/50 flex flex-wrap items-center justify-between gap-3 text-[11px]">
              <div className="flex flex-wrap items-center gap-4">
                <span className="font-bold text-slate-600">
                  분석 대상 품목: <strong className="text-slate-900">{stats.totalItems.toLocaleString()}</strong>개
                </span>
                <span className="font-bold text-slate-600">
                  기간 내 변동 발생: <strong className="text-indigo-600 font-black">{stats.changedCount.toLocaleString()}</strong>개
                </span>
                {stats.changedCount > 0 && (
                  <span className="font-bold text-slate-600">
                    평균 변동폭: <strong className="text-slate-900">{stats.avgDiff.toLocaleString()}</strong>원
                  </span>
                )}
              </div>

              {/* Quick Extremes Pills */}
              <div className="flex items-center gap-3">
                {stats.maxIncreaseItem && (
                  <div className="flex items-center gap-1.5 text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded-md font-bold text-[10.5px]">
                    <TrendingUp className="h-3 w-3" />
                    <span>최대 인상:</span>
                    <span className="font-black truncate max-w-[120px]">{stats.maxIncreaseItem.itemName}</span>
                    <span className="font-mono font-black">+{stats.maxIncreaseItem.priceDiff.toLocaleString()}원</span>
                  </div>
                )}
                {stats.maxDecreaseItem && (
                  <div className="flex items-center gap-1.5 text-blue-700 bg-blue-50 border border-blue-200 px-2 py-0.5 rounded-md font-bold text-[10.5px]">
                    <TrendingDown className="h-3 w-3" />
                    <span>최대 인하:</span>
                    <span className="font-black truncate max-w-[120px]">{stats.maxDecreaseItem.itemName}</span>
                    <span className="font-mono font-black">{stats.maxDecreaseItem.priceDiff.toLocaleString()}원</span>
                  </div>
                )}
              </div>
            </div>

            {/* 3. TOP 5 CARDS GRID */}
            <div className="p-6">
              {topItems.length === 0 ? (
                <div className="bg-white border border-slate-200 rounded-2xl p-8 text-center shadow-xs">
                  <div className="w-12 h-12 bg-slate-100 text-slate-400 rounded-full flex items-center justify-center mx-auto mb-3">
                    <AlertCircle className="h-6 w-6" />
                  </div>
                  <h4 className="text-sm font-bold text-slate-800 mb-1">
                    선택한 기간 ({periodLabel}) 내에 단가 변동이 기록된 품목이 없습니다.
                  </h4>
                  <p className="text-xs text-slate-500 mb-4 max-w-md mx-auto">
                    단가 변동 이력(priceHistory)이 있는 품목을 확인하시려면 분석 기간을 더 넓히거나 '전체 기간'을 선택해주세요.
                  </p>
                  <div className="flex justify-center gap-2">
                    <button
                      onClick={() => setPeriodType('6m')}
                      className="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition-colors"
                    >
                      최근 6개월로 조회
                    </button>
                    <button
                      onClick={() => setPeriodType('1y')}
                      className="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition-colors"
                    >
                      최근 1년으로 조회
                    </button>
                    <button
                      onClick={() => setPeriodType('all')}
                      className="px-3.5 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs transition-colors shadow-xs"
                    >
                      전체 기간으로 조회
                    </button>
                  </div>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3.5">
                  {topItems.map((item, idx) => {
                    const isUp = item.priceDiff > 0;
                    const rankColors = [
                      'from-amber-500 to-amber-600 text-white shadow-amber-200', // 1st
                      'from-slate-600 to-slate-700 text-white shadow-slate-200', // 2nd
                      'from-amber-700 to-amber-800 text-white shadow-amber-300', // 3rd
                      'from-indigo-600 to-indigo-700 text-white shadow-indigo-200', // 4th
                      'from-indigo-800 to-slate-800 text-white shadow-slate-200', // 5th
                    ];

                    return (
                      <div
                        key={`top-fluctuation-${item.key}-${idx}`}
                        className="bg-white border border-slate-200/90 rounded-2xl p-4 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group relative overflow-hidden"
                      >
                        {/* Top Indicator bar */}
                        <div 
                          className={`absolute top-0 left-0 right-0 h-1 ${
                            isUp ? 'bg-gradient-to-r from-rose-500 to-red-500' : 'bg-gradient-to-r from-blue-500 to-indigo-500'
                          }`} 
                        />

                        {/* Card Header: Rank & Category */}
                        <div>
                          <div className="flex items-center justify-between mb-2">
                            <span className={`px-2 py-0.5 rounded-md text-[10px] font-black bg-gradient-to-r shadow-xs ${rankColors[idx] || rankColors[3]}`}>
                              TOP {idx + 1}
                            </span>
                            <span className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-md border border-indigo-100">
                              {item.category}
                            </span>
                          </div>

                          {/* Item Name & Spec */}
                          <h3 className="text-xs font-black text-slate-900 group-hover:text-indigo-600 transition-colors line-clamp-1" title={item.itemName}>
                            {item.itemName}
                          </h3>
                          <div className="flex items-center gap-1.5 text-[11px] text-slate-500 font-medium mt-0.5 mb-2.5">
                            <span className="truncate max-w-[120px]" title={item.spec}>{item.spec || '-'}</span>
                            <span className="text-slate-300">|</span>
                            <span className="text-slate-400 font-bold">{item.unit || 'EA'}</span>
                          </div>

                          {/* Primary Vendor Badge */}
                          <div className="flex items-center gap-1.5 py-1 px-2 rounded-lg bg-slate-50 border border-slate-100 text-[10px] text-slate-600 font-bold mb-3">
                            <Building2 className="h-3 w-3 text-slate-400 shrink-0" />
                            <span className="truncate font-semibold" title={item.primaryVendorName}>
                              {item.primaryVendorName}
                            </span>
                            {item.allVendorNames.length > 1 && (
                              <span className="text-[9px] text-indigo-600 bg-indigo-50 px-1 py-0.2 rounded font-black shrink-0">
                                외 {item.allVendorNames.length - 1}곳
                              </span>
                            )}
                          </div>
                        </div>

                        {/* Price Change Figures */}
                        <div className="pt-2 border-t border-slate-100">
                          {/* Price Movement: Start -> End */}
                          <div className="flex items-center justify-between text-[10.5px] text-slate-400 font-medium mb-1">
                            <span className="line-through decoration-slate-300">
                              {item.startPrice.toLocaleString()}원
                            </span>
                            <ArrowRight className="h-3 w-3 text-slate-300" />
                            <span className="text-slate-900 font-black text-xs">
                              {item.endPrice.toLocaleString()}원
                            </span>
                          </div>

                          {/* Big Fluctuation Pill */}
                          <div 
                            className={`px-2.5 py-1.5 rounded-xl border flex items-center justify-between font-mono font-black ${
                              isUp 
                                ? 'bg-rose-50/80 border-rose-200 text-rose-700' 
                                : 'bg-blue-50/80 border-blue-200 text-blue-700'
                            }`}
                          >
                            <div className="flex items-center gap-1 text-[11px]">
                              {isUp ? <TrendingUp className="h-3.5 w-3.5 text-rose-600" /> : <TrendingDown className="h-3.5 w-3.5 text-blue-600" />}
                              <span>{isUp ? '+' : ''}{item.priceDiff.toLocaleString()}원</span>
                            </div>
                            <span className="text-[10px] px-1.5 py-0.5 rounded-md bg-white/80 border border-current shadow-2xs">
                              {isUp ? '▲' : '▼'} {Math.abs(item.percentDiff).toFixed(1)}%
                            </span>
                          </div>

                          {/* Action Button: Jump to Matrix Row */}
                          {onSelectRow && (
                            <button
                              type="button"
                              onClick={() => onSelectRow(item.itemName, item.spec)}
                              className="w-full mt-2.5 py-1.5 px-2 rounded-lg bg-slate-50 hover:bg-indigo-50 hover:text-indigo-600 text-slate-500 font-bold text-[10.5px] flex items-center justify-center gap-1 transition-all border border-slate-100 hover:border-indigo-200 group/btn"
                            >
                              <span>매트릭스 위치 확인</span>
                              <ArrowRight className="h-3 w-3 group-hover/btn:translate-x-0.5 transition-transform" />
                            </button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </section>
  );
};
