"use client";

import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";


// --- Types ---

interface Prepayment {
  month: number;
  amount: number;
}

interface InterestChange {
  month: number;
  rate: number;
}

interface TabData {
  id: string;
  name: string;
  principal: number;
  interestRate: number;
  tenureYears: number;
  startMonth: string;
  prepayments: Prepayment[];
  interestChanges: InterestChange[];
}

interface AmortizationRow {
  monthIndex: number;
  monthLabel: string;
  emi: number;
  principalPaid: number;
  interestPaid: number;
  prepayment: number;
  remainingPrincipal: number;
  interestRate: number;
}

interface AmortizationResult {
  schedule: AmortizationRow[];
  totalInterest: number;
  totalAmount: number;
  totalSavings: number;
  originalTotalInterest: number;
  monthsSaved: number;
}

// --- Utils ---

const CURRENCIES = [
  { code: "INR", symbol: "₹", name: "Indian Rupee", locale: "en-IN" },
  { code: "USD", symbol: "$", name: "US Dollar", locale: "en-US" },
  { code: "EUR", symbol: "€", name: "Euro", locale: "en-DE" },
  { code: "GBP", symbol: "£", name: "British Pound", locale: "en-GB" },
  { code: "AED", symbol: "د.إ", name: "UAE Dirham", locale: "ar-AE" },
];

const formatCurrency = (amount: number, locale: string = "en-IN", decimals: number = 0) => {
  return new Intl.NumberFormat(locale, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(amount);
};

const getMonthLabel = (startMonth: string, offset: number) => {
  const date = new Date(startMonth + "-01");
  date.setMonth(date.getMonth() + offset);
  return date.toLocaleDateString("en-US", { month: "short", year: "numeric" });
};

const calculateAmortization = (data: TabData): AmortizationResult => {
  const {
    principal,
    interestRate,
    tenureYears,
    startMonth,
    prepayments,
    interestChanges,
  } = data;

  const schedule: AmortizationRow[] = [];
  let remainingPrincipal = principal;
  let currentInterestRate = interestRate;

  const totalMonthsOriginal = tenureYears * 12;
  const monthlyRateOriginal = interestRate / 12 / 100;
  const originalEMI =
    (principal *
      monthlyRateOriginal *
      Math.pow(1 + monthlyRateOriginal, totalMonthsOriginal)) /
    (Math.pow(1 + monthlyRateOriginal, totalMonthsOriginal) - 1);

  let currentEMI = originalEMI;
  let totalInterest = 0;
  let monthIndex = 0;

  const originalTotalInterest = originalEMI * totalMonthsOriginal - principal;

  while (remainingPrincipal > 0 && monthIndex < 600) {
    const monthLabel = getMonthLabel(startMonth, monthIndex);
    const rateChange = interestChanges.find((c) => c.month === monthIndex);
    if (rateChange) {
      currentInterestRate = rateChange.rate;
    }

    const monthlyRate = currentInterestRate / 12 / 100;
    const interestForMonth = remainingPrincipal * monthlyRate;

    if (currentEMI <= interestForMonth) {
      currentEMI = interestForMonth + principal / totalMonthsOriginal;
    }

    let principalForMonth = currentEMI - interestForMonth;
    const prepaymentData = prepayments.find((p) => p.month === monthIndex);
    const prepayment = prepaymentData ? prepaymentData.amount : 0;

    if (principalForMonth + prepayment > remainingPrincipal) {
      principalForMonth = remainingPrincipal;
      remainingPrincipal = 0;
    } else {
      remainingPrincipal -= principalForMonth + prepayment;
    }

    totalInterest += interestForMonth;

    schedule.push({
      monthIndex,
      monthLabel,
      emi: interestForMonth + principalForMonth,
      principalPaid: principalForMonth,
      interestPaid: interestForMonth,
      prepayment,
      remainingPrincipal,
      interestRate: currentInterestRate,
    });

    if (remainingPrincipal <= 0) break;
    monthIndex++;
  }

  return {
    schedule,
    totalInterest,
    totalAmount: principal + totalInterest,
    totalSavings: originalTotalInterest - totalInterest,
    originalTotalInterest,
    monthsSaved: totalMonthsOriginal - schedule.length,
  };
};

// --- Components ---

export default function ExcelTaxCalculator() {
  const [tabs, setTabs] = useState<TabData[]>([
    {
      id: "1",
      name: "Sheet1",
      principal: 5000000,
      interestRate: 8.5,
      tenureYears: 20,
      startMonth: new Date().toISOString().slice(0, 7),
      prepayments: [],
      interestChanges: [],
    },
  ]);

  const [activeTabId, setActiveTabId] = useState("1");
  const [currency, setCurrency] = useState(CURRENCIES[0]);
  const [editingTabId, setEditingTabId] = useState<string | null>(null);
  const [selection, setSelection] = useState<{
    start: { r: number; c: string };
    end: { r: number; c: string };
  } | null>(null);
  const [isSelecting, setIsSelecting] = useState(false);
  const [scrollTop, setScrollTop] = useState(0);
  const [balanceHoverIdx, setBalanceHoverIdx] = useState<number | null>(null);
  const [emiHoverIdx, setEmiHoverIdx] = useState<number | null>(null);
  const [isLoaded, setIsLoaded] = useState(false);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const tabsContainerRef = React.useRef<HTMLDivElement>(null);

  // Load from localStorage on mount
  useEffect(() => {
    const savedTabs = localStorage.getItem("tax-calculator-tabs");
    const savedActiveTabId = localStorage.getItem("tax-calculator-activeTabId");
    const savedCurrencyCode = localStorage.getItem("tax-calculator-currencyCode");

    if (savedTabs) {
      try {
        setTabs(JSON.parse(savedTabs));
      } catch (e) {
        console.error("Failed to parse saved tabs", e);
      }
    }

    if (savedActiveTabId) {
      setActiveTabId(savedActiveTabId);
    }

    if (savedCurrencyCode) {
      const found = CURRENCIES.find((c) => c.code === savedCurrencyCode);
      if (found) setCurrency(found);
    }

    setIsLoaded(true);
  }, []);

  // Autosave to localStorage on any change
  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem("tax-calculator-tabs", JSON.stringify(tabs));
      localStorage.setItem("tax-calculator-activeTabId", activeTabId);
      localStorage.setItem("tax-calculator-currencyCode", currency.code);
    }
  }, [tabs, activeTabId, currency, isLoaded]);

  const activeTab = tabs.find((t) => t.id === activeTabId) || tabs[0];
  const amortization = useMemo(
    () => calculateAmortization(activeTab),
    [activeTab],
  );

  const ROW_HEIGHT = 32; // Standard Excel row height
  const VISIBLE_ROWS = 30;
  const BUFFER_ROWS = 10;

  const COLUMNS = ["", "A", "B", "C", "D", "E", "F", "G", "H"];
  const colIdx = (c: string) => COLUMNS.indexOf(c);

  const isCellSelected = (r: number, c: string) => {
    if (!selection) return false;
    const minR = Math.min(selection.start.r, selection.end.r);
    const maxR = Math.max(selection.start.r, selection.end.r);
    const minC = Math.min(colIdx(selection.start.c), colIdx(selection.end.c));
    const maxC = Math.max(colIdx(selection.start.c), colIdx(selection.end.c));
    const currC = colIdx(c);
    return r >= minR && r <= maxR && currC >= minC && currC <= maxC;
  };

  const isRowSelected = (r: number) => {
    if (!selection) return false;
    const minR = Math.min(selection.start.r, selection.end.r);
    const maxR = Math.max(selection.start.r, selection.end.r);
    return r >= minR && r <= maxR;
  };

  const isColSelected = (c: string) => {
    if (!selection) return false;
    const minC = Math.min(colIdx(selection.start.c), colIdx(selection.end.c));
    const maxC = Math.max(colIdx(selection.start.c), colIdx(selection.end.c));
    const currC = colIdx(c);
    return currC >= minC && currC <= maxC;
  };

  const getSelectionStats = () => {
    if (!selection) return null;
    const values: number[] = [];

    const minR = Math.min(selection.start.r, selection.end.r);
    const maxR = Math.max(selection.start.r, selection.end.r);
    const minC = Math.min(colIdx(selection.start.c), colIdx(selection.end.c));
    const maxC = Math.max(colIdx(selection.start.c), colIdx(selection.end.c));

    for (let r = minR; r <= maxR; r++) {
      for (let c_idx = minC; c_idx <= maxC; c_idx++) {
        const c = COLUMNS[c_idx];
        let val: any = null;

        // Map grid cells to data
        if (r === 2 && c === "B") val = activeTab.principal;
        else if (r === 3 && c === "B") val = activeTab.interestRate;
        else if (r === 4 && c === "B") val = activeTab.tenureYears;
        else if (r >= 8) {
          const row = amortization.schedule[r - 8];
          if (row) {
            if (c === "B") val = row.emi;
            else if (c === "C") val = row.interestPaid;
            else if (c === "D") val = row.principalPaid;
            else if (c === "E") val = row.prepayment;
            else if (c === "F") val = row.interestRate;
            else if (c === "G") val = row.remainingPrincipal;
          }
        }

        if (typeof val === "number" && !isNaN(val)) values.push(val);
      }
    }

    if (values.length === 0) return null;
    const sum = values.reduce((a, b) => a + b, 0);
    return {
      sum,
      avg: sum / values.length,
      count: values.length,
    };
  };

  const stats = useMemo(getSelectionStats, [
    selection,
    activeTab,
    amortization,
  ]);

  const handleMouseDown = (e: React.MouseEvent, r: number, c: string) => {
    if (e.shiftKey && selection) {
      setSelection({ ...selection, end: { r, c } });
    } else {
      setSelection({ start: { r, c }, end: { r, c } });
    }
    setIsSelecting(true);
  };

  const handleMouseEnter = (r: number, c: string) => {
    if (isSelecting && selection) {
      setSelection({ ...selection, end: { r, c } });
    }
  };

  const handleMouseUp = () => {
    setIsSelecting(false);
  };

  useEffect(() => {
    window.addEventListener("mouseup", handleMouseUp);
    return () => window.removeEventListener("mouseup", handleMouseUp);
  }, []);

  const handleScroll = (e: React.UIEvent<HTMLDivElement>) => {
    setScrollTop(e.currentTarget.scrollTop);
  };

  const virtualizedData = useMemo(() => {
    // We only virtualize the schedule, not the inputs
    // Input rows are 1-5 (approx 5 rows * ROW_HEIGHT)
    const tableHeaderPos = 180; // Approximate height of inputs + margin
    const scrollAdjusted = Math.max(0, scrollTop - tableHeaderPos);

    const startIndex = Math.max(
      0,
      Math.floor(scrollAdjusted / ROW_HEIGHT) - BUFFER_ROWS,
    );
    const endIndex = Math.min(
      amortization.schedule.length,
      startIndex + VISIBLE_ROWS + BUFFER_ROWS * 2,
    );

    return {
      rows: amortization.schedule.slice(startIndex, endIndex),
      startIndex,
      totalHeight: amortization.schedule.length * ROW_HEIGHT,
      topOffset: startIndex * ROW_HEIGHT,
    };
  }, [scrollTop, amortization.schedule]);

  const getSelectionBoxStyle = () => {
    if (!selection) return { display: "none" };

    const minR = Math.min(selection.start.r, selection.end.r);
    const maxR = Math.max(selection.start.r, selection.end.r);
    const minC = Math.min(colIdx(selection.start.c), colIdx(selection.end.c));
    const maxC = Math.max(colIdx(selection.start.c), colIdx(selection.end.c));

    const cellWidth = 128;
    const cellHeight = 32;
    const rowNumWidth = 40;
    const columnHeaderHeight = 25;

    // With rows 1-4, then 5-6 (empty), then 7, 8... all 32px
    // The logic becomes a simple linear calculation
    const top = columnHeaderHeight + (minR - 1) * cellHeight;
    const left = rowNumWidth + (minC - 1) * cellWidth;

    const width = (maxC - minC + 1) * cellWidth + 1;
    const height = (maxR - minR + 1) * cellHeight + 1;

    return {
      position: "absolute" as const,
      top,
      left,
      width,
      height,
      border: "2px solid #107c41",
      backgroundColor: "rgba(16, 124, 65, 0.05)",
      pointerEvents: "none" as const,
      zIndex: 100,
      transition: "none",
    };
  };

  const updateActiveTab = (updates: Partial<TabData>) => {
    setTabs(tabs.map((t) => (t.id === activeTabId ? { ...t, ...updates } : t)));
  };

  const addTab = () => {
    const newId = Math.random().toString(36).substr(2, 9);
    const newName = `Sheet${tabs.length + 1}`;
    setTabs([
      ...tabs,
      {
        ...activeTab,
        id: newId,
        name: newName,
        prepayments: [],
        interestChanges: [],
      },
    ]);
    setActiveTabId(newId);

    // Scroll to the new tab after state update
    setTimeout(() => {
      if (tabsContainerRef.current) {
        tabsContainerRef.current.scrollTo({
          left: tabsContainerRef.current.scrollWidth,
          behavior: "smooth",
        });
      }
    }, 50);
  };

  const scrollTabs = (direction: "left" | "right") => {
    if (tabsContainerRef.current) {
      const scrollAmount = 200;
      tabsContainerRef.current.scrollBy({
        left: direction === "left" ? -scrollAmount : scrollAmount,
        behavior: "smooth",
      });
    }
  };

  const removeTab = (id: string) => {
    if (tabs.length > 1) {
      const newTabs = tabs.filter((t) => t.id !== id);
      setTabs(newTabs);
      if (activeTabId === id) setActiveTabId(newTabs[0].id);
    }
  };

  const handlePrepaymentChange = (month: number, val: string) => {
    // Strip everything except digits and decimal point to be robust against symbols/commas
    const cleaned = val.replace(/[^\d.]/g, "");
    const amount = cleaned === "" ? 0 : Number(cleaned);
    if (isNaN(amount)) return;
    const existing = activeTab.prepayments.filter((p) => p.month !== month);
    updateActiveTab({
      prepayments: amount > 0 ? [...existing, { month, amount }] : existing,
    });
  };

  const handleRateChange = (month: number, val: string) => {
    const rate = val === "" ? activeTab.interestRate : Number(val);
    if (isNaN(rate)) return;
    // Remove the current override and any subsequent overrides to propagate the new rate
    const existing = activeTab.interestChanges.filter((c) => c.month < month);
    updateActiveTab({ interestChanges: [...existing, { month, rate }] });
  };

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();

    tabs.forEach((tab) => {
      const tabAmortization = calculateAmortization(tab);
      const ws_data = [
        ["LOAN FINANCIAL REPORT", "", "", "", "", "", `Currency: ${currency.code}`],
        ["Sheet Name", tab.name, "", "", "", "", ""],
        [],
        ["LOAN INPUTS", "", "", "SUMMARY METRICS", "", "", ""],
        ["Principal", tab.principal, "", "Total Paid", tabAmortization.totalAmount, "", ""],
        ["Interest Rate (%)", tab.interestRate, "", "Total Interest", tabAmortization.totalInterest, "", ""],
        ["Interest Rate (%)", tab.interestRate, "", tabAmortization.totalSavings < 0 ? "Interest Added" : "Interest Saved", Math.abs(tabAmortization.totalSavings), "", ""],
        ["Tenure (Years)", tab.tenureYears, "", tabAmortization.monthsSaved < 0 ? "Months Added" : "Months Saved", Math.abs(tabAmortization.monthsSaved), "", ""],
        ["Start Month", tab.startMonth, "", "", "", "", ""],
        [],
        ["AMORTIZATION SCHEDULE", "", "", "", "", "", ""],
        ["Month", "EMI", "Interest", "Principal", "Prepayment", "Rate %", "Balance"],
        ...tabAmortization.schedule.map((row) => [
          row.monthLabel,
          row.emi,
          row.interestPaid,
          row.principalPaid,
          row.prepayment,
          row.interestRate,
          row.remainingPrincipal,
        ]),
      ];

      const ws = XLSX.utils.aoa_to_sheet(ws_data);
      const wscols = [
        { wch: 15 }, // Month
        { wch: 12 }, // EMI
        { wch: 12 }, // Interest
        { wch: 12 }, // Principal
        { wch: 12 }, // Prepayment
        { wch: 10 }, // Rate
        { wch: 15 }, // Balance
      ];
      ws["!cols"] = wscols;

      XLSX.utils.book_append_sheet(wb, ws, tab.name);
    });

    XLSX.writeFile(wb, `Financial_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div
      className="flex flex-col h-screen bg-white text-[#212529] select-none overflow-hidden"
      onMouseUp={handleMouseUp}
    >
      {/* Premium Toolbar */}
      <div className="flex bg-[#f8f9fa] border-b border-[#e1e1e1] p-3 gap-3 items-center shadow-[0_1px_3px_0_rgba(0,0,0,0.05)] z-30">
        <div className="flex items-center gap-3 px-4 border-r border-[#e1e1e1]">
          <div className="w-9 h-9 bg-[#107c41] rounded-lg flex items-center justify-center text-white font-bold shadow-lg shadow-green-500/20">
            X
          </div>
          <div className="flex flex-col">
            <span className="font-bold text-sm tracking-tight">
              TaxCalc Excel
            </span>
          </div>
        </div>

        {/* Currency Selector */}
        <div className="flex items-center gap-3 px-4 border-r border-[#e1e1e1]">
          <div className="flex flex-col">
            <span className="text-[9px] text-gray-400 font-bold uppercase tracking-widest leading-none mb-1">
              Currency
            </span>
            <select
              className="bg-transparent border-none p-0 text-[11px] font-bold outline-none cursor-pointer text-gray-700 hover:text-[#107c41] transition-colors"
              value={currency.code}
              onChange={(e) => {
                const selected = CURRENCIES.find((c) => c.code === e.target.value);
                if (selected) setCurrency(selected);
              }}
            >
              {CURRENCIES.map((c) => (
                <option key={c.code} value={c.code}>
                  {c.symbol} {c.code} — {c.name}
                </option>
              ))}
            </select>
          </div>
        </div>
        <div className="flex gap-1 ml-2">
          <div className="excel-ribbon-btn" onClick={addTab}>
            <div className="text-lg">✨</div>
            <span className="text-[10px] font-bold">New Sheet</span>
          </div>
          <div className="excel-ribbon-btn" onClick={exportToExcel}>
            <div className="text-lg">📊</div>
            <span className="text-[10px] font-bold">Export Excel</span>
          </div>

        </div>

        <div className="ml-auto flex items-center gap-4 bg-white/50 p-1 rounded-xl border border-white/50 backdrop-blur-sm">
          <div className="flex flex-col justify-center px-4 border-r border-[#e1e1e1]">
            <div className="text-[9px] text-gray-400 uppercase font-black tracking-tighter">
              {amortization.totalSavings < 0 ? "Interest Added" : "Interest Saved"}
            </div>
            <div className={`text-md font-bold ${amortization.totalSavings < 0 ? "text-red-600" : "text-[#107c41]"}`}>
              {currency.symbol} {formatCurrency(Math.abs(amortization.totalSavings), currency.locale)}
            </div>
          </div>
          <div className="flex flex-col justify-center px-4">
            <div className="text-[9px] text-gray-400 uppercase font-black tracking-tighter">
              {amortization.monthsSaved < 0 ? "Time Added" : "Time Saved"}
            </div>
            <div className={`text-md font-bold ${amortization.monthsSaved < 0 ? "text-red-600" : "text-blue-600"}`}>
              {Math.abs(amortization.monthsSaved)} Months
            </div>
          </div>
        </div>
      </div>

      {/* Main Container */}
      <div className="flex-1 flex overflow-hidden bg-[#f1f3f4]">
        {/* Spreadsheet Content */}
        <div
          className="flex-1 overflow-auto no-scrollbar-x relative"
          onScroll={handleScroll}
          ref={containerRef}
        >
          <div className="inline-block min-w-full bg-white relative">
            {/* Selection Border Overlay */}
            <div style={getSelectionBoxStyle()} />

            {/* Column Headers */}
            <div className="flex sticky top-0 z-20">
              <div className="w-10 h-6 bg-[#f3f3f3] border-r border-b border-[#d4d4d4]" />
              {COLUMNS.slice(1).map((c) => (
                <div
                  key={c}
                  className={`w-32 h-6 excel-header border-r border-b border-[#d4d4d4] transition-colors duration-200 ${isColSelected(c) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-b-2 border-b-[#107c41]" : ""}`}
                >
                  {c}
                </div>
              ))}
            </div>

            <div className="flex">
              <div
                className={`excel-row-num transition-colors duration-200 ${isRowSelected(1) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
              >
                1
              </div>
              <div
                className={`excel-cell header w-32 bg-yellow-50 font-bold ${isCellSelected(1, "A") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "A")}
                onMouseEnter={() => handleMouseEnter(1, "A")}
              >
                LOAN INPUTS
              </div>
              <div
                className={`excel-cell w-32 bg-yellow-50 ${isCellSelected(1, "B") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "B")}
                onMouseEnter={() => handleMouseEnter(1, "B")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(1, "C") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "C")}
                onMouseEnter={() => handleMouseEnter(1, "C")}
              ></div>
              <div
                className={`excel-cell header w-32 bg-green-50 font-bold ${isCellSelected(1, "D") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "D")}
                onMouseEnter={() => handleMouseEnter(1, "D")}
              >
                SUMMARY
              </div>
              <div
                className={`excel-cell w-32 bg-green-50 ${isCellSelected(1, "E") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "E")}
                onMouseEnter={() => handleMouseEnter(1, "E")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(1, "F") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "F")}
                onMouseEnter={() => handleMouseEnter(1, "F")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(1, "G") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 1, "G")}
                onMouseEnter={() => handleMouseEnter(1, "G")}
              ></div>
            </div>

            <div className="flex">
              <div
                className={`excel-row-num transition-colors duration-200 ${isRowSelected(2) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
              >
                2
              </div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(2, "A") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "A")}
                onMouseEnter={() => handleMouseEnter(2, "A")}
              >
                Principal
              </div>
              <div
                className={`excel-cell w-32 p-0 ${isCellSelected(2, "B") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "B")}
                onMouseEnter={() => handleMouseEnter(2, "B")}
              >
                <input
                  className="excel-input font-bold"
                  type="number"
                  value={activeTab.principal}
                  onChange={(e) =>
                    updateActiveTab({ principal: Number(e.target.value) })
                  }
                />
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(2, "C") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "C")}
                onMouseEnter={() => handleMouseEnter(2, "C")}
              ></div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(2, "D") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "D")}
                onMouseEnter={() => handleMouseEnter(2, "D")}
              >
                Total Paid
              </div>
              <div
                className={`excel-cell w-32 font-bold ${isCellSelected(2, "E") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "E")}
                onMouseEnter={() => handleMouseEnter(2, "E")}
              >
                {currency.symbol} {formatCurrency(amortization.totalAmount, currency.locale)}
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(2, "F") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "F")}
                onMouseEnter={() => handleMouseEnter(2, "F")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(2, "G") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 2, "G")}
                onMouseEnter={() => handleMouseEnter(2, "G")}
              ></div>
            </div>

            <div className="flex">
              <div
                className={`excel-row-num transition-colors duration-200 ${isRowSelected(3) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
              >
                3
              </div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(3, "A") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "A")}
                onMouseEnter={() => handleMouseEnter(3, "A")}
              >
                Int. Rate (%)
              </div>
              <div
                className={`excel-cell w-32 p-0 ${isCellSelected(3, "B") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "B")}
                onMouseEnter={() => handleMouseEnter(3, "B")}
              >
                <input
                  className="excel-input font-bold"
                  type="number"
                  step="0.1"
                  value={activeTab.interestRate}
                  onChange={(e) =>
                    updateActiveTab({ interestRate: Number(e.target.value) })
                  }
                />
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(3, "C") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "C")}
                onMouseEnter={() => handleMouseEnter(3, "C")}
              ></div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(3, "D") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "D")}
                onMouseEnter={() => handleMouseEnter(3, "D")}
              >
                Total Interest
              </div>
              <div
                className={`excel-cell w-32 font-bold text-red-600 ${isCellSelected(3, "E") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "E")}
                onMouseEnter={() => handleMouseEnter(3, "E")}
              >
                {currency.symbol} {formatCurrency(amortization.totalInterest, currency.locale)}
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(3, "F") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "F")}
                onMouseEnter={() => handleMouseEnter(3, "F")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(3, "G") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 3, "G")}
                onMouseEnter={() => handleMouseEnter(3, "G")}
              ></div>
            </div>

            <div className="flex">
              <div
                className={`excel-row-num transition-colors duration-200 ${isRowSelected(4) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
              >
                4
              </div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(4, "A") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "A")}
                onMouseEnter={() => handleMouseEnter(4, "A")}
              >
                Tenure (Yrs)
              </div>
              <div
                className={`excel-cell w-32 p-0 ${isCellSelected(4, "B") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "B")}
                onMouseEnter={() => handleMouseEnter(4, "B")}
              >
                <input
                  className="excel-input font-bold"
                  type="number"
                  value={activeTab.tenureYears}
                  onChange={(e) =>
                    updateActiveTab({ tenureYears: Number(e.target.value) })
                  }
                />
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(4, "C") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "C")}
                onMouseEnter={() => handleMouseEnter(4, "C")}
              ></div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(4, "D") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "D")}
                onMouseEnter={() => handleMouseEnter(4, "D")}
              >
                {amortization.monthsSaved < 0 ? "Months Added" : "Months Saved"}
              </div>
              <div
                className={`excel-cell w-32 font-bold ${amortization.monthsSaved < 0 ? "text-red-600" : "text-green-600"} ${isCellSelected(4, "E") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "E")}
                onMouseEnter={() => handleMouseEnter(4, "E")}
              >
                {Math.abs(amortization.monthsSaved)}
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(4, "F") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "F")}
                onMouseEnter={() => handleMouseEnter(4, "F")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(4, "G") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 4, "G")}
                onMouseEnter={() => handleMouseEnter(4, "G")}
              ></div>
            </div>

            <div className="flex">
              <div
                className={`excel-row-num transition-colors duration-200 ${isRowSelected(5) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
              >
                5
              </div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(5, "A") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "A")}
                onMouseEnter={() => handleMouseEnter(5, "A")}
              >
                Start Month
              </div>
              <div
                className={`excel-cell w-32 p-0 ${isCellSelected(5, "B") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "B")}
                onMouseEnter={() => handleMouseEnter(5, "B")}
              >
                <input
                  className="excel-input font-bold"
                  type="month"
                  value={activeTab.startMonth}
                  onChange={(e) => updateActiveTab({ startMonth: e.target.value })}
                />
              </div>
              <div
                className={`excel-cell w-32 ${isCellSelected(5, "C") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "C")}
                onMouseEnter={() => handleMouseEnter(5, "C")}
              ></div>
              <div
                className={`excel-cell w-32 text-gray-500 ${isCellSelected(5, "D") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "D")}
                onMouseEnter={() => handleMouseEnter(5, "D")}
              ></div>
              <div
                className={`excel-cell w-32 font-bold ${isCellSelected(5, "E") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "E")}
                onMouseEnter={() => handleMouseEnter(5, "E")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(5, "F") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "F")}
                onMouseEnter={() => handleMouseEnter(5, "F")}
              ></div>
              <div
                className={`excel-cell w-32 ${isCellSelected(5, "G") ? "selected" : ""}`}
                onMouseDown={(e) => handleMouseDown(e, 5, "G")}
                onMouseEnter={() => handleMouseEnter(5, "G")}
              ></div>
            </div>

            {/* Empty spacer row 6 to maintain Excel grid consistency */}
            {[6].map((r) => (
              <div key={r} className="flex">
                <div
                  className={`excel-row-num transition-colors duration-200 ${isRowSelected(r) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
                >
                  {r}
                </div>
                {COLUMNS.slice(1).map((c) => (
                  <div
                    key={c}
                    className={`excel-cell w-32 ${isCellSelected(r, c) ? "selected" : ""}`}
                    onMouseDown={(e) => handleMouseDown(e, r, c)}
                    onMouseEnter={() => handleMouseEnter(r, c)}
                  ></div>
                ))}
              </div>
            ))}

            {/* Table Header (Row 7) */}
            <div className="flex">
              <div
                className={`excel-row-num transition-colors duration-200 ${isRowSelected(7) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
              >
                7
              </div>
              <div className="excel-cell header w-32 bg-[#f3f3f3]">Month</div>
              <div className="excel-cell header w-32 bg-[#f3f3f3]">EMI</div>
              <div className="excel-cell header w-32 bg-[#f3f3f3]">
                Interest
              </div>
              <div className="excel-cell header w-32 bg-[#f3f3f3]">
                Principal
              </div>
              <div className="excel-cell header w-32 bg-blue-50 text-blue-800">
                Prepayment
              </div>
              <div className="excel-cell header w-32 bg-red-50 text-red-800">
                Rate %
              </div>
              <div className="excel-cell header w-32 bg-[#f3f3f3]">Balance</div>
            </div>

            {/* Table Body - VIRTUALIZED */}
            <div
              style={{
                height: virtualizedData.totalHeight,
                position: "relative",
              }}
            >
              <div
                style={{
                  transform: `translateY(${virtualizedData.topOffset}px)`,
                }}
              >
                {virtualizedData.rows.map((row) => {
                  const globalIdx = row.monthIndex + 8;
                  return (
                    <div
                      key={row.monthIndex}
                      className="flex group transition-opacity duration-200"
                    >
                      <div
                        className={`excel-row-num transition-colors duration-200 ${isRowSelected(globalIdx) ? "bg-[#107c41]/10 text-[#107c41] font-bold border-r-2 border-r-[#107c41]" : ""}`}
                      >
                        {globalIdx}
                      </div>
                      <div
                        className={`excel-cell w-32 font-medium bg-[#fafafa] ${isCellSelected(globalIdx, "A") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "A")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "A")}
                      >
                        {row.monthLabel}
                      </div>
                      <div
                        className={`excel-cell w-32 text-right ${isCellSelected(globalIdx, "B") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "B")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "B")}
                      >
                        {currency.symbol} {formatCurrency(row.emi, currency.locale)}
                      </div>
                      <div
                        className={`excel-cell w-32 text-right text-rose-600 ${isCellSelected(globalIdx, "C") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "C")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "C")}
                      >
                        {currency.symbol} {formatCurrency(row.interestPaid, currency.locale)}
                      </div>
                      <div
                        className={`excel-cell w-32 text-right text-emerald-600 ${isCellSelected(globalIdx, "D") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "D")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "D")}
                      >
                        {currency.symbol} {formatCurrency(row.principalPaid, currency.locale)}
                      </div>
                      <div
                        className={`excel-cell w-32 p-0 bg-blue-50/30 flex items-center ${isCellSelected(globalIdx, "E") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "E")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "E")}
                      >
                        {row.prepayment > 0 && (
                          <span className="pl-2 text-blue-400 text-[10px] shrink-0 select-none">
                            {currency.symbol}
                          </span>
                        )}
                        <input
                          className="excel-input text-right text-blue-700 font-medium flex-1 bg-transparent border-none outline-none pr-2 h-full w-full"
                          type="number"
                          placeholder={row.prepayment === 0 ? "0" : ""}
                          value={row.prepayment || ""}
                          onChange={(e) =>
                            handlePrepaymentChange(
                              row.monthIndex,
                              e.target.value,
                            )
                          }
                        />
                      </div>
                      <div
                        className={`excel-cell w-32 p-0 bg-rose-50/30 ${isCellSelected(globalIdx, "F") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "F")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "F")}
                      >
                        <input
                          className="excel-input text-right text-rose-700 font-medium"
                          type="number"
                          step="0.1"
                          value={row.interestRate}
                          onChange={(e) =>
                            handleRateChange(row.monthIndex, e.target.value)
                          }
                        />
                      </div>
                      <div
                        className={`excel-cell w-32 text-right font-bold ${isCellSelected(globalIdx, "G") ? "selected" : ""}`}
                        onMouseDown={(e) => handleMouseDown(e, globalIdx, "G")}
                        onMouseEnter={() => handleMouseEnter(globalIdx, "G")}
                      >
                        {currency.symbol} {formatCurrency(row.remainingPrincipal, currency.locale)}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* ── Analytics Panel ── */}
        <div className="flex-1 min-w-[480px] bg-[#f8f9fa] flex flex-col overflow-hidden z-10 border-l border-[#e1e1e1]">
          <div className="px-6 py-3 bg-white border-b border-[#e1e1e1] flex items-center justify-between shrink-0">
            <div>
              <h2 className="text-[13px] font-extrabold text-gray-800 tracking-tight">
                Loan Analytics
              </h2>
              <p className="text-[10px] text-gray-400 mt-0.5 tracking-wide">
                Real-time financial insights
              </p>
            </div>
            <div className="flex items-center gap-1.5 px-2 py-0.5 rounded-full bg-green-50">
              <div className="w-1.5 h-1.5 rounded-full bg-[#107c41] animate-pulse" />
              <span className="text-[9px] text-[#107c41] font-bold uppercase tracking-wider">
                Live
              </span>
            </div>
          </div>
 
          {/* Scrollable body */}
          <div
            className="flex-1 overflow-y-auto px-8 py-8 flex flex-col gap-10"
            style={{
              scrollbarWidth: "thin",
              scrollbarColor: "#d1d5db transparent",
            }}
          >
            {/* KPI strip */}
            <div className="grid grid-cols-3 gap-5">
              {[
                {
                  label: amortization.totalSavings < 0 ? "Interest Added" : "Interest Saved",
                  value: `${currency.symbol}${formatCurrency(Math.abs(amortization.totalSavings), currency.locale)}`,
                  sub: amortization.totalSavings < 0 ? "due to rate changes" : "via prepayments",
                  accent: amortization.totalSavings < 0 ? "#e11d48" : "#107c41",
                  bg: amortization.totalSavings < 0 ? "#fff1f2" : "#f0fdf4",
                },
                {
                  label: amortization.monthsSaved < 0 ? "Time Added" : "Months Freed",
                  value: `${Math.abs(amortization.monthsSaved)} Months`,
                  sub: amortization.monthsSaved < 0 
                    ? `${(Math.abs(amortization.monthsSaved) / 12).toFixed(1)} yrs extra`
                    : `${(amortization.monthsSaved / 12).toFixed(1)} yrs early`,
                  accent: amortization.monthsSaved < 0 ? "#e11d48" : "#2563eb",
                  bg: amortization.monthsSaved < 0 ? "#fff1f2" : "#eff6ff",
                },
                {
                  label: "EMI / Month",
                  value: `${currency.symbol}${amortization.schedule[0] ? formatCurrency(amortization.schedule[0].emi, currency.locale) : "—"}`,
                  sub: `at ${activeTab.interestRate}%`,
                  accent: "#7c3aed",
                  bg: "#f5f3ff",
                },
              ].map((k) => (
                <div
                  key={k.label}
                  className="rounded-2xl p-4 flex flex-col gap-1.5"
                  style={{ background: k.bg }}
                >
                  <p
                    className="text-[9px] font-bold uppercase tracking-widest"
                    style={{ color: k.accent + "99" }}
                  >
                    {k.label}
                  </p>
                  <p className="text-lg font-black text-gray-900 leading-none tracking-tight">
                    {k.value}
                  </p>
                  <p
                    className="text-[9px] font-semibold"
                    style={{ color: k.accent }}
                  >
                    {k.sub}
                  </p>
                </div>
              ))}
            </div>
 
            {/* Donut + Completion row */}
            <div className="grid grid-cols-2 gap-8">
              {/* Donut */}
              <div>
                <p className="text-[10px] font-bold text-gray-500 uppercase tracking-wider mb-4">
                  Cost Split
                </p>
                <div className="flex items-center gap-5">
                  {(() => {
                    const size = 140;
                    const cx = 70;
                    const cy = 70;
                    const r = 52;
                    const circ = 2 * Math.PI * r;
                    const pPct = activeTab.principal / amortization.totalAmount;
                    const pDash = pPct * circ;
                    const iDash = circ - pDash;
                    return (
                      <svg width={size} height={size} className="shrink-0">
                        <circle
                          cx={cx}
                          cy={cy}
                          r={r}
                          fill="none"
                          stroke="#f3f4f6"
                          strokeWidth="22"
                        />
                        <circle
                          cx={cx}
                          cy={cy}
                          r={r}
                          fill="none"
                          stroke="#fca5a5"
                          strokeWidth="22"
                          strokeDasharray={`${iDash} ${circ}`}
                          strokeDashoffset={-pDash}
                          transform={`rotate(-90 ${cx} ${cy})`}
                          strokeLinecap="butt"
                        />
                        <circle
                          cx={cx}
                          cy={cy}
                          r={r}
                          fill="none"
                          stroke="#107c41"
                          strokeWidth="22"
                          strokeDasharray={`${pDash} ${circ}`}
                          transform={`rotate(-90 ${cx} ${cy})`}
                          strokeLinecap="butt"
                        />
                        <text
                          x={cx}
                          y={cy - 5}
                          textAnchor="middle"
                          fill="#111827"
                          fontSize="18"
                          fontWeight="800"
                        >
                          {Math.round(pPct * 100)}%
                        </text>
                        <text
                          x={cx}
                          y={cy + 12}
                          textAnchor="middle"
                          fill="#9ca3af"
                          fontSize="9"
                        >
                          principal
                        </text>
                      </svg>
                    );
                  })()}
                  <div className="flex flex-col gap-4">
                    <div>
                      <div className="flex items-center gap-2 mb-1">
                        <div className="w-3 h-3 rounded bg-[#107c41]" />
                        <span className="text-[11px] text-gray-500">
                          Principal
                        </span>
                      </div>
                      <p className="text-sm font-bold text-gray-900 pl-5">
                        {currency.symbol}{formatCurrency(activeTab.principal, currency.locale)}
                      </p>
                    </div>
                    <div>
                      <div className="flex items-center gap-2 mb-1">
                        <div className="w-3 h-3 rounded bg-rose-400" />
                        <span className="text-[11px] text-gray-500">
                          Interest
                        </span>
                      </div>
                      <p className="text-sm font-bold text-rose-600 pl-5">
                        {currency.symbol}{formatCurrency(amortization.totalInterest, currency.locale)}
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              {/* Completion + Efficiency */}
              <div className="flex flex-col gap-5">
                <div>
                  <p className="text-[10px] font-bold text-gray-500 uppercase tracking-wider mb-2">
                    Completion
                  </p>
                  {(() => {
                    const pct = Math.min(
                      100,
                      Math.round(
                        (amortization.schedule.length /
                          (activeTab.tenureYears * 12)) *
                          100,
                      ),
                    );
                    return (
                      <>
                        <div className="flex items-baseline gap-1.5">
                          <span className="text-4xl font-black text-gray-900 tracking-tight">
                            {pct}
                          </span>
                          <span className="text-lg font-bold text-gray-300">
                            %
                          </span>
                        </div>
                        <p className="text-[10px] text-gray-400 mt-1">
                          {amortization.schedule.length} of{" "}
                          {activeTab.tenureYears * 12} months
                        </p>
                        <div className="w-full h-2.5 bg-gray-100 rounded-full mt-2.5">
                          <div
                            className="h-full rounded-full transition-all duration-700"
                            style={{
                              width: `${pct}%`,
                              background:
                                "linear-gradient(90deg,#107c41,#2ebc6c)",
                            }}
                          />
                        </div>
                      </>
                    );
                  })()}
                </div>
                <div className="pt-4 border-t border-gray-100">
                  <div className="flex items-center gap-1.5 mb-2">
                    <p className="text-[10px] font-bold text-gray-500 uppercase tracking-wider">
                      Prepay Efficiency
                    </p>
                    <div className="group relative">
                      <div className="w-3 h-3 rounded-full bg-gray-100 flex items-center justify-center text-[8px] text-gray-500 cursor-help hover:bg-gray-200 transition-colors">
                        ?
                      </div>
                      <div className="absolute bottom-full left-0 mb-2 w-56 p-2.5 bg-gray-900 text-white text-[9px] rounded-xl opacity-0 group-hover:opacity-100 transition-all transform scale-95 group-hover:scale-100 pointer-events-none z-50 shadow-xl border border-white/10">
                        <p className="font-bold mb-1 text-orange-400">
                          Money's Power
                        </p>
                        <p className="mb-2">
                          This shows how many {currency.code} of future interest you save
                          for every {currency.symbol}1 you pay early today.
                        </p>
                        <div className="space-y-2 border-t border-white/10 pt-2">
                          <p>
                            <span className="font-bold text-emerald-400">
                              If &gt; 1.0x (e.g. 2.40x):
                            </span>{" "}
                            Prepaying is a huge win. Every {currency.symbol}1 destroys {currency.symbol}2.40 of
                            debt. This is usually better than most other
                            investments.
                          </p>
                          <p>
                            <span className="font-bold text-rose-400">
                              If &lt; 1.0x (e.g. 0.60x):
                            </span>{" "}
                            Diminishing returns. You might earn more by putting
                            that money in a Mutual Fund or FD instead of
                            prepaying.
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>
                  {(() => {
                    const totalPrepaid = amortization.schedule.reduce(
                      (a, r) => a + r.prepayment,
                      0,
                    );
                    const effNum =
                      totalPrepaid > 0
                        ? amortization.totalSavings / totalPrepaid
                        : 0;
                    const eff = totalPrepaid > 0 ? effNum.toFixed(2) : null;
 
                    let statusNote = "";
                    let noteColor = "text-gray-500";
 
                    if (effNum >= 2) {
                      statusNote = "Exceptional!";
                      noteColor = "text-emerald-600";
                    } else if (effNum >= 1) {
                      statusNote = "Good ROI.";
                      noteColor = "text-blue-600";
                    } else if (effNum > 0) {
                      statusNote = "Diminishing returns.";
                      noteColor = "text-orange-600";
                    }
 
                    return eff ? (
                      <div className="flex flex-col gap-1">
                        <div className="flex items-baseline gap-1.5">
                          <span className="text-3xl font-black text-orange-500">
                            {eff}
                          </span>
                          <span className="text-sm font-bold text-gray-300">
                            x
                          </span>
                          <span className="text-[10px] text-gray-400 ml-1 font-medium">
                            per {currency.symbol}1 prepaid
                          </span>
                        </div>
                        <p
                          className={`text-[9px] font-bold leading-tight ${noteColor}`}
                        >
                          {statusNote} Every {currency.symbol}1 prepaid saves {currency.symbol}{effNum.toFixed(2)} in interest.
                        </p>
                      </div>
                    ) : (
                      <p className="text-sm text-gray-300 italic">
                        No prepayments yet
                      </p>
                    );
                  })()}
                </div>
              </div>
            </div>

            {/* Balance Area Chart */}
            <div>
              <div className="flex justify-between items-center mb-4">
                <div>
                  <p className="text-[10px] font-bold text-gray-500 uppercase tracking-wider">
                    Loan Balance Over Time
                  </p>
                  <p className="text-[9px] text-gray-400 mt-0.5">
                    {currency.symbol}{formatCurrency(activeTab.principal, currency.locale)} → {currency.symbol}0
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-4 h-[2px] bg-[#107c41] rounded" />
                  <span className="text-[9px] text-gray-400">Balance</span>
                </div>
              </div>
              {(() => {
                const schedule = amortization.schedule;
                if (schedule.length < 2)
                  return (
                    <div className="h-36 flex items-center justify-center text-gray-300 text-xs">
                      Enter loan details above
                    </div>
                  );
                const W = 600;
                const H = 150;
                const step = Math.max(1, Math.floor(schedule.length / 120));
                const samples: typeof schedule = [];
                for (let i = 0; i < schedule.length; i += step)
                  samples.push(schedule[i]);
                if (
                  samples[samples.length - 1] !== schedule[schedule.length - 1]
                )
                  samples.push(schedule[schedule.length - 1]);
                const maxBal = activeTab.principal;
                const pad = 10;
                const pts = samples.map(
                  (row, i) =>
                    [
                      (i / (samples.length - 1)) * W,
                      H -
                        pad -
                        (row.remainingPrincipal / maxBal) * (H - pad * 2),
                    ] as [number, number],
                );
                const linePath = pts
                  .map(
                    (p, i) =>
                      `${i === 0 ? "M" : "L"} ${p[0].toFixed(1)} ${p[1].toFixed(1)}`,
                  )
                  .join(" ");
                const areaPath = linePath + ` L ${W} ${H} L 0 ${H} Z`;
                const handleMouseMove = (e: React.MouseEvent) => {
                  const rect = e.currentTarget.getBoundingClientRect();
                  const x = e.clientX - rect.left;
                  const ratio = Math.max(0, Math.min(1, x / rect.width));
                  const idx = Math.round(ratio * (samples.length - 1));
                  setBalanceHoverIdx(idx);
                };
 
                const hoverRow =
                  balanceHoverIdx !== null ? samples[balanceHoverIdx] : null;
                const hoverX =
                  balanceHoverIdx !== null
                    ? (balanceHoverIdx / (samples.length - 1)) * W
                    : 0;
 
                return (
                  <div className="relative group">
                    <svg
                      width="100%"
                      viewBox={`0 0 ${W} ${H}`}
                      preserveAspectRatio="none"
                      style={{ height: 150 }}
                      onMouseMove={handleMouseMove}
                      onMouseLeave={() => setBalanceHoverIdx(null)}
                      className="cursor-crosshair"
                    >
                      <defs>
                        <linearGradient id="aGrad" x1="0" y1="0" x2="0" y2="1">
                          <stop
                            offset="0%"
                            stopColor="#107c41"
                            stopOpacity="0.12"
                          />
                          <stop
                            offset="100%"
                            stopColor="#107c41"
                            stopOpacity="0.01"
                          />
                        </linearGradient>
                      </defs>
                      {[0.25, 0.5, 0.75].map((f) => (
                        <line
                          key={f}
                          x1="0"
                          y1={H * f}
                          x2={W}
                          y2={H * f}
                          stroke="#f0f0f0"
                          strokeWidth="1"
                        />
                      ))}
                      <path d={areaPath} fill="url(#aGrad)" />
                      <path
                        d={linePath}
                        fill="none"
                        stroke="#107c41"
                        strokeWidth="2.5"
                        strokeLinejoin="round"
                      />
                      <circle
                        cx={pts[0][0]}
                        cy={pts[0][1]}
                        r="4"
                        fill="#107c41"
                      />
                      <circle
                        cx={pts[pts.length - 1][0]}
                        cy={pts[pts.length - 1][1]}
                        r="4"
                        fill="#107c41"
                      />
 
                      {/* Hover line */}
                      {hoverRow && (
                        <line
                          x1={hoverX}
                          y1="0"
                          x2={hoverX}
                          y2={H}
                          stroke="#111827"
                          strokeWidth="1"
                          strokeDasharray="3 2"
                        />
                      )}
                    </svg>
 
                    {/* Tooltip */}
                    {hoverRow && (
                      <div
                        className="absolute top-0 z-50 pointer-events-none bg-white/95 backdrop-blur-sm border border-gray-200 shadow-xl rounded-lg p-2.5 flex flex-col gap-1.5 min-w-[140px]"
                        style={{
                          left: hoverX > W / 2 ? "auto" : hoverX + 10,
                          right: hoverX > W / 2 ? W - hoverX + 10 : "auto",
                          top: 10,
                        }}
                      >
                        <p className="text-[10px] font-bold text-gray-900 border-b border-gray-100 pb-1 mb-0.5">
                          Month {hoverRow.monthIndex + 1}
                        </p>
                        <div className="flex items-center justify-between gap-4">
                          <div className="flex items-center gap-1.5">
                            <div className="w-1.5 h-1.5 rounded-full bg-[#107c41]" />
                            <span className="text-[9px] text-gray-500 font-medium">
                              Balance
                            </span>
                          </div>
                          <span className="text-[10px] font-bold text-gray-900">
                            {currency.symbol}{formatCurrency(hoverRow.remainingPrincipal, currency.locale)}
                          </span>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })()}
              <div className="flex justify-between mt-2">
                <span className="text-[9px] text-gray-400">Month 1</span>
                <span className="text-[9px] text-gray-400">
                  Month {amortization.schedule.length} — Loan Cleared
                </span>
              </div>
            </div>
 
            {/* EMI Composition Chart */}
            <div className="pt-6 border-t border-gray-100">
              <div className="flex justify-between items-center mb-4">
                <div>
                  <p className="text-[10px] font-bold text-gray-500 uppercase tracking-wider">
                    EMI Composition Over Time
                  </p>
                  <p className="text-[9px] text-gray-400 mt-0.5">
                    Monthly Interest vs. Principal split
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="flex items-center gap-1.5">
                    <div className="w-2 h-2 rounded-sm bg-rose-400" />
                    <span className="text-[9px] text-gray-400 font-medium">
                      Interest
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <div className="w-2 h-2 rounded-sm bg-[#107c41]" />
                    <span className="text-[9px] text-gray-400 font-medium">
                      Principal
                    </span>
                  </div>
                </div>
              </div>
              {(() => {
                const schedule = amortization.schedule;
                if (schedule.length < 2) return null;
                const W = 600;
                const H = 120;
                const step = Math.max(1, Math.floor(schedule.length / 120));
                const samples: typeof schedule = [];
                for (let i = 0; i < schedule.length; i += step)
                  samples.push(schedule[i]);
                if (
                  samples[samples.length - 1] !== schedule[schedule.length - 1]
                )
                  samples.push(schedule[schedule.length - 1]);
 
                const maxEMI = Math.max(...samples.map((r) => r.emi)) || 1;
                const pad = 5;
                const usableH = H - pad * 2;
 
                const totalPts = samples.map((row, i) => [
                  (i / (samples.length - 1)) * W,
                  H -
                    pad -
                    ((row.interestPaid + row.principalPaid) / maxEMI) * usableH,
                ]);
                const interestPts = samples.map((row, i) => [
                  (i / (samples.length - 1)) * W,
                  H - pad - (row.interestPaid / maxEMI) * usableH,
                ]);
 
                const totalAreaPath =
                  totalPts
                    .map(
                      (p, i) =>
                        `${i === 0 ? "M" : "L"} ${p[0].toFixed(1)} ${p[1].toFixed(1)}`,
                    )
                    .join(" ") + ` L ${W} ${H} L 0 ${H} Z`;
 
                const interestAreaPath =
                  interestPts
                    .map(
                      (p, i) =>
                        `${i === 0 ? "M" : "L"} ${p[0].toFixed(1)} ${p[1].toFixed(1)}`,
                    )
                    .join(" ") + ` L ${W} ${H} L 0 ${H} Z`;
 
                const handleMouseMove = (e: React.MouseEvent) => {
                  const rect = e.currentTarget.getBoundingClientRect();
                  const x = e.clientX - rect.left;
                  const ratio = Math.max(0, Math.min(1, x / rect.width));
                  const idx = Math.round(ratio * (samples.length - 1));
                  setEmiHoverIdx(idx);
                };
 
                const hoverRow =
                  emiHoverIdx !== null ? samples[emiHoverIdx] : null;
                const hoverX =
                  emiHoverIdx !== null
                    ? (emiHoverIdx / (samples.length - 1)) * W
                    : 0;
 
                return (
                  <div className="relative group">
                    <svg
                      width="100%"
                      viewBox={`0 0 ${W} ${H}`}
                      preserveAspectRatio="none"
                      style={{ height: 120 }}
                      onMouseMove={handleMouseMove}
                      onMouseLeave={() => setEmiHoverIdx(null)}
                      className="cursor-crosshair"
                    >
                      {/* Principal Area (Background) */}
                      <path
                        d={totalAreaPath}
                        fill="#107c41"
                        fillOpacity="0.8"
                      />
                      {/* Interest Area (Foreground) */}
                      <path d={interestAreaPath} fill="#fb7185" />
                      {/* Separation Line */}
                      <path
                        d={interestPts
                          .map(
                            (p, i) =>
                              `${i === 0 ? "M" : "L"} ${p[0].toFixed(1)} ${p[1].toFixed(1)}`,
                          )
                          .join(" ")}
                        fill="none"
                        stroke="white"
                        strokeWidth="0.5"
                        strokeOpacity="0.5"
                      />
 
                      {/* Hover line */}
                      {hoverRow && (
                        <line
                          x1={hoverX}
                          y1="0"
                          x2={hoverX}
                          y2={H}
                          stroke="#111827"
                          strokeWidth="1"
                          strokeDasharray="3 2"
                        />
                      )}
                    </svg>
 
                    {/* Tooltip */}
                    {hoverRow && (
                      <div
                        className="absolute top-0 z-50 pointer-events-none bg-white/95 backdrop-blur-sm border border-gray-200 shadow-xl rounded-lg p-2.5 flex flex-col gap-1.5 min-w-[120px]"
                        style={{
                          left: hoverX > W / 2 ? "auto" : hoverX + 10,
                          right: hoverX > W / 2 ? W - hoverX + 10 : "auto",
                          top: 10,
                        }}
                      >
                        <p className="text-[10px] font-bold text-gray-900 border-b border-gray-100 pb-1 mb-0.5">
                          Month {hoverRow.monthIndex + 1}
                        </p>
                        <div className="flex items-center justify-between gap-4">
                          <div className="flex items-center gap-1.5">
                            <div className="w-1.5 h-1.5 rounded-full bg-[#107c41]" />
                            <span className="text-[9px] text-gray-500 font-medium">
                              Principal
                            </span>
                          </div>
                          <span className="text-[10px] font-bold text-gray-900">
                            {currency.symbol}{formatCurrency(hoverRow.principalPaid, currency.locale)}
                          </span>
                        </div>
                        <div className="flex items-center justify-between gap-4">
                          <div className="flex items-center gap-1.5">
                            <div className="w-1.5 h-1.5 rounded-full bg-rose-400" />
                            <span className="text-[9px] text-gray-500 font-medium">
                              Interest
                            </span>
                          </div>
                          <span className="text-[10px] font-bold text-rose-600">
                            {currency.symbol}{formatCurrency(hoverRow.interestPaid, currency.locale)}
                          </span>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })()}
              <div className="flex justify-between mt-2">
                <span className="text-[9px] text-gray-400">Month 1</span>
                <span className="text-[9px] text-gray-400">
                  Month {amortization.schedule.length}
                </span>
              </div>
            </div>


          </div>
        </div>
      </div>

      {/* Bottom Tabs Bar */}
      <div className="flex bg-[#f8f9fa] border-t border-[#e1e1e1] h-10 items-center px-2 gap-1 z-30 shadow-[0_-1px_3px_0_rgba(0,0,0,0.05)]">
        <div className="flex items-center h-full mr-1 shrink-0">
          <div
            className="w-8 h-8 flex items-center justify-center text-gray-400 hover:text-gray-600 transition-colors cursor-pointer hover:bg-gray-200 rounded"
            onClick={() => scrollTabs("left")}
          >
            ◀
          </div>
          <div
            className="w-8 h-8 flex items-center justify-center text-gray-400 hover:text-gray-600 transition-colors cursor-pointer hover:bg-gray-200 rounded"
            onClick={() => scrollTabs("right")}
          >
            ▶
          </div>
        </div>

        <div
          ref={tabsContainerRef}
          className="flex h-full gap-0.5 overflow-x-auto scrollbar-hide flex-1 items-center"
        >
          {tabs.map((tab) => (
            <div
              key={tab.id}
              className={`excel-tab shrink-0 ${activeTabId === tab.id ? "active" : ""}`}
              onClick={() => setActiveTabId(tab.id)}
              onDoubleClick={() => setEditingTabId(tab.id)}
            >
              <div className="flex-1 flex justify-start pr-6 min-w-0 overflow-hidden">
                {editingTabId === tab.id ? (
                  <input
                    autoFocus
                    className="bg-transparent border-none outline-none text-[13px] w-full font-bold text-[#107c41] text-left"
                    defaultValue={tab.name}
                    onBlur={(e) => {
                      const newName = e.target.value.trim();
                      if (newName)
                        setTabs(
                          tabs.map((t) =>
                            t.id === tab.id ? { ...t, name: newName } : t,
                          ),
                        );
                      setEditingTabId(null);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        const newName = e.currentTarget.value.trim();
                        if (newName)
                          setTabs(
                            tabs.map((t) =>
                              t.id === tab.id ? { ...t, name: newName } : t,
                            ),
                          );
                        setEditingTabId(null);
                      } else if (e.key === "Escape") {
                        setEditingTabId(null);
                      }
                    }}
                    onClick={(e) => e.stopPropagation()}
                  />
                ) : (
                  <span className="whitespace-nowrap truncate block w-full">
                    {tab.name}
                  </span>
                )}
              </div>
              {tabs.length > 1 && editingTabId !== tab.id && (
                <span
                  className="absolute right-2 w-4 h-4 rounded-full hover:bg-black/10 flex items-center justify-center text-[10px] transition-colors"
                  onClick={(e) => {
                    e.stopPropagation();
                    removeTab(tab.id);
                  }}
                >
                  ×
                </span>
              )}
            </div>
          ))}
        </div>

        <button
          onClick={addTab}
          className="w-8 h-8 flex items-center justify-center hover:bg-gray-200 rounded-md transition-all text-lg text-gray-500 hover:text-[#107c41] shrink-0 ml-1"
          title="Add New Sheet"
        >
          ⊕
        </button>
      </div>

      {/* Excel-style Status Bar (Dedicated for Stats) */}
      <div className="flex bg-[#107c41] text-white h-7 items-center px-4 text-[11px] font-medium z-40 select-none shadow-[0_-2px_10px_rgba(0,0,0,0.1)]">
        <div className="flex items-center gap-2 flex-1">
          <div className="w-2 h-2 bg-white rounded-full animate-pulse" />
          <span className="uppercase tracking-wider">Ready</span>
          <span className="ml-2 text-[10px] opacity-70 font-normal">Autosave on</span>
        </div>

        <div className="flex items-center justify-center flex-1">
          {stats && (
            <div className="flex gap-5 animate-in fade-in slide-in-from-bottom-1">
              <div className="flex items-center gap-1.5 opacity-90">
                <span className="text-[10px] font-bold text-green-200 uppercase">
                  Average:
                </span>
                <span className="text-white">{formatCurrency(stats.avg, currency.locale, 2)}</span>
              </div>
              <div className="w-px h-3 bg-white/20 self-center" />
              <div className="flex items-center gap-1.5 opacity-90">
                <span className="text-[10px] font-bold text-green-200 uppercase">
                  Count:
                </span>
                <span className="text-white">{stats.count}</span>
              </div>
              <div className="w-px h-3 bg-white/20 self-center" />
              <div className="flex items-center gap-1.5 opacity-90">
                <span className="text-[10px] font-bold text-green-200 uppercase">
                  Sum:
                </span>
                <span className="text-white">{formatCurrency(stats.sum, currency.locale, 2)}</span>
              </div>
            </div>
          )}
        </div>

        {process.env.NEXT_PUBLIC_GITHUB_REPO_URL && (
          <div className="flex-1 flex justify-end items-center pr-2">
            <a
              href={process.env.NEXT_PUBLIC_GITHUB_REPO_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-1.5 hover:text-green-200 transition-colors opacity-90 hover:opacity-100"
            >
              <svg className="w-3.5 h-3.5 fill-current" viewBox="0 0 24 24">
                <path d="M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.43.372.823 1.102.823 2.222 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12" />
              </svg>
              <span className="hidden sm:inline">Contribute on GitHub</span>
            </a>
          </div>
        )}
        {!process.env.NEXT_PUBLIC_GITHUB_REPO_URL && <div className="flex-1" />}
      </div>

      <style jsx global>{`
        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }

        .excel-input:focus {
          z-index: 50;
          position: relative;
        }

        .excel-cell.selected {
          background-color: rgba(16, 124, 65, 0.1) !important;
          box-shadow: none !important;
          position: relative;
          z-index: 5;
        }
      `}</style>
    </div>
  );
}
