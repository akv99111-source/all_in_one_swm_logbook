import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, ShieldCheck, Plus, Trash2, ArrowLeft, Settings2, FileSpreadsheet, BookOpen, CheckCircle2, AlertTriangle, Zap } from 'lucide-react';
import * as XLSX from 'xlsx';

const MONTHS = [
  { id: 1, shortEn: 'Jan', shortHi: 'जनवरी', fullEn: 'January' },
  { id: 2, shortEn: 'Feb', shortHi: 'फरवरी', fullEn: 'February' },
  { id: 3, shortEn: 'Mar', shortHi: 'मार्च', fullEn: 'March' },
  { id: 4, shortEn: 'Apr', shortHi: 'अप्रैल', fullEn: 'April' },
  { id: 5, shortEn: 'May', shortHi: 'मई', fullEn: 'May' },
  { id: 6, shortEn: 'Jun', shortHi: 'जून', fullEn: 'June' },
  { id: 7, shortEn: 'Jul', shortHi: 'जुलाई', fullEn: 'July' },
  { id: 8, shortEn: 'Aug', shortHi: 'अगस्त', fullEn: 'August' },
  { id: 9, shortEn: 'Sep', shortHi: 'सितंबर', fullEn: 'September' },
  { id: 10, shortEn: 'Oct', shortHi: 'अक्टूबर', fullEn: 'October' },
  { id: 11, shortEn: 'Nov', shortHi: 'नवंबर', fullEn: 'November' },
  { id: 12, shortEn: 'Dec', shortHi: 'दिसंबर', fullEn: 'December' }
];

const STATES_LIST = [
  { nameEn: 'Andaman & Nicobar Islands', nameHi: 'अंडमान और निकोबार' },
  { nameEn: 'Andhra Pradesh', nameHi: 'आंध्र प्रदेश' },
  { nameEn: 'Arunachal Pradesh', nameHi: 'अरुणाचल प्रदेश' },
  { nameEn: 'Assam', nameHi: 'असम' },
  { nameEn: 'Bihar', nameHi: 'बिहार' },
  { nameEn: 'Chandigarh', nameHi: 'चंडीगढ़' },
  { nameEn: 'Chhattisgarh', nameHi: 'छत्तीसगढ़' },
  { nameEn: 'Delhi (NCR)', nameHi: 'दिल्ली (एनसीआर)' },
  { nameEn: 'Goa', nameHi: 'गोवा' },
  { nameEn: 'Gujarat', nameHi: 'गुजरात' },
  { nameEn: 'Haryana', nameHi: 'हरियाणा' },
  { nameEn: 'Himachal Pradesh', nameHi: 'हिमाचल प्रदेश' },
  { nameEn: 'Jammu and Kashmir', nameHi: 'जम्मू और कश्मीर' },
  { nameEn: 'Jharkhand', nameHi: 'झारखंड' },
  { nameEn: 'Karnataka', nameHi: 'कर्नाटक' },
  { nameEn: 'Kerala', nameHi: 'केरल' },
  { nameEn: 'Madhya Pradesh', nameHi: 'मध्य प्रदेश' },
  { nameEn: 'Maharashtra', nameHi: 'महाराष्ट्र' },
  { nameEn: 'Odisha', nameHi: 'ओडिशा' },
  { nameEn: 'Punjab', nameHi: 'पंजाब' },
  { nameEn: 'Rajasthan', nameHi: 'राजस्थान' },
  { nameEn: 'Tamil Nadu', nameHi: 'तमिलनाडु' },
  { nameEn: 'Telangana', nameHi: 'तेलंगाना' },
  { nameEn: 'Uttar Pradesh', nameHi: 'उत्तर प्रदेश' },
  { nameEn: 'Uttarakhand', nameHi: 'उत्तराखंड' },
  { nameEn: 'West Bengal', nameHi: 'पश्चिम बंगाल' },
  { nameEn: 'Other / Pan-India Standard', nameHi: 'अन्य / राष्ट्रीय मानक' }
];

const cyrb128 = (str) => {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) {
    let k = str.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  return (Math.imul(h3 ^ (h1 >>> 18), 597399067) ^ Math.imul(h4 ^ (h2 >>> 22), 2869860233) ^ Math.imul(h1 ^ (h3 >>> 17), 951274213) ^ Math.imul(h2 ^ (h4 >>> 19), 2716044179)) >>> 0;
};

const mulberry32 = (a) => {
  return function() {
    var t = a += 0x6D2B79F5;
    t = Math.imul(t ^ t >>> 15, t | 1);
    t ^= t + Math.imul(t ^ t >>> 7, t | 61);
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  };
};

const inputStyle = { width: '100%', padding: '9px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', boxSizing: 'border-box' };

export default function App() {
  const [lang, setLang] = useState('hi');
  const [selectedState, setSelectedState] = useState('Uttar Pradesh');
  const [name, setName] = useState('Nagar Palika Parishad');
  const [phone, setPhone] = useState('');
  
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(50000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [actualAverageTpd, setActualAverageTpd] = useState(22.5);
  
  // Bring Back Segregation Rate
  const [segregationRate, setSegregationRate] = useState(80);

  const [facilities, setFacilities] = useState([
    { id: 'f1', name: 'Windrow Compost Pad', type: 'wet_compost', designCapacity: 12, avgProcessing: 10.8 },
    { id: 'f2', name: 'Dry MRF Sorting Shed', type: 'dry_mrf', designCapacity: 8, avgProcessing: 7.2 },
    { id: 'f3', name: 'Trommel Screening Unit', type: 'mixed_trommel', designCapacity: 6, avgProcessing: 4.5 }
  ]);
  
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [activeAssetView, setActiveAssetView] = useState('gate');
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [activePolicyModal, setActivePolicyModal] = useState(null);

  const resultsRef = useRef(null);
  const parsedPerCapita = Number(perCapitaOption);
  
  // CORE DUAL-STREAM MASS BALANCE CALCULATIONS
  const targetTotalTpd = ulbCalculationMode === 'population' 
    ? Number(((Number(population) * parsedPerCapita) / 1000000).toFixed(2))
    : Number(Number(actualAverageTpd || 0).toFixed(2));

  const targetSegregatedTpd = Number((targetTotalTpd * (segregationRate / 100)).toFixed(2));
  const targetMixedTpd = Number((targetTotalTpd - targetSegregatedTpd).toFixed(2));

  const allocatedMixed = Number(facilities.filter(f => f.type === 'mixed_trommel').reduce((acc, f) => acc + Number(f.avgProcessing || 0), 0).toFixed(2));
  const allocatedSegregated = Number(facilities.filter(f => f.type !== 'mixed_trommel').reduce((acc, f) => acc + Number(f.avgProcessing || 0), 0).toFixed(2));

  const isMixedBalanced = Math.abs(targetMixedTpd - allocatedMixed) <= 0.02;
  const isSegBalanced = Math.abs(targetSegregatedTpd - allocatedSegregated) <= 0.02;
  const isMassBalanced = isMixedBalanced && isSegBalanced;

  let validationMsg = "✓ Dual-Stream Mass Balance 100% Validated";
  if (!isSegBalanced) validationMsg = `⚠️ Segregated Stream Imbalance (Target ${targetSegregatedTpd} vs Allocated ${allocatedSegregated})`;
  else if (!isMixedBalanced) validationMsg = `⚠️ Mixed Stream Imbalance (Target ${targetMixedTpd} vs Allocated ${allocatedMixed})`;

  const addFacility = () => {
    setFacilities([
      ...facilities,
      { id: `proc_${Date.now()}`, name: `New Processing Facility`, type: 'wet_compost', designCapacity: 5, avgProcessing: 0 }
    ]);
  };

  const removeFacility = (id) => {
    if (facilities.length > 1) setFacilities(facilities.filter(f => f.id !== id));
  };

  const updateFacility = (id, field, value) => {
    setFacilities(facilities.map(f => f.id === id ? { ...f, [field]: value } : f));
  };

  const autoBalanceAllocation = () => {
    const trommels = facilities.filter(f => f.type === 'mixed_trommel');
    const segregated = facilities.filter(f => f.type !== 'mixed_trommel');
    
    const totalTrommelCap = trommels.reduce((sum, f) => sum + (Number(f.designCapacity)||1), 0);
    const totalSegCap = segregated.reduce((sum, f) => sum + (Number(f.designCapacity)||1), 0);

    let updated = facilities.map(f => {
      if (f.type === 'mixed_trommel') {
        if (totalTrommelCap === 0) return { ...f, avgProcessing: 0 };
        const share = Number(((Number(f.designCapacity)||1) / totalTrommelCap) * targetMixedTpd).toFixed(2);
        return { ...f, avgProcessing: share };
      } else {
        if (totalSegCap === 0) return { ...f, avgProcessing: 0 };
        const share = Number(((Number(f.designCapacity)||1) / totalSegCap) * targetSegregatedTpd).toFixed(2);
        return { ...f, avgProcessing: share };
      }
    });

    // Fix minor floating point rounding errors to force exact 0.00 match
    const sumTrommel = updated.filter(f => f.type === 'mixed_trommel').reduce((s,f) => s + f.avgProcessing, 0);
    if (trommels.length > 0 && Math.abs(sumTrommel - targetMixedTpd) > 0.001) {
       const lastT = updated.findLast(f => f.type === 'mixed_trommel');
       lastT.avgProcessing = Number((lastT.avgProcessing + (targetMixedTpd - sumTrommel)).toFixed(2));
    }

    const sumSeg = updated.filter(f => f.type !== 'mixed_trommel').reduce((s,f) => s + f.avgProcessing, 0);
    if (segregated.length > 0 && Math.abs(sumSeg - targetSegregatedTpd) > 0.001) {
       const lastS = updated.findLast(f => f.type !== 'mixed_trommel');
       lastS.avgProcessing = Number((lastS.avgProcessing + (targetSegregatedTpd - sumSeg)).toFixed(2));
    }

    setFacilities(updated);
  };

  // v4 cache tag resets local storage completely to lock preview
  const getSessionKey = () => `crf_paid_INT_v4_${name.trim().toLowerCase().replace(/\s+/g, '_')}_${selectedMonths.join('_')}_${startYear}`;

  useEffect(() => {
    const rawData = localStorage.getItem(getSessionKey());
    if (rawData) {
      try {
        const parsed = JSON.parse(rawData);
        if (parsed.paid && (Date.now() - parsed.timestamp < 12 * 60 * 60 * 1000)) {
          setIsPaid(true); return;
        }
      } catch (e) {
        if (rawData === 'true') { setIsPaid(true); return; }
      }
    }
    setIsPaid(false);
  }, [name, selectedMonths, startYear]);

  const toggleMonth = (mId) => {
    if (selectedMonths.includes(mId)) {
      if (selectedMonths.length > 1) setSelectedMonths(selectedMonths.filter(m => m !== mId));
    } else {
      setSelectedMonths([...selectedMonths, mId].sort((a, b) => a - b));
    }
  };

  const getPricingDetails = () => {
    const count = selectedMonths.length;
    const freeMonths = Math.floor(count / 6);
    const billableMonths = count - freeMonths;
    const baseRate = 500;
    const baseTotal = billableMonths * baseRate;
    const finalTotalWithCharges = Math.round(baseTotal / (1 - 0.0236));
    return { count, freeMonths, billableMonths, baseTotal, total: finalTotalWithCharges };
  };

  const pricing = getPricingDetails();

  const handleGenerate = (e) => {
    e.preventDefault();
    if (!isMassBalanced) {
      alert(`Validation Error: ${validationMsg}\n\nClick the "Auto-Balance Mass Allocation" button to mathematically distribute waste properly.`);
      return;
    }

    let monthlyDataMap = {};

    selectedMonths.forEach((m) => {
      const days = new Date(startYear, m, 0).getDate();
      const seedString = `INTEGRATED-MASS-BALANCE-V4-${selectedState}-${name}-${startYear}-${m}-${targetTotalTpd}-${segregationRate}`;
      const random = mulberry32(cyrb128(seedString));
      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayName = new Date(startYear, m - 1, day).toLocaleDateString('en-US', { weekday: 'short' });

        let noise = 0.95 + random() * 0.10;
        const dailyGateTotal = Number((targetTotalTpd * noise).toFixed(3));
        
        // Split Daily Generation into Dual Streams based on Segregation %
        const dailySegregated = Number((dailyGateTotal * (segregationRate / 100)).toFixed(3));
        const dailyMixed = Number((dailyGateTotal - dailySegregated).toFixed(3));

        let facilityBreakdown = {};
        facilities.forEach((f) => {
          let fIntake = 0;
          let outputs = {};

          if (f.type === 'mixed_trommel') {
            const ratio = allocatedMixed > 0 ? f.avgProcessing / allocatedMixed : 0;
            fIntake = Number((dailyMixed * ratio).toFixed(3));
            outputs = {
              intake: fIntake,
              organicFines: Number((fIntake * 0.45).toFixed(3)),
              coarseRdf: Number((fIntake * 0.35).toFixed(3)),
              heavyInerts: Number((fIntake * 0.20).toFixed(3))
            };
          } else {
            const ratio = allocatedSegregated > 0 ? f.avgProcessing / allocatedSegregated : 0;
            fIntake = Number((dailySegregated * ratio).toFixed(3));
            
            if (f.type === 'wet_compost' || f.type === 'vermicompost') {
              outputs = {
                intake: fIntake,
                enzyme: Number((fIntake * 2.5).toFixed(2)),
                compostYield: Number((fIntake * 0.18).toFixed(3)),
                rejects: Number((fIntake * 0.05).toFixed(3))
              };
            } else if (f.type === 'dry_mrf') {
              outputs = {
                intake: fIntake,
                recyclables: Number((fIntake * 0.55).toFixed(3)),
                rdf: Number((fIntake * 0.35).toFixed(3)),
                inerts: Number((fIntake * 0.10).toFixed(3))
              };
            } else if (f.type === 'biomethanation') {
              outputs = {
                intake: fIntake,
                biogasGenerated: Number((fIntake * 65).toFixed(1)),
                digestate: Number((fIntake * 0.25).toFixed(3))
              };
            } else {
              outputs = {
                intake: fIntake,
                safeStorage: Number((fIntake * 0.90).toFixed(3)),
                dispatchedTsdf: Number((fIntake * 0.10).toFixed(3))
              };
            }
          }

          facilityBreakdown[f.id] = outputs;
        });

        logs.push({
          date: dateStr,
          dayName,
          totalIntake: dailyGateTotal,
          dailySegregated,
          dailyMixed,
          facilityBreakdown
        });
      }
      monthlyDataMap[m] = logs;
    });

    const rawData = localStorage.getItem(getSessionKey());
    let verifiedPaid = false;
    if (rawData) {
      try {
        const parsed = JSON.parse(rawData);
        if (parsed.paid && (Date.now() - parsed.timestamp < 12 * 60 * 60 * 1000)) verifiedPaid = true;
      } catch (e) { if (rawData === 'true') verifiedPaid = true; }
    }
    setIsPaid(verifiedPaid);

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);
    setActiveAssetView('gate');
    setTimeout(() => resultsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
  };

  const handlePayment = async () => {
    if (!phone || phone.length < 10) {
      alert(lang === 'hi' ? 'कृपया एक वैध 10-अंकों का मोबाइल नंबर दर्ज करें।' : 'Please enter a valid 10-digit mobile number.');
      return;
    }
    setIsProcessing(true);

    if (!window.Cashfree) {
      await new Promise((res) => {
        if (document.querySelector('script[src*="cashfree.com"]')) return res(true);
        const s = document.createElement('script');
        s.src = 'https://sdk.cashfree.com/js/v3/cashfree.js';
        s.onload = () => res(true);
        document.body.appendChild(s);
      });
    }

    try {
      const res = await fetch('/api/create-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ amount: pricing.total, customerName: name, customerPhone: phone })
      });

      const order = await res.json();
      if (!order.payment_session_id) throw new Error(order.message || 'Failed to initialize payment session.');

      const cashfree = window.Cashfree({ mode: import.meta.env.VITE_CASHFREE_MODE || 'production' });

      cashfree.checkout({
        paymentSessionId: order.payment_session_id,
        redirectTarget: '_modal'
      }).then((result) => {
        if (result.error) {
          alert('Payment Failed: ' + result.error.message);
          setIsProcessing(false);
        } else if (result.paymentDetails) {
          setIsPaid(true);
          setIsProcessing(false);
          localStorage.setItem(getSessionKey(), JSON.stringify({ paid: true, timestamp: Date.now() }));
          downloadExcel();
        }
      });
    } catch (err) {
      alert('Payment Error: ' + err.message);
      setIsProcessing(false);
    }
  };

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(Number(v || 0) * 1000) : Number(v || 0).toFixed(3);

  const downloadExcel = () => {
    if (!generatedMonthlyData) return;
    if (!isPaid) return alert('Payment verified flag missing. Please complete checkout to download.');

    try {
      const u = displayUnit === 'kg' ? 'kg' : 'Tons';
      const wb = XLSX.utils.book_new();

      selectedMonths.forEach((mId) => {
        const monthName = MONTHS.find(m => m.id === mId)?.shortEn || `M${mId}`;

        // 1. Master Gate Intake Sheet (Now Shows Dual Streams!)
        const gateHeaders = ["Date", "Day", `Total Gate Intake (${u})`, `Segregated Stream (${u})`, `Mixed Stream (${u})`, ...facilities.map(f => `${f.name} Allocated (${u})`)];
        const gateRows = generatedMonthlyData[mId].map(r => [
          r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.dailySegregated), formatVal(r.dailyMixed),
          ...facilities.map(f => formatVal(r.facilityBreakdown[f.id]?.intake))
        ]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_GateIntake`);

        // 2. Individual Dynamic Facility Sheets
        facilities.forEach((f, idx) => {
          let headers = [];
          let rows = [];

          if (f.type === 'wet_compost' || f.type === 'vermicompost') {
            headers = ["Date", "Day", `Organic Feed (${u})`, "Enzyme Dose (Liters)", `Compost Yield (${u})`, `Inert Rejects (${u})`];
            rows = generatedMonthlyData[mId].map(r => [
              r.date, r.dayName,
              formatVal(r.facilityBreakdown[f.id]?.intake),
              r.facilityBreakdown[f.id]?.enzyme,
              formatVal(r.facilityBreakdown[f.id]?.compostYield),
              formatVal(r.facilityBreakdown[f.id]?.rejects)
            ]);
          } else if (f.type === 'dry_mrf') {
            headers = ["Date", "Day", `Dry Feed (${u})`, `Sorted Recyclables (${u})`, `RDF Dispatched (${u})`, `Inerts (${u})`];
            rows = generatedMonthlyData[mId].map(r => [
              r.date, r.dayName,
              formatVal(r.facilityBreakdown[f.id]?.intake),
              formatVal(r.facilityBreakdown[f.id]?.recyclables),
              formatVal(r.facilityBreakdown[f.id]?.rdf),
              formatVal(r.facilityBreakdown[f.id]?.inerts)
            ]);
          } else if (f.type === 'mixed_trommel') {
            headers = ["Date", "Day", `Mixed Feed (${u})`, `Organic Fines (${u})`, `Coarse Screen RDF (${u})`, `Heavy Inerts (${u})`];
            rows = generatedMonthlyData[mId].map(r => [
              r.date, r.dayName,
              formatVal(r.facilityBreakdown[f.id]?.intake),
              formatVal(r.facilityBreakdown[f.id]?.organicFines),
              formatVal(r.facilityBreakdown[f.id]?.coarseRdf),
              formatVal(r.facilityBreakdown[f.id]?.heavyInerts)
            ]);
          } else if (f.type === 'biomethanation') {
            headers = ["Date", "Day", `Organic Feed (${u})`, "Biogas Generated (m³)", `Bio-Digestate (${u})`];
            rows = generatedMonthlyData[mId].map(r => [
              r.date, r.dayName,
              formatVal(r.facilityBreakdown[f.id]?.intake),
              r.facilityBreakdown[f.id]?.biogasGenerated,
              formatVal(r.facilityBreakdown[f.id]?.digestate)
            ]);
          } else {
            headers = ["Date", "Day", `Daily Intake (${u})`, `Safe Stored (${u})`, `TSDF Transfer (${u})`];
            rows = generatedMonthlyData[mId].map(r => [
              r.date, r.dayName,
              formatVal(r.facilityBreakdown[f.id]?.intake),
              formatVal(r.facilityBreakdown[f.id]?.safeStorage),
              formatVal(r.facilityBreakdown[f.id]?.dispatchedTsdf)
            ]);
          }

          const safeSheetName = `${monthName}_F${idx + 1}_${f.name.replace(/[^a-zA-Z0-9]/g, '')}`.substring(0, 31);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), safeSheetName);
        });
      });

      XLSX.writeFile(wb, `Integrated_Master_Suite_${name.replace(/\s+/g, '_')}.xlsx`);
    } catch (err) {
      alert('Excel Generation Error: ' + err.message);
    }
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  
  // EXPLICIT LOCK ENFORCEMENT - IF !isPaid, STRICTLY CUT TO FIRST 5 DAYS
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh', padding: '15px' }}>
      <div style={{ maxWidth: '1000px', margin: '0 auto' }}>
        
        {/* HEADER */}
        <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                <ShieldCheck size={12} /> MASS-BALANCE SWM ESTIMATION SUITE
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {lang === 'hi' ? 'एकीकृत 3-इन-1 मास्टर लॉग-बुक सुइट' : 'Integrated 3-in-1 Master Logbook Suite'}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* CROSS-LINK BANNER */}
        <div style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', padding: '12px 16px', borderRadius: '8px', marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
          <div>
            <h3 style={{ margin: 0, color: '#334155', fontSize: '14px' }}>{lang === 'hi' ? 'सिंगल-फैसिलिटी लॉग-बुक चाहिए?' : 'Need Single-Facility Logbooks?'}</h3>
            <p style={{ margin: 0, color: '#64748b', fontSize: '12px' }}>{lang === 'hi' ? 'केवल ₹100/माह में साधारण ULB/MRF जनरेटर खोलें।' : 'Use our standalone ULB, MRF, or Mixed waste tool starting at ₹100/mo.'}</p>
          </div>
          <a href="https://ulb-waste-generator.vercel.app/" style={{ textDecoration: 'none', padding: '8px 14px', background: '#334155', color: '#fff', borderRadius: '6px', fontSize: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}>
            <ArrowLeft size={14} /> Open Standalone App
          </a>
        </div>

        {/* BILINGUAL USER GUIDE CONTAINER */}
        <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd', padding: '16px', borderRadius: '8px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', borderBottom: '1px solid #cbd5e1', paddingBottom: '8px' }}>
            <h3 style={{ margin: 0, color: '#0369a1', fontSize: '15px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <BookOpen size={18} /> {lang === 'hi' ? 'लॉग-बुक जनरेटर उपयोग निर्देशिका (User Guide)' : 'User Guide & Operator Instructions'}
            </h3>
            <span style={{ fontSize: '11px', background: '#e0f2fe', color: '#0369a1', padding: '2px 8px', borderRadius: '12px', fontWeight: 'bold' }}>
              {lang === 'hi' ? 'सहायता गाइड' : 'Help Manual'}
            </span>
          </div>

          {lang === 'hi' ? (
            <div style={{ fontSize: '13px', color: '#0c4a6e', lineHeight: '1.6' }}>
              <p style={{ margin: '0 0 8px 0', fontWeight: 'bold', color: '#0284c7' }}>
                इस टूल का उपयोग करके 30-दिवसीय ऑडिट-रेडी SWM एक्सेल लॉग-बुक कैसे बनाएं:
              </p>
              <ol style={{ margin: 0, paddingLeft: '20px' }}>
                <li style={{ marginBottom: '6px' }}>
                  <strong>बुनियादी विवरण दर्ज करें:</strong> अपने राज्य का चयन करें, अपनी निकाय/प्लांट का नाम लिखें और 10-अंकों का मोबाइल नंबर दर्ज करें।
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>कचरा क्षमता मोड चुनें:</strong> 
                  <ul style={{ paddingLeft: '16px', margin: '2px 0' }}>
                    <li><em>जनसंख्या आधारित:</em> बिना तौल-कांटा के मानक दर (450 ग्राम/दिन) से ऑटोमैटिक कैलकुलेट करें।</li>
                    <li><em>वास्तविक TPD:</em> यदि आपके पास वास्तविक तौल का आंकड़ा है, तो दैनिक औसत TPD दर्ज करें।</li>
                  </ul>
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>स्रोत पृथक्करण दर (%):</strong> डोर-टू-डोर पृथक्कृत (Segregated) और मिश्रित (Mixed) कचरे का प्रतिशत सेट करें।
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>प्रोसेसिंग फैसिलिटीज और मास-बैलेंस सेट करें:</strong> अपनी सभी प्रोसेसिंग यूनिट्स (कम्पोस्ट, MRF, ट्रॉमेल, बायो-सीएनजी) जोड़ें। सुनिश्चित करें कि सभी यूनिट्स की औसत प्रोसेसिंग (Average Processing TPD) का योग कुल गेट कचरे के बराबर हो। ट्रॉमेल प्लांट हमेशा मिश्रित कचरा लेगा।
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>महीने चुनें:</strong> आवश्यकतानुसार महीने चुनें (हर 6ठा महीना बिल्कुल मुफ्त है)।
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>डेटासेट जनरेट करें:</strong> "Generate Master Dataset" पर क्लिक करें। पहले 5 दिनों का मुफ्त पूर्वावलोकन (Preview) देखें, फिर भुगतान पूरा करके पूरे महीने की Multi-Sheet Excel Workbook (.xlsx) डाउनलोड करें।
                </li>
              </ol>
            </div>
          ) : (
            <div style={{ fontSize: '13px', color: '#0c4a6e', lineHeight: '1.6' }}>
              <p style={{ margin: '0 0 8px 0', fontWeight: 'bold', color: '#0284c7' }}>
                How to generate your 30-day audit-ready SWM Excel logbooks step-by-step:
              </p>
              <ol style={{ margin: 0, paddingLeft: '20px' }}>
                <li style={{ marginBottom: '6px' }}>
                  <strong>Enter Basic Details:</strong> Select your State, type your ULB/Facility Name, and provide a valid 10-digit mobile number.
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>Choose Waste Calculation Mode:</strong>
                  <ul style={{ paddingLeft: '16px', margin: '2px 0' }}>
                    <li><em>Population Based:</em> Automatically estimates daily waste using municipal per-capita benchmarks (450 g/person/day).</li>
                    <li><em>Actual TPD:</em> Select this if you weigh incoming trucks at a weighbridge and have an exact daily average.</li>
                  </ul>
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>Set Segregation Rate (%):</strong> Use the slider to balance Segregated waste vs Mixed unsegregated waste.
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>Configure Processing Assets & Mass Balance:</strong> Add custom processing facilities. Any Mixed Waste Trommel will automatically process your mixed stream. Ensure the sum of Average Processing TPD exactly equals your Total Gate Generation.
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>Select Duration:</strong> Click month buttons to choose duration (every 6th month is free).
                </li>
                <li style={{ marginBottom: '6px' }}>
                  <strong>Preview & Export:</strong> Click "Generate Master Dataset" to review the first 5 days for free, then complete payment to download the full Multi-Sheet Excel workbook (`.xlsx`).
                </li>
              </ol>
            </div>
          )}
        </div>

        {/* MAIN FORM */}
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>Select State</label>
              <select style={inputStyle} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>ULB / Facility Name</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>Mobile Number</label>
              <input style={inputStyle} type="tel" maxLength={10} required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
            </div>
          </div>

          {/* GATE INTAKE & SEGREGATION BASIS */}
          <div style={{ background: '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #e2e8f0', marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '10px' }}>1. Master Gate Refuse Estimation & Segregation</strong>
            
            <div style={{ display: 'flex', gap: '15px', marginBottom: '10px', fontSize: '13px' }}>
              <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'population'} onChange={() => setUlbCalculationMode('population')} /> {lang === 'hi' ? 'जनसंख्या आधारित' : 'Population Based'}</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'actual'} onChange={() => setUlbCalculationMode('actual')} /> {lang === 'hi' ? 'वास्तविक TPD' : 'Actual Observed TPD'}</label>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '15px', alignItems: 'center' }}>
              {ulbCalculationMode === 'population' ? (
                <>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>Population (Approx.)</label>
                    <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
                  </div>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>Per Capita Rate</label>
                    <select style={inputStyle} value={perCapitaOption} onChange={(e) => setPerCapitaOption(e.target.value)}>
                      <option value="300">300 g/day</option>
                      <option value="450">450 g/day</option>
                      <option value="500">500 g/day</option>
                    </select>
                  </div>
                </>
              ) : (
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>Actual Waste Generation (TPD)</label>
                  <input style={inputStyle} type="number" step="0.1" value={actualAverageTpd} onChange={(e) => setActualAverageTpd(e.target.value)} />
                </div>
              )}

              <div>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#0f172a' }}>Source Segregation Rate: {segregationRate}%</label>
                <input type="range" min="0" max="100" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '6px' }} />
              </div>
            </div>

            {/* LIVE SEGREGATION BREAKDOWN DISPLAY */}
            <div style={{ display: 'flex', gap: '10px', marginTop: '12px' }}>
               <div style={{ flex: 1, background: '#ecfdf5', padding: '8px 12px', borderRadius: '6px', border: '1px solid #a7f3d0' }}>
                 <span style={{ fontSize: '11px', color: '#065f46', fontWeight: 'bold', display: 'block' }}>Total Gate Refuse</span>
                 <span style={{ fontSize: '16px', color: '#047857', fontWeight: '900' }}>{targetTotalTpd} TPD</span>
               </div>
               <div style={{ flex: 1, background: '#f0f9ff', padding: '8px 12px', borderRadius: '6px', border: '1px solid #bae6fd' }}>
                 <span style={{ fontSize: '11px', color: '#0369a1', fontWeight: 'bold', display: 'block' }}>Target Segregated ({segregationRate}%)</span>
                 <span style={{ fontSize: '16px', color: '#0284c7', fontWeight: '900' }}>{targetSegregatedTpd} TPD</span>
               </div>
               <div style={{ flex: 1, background: '#fffbeb', padding: '8px 12px', borderRadius: '6px', border: '1px solid #fde68a' }}>
                 <span style={{ fontSize: '11px', color: '#92400e', fontWeight: 'bold', display: 'block' }}>Target Mixed ({100 - segregationRate}%)</span>
                 <span style={{ fontSize: '16px', color: '#d97706', fontWeight: '900' }}>{targetMixedTpd} TPD</span>
               </div>
            </div>
          </div>

          {/* DYNAMIC MULTI-ASSET FACILITY CONFIGURATOR */}
          <div style={{ background: '#f8fafc', padding: '14px', borderRadius: '6px', border: '1px solid #cbd5e1', marginBottom: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexWrap: 'wrap', gap: '10px' }}>
              <div>
                <strong style={{ fontSize: '13px', color: '#0f172a' }}>2. Processing Facilities & Capacity Allocation</strong>
                <span style={{ display: 'block', fontSize: '11px', color: '#64748b' }}>Trommels will auto-route the Mixed Stream. Other facilities route Segregated Streams.</span>
              </div>

              {/* DUAL-STREAM Mass Balance Indicator Badge */}
              <div style={{ 
                padding: '6px 12px', borderRadius: '6px', fontWeight: 'bold', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '6px',
                background: isMassBalanced ? '#ecfdf5' : '#fef2f2',
                color: isMassBalanced ? '#047857' : '#dc2626',
                border: `1px solid ${isMassBalanced ? '#a7f3d0' : '#fca5a5'}`
              }}>
                {isMassBalanced ? <CheckCircle2 size={16} /> : <AlertTriangle size={16} />}
                {validationMsg}
              </div>
            </div>

            {/* COLUMN HEADERS */}
            <div style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr 1fr auto', gap: '8px', marginBottom: '6px', fontSize: '11px', fontWeight: 'bold', color: '#475569', paddingRight: '10px' }}>
              <span>Facility Name</span>
              <span>Processing Stream Type</span>
              <span>Design Capacity (TPD)</span>
              <span>Avg Processing (TPD)</span>
              <span></span>
            </div>

            {/* DYNAMIC ROWS */}
            {facilities.map((f) => {
              const isTrommel = f.type === 'mixed_trommel';
              const inputColor = isTrommel ? (isMixedBalanced ? '#cbd5e1' : '#f87171') : (isSegBalanced ? '#cbd5e1' : '#f87171');
              
              return (
                <div key={f.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                  <input type="text" value={f.name} onChange={(e) => updateFacility(f.id, 'name', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="Unit Name" />
                  
                  <select value={f.type} onChange={(e) => updateFacility(f.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                    <option value="wet_compost">🌱 Wet Waste Composting Pad</option>
                    <option value="vermicompost">🪱 Vermicomposting Pit</option>
                    <option value="dry_mrf">📦 Dry Material Recovery (MRF)</option>
                    <option value="biomethanation">⚡ Bio-methanation / Bio-CNG</option>
                    <option value="hazardous_sanitary">☣️ Domestic Hazardous & Sanitary</option>
                    <option value="mixed_trommel">⚙️ Mixed Waste Trommel Line</option>
                  </select>

                  <input type="number" step="0.1" value={f.designCapacity} onChange={(e) => updateFacility(f.id, 'designCapacity', Number(e.target.value))} style={{ ...inputStyle, marginTop: 0 }} placeholder="Cap TPD" />
                  <input type="number" step="0.1" value={f.avgProcessing} onChange={(e) => updateFacility(f.id, 'avgProcessing', Number(e.target.value))} style={{ ...inputStyle, marginTop: 0, fontWeight: 'bold', borderColor: inputColor }} placeholder="Avg TPD" />

                  {facilities.length > 1 && (
                    <button type="button" onClick={() => removeFacility(f.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={16} /></button>
                  )}
                </div>
              );
            })}

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '12px' }}>
              <button type="button" onClick={addFacility} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={14} /> Add Processing Facility
              </button>

              <button type="button" onClick={autoBalanceAllocation} style={{ padding: '6px 12px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Zap size={14} /> Auto-Balance Streams
              </button>
            </div>
          </div>

          {/* MONTH SELECTOR GRID */}
          <div style={{ marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px' }}>{lang === 'hi' ? `महीने चुनें (${pricing.count} चयनित — ₹${pricing.total}):` : `Select Months (${pricing.count} Selected — ₹${pricing.total}):`}</strong>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(75px, 1fr))', gap: '6px', marginTop: '6px' }}>
              {MONTHS.map((m) => (
                <button key={m.id} type="button" onClick={() => toggleMonth(m.id)} style={{ padding: '6px 2px', borderRadius: '4px', border: selectedMonths.includes(m.id) ? '2px solid #047857' : '1px solid #cbd5e1', background: selectedMonths.includes(m.id) ? '#ecfdf5' : '#fff', fontWeight: selectedMonths.includes(m.id) ? 'bold' : 'normal', cursor: 'pointer', fontSize: '12px' }}>
                  {lang === 'hi' ? m.shortHi : m.shortEn}
                </button>
              ))}
            </div>
          </div>

          <button type="submit" style={{ width: '100%', padding: '14px', background: isMassBalanced ? '#047857' : '#94a3b8', color: '#fff', border: 'none', borderRadius: '6px', cursor: isMassBalanced ? 'pointer' : 'not-allowed', fontWeight: 'bold', fontSize: '15px' }}>
            Generate Master Dataset (₹{pricing.total}) →
          </button>
        </form>

        {/* PREVIEW CONTAINER WITH ASSET TABS */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            
            {/* MONTH TABS */}
            <div style={{ display: 'flex', gap: '6px', marginBottom: '12px', overflowX: 'auto', borderBottom: '1px solid #e2e8f0', paddingBottom: '6px' }}>
              {selectedMonths.map(mId => (
                <button key={mId} type="button" onClick={() => setActiveTabMonth(mId)} style={{ padding: '6px 12px', borderRadius: '4px', border: activeTabMonth === mId ? '2px solid #047857' : '1px solid #cbd5e1', background: activeTabMonth === mId ? '#047857' : '#f8fafc', color: activeTabMonth === mId ? '#fff' : '#334155', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px' }}>
                  {MONTHS.find(m => m.id === mId)?.fullEn}
                </button>
              ))}
            </div>

            {/* ASSET SELECTOR SWITCHER */}
            <div style={{ display: 'flex', gap: '8px', marginBottom: '12px', flexWrap: 'wrap', background: '#f1f5f9', padding: '8px', borderRadius: '6px' }}>
              <button type="button" onClick={() => setActiveAssetView('gate')} style={{ padding: '6px 10px', borderRadius: '4px', border: 'none', background: activeAssetView === 'gate' ? '#0f172a' : 'transparent', color: activeAssetView === 'gate' ? '#fff' : '#475569', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px' }}>
                🏢 Master Gate Intake Sheet
              </button>

              {facilities.map(f => (
                <button key={f.id} type="button" onClick={() => setActiveAssetView(f.id)} style={{ padding: '6px 10px', borderRadius: '4px', border: 'none', background: activeAssetView === f.id ? '#047857' : 'transparent', color: activeAssetView === f.id ? '#fff' : '#047857', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px' }}>
                  ⚙️ {f.name}
                </button>
              ))}
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong>{name} — Sheet Preview</strong>
              {isPaid ? (
                <button onClick={downloadExcel} style={{ padding: '6px 12px', background: '#047857', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <FileSpreadsheet size={14} /> Download Excel Workbook
                </button>
              ) : (
                <span style={{ fontSize: '12px', color: '#dc2626', fontWeight: 'bold' }}>🔒 Locked Preview (Days 1–5 Only)</span>
              )}
            </div>

            {/* PREVIEW TABLE */}
            <div onContextMenu={(e) => !isPaid && e.preventDefault()} style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', userSelect: isPaid ? 'text' : 'none' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '600px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1', textAlign: 'left' }}>
                    <th>Date</th><th>Day</th>
                    {activeAssetView === 'gate' && (
                      <>
                        <th style={{ textAlign: 'right' }}>Total Gate Refuse</th>
                        <th style={{ textAlign: 'right' }}>Segregated Stream</th>
                        <th style={{ textAlign: 'right' }}>Mixed Stream</th>
                        {facilities.map(f => <th key={f.id} style={{ textAlign: 'right' }}>{f.name}</th>)}
                      </>
                    )}
                    {facilities.some(f => f.id === activeAssetView) && (
                      <>
                        <th style={{ textAlign: 'right' }}>Facility Feed (Tons)</th>
                        <th style={{ textAlign: 'right' }}>Primary Output</th>
                        <th style={{ textAlign: 'right' }}>Secondary Fraction</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>

                      {activeAssetView === 'gate' && (
                        <>
                          <td style={{ textAlign: 'right' }}><strong>{formatVal(r.totalIntake)}</strong></td>
                          <td style={{ textAlign: 'right', color: '#0284c7' }}>{formatVal(r.dailySegregated)}</td>
                          <td style={{ textAlign: 'right', color: '#d97706' }}>{formatVal(r.dailyMixed)}</td>
                          {facilities.map(f => (
                            <td key={f.id} style={{ textAlign: 'right' }}>{formatVal(r.facilityBreakdown[f.id]?.intake)}</td>
                          ))}
                        </>
                      )}

                      {facilities.some(f => f.id === activeAssetView) && (
                        <>
                          <td style={{ textAlign: 'right' }}><strong>{formatVal(r.facilityBreakdown[activeAssetView]?.intake)}</strong></td>
                          <td style={{ textAlign: 'right' }}>
                            {formatVal(r.facilityBreakdown[activeAssetView]?.compostYield || r.facilityBreakdown[activeAssetView]?.recyclables || r.facilityBreakdown[activeAssetView]?.organicFines || r.facilityBreakdown[activeAssetView]?.biogasGenerated || r.facilityBreakdown[activeAssetView]?.safeStorage)}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            {formatVal(r.facilityBreakdown[activeAssetView]?.rejects || r.facilityBreakdown[activeAssetView]?.rdf || r.facilityBreakdown[activeAssetView]?.coarseRdf || r.facilityBreakdown[activeAssetView]?.digestate || r.facilityBreakdown[activeAssetView]?.dispatchedTsdf)}
                          </td>
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {!isPaid && (
              <div style={{ border: '2px dashed #047857', background: '#ecfdf5', padding: '15px', textAlign: 'center', marginTop: '12px', borderRadius: '6px' }}>
                <Lock style={{ color: '#047857' }} size={18} />
                <h4 style={{ margin: '4px 0', color: '#065f46' }}>Preview Locked (Days 1–5 Only)</h4>
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#047857', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '14px', fontWeight: 'bold', cursor: 'pointer', marginTop: '8px' }}>
                  {isProcessing ? 'Connecting...' : `Pay ₹${pricing.total} & Download File`}
                </button>
              </div>
            )}
          </div>
        )}

        {/* COMPLIANCE FOOTER */}
        <footer style={{ marginTop: '40px', paddingTop: '20px', borderTop: '1px solid #cbd5e1', textAlign: 'center', fontSize: '12px', color: '#64748b' }}>
          <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', flexWrap: 'wrap', marginBottom: '10px' }}>
            <button type="button" onClick={() => setActivePolicyModal('contact')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Contact Us</button> |
            <button type="button" onClick={() => setActivePolicyModal('terms')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Terms & Conditions</button> |
            <button type="button" onClick={() => setActivePolicyModal('refunds')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Refunds & Cancellations</button> |
            <button type="button" onClick={() => setActivePolicyModal('pricing')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Services & Pricing (INR)</button>
          </div>
          <p style={{ margin: 0 }}>© {new Date().getFullYear()} Consilience Research Foundation / SWM Suite. All Rights Reserved.</p>
        </footer>

        {/* POLICY MODAL */}
        {activePolicyModal && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000, padding: '15px' }}>
            <div style={{ background: '#fff', padding: '24px', borderRadius: '8px', maxWidth: '600px', width: '100%', maxHeight: '80vh', overflowY: 'auto' }}>
              {activePolicyModal === 'contact' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Contact Us</h2>
                  <p style={{ fontSize: '13px' }}><strong>Organisation:</strong> Consilience Research Foundation</p>
                  <p style={{ fontSize: '13px' }}><strong>Address:</strong> Arjunganj, Lucknow, Uttar Pradesh, India</p>
                  <p style={{ fontSize: '13px' }}><strong>Email:</strong> support@consilience.res.in</p>
                </div>
              )}
              {activePolicyModal === 'terms' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Terms & Conditions</h2>
                  <p style={{ fontSize: '12px' }}>This tool provides engineered multi-asset estimations for solid waste management facilities.</p>
                </div>
              )}
              {activePolicyModal === 'refunds' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Refunds & Cancellations</h2>
                  <p style={{ fontSize: '12px' }}>Digital Excel files are unlocked instantly upon payment confirmation. Failed unlocks after debit are refunded in 5-7 business days.</p>
                </div>
              )}
              {activePolicyModal === 'pricing' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Services & Pricing (INR)</h2>
                  <ul style={{ fontSize: '12px', lineHeight: '1.8' }}>
                    <li>Integrated 3-in-1 Master Suite: ₹500 / Month (Includes Mass Balance Gate & Dynamic Processing Tabs)</li>
                  </ul>
                </div>
              )}
              <button type="button" onClick={() => setActivePolicyModal(null)} style={{ marginTop: '15px', padding: '8px 16px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Close</button>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
