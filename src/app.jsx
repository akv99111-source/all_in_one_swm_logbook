import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck, MapPin, AlertCircle, Phone, Plus, Trash2, ArrowLeft, Layers, Sparkles, BookOpen } from 'lucide-react';
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
  }
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
  const [actualAverageTpd, setActualAverageTpd] = useState(10);
  const [segregationRate, setSegregationRate] = useState(80);

  // Dynamic Assets for Integrated 3-in-1 Mode
  const [compostUnits, setCompostUnits] = useState([
    { id: 'c1', label: 'Windrow Pad Alpha', type: 'Windrow Pad', capacity: 10 }
  ]);
  const [mrfUnits, setMrfUnits] = useState([
    { id: 'm1', label: 'MRF Shed 1', type: 'Manual Sorting Shed', capacity: 5 }
  ]);
  
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const resultsRef = useRef(null);
  const parsedPerCapita = Number(perCapitaOption);
  
  // TPD Calculation for display
  const calculatedTpdDisplay = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);

  // Dynamic Handlers
  const addCompostUnit = () => {
    setCompostUnits([...compostUnits, { id: `c${Date.now()}`, label: lang === 'hi' ? `नई कम्पोस्ट यूनिट` : `New Compost Unit`, type: 'Windrow Pad', capacity: 5 }]);
  };

  const removeCompostUnit = (id) => {
    if (compostUnits.length > 1) setCompostUnits(compostUnits.filter(u => u.id !== id));
  };

  const updateCompostUnit = (id, field, value) => {
    setCompostUnits(compostUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

  const addMrfUnit = () => {
    setMrfUnits([...mrfUnits, { id: `m${Date.now()}`, label: lang === 'hi' ? `नया एमआरएफ शेड` : `New MRF Shed`, type: 'Manual Sorting Shed', capacity: 5 }]);
  };

  const removeMrfUnit = (id) => {
    if (mrfUnits.length > 1) setMrfUnits(mrfUnits.filter(u => u.id !== id));
  };

  const updateMrfUnit = (id, field, value) => {
    setMrfUnits(mrfUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

  // Session Management
  const getSessionKey = () => {
    const cleanName = name.trim().toLowerCase().replace(/\s+/g, '_');
    return `crf_paid_INTEGRATED_${cleanName}_${selectedMonths.length}M`;
  };

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
  }, [name, selectedMonths.length]);

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
    const baseRate = 500; // Flat ₹500 for integrated suite
    
    const baseTotal = billableMonths * baseRate;
    const effectiveFeeRate = 0.0236; 
    const finalTotalWithCharges = Math.round(baseTotal / (1 - effectiveFeeRate));

    return { count, freeMonths, billableMonths, baseRate, baseTotal, total: finalTotalWithCharges };
  };

  const pricing = getPricingDetails();

  const handleGenerate = (e) => {
    e.preventDefault();
    let monthlyDataMap = {};

    selectedMonths.forEach((m) => {
      const days = new Date(startYear, m, 0).getDate();
      let targetTons = ulbCalculationMode === 'population' 
        ? (Number(population) * parsedPerCapita) / 1000000 
        : Number(actualAverageTpd);

      const seedString = `INTEGRATED-3IN1-${selectedState}-${name}-${startYear}-${m}-${ulbCalculationMode}-${targetTons}-${segregationRate}`;
      const random = mulberry32(cyrb128(seedString));
      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayName = new Date(startYear, m - 1, day).toLocaleDateString('en-US', { weekday: 'short' });

        let noise = 0.95 + random() * 0.10;
        const dailyTotal = targetTons * noise;

        const segFrac = segregationRate / 100;
        const unsegFrac = 1 - segFrac;

        const segregatedTotal = dailyTotal * segFrac;
        const unsegregatedMixed = Number((dailyTotal * unsegFrac).toFixed(3));
        
        const wetSeg = Number((segregatedTotal * 0.60).toFixed(3));
        const drySeg = Number((segregatedTotal * 0.32).toFixed(3));
        const hazSeg = Number((segregatedTotal * 0.03).toFixed(3));
        const sanSeg = Number((segregatedTotal * 0.05).toFixed(3));

        const organicFines = Number((unsegregatedMixed * 0.45).toFixed(3));
        const dryOversize = Number((unsegregatedMixed * 0.35).toFixed(3));
        const heavyInerts = Number((unsegregatedMixed * 0.20).toFixed(3));

        // Compost Mass Balance Distribution
        const totalCompostFeed = wetSeg + organicFines;
        const totalCompostCapacity = compostUnits.reduce((acc, u) => acc + Number(u.capacity || 1), 0);
        let compostUnitBreakdown = {};
        compostUnits.forEach(unit => {
          const unitShare = (Number(unit.capacity || 1) / totalCompostCapacity) * totalCompostFeed;
          compostUnitBreakdown[unit.id] = {
            feed: Number(unitShare.toFixed(3)),
            compostYield: Number((unitShare * 0.18).toFixed(3))
          };
        });

        // MRF Mass Balance Distribution
        const totalMrfFeed = drySeg + dryOversize;
        const totalMrfCapacity = mrfUnits.reduce((acc, u) => acc + Number(u.capacity || 1), 0);
        let mrfUnitBreakdown = {};
        mrfUnits.forEach(unit => {
          const unitShare = (Number(unit.capacity || 1) / totalMrfCapacity) * totalMrfFeed;
          mrfUnitBreakdown[unit.id] = {
            feed: Number(unitShare.toFixed(3)),
            recyclables: Number((unitShare * 0.65).toFixed(3)),
            rdf: Number((unitShare * 0.25).toFixed(3))
          };
        });

        logs.push({
          date: dateStr, dayName, totalIntake: Number(dailyTotal.toFixed(3)),
          wetSeg, drySeg, hazSeg, sanSeg, unsegregatedMixed,
          organicFines, dryOversize, heavyInerts,
          totalCompostFeed: Number(totalCompostFeed.toFixed(3)),
          totalMrfFeed: Number(totalMrfFeed.toFixed(3)),
          compostUnitBreakdown, mrfUnitBreakdown
        });
      }
      monthlyDataMap[m] = logs;
    });

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);
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

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(v * 1000) : Number(v || 0).toFixed(3);

  const downloadExcel = () => {
    if (!generatedMonthlyData) return;
    const u = displayUnit === 'kg' ? 'kg' : 'Tons';
    const wb = XLSX.utils.book_new();

    selectedMonths.forEach((mId) => {
      const monthName = MONTHS.find(m => m.id === mId)?.fullEn;

      // Tab 1: Gate Register
      const gateHeaders = ["Date", "Day", `Total Gate Intake (${u})`, `Segregated Wet (${u})`, `Segregated Dry (${u})`, `Domestic Hazardous (${u})`, `Domestic Sanitary (${u})`, `Unsegregated Mixed (${u})`];
      const gateRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.wetSeg), formatVal(r.drySeg), formatVal(r.hazSeg), formatVal(r.sanSeg), formatVal(r.unsegregatedMixed)]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_Gate`);

      // Tab 2: Pre-Sorting (Mixed Waste Trommel)
      const preHeaders = ["Date", "Day", `Mixed Intake (${u})`, `Fine Screen Fraction (${u})`, `Coarse Screen Fraction (${u})`, `Heavy Inerts (${u})`];
      const preRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.unsegregatedMixed), formatVal(r.organicFines), formatVal(r.dryOversize), formatVal(r.heavyInerts)]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([preHeaders, ...preRows]), `${monthName}_PreSort`);

      // Tabs 3+: Dynamic Compost Units
      compostUnits.forEach(unit => {
        const cHeaders = ["Date", "Day", `Unit Feed (${u})`, `Compost Yield (${u})`];
        const cRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.compostUnitBreakdown[unit.id]?.feed), formatVal(r.compostUnitBreakdown[unit.id]?.compostYield)]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([cHeaders, ...cRows]), `${monthName}_${unit.label.substring(0, 10)}`);
      });

      // Tabs 4+: Dynamic MRF Units
      mrfUnits.forEach(unit => {
        const mHeaders = ["Date", "Day", `Unit Feed (${u})`, `Sorted Recyclables (${u})`, `RDF Dispatched (${u})`];
        const mRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.mrfUnitBreakdown[unit.id]?.feed), formatVal(r.mrfUnitBreakdown[unit.id]?.recyclables), formatVal(r.mrfUnitBreakdown[unit.id]?.rdf)]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([mHeaders, ...mRows]), `${monthName}_${unit.label.substring(0, 10)}`);
      });
    });

    XLSX.writeFile(wb, `Integrated_Suite_${name.replace(/\s+/g, '_')}.xlsx`);
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh', padding: '15px' }}>
      <div style={{ maxWidth: '1000px', margin: '0 auto' }}>
        
        {/* HEADER BANNER */}
        <div style={{ background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                <Layers size={12} /> {lang === 'hi' ? 'एडवांस्ड SWM सुइट' : 'ADVANCED SWM SUITE'}
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {lang === 'hi' ? 'एकीकृत 3-इन-1 मास्टर लॉग-बुक जनरेटर' : 'Integrated 3-in-1 Master Logbook Generator'}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#0f172a', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* HELP TEXT BOX */}
        <div style={{ background: '#f0f9ff', border: '1px solid #7dd3fc', padding: '12px 16px', borderRadius: '8px', marginBottom: '16px', display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
          <BookOpen style={{ color: '#0284c7', marginTop: '2px' }} size={20} />
          <div>
            <strong style={{ color: '#0369a1', fontSize: '14px', display: 'block', marginBottom: '4px' }}>
              {lang === 'hi' ? 'मल्टी-यूनिट सुइट का उपयोग कैसे करें?' : 'How to use the Multi-Unit Suite?'}
            </strong>
            <p style={{ margin: 0, fontSize: '13px', color: '#0c4a6e', lineHeight: '1.5' }}>
              {lang === 'hi' 
                ? '1. शहर की क्षमता दर्ज करें। 2. अपनी आवश्यकता के अनुसार कस्टम कम्पोस्ट यूनिट्स और MRF शेड जोड़ें। 3. सेग्रीगेशन दर सेट करें। 4. सुरक्षित भुगतान करके गेट, प्री-सॉर्टिंग, कम्पोस्ट, और MRF के लिए मल्टी-टैब एक्सेल डाउनलोड करें।' 
                : '1. Enter city capacity. 2. Dynamically add Custom Compost Units and MRF Sheds based on your infrastructure. 3. Set Segregation Rate. 4. Pay to download a multi-tab Excel for Gate, Pre-sorting, Compost, and MRF.'}
            </p>
          </div>
        </div>

        {/* EXTERNAL LINK BANNER TO STANDALONE APP */}
        <div style={{ background: 'linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)', border: '1px solid #cbd5e1', padding: '14px 18px', borderRadius: '8px', marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
          <div>
            <h3 style={{ margin: '0 0 2px 0', color: '#334155', fontSize: '14px', fontWeight: 'bold' }}>
              {lang === 'hi' ? 'क्या आप केवल सिंगल-फैसिलिटी (ULB/MRF) लॉग-बुक खोज रहे हैं?' : 'Looking for a Single-Facility (ULB/MRF) Logbook?'}
            </h3>
            <p style={{ margin: 0, color: '#475569', fontSize: '12px' }}>
              {lang === 'hi' ? 'केवल ₹100 में साधारण स्टैंडअलोन जनरेटर का उपयोग करें।' : 'Use the simple standalone generator starting at ₹100/mo.'}
            </p>
          </div>
          {/* UPDATE THIS URL TO POINT TO YOUR STANDALONE DEPLOYMENT */}
          <a href="https://your-standalone-app-link.com" target="_blank" rel="noopener noreferrer" style={{ textDecoration: 'none', padding: '8px 16px', background: '#334155', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '6px' }}>
            <ArrowLeft size={14} /> {lang === 'hi' ? 'स्टैंडअलोन ऐप पर जाएँ' : 'Go to Standalone App'}
          </a>
        </div>

        {/* FORM SECTION */}
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'राज्य चुनें' : 'Select State'}</label>
              <select style={inputStyle} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'निकाय / प्लांट का नाम' : 'City / Facility Name'}</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'मोबाइल नंबर' : 'Mobile Number'}</label>
              <input style={inputStyle} type="tel" maxLength={10} placeholder="9876543210" required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
            </div>

            <div style={{ gridColumn: '1 / -1', background: '#f1f5f9', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
              <strong style={{ fontSize: '13px' }}>{lang === 'hi' ? 'कचरा उत्पादन का आधार (संपूर्ण प्लांट)' : 'Total Facility Waste Estimation Basis'}</strong>
              <div style={{ display: 'flex', gap: '15px', marginTop: '6px', fontSize: '13px' }}>
                <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'population'} onChange={() => setUlbCalculationMode('population')} /> {lang === 'hi' ? 'जनसंख्या आधारित' : 'Population Based'}</label>
                <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'actual'} onChange={() => setUlbCalculationMode('actual')} /> {lang === 'hi' ? 'वास्तविक TPD (देखा गया)' : 'Actual TPD'}</label>
              </div>
            </div>

            {ulbCalculationMode === 'population' ? (
              <>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'अनुमानित जनसंख्या' : 'Population'}</label>
                  <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'प्रति व्यक्ति दर' : 'Per Capita Rate'}</label>
                  <select style={inputStyle} value={perCapitaOption} onChange={(e) => setPerCapitaOption(e.target.value)}>
                    <option value="300">300 {lang === 'hi' ? 'ग्राम/दिन' : 'g/day'}</option>
                    <option value="450">450 {lang === 'hi' ? 'ग्राम/दिन' : 'g/day'}</option>
                    <option value="500">500 {lang === 'hi' ? 'ग्राम/दिन' : 'g/day'}</option>
                  </select>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', background: '#ecfdf5', padding: '10px', borderRadius: '6px', border: '1px solid #a7f3d0' }}>
                  <span style={{ fontSize: '11px', color: '#065f46', fontWeight: 'bold' }}>{lang === 'hi' ? 'कुल अनुमानित गेट कचरा' : 'Calculated Gate Waste'}</span>
                  <span style={{ fontSize: '18px', color: '#047857', fontWeight: '900' }}>{calculatedTpdDisplay} TPD</span>
                </div>
              </>
            ) : (
              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'कुल गेट आवक (TPD)' : 'Total Gate Intake (TPD)'}</label>
                <input style={inputStyle} type="number" value={actualAverageTpd} onChange={(e) => setActualAverageTpd(e.target.value)} />
              </div>
            )}
          </div>

          <div style={{ background: '#fef3c7', padding: '12px', borderRadius: '6px', border: '1px solid #fde68a', marginBottom: '14px' }}>
            <label style={{ fontSize: '13px', fontWeight: 'bold', color: '#92400e' }}>
              {lang === 'hi' ? `स्रोत पर कचरा पृथक्करण दर (%): ${segregationRate}%` : `Source Segregation Rate (%): ${segregationRate}%`}
            </label>
            <input type="range" min="20" max="95" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '8px' }} />
            <p style={{ margin: '4px 0 0 0', fontSize: '11px', color: '#b45309' }}>
              {lang === 'hi' ? 'यह स्लाइडर प्लांट में आने वाले मिश्रित, गीले और सूखे कचरे का मास-बैलेंस तय करेगा।' : 'This determines the mass-balance of Mixed vs Segregated waste entering the facility.'}
            </p>
          </div>

          {/* DYNAMIC COMPOST UNITS */}
          <div style={{ background: '#f0fdf4', padding: '12px', borderRadius: '6px', border: '1px solid #bbf7d0', marginBottom: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <strong style={{ fontSize: '13px', color: '#166534' }}>{lang === 'hi' ? 'कम्पोस्टिंग / गीला कचरा यूनिट्स' : 'Composting Assets (Wet Waste Line)'}</strong>
              <button type="button" onClick={addCompostUnit} style={{ padding: '6px 10px', background: '#166534', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={14} /> {lang === 'hi' ? 'यूनिट जोड़ें' : 'Add Unit'}
              </button>
            </div>
            {compostUnits.map((u) => (
              <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                <input type="text" value={u.label} onChange={(e) => updateCompostUnit(u.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder={lang === 'hi' ? 'यूनिट का नाम' : 'Unit Label'} />
                <select value={u.type} onChange={(e) => updateCompostUnit(u.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                  <option value="Windrow Pad">{lang === 'hi' ? 'विंड्रो कम्पोस्ट पैड' : 'Windrow Pad'}</option>
                  <option value="Vermicompost Pit">{lang === 'hi' ? 'वर्मीकम्पोस्ट (केंचुआ) पिट' : 'Vermicompost Pit'}</option>
                  <option value="Biomethanation">{lang === 'hi' ? 'बायोमेथेनेशन प्लांट' : 'Biomethanation Plant'}</option>
                </select>
                <input type="number" value={u.capacity} onChange={(e) => updateCompostUnit(u.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="TPD" />
                {compostUnits.length > 1 && <button type="button" onClick={() => removeCompostUnit(u.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer', padding: '6px' }}><Trash2 size={16} /></button>}
              </div>
            ))}
          </div>

          {/* DYNAMIC MRF UNITS */}
          <div style={{ background: '#f0f9ff', padding: '12px', borderRadius: '6px', border: '1px solid #bae6fd', marginBottom: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <strong style={{ fontSize: '13px', color: '#0369a1' }}>{lang === 'hi' ? 'एमआरएफ / सूखा कचरा सॉर्टिंग शेड' : 'MRF Assets (Dry Waste Line)'}</strong>
              <button type="button" onClick={addMrfUnit} style={{ padding: '6px 10px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={14} /> {lang === 'hi' ? 'शेड जोड़ें' : 'Add Shed'}
              </button>
            </div>
            {mrfUnits.map((u) => (
              <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                <input type="text" value={u.label} onChange={(e) => updateMrfUnit(u.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder={lang === 'hi' ? 'शेड का नाम' : 'Unit Label'} />
                <select value={u.type} onChange={(e) => updateMrfUnit(u.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                  <option value="Manual Sorting Shed">{lang === 'hi' ? 'मैनुअल सॉर्टिंग शेड' : 'Manual Sorting Shed'}</option>
                  <option value="Semi-Automated Line">{lang === 'hi' ? 'सेमी-ऑटोमेटेड मशीन' : 'Semi-Automated Line'}</option>
                  <option value="Baling Unit">{lang === 'hi' ? 'बेलिंग (Baling) यूनिट' : 'Baling Unit'}</option>
                </select>
                <input type="number" value={u.capacity} onChange={(e) => updateMrfUnit(u.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="TPD" />
                {mrfUnits.length > 1 && <button type="button" onClick={() => removeMrfUnit(u.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer', padding: '6px' }}><Trash2 size={16} /></button>}
              </div>
            ))}
          </div>

          <div style={{ marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px', display: 'flex', alignItems: 'center', gap: '10px' }}>
              {lang === 'hi' ? `महीने चुनें (${pricing.count} चयनित — ₹${pricing.total}):` : `Select Months (${pricing.count} Selected — ₹${pricing.total}):`}
              {pricing.freeMonths > 0 && (
                <span style={{ color: '#dc2626', fontSize: '11px', background: '#fee2e2', padding: '3px 8px', borderRadius: '12px' }}>
                  {lang === 'hi' ? `🎉 ${pricing.freeMonths} महीना मुफ़्त!` : `🎉 ${pricing.freeMonths} Month Free!`}
                </span>
              )}
            </strong>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(75px, 1fr))', gap: '6px', marginTop: '6px' }}>
              {MONTHS.map((m) => {
                const active = selectedMonths.includes(m.id);
                return (
                  <button key={m.id} type="button" onClick={() => toggleMonth(m.id)} style={{ padding: '6px 2px', borderRadius: '4px', border: active ? '2px solid #0f172a' : '1px solid #cbd5e1', background: active ? '#f1f5f9' : '#fff', fontWeight: active ? 'bold' : 'normal', cursor: 'pointer', fontSize: '12px' }}>
                    {lang === 'hi' ? m.shortHi : m.shortEn}
                  </button>
                );
              })}
            </div>
          </div>

          <button type="submit" style={{ width: '100%', padding: '14px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            {lang === 'hi' 
              ? `मल्टी-यूनिट मास्टर लॉग जनरेट करें (₹${pricing.total}) →` 
              : `Generate Multi-Unit Master Log (₹${pricing.total}) →`}
          </button>
        </form>

        {/* RESULTS PREVIEW */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexWrap: 'wrap', gap: '10px' }}>
              <strong style={{ fontSize: '14px' }}>{name} — {lang === 'hi' ? 'प्लांट गेट डेटासेट प्रीव्यू' : 'Facility Gate Dataset Preview'}</strong>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1', cursor: 'pointer' }}>
                  {lang === 'hi' ? 'यूनिट:' : 'Unit:'} <strong>{displayUnit}</strong>
                </button>
                {isPaid && (
                  <button onClick={downloadExcel} style={{ padding: '6px 12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    <Download size={13} /> Export Master Excel
                  </button>
                )}
              </div>
            </div>

            <div onContextMenu={(e) => !isPaid && e.preventDefault()} style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', userSelect: isPaid ? 'text' : 'none' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '600px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1', textAlign: 'left' }}>
                    <th>Date</th><th>Day</th>
                    <th>{lang === 'hi' ? 'कुल गेट आवक' : 'Gate Intake'}</th>
                    <th>{lang === 'hi' ? 'अलग किया गीला' : 'Seg. Wet'}</th>
                    <th>{lang === 'hi' ? 'अलग किया सूखा' : 'Seg. Dry'}</th>
                    <th>{lang === 'hi' ? 'मिश्रित कचरा' : 'Mixed Waste'}</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>
                      <td><strong>{formatVal(r.totalIntake)}</strong></td>
                      <td>{formatVal(r.wetSeg)}</td>
                      <td>{formatVal(r.drySeg)}</td>
                      <td>{formatVal(r.unsegregatedMixed)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {!isPaid && (
              <div style={{ border: '2px dashed #0f172a', background: '#f8fafc', padding: '15px', textAlign: 'center', marginTop: '12px', borderRadius: '6px' }}>
                <Lock style={{ color: '#0f172a' }} size={18} />
                <h4 style={{ margin: '4px 0', color: '#1e293b', fontSize: '15px' }}>
                  {lang === 'hi' ? 'प्रीव्यू लॉक है (केवल 1-5 दिन दिख रहे हैं)' : 'Preview Locked (Days 1–5 Only)'}
                </h4>
                <p style={{ margin: '4px 0 10px 0', color: '#475569', fontSize: '13px' }}>
                  {lang === 'hi' 
                    ? `Gate, Pre-sorting, Compost और MRF टैब सहित पूरी मास्टर एक्सेल डाउनलोड करने के लिए ₹${pricing.total} का भुगतान करें।` 
                    : `Pay ₹${pricing.total} to download your complete Multi-Tab Master Excel (Includes Gate, Pre-Sort, Compost & MRF sheets).`}
                </p>
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '14px', fontWeight: 'bold', cursor: 'pointer' }}>
                  {isProcessing 
                    ? (lang === 'hi' ? 'प्रक्रिया जारी है...' : 'Connecting...') 
                    : (lang === 'hi' ? `₹${pricing.total} सुरक्षित भुगतान करें` : `Pay ₹${pricing.total} Securely`)}
                </button>
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
}