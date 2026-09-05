import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, ShieldCheck, Plus, Trash2, BookOpen, Layers } from 'lucide-react';
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
  const [actualAverageTpd, setActualAverageTpd] = useState(10);
  const [segregationRate, setSegregationRate] = useState(80);

  const [compostUnits, setCompostUnits] = useState([
    { id: 'c1', label: 'Windrow Pad Alpha', type: 'Windrow Pad', capacity: 10 }
  ]);
  const [mrfUnits, setMrfUnits] = useState([
    { id: 'm1', label: 'MRF Shed 2', type: 'Manual Sorting Shed', capacity: 5 }
  ]);
  
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [activePolicyModal, setActivePolicyModal] = useState(null);

  const resultsRef = useRef(null);
  const parsedPerCapita = Number(perCapitaOption);
  const calculatedTpdDisplay = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);

  const addCompostUnit = () => {
    setCompostUnits([...compostUnits, { id: `c${Date.now()}`, label: lang === 'hi' ? 'नई कम्पोस्ट यूनिट' : 'New Compost Unit', type: 'Windrow Pad', capacity: 5 }]);
  };

  const removeCompostUnit = (id) => {
    if (compostUnits.length > 1) setCompostUnits(compostUnits.filter(u => u.id !== id));
  };

  const updateCompostUnit = (id, field, value) => {
    setCompostUnits(compostUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

  const addMrfUnit = () => {
    setMrfUnits([...mrfUnits, { id: `m${Date.now()}`, label: lang === 'hi' ? 'नया एमआरएफ शेड' : 'New MRF Shed', type: 'Manual Sorting Shed', capacity: 5 }]);
  };

  const removeMrfUnit = (id) => {
    if (mrfUnits.length > 1) setMrfUnits(mrfUnits.filter(u => u.id !== id));
  };

  const updateMrfUnit = (id, field, value) => {
    setMrfUnits(mrfUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

  const getSessionKey = () => `crf_paid_INTEGRATED_${name.trim().toLowerCase().replace(/\s+/g, '_')}_${selectedMonths.length}M`;

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
    const baseRate = 500;
    const baseTotal = billableMonths * baseRate;
    const finalTotalWithCharges = Math.round(baseTotal / (1 - 0.0236));
    return { count, freeMonths, billableMonths, baseTotal, total: finalTotalWithCharges };
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

      const gateHeaders = ["Date", "Day", `Total Gate Intake (${u})`, `Segregated Wet (${u})`, `Segregated Dry (${u})`, `Domestic Hazardous (${u})`, `Domestic Sanitary (${u})`, `Unsegregated Mixed (${u})`];
      const gateRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.wetSeg), formatVal(r.drySeg), formatVal(r.hazSeg), formatVal(r.sanSeg), formatVal(r.unsegregatedMixed)]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_Gate`);

      const preHeaders = ["Date", "Day", `Mixed Intake (${u})`, `Fine Screen Fraction (${u})`, `Coarse Screen Fraction (${u})`, `Heavy Inerts (${u})`];
      const preRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.unsegregatedMixed), formatVal(r.organicFines), formatVal(r.dryOversize), formatVal(r.heavyInerts)]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([preHeaders, ...preRows]), `${monthName}_PreSort`);

      compostUnits.forEach(unit => {
        const cHeaders = ["Date", "Day", `Unit Feed (${u})`, `Compost Yield (${u})`];
        const cRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.compostUnitBreakdown[unit.id]?.feed), formatVal(r.compostUnitBreakdown[unit.id]?.compostYield)]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([cHeaders, ...cRows]), `${monthName}_${unit.label.substring(0, 10)}`);
      });

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
        
        <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                <ShieldCheck size={12} /> {lang === 'hi' ? 'MULTI-ASSET SWM ESTIMATION ENGINE' : 'MULTI-ASSET SWM ESTIMATION ENGINE'}
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {lang === 'hi' ? 'Integrated 3-in-1 Multi-Unit Logbook Suite' : 'Integrated 3-in-1 Multi-Unit Logbook Suite'}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'Select State' : 'Select State'}</label>
              <select style={inputStyle} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'ULB / Facility Name' : 'ULB / Facility Name'}</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'Mobile Number' : 'Mobile Number'}</label>
              <input style={inputStyle} type="tel" maxLength={10} placeholder="" required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
            </div>
          </div>

          <div style={{ background: '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #e2e8f0', marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '10px' }}>Gate Intake & Segregation Efficiency</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '15px', alignItems: 'center' }}>
              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>Population (Approx.)</label>
                <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
              </div>
              <div>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#0f172a' }}>Source Segregation Rate: {segregationRate}%</label>
                <input type="range" min="20" max="95" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '8px' }} />
              </div>
            </div>
          </div>

          <div style={{ background: '#f0fdf4', padding: '12px', borderRadius: '6px', border: '1px solid #bbf7d0', marginBottom: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <strong style={{ fontSize: '13px', color: '#166534' }}>Composting Assets (Wet Line)</strong>
              <button type="button" onClick={addCompostUnit} style={{ padding: '6px 10px', background: '#047857', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={14} /> Add Compost Unit
              </button>
            </div>
            {compostUnits.map((u) => (
              <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                <input type="text" value={u.label} onChange={(e) => updateCompostUnit(u.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} />
                <select value={u.type} onChange={(e) => updateCompostUnit(u.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                  <option value="Windrow Pad">Windrow Pad</option>
                  <option value="Vermicompost Pit">Vermicompost Pit</option>
                </select>
                <input type="number" value={u.capacity} onChange={(e) => updateCompostUnit(u.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} />
                {compostUnits.length > 1 && <button type="button" onClick={() => removeCompostUnit(u.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={16} /></button>}
              </div>
            ))}
          </div>

          <div style={{ background: '#f0f9ff', padding: '12px', borderRadius: '6px', border: '1px solid #bae6fd', marginBottom: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <strong style={{ fontSize: '13px', color: '#0369a1' }}>MRF / Sorting Shed Assets (Dry Line)</strong>
              <button type="button" onClick={addMrfUnit} style={{ padding: '6px 10px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={14} /> Add MRF Shed
              </button>
            </div>
            {mrfUnits.map((u) => (
              <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                <input type="text" value={u.label} onChange={(e) => updateMrfUnit(u.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} />
                <select value={u.type} onChange={(e) => updateMrfUnit(u.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                  <option value="Manual Sorting Shed">Manual Sorting Shed</option>
                  <option value="Semi-Automated Line">Semi-Automated Line</option>
                </select>
                <input type="number" value={u.capacity} onChange={(e) => updateMrfUnit(u.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} />
                {mrfUnits.length > 1 && <button type="button" onClick={() => removeMrfUnit(u.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={16} /></button>}
              </div>
            ))}
          </div>

          <button type="submit" style={{ width: '100%', padding: '14px', background: '#047857', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            Generate Master Dataset (₹500/mo Suite) →
          </button>
        </form>

        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong>{name} — Gate Dataset Preview</strong>
              {isPaid && (
                <button onClick={downloadExcel} style={{ padding: '6px 12px', background: '#047857', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                  <Download size={13} /> Export Master Excel
                </button>
              )}
            </div>

            <div onContextMenu={(e) => !isPaid && e.preventDefault()} style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', userSelect: isPaid ? 'text' : 'none' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '600px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1', textAlign: 'left' }}>
                    <th>Date</th><th>Day</th><th>Gate Intake</th><th>Seg. Wet</th><th>Seg. Dry</th><th>Mixed Waste</th>
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
                    <li>Integrated 3-in-1 Master Suite: ₹500 / Month (Includes Gate, Compost & MRF Tabs)</li>
                  </ul>
                </div>
              )}
              <button type="button" onClick={() => setActivePolicyModal(null)} style={{ marginTop: '15px', padding: '8px 16px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Close</button>
            </div>
          </div>
        )}

      </div>
    </div>

      </div>
    </div>
  );
}

  );
}
