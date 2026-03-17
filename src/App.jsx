import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  ScatterChart, Scatter, ZAxis, PieChart, Pie, Cell, ComposedChart, Area
} from 'recharts';
import { 
  LayoutDashboard, Filter, Layers, DollarSign, Users, 
  Target, CheckCircle, Smartphone, AlertTriangle, Activity, 
  FileText, Clock, BarChart2, TrendingUp, TrendingDown, ChevronDown, Calendar, Award, Zap, XCircle, Check, Lock, LogOut, Loader2, Briefcase
} from 'lucide-react';

// --- Configuration ---
const MAIN_DATA_SHEET_ID = '1f1cUsWsRcS-I7VdVEVj1oLyalTtJLlCVnzUiWh77ff0';
const AUTH_DATA_SHEET_ID = '144YySNLbFulSD3bRVeRCxe5PoyrLPl5-vvuVLS8uVds';

const getCsvUrl = (id) => `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv`;

// --- Utility Functions ---
const parseCSV = (csv) => {
  const lines = csv.split('\n').filter(line => line.trim() !== '');
  if (lines.length === 0) return [];
  
  const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  const data = lines.slice(1).map(line => {
    const values = [];
    let inQuotes = false;
    let value = '';
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') inQuotes = !inQuotes;
      else if (char === ',' && !inQuotes) { values.push(value.trim().replace(/^"|"$/g, '')); value = ''; }
      else value += char;
    }
    values.push(value.trim().replace(/^"|"$/g, ''));
    const row = {};
    headers.forEach((header, index) => { row[header] = values[index] || ''; });
    return row;
  });
  return data;
};

const processData = (rawData) => {
  const processed = [];
  const groups = {};

  rawData.forEach(row => {
    const year = row.year;
    const month = row.month;
    const projectType = row.project_type === 'Incident Check' ? 'Incident Check' : 'Client Project';
    const projectNo = row.project_no;
    
    const quota = parseFloat(row.quota) || 0;
    const apComplete = parseFloat(row.ap_complete) || 0;
    const meowComplete = parseFloat(row.meow_complete) || 0;
    const fwComplete = parseFloat(row.fw_complete) || 0; 
    const badSample = parseFloat(row.bad_sample) || 0;
    const answers = parseFloat(row.Answers) || 0;
    const apCost = parseFloat(row.ap_total_price_ap) || 0;
    const totalThb = parseFloat(row.total_thb) || 0;
    const ir = parseFloat(row.ir) || 0;
    const loi = parseFloat(row.loi) || 0;
    const workingDay = parseFloat(row.working_day) || 0;
    
    const apMobile = parseFloat(row.ap_mobile) || 0;
    const ap3party = parseFloat(row.ap_3party) || 0;
    const apRabbit = parseFloat(row.ap_rabbit) || 0;
    const apPc = parseFloat(row.ap_pc) || 0;
    const cint = parseFloat(row.cint) || 0;

    const key = `${year}-${month}-${projectNo}`;

    if (projectType === 'Incident Check') {
      processed.push({
        ...row,
        project_type_mapped: projectType,
        quota, apComplete, meowComplete, fwComplete, badSample, answers, apCost, totalThb, ir, loi, mbokakr: parseFloat(row.MBOKAKR) || 0, workingDay,
        totalComplete: apComplete + meowComplete,
        percentComplete: quota > 0 ? ((apComplete + meowComplete) / quota) * 100 : 0,
        apMobile, ap3party, apRabbit, apPc, cint
      });
    } else {
      if (!groups[key]) {
        groups[key] = {
          ...row,
          project_type_mapped: projectType,
          quota: 0, apComplete: 0, meowComplete: 0, fwComplete: 0, badSample: 0, answers: 0,
          apCost: 0, totalThb: 0,
          sumIR: 0, countIR: 0,
          sumLOI: 0, countLOI: 0,
          sumMBOKAKR: 0, countMBOKAKR: 0,
          sumWorkingDay: 0, countWorkingDay: 0,
          projectNames: [],
          apMobile: 0, ap3party: 0, apRabbit: 0, apPc: 0, cint: 0
        };
      }
      
      const g = groups[key];
      g.quota += quota;
      g.apComplete += apComplete;
      g.meowComplete += meowComplete;
      g.fwComplete += fwComplete;
      g.badSample += badSample;
      g.answers += answers;
      g.apCost += apCost;
      g.totalThb += totalThb;
      
      if (ir > 0) { g.sumIR += ir; g.countIR++; }
      if (loi > 0) { g.sumLOI += loi; g.countLOI++; }
      if (parseFloat(row.MBOKAKR) > 0) { g.sumMBOKAKR += parseFloat(row.MBOKAKR); g.countMBOKAKR++; }
      if (workingDay > 0) { g.sumWorkingDay += workingDay; g.countWorkingDay++; }

      g.apMobile = Math.max(g.apMobile, apMobile);
      g.ap3party = Math.max(g.ap3party, ap3party);
      g.apRabbit = Math.max(g.apRabbit, apRabbit);
      g.apPc = Math.max(g.apPc, apPc);
      g.cint = Math.max(g.cint, cint);

      if (!g.projectNames.includes(row.project_name)) g.projectNames.push(row.project_name);
    }
  });

  Object.values(groups).forEach(g => {
    const totalComplete = g.apComplete + g.meowComplete;
    processed.push({
      ...g,
      project_name: g.projectNames.join(', '),
      ir: g.countIR ? g.sumIR / g.countIR : 0,
      loi: g.countLOI ? g.sumLOI / g.countLOI : 0,
      MBOKAKR: g.countMBOKAKR ? g.sumMBOKAKR / g.countMBOKAKR : 0,
      workingDay: g.countWorkingDay ? g.sumWorkingDay / g.countWorkingDay : 0,
      totalComplete,
      percentComplete: g.quota > 0 ? (totalComplete / g.quota) * 100 : 0,
      per_cpi_usd: totalComplete > 0 ? g.apCost / totalComplete : 0,
      per_cpi_thb: totalComplete > 0 ? g.totalThb / totalComplete : 0
    });
  });

  return processed;
};

const formatCurrency = (val, currency = 'USD') => new Intl.NumberFormat('en-US', { style: 'currency', currency }).format(val || 0);
const formatNumber = (val) => new Intl.NumberFormat('en-US').format(val || 0);

// --- Components ---
const DropdownFilter = ({ label, options, selected, onChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef(null);
  useEffect(() => {
    const handleClickOutside = (e) => { if (containerRef.current && !containerRef.current.contains(e.target)) setIsOpen(false); };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);
  const toggleOption = (option) => {
    const newSelected = selected.includes(option) ? selected.filter((item) => item !== option) : [...selected, option];
    onChange(newSelected);
  };
  const displayLabel = selected.length === 0 ? `All ${label}s` : selected.length === 1 ? String(selected[0]) : `${selected.length} Selected`;
  return (
    <div className="flex flex-col min-w-[170px]" ref={containerRef}>
      <label className="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1.5 ml-1">{label}</label>
      <div className="relative">
        <button type="button" onClick={() => setIsOpen(!isOpen)} className={`w-full flex items-center justify-between bg-white border border-gray-200 text-gray-700 py-2.5 px-4 rounded-lg leading-tight focus:outline-none transition-all shadow-sm text-sm hover:border-indigo-300 ${isOpen ? 'ring-2 ring-indigo-100 border-indigo-500' : ''}`}>
          <span className="truncate mr-2 font-medium">{displayLabel}</span>
          <ChevronDown className={`w-4 h-4 transition-transform duration-200 ${isOpen ? 'rotate-180 text-indigo-500' : 'text-gray-400'}`} />
        </button>
        {isOpen && (
          <div className="absolute z-50 mt-1 w-full bg-white border border-gray-100 rounded-xl shadow-xl max-h-64 overflow-y-auto overflow-x-hidden animate-in fade-in zoom-in duration-150">
            <div className="p-1">
              <button onClick={() => onChange([])} className={`w-full text-left px-3 py-2 text-sm rounded-lg transition-colors flex items-center justify-between ${selected.length === 0 ? 'bg-indigo-50 text-indigo-700 font-bold' : 'text-gray-600 hover:bg-gray-50'}`}><span>All {label}s</span>{selected.length === 0 && <Check className="w-4 h-4" />}</button>
              <div className="h-px bg-gray-100 my-1 mx-2" />
              {options.map((opt) => (
                <button key={opt} onClick={() => toggleOption(opt)} className={`w-full text-left px-3 py-2 text-sm rounded-lg transition-colors flex items-center justify-between group ${selected.includes(opt) ? 'bg-indigo-50 text-indigo-700 font-bold' : 'text-gray-600 hover:bg-gray-50'}`}>
                  <span className="truncate">{String(opt)}</span>
                  <div className={`w-4 h-4 border rounded flex items-center justify-center transition-all ${selected.includes(opt) ? 'bg-indigo-600 border-indigo-600' : 'border-gray-300 group-hover:border-indigo-400'}`}>{selected.includes(opt) && <Check className="w-3 h-3 text-white" />}</div>
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

const SummaryCard = ({ title, value, subValue, icon: Icon, color = "blue", comparison, comparisonYear }) => {
  const colorClasses = { blue: "bg-blue-50 text-blue-600", green: "bg-green-50 text-green-600", purple: "bg-purple-50 text-purple-600", orange: "bg-orange-50 text-orange-600", red: "bg-red-50 text-red-600", indigo: "bg-indigo-50 text-indigo-600" };
  const hasComparison = typeof comparison === 'number' && !isNaN(comparison);
  
  return (
    <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-5 flex items-start justify-between hover:shadow-md transition-shadow duration-200 h-full">
      <div>
        <p className="text-sm font-medium text-gray-500 mb-1 line-clamp-1" title={title}>{title}</p>
        <h3 className="text-xl font-bold text-gray-800 tracking-tight">{value}</h3>
        {subValue && <p className="text-xs text-gray-400 mt-1">{subValue}</p>}
        {hasComparison && (
          <div className={`text-[10px] mt-2 flex items-center font-bold ${comparison >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
            {comparison >= 0 ? <TrendingUp className="w-3 h-3 mr-1" /> : <TrendingDown className="w-3 h-3 mr-1" />}
            {Math.abs(comparison).toFixed(1)}% vs {comparisonYear || 'Prev Year'}
          </div>
        )}
      </div>
      <div className={`p-2 rounded-lg ${colorClasses[color]}`}><Icon className="w-5 h-5" /></div>
    </div>
  );
};

const WordCloud = ({ data }) => {
  const validData = Array.isArray(data) ? data : [];
  const maxVal = Math.max(...validData.map(d => d.value || 0), 1);
  const colors = ['#6366f1', '#ec4899', '#8b5cf6', '#10b981', '#f59e0b', '#3b82f6'];
  return (
    <div className="flex flex-wrap gap-2 justify-center items-center h-full p-4 overflow-y-auto">
      {validData.map((item, idx) => {
        const size = 0.8 + ((item.value || 0) / maxVal) * 1.5;
        return (
          <span key={idx} style={{ fontSize: `${size}rem`, color: colors[idx % colors.length], opacity: 0.8 + ((item.value || 0) / maxVal) * 0.2 }} className="font-bold px-2 py-1 bg-gray-50 rounded-md hover:bg-gray-100 transition-colors cursor-default" title={`${item.name}: ${item.value} Projects`}>
            {String(item.name)}
          </span>
        )
      })}
    </div>
  )
}

// --- Main App ---
const App = () => {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [isLoadingAuth, setIsLoadingAuth] = useState(true);
  const [isDataLoading, setIsDataLoading] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [loginForm, setLoginForm] = useState({ user: '', pass: '' });
  const [loginError, setLoginError] = useState('');
  const [authData, setAuthData] = useState([]);
  
  const [data, setData] = useState([]);
  const [filters, setFilters] = useState({ projectType: ['Client Project'], month: [], year: [], quater: [] });
  const [displayCount, setDisplayCount] = useState(10);

  // Helper to calculate basic metrics
  const calculateAllMetrics = (dataset) => {
    if (!dataset || dataset.length === 0) return null;
    const clientProjects = dataset.filter(d => d.project_type_mapped === 'Client Project');
    const totalProjects = clientProjects.length;
    const apCost = dataset.reduce((acc, curr) => acc + (curr.apCost || 0), 0);
    const thbCost = dataset.reduce((acc, curr) => acc + (curr.totalThb || 0), 0);
    const targetQuota = dataset.reduce((acc, curr) => acc + (curr.quota || 0), 0);
    const apComplete = dataset.reduce((acc, curr) => acc + (curr.apComplete || 0), 0);
    const meowComplete = dataset.reduce((acc, curr) => acc + (curr.meowComplete || 0), 0);
    const fwComplete = dataset.reduce((acc, curr) => acc + (curr.fwComplete || 0), 0);
    const allComplete = apComplete + meowComplete + fwComplete;
    const badSample = dataset.reduce((acc, curr) => acc + (curr.badSample || 0), 0);
    const allAnswers = dataset.reduce((acc, curr) => acc + (curr.answers || 0), 0);
    const validIr = dataset.filter(d => (d.ir || 0) > 0);
    const avgIr = validIr.length ? validIr.reduce((a, b) => a + (b.ir || 0), 0) / validIr.length : 0;
    const validLoi = dataset.filter(d => (d.loi || 0) > 0);
    const avgLoi = validLoi.length ? validLoi.reduce((a, b) => a + (b.loi || 0), 0) / validLoi.length : 0;
    const validCpiUsd = dataset.filter(d => d.per_cpi_usd > 0);
    const avgCpiUsd = validCpiUsd.length ? validCpiUsd.reduce((a, b) => a + (b.per_cpi_usd || 0), 0) / validCpiUsd.length : 0;
    const validCpiThb = dataset.filter(d => d.per_cpi_thb > 0);
    const avgCpiThb = validCpiThb.length ? validCpiThb.reduce((a, b) => a + (b.per_cpi_thb || 0), 0) / validCpiThb.length : 0;
    const kpiPassCount = dataset.filter(d => d.kpi_39 === 'Pass').length;
    const kpiRate = dataset.length ? (kpiPassCount / dataset.length) * 100 : 0;

    return { 
      totalProjects, thbCost, apCost, targetQuota, allComplete, apComplete, meowComplete, fwComplete, 
      badSample, allAnswers, avgIr, avgLoi, avgCpiUsd, avgCpiThb, kpiRate 
    };
  };

  // Initial Fetch: account_permission
  useEffect(() => {
    const fetchAuth = async () => {
      try {
        const response = await fetch(getCsvUrl(AUTH_DATA_SHEET_ID));
        const csvText = await response.text();
        const parsed = parseCSV(csvText);
        setAuthData(parsed);
      } catch (err) {
        console.error("Auth Fetch Error:", err);
      } finally {
        setIsLoadingAuth(false);
      }
    };
    fetchAuth();
  }, []);

  // Main Data Refresh
  const refreshMainData = async () => {
    setIsDataLoading(true);
    try {
      const response = await fetch(getCsvUrl(MAIN_DATA_SHEET_ID));
      const csvText = await response.text();
      const parsed = parseCSV(csvText);
      const processed = processData(parsed);
      setData(processed);
      
      const years = [...new Set(processed.map(d => d.year.toString()))].sort((a,b) => b-a);
      if (years.length > 0) {
        setFilters(prev => ({ ...prev, year: [years[0]] }));
      }
    } catch (err) {
      console.error("Data Fetch Error:", err);
    } finally {
      setIsDataLoading(false);
    }
  };

  useEffect(() => {
    if (isLoggedIn) {
      refreshMainData();
    }
  }, [isLoggedIn]);

  const handleLogin = (e) => {
    e.preventDefault();
    const foundUser = authData.find(u => 
      u.Username === loginForm.user && u.Password === loginForm.pass
    );
    if (foundUser) {
      setIsLoggedIn(true);
      setCurrentUser(foundUser);
      setLoginError('');
    } else {
      setLoginError('Username หรือ Password ไม่ถูกต้อง');
    }
  };

  const filteredData = useMemo(() => {
    return data.filter(d => 
      (filters.projectType.length === 0 || filters.projectType.includes(d.project_type_mapped)) &&
      (filters.month.length === 0 || filters.month.includes(d.month)) &&
      (filters.year.length === 0 || filters.year.includes(String(d.year))) &&
      (filters.quater.length === 0 || filters.quater.includes(d.quater))
    );
  }, [data, filters]);

  const stats = useMemo(() => {
    const current = calculateAllMetrics(filteredData) || { totalProjects:0, thbCost:0, apCost:0, targetQuota:0, allComplete:0, apComplete:0, meowComplete:0, fwComplete:0, badSample:0, allAnswers:0, avgIr:0, avgLoi:0, avgCpiUsd:0, avgCpiThb:0, kpiRate:0 };
    
    // Insights
    const mobileProjects = filteredData.filter(d => d.apMobile > 0).length;
    const mobileRate = filteredData.length ? (mobileProjects / filteredData.length) * 100 : 0;
    const thirdPartyProjects = filteredData.filter(d => (d.ap3party > 0 || d.apRabbit > 0 || d.apPc > 0 || d.cint > 0)).length;
    const thirdPartyRate = filteredData.length ? (thirdPartyProjects / filteredData.length) * 100 : 0;
    const validWorkingDay = filteredData.filter(d => d.workingDay > 0);
    const avgWorkingDay = validWorkingDay.length ? validWorkingDay.reduce((a, b) => a + (b.workingDay || 0), 0) / validWorkingDay.length : 0;
    
    const buckets = { 
      less70: filteredData.filter(d => d.percentComplete < 70).length, 
      bet70_90: filteredData.filter(d => d.percentComplete >= 70 && d.percentComplete <= 90).length, 
      more90: filteredData.filter(d => d.percentComplete > 90 && d.percentComplete < 100).length, 
      more100: filteredData.filter(d => d.percentComplete >= 100).length 
    };
    
    const mbokakrPassCount = filteredData.filter(d => (d.quota || 0) > 0 && (((d.apComplete || 0) + (d.meowComplete || 0)) / d.quota) >= 1).length;
    const avgPercentComplete = filteredData.length ? filteredData.reduce((a, b) => a + (b.percentComplete || 0), 0) / filteredData.length : 0;
    
    // Best Category & Project
    const catStats = {};
    filteredData.forEach(d => { if (d.category) { if (!catStats[d.category]) catStats[d.category] = { sum: 0, count: 0 }; catStats[d.category].sum += (d.percentComplete || 0); catStats[d.category].count++; } });
    let bestCategory = { name: 'N/A', val: 0 };
    Object.keys(catStats).forEach(k => { const avg = catStats[k].sum / catStats[k].count; if (avg > bestCategory.val) bestCategory = { name: k, val: avg }; });

    let bestProject = { name: 'N/A', val: 0 };
    if (filteredData.length > 0) {
        const sortedProjects = [...filteredData].sort((a, b) => (b.percentComplete || 0) - (a.percentComplete || 0));
        if (sortedProjects[0]) {
            bestProject = { name: sortedProjects[0].project_name, val: sortedProjects[0].percentComplete || 0 };
        }
    }

    // Comparison Logic
    let comparison = null;
    let prevYearLabel = null;
    if (filters.year.length === 1) {
      prevYearLabel = (parseInt(filters.year[0]) - 1).toString();
      const prevYearData = data.filter(d => d.year.toString() === prevYearLabel && (filters.projectType.length === 0 || filters.projectType.includes(d.project_type_mapped)));
      const prev = calculateAllMetrics(prevYearData);
      if (prev) {
        const getPct = (c, p) => p !== 0 ? ((c - p) / Math.abs(p)) * 100 : 0;
        comparison = {
          projects: getPct(current.totalProjects, prev.totalProjects),
          quota: getPct(current.targetQuota, prev.targetQuota),
          completes: getPct(current.allComplete, prev.allComplete),
          apCompletes: getPct(current.apComplete, prev.apComplete),
          meowCompletes: getPct(current.meowComplete, prev.meowComplete),
          fwCompletes: getPct(current.fwComplete, prev.fwComplete),
          badSample: getPct(current.badSample, prev.badSample),
          answers: getPct(current.allAnswers, prev.allAnswers),
          ir: getPct(current.avgIr, prev.avgIr),
          loi: getPct(current.avgLoi, prev.avgLoi),
          apCost: getPct(current.apCost, prev.apCost),
          thbCost: getPct(current.thbCost, prev.thbCost),
          avgCpiUsd: getPct(current.avgCpiUsd, prev.avgCpiUsd),
          avgCpiThb: getPct(current.avgCpiThb, prev.avgCpiThb),
          kpiRate: getPct(current.kpiRate, prev.kpiRate)
        };
      }
    }

    return { ...current, mobileRate, thirdPartyRate, avgWorkingDay, buckets, mbokakrPassCount, avgPercentComplete, bestCategory, bestProject, comparison, prevYearLabel };
  }, [filteredData, data, filters.year, filters.projectType]);

  // Scatter Chart: X = IR%, Y = LOI
  const scatterData = filteredData.filter(d => (d.ir || 0) > 0 && (d.loi || 0) > 0).map(d => ({ x: d.ir, y: d.loi, z: d.totalComplete, name: d.project_name }));
  const categoryData = useMemo(() => {
    const cats = {};
    filteredData.forEach(d => { if(d.category) { if(!cats[d.category]) cats[d.category] = 0; cats[d.category]++; } });
    return Object.keys(cats).map(k => ({ name: k, value: cats[k] }));
  }, [filteredData]);

  const kpiTrendData = useMemo(() => {
    const monthOrder = { 'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'June': 6, 'July': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12 };
    const monthlyStats = {};
    filteredData.forEach(d => { if (d.month) { if (!monthlyStats[d.month]) monthlyStats[d.month] = { total: 0, pass: 0 }; monthlyStats[d.month].total++; if (d.kpi_39 === 'Pass') monthlyStats[d.month].pass++; } });
    return Object.keys(monthlyStats).sort((a, b) => monthOrder[a] - monthOrder[b]).map(m => ({ name: m, rate: (monthlyStats[m].pass / monthlyStats[m].total) * 100 }));
  }, [filteredData]);

  const mbokakrStatusData = useMemo(() => [
      { name: 'Achieved Target', value: stats.mbokakrPassCount, color: '#10B981' },
      { name: 'Missed Target', value: Math.max(0, filteredData.length - stats.mbokakrPassCount), color: '#F59E0B' }
  ], [stats.mbokakrPassCount, filteredData.length]);

  const completionCardData = useMemo(() => [
      { name: 'Target Met', desc: '≥ 100% Complete', value: stats.buckets.more100, color: 'bg-green-50 text-green-700', border: 'border-green-200', icon: CheckCircle, iconColor: 'text-green-600' },
      { name: 'High', desc: '90-99% Complete', value: stats.buckets.more90, color: 'bg-teal-50 text-teal-700', border: 'border-teal-200', icon: TrendingUp, iconColor: 'text-teal-600' },
      { name: 'Medium', desc: '70-90% Complete', value: stats.buckets.bet70_90, color: 'bg-yellow-50 text-yellow-700', border: 'border-yellow-200', icon: AlertTriangle, iconColor: 'text-yellow-600' },
      { name: 'Low', desc: '< 70% Complete', value: stats.buckets.less70, color: 'bg-red-50 text-red-700', border: 'border-red-200', icon: XCircle, iconColor: 'text-red-600' },
  ], [stats.buckets]);

  const monthlyTrendData = useMemo(() => {
    const monthOrder = { 'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'June': 6, 'July': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12 };
    const monthlyStats = {};
    filteredData.forEach(d => {
        if (d.month) {
            if (!monthlyStats[d.month]) monthlyStats[d.month] = { client: 0, incident: 0, totalCompletes: 0 };
            if (d.project_type_mapped === 'Client Project') monthlyStats[d.month].client++;
            else monthlyStats[d.month].incident++;
            monthlyStats[d.month].totalCompletes += (d.totalComplete || 0);
        }
    });
    return Object.keys(monthlyStats).sort((a, b) => monthOrder[a] - monthOrder[b]).map(m => ({ name: m, Client: monthlyStats[m].client, Incident: monthlyStats[m].incident, Completes: monthlyStats[m].totalCompletes }));
  }, [filteredData]);

  const monthlyInsight = useMemo(() => {
      if (monthlyTrendData.length === 0) return { busiestMonth: '-', avgProjects: 0 };
      const busiest = monthlyTrendData.reduce((p, c) => (p.Client + p.Incident) > (c.Client + c.Incident) ? p : c);
      const totalP = monthlyTrendData.reduce((a, c) => a + c.Client + c.Incident, 0);
      return { busiestMonth: busiest.name, avgProjects: (totalP / monthlyTrendData.length).toFixed(1) };
  }, [monthlyTrendData]);

  const filterOptions = {
    projectType: ['Client Project', 'Incident Check'],
    month: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    year: ['2023', '2024', '2025', '2026'],
    quater: ["4'23", "1'24", "2'24", "3'24", "4'24", "1'25", "2'25", "3'25", "4'25", "1'26"]
  };

  if (isLoadingAuth) {
    return (
      <div className="min-h-screen bg-indigo-950 flex flex-col items-center justify-center p-6 text-white gap-4">
        <Loader2 className="w-12 h-12 animate-spin text-indigo-400" />
        <p className="font-medium animate-pulse text-indigo-200 tracking-widest text-xs uppercase">Connecting to account_permission...</p>
      </div>
    );
  }

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen bg-indigo-950 flex items-center justify-center p-6 font-sans">
        <div className="max-w-md w-full bg-white rounded-3xl shadow-2xl overflow-hidden p-10 animate-in fade-in slide-in-from-bottom-8 duration-700 border border-white/20">
          <div className="text-center mb-10">
            <div className="bg-indigo-100 w-16 h-16 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-inner">
              <Lock className="w-8 h-8 text-indigo-600" />
            </div>
            <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Panel Login</h1>
            <p className="text-gray-500 text-sm mt-2 font-medium">กรุณาเข้าสู่ระบบด้วยบัญชี Google Sheet</p>
          </div>
          <form onSubmit={handleLogin} className="space-y-6">
            <div className="space-y-1">
              <label className="block text-xs font-bold text-gray-400 uppercase ml-1 tracking-wider">Username</label>
              <input type="text" required className="w-full px-4 py-3.5 rounded-xl border border-gray-100 bg-gray-50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all font-medium" placeholder="ชื่อผู้ใช้" value={loginForm.user} onChange={e => setLoginForm({...loginForm, user: e.target.value})} />
            </div>
            <div className="space-y-1">
              <label className="block text-xs font-bold text-gray-400 uppercase ml-1 tracking-wider">Password</label>
              <input type="password" required className="w-full px-4 py-3.5 rounded-xl border border-gray-100 bg-gray-50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all font-medium" placeholder="รหัสผ่าน" value={loginForm.pass} onChange={e => setLoginForm({...loginForm, pass: e.target.value})} />
            </div>
            {loginError && <div className="flex items-center gap-2 text-red-500 text-xs font-bold justify-center bg-red-50 py-3 rounded-xl border border-red-100"><AlertTriangle className="w-3.5 h-3.5" />{loginError}</div>}
            <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-4 rounded-xl transition-all shadow-lg shadow-indigo-200 active:scale-[0.98] flex items-center justify-center gap-2">Sign In <Check className="w-4 h-4" /></button>
          </form>
          <div className="mt-8 pt-6 border-t border-gray-50 text-center"><p className="text-[10px] text-gray-400 font-bold uppercase tracking-[0.2em] leading-relaxed">Security Verified: account_permission</p></div>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-gray-50 min-h-screen font-sans text-gray-800 p-8">
      {/* Header */}
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-3xl font-extrabold text-indigo-900 tracking-tight">Panel Report Dashboard 2025</h1>
          <div className="flex items-center gap-2 mt-1">
            <span className="text-gray-500">Interactive Insights & Performance Overview</span>
            <span className="text-xs bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider">Role: {String(currentUser?.Role || 'User')}</span>
            <span className="text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider ml-1">Auth: {String(currentUser?.Username || 'N/A')}</span>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <button onClick={refreshMainData} disabled={isDataLoading} className="flex items-center gap-2 bg-white px-4 py-2 rounded-lg border border-gray-200 shadow-sm text-sm font-semibold text-gray-600 hover:bg-gray-50 transition-all disabled:opacity-50">
            <Zap className={`w-4 h-4 text-orange-400 ${isDataLoading ? 'animate-pulse' : ''}`} /> {isDataLoading ? 'Syncing...' : 'Refresh Data'}
          </button>
          <button onClick={() => { setIsLoggedIn(false); setData([]); }} className="flex items-center gap-2 text-gray-500 hover:text-red-600 transition-colors font-bold text-sm bg-white px-4 py-2 rounded-lg border border-gray-200 shadow-sm"><LogOut className="w-4 h-4" /> Sign Out</button>
          <div className="bg-white px-4 py-2 rounded-lg shadow-sm text-sm text-gray-500 flex items-center gap-2 border border-gray-100"><Clock className="w-4 h-4" /> Live Sync: Enabled</div>
        </div>
      </div>

      {/* Filter Section */}
      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-6">
        <div className="flex items-center gap-2 mb-4 text-indigo-700 font-semibold text-lg">
          <Filter className="w-5 h-5" /> Filter Data
        </div>
        <div className="flex flex-wrap gap-6">
          <DropdownFilter label="Project Type" options={filterOptions.projectType} selected={filters.projectType} onChange={(val) => setFilters({...filters, projectType: val})} />
          <DropdownFilter label="Year" options={filterOptions.year} selected={filters.year} onChange={(val) => setFilters({...filters, year: val})} />
          <DropdownFilter label="Quarter" options={filterOptions.quater} selected={filters.quater} onChange={(val) => setFilters({...filters, quater: val})} />
          <DropdownFilter label="Month" options={filterOptions.month} selected={filters.month} onChange={(val) => setFilters({...filters, month: val})} />
        </div>
      </div>

      {/* Executive Summary */}
      <div className="mb-8">
          <h2 className="text-xl font-bold text-gray-700 mb-4 flex items-center gap-2"><LayoutDashboard className="w-5 h-5"/> Executive Summary</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-4 mb-4">
            <SummaryCard title="Total Projects" value={stats.totalProjects} icon={Layers} color="indigo" comparison={stats.comparison?.projects} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Target Quota" value={formatNumber(stats.targetQuota)} icon={Target} color="purple" comparison={stats.comparison?.quota} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="All Participants" value={formatNumber(stats.allAnswers)} subValue="(RD)" icon={Users} color="purple" comparison={stats.comparison?.answers} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. IR" value={`${(stats.avgIr).toFixed(2)}%`} icon={Activity} color="orange" comparison={stats.comparison?.ir} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. LOI" value={`${stats.avgLoi.toFixed(1)} min`} icon={Clock} color="orange" comparison={stats.comparison?.loi} comparisonYear={stats.prevYearLabel} />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-4 mb-4">
            <SummaryCard title="All Completes" value={formatNumber(stats.allComplete)} subValue="AP + Meow + FW" icon={CheckCircle} color="indigo" comparison={stats.comparison?.completes} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="AP Completes" value={formatNumber(stats.apComplete)} icon={Users} color="indigo" comparison={stats.comparison?.apCompletes} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Meow Completes" value={formatNumber(stats.meowComplete)} icon={Users} color="indigo" comparison={stats.comparison?.meowCompletes} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="FW (CLT) Completes" value={formatNumber(stats.fwComplete)} icon={Users} color="blue" comparison={stats.comparison?.fwCompletes} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Total Bad Sample" value={formatNumber(stats.badSample)} icon={AlertTriangle} color="red" comparison={stats.comparison?.badSample} comparisonYear={stats.prevYearLabel} />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-4">
            <SummaryCard title="Total Cost (USD)" value={formatCurrency(stats.apCost)} icon={DollarSign} color="green" comparison={stats.comparison?.apCost} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Total Cost (THB)" value={formatNumber(stats.thbCost)} subValue="THB" icon={DollarSign} color="green" comparison={stats.comparison?.thbCost} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. CPI (USD)" value={formatCurrency(stats.avgCpiUsd)} icon={TrendingUp} color="blue" comparison={stats.comparison?.avgCpiUsd} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. CPI (THB)" value={formatNumber(stats.avgCpiThb.toFixed(2))} subValue="THB" icon={TrendingUp} color="blue" comparison={stats.comparison?.avgCpiThb} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="KPI Pass Rate" value={`${stats.kpiRate.toFixed(1)}%`} icon={CheckCircle} color={stats.kpiRate > 90 ? "green" : "red"} comparison={stats.comparison?.kpiRate} comparisonYear={stats.prevYearLabel} />
          </div>
      </div>

      {/* Insights Row */}
      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-8">
          <h2 className="text-xl font-bold text-gray-700 mb-6 flex items-center gap-2"><Smartphone className="w-5 h-5"/> Mobile & Panel Insights</h2>
          <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
            <div className="flex flex-col gap-4 lg:col-span-1">
               <div className="bg-blue-50 p-6 rounded-lg flex items-center justify-between shadow-sm">
                   <div><p className="text-sm text-gray-500 uppercase font-bold mb-1">Mobile Panel Usage</p><p className="text-3xl font-bold text-blue-600">{stats.mobileRate.toFixed(1)}%</p></div>
                   <Smartphone className="w-10 h-10 text-blue-200" />
               </div>
               <div className="bg-purple-50 p-6 rounded-lg flex items-center justify-between shadow-sm">
                   <div><p className="text-sm text-gray-500 uppercase font-bold mb-1">3rd Party Usage</p><p className="text-3xl font-bold text-purple-600">{stats.thirdPartyRate.toFixed(1)}%</p></div>
                   <Users className="w-10 h-10 text-purple-200" />
               </div>
               <div className="bg-orange-50 p-6 rounded-lg flex items-center justify-between shadow-sm">
                   <div><p className="text-sm text-gray-500 uppercase font-bold mb-1">Avg. Working Days</p><p className="text-3xl font-bold text-orange-600">{stats.avgWorkingDay.toFixed(1)}</p><p className="text-xs text-orange-400 font-medium">Days per project</p></div>
                   <Calendar className="w-10 h-10 text-orange-200" />
               </div>
            </div>
            <div className="lg:col-span-3 bg-gray-50 rounded-lg p-6 border border-gray-200 flex flex-col md:flex-row gap-8">
                <div className="flex-1">
                    <h3 className="text-sm font-semibold text-gray-600 mb-6 flex items-center gap-2"><Target className="w-4 h-4"/> Completion Rate Breakdown</h3>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        {completionCardData.map((item, idx) => (
                            <div key={idx} className={`p-4 rounded-xl border ${item.border} ${item.color} shadow-sm flex flex-col justify-between h-32 hover:shadow-md transition-shadow`}>
                                <div className="flex justify-between items-start">
                                    <div><p className="font-bold text-lg">{String(item.name)}</p><p className="text-xs opacity-80">{String(item.desc)}</p></div>
                                    <div className="p-2 bg-white rounded-full bg-opacity-50"><item.icon className={`w-5 h-5 ${item.iconColor}`} /></div>
                                </div>
                                <div className="flex items-baseline gap-2"><span className="text-4xl font-bold">{item.value}</span><span className="text-sm opacity-80">Projects ({filteredData.length > 0 ? ((item.value / filteredData.length) * 100).toFixed(0) : 0}%)</span></div>
                            </div>
                        ))}
                    </div>
                </div>
                <div className="md:w-1/3 flex flex-col justify-center border-t md:border-t-0 md:border-l border-gray-200 pt-6 md:pt-0 md:pl-8 space-y-6">
                    <div>
                        <p className="text-sm text-gray-500 font-medium mb-1 flex items-center gap-1"><Zap className="w-3 h-3"/> Overall Avg. Completion</p>
                        <div className="flex items-baseline gap-2"><span className="text-4xl font-extrabold text-indigo-600">{stats.avgPercentComplete.toFixed(1)}%</span><span className="text-xs text-gray-400">across all projects</span></div>
                    </div>
                    <div className="space-y-4">
                        <div>
                            <p className="text-sm text-gray-500 font-medium mb-2 flex items-center gap-1"><Award className="w-3 h-3"/> Best Performing Category</p>
                            <div className="bg-white p-3 rounded-lg border border-gray-100 shadow-sm">
                                <div className="flex justify-between items-center mb-1"><span className="text-sm font-bold text-gray-700 truncate mr-2" title={String(stats.bestCategory.name)}>{String(stats.bestCategory.name)}</span><span className="bg-green-100 text-green-700 text-[10px] px-2 py-0.5 rounded-full font-bold">{stats.bestCategory.val.toFixed(0)}% Avg</span></div>
                                <div className="w-full bg-gray-100 rounded-full h-1.5 mt-2"><div className="bg-green-500 h-1.5 rounded-full" style={{ width: `${Math.min(stats.bestCategory.val, 100)}%` }}></div></div>
                            </div>
                        </div>
                        <div>
                            <p className="text-sm text-gray-500 font-medium mb-2 flex items-center gap-1"><Briefcase className="w-3 h-3"/> Best Performing Project</p>
                            <div className="bg-white p-3 rounded-lg border border-gray-100 shadow-sm">
                                <div className="flex justify-between items-center mb-1"><span className="text-sm font-bold text-indigo-700 truncate mr-2" title={String(stats.bestProject.name)}>{String(stats.bestProject.name)}</span><span className="bg-indigo-100 text-indigo-700 text-[10px] px-2 py-0.5 rounded-full font-bold">{stats.bestProject.val.toFixed(0)}%</span></div>
                                <div className="w-full bg-indigo-50 rounded-full h-1.5 mt-2"><div className="bg-indigo-500 h-1.5 rounded-full" style={{ width: `${Math.min(stats.bestProject.val, 100)}%` }}></div></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
          </div>
      </div>

      {/* Charts Row */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
             <div className="flex justify-between items-center mb-6"><h2 className="text-xl font-bold text-gray-700 flex items-center gap-2"><BarChart2 className="w-5 h-5"/> Performance Analysis</h2></div>
             <div className="h-80">
                <h3 className="text-sm font-semibold text-gray-500 mb-2 text-center">IR% vs LOI Analysis (Size = Completes)</h3>
                <ResponsiveContainer width="100%" height="100%">
                    <ScatterChart margin={{ top: 20, right: 20, bottom: 20, left: 20 }}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis type="number" dataKey="x" name="IR" unit="%" />
                        <YAxis type="number" dataKey="y" name="LOI" unit=" min" />
                        <ZAxis type="number" dataKey="z" range={[50, 400]} name="Completes" />
                        <RechartsTooltip cursor={{ strokeDasharray: '3 3' }} />
                        <Legend />
                        <Scatter name="Projects" data={scatterData} fill="#6366f1" />
                    </ScatterChart>
                </ResponsiveContainer>
             </div>
          </div>
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
             <div className="flex justify-between items-center mb-6"><h2 className="text-xl font-bold text-gray-700 flex items-center gap-2"><Layers className="w-5 h-5"/> Top Categories</h2></div>
             <div className="h-80 border-2 border-dashed border-gray-100 rounded-lg flex items-center justify-center"><WordCloud data={categoryData} /></div>
          </div>
      </div>

      {/* Monthly and Targets */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
           <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                <h2 className="text-lg font-bold text-gray-700 mb-4 flex items-center gap-2"><TrendingUp className="w-5 h-5"/> KPI Pass Rate Trend (Monthly)</h2>
                <div className="h-80">
                    <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={kpiTrendData} margin={{ top: 5, right: 30, left: 20, bottom: 5 }}><CartesianGrid strokeDasharray="3 3" vertical={false} /><XAxis dataKey="name" /><YAxis domain={[0, 100]} unit="%" /><RechartsTooltip /><Legend /><Line type="monotone" dataKey="rate" name="Pass Rate %" stroke="#10B981" strokeWidth={3} activeDot={{ r: 8 }} /></LineChart>
                    </ResponsiveContainer>
                </div>
           </div>
           <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                <h2 className="text-lg font-bold text-gray-700 mb-2 flex items-center gap-2"><Award className="w-5 h-5"/> MBOKAKR Target Achievement</h2>
                <p className="text-xs text-gray-500 mb-6">Target: (AP + Meow Completes) ≥ 100% of Quota</p>
                <div className="flex flex-col md:flex-row items-center justify-around h-72">
                    <div className="h-64 w-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                                <Pie data={mbokakrStatusData} cx="50%" cy="50%" innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                                    {mbokakrStatusData.map((e, i) => <Cell key={`c-${i}`} fill={e.color} />)}
                                </Pie><RechartsTooltip /><Legend verticalAlign="bottom" height={36}/>
                            </PieChart>
                        </ResponsiveContainer>
                    </div>
                    <div className="flex flex-col gap-4 p-4">
                        <div className="bg-green-50 border border-green-100 p-4 rounded-xl text-center min-w-[140px]"><span className="block text-3xl font-bold text-green-600">{stats.mbokakrPassCount}</span><span className="text-xs font-semibold text-green-700 uppercase">Projects Passed</span></div>
                        <div className="bg-yellow-50 border border-yellow-100 p-4 rounded-xl text-center min-w-[140px]"><span className="block text-3xl font-bold text-yellow-600">{Math.max(0, filteredData.length - stats.mbokakrPassCount)}</span><span className="text-xs font-semibold text-yellow-700 uppercase">Projects Missed</span></div>
                    </div>
                </div>
           </div>
      </div>

      {/* Monthly Volume */}
      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-8">
          <div className="flex items-center justify-between mb-6">
              <h2 className="text-xl font-bold text-gray-700 flex items-center gap-2"><BarChart2 className="w-5 h-5"/> Monthly Volume & Trends</h2>
              <div className="flex gap-4">
                  <div className="text-right"><p className="text-xs text-gray-500 font-bold uppercase">Busiest Month</p><p className="text-lg font-bold text-indigo-600">{String(monthlyInsight.busiestMonth)}</p></div>
                  <div className="text-right"><p className="text-xs text-gray-500 font-bold uppercase">Avg. Projects/Month</p><p className="text-lg font-bold text-purple-600">{monthlyInsight.avgProjects}</p></div>
              </div>
          </div>
          <div className="h-96">
              <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={monthlyTrendData} margin={{ top: 20, right: 20, bottom: 20, left: 20 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} /><XAxis dataKey="name" /><YAxis yAxisId="left" orientation="left" stroke="#8884d8" /><YAxis yAxisId="right" orientation="right" stroke="#82ca9d" /><RechartsTooltip /><Legend /><Bar yAxisId="left" dataKey="Client" name="Client Projects" stackId="a" fill="#8884d8" /><Bar yAxisId="left" dataKey="Incident" name="Incident Checks" stackId="a" fill="#ffc658" /><Line yAxisId="right" type="monotone" dataKey="Completes" name="Total Completes" stroke="#82ca9d" strokeWidth={3} />
                  </ComposedChart>
              </ResponsiveContainer>
          </div>
      </div>

      {/* Data Table */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="p-6 border-b border-gray-100 flex justify-between items-center">
             <h2 className="text-xl font-bold text-gray-700 flex items-center gap-2"><FileText className="w-5 h-5"/> Project Details</h2>
             <span className="text-sm text-gray-500">Showing {Math.min(displayCount, filteredData.length)} of {filteredData.length} Projects</span>
          </div>
          <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                  <thead>
                      <tr className="bg-gray-50 text-gray-500 text-xs uppercase tracking-wider">
                          <th className="p-4 font-semibold">Month</th><th className="p-4 font-semibold">Project No</th><th className="p-4 font-semibold w-1/4">Project Name</th><th className="p-4 font-semibold">Type</th><th className="p-4 font-semibold text-right">Quota</th><th className="p-4 font-semibold text-right">Completes</th><th className="p-4 font-semibold text-right">%</th><th className="p-4 font-semibold text-right">Total Cost (THB)</th><th className="p-4 font-semibold text-center">Status</th>
                      </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                      {filteredData.slice(0, displayCount).map((row, idx) => (
                          <tr key={idx} className="hover:bg-gray-50 transition-colors">
                              <td className="p-4 text-sm text-gray-600">{String(row.month)}</td>
                              <td className="p-4 text-sm font-medium text-indigo-600">{String(row.project_no)}</td>
                              <td className="p-4 text-sm text-gray-800 truncate max-w-xs" title={String(row.project_name)}>{String(row.project_name)}</td>
                              <td className="p-4 text-sm text-gray-500"><span className={`px-2 py-1 rounded-full text-xs ${row.project_type_mapped === 'Incident Check' ? 'bg-orange-100 text-orange-600' : 'bg-blue-100 text-blue-600'}`}>{row.project_type_mapped === 'Incident Check' ? 'Incident' : 'Client'}</span></td>
                              <td className="p-4 text-sm text-gray-600 text-right">{formatNumber(row.quota)}</td>
                              <td className="p-4 text-sm text-gray-600 text-right font-medium">{formatNumber(row.totalComplete)}</td>
                              <td className="p-4 text-sm text-right"><span className={`font-bold ${(row.percentComplete || 0) >= 100 ? 'text-green-600' : (row.percentComplete || 0) < 70 ? 'text-red-500' : 'text-yellow-600'}`}>{(row.percentComplete || 0).toFixed(1)}%</span></td>
                              <td className="p-4 text-sm text-gray-600 text-right">{formatNumber(row.totalThb)}</td>
                              <td className="p-4 text-center">{row.kpi_39 === 'Pass' ? <CheckCircle className="w-5 h-5 text-green-500 mx-auto" /> : <AlertTriangle className="w-5 h-5 text-red-400 mx-auto" />}</td>
                          </tr>
                      ))}
                  </tbody>
              </table>
          </div>
          {displayCount < filteredData.length && (
              <div className="p-4 text-center border-t border-gray-100">
                  <button onClick={() => setDisplayCount(prev => prev + 10)} className="px-6 py-2 bg-white border border-gray-300 rounded-lg text-sm font-medium text-gray-700 hover:bg-gray-50 transition-colors shadow-sm">Load More Projects <ChevronDown className="w-4 h-4 inline ml-1"/></button>
              </div>
          )}
      </div>
    </div>
  );
};

export default App;