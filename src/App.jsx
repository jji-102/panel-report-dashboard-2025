import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  ScatterChart, Scatter, ZAxis, PieChart, Pie, Cell, ComposedChart, Area
} from 'recharts';
import { 
  LayoutDashboard, Filter, Layers, DollarSign, Users, 
  Target, CheckCircle, Smartphone, AlertTriangle, Activity, 
  FileText, Clock, BarChart2, TrendingUp, TrendingDown, ChevronDown, Calendar, Award, Zap, XCircle, Check, Lock, LogOut, Loader2, Briefcase, X, Info, Tag, Search, Hash, Type, Percent, Globe, ExternalLink, List
} from 'lucide-react';

// --- Configuration ---
const MAIN_DATA_SHEET_ID = '1f1cUsWsRcS-I7VdVEVj1oLyalTtJLlCVnzUiWh77ff0';
const AUTH_DATA_SHEET_ID = '144YySNLbFulSD3bRVeRCxe5PoyrLPl5-vvuVLS8uVds';
const DASHBOARD_VERSION = "09-0326OP-DA"; // Update Version: 09-0326OP-DA (Interactive Completion Breakdown)
const RATE_CARD_URL = "https://ratecard-gold-theta.vercel.app/";

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
        totalComplete: apComplete + meowComplete + fwComplete,
        percentComplete: quota > 0 ? ((apComplete + meowComplete + fwComplete) / quota) * 100 : 0,
        per_cpi_usd: apComplete > 0 ? apCost / apComplete : 0, 
        per_cpi_thb: apComplete > 0 ? totalThb / apComplete : 0,
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
    const totalComplete = g.apComplete + g.meowComplete + g.fwComplete;
    processed.push({
      ...g,
      project_name: g.projectNames.join(', '),
      ir: g.countIR ? g.sumIR / g.countIR : 0,
      loi: g.countLOI ? g.sumLOI / g.countLOI : 0,
      MBOKAKR: g.countMBOKAKR ? g.sumMBOKAKR / g.countMBOKAKR : 0,
      workingDay: g.countWorkingDay ? g.sumWorkingDay / g.countWorkingDay : 0,
      totalComplete,
      percentComplete: g.quota > 0 ? (totalComplete / g.quota) * 100 : 0,
      per_cpi_usd: g.apComplete > 0 ? g.apCost / g.apComplete : 0,
      per_cpi_thb: g.apComplete > 0 ? g.totalThb / g.apComplete : 0 
    });
  });

  return processed;
};

const formatCurrency = (val, currency = 'USD') => new Intl.NumberFormat('en-US', { style: 'currency', currency }).format(val || 0);
const formatNumber = (val) => new Intl.NumberFormat('en-US').format(val || 0);

// --- Status Helpers ---
const getIrStatus = (val) => {
  const v = Math.round(val);
  if (v >= 80) return { label: 'High', color: 'text-emerald-600', bg: 'bg-emerald-50' };
  if (v >= 30) return { label: 'Mid', color: 'text-amber-600', bg: 'bg-amber-50' };
  if (v > 5) return { label: 'Low', color: 'text-orange-600', bg: 'bg-orange-50' };
  return { label: 'Alert', color: 'text-rose-600', bg: 'bg-rose-50' };
};

const getLoiStatus = (val) => {
  if (Math.round(val) > 15) return { label: 'Overtime', color: 'text-rose-600', bg: 'bg-rose-50' };
  return null;
};

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

const WordCloud = ({ data, onCategoryDoubleClick }) => {
  const validData = Array.isArray(data) ? data : [];
  const maxVal = Math.max(...validData.map(d => d.value || 0), 1);
  const colors = ['#6366f1', '#ec4899', '#8b5cf6', '#10b981', '#f59e0b', '#3b82f6'];
  return (
    <div className="flex flex-wrap gap-2 justify-center items-center h-full p-4 overflow-y-auto">
      {validData.map((item, idx) => {
        const size = 0.8 + ((item.value || 0) / maxVal) * 1.5;
        return (
          <span 
            key={idx} 
            style={{ fontSize: `${size}rem`, color: colors[idx % colors.length], opacity: 0.8 + ((item.value || 0) / maxVal) * 0.2 }} 
            className="font-bold px-3 py-1 bg-gray-50 rounded-lg hover:bg-indigo-50 hover:text-indigo-600 transition-all cursor-pointer select-none border border-transparent hover:border-indigo-100" 
            title={`Double-click to view all ${item.name} projects`}
            onDoubleClick={() => onCategoryDoubleClick(item.name)}
          >
            {String(item.name)}
          </span>
        )
      })}
    </div>
  )
}

const ProjectModal = ({ project, onClose }) => {
  if (!project) return null;

  const totalComp = (parseFloat(project.apComplete) || 0) + (parseFloat(project.meowComplete) || 0) + (parseFloat(project.fwComplete) || 0);
  const quotaNum = parseFloat(project.quota) || 0;
  const quotaPct = quotaNum > 0 ? ((totalComp / quotaNum) * 100).toFixed(1) : "0.0";
  
  const getChannelPct = (val) => totalComp > 0 ? ((parseFloat(val) || 0) / totalComp * 100).toFixed(1) : "0.0";

  const dataGroups = [
    {
      title: "General Information",
      icon: Info,
      items: [
        { label: "Project No.", value: project.project_no, icon: Hash },
        { label: "Project Name", value: project.project_name, icon: FileText },
        { label: "Project Type", value: project.project_type_mapped, icon: Type },
        { label: "Period", value: `${project.month} ${project.year}`, icon: Calendar },
        { label: "Quarter", value: project.quater, icon: Layers },
        { label: "Category", value: project.category || 'N/A', icon: Award },
      ]
    },
    {
      title: "Performance Metrics",
      icon: Activity,
      items: [
        { label: "Target Quota", value: formatNumber(project.quota), icon: Target },
        { label: "Total Completes", value: `${formatNumber(totalComp)} (${quotaPct}% of Quota)`, icon: CheckCircle, highlight: true },
        { label: "AP Completes", value: `${formatNumber(project.apComplete)} (${getChannelPct(project.apComplete)}%)`, icon: Users },
        { label: "Meow Completes", value: `${formatNumber(project.meowComplete)} (${getChannelPct(project.meowComplete)}%)`, icon: Users },
        { label: "FW Completes", value: `${formatNumber(project.fwComplete)} (${getChannelPct(project.fwComplete)}%)`, icon: Users },
        { label: "Bad Samples", value: formatNumber(project.badSample), icon: AlertTriangle },
        { label: "Incidence Rate (IR)", value: `${Math.round(project.ir || 0)}%`, icon: Percent },
        { label: "Length of Interview", value: `${Math.round(project.loi || 0)} min`, icon: Clock },
        { label: "Working Days", value: `${Math.round(project.workingDay || 0)} d`, icon: Calendar },
      ]
    },
    {
      title: "Financial Analysis (AP Based)",
      icon: DollarSign,
      items: [
        { label: "Total Budget (USD)", value: formatCurrency(project.apCost, 'USD'), icon: DollarSign },
        { label: "Total Budget (THB)", value: formatNumber(project.totalThb) + " THB", icon: DollarSign },
        { label: "Unit Cost (CPI USD)", value: formatCurrency(project.per_cpi_usd, 'USD'), icon: TrendingUp },
        { label: "Unit Cost (CPI THB)", value: (project.per_cpi_thb || 0).toFixed(2) + " THB", icon: TrendingUp },
        { label: "KPI Status", value: project.kpi_39, status: true, icon: Zap },
      ]
    },
    {
      title: "Panel Distribution",
      icon: Smartphone,
      items: [
        { label: "Mobile Usage", value: project.apMobile > 0 ? "Enabled" : "Disabled", icon: Smartphone },
        { label: "3rd Party Access", value: (project.ap3party > 0 || project.apRabbit > 0 || project.apPc > 0 || project.cint > 0) ? "Active" : "None", icon: Globe },
        { label: "Cint Integration", value: project.cint > 0 ? "Integrated" : "No", icon: Zap },
      ]
    }
  ];

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-indigo-950/60 backdrop-blur-md p-4 animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-4xl max-h-[90vh] rounded-3xl shadow-2xl overflow-hidden flex flex-col animate-in zoom-in-95 duration-300">
        <div className="bg-indigo-600 p-6 flex justify-between items-center text-white shadow-lg">
          <div className="flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-xl"><Briefcase className="w-6 h-6" /></div>
            <div>
              <h2 className="text-xl font-bold leading-tight">Project Analysis: {project.project_no}</h2>
              <p className="text-indigo-100 text-sm font-medium opacity-90">{project.project_name}</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-white/20 rounded-full transition-colors"><X className="w-6 h-6" /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-8 bg-gray-50/50">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {dataGroups.map((group, idx) => (
              <div key={idx} className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100">
                <div className="flex items-center gap-2 mb-4 pb-2 border-b border-gray-50">
                  <group.icon className="w-5 h-5 text-indigo-500" />
                  <h3 className="font-bold text-gray-700 text-sm uppercase tracking-wider">{group.title}</h3>
                </div>
                <div className="space-y-4">
                  {group.items.map((item, i) => (
                    <div key={i} className="flex items-start justify-between text-sm">
                      <div className="flex items-center gap-2 text-gray-500">
                        <item.icon className="w-3.5 h-3.5 opacity-60" />
                        <span className="font-medium">{item.label}</span>
                      </div>
                      <span className={`font-bold text-right ml-4 ${item.highlight ? 'text-indigo-600' : item.status ? (item.value === 'Pass' ? 'text-green-600' : 'text-red-500') : 'text-gray-800'}`}>
                        {item.value}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
        <div className="p-6 bg-white border-t border-gray-100 flex justify-end">
          <button onClick={onClose} className="px-8 py-2.5 bg-indigo-600 text-white font-bold rounded-xl hover:bg-indigo-700 transition-all shadow-lg active:scale-95">Close</button>
        </div>
      </div>
    </div>
  );
};

// --- Generic Project List Modal ---
const ProjectListModal = ({ title, subtitle, projects, onProjectSelect, onClose }) => {
  if (!title) return null;
  return (
    <div className="fixed inset-0 z-[105] flex items-center justify-center bg-indigo-950/40 backdrop-blur-sm p-4 animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-5xl h-[80vh] rounded-3xl shadow-2xl overflow-hidden flex flex-col animate-in slide-in-from-bottom-8 duration-300">
        <div className="bg-indigo-800 p-6 flex justify-between items-center text-white">
          <div className="flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-xl"><Layers className="w-6 h-6" /></div>
            <div>
              <h2 className="text-xl font-bold leading-tight">{title}</h2>
              <p className="text-indigo-200 text-sm font-medium">{subtitle || `Found ${projects.length} projects.`}</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-white/20 rounded-full transition-colors"><X className="w-6 h-6" /></button>
        </div>
        <div className="p-4 bg-amber-50 border-b border-amber-100 text-amber-700 text-xs font-bold uppercase tracking-widest text-center">
          Double-click any project name to view deep analysis
        </div>
        <div className="flex-1 overflow-y-auto p-4 bg-white">
          {projects.length === 0 ? (
            <div className="flex flex-col items-center justify-center h-full text-gray-400">
              <Activity className="w-12 h-12 mb-2 opacity-20" />
              <p className="font-bold">No projects found in this selection.</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {projects.map((p, idx) => (
                <div 
                  key={idx} 
                  onDoubleClick={() => onProjectSelect(p)}
                  className="p-5 bg-white border border-gray-100 rounded-2xl shadow-sm hover:border-indigo-400 hover:shadow-md transition-all cursor-pointer group"
                >
                  <div className="flex justify-between items-start mb-2">
                    <span className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2 py-1 rounded-md uppercase">{p.project_no}</span>
                    <span className={`text-[10px] font-bold px-2 py-1 rounded-md uppercase ${p.kpi_39 === 'Pass' ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600'}`}>
                      {p.kpi_39}
                    </span>
                  </div>
                  <h4 className="text-sm font-bold text-gray-800 line-clamp-2 mb-3 group-hover:text-indigo-700 transition-colors">{p.project_name}</h4>
                  <div className="flex justify-between items-center text-[10px] text-gray-400 font-bold uppercase border-t pt-3">
                    <div className="flex items-center gap-1"><Users className="w-3 h-3" /> {formatNumber(p.totalComplete)} Complete</div>
                    <div className="flex items-center gap-1"><Target className="w-3 h-3" /> {Math.round(p.percentComplete)}% Goal</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="p-6 bg-gray-50 border-t border-gray-100 flex justify-end">
          <button onClick={onClose} className="px-6 py-2 bg-gray-200 text-gray-600 font-bold rounded-xl hover:bg-gray-300 transition-all">Close View</button>
        </div>
      </div>
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
  const [searchTerm, setSearchTerm] = useState('');
  const [displayCount, setDisplayCount] = useState(10);
  const [selectedProject, setSelectedProject] = useState(null);
  const [selectedCategoryName, setSelectedCategoryName] = useState(null);
  const [selectedBucketName, setSelectedBucketName] = useState(null);

  const calculateAllMetrics = (dataset) => {
    if (!dataset || dataset.length === 0) return null;
    const totalProjects = dataset.length; 
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
    
    const totalApComplete = dataset.reduce((acc, curr) => acc + (curr.apComplete || 0), 0);
    const avgCpiUsd = totalApComplete > 0 ? apCost / totalApComplete : 0;
    const avgCpiThb = totalApComplete > 0 ? thbCost / totalApComplete : 0;

    const kpiPassCount = dataset.filter(d => d.kpi_39 === 'Pass').length;
    const kpiRate = dataset.length ? (kpiPassCount / dataset.length) * 100 : 0;
    return { totalProjects, thbCost, apCost, targetQuota, allComplete, apComplete, meowComplete, fwComplete, badSample, allAnswers, avgIr, avgLoi, avgCpiUsd, avgCpiThb, kpiRate };
  };

  useEffect(() => {
    const fetchAuth = async () => {
      try {
        const response = await fetch(getCsvUrl(AUTH_DATA_SHEET_ID));
        const csvText = await response.text();
        const parsed = parseCSV(csvText);
        setAuthData(parsed);
      } catch (err) { console.error("Auth Fetch Error:", err); } 
      finally { setIsLoadingAuth(false); }
    };
    fetchAuth();
  }, []);

  const refreshMainData = async () => {
    setIsDataLoading(true);
    try {
      const response = await fetch(getCsvUrl(MAIN_DATA_SHEET_ID));
      const csvText = await response.text();
      const parsed = parseCSV(csvText);
      const processed = processData(parsed);
      setData(processed);
      const years = [...new Set(processed.map(d => d.year.toString()))].sort((a,b) => b-a);
      if (years.length > 0) { setFilters(prev => ({ ...prev, year: [years[0]] })); }
    } catch (err) { console.error("Data Fetch Error:", err); } 
    finally { setIsDataLoading(false); }
  };

  useEffect(() => { if (isLoggedIn) { refreshMainData(); } }, [isLoggedIn]);

  const handleLogin = (e) => {
    e.preventDefault();
    const foundUser = authData.find(u => u.Username === loginForm.user && u.Password === loginForm.pass);
    if (foundUser) { setIsLoggedIn(true); setCurrentUser(foundUser); setLoginError(''); } 
    else { setLoginError('Username หรือ Password ไม่ถูกต้อง'); }
  };

  const filteredData = useMemo(() => {
    return data.filter(d => {
      const matchesFilters = (
        (filters.projectType.length === 0 || filters.projectType.includes(d.project_type_mapped)) &&
        (filters.month.length === 0 || filters.month.includes(d.month)) &&
        (filters.year.length === 0 || filters.year.includes(String(d.year))) &&
        (filters.quater.length === 0 || filters.quater.includes(d.quater))
      );
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch = (searchTerm === '' || d.project_no?.toLowerCase().includes(searchLower) || d.project_name?.toLowerCase().includes(searchLower));
      return matchesFilters && matchesSearch;
    });
  }, [data, filters, searchTerm]);

  const stats = useMemo(() => {
    const current = calculateAllMetrics(filteredData) || { totalProjects:0, thbCost:0, apCost:0, targetQuota:0, allComplete:0, apComplete:0, meowComplete:0, fwComplete:0, badSample:0, allAnswers:0, avgIr:0, avgLoi:0, avgCpiUsd:0, avgCpiThb:0, kpiRate:0 };
    const mobileProjects = filteredData.filter(d => d.apMobile > 0).length;
    const mobileRate = filteredData.length ? (mobileProjects / filteredData.length) * 100 : 0;
    const thirdPartyProjects = filteredData.filter(d => (d.ap3party > 0 || d.apRabbit > 0 || d.apPc > 0 || d.cint > 0)).length;
    const thirdPartyRate = filteredData.length ? (thirdPartyProjects / filteredData.length) * 100 : 0;
    const validWorkingDay = filteredData.filter(d => d.workingDay > 0);
    const avgWorkingDay = validWorkingDay.length ? validWorkingDay.reduce((a, b) => a + (b.workingDay || 0), 0) / validWorkingDay.length : 0;
    
    // --- Specific Calculation for Completion Rate Breakdown (AP + Meow only) ---
    const apMeowData = filteredData.map(d => ({
      ...d,
      apMeowPct: d.quota > 0 ? (((d.apComplete || 0) + (d.meowComplete || 0)) / d.quota) * 100 : 0
    }));

    const buckets = { 
      less70: apMeowData.filter(d => d.apMeowPct < 70).length, 
      bet70_90: apMeowData.filter(d => d.apMeowPct >= 70 && d.apMeowPct <= 90).length, 
      more90: apMeowData.filter(d => d.apMeowPct > 90 && d.apMeowPct < 100).length, 
      more100: apMeowData.filter(d => d.apMeowPct >= 100).length 
    };

    const avgPercentComplete = apMeowData.length ? apMeowData.reduce((a, b) => a + b.apMeowPct, 0) / apMeowData.length : 0;

    const mbokakrPassCount = filteredData.filter(d => (d.quota || 0) > 0 && (((d.apComplete || 0) + (d.meowComplete || 0)) / d.quota) >= 1).length;
    
    const catStats = {};
    filteredData.forEach(d => { if (d.category) { if (!catStats[d.category]) catStats[d.category] = { sum: 0, count: 0 }; catStats[d.category].sum += (d.percentComplete || 0); catStats[d.category].count++; } });
    let bestCategory = { name: 'N/A', val: 0 };
    Object.keys(catStats).forEach(k => { const avg = catStats[k].sum / catStats[k].count; if (avg > bestCategory.val) bestCategory = { name: k, val: avg }; });
    let bestProject = { name: 'N/A', val: 0 };
    if (filteredData.length > 0) {
        const sortedProjects = [...filteredData].sort((a, b) => (b.percentComplete || 0) - (a.percentComplete || 0));
        if (sortedProjects[0]) { bestProject = { name: sortedProjects[0].project_name, val: sortedProjects[0].percentComplete || 0 }; }
    }
    let comparison = null; let prevYearLabel = null;
    if (filters.year.length === 1) {
      prevYearLabel = (parseInt(filters.year[0]) - 1).toString();
      const prevYearData = data.filter(d => d.year.toString() === prevYearLabel && (filters.projectType.length === 0 || filters.projectType.includes(d.project_type_mapped)));
      const prev = calculateAllMetrics(prevYearData);
      if (prev) {
        const getPct = (c, p) => p !== 0 ? ((c - p) / Math.abs(p)) * 100 : 0;
        comparison = { projects: getPct(current.totalProjects, prev.totalProjects), quota: getPct(current.targetQuota, prev.targetQuota), completes: getPct(current.allComplete, prev.allComplete), apCompletes: getPct(current.apComplete, prev.apComplete), meowCompletes: getPct(current.meowComplete, prev.meowComplete), fwCompletes: getPct(current.fwComplete, prev.fwComplete), badSample: getPct(current.badSample, prev.badSample), answers: getPct(current.allAnswers, prev.allAnswers), ir: getPct(current.avgIr, prev.avgIr), loi: getPct(current.avgLoi, prev.avgLoi), apCost: getPct(current.apCost, prev.apCost), thbCost: getPct(current.thbCost, prev.thbCost), avgCpiUsd: getPct(current.avgCpiUsd, prev.avgCpiUsd), avgCpiThb: getPct(current.avgCpiThb, prev.avgCpiThb), kpiRate: getPct(current.kpiRate, prev.kpiRate) };
      }
    }
    return { ...current, mobileRate, thirdPartyRate, avgWorkingDay, buckets, mbokakrPassCount, avgPercentComplete, bestCategory, bestProject, comparison, prevYearLabel };
  }, [filteredData, data, filters.year, filters.projectType]);

  const categoryProjects = useMemo(() => {
    if (!selectedCategoryName) return [];
    return filteredData.filter(d => d.category === selectedCategoryName);
  }, [filteredData, selectedCategoryName]);

  const bucketProjects = useMemo(() => {
    if (!selectedBucketName) return [];
    const apMeowData = filteredData.map(d => ({
      ...d,
      apMeowPct: d.quota > 0 ? (((d.apComplete || 0) + (d.meowComplete || 0)) / d.quota) * 100 : 0
    }));

    switch (selectedBucketName) {
      case 'Target Met': return apMeowData.filter(d => d.apMeowPct >= 100);
      case 'High': return apMeowData.filter(d => d.apMeowPct >= 90 && d.apMeowPct < 100);
      case 'Medium': return apMeowData.filter(d => d.apMeowPct >= 70 && d.apMeowPct <= 90);
      case 'Low': return apMeowData.filter(d => d.apMeowPct < 70);
      default: return [];
    }
  }, [filteredData, selectedBucketName]);

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

  const mbokakrStatusData = useMemo(() => [ { name: 'Achieved Target', value: stats.mbokakrPassCount, color: '#10B981' }, { name: 'Missed Target', value: Math.max(0, filteredData.length - stats.mbokakrPassCount), color: '#F59E0B' } ], [stats.mbokakrPassCount, filteredData.length]);
  
  const completionCardData = useMemo(() => [ 
    { name: 'Target Met', desc: '≥ 100% Complete', value: stats.buckets.more100, color: 'bg-green-50 text-green-700', border: 'border-green-200', icon: CheckCircle, iconColor: 'text-green-600' }, 
    { name: 'High', desc: '90-99% Complete', value: stats.buckets.more90, color: 'bg-teal-50 text-teal-700', border: 'border-teal-200', icon: TrendingUp, iconColor: 'text-teal-600' }, 
    { name: 'Medium', desc: '70-90% Complete', value: stats.buckets.bet70_90, color: 'bg-yellow-50 text-yellow-700', border: 'border-yellow-200', icon: AlertTriangle, iconColor: 'text-yellow-600' }, 
    { name: 'Low', desc: '< 70% Complete', value: stats.buckets.less70, color: 'bg-red-50 text-red-700', border: 'border-red-200', icon: XCircle, iconColor: 'text-red-600' }, 
  ], [stats.buckets]);

  const monthlyTrendData = useMemo(() => {
    const monthOrder = { 'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'June': 6, 'July': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12 };
    const monthlyStats = {};
    filteredData.forEach(d => { if (d.month) { if (!monthlyStats[d.month]) monthlyStats[d.month] = { client: 0, incident: 0, totalCompletes: 0 }; if (d.project_type_mapped === 'Client Project') monthlyStats[d.month].client++; else monthlyStats[d.month].incident++; monthlyStats[d.month].totalCompletes += (d.totalComplete || 0); } });
    return Object.keys(monthlyStats).sort((a, b) => monthOrder[a] - monthOrder[b]).map(m => ({ name: m, Client: monthlyStats[m].client, Incident: monthlyStats[m].incident, Completes: monthlyStats[m].totalCompletes }));
  }, [filteredData]);

  const monthlyInsight = useMemo(() => { if (monthlyTrendData.length === 0) return { busiestMonth: '-', avgProjects: 0 }; const busiest = monthlyTrendData.reduce((p, c) => (p.Client + p.Incident) > (c.Client + c.Incident) ? p : c); const totalP = monthlyTrendData.reduce((a, c) => a + c.Client + c.Incident, 0); return { busiestMonth: busiest.name, avgProjects: (totalP / monthlyTrendData.length).toFixed(1) }; }, [monthlyTrendData]);

  const filterOptions = { projectType: ['Client Project', 'Incident Check'], month: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], year: ['2023', '2024', '2025', '2026'], quater: ["4'23", "1'24", "2'24", "3'24", "4'24", "1'25", "2'25", "3'25", "4'25", "1'26"] };

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
            <div className="bg-indigo-100 w-16 h-16 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-inner"><Lock className="w-8 h-8 text-indigo-600" /></div>
            <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Panel Login</h1>
            <p className="text-gray-500 text-sm mt-2 font-medium">Please sign in with your credentials</p>
          </div>
          <form onSubmit={handleLogin} className="space-y-6">
            <div className="space-y-1">
              <label className="block text-xs font-bold text-gray-400 uppercase ml-1 tracking-wider">Username</label>
              <input type="text" required className="w-full px-4 py-3.5 rounded-xl border border-gray-100 bg-gray-50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all font-medium" value={loginForm.user} onChange={e => setLoginForm({...loginForm, user: e.target.value})} />
            </div>
            <div className="space-y-1">
              <label className="block text-xs font-bold text-gray-400 uppercase ml-1 tracking-wider">Password</label>
              <input type="password" required className="w-full px-4 py-3.5 rounded-xl border border-gray-100 bg-gray-50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all font-medium" value={loginForm.pass} onChange={e => setLoginForm({...loginForm, pass: e.target.value})} />
            </div>
            {loginError && <div className="flex items-center gap-2 text-red-500 text-xs font-bold justify-center bg-red-50 py-3 rounded-xl border border-red-100"><AlertTriangle className="w-3.5 h-3.5" />{loginError}</div>}
            <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-4 rounded-xl transition-all shadow-lg active:scale-[0.98] flex items-center justify-center gap-2">Sign In <Check className="w-4 h-4" /></button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-gray-50 min-h-screen font-sans text-gray-800 p-8">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-3xl font-extrabold text-indigo-900 tracking-tight">Panel Report Dashboard 2025</h1>
          <div className="flex flex-wrap items-center gap-2 mt-1">
            <span className="text-gray-500 text-sm">Real-time Performance Overview</span>
            <span className="text-xs bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider">Role: {String(currentUser?.Role || 'User')}</span>
            <span className="text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider ml-1">Auth: {String(currentUser?.Username || 'N/A')}</span>
            <span className="text-xs bg-indigo-600 text-white px-3 py-0.5 rounded-full font-bold uppercase tracking-wider ml-1 flex items-center gap-1 shadow-sm">
              <Tag className="w-3 h-3" /> Ver: {DASHBOARD_VERSION}
            </span>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <a 
            href={RATE_CARD_URL} 
            target="_blank" 
            rel="noopener noreferrer" 
            className="flex items-center gap-2 bg-amber-500 hover:bg-amber-600 text-white px-4 py-2 rounded-lg shadow-sm text-sm font-bold transition-all hover:shadow-md active:scale-95"
          >
            <ExternalLink className="w-4 h-4" /> View Rate Card
          </a>
          <button onClick={refreshMainData} disabled={isDataLoading} className="flex items-center gap-2 bg-white px-4 py-2 rounded-lg border border-gray-200 shadow-sm text-sm font-semibold text-gray-600 hover:bg-gray-50 transition-all disabled:opacity-50">
            <Zap className={`w-4 h-4 text-orange-400 ${isDataLoading ? 'animate-pulse' : ''}`} /> {isDataLoading ? 'Syncing...' : 'Refresh Data'}
          </button>
          <button onClick={() => { setIsLoggedIn(false); setData([]); }} className="flex items-center gap-2 text-gray-500 hover:text-red-600 transition-colors font-bold text-sm bg-white px-4 py-2 rounded-lg border border-gray-200 shadow-sm"><LogOut className="w-4 h-4" /> Sign Out</button>
        </div>
      </div>

      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-6">
        <div className="flex items-center gap-2 mb-4 text-indigo-700 font-semibold text-lg"><Filter className="w-5 h-5" /> Filter Data</div>
        <div className="flex flex-wrap gap-6">
          <DropdownFilter label="Project Type" options={filterOptions.projectType} selected={filters.projectType} onChange={(val) => setFilters({...filters, projectType: val})} />
          <DropdownFilter label="Year" options={filterOptions.year} selected={filters.year} onChange={(val) => setFilters({...filters, year: val})} />
          <DropdownFilter label="Quarter" options={filterOptions.quater} selected={filters.quater} onChange={(val) => setFilters({...filters, quater: val})} />
          <DropdownFilter label="Month" options={filterOptions.month} selected={filters.month} onChange={(val) => setFilters({...filters, month: val})} />
        </div>
      </div>

      <div className="mb-8">
          <h2 className="text-xl font-bold text-gray-700 mb-4 flex items-center gap-2"><LayoutDashboard className="w-5 h-5"/> Executive Summary</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-4 mb-4">
            <SummaryCard title="Total Projects" value={stats.totalProjects} icon={Layers} color="indigo" comparison={stats.comparison?.projects} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Target Quota" value={formatNumber(stats.targetQuota)} icon={Target} color="purple" comparison={stats.comparison?.quota} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="All Participants" value={formatNumber(stats.allAnswers)} subValue="(RD)" icon={Users} color="purple" comparison={stats.comparison?.answers} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. IR" value={`${Math.round(stats.avgIr)}%`} icon={Activity} color="orange" comparison={stats.comparison?.ir} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. LOI" value={`${Math.round(stats.avgLoi)} min`} icon={Clock} color="orange" comparison={stats.comparison?.loi} comparisonYear={stats.prevYearLabel} />
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
            <SummaryCard title="Avg. CPI (USD)" value={formatCurrency(stats.avgCpiUsd)} subValue="AP Basis" icon={TrendingUp} color="blue" comparison={stats.comparison?.avgCpiUsd} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="Avg. CPI (THB)" value={formatNumber(stats.avgCpiThb.toFixed(2))} subValue="THB (AP Basis)" icon={TrendingUp} color="blue" comparison={stats.comparison?.avgCpiThb} comparisonYear={stats.prevYearLabel} />
            <SummaryCard title="KPI Pass Rate" value={`${stats.kpiRate.toFixed(1)}%`} icon={CheckCircle} color={stats.kpiRate > 90 ? "green" : "red"} comparison={stats.comparison?.kpiRate} comparisonYear={stats.prevYearLabel} />
          </div>
      </div>

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
                   <div><p className="text-sm text-gray-500 uppercase font-bold mb-1">Avg. Working Days</p><p className="text-3xl font-bold text-orange-600">{Math.round(stats.avgWorkingDay)}</p><p className="text-xs text-orange-400 font-medium">Days per project</p></div>
                   <Calendar className="w-10 h-10 text-orange-200" />
               </div>
            </div>
            <div className="lg:col-span-3 bg-gray-50 rounded-lg p-6 border border-gray-200 flex flex-col md:flex-row gap-8">
                <div className="flex-1">
                    <h3 className="text-sm font-semibold text-gray-600 mb-4 flex items-center gap-2"><Target className="w-4 h-4"/> Completion Rate Breakdown</h3>
                    <p className="text-[10px] text-gray-400 font-bold uppercase mb-4 tracking-widest">Double-click a card to explore projects in that level</p>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        {completionCardData.map((item, idx) => (
                            <div 
                              key={idx} 
                              onDoubleClick={() => setSelectedBucketName(item.name)}
                              className={`p-4 rounded-xl border ${item.border} ${item.color} shadow-sm flex flex-col justify-between h-32 hover:shadow-md transition-all cursor-pointer select-none active:scale-[0.98] hover:border-opacity-100 border-opacity-40`}
                            >
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
                        <div className="flex items-baseline gap-2"><span className="text-4xl font-extrabold text-indigo-600">{stats.avgPercentComplete.toFixed(1)}%</span><span className="text-xs text-gray-400">across all projects (AP+Meow)</span></div>
                    </div>
                    <div className="space-y-4">
                        <div>
                            <p className="text-sm text-gray-500 font-medium mb-2 flex items-center gap-1"><Award className="w-3 h-3"/> Best Performing Category</p>
                            <div className="bg-white p-3 rounded-lg border border-gray-100 shadow-sm">
                                <div className="flex justify-between items-center mb-1"><span className="text-sm font-bold text-gray-700 truncate mr-2">{String(stats.bestCategory.name)}</span><span className="bg-green-100 text-green-700 text-[10px] px-2 py-0.5 rounded-full font-bold">{stats.bestCategory.val.toFixed(0)}% Avg</span></div>
                                <div className="w-full bg-gray-100 rounded-full h-1.5 mt-2"><div className="bg-green-500 h-1.5 rounded-full" style={{ width: `${Math.min(stats.bestCategory.val, 100)}%` }}></div></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
          </div>
      </div>

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
                        <Tooltip cursor={{ strokeDasharray: '3 3' }} />
                        <Legend />
                        <Scatter name="Projects" data={scatterData} fill="#6366f1" />
                    </ScatterChart>
                </ResponsiveContainer>
             </div>
          </div>
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
             <div className="flex justify-between items-center mb-6">
                <h2 className="text-xl font-bold text-gray-700 flex items-center gap-2"><Layers className="w-5 h-5"/> Top Categories</h2>
             </div>
             <div className="h-80 border-2 border-dashed border-gray-100 rounded-2xl bg-gray-50/30">
                <WordCloud data={categoryData} onCategoryDoubleClick={(name) => setSelectedCategoryName(name)} />
             </div>
             <p className="text-[10px] text-center text-gray-400 font-bold uppercase mt-4 tracking-widest">Double-click category name to explore projects</p>
          </div>
      </div>

      <div className="bg-white rounded-3xl shadow-xl border border-gray-100 overflow-hidden mb-12">
          <div className="p-8 border-b border-gray-100 bg-white flex flex-col md:flex-row md:justify-between md:items-center gap-4">
             <div>
                <h2 className="text-2xl font-bold text-gray-800 flex items-center gap-3"><FileText className="w-7 h-7 text-indigo-500"/> Project Details</h2>
             </div>
             <div className="flex flex-col sm:flex-row items-center gap-3">
                <div className="relative w-full sm:w-80">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                  <input type="text" placeholder="Search by ID or Name..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all" />
                </div>
                <span className="text-xs bg-indigo-50 text-indigo-600 px-4 py-2 rounded-full font-bold uppercase tracking-wider border border-indigo-100 whitespace-nowrap">Total: {filteredData.length} Projects</span>
             </div>
          </div>
          <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[1500px]">
                  <thead>
                      <tr className="bg-gray-50/80 text-gray-500 text-[11px] uppercase tracking-[0.1em]">
                          <th className="p-5 font-bold border-b border-gray-100">Period</th>
                          <th className="p-5 font-bold border-b border-gray-100">Project ID</th>
                          <th className="p-5 font-bold border-b border-gray-100 w-64">Project Name</th>
                          <th className="p-5 font-bold border-b border-gray-100">Category</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">Quota</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">Total Comp.</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">AP Comp.</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-center">IR% Status</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-center">LOI Status</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">Working Days</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">Total Cost</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">CPI (USD)</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-right">CPI (THB)</th>
                          <th className="p-5 font-bold border-b border-gray-100 text-center">Status</th>
                      </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 bg-white">
                      {filteredData.slice(0, displayCount).map((row, idx) => {
                          const allCompletes = (parseFloat(row.apComplete) || 0) + (parseFloat(row.meowComplete) || 0) + (parseFloat(row.fwComplete) || 0);
                          const irVal = Math.round(row.ir || 0);
                          const loiVal = Math.round(row.loi || 0);
                          const irStatus = getIrStatus(irVal);
                          const loiStatus = getLoiStatus(loiVal);
                          
                          return (
                            <tr key={idx} onDoubleClick={() => setSelectedProject(row)} className="hover:bg-indigo-50/30 transition-all cursor-pointer group">
                                <td className="p-5"><div className="flex flex-col"><span className="text-sm font-bold text-gray-800">{String(row.month)}</span><span className="text-[10px] text-gray-400 font-bold uppercase">{row.quater} / {row.year}</span></div></td>
                                <td className="p-5"><span className="text-sm font-bold text-indigo-600 bg-indigo-50 px-2 py-1 rounded-lg border border-indigo-100 group-hover:bg-indigo-600 group-hover:text-white transition-colors">{String(row.project_no)}</span></td>
                                <td className="p-5"><p className="text-sm text-gray-800 font-medium line-clamp-2" title={String(row.project_name)}>{String(row.project_name)}</p></td>
                                <td className="p-5"><span className="text-xs text-gray-500 font-semibold">{row.category || 'N/A'}</span></td>
                                <td className="p-5 text-sm text-gray-600 text-right font-medium">{formatNumber(row.quota)}</td>
                                <td className="p-5 text-sm text-gray-800 text-right font-extrabold">{formatNumber(allCompletes)}</td>
                                <td className="p-5 text-sm text-indigo-500 text-right font-bold">{formatNumber(row.apComplete)}</td>
                                <td className="p-5 text-center">
                                    <div className={`inline-flex flex-col px-3 py-1 rounded-full border ${irStatus.bg} ${irStatus.color} border-current min-w-[70px]`}>
                                        <span className="text-xs font-black">{irVal}%</span>
                                        <span className="text-[9px] font-bold uppercase leading-none opacity-80">{irStatus.label}</span>
                                    </div>
                                </td>
                                <td className="p-5 text-center">
                                    <div className={`inline-flex items-center gap-1 font-bold ${loiStatus ? loiStatus.color : 'text-gray-600'}`}>
                                        {loiVal} min
                                        {loiStatus && <AlertTriangle className="w-3 h-3 animate-pulse" />}
                                    </div>
                                </td>
                                <td className="p-5 text-sm text-gray-600 text-right font-medium">{Math.round(row.workingDay || 0)} d</td>
                                <td className="p-5 text-sm text-emerald-600 text-right font-bold">{formatCurrency(row.apCost, 'USD')}</td>
                                <td className="p-5 text-sm text-emerald-600 text-right font-bold">{(row.per_cpi_usd || 0).toFixed(2)}</td>
                                <td className="p-5 text-sm text-indigo-600 text-right font-bold">{(row.per_cpi_thb || 0).toFixed(2)}</td>
                                <td className="p-5 text-center">
                                  {row.kpi_39 === 'Pass' ? 
                                    <div className="flex items-center justify-center bg-emerald-50 w-8 h-8 rounded-full mx-auto"><CheckCircle className="w-5 h-5 text-emerald-500" /></div> : 
                                    <div className="flex items-center justify-center bg-rose-50 w-8 h-8 rounded-full mx-auto"><AlertTriangle className="w-5 h-5 text-rose-500" /></div>
                                  }
                                </td>
                            </tr>
                          );
                      })}
                  </tbody>
              </table>
          </div>
          <div className="p-8 text-center border-t border-gray-100 bg-gray-50/30">
              {displayCount < filteredData.length ? (
                <button onClick={() => setDisplayCount(prev => prev + 10)} className="px-10 py-3 bg-white border border-gray-200 rounded-2xl text-sm font-bold text-gray-700 hover:border-indigo-400 hover:text-indigo-600 transition-all shadow-sm hover:shadow-md flex items-center gap-2 mx-auto">Load More <ChevronDown className="w-4 h-4"/></button>
              ) : ( <p className="text-gray-400 text-sm font-bold uppercase tracking-widest">End of project list</p> )}
          </div>
      </div>

      {/* Popups */}
      {selectedCategoryName && (
        <ProjectListModal 
          title={`Category: ${selectedCategoryName}`} 
          subtitle={`Detailed view of all projects tagged under ${selectedCategoryName}`}
          projects={categoryProjects} 
          onProjectSelect={(p) => setSelectedProject(p)}
          onClose={() => setSelectedCategoryName(null)} 
        />
      )}
      {selectedBucketName && (
        <ProjectListModal 
          title={`Completion Level: ${selectedBucketName}`} 
          subtitle={`Analysis based on AP+Meow completion rate criteria.`}
          projects={bucketProjects} 
          onProjectSelect={(p) => setSelectedProject(p)}
          onClose={() => setSelectedBucketName(null)} 
        />
      )}
      {selectedProject && <ProjectModal project={selectedProject} onClose={() => setSelectedProject(null)} />}
    </div>
  );
};

export default App;