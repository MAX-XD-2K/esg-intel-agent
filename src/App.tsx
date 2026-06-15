/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { GoogleGenAI, Type, GenerateContentResponse } from "@google/genai";
import { 
  FileText, 
  Upload, 
  Send, 
  TrendingUp, 
  BarChart3, 
  AlertCircle, 
  CheckCircle2, 
  Loader2,
  Table as TableIcon,
  ChevronRight,
  History,
  Info,
  FileSpreadsheet,
  File,
  Image as ImageIcon
} from 'lucide-react';
import { 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  BarChart, 
  Bar,
  Legend
} from 'recharts';
import Markdown from 'react-markdown';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

// --- Utility ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

async function callWithRetry(fn: () => Promise<any>, onRetry?: (attempt: number, delay: number) => void, maxRetries = 5, initialDelay = 3000) {
  let lastError: any;
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (err: any) {
      lastError = err;
      const isRateLimit = 
        err.message?.includes('429') || 
        err.message?.includes('RESOURCE_EXHAUSTED') || 
        err.message?.includes('quota') ||
        err.status === 429 || 
        err.code === 429;

      if (isRateLimit && i < maxRetries - 1) {
        const delay = initialDelay * Math.pow(2, i);
        if (onRetry) onRetry(i + 1, delay);
        console.warn(`Rate limit hit. Retrying in ${delay}ms... (Attempt ${i + 1}/${maxRetries})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }
      throw err;
    }
  }
  throw lastError;
}

// --- Types ---
interface ESGMetric {
  year: string;
  metric_name: string;
  value: string;
  unit: string;
  category: string; // 'Environmental' | 'Social' | 'Governance' or other category found
}

interface EcosystemItem {
  name: string;
  fullName: string;
  type: string;
  shortDescription: string;
}

interface ReportingEcosystem {
  frameworks: EcosystemItem[];
  standards: EcosystemItem[];
  ratings: EcosystemItem[];
  certifications: EcosystemItem[];
  assuranceStandards: EcosystemItem[];
}

interface DataFormats {
  tables: boolean;
  charts: boolean;
  graphs: boolean;
  kpiDashboards: boolean;
  narrativeSections: boolean;
}

interface ESGDocAnalysis {
  documentType: string;
  documentSubtype: string | null;
  companyName: string | null;
  reportingYears: string[];
  summary: string;
  confidence_score: number;
  documentStructure: string[];
  reportingEcosystem: ReportingEcosystem;
  dataFormats: DataFormats;
  esgDataLocations: string[];
  reportEvolution: string[];
  metrics: ESGMetric[];
}

interface ESGDataset {
  companyName: string;
  metrics: ESGMetric[];
  summary: string;
}

interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
  data?: any;
  chartData?: any[];
  chartType?: 'line' | 'bar';
  percentageChange?: string;
}

// --- Constants ---
const SYSTEM_INSTRUCTION = `You are an ESG Intelligence Agent operating in STRICT DATA MODE.  
All analysis must follow a high-integrity, document-first methodology.

CORE OPERATING PRINCIPLES

CRITICAL RULES
1. Use ONLY:
   - the uploaded document(s), OR
   - structured data explicitly provided in the conversation.

2. NEVER use:
   - outside knowledge,
   - pre-trained corporate ESG knowledge,
   - industry assumptions,
   - inferred metrics,
   - estimated values.

3. If requested information is not explicitly present in the uploaded material, state exactly:
   "This information is not available in the uploaded document."

4. Do not:
   - guess,
   - speculate,
   - estimate,
   - infer missing values.

5. Maintain a:
   - professional,
   - board-ready,
   - audit-safe tone.

PRIMARY REPORT ANALYSIS WORKFLOW

Before answering any ESG metric question, FIRST perform a full document-level analysis.

STEP 1 — FULL DOCUMENT ANALYSIS

Analyze the complete uploaded file before extracting metrics.

Identify and summarize:
- Report Type
- Document Structure
- Main Sections
- Reporting Frameworks Mentioned
- ESG Categories Covered
- Data Presentation Formats
- Tables / Charts / KPI Sections
- Reporting Years Covered

Determine whether the uploaded file is:
- ESG Report
- Sustainability Report
- Annual Report
- Climate Report
- CSR Report
- Integrated Report
- Regulatory Filing
- Mixed-format corporate report

REQUIRED INITIAL OUTPUT FORMAT

DOCUMENT OVERVIEW

1. Report Type
Identify exactly what type of report/file it is.

2. File Structure Analysis

| Section | Purpose |
|---|---|

3. Data Format Mapping

| Report Area | Data Format |
|---|---|

4. Reporting Periods Covered
List all years explicitly mentioned in the uploaded document.

5. Frameworks or Standards Mentioned
Only include frameworks explicitly written in the document.

If absent, state:
"No reporting framework explicitly identified in the uploaded document."

REPORT EVOLUTION & FORMAT LEARNING RULE

When multiple reports, years, or versions are uploaded, analyze how the reporting structure evolves over time.

REQUIRED REPORT EVOLUTION ANALYSIS

Identify:
- Structural changes in report layout
- Added or removed ESG sections
- Changes in KPI presentation style
- Changes in reporting frameworks used
- Changes in disclosure depth
- New sustainability topics introduced
- Shifts from narrative reporting to quantitative reporting
- Changes in data visualization formats
- Movement of ESG information between sections
- Expansion or reduction of environmental/social/governance coverage

REQUIRED OUTPUT — REPORT EVOLUTION ANALYSIS

| Area | Previous Format | Current Format | Change Observed |
|---|---|---|---|

LEARNING & ADAPTIVE ANALYSIS RULE

The agent should:
- learn recurring report structures from uploaded documents,
- recognize patterns in ESG disclosure formatting,
- identify consistency or inconsistency in reporting methods,
- track how the organization evolves ESG disclosure maturity over time.

However:
- learning must remain STRICTLY limited to uploaded documents,
- no external ESG assumptions may be introduced,
- no industry benchmarking may be inferred.

SECONDARY ESG ANALYSIS RULES

ONLY after the complete report overview is finished:
- Answer detailed ESG metric questions
- Extract KPI data
- Perform trend analysis
- Compare reporting years
- Calculate percentage changes

CALCULATION RULES

When calculating percentage change, ALWAYS use:

percentage_change = ((New Value - Old Value) / Old Value) * 100

Requirements:
- Show formula
- Show source values
- Round results to exactly 2 decimal places

REQUIRED ESG METRIC RESPONSE STRUCTURE

1. DIRECT ANSWER
Provide a concise factual answer.

2. SUPPORTING DATA TABLE

| Metric | Year | Value | Unit | Source Section |
|---|---|---|---|---|

3. PERCENTAGE CHANGE
Show:
- formula
- calculation
- exact result

4. INTERPRETATION
Use ONLY these labels:
- Improving
- Declining
- Stable

Support interpretation strictly using uploaded document evidence.

5. MISSING DATA NOTE

If periods are incomplete or data is partially unavailable, append:
"Analysis limited due to incomplete reporting periods."

STRICT DATA ENFORCEMENT

If any requested ESG metric, KPI, target, or reporting year is missing:

Respond exactly:
"This information is not available in the uploaded document."

Do NOT:
- fill gaps
- estimate trends
- assume continuity
- infer unstated values

FINAL BEHAVIORAL RULE

The agent must behave like:
- a compliance-grade ESG analyst,
- an audit-safe reporting assistant,
- a document-constrained intelligence system.

All conclusions must remain fully traceable to uploaded evidence only.

If the user asks for charts, visual comparisons, or trends, you must also provide a JSON block representing data points that can be fed into a chart.`;

// --- Components ---

const MetricCard = ({ title, value, unit, trend, icon: Icon }: any) => (
  <div className="bg-zinc-900 border border-zinc-800 p-4 rounded-xl shadow-sm">
    <div className="flex items-center justify-between mb-2">
      <span className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">{title}</span>
      <Icon className="w-4 h-4 text-zinc-500" />
    </div>
    <div className="flex items-baseline gap-1">
      <span className="text-2xl font-bold text-zinc-100">{value}</span>
      <span className="text-sm text-zinc-500">{unit}</span>
    </div>
    {trend && (
      <div className={cn("text-xs mt-2 flex items-center gap-1", trend > 0 ? "text-emerald-400" : "text-rose-400")}>
        <TrendingUp className={cn("w-3 h-3", trend < 0 && "rotate-180")} />
        {Math.abs(trend)}% vs last year
      </div>
    )}
  </div>
);

interface UploadedFileMeta {
  name: string;
  type: string;
  size: string;
  kindText: string;
  metricsCount: number;
  metrics?: ESGMetric[];
  analysis: ESGDocAnalysis;
}

function formatBytes(bytes: number, decimals = 2) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function getFileKindText(type: string, name: string) {
  const ext = name.split('.').pop()?.toLowerCase();
  switch (ext) {
    case 'pdf':
      return 'PDF Document';
    case 'xlsx':
    case 'xls':
      return 'Excel Spreadsheet';
    case 'docx':
    case 'doc':
      return 'Word Document';
    case 'csv':
      return 'CSV Data File';
    case 'txt':
      return 'Plain Text Document';
    case 'md':
      return 'Markdown Document';
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'webp':
      return 'Image File';
    default:
      return type || 'Document File';
  }
}

interface AnalysisDashboardProps {
  analysis: ESGDocAnalysis;
}

function AnalysisDashboard({ analysis }: AnalysisDashboardProps) {
  const [metricsFilter, setMetricsFilter] = useState<'All' | 'Environmental' | 'Social' | 'Governance'>('All');
  const [metricsSearch, setMetricsSearch] = useState('');

  const filteredMetrics = (analysis.metrics || []).filter(m => {
    const matchesCategory = metricsFilter === 'All' || m.category?.toLowerCase() === metricsFilter.toLowerCase();
    const matchesSearch = m.metric_name?.toLowerCase().includes(metricsSearch.toLowerCase()) || 
                          m.year?.toLowerCase().includes(metricsSearch.toLowerCase());
    return matchesCategory && matchesSearch;
  });

  const confidencePercentage = Math.round((analysis.confidence_score || 0) * 100);

  return (
    <div className="flex-1 overflow-y-auto p-6 space-y-8 custom-scrollbar">
      {/* 1. Overview & Classification */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {/* Left main info */}
        <div className="md:col-span-2 bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl flex flex-col justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-2">
              <span className="text-[10px] bg-white/10 text-white border border-white/10 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider">
                {analysis.documentType || 'Document'}
              </span>
              {analysis.documentSubtype && (
                <span className="text-[10px] bg-zinc-800 text-zinc-300 border border-zinc-700/50 px-2 py-0.5 rounded-full font-semibold">
                  {analysis.documentSubtype}
                </span>
              )}
            </div>
            <h2 className="text-xl font-extrabold text-white tracking-tight">
              {analysis.companyName || 'Unknown Company'}
            </h2>
            <p className="text-xs text-zinc-500 mt-1 font-mono">
              Reporting Years: {analysis.reportingYears?.join(', ') || 'N/A'}
            </p>
          </div>
          <div>
            <p className="text-xs font-bold text-zinc-500 uppercase tracking-wider mb-1">Executive Summary</p>
            <p className="text-sm text-zinc-300 leading-relaxed italic">
              "{analysis.summary || 'No summary available.'}"
            </p>
          </div>
        </div>

        {/* Right confidence score */}
        <div className="bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl flex flex-col items-center justify-center text-center">
          <p className="text-xs font-bold text-zinc-500 uppercase tracking-wider mb-4">Confidence Score</p>
          <div className="relative w-24 h-24 flex items-center justify-center">
            {/* Radial progress ring */}
            <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
              <circle 
                className="text-zinc-800" 
                strokeWidth="8" 
                stroke="currentColor" 
                fill="transparent" 
                r="40" 
                cx="50" 
                cy="50" 
              />
              <circle 
                className={cn(
                  confidencePercentage > 80 ? "text-emerald-500" : confidencePercentage > 50 ? "text-amber-500" : "text-rose-500"
                )}
                strokeWidth="8" 
                strokeDasharray={2 * Math.PI * 40}
                strokeDashoffset={2 * Math.PI * 40 * (1 - confidencePercentage / 100)}
                strokeLinecap="round" 
                stroke="currentColor" 
                fill="transparent" 
                r="40" 
                cx="50" 
                cy="50" 
              />
            </svg>
            <span className="absolute text-xl font-black text-white">{confidencePercentage}%</span>
          </div>
          <p className="text-[10px] text-zinc-500 mt-3 max-w-[150px]">
            Based exclusively on explicit document evidence
          </p>
        </div>
      </div>

      {/* 2. Structure & Formats Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Document Structure */}
        <div className="bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl flex flex-col gap-3">
          <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
            <FileText className="w-3.5 h-3.5" />
            Document Structure
          </h3>
          <div className="max-h-60 overflow-y-auto space-y-1.5 pr-2 custom-scrollbar">
            {analysis.documentStructure && analysis.documentStructure.length > 0 ? (
              analysis.documentStructure.map((section, sIdx) => (
                <div key={sIdx} className="flex items-start gap-2 text-xs py-1 border-b border-zinc-800/30 last:border-0">
                  <ChevronRight className="w-3 h-3 text-zinc-600 shrink-0 mt-0.5" />
                  <span className="text-zinc-300 font-medium">{section}</span>
                </div>
              ))
            ) : (
              <p className="text-xs text-zinc-500 italic">No structure parsed.</p>
            )}
          </div>
        </div>

        {/* Data Formats & Locations */}
        <div className="bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl flex flex-col gap-5">
          <div className="flex flex-col gap-3">
            <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
              <BarChart3 className="w-3.5 h-3.5" />
              Detected Data Formats
            </h3>
            <div className="grid grid-cols-2 gap-2">
              {Object.entries(analysis.dataFormats || {}).map(([key, value]) => {
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
                return (
                  <div key={key} className="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-zinc-900/50 border border-zinc-800/50">
                    <div className={cn(
                      "w-2 h-2 rounded-full",
                      value ? "bg-emerald-500 shadow-sm shadow-emerald-500/20" : "bg-zinc-700"
                    )} />
                    <span className={cn("text-xs font-medium", value ? "text-zinc-200" : "text-zinc-500")}>
                      {label}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="flex flex-col gap-3">
            <h4 className="text-xs font-bold text-zinc-500 uppercase tracking-wider">
              ESG Data Locations
            </h4>
            <div className="flex flex-wrap gap-2">
              {analysis.esgDataLocations && analysis.esgDataLocations.length > 0 ? (
                analysis.esgDataLocations.map((loc, lIdx) => (
                  <span key={lIdx} className="text-[10px] px-2.5 py-1 rounded-lg bg-zinc-900 border border-zinc-800 text-zinc-400 font-medium">
                    {loc}
                  </span>
                ))
              ) : (
                <span className="text-xs text-zinc-500 italic">No locations specified.</span>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* 3. Reporting Ecosystem Registry */}
      <div className="bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl space-y-4">
        <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
          <Info className="w-3.5 h-3.5" />
          Reporting Ecosystem & Standards
        </h3>
        
        <div className="space-y-4 divide-y divide-zinc-800/40">
          {Object.entries(analysis.reportingEcosystem || {}).map(([category, items]) => {
            const displayTitle = category.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
            return (
              <div key={category} className="pt-4 first:pt-0 flex flex-col md:flex-row md:items-start gap-4">
                <div className="md:w-48 shrink-0">
                  <span className="text-xs font-bold text-zinc-400 uppercase tracking-wider">
                    {displayTitle}
                  </span>
                </div>
                <div className="flex-1 flex flex-wrap gap-2">
                  {items && items.length > 0 ? (
                    items.map((item: EcosystemItem, iIdx: number) => (
                      <div 
                        key={iIdx} 
                        className="relative group cursor-help px-3 py-1.5 bg-zinc-900 hover:bg-zinc-800 border border-zinc-800 hover:border-zinc-700 rounded-xl transition-all"
                      >
                        <span className="text-xs font-semibold text-zinc-200">{item.name}</span>
                        {/* Custom Hover Tooltip */}
                        <div className="absolute z-50 bottom-full left-1/2 -translate-x-1/2 mb-2 w-64 p-3 bg-zinc-950 border border-zinc-800 rounded-xl shadow-xl opacity-0 group-hover:opacity-100 pointer-events-none transition-all duration-200 flex flex-col gap-1 text-left">
                          <span className="text-xs font-bold text-zinc-100">{item.fullName || item.name}</span>
                          <span className="text-[9px] uppercase font-bold tracking-wider text-zinc-500">{item.type}</span>
                          <p className="text-[10px] leading-normal text-zinc-400 mt-1">{item.shortDescription}</p>
                        </div>
                      </div>
                    ))
                  ) : (
                    <span className="text-xs text-zinc-500 italic">None identified</span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* 4. Report Evolution */}
      {analysis.reportEvolution && analysis.reportEvolution.length > 0 && (
        <div className="bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl space-y-3">
          <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
            <TrendingUp className="w-3.5 h-3.5" />
            Report Evolution & Signals
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {analysis.reportEvolution.map((signal, sIdx) => (
              <div key={sIdx} className="p-3 rounded-xl bg-zinc-900/30 border border-zinc-800/50 flex items-start gap-3">
                <div className="w-1.5 h-1.5 rounded-full bg-emerald-400 mt-1.5 shrink-0" />
                <p className="text-xs text-zinc-300 leading-normal">{signal}</p>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 5. Metrics Explorer */}
      <div className="bg-zinc-900/40 border border-zinc-800/80 p-6 rounded-2xl space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
            <TableIcon className="w-3.5 h-3.5" />
            Extracted ESG Metrics ({filteredMetrics.length})
          </h3>
          
          <div className="flex flex-wrap items-center gap-3">
            {/* Search Input */}
            <input 
              type="text" 
              placeholder="Search metrics or year..." 
              value={metricsSearch}
              onChange={(e) => setMetricsSearch(e.target.value)}
              className="bg-zinc-900/80 border border-zinc-800 rounded-lg px-3 py-1.5 text-xs text-white placeholder:text-zinc-500 focus:outline-none focus:border-zinc-700"
            />
            {/* Filter Tabs */}
            <div className="flex rounded-lg bg-zinc-950 p-1 border border-zinc-800">
              {(['All', 'Environmental', 'Social', 'Governance'] as const).map(tab => (
                <button 
                  key={tab} 
                  onClick={() => setMetricsFilter(tab)}
                  className={cn(
                    "px-3 py-1 rounded-md text-xs font-semibold transition-all",
                    metricsFilter === tab 
                      ? "bg-zinc-900 text-white border border-zinc-800" 
                      : "text-zinc-400 hover:text-zinc-200"
                  )}
                >
                  {tab}
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* Metrics Table */}
        <div className="overflow-x-auto border border-zinc-800 rounded-xl bg-zinc-950">
          <table className="w-full text-xs text-left">
            <thead className="bg-zinc-900/80 border-b border-zinc-800 text-zinc-500 font-bold uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3">Metric Name</th>
                <th className="px-4 py-3 w-24">Year</th>
                <th className="px-4 py-3 w-40">Value & Unit</th>
                <th className="px-4 py-3 w-32">Category</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-800/40">
              {filteredMetrics.length > 0 ? (
                filteredMetrics.map((m, mIdx) => {
                  const isEnv = m.category?.toLowerCase() === 'environmental';
                  const isSoc = m.category?.toLowerCase() === 'social';
                  const badgeColor = isEnv 
                    ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' 
                    : isSoc 
                      ? 'bg-blue-500/10 text-blue-400 border-blue-500/20' 
                      : 'bg-amber-500/10 text-amber-400 border-amber-500/20';

                  return (
                    <tr key={mIdx} className="hover:bg-zinc-900/20 transition-all">
                      <td className="px-4 py-3 font-semibold text-zinc-200">{m.metric_name}</td>
                      <td className="px-4 py-3 font-mono text-zinc-400">{m.year}</td>
                      <td className="px-4 py-3 font-semibold text-zinc-100">
                        {m.value} <span className="text-zinc-500 font-normal ml-0.5">{m.unit}</span>
                      </td>
                      <td className="px-4 py-3">
                        <span className={cn("text-[9px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full border", badgeColor)}>
                          {m.category}
                        </span>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={4} className="px-4 py-6 text-center text-zinc-500 italic">
                    No matching metrics found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [isIngesting, setIsIngesting] = useState(false);
  const [retryStatus, setRetryStatus] = useState<string | null>(null);
  const [dataset, setDataset] = useState<ESGDataset | null>(null);
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFileMeta[]>([]);
  const [selectedFileIndex, setSelectedFileIndex] = useState<number | null>(null);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'chat'>('dashboard');
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState('');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [messages]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    setIsIngesting(true);
    setError(null);

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY! });
      
      // Concurrency control: process max 2 files at a time to avoid aggressive 429s
      const CONCURRENCY_LIMIT = 2;
      const results: Array<{
        fileResult: ESGDocAnalysis;
        fileName: string;
        fileSize: number;
        fileType: string;
      }> = [];
      const queue = [...files];
      let processedCount = 0;

      const processQueue = async () => {
        while (queue.length > 0) {
          const file = queue.shift();
          if (!file) break;

          setRetryStatus(`Scanning ${file.name} (${++processedCount}/${files.length})...`);
          
          let part: any;
          const ext = file.name.split('.').pop()?.toLowerCase();
          
          if (file.type === 'application/pdf' || 
              file.type.startsWith('image/') || 
              file.type.startsWith('audio/') || 
              file.type.startsWith('video/')) {
            const reader = new FileReader();
            const base64Promise = new Promise<string>((resolve) => {
              reader.onload = () => {
                const base64 = (reader.result as string).split(',')[1];
                resolve(base64);
              };
            });
            reader.readAsDataURL(file);
            const base64Data = await base64Promise;
            part = { inlineData: { data: base64Data, mimeType: file.type } };
          } else if (ext === 'xlsx' || ext === 'xls') {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const csv = XLSX.utils.sheet_to_csv(worksheet);
            part = { text: `Excel Data from ${file.name} (Sheet: ${firstSheetName}):\n${csv}` };
          } else if (ext === 'docx') {
            const buffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer: buffer });
            part = { text: `Word Document Data from ${file.name}:\n${result.value}` };
          } else {
            try {
              const text = await file.text();
              const isBinary = /[\x00-\x08\x0E-\x1F]/.test(text.slice(0, 1000));
              if (isBinary) {
                const reader = new FileReader();
                const base64Promise = new Promise<string>((resolve) => {
                  reader.onload = () => {
                    const base64 = (reader.result as string).split(',')[1];
                    resolve(base64);
                  };
                });
                reader.readAsDataURL(file);
                const base64Data = await base64Promise;
                part = { inlineData: { data: base64Data, mimeType: file.type || 'application/octet-stream' } };
              } else {
                part = { text: `File Content from ${file.name}:\n${text}` };
              }
            } catch (err) {
              const reader = new FileReader();
              const base64Promise = new Promise<string>((resolve) => {
                reader.onload = () => {
                  const base64 = (reader.result as string).split(',')[1];
                  resolve(base64);
                };
              });
              reader.readAsDataURL(file);
              const base64Data = await base64Promise;
              part = { inlineData: { data: base64Data, mimeType: file.type || 'application/octet-stream' } };
            }
          }

          const ecosystemItemSchema = {
            type: Type.OBJECT,
            properties: {
              name: { type: Type.STRING },
              fullName: { type: Type.STRING },
              type: { type: Type.STRING },
              shortDescription: { type: Type.STRING }
            },
            required: ["name", "fullName", "type", "shortDescription"]
          };

          const docAnalysisSchema = {
            type: Type.OBJECT,
            properties: {
              documentType: { type: Type.STRING },
              documentSubtype: { type: Type.STRING, nullable: true },
              companyName: { type: Type.STRING, nullable: true },
              reportingYears: { type: Type.ARRAY, items: { type: Type.STRING } },
              summary: { type: Type.STRING },
              confidence_score: { type: Type.NUMBER },
              documentStructure: { type: Type.ARRAY, items: { type: Type.STRING } },
              reportingEcosystem: {
                type: Type.OBJECT,
                properties: {
                  frameworks: { type: Type.ARRAY, items: ecosystemItemSchema },
                  standards: { type: Type.ARRAY, items: ecosystemItemSchema },
                  ratings: { type: Type.ARRAY, items: ecosystemItemSchema },
                  certifications: { type: Type.ARRAY, items: ecosystemItemSchema },
                  assuranceStandards: { type: Type.ARRAY, items: ecosystemItemSchema }
                },
                required: ["frameworks", "standards", "ratings", "certifications", "assuranceStandards"]
              },
              dataFormats: {
                type: Type.OBJECT,
                properties: {
                  tables: { type: Type.BOOLEAN },
                  charts: { type: Type.BOOLEAN },
                  graphs: { type: Type.BOOLEAN },
                  kpiDashboards: { type: Type.BOOLEAN },
                  narrativeSections: { type: Type.BOOLEAN }
                },
                required: ["tables", "charts", "graphs", "kpiDashboards", "narrativeSections"]
              },
              esgDataLocations: { type: Type.ARRAY, items: { type: Type.STRING } },
              reportEvolution: { type: Type.ARRAY, items: { type: Type.STRING } },
              metrics: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    year: { type: Type.STRING },
                    metric_name: { type: Type.STRING },
                    value: { type: Type.STRING },
                    unit: { type: Type.STRING },
                    category: { type: Type.STRING }
                  },
                  required: ["year", "metric_name", "value", "unit", "category"]
                }
              }
            },
            required: [
              "documentType",
              "documentSubtype",
              "companyName",
              "reportingYears",
              "summary",
              "confidence_score",
              "documentStructure",
              "reportingEcosystem",
              "dataFormats",
              "esgDataLocations",
              "reportEvolution",
              "metrics"
            ]
          };

          const documentPrompt = `You are an expert Document Classification and ESG Extraction Engine operating in STRICT DOCUMENT MODE.

Analyze ONLY the provided document content.

RULES:
1. Use only information explicitly present in the document.
2. Do not use external knowledge for ESG metrics, company information, reporting performance, or missing disclosures.
3. Do not infer missing values.
4. If information cannot be found, return null (or appropriate default for types, like false for boolean or empty array for lists).
5. Return ONLY valid JSON.
6. Do not include markdown, explanations, comments, or additional text.
7. All extracted information must be traceable to content explicitly present in the provided document.

TASKS:
1. Classify the document type using exactly one of: "Report", "Invoice", "Resume", "Medical Record", "Prescription", "Legal Document", "Bank Statement", "Research Paper", "Academic Certificate", "Other".
2. Determine document subtype when applicable: "Blood Report", "MRI Report", "CT Scan Report", "X-Ray Report", "Pathology Report", "Diagnostic Report", "ESG Report", "Sustainability Report", "Annual Report", "Integrated Report", "Climate Report", "CSR Report", "Other Report". If not applicable, return null.
3. Extract: companyName, reportingYears, and summary (maximum 25 words).
4. Generate confidence_score between 0.0 and 1.0 based only on document evidence.
5. Extract document structure. Return all major sections and subsections when available.
6. Extract reporting frameworks, standards, certifications, ratings, and assurance references explicitly mentioned in the document. Organise them under reportingEcosystem.
7. Extract all key metrics and indicators explicitly stated in the document (such as ESG parameters, blood test values, invoice amounts/costs, resume experience duration, academic grades, bank transactions, etc.). Organise them under metrics.
For each metric, determine year, metric_name, value, unit, and category. If standard ESG categories (Environmental, Social, Governance) do not apply, use appropriate domain categories (e.g., Financial, Hematology, Experience, Education, Personal, etc.).
8. Identify ESG data locations within the document.
9. Detect whether the document contains: Tables, Charts, Graphs, KPI Dashboards, Narrative Sections.
10. If multiple reporting years are present, identify reporting evolution signals under reportEvolution.

Use the specified JSON schema structure.`;

          const response = await callWithRetry(() => ai.models.generateContent({
            model: "gemini-3.5-flash",
            contents: [
              {
                parts: [
                  part,
                  { text: documentPrompt }
                ]
              }
            ],
            config: {
              responseMimeType: "application/json",
              responseSchema: docAnalysisSchema as any
            }
          }), (attempt, delay) => {
            setRetryStatus(`Rate limit hit. Resuming in ${Math.round(delay/1000)}s...`);
          });

          const fileResult = JSON.parse(response.text || '{}') as ESGDocAnalysis;
          results.push({
            fileResult,
            fileName: file.name,
            fileSize: file.size,
            fileType: file.type
          });
        }
      };

      // Run workers in parallel up to limit
      const workers = Array(Math.min(CONCURRENCY_LIMIT, files.length))
        .fill(null)
        .map(() => processQueue());
      
      await Promise.all(workers);
      
      let mergedDataset: ESGDataset = dataset 
        ? { ...dataset, metrics: [...dataset.metrics] } 
        : { companyName: '', summary: '', metrics: [] };

      const newUploadedFiles: UploadedFileMeta[] = [];

      for (const item of results) {
        const { fileResult, fileName, fileSize, fileType } = item;
        if (!mergedDataset.companyName) mergedDataset.companyName = fileResult.companyName || 'Unknown Company';
        mergedDataset.summary += (mergedDataset.summary ? " | " : "") + fileResult.summary;
        if (fileResult.metrics) {
          const mapped = fileResult.metrics.map(m => ({
            year: m.year,
            metric_name: m.metric_name,
            value: m.value,
            unit: m.unit,
            category: m.category
          }));
          mergedDataset.metrics = [...mergedDataset.metrics, ...mapped];
        }

        newUploadedFiles.push({
          name: fileName,
          type: fileType,
          size: formatBytes(fileSize),
          kindText: getFileKindText(fileType, fileName),
          metricsCount: fileResult.metrics?.length || 0,
          metrics: fileResult.metrics ? fileResult.metrics.map(m => ({
            year: m.year,
            metric_name: m.metric_name,
            value: m.value,
            unit: m.unit,
            category: m.category
          })) : [],
          analysis: fileResult
        });
      }

      const prevLen = uploadedFiles.length;
      setUploadedFiles(prev => [...prev, ...newUploadedFiles]);
      if (selectedFileIndex === null && newUploadedFiles.length > 0) {
        setSelectedFileIndex(prevLen);
      }

      // Deduplicate metrics if necessary (same year, same metric name)
      const uniqueMetrics = mergedDataset.metrics.reduce((acc: ESGMetric[], current) => {
        const x = acc.find(item => item.year === current.year && item.metric_name === current.metric_name);
        if (!x) {
          return acc.concat([current]);
        } else {
          return acc;
        }
      }, []);
      mergedDataset.metrics = uniqueMetrics;

      setDataset(mergedDataset);
      const newMsg: ChatMessage = {
        role: 'assistant',
        content: dataset 
          ? `Added data from **${files.length}** new file(s). Total unique metrics: ${mergedDataset.metrics.length}.`
          : `Successfully ingested and classified **${files.length}** file(s) for **${mergedDataset.companyName}**. 

I have performed a strict document extraction and classification. You can view the full details in the **Analysis Dashboard** or ask questions in the **Interactive Chat** tab. How can I help you today?`
      };
      setMessages(prev => [...prev, newMsg]);
    } catch (err: any) {
      console.error(err);
      setError("Failed to process the documents. Please ensure they are valid PDFs or images containing ESG data.");
    } finally {
      setIsIngesting(false);
      setRetryStatus(null);
    }
  };

  const handleSendMessage = async () => {
    if (!input.trim() || !dataset || isAnalyzing) return;

    const userMessage = input;
    setInput('');
    setMessages(prev => [...prev, { role: 'user', content: userMessage }]);
    setIsAnalyzing(true);

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY! });
      const response = await callWithRetry(() => ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: [
          {
            parts: [
              { text: `Context Dataset: ${JSON.stringify(dataset)}` },
              ...(selectedFileIndex !== null && uploadedFiles[selectedFileIndex]
                ? [{ text: `Active Document Analysis: ${JSON.stringify(uploadedFiles[selectedFileIndex].analysis)}` }]
                : []),
              { text: `User Question: ${userMessage}` }
            ]
          }
        ],
        config: {
          systemInstruction: SYSTEM_INSTRUCTION,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              answer: { type: Type.STRING, description: "The direct answer and interpretation in markdown format." },
              tableData: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    Year: { type: Type.STRING },
                    Metric: { type: Type.STRING },
                    Value: { type: Type.STRING }
                  }
                }
              },
              percentageChange: { type: Type.STRING },
              chartData: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    name: { type: Type.STRING, description: "Year or Label" },
                    value: { type: Type.NUMBER }
                  }
                }
              },
              chartType: { type: Type.STRING, enum: ["line", "bar"] },
              missingData: { type: Type.STRING }
            },
            required: ["answer"]
          }
        }
      }), (attempt, delay) => {
        setRetryStatus(`Optimizing throughput... Retrying in ${Math.round(delay/1000)}s...`);
      });

      const result = JSON.parse(response.text || '{}');
      setMessages(prev => [...prev, { 
        role: 'assistant', 
        content: result.answer,
        data: result.tableData,
        chartData: result.chartData,
        chartType: result.chartType,
        percentageChange: result.percentageChange
      } as any]);
    } catch (err) {
      console.error(err);
      setMessages(prev => [...prev, { role: 'assistant', content: "I encountered an error while analyzing the data. Please try rephrasing your question." }]);
    } finally {
      setIsAnalyzing(false);
      setRetryStatus(null);
    }
  };

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 font-sans selection:bg-zinc-800">
      {/* Header */}
      <header className="sticky top-0 z-50 bg-zinc-900/80 backdrop-blur-md border-b border-zinc-800 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-white rounded-lg flex items-center justify-center">
            <BarChart3 className="text-zinc-900 w-6 h-6" />
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight text-white">ESG Intel Agent</h1>
            <p className="text-[10px] uppercase tracking-[0.2em] font-semibold text-zinc-500">Strict Data Mode Active</p>
          </div>
        </div>
        
        {dataset && (
          <div className="flex items-center gap-4">
            <div className="px-3 py-1 bg-emerald-500/10 border border-emerald-500/20 rounded-full flex items-center gap-2">
              <CheckCircle2 className="w-3 h-3 text-emerald-400" />
              <span className="text-xs font-medium text-emerald-400">{dataset.companyName} Loaded</span>
            </div>

            <label className={cn(
              "cursor-pointer text-xs font-medium text-zinc-400 hover:text-white transition-colors flex items-center gap-1.5",
              isIngesting && "opacity-50 pointer-events-none"
            )}>
              {isIngesting ? <Loader2 className="w-3 h-3 animate-spin" /> : <Upload className="w-3 h-3" />}
              {isIngesting ? "Processing..." : "Add Files"}
              <input type="file" className="hidden" onChange={handleFileUpload} disabled={isIngesting} multiple />
            </label>

            <button 
              onClick={() => { setDataset(null); setMessages([]); setUploadedFiles([]); }}
              className="text-xs font-medium text-zinc-500 hover:text-zinc-100 transition-colors"
            >
              Reset Session
            </button>
          </div>
        )}
      </header>

      <main className="max-w-[1600px] mx-auto p-6 grid grid-cols-1 lg:grid-cols-12 gap-6 h-[calc(100vh-88px)]">
        {/* Left Sidebar: Data Overview */}
        <div className="lg:col-span-4 space-y-6 overflow-y-auto pr-2">
          {!dataset ? (
            <div className="h-full flex flex-col items-center justify-center border-2 border-dashed border-zinc-800 rounded-3xl bg-transparent p-8 text-center">
              <div className="w-16 h-16 bg-zinc-900 border border-zinc-800 rounded-2xl flex items-center justify-center mb-6">
                <Upload className="w-8 h-8 text-zinc-500" />
              </div>
              <h2 className="text-xl font-bold mb-2 text-white">Upload Documents</h2>
              <p className="text-sm text-zinc-500 mb-8 max-w-xs">
                Upload any report, medical record, resume, invoice, spreadsheet, text, code, or media file to begin your analysis.
              </p>
              <label className="relative group cursor-pointer">
                <div className={cn(
                  "px-8 py-3 bg-white text-zinc-900 rounded-xl font-medium transition-all group-hover:scale-105 active:scale-95 flex items-center gap-2",
                  isIngesting && "opacity-50 pointer-events-none"
                )}>
                  {isIngesting ? <Loader2 className="w-4 h-4 animate-spin" /> : <FileText className="w-4 h-4" />}
                  {isIngesting ? (retryStatus || "Processing Files...") : "Select Documents"}
                </div>
                <input type="file" className="hidden" onChange={handleFileUpload} disabled={isIngesting} multiple />
              </label>
              {error && (
                <div className="mt-6 p-4 bg-rose-500/10 border border-rose-500/20 rounded-xl flex items-start gap-3 text-left">
                  <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
                  <p className="text-xs text-rose-400 leading-relaxed">{error}</p>
                </div>
              )}
            </div>
          ) : (
            <>
              {/* Loaded Documents & Categories */}
              <section className="space-y-4">
                <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
                  <File className="w-3.5 h-3.5 text-zinc-400" />
                  Ingested Files ({uploadedFiles.length})
                </h3>
                <div className="space-y-3">
                  {uploadedFiles.map((file, idx) => {
                    const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
                    const isCsv = file.name.endsWith('.csv');
                    const isWord = file.name.endsWith('.docx') || file.name.endsWith('.doc');
                    const isPdf = file.name.endsWith('.pdf');
                    const isImg = file.type?.startsWith('image/');
                    
                    let FileIconComp = FileText;
                    let iconColor = 'text-blue-400 bg-blue-500/10 border-blue-500/20';
                    
                    if (isExcel || isCsv) {
                      FileIconComp = FileSpreadsheet;
                      iconColor = 'text-emerald-400 bg-emerald-500/10 border-emerald-500/20';
                    } else if (isPdf) {
                      iconColor = 'text-rose-400 bg-rose-500/10 border-rose-500/20';
                    } else if (isImg) {
                      FileIconComp = ImageIcon;
                      iconColor = 'text-purple-400 bg-purple-500/10 border-purple-500/20';
                    } else if (isWord) {
                      iconColor = 'text-indigo-400 bg-indigo-500/10 border-indigo-500/20';
                    }
                    
                    return (
                      <div 
                        key={idx} 
                        onClick={() => { setSelectedFileIndex(idx); setActiveTab('dashboard'); }}
                        className={cn(
                          "cursor-pointer bg-transparent border rounded-2xl flex flex-col gap-3 hover:bg-zinc-900/30 hover:border-zinc-700/80 transition-all text-left",
                          selectedFileIndex === idx 
                            ? "border-white bg-zinc-900/40" 
                            : "border-zinc-800/80"
                        )}
                      >
                        <div className="flex items-start gap-3">
                          <div className={cn("p-2.5 rounded-xl shrink-0 border", iconColor)}>
                            <FileIconComp className="w-5 h-5" />
                          </div>
                          <div className="min-w-0 flex-1">
                            <p className="text-sm font-semibold text-zinc-100 truncate" title={file.name}>
                              {file.name}
                            </p>
                            <div className="flex items-center gap-2 mt-1.5 flex-wrap">
                              <span className="text-[10px] bg-zinc-900/40 text-zinc-400 border border-zinc-800 px-2 py-0.5 rounded-full font-medium">
                                {file.kindText}
                              </span>
                              <span className="text-[10px] text-zinc-500 font-mono">
                                {file.size}
                              </span>
                            </div>
                            {file.metricsCount > 0 ? (
                              <p className="text-[10px] text-emerald-400 font-medium mt-2 flex items-center gap-1">
                                <CheckCircle2 className="w-2.5 h-2.5" />
                                Holds {file.metricsCount} metrics
                              </p>
                            ) : (
                              <p className="text-[10px] text-zinc-500 italic mt-2">
                                Parsed content successfully
                              </p>
                            )}
                          </div>
                        </div>

                        {file.metrics && file.metrics.length > 0 && (
                          <div className="pt-3 border-t border-zinc-800/40 space-y-2">
                            <p className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">
                              Extracted Metrics
                            </p>
                            <div className="max-h-48 overflow-y-auto space-y-2 pr-1 divide-y divide-zinc-800/40 custom-scrollbar">
                              {file.metrics.map((m, mIdx) => {
                                const isEnv = m.category === 'Environmental';
                                const isSoc = m.category === 'Social';
                                const badgeColor = isEnv 
                                  ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' 
                                  : isSoc 
                                    ? 'bg-blue-500/10 text-blue-400 border-blue-500/20' 
                                    : 'bg-amber-500/10 text-amber-400 border-amber-500/20';

                                return (
                                  <div key={mIdx} className="pt-2 first:pt-0 flex flex-col gap-0.5 text-xs">
                                    <div className="flex items-start justify-between gap-2">
                                      <span className="font-medium text-zinc-300 leading-tight">
                                        {m.metric_name}
                                      </span>
                                      <span className={cn("text-[8px] font-mono shrink-0 px-1.5 py-0.5 rounded-md border", badgeColor)}>
                                        {m.year}
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between mt-1 text-[11px]">
                                      <span className="text-zinc-500 text-[10px]">{m.category}</span>
                                      <span className="font-semibold text-zinc-200 font-mono">
                                        {m.value} <span className="text-zinc-400 font-normal text-[10px]">{m.unit}</span>
                                      </span>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </section>

              {/* Executive Summary Section */}
              <section className="space-y-4">
                <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
                  <Info className="w-3.5 h-3.5 text-zinc-400" />
                  Executive Summary
                </h3>
                <div className="bg-transparent border border-zinc-800/80 p-5 rounded-2xl shadow-sm leading-relaxed">
                  <p className="text-sm text-zinc-400 italic">
                    "{dataset.summary}"
                  </p>
                </div>
              </section>
            </>
          )}
        </div>

        {/* Middle Content: Chat & Analysis */}
        <div className="lg:col-span-8 flex flex-col bg-zinc-900 border border-zinc-800 rounded-3xl shadow-sm overflow-hidden">
          {/* Tabs Navigation Header */}
          <div className="flex items-center justify-between border-b border-zinc-800 px-6 py-3 bg-zinc-900/85 backdrop-blur-sm">
            <div className="flex gap-4">
              <button 
                onClick={() => setActiveTab('dashboard')}
                className={cn(
                  "px-4 py-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all",
                  activeTab === 'dashboard' 
                    ? "border-white text-white" 
                    : "border-transparent text-zinc-500 hover:text-zinc-300"
                )}
              >
                Analysis Dashboard
              </button>
              <button 
                onClick={() => setActiveTab('chat')}
                className={cn(
                  "px-4 py-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all",
                  activeTab === 'chat' 
                    ? "border-white text-white" 
                    : "border-transparent text-zinc-500 hover:text-zinc-300"
                )}
              >
                Interactive Chat
              </button>
            </div>
            {selectedFileIndex !== null && uploadedFiles[selectedFileIndex] && (
              <div className="text-[10px] text-zinc-400 font-mono truncate max-w-xs">
                Active: <span className="text-zinc-200 font-semibold">{uploadedFiles[selectedFileIndex].name}</span>
              </div>
            )}
          </div>

          {activeTab === 'dashboard' ? (
            selectedFileIndex !== null && uploadedFiles[selectedFileIndex] ? (
              <AnalysisDashboard analysis={uploadedFiles[selectedFileIndex].analysis} />
            ) : (
              <div className="flex-1 flex flex-col items-center justify-center text-center p-8 text-zinc-500 italic">
                No active document selected.
              </div>
            )
          ) : (
            <>
              {/* Chat Messages */}
              <div ref={scrollRef} className="flex-1 overflow-y-auto p-6 space-y-8">
                {messages.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center text-center space-y-4 opacity-20">
                    <History className="w-12 h-12 text-zinc-500" />
                    <p className="text-sm font-medium text-zinc-400">No analysis history yet.</p>
                  </div>
                ) : (
                  messages.map((msg, i) => (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      key={i} 
                      className={cn(
                        "flex flex-col gap-4",
                        msg.role === 'user' ? "items-end" : "items-start"
                      )}
                    >
                      <div className={cn(
                        "max-w-[85%] p-4 rounded-2xl text-sm leading-relaxed",
                        msg.role === 'user' 
                          ? "bg-white text-zinc-900 rounded-tr-none" 
                          : "bg-zinc-800 text-zinc-100 border border-zinc-700 rounded-tl-none"
                      )}>
                        <div className="prose prose-sm prose-invert prose-zinc max-w-none">
                          <Markdown>{msg.content}</Markdown>
                        </div>
                        {msg.percentageChange && (
                          <div className="mt-3 pt-3 border-t border-zinc-700 flex items-center gap-2 text-xs font-bold text-zinc-500 uppercase tracking-wider">
                            <TrendingUp className="w-3 h-3" />
                            Change: <span className="text-zinc-100">{msg.percentageChange}</span>
                          </div>
                        )}
                      </div>

                      {msg.data && (
                        <div className="w-full max-w-[90%] overflow-x-auto border border-zinc-700 rounded-xl bg-zinc-900 shadow-sm">
                          <table className="w-full text-xs text-left">
                            <thead className="bg-zinc-800 border-b border-zinc-700">
                              <tr>
                                {Object.keys(msg.data[0]).map(k => (
                                  <th key={k} className="px-4 py-2 font-bold text-zinc-500 uppercase tracking-wider">{k}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-zinc-800">
                              {msg.data.map((row: any, ri: number) => (
                                <tr key={ri}>
                                  {Object.values(row).map((v: any, ci: number) => (
                                    <td key={ci} className="px-4 py-2 font-medium text-zinc-300">{v}</td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}

                      {msg.chartData && (
                        <div className="w-full max-w-[90%] h-64 bg-zinc-900 border border-zinc-700 rounded-2xl p-4 shadow-sm">
                          <ResponsiveContainer width="100%" height="100%">
                            {msg.chartType === 'bar' ? (
                              <BarChart data={msg.chartData}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#27272a" />
                                <XAxis dataKey="name" fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                                <YAxis fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                                <Tooltip 
                                  contentStyle={{ backgroundColor: '#18181b', border: '1px solid #27272a', borderRadius: '12px' }}
                                />
                                <Bar dataKey="value" fill="#f4f4f5" radius={[4, 4, 0, 0]} />
                              </BarChart>
                            ) : (
                              <LineChart data={msg.chartData}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#27272a" />
                                <XAxis dataKey="name" fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                                <YAxis fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                                <Tooltip 
                                  contentStyle={{ backgroundColor: '#18181b', border: '1px solid #27272a', borderRadius: '12px' }}
                                />
                                <Line type="monotone" dataKey="value" stroke="#f4f4f5" strokeWidth={2} dot={{ r: 4, fill: '#f4f4f5' }} />
                              </LineChart>
                            )}
                          </ResponsiveContainer>
                        </div>
                      )}
                    </motion.div>
                  ))
                )}
              </div>

              {/* Input Area */}
              <div className="p-6 bg-zinc-900 border-t border-zinc-800">
                <div className="relative group">
                  <input 
                    type="text"
                    value={input}
                    onChange={(e) => setInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleSendMessage()}
                    placeholder={dataset ? "Ask about Scope 1 emissions, diversity ratios, or trends..." : "Upload a document to start analysis"}
                    disabled={!dataset || isAnalyzing}
                    className="w-full bg-zinc-800 border border-zinc-700 rounded-2xl px-5 py-4 pr-14 text-sm text-white placeholder:text-zinc-500 focus:outline-none focus:ring-2 focus:ring-white/5 focus:border-zinc-500 transition-all disabled:opacity-50"
                  />
                  <button 
                    onClick={handleSendMessage}
                    disabled={!dataset || isAnalyzing || !input.trim()}
                    className="absolute right-3 top-1/2 -translate-y-1/2 w-10 h-10 bg-white text-zinc-900 rounded-xl flex items-center justify-center hover:scale-105 active:scale-95 transition-all disabled:opacity-50 disabled:hover:scale-100"
                  >
                    {isAnalyzing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Send className="w-4 h-4" />}
                  </button>
                </div>
                {retryStatus && (
                  <div className="mt-2 text-[10px] font-bold text-rose-400 uppercase tracking-wider animate-pulse">
                    {retryStatus}
                  </div>
                )}
                <div className="mt-3 flex items-center gap-4 px-2">
                  <div className="flex items-center gap-1.5">
                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                    <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">Agent Online</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <div className="w-1.5 h-1.5 rounded-full bg-zinc-700" />
                    <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">Strict Mode</span>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>

      </main>
    </div>
  );
}
