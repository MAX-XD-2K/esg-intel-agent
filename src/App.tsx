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
        err.message?.includes('503') ||
        err.message?.includes('UNAVAILABLE') ||
        err.message?.includes('temporary') ||
        err.status === 429 || 
        err.code === 429 ||
        err.status === 503 ||
        err.code === 503 ||
        err.status === 'UNAVAILABLE';

      if (isRateLimit && i < maxRetries - 1) {
        const delay = initialDelay * Math.pow(2, i);
        if (onRetry) onRetry(i + 1, delay);
        console.warn(`Rate limit/transient error hit. Retrying in ${delay}ms... (Attempt ${i + 1}/${maxRetries})`);
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
  suggestedPeer?: {
    companyName: string;
    summary: string;
    metrics: ESGMetric[];
  };
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
          <div className="flex-1 flex flex-col min-h-0">
            <p className="text-xs font-bold text-zinc-500 uppercase tracking-wider mb-2">Executive Summary</p>
            <div className="text-xs text-zinc-300 leading-relaxed overflow-y-auto max-h-40 pr-1 custom-scrollbar prose prose-xs prose-invert prose-zinc">
              <Markdown>{analysis.summary || 'No summary available.'}</Markdown>
            </div>
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

// --- Peer Datasets Constants & Comparison Dashboard ---

const GOOGLE_PEER_META: UploadedFileMeta = {
  name: 'Google-2024-ESG-Peer-Data.json',
  type: 'application/json',
  size: '2.4 KB',
  kindText: 'JSON Peer Data',
  metricsCount: 7,
  metrics: [
    {
      year: '2024',
      metric_name: 'Greenhouse Gas Emissions (Scope 1 & 2)',
      value: '1.20',
      unit: 'million metric tons CO2e',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Greenhouse Gas Emissions (Scope 3)',
      value: '13.10',
      unit: 'million metric tons CO2e',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Carbon-Free Energy (CFE) Share',
      value: '64.0',
      unit: '% global average',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Water Replenishment of Consumption',
      value: '18.0',
      unit: '% of total water consumption',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Women in Leadership Roles',
      value: '31.2',
      unit: '% of global leadership',
      category: 'Social'
    },
    {
      year: '2024',
      metric_name: 'Women in Tech Roles',
      value: '26.2',
      unit: '% of tech workforce',
      category: 'Social'
    },
    {
      year: '2024',
      metric_name: 'Community Investment / Tech Donations',
      value: '3.10',
      unit: 'billion USD',
      category: 'Social'
    }
  ],
  analysis: {
    documentType: 'Report',
    documentSubtype: 'ESG Report',
    companyName: 'Google LLC',
    reportingYears: ['2024'],
    summary: `Google's 2024 ESG performance underscores a **64% global average carbon-free energy** achievement, alongside a **48% increase in GHG emissions** from a 2019 baseline due to data center expansion. Social efforts focus on expanding workforce diversity with **26.2% women in tech** and distributing **$3.1 billion in technology donations** to nonprofits.`,
    confidence_score: 0.95,
    documentStructure: [
      'Environmental Report: Carbon, Water, Circular Economy',
      'Social Impact: Diversity, Community Investment, Digital Skills',
      'Governance & Oversight: Compliance, Privacy, Leadership'
    ],
    reportingEcosystem: {
      frameworks: [
        { name: 'TCFD', fullName: 'Task Force on Climate-related Financial Disclosures', type: 'Framework', shortDescription: 'TCFD disclosure recommendations' },
        { name: 'GRI', fullName: 'Global Reporting Initiative', type: 'Standard', shortDescription: 'GRI standard-compliant sustainability index' }
      ],
      standards: [
        { name: 'SASB', fullName: 'Sustainability Accounting Standards Board', type: 'Standard', shortDescription: 'Software & IT Services Industry Standard' }
      ],
      ratings: [
        { name: 'MSCI', fullName: 'MSCI ESG Rating', type: 'Rating', shortDescription: 'Leader rating based on corporate resilience' }
      ],
      certifications: [
        { name: 'ISO 50001', fullName: 'ISO Energy Management Standard', type: 'Certification', shortDescription: 'Data centers global certification' }
      ],
      assuranceStandards: [
        { name: 'ISAE 3000', fullName: 'International Standard on Assurance Engagements', type: 'Assurance', shortDescription: 'Reasonable and limited assurance on GHG metrics' }
      ]
    },
    dataFormats: {
      tables: true,
      charts: true,
      graphs: false,
      kpiDashboards: true,
      narrativeSections: true
    },
    esgDataLocations: ['Environmental Report Page 12-24', 'Diversity Annual Report Page 5-11'],
    reportEvolution: [
      'Shifted to granular regional CFE metrics tracking',
      'Expanded scope of Supply Chain Scope 3 calculations',
      'Unified reporting across Alphabet subsidiaries'
    ],
    metrics: []
  }
};
GOOGLE_PEER_META.analysis.metrics = GOOGLE_PEER_META.metrics as any;

const APPLE_PEER_META: UploadedFileMeta = {
  name: 'Apple-2024-ESG-Peer-Data.json',
  type: 'application/json',
  size: '2.6 KB',
  kindText: 'JSON Peer Data',
  metricsCount: 7,
  metrics: [
    {
      year: '2024',
      metric_name: 'Greenhouse Gas Emissions (Scope 1 & 2)',
      value: '0.32',
      unit: 'million metric tons CO2e',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Greenhouse Gas Emissions (Scope 3)',
      value: '15.28',
      unit: 'million metric tons CO2e',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Carbon-Free Energy (CFE) Share',
      value: '100.0',
      unit: '% for corporate facilities',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Recycled Materials Content in Products',
      value: '22.0',
      unit: '% of all materials shipped',
      category: 'Environmental'
    },
    {
      year: '2024',
      metric_name: 'Women in Leadership Roles',
      value: '34.0',
      unit: '% of global leadership',
      category: 'Social'
    },
    {
      year: '2024',
      metric_name: 'Women in Tech Roles',
      value: '25.0',
      unit: '% of tech workforce',
      category: 'Social'
    },
    {
      year: '2024',
      metric_name: 'Supplier Clean Energy Program Commitments',
      value: '95.0',
      unit: '% of direct supplier spend',
      category: 'Environmental'
    }
  ],
  analysis: {
    documentType: 'Report',
    documentSubtype: 'Sustainability Report',
    companyName: 'Apple Inc.',
    reportingYears: ['2024'],
    summary: `Apple's 2024 Environmental Progress Report details their path toward **Apple 2030** carbon neutrality. Highlights include a **55% reduction in GHG emissions** since 2015, using **22% recycled/renewable materials** in shipped products, achieving **100% renewable electricity** across corporate offices, and securing clean energy commitments from **95% of direct suppliers**.`,
    confidence_score: 0.98,
    documentStructure: [
      'Resources: Recycled Materials, Water Stewardship, Zero Waste',
      'Climate Change: Low-carbon Design, Energy Efficiency, Renewable Electricity, Carbon Removal',
      'Smarter Chemistry: Chemical Safety, Material Sourcing',
      'Social Responsibility & Inclusion'
    ],
    reportingEcosystem: {
      frameworks: [
        { name: 'TCFD', fullName: 'Task Force on Climate-related Financial Disclosures', type: 'Framework', shortDescription: 'TCFD metrics reporting' }
      ],
      standards: [
        { name: 'GRI', fullName: 'Global Reporting Initiative', type: 'Standard', shortDescription: 'GRI Content Index' },
        { name: 'SASB', fullName: 'Sustainability Accounting Standards Board', type: 'Standard', shortDescription: 'Hardware industry standards compliance' }
      ],
      ratings: [
        { name: 'S&P ESG', fullName: 'S&P Global ESG Score', type: 'Rating', shortDescription: 'Technology Hardware, Storage & Peripherals rating' }
      ],
      certifications: [
        { name: 'Alliance for Water Stewardship (AWS)', fullName: 'AWS Certification', type: 'Certification', shortDescription: 'Corporate and supplier site certificates' }
      ],
      assuranceStandards: [
        { name: 'Apex Assurance', fullName: 'Apex Companies Assurance', type: 'Assurance', shortDescription: 'Independent third-party verification of GHG emissions' }
      ]
    },
    dataFormats: {
      tables: true,
      charts: true,
      graphs: true,
      kpiDashboards: false,
      narrativeSections: true
    },
    esgDataLocations: ['Environmental Progress Report Page 6-35', 'Supplier Responsibility Report Page 10-18'],
    reportEvolution: [
      'Transitioned packaging materials to 100% fiber-based solutions',
      'Accelerated supplier renewable energy tracking metrics',
      'Introduced high-integrity carbon removal project metrics'
    ],
    metrics: []
  }
};
APPLE_PEER_META.analysis.metrics = APPLE_PEER_META.metrics as any;

interface ComparisonDashboardProps {
  files: UploadedFileMeta[];
  selectedIndices: number[];
  onToggleIndex: (idx: number) => void;
  onLoadPeer: (peer: 'Google' | 'Apple') => void;
}

function ComparisonDashboard({ files, selectedIndices, onToggleIndex, onLoadPeer }: ComparisonDashboardProps) {
  const [metricQuery, setMetricQuery] = useState('');
  const [chartMetricType, setChartMetricType] = useState<'emissions' | 'diversity' | 'donations'>('emissions');

  const selectedFiles = files.filter((_, idx) => selectedIndices.includes(idx));

  const getCompanyEmissions = (file: UploadedFileMeta) => {
    const s12 = file.metrics?.find(m => 
      /Scope 1.*2|Scope 1\s*\+\s*2|Scope 1 and 2/i.test(m.metric_name) ||
      (/Scope 1/i.test(m.metric_name) && /Scope 2/i.test(m.metric_name))
    );
    const scope1 = file.metrics?.find(m => /Scope 1/i.test(m.metric_name) && !/Scope 3/i.test(m.metric_name));
    const scope2 = file.metrics?.find(m => /Scope 2/i.test(m.metric_name) && !/Scope 3/i.test(m.metric_name));
    
    let s12Val: number | null = null;
    if (s12) {
      s12Val = parseFloat(s12.value);
    } else if (scope1 || scope2) {
      s12Val = (scope1 ? parseFloat(scope1.value) : 0) + (scope2 ? parseFloat(scope2.value) : 0);
    }

    const s3 = file.metrics?.find(m => /Scope 3/i.test(m.metric_name));
    let s3Val: number | null = s3 ? parseFloat(s3.value) : null;

    return { s12: s12Val, s3: s3Val };
  };

  const getCompanyDiversity = (file: UploadedFileMeta) => {
    const leadership = file.metrics?.find(m => /leadership/i.test(m.metric_name) && /women/i.test(m.metric_name));
    const tech = file.metrics?.find(m => /tech/i.test(m.metric_name) && /women/i.test(m.metric_name));
    return {
      leadership: leadership ? parseFloat(leadership.value) : null,
      tech: tech ? parseFloat(tech.value) : null
    };
  };

  const getCompanyDonations = (file: UploadedFileMeta) => {
    const donation = file.metrics?.find(m => 
      /donat/i.test(m.metric_name) || 
      /technology provided/i.test(m.metric_name) || 
      /nonprofit/i.test(m.metric_name)
    );
    if (!donation) return null;
    let val = parseFloat(donation.value);
    if (/million/i.test(donation.unit) || /million/i.test(donation.value)) {
      val = val / 1000;
    }
    return val;
  };

  const emissionsChartData = selectedFiles.map(file => {
    const data = getCompanyEmissions(file);
    return {
      name: file.analysis.companyName || file.name.split('.')[0],
      'Scope 1 & 2': data.s12 || 0,
      'Scope 3': data.s3 || 0
    };
  });

  const diversityChartData = selectedFiles.map(file => {
    const data = getCompanyDiversity(file);
    return {
      name: file.analysis.companyName || file.name.split('.')[0],
      'Women in Leadership': data.leadership || 0,
      'Women in Tech': data.tech || 0
    };
  });

  const donationsChartData = selectedFiles.map(file => {
    const data = getCompanyDonations(file);
    return {
      name: file.analysis.companyName || file.name.split('.')[0],
      'Donations': data || 0
    };
  });

  const allMetricNames = Array.from(new Set(
    selectedFiles.flatMap(f => (f.metrics || []).map(m => m.metric_name))
  )).filter(name => name.toLowerCase().includes(metricQuery.toLowerCase()));

  const googleLoaded = files.some(f => f.name === GOOGLE_PEER_META.name);
  const appleLoaded = files.some(f => f.name === APPLE_PEER_META.name);

  return (
    <div className="flex-1 overflow-y-auto p-6 space-y-8 custom-scrollbar">
      {(!googleLoaded || !appleLoaded) && (
        <div className="bg-zinc-900/40 border border-zinc-850 p-5 rounded-2xl flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="text-left">
            <h3 className="text-sm font-bold text-zinc-100 mb-1">Add Tech Peers for Comparative Insights</h3>
            <p className="text-xs text-zinc-400">Load pre-extracted authentic 2024 ESG reports for Google or Apple to enable side-by-side comparison.</p>
          </div>
          <div className="flex gap-3 shrink-0">
            {!googleLoaded && (
              <button 
                onClick={() => onLoadPeer('Google')}
                className="px-3.5 py-2 bg-zinc-850 hover:bg-zinc-800 border border-zinc-700 text-zinc-200 text-xs font-semibold rounded-xl transition-all cursor-pointer"
              >
                + Google 2024
              </button>
            )}
            {!appleLoaded && (
              <button 
                onClick={() => onLoadPeer('Apple')}
                className="px-3.5 py-2 bg-zinc-850 hover:bg-zinc-800 border border-zinc-700 text-zinc-200 text-xs font-semibold rounded-xl transition-all cursor-pointer"
              >
                + Apple 2024
              </button>
            )}
          </div>
        </div>
      )}

      <div className="space-y-3">
        <h4 className="text-xs font-bold text-zinc-500 uppercase tracking-widest">Select Reports to Compare</h4>
        <div className="flex flex-wrap gap-3">
          {files.map((file, idx) => {
            const isSelected = selectedIndices.includes(idx);
            return (
              <button
                key={idx}
                onClick={() => onToggleIndex(idx)}
                className={cn(
                  "px-4 py-2.5 rounded-xl border text-xs font-semibold flex items-center gap-2 transition-all cursor-pointer",
                  isSelected 
                    ? "bg-white text-zinc-950 border-white"
                    : "bg-zinc-900/50 text-zinc-400 border-zinc-850 hover:border-zinc-750 hover:text-zinc-200"
                )}
              >
                <div className={cn(
                  "w-3.5 h-3.5 rounded-md border flex items-center justify-center transition-all",
                  isSelected ? "bg-zinc-950 border-zinc-950 text-white" : "border-zinc-700"
                )}>
                  {isSelected && <div className="w-1.5 h-1.5 bg-white rounded-sm" />}
                </div>
                {file.analysis.companyName || file.name}
              </button>
            );
          })}
        </div>
      </div>

      {selectedFiles.length === 0 ? (
        <div className="h-64 flex flex-col items-center justify-center text-center p-8 border border-zinc-850 border-dashed rounded-2xl text-zinc-500 italic">
          Please select at least one report above to compare.
        </div>
      ) : (
        <>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {selectedFiles.map((file, fIdx) => (
              <div key={fIdx} className="bg-zinc-900/40 border border-zinc-850 p-5 rounded-2xl flex flex-col gap-4 text-left shadow-sm">
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <span className="text-[10px] bg-white/10 text-white border border-white/10 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider">
                      {file.analysis.documentSubtype || file.kindText}
                    </span>
                    <span className="text-xs font-mono font-bold text-zinc-500">
                      {file.analysis.reportingYears?.join(', ') || '2024'}
                    </span>
                  </div>
                  <h3 className="text-md font-extrabold text-white tracking-tight">
                    {file.analysis.companyName || file.name}
                  </h3>
                </div>
                <div className="flex-1 text-xs text-zinc-300 leading-relaxed prose prose-xs prose-invert prose-zinc max-h-40 overflow-y-auto custom-scrollbar pr-1">
                  <Markdown>{file.analysis.summary}</Markdown>
                </div>
                <div className="pt-3 border-t border-zinc-800/50 flex items-center justify-between">
                  <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">Confidence</span>
                  <span className={cn(
                    "text-xs font-extrabold px-2 py-0.5 rounded-md border",
                    file.analysis.confidence_score > 0.8 
                      ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/20" 
                      : "bg-amber-500/10 text-amber-400 border-amber-500/20"
                  )}>
                    {Math.round(file.analysis.confidence_score * 100)}%
                  </span>
                </div>
              </div>
            ))}
          </div>

          <div className="bg-zinc-900/40 border border-zinc-850 p-6 rounded-2xl space-y-6">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="text-left">
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <BarChart3 className="w-4 h-4 text-zinc-400" />
                  Key Dimensions Visualized
                </h3>
                <p className="text-[11px] text-zinc-500 mt-0.5">Visual side-by-side comparison of normalized dimensions.</p>
              </div>
              <div className="flex rounded-lg bg-zinc-950 p-1 border border-zinc-850">
                {(['emissions', 'diversity', 'donations'] as const).map(type => (
                  <button 
                    key={type} 
                    onClick={() => setChartMetricType(type)}
                    className={cn(
                      "px-3 py-1 rounded-md text-xs font-semibold transition-all cursor-pointer",
                      chartMetricType === type 
                        ? "bg-zinc-900 text-white border border-zinc-800" 
                        : "text-zinc-400 hover:text-zinc-200"
                    )}
                  >
                    {type === 'emissions' && 'GHG Emissions'}
                    {type === 'diversity' && 'Workforce Diversity'}
                    {type === 'donations' && 'Donations'}
                  </button>
                ))}
              </div>
            </div>

            <div className="h-72 w-full bg-zinc-950/40 border border-zinc-900 p-4 rounded-xl flex items-center justify-center">
              <ResponsiveContainer width="100%" height="100%">
                {chartMetricType === 'emissions' ? (
                  <BarChart data={emissionsChartData}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#27272a" />
                    <XAxis dataKey="name" fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                    <YAxis fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                    <Tooltip 
                      contentStyle={{ backgroundColor: '#18181b', border: '1px solid #27272a', borderRadius: '12px' }}
                    />
                    <Legend wrapperStyle={{ fontSize: 10, paddingTop: 10 }} />
                    <Bar dataKey="Scope 1 & 2" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="Scope 3" fill="#10b981" radius={[4, 4, 0, 0]} />
                  </BarChart>
                ) : chartMetricType === 'diversity' ? (
                  <BarChart data={diversityChartData}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#27272a" />
                    <XAxis dataKey="name" fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                    <YAxis fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} unit="%" />
                    <Tooltip 
                      contentStyle={{ backgroundColor: '#18181b', border: '1px solid #27272a', borderRadius: '12px' }}
                    />
                    <Legend wrapperStyle={{ fontSize: 10, paddingTop: 10 }} />
                    <Bar dataKey="Women in Leadership" fill="#ec4899" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="Women in Tech" fill="#8b5cf6" radius={[4, 4, 0, 0]} />
                  </BarChart>
                ) : (
                  <BarChart data={donationsChartData}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#27272a" />
                    <XAxis dataKey="name" fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                    <YAxis fontSize={10} tickLine={false} axisLine={false} tick={{ fill: '#71717a' }} />
                    <Tooltip 
                      contentStyle={{ backgroundColor: '#18181b', border: '1px solid #27272a', borderRadius: '12px' }}
                    />
                    <Bar dataKey="Donations" fill="#f59e0b" radius={[4, 4, 0, 0]} />
                  </BarChart>
                )}
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-zinc-900/40 border border-zinc-850 p-6 rounded-2xl space-y-4">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="text-left">
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <TableIcon className="w-4 h-4 text-zinc-400" />
                  Granular Metric Matrix
                </h3>
                <p className="text-[11px] text-zinc-500 mt-0.5">Explore how specific disclosures compare across entities.</p>
              </div>
              <input 
                type="text" 
                placeholder="Search comparison metrics..." 
                value={metricQuery}
                onChange={(e) => setMetricQuery(e.target.value)}
                className="bg-zinc-950 border border-zinc-850 rounded-lg px-3 py-1.5 text-xs text-white placeholder:text-zinc-500 focus:outline-none focus:border-zinc-700 w-full md:w-64"
              />
            </div>

            <div className="overflow-x-auto border border-zinc-850 rounded-xl bg-zinc-950">
              <table className="w-full text-xs text-left">
                <thead className="bg-zinc-900 border-b border-zinc-850 text-zinc-500 font-bold uppercase tracking-wider">
                  <tr>
                    <th className="px-4 py-3 min-w-[200px]">Metric Disclosed</th>
                    {selectedFiles.map((file, idx) => (
                      <th key={idx} className="px-4 py-3 min-w-[150px]">{file.analysis.companyName || file.name}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-850">
                  {allMetricNames.length > 0 ? (
                    allMetricNames.map((metricName, mIdx) => (
                      <tr key={mIdx} className="hover:bg-zinc-900/10 transition-colors">
                        <td className="px-4 py-3.5 font-semibold text-zinc-300 leading-normal">{metricName}</td>
                        {selectedFiles.map((file, fIdx) => {
                          const m = file.metrics?.find(met => met.metric_name.toLowerCase() === metricName.toLowerCase());
                          if (m) {
                            const isEnv = m.category?.toLowerCase() === 'environmental';
                            const isSoc = m.category?.toLowerCase() === 'social';
                            const badgeColor = isEnv 
                              ? 'text-emerald-400' 
                              : isSoc 
                                ? 'text-blue-400' 
                                : 'text-amber-400';
                            return (
                              <td key={fIdx} className="px-4 py-3.5">
                                <div className="font-bold text-zinc-100">{m.value} <span className="font-normal text-zinc-550 text-[10px] ml-0.5">{m.unit}</span></div>
                                <span className={cn("text-[9px] uppercase tracking-wider font-semibold font-mono", badgeColor)}>
                                  {m.category} ({m.year})
                                </span>
                              </td>
                            );
                          }
                          return (
                            <td key={fIdx} className="px-4 py-3.5 text-zinc-700 italic">
                              -
                            </td>
                          );
                        })}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={selectedFiles.length + 1} className="px-4 py-8 text-center text-zinc-500 italic">
                        No comparison metrics found matching filter.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

export default function App() {
  const [isIngesting, setIsIngesting] = useState(false);
  const [retryStatus, setRetryStatus] = useState<string | null>(null);
  const [dataset, setDataset] = useState<ESGDataset | null>(null);
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFileMeta[]>([]);
  const [selectedFileIndex, setSelectedFileIndex] = useState<number | null>(null);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'chat' | 'comparison'>('dashboard');
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState('');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const [comparedFileIndices, setComparedFileIndices] = useState<number[]>([]);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [messages]);

  useEffect(() => {
    if (uploadedFiles.length > 0) {
      setComparedFileIndices(prev => {
        const next = [...prev];
        uploadedFiles.forEach((_, idx) => {
          if (!next.includes(idx)) {
            next.push(idx);
          }
        });
        return next.filter(idx => idx < uploadedFiles.length);
      });
    } else {
      setComparedFileIndices([]);
    }
  }, [uploadedFiles]);

  const toggleComparedFile = (idx: number) => {
    setComparedFileIndices(prev => {
      if (prev.includes(idx)) {
        if (prev.length === 1) return prev;
        return prev.filter(i => i !== idx);
      } else {
        return [...prev, idx];
      }
    });
  };

  const loadPeerDataset = (peer: 'Google' | 'Apple') => {
    setError(null);
    const peerMeta = peer === 'Google' ? GOOGLE_PEER_META : APPLE_PEER_META;
    
    if (uploadedFiles.some(f => f.name === peerMeta.name)) {
      return;
    }

    const newUploadedFiles = [...uploadedFiles, peerMeta];
    setUploadedFiles(newUploadedFiles);
    
    let mergedDataset: ESGDataset = dataset 
      ? { ...dataset, metrics: [...dataset.metrics] } 
      : { companyName: peerMeta.analysis.companyName || '', summary: '', metrics: [] };

    if (!mergedDataset.companyName) {
      mergedDataset.companyName = peerMeta.analysis.companyName || '';
    }
    
    mergedDataset.summary += (mergedDataset.summary ? " | " : "") + peerMeta.analysis.summary;
    
    if (peerMeta.metrics) {
      const mapped = peerMeta.metrics.map(m => ({
        year: m.year,
        metric_name: m.metric_name,
        value: m.value,
        unit: m.unit,
        category: m.category
      }));
      mergedDataset.metrics = [...mergedDataset.metrics, ...mapped];
    }

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
    setSelectedFileIndex(newUploadedFiles.length - 1);
    setActiveTab('dashboard');

    const newMsg: ChatMessage = {
      role: 'assistant',
      content: `Loaded **${peerMeta.name}** peer dataset. You can now compare it with other reports in the **Peer Comparison** tab or ask questions in **Interactive Chat**.`
    };
    setMessages(prev => [...prev, newMsg]);
  };

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
              },
              suggestedPeer: {
                type: Type.OBJECT,
                properties: {
                  companyName: { type: Type.STRING },
                  summary: { type: Type.STRING },
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
                required: ["companyName", "summary", "metrics"]
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
              "metrics",
              "suggestedPeer"
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
3. Extract: companyName, reportingYears, and summary (a detailed executive summary of 100-200 words in markdown format, describing the report scope, main achievements, and any challenges or sustainability milestones).
4. Generate confidence_score between 0.0 and 1.0 based only on document evidence.
5. Extract document structure. Return all major sections and subsections when available.
6. Extract reporting frameworks, standards, certifications, ratings, and assurance references explicitly mentioned in the document. Organise them under reportingEcosystem.
7. Extract all key metrics and indicators explicitly stated in the document (such as ESG parameters, blood test values, invoice amounts/costs, resume experience duration, academic grades, bank transactions, etc.). Organise them under metrics.
For each metric, determine year, metric_name, value, unit, and category. If standard ESG categories (Environmental, Social, Governance) do not apply, use appropriate domain categories (e.g., Financial, Hematology, Experience, Education, Personal, etc.).
8. Identify ESG data locations within the document.
9. Detect whether the document contains: Tables, Charts, Graphs, KPI Dashboards, Narrative Sections.
10. If multiple reporting years are present, identify reporting evolution signals under reportEvolution.
11. Generate a relatable suggestedPeer dataset representing standard peer benchmarks, industry averages, normal reference ranges, or a comparable competitor baseline corresponding to this document's content and metrics. It must have matching metric names and units so they can be compared directly side-by-side. Make sure to generate realistic values (e.g. standard healthy reference values for a blood test, comparable engineer stats for a resume, industry standard averages for an invoice).

Use the specified JSON schema structure.`;

          const response = await callWithRetry(() => ai.models.generateContent({
            model: "gemini-2.5-flash",
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

        // Generate virtual peer dataset if returned by Gemini
        if (fileResult.suggestedPeer) {
          const peer = fileResult.suggestedPeer;
          const peerName = `Peer-Baseline-For-${fileName}`;
          const virtualPeerMeta: UploadedFileMeta = {
            name: peerName,
            type: 'application/json',
            size: '1.5 KB',
            kindText: 'Auto-Generated Peer Baseline',
            metricsCount: peer.metrics?.length || 0,
            metrics: peer.metrics ? peer.metrics.map(m => ({
              year: m.year,
              metric_name: m.metric_name,
              value: m.value,
              unit: m.unit,
              category: m.category
            })) : [],
            analysis: {
              documentType: fileResult.documentType,
              documentSubtype: 'Peer Baseline',
              companyName: peer.companyName,
              reportingYears: fileResult.reportingYears,
              summary: peer.summary,
              confidence_score: 1.0,
              documentStructure: ['Comparative Baseline'],
              reportingEcosystem: { frameworks: [], standards: [], ratings: [], certifications: [], assuranceStandards: [] },
              dataFormats: { tables: true, charts: false, graphs: false, kpiDashboards: false, narrativeSections: false },
              esgDataLocations: [],
              reportEvolution: [],
              metrics: peer.metrics || []
            }
          };

          if (peer.metrics) {
            const mapped = peer.metrics.map(m => ({
              year: m.year,
              metric_name: m.metric_name,
              value: m.value,
              unit: m.unit,
              category: m.category
            }));
            mergedDataset.metrics = [...mergedDataset.metrics, ...mapped];
          }

          newUploadedFiles.push(virtualPeerMeta);
        }
      }

      const prevLen = uploadedFiles.length;
      setUploadedFiles(prev => [...prev, ...newUploadedFiles]);
      
      const totalCount = prevLen + newUploadedFiles.length;
      if (totalCount > 1) {
        setActiveTab('comparison');
      } else {
        if (selectedFileIndex === null && newUploadedFiles.length > 0) {
          setSelectedFileIndex(prevLen);
        }
        setActiveTab('dashboard');
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
      setError("Failed to process the documents. " + (err.message || "Please check that they are valid files."));
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
        model: "gemini-2.5-flash",
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
    <div className="min-h-screen bg-[#050507] text-zinc-100 font-sans selection:bg-zinc-800 relative overflow-hidden">
      {/* Background Radial Glow */}
      <div className="absolute top-0 left-1/4 w-[500px] h-[500px] bg-indigo-500/5 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute bottom-0 right-1/4 w-[600px] h-[600px] bg-purple-500/5 rounded-full blur-[140px] pointer-events-none" />

      {/* Header */}
      <header className="sticky top-0 z-50 bg-[#050507]/40 backdrop-blur-xl border-b border-zinc-900/60 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-gradient-to-tr from-indigo-500 via-purple-500 to-pink-500 rounded-xl flex items-center justify-center shadow-lg shadow-purple-500/10">
            <BarChart3 className="text-white w-5 h-5" />
          </div>
          <div>
            <h1 className="text-md font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-white via-zinc-200 to-zinc-400">ESG Intel Agent</h1>
            <p className="text-[9px] uppercase tracking-[0.2em] font-bold text-zinc-500">Strict Data Mode Active</p>
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
              onClick={() => { setDataset(null); setMessages([]); setUploadedFiles([]); setComparedFileIndices([]); setActiveTab('dashboard'); }}
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
            <div className="h-full flex flex-col justify-between border border-zinc-850 rounded-3xl bg-zinc-900/10 p-6 text-center">
              <div className="flex-1 flex flex-col items-center justify-center py-6">
                <div className="w-14 h-14 bg-zinc-900 border border-zinc-800 rounded-2xl flex items-center justify-center mb-5">
                  <Upload className="w-6 h-6 text-zinc-500" />
                </div>
                <h2 className="text-lg font-bold mb-1 text-white">Upload Documents</h2>
                <p className="text-xs text-zinc-500 mb-6 max-w-xs">
                  Upload any report, spreadsheet, text, or PDF file to begin your ESG analysis.
                </p>
                <label className="relative group cursor-pointer">
                  <div className={cn(
                    "px-6 py-2.5 bg-white text-zinc-900 rounded-xl text-xs font-semibold transition-all group-hover:scale-105 active:scale-95 flex items-center gap-2",
                    isIngesting && "opacity-50 pointer-events-none"
                  )}>
                    {isIngesting ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <FileText className="w-3.5 h-3.5" />}
                    {isIngesting ? (retryStatus || "Processing...") : "Select Documents"}
                  </div>
                  <input type="file" className="hidden" onChange={handleFileUpload} disabled={isIngesting} multiple />
                </label>
                {error && (
                  <div className="mt-4 p-3 bg-rose-500/10 border border-rose-500/20 rounded-xl flex items-start gap-2 text-left">
                    <AlertCircle className="w-3.5 h-3.5 text-rose-400 shrink-0 mt-0.5" />
                    <p className="text-[11px] text-rose-400 leading-normal">{error}</p>
                  </div>
                )}
              </div>
              
              <div className="pt-6 border-t border-zinc-800/40 text-left space-y-3">
                <div>
                  <h3 className="text-xs font-bold text-zinc-400 uppercase tracking-wider">Try with Peer Data</h3>
                  <p className="text-[10px] text-zinc-500 leading-normal mt-0.5">Explore the dashboard instantly using authentic 2024 ESG peer datasets.</p>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    onClick={() => loadPeerDataset('Google')}
                    className="px-3 py-2.5 rounded-xl border border-zinc-850 bg-zinc-900/40 text-zinc-300 hover:border-zinc-700 hover:bg-zinc-900/70 text-xs font-semibold transition-all cursor-pointer text-center"
                  >
                    Google 2024
                  </button>
                  <button
                    onClick={() => loadPeerDataset('Apple')}
                    className="px-3 py-2.5 rounded-xl border border-zinc-850 bg-zinc-900/40 text-zinc-300 hover:border-zinc-700 hover:bg-zinc-900/70 text-xs font-semibold transition-all cursor-pointer text-center"
                  >
                    Apple 2024
                  </button>
                </div>
              </div>
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
                          "cursor-pointer bg-zinc-900/10 hover:bg-zinc-900/40 border rounded-2xl flex flex-col gap-4 p-5 transition-all text-left shadow-sm hover:shadow-md",
                          selectedFileIndex === idx 
                            ? "border-zinc-400 bg-zinc-900/60 ring-1 ring-zinc-400/20" 
                            : "border-zinc-800/80 hover:border-zinc-700/80"
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
                          <div className="pt-3.5 border-t border-zinc-800/60 space-y-2.5">
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
                                  <div key={mIdx} className="pt-2.5 pb-2.5 first:pt-1 last:pb-1 flex flex-col gap-1 text-xs">
                                    <div className="flex items-start justify-between gap-2">
                                      <span className="font-semibold text-zinc-300 leading-tight">
                                        {m.metric_name}
                                      </span>
                                      <span className={cn("text-[8px] font-mono font-semibold shrink-0 px-1.5 py-0.5 rounded-md border", badgeColor)}>
                                        {m.year}
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between mt-0.5 text-[11px]">
                                      <span className="text-zinc-500 text-[10px]">{m.category}</span>
                                      <span className="font-bold text-zinc-200 font-mono">
                                        {m.value} <span className="text-zinc-500 font-normal text-[10px] ml-0.5">{m.unit}</span>
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
                <div className="bg-transparent border border-zinc-800/80 p-5 rounded-2xl shadow-sm leading-relaxed text-xs text-zinc-300 prose prose-xs prose-invert prose-zinc max-h-60 overflow-y-auto custom-scrollbar">
                  <Markdown>{dataset.summary || 'No summary available.'}</Markdown>
                </div>
              </section>

              {/* Suggested Peer Datasets (Inside Ingested state) */}
              <section className="space-y-4 pt-4 border-t border-zinc-800/40">
                <h3 className="text-xs font-bold text-zinc-500 uppercase tracking-widest flex items-center gap-2">
                  <TrendingUp className="w-3.5 h-3.5 text-zinc-400" />
                  Suggested Peer Datasets
                </h3>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    disabled={uploadedFiles.some(f => f.name === GOOGLE_PEER_META.name)}
                    onClick={() => loadPeerDataset('Google')}
                    className={cn(
                      "px-3 py-3 rounded-xl border text-xs font-semibold flex flex-col items-center gap-1.5 transition-all text-center",
                      uploadedFiles.some(f => f.name === GOOGLE_PEER_META.name)
                        ? "border-zinc-800/40 bg-zinc-900/10 text-zinc-650 cursor-not-allowed"
                        : "border-zinc-800 bg-zinc-900/30 text-zinc-300 hover:border-zinc-700 hover:bg-zinc-900/50 cursor-pointer"
                    )}
                  >
                    <span className="font-extrabold text-[13px] text-zinc-100">Google</span>
                    <span className="text-[9px] text-zinc-500">2024 ESG Data</span>
                  </button>
                  <button
                    disabled={uploadedFiles.some(f => f.name === APPLE_PEER_META.name)}
                    onClick={() => loadPeerDataset('Apple')}
                    className={cn(
                      "px-3 py-3 rounded-xl border text-xs font-semibold flex flex-col items-center gap-1.5 transition-all text-center",
                      uploadedFiles.some(f => f.name === APPLE_PEER_META.name)
                        ? "border-zinc-800/40 bg-zinc-900/10 text-zinc-650 cursor-not-allowed"
                        : "border-zinc-800 bg-zinc-900/30 text-zinc-300 hover:border-zinc-700 hover:bg-zinc-900/50 cursor-pointer"
                    )}
                  >
                    <span className="font-extrabold text-[13px] text-zinc-100">Apple</span>
                    <span className="text-[9px] text-zinc-500">2024 ESG Data</span>
                  </button>
                </div>
              </section>
            </>
          )}
        </div>

        {/* Middle Content: Chat & Analysis */}
        <div className="lg:col-span-8 flex flex-col bg-[#0a0a0c]/20 backdrop-blur-xl border border-zinc-900/60 rounded-3xl shadow-[0_8px_30px_rgb(0,0,0,0.2)] overflow-hidden">
          {/* Tabs Navigation Header */}
          <div className="flex items-center justify-between border-b border-zinc-900/60 px-6 py-3 bg-[#050507]/40 backdrop-blur-md">
            <div className="flex gap-4">
              <button 
                onClick={() => setActiveTab('dashboard')}
                className={cn(
                  "px-4 py-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all cursor-pointer",
                  activeTab === 'dashboard' 
                    ? "border-white text-white" 
                    : "border-transparent text-zinc-500 hover:text-zinc-350"
                )}
              >
                Analysis Dashboard
              </button>
              <button 
                onClick={() => setActiveTab('chat')}
                className={cn(
                  "px-4 py-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all cursor-pointer",
                  activeTab === 'chat' 
                    ? "border-white text-white" 
                    : "border-transparent text-zinc-500 hover:text-zinc-350"
                )}
              >
                Interactive Chat
              </button>
              {uploadedFiles.length > 0 && (
                <button 
                  onClick={() => setActiveTab('comparison')}
                  className={cn(
                    "px-4 py-2 text-xs font-bold uppercase tracking-wider border-b-2 transition-all cursor-pointer",
                    activeTab === 'comparison' 
                      ? "border-white text-white" 
                      : "border-transparent text-zinc-500 hover:text-zinc-350"
                  )}
                >
                  Peer Comparison
                </button>
              )}
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
          ) : activeTab === 'comparison' ? (
            <ComparisonDashboard 
              files={uploadedFiles} 
              selectedIndices={comparedFileIndices} 
              onToggleIndex={toggleComparedFile} 
              onLoadPeer={loadPeerDataset}
            />
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
                          ? "bg-gradient-to-tr from-indigo-600 to-indigo-500 text-white rounded-2xl rounded-tr-none shadow-lg shadow-indigo-600/10 font-medium" 
                          : "bg-zinc-900/60 backdrop-blur-md text-zinc-100 border border-zinc-850 rounded-2xl rounded-tl-none shadow-sm"
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
              <div className="p-6 bg-[#050507]/65 backdrop-blur-md border-t border-zinc-900/60">
                <div className="relative group">
                  <input 
                    type="text"
                    value={input}
                    onChange={(e) => setInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleSendMessage()}
                    placeholder={dataset ? "Ask about Scope 1 emissions, diversity ratios, or trends..." : "Upload a document to start analysis"}
                    disabled={!dataset || isAnalyzing}
                    className="w-full bg-zinc-900/65 border border-zinc-850 focus:border-zinc-700 rounded-2xl px-5 py-4 pr-14 text-sm text-white placeholder:text-zinc-550 focus:outline-none focus:ring-1 focus:ring-zinc-700/50 transition-all disabled:opacity-50"
                  />
                  <button 
                    onClick={handleSendMessage}
                    disabled={!dataset || isAnalyzing || !input.trim()}
                    className="absolute right-3 top-1/2 -translate-y-1/2 w-10 h-10 bg-gradient-to-tr from-indigo-500 to-indigo-600 hover:from-indigo-650 hover:to-indigo-550 text-white rounded-xl flex items-center justify-center hover:scale-105 active:scale-95 transition-all disabled:from-zinc-800 disabled:to-zinc-850 disabled:text-zinc-500 disabled:hover:scale-100 disabled:cursor-not-allowed shadow-md"
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
