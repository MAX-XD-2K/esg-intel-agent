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
  year: number;
  metric: string;
  value: number;
  unit: string;
  category: 'Environmental' | 'Social' | 'Governance';
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

export default function App() {
  const [isIngesting, setIsIngesting] = useState(false);
  const [retryStatus, setRetryStatus] = useState<string | null>(null);
  const [dataset, setDataset] = useState<ESGDataset | null>(null);
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFileMeta[]>([]);
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
        fileResult: ESGDataset;
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
          if (file.type === 'application/pdf' || file.type.startsWith('image/')) {
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
          } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const csv = XLSX.utils.sheet_to_csv(worksheet);
            part = { text: `Excel Data from ${file.name} (Sheet: ${firstSheetName}):\n${csv}` };
          } else if (file.name.endsWith('.docx')) {
            const buffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer: buffer });
            part = { text: `Word Document Data from ${file.name}:\n${result.value}` };
          } else {
            const text = await file.text();
            part = { text: `File Content from ${file.name}:\n${text}` };
          }

          const response = await callWithRetry(() => ai.models.generateContent({
            model: "gemini-3.5-flash",
            contents: [
              {
                parts: [
                  part,
                  { text: "Extract all ESG metrics from this data into a structured JSON format. Include company name, a brief summary, and a list of metrics with year, metric name, value, unit, and category (Environmental, Social, or Governance). Ensure all numerical values are extracted accurately." }
                ]
              }
            ],
            config: {
              responseMimeType: "application/json",
              responseSchema: {
                type: Type.OBJECT,
                properties: {
                  companyName: { type: Type.STRING },
                  summary: { type: Type.STRING },
                  metrics: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        year: { type: Type.INTEGER },
                        metric: { type: Type.STRING },
                        value: { type: Type.NUMBER },
                        unit: { type: Type.STRING },
                        category: { type: Type.STRING, enum: ["Environmental", "Social", "Governance"] }
                      },
                      required: ["year", "metric", "value", "unit", "category"]
                    }
                  }
                },
                required: ["companyName", "summary", "metrics"]
              }
            }
          }), (attempt, delay) => {
            setRetryStatus(`Rate limit hit. Resuming in ${Math.round(delay/1000)}s...`);
          });

          const fileResult = JSON.parse(response.text || '{}') as ESGDataset;
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
        if (!mergedDataset.companyName) mergedDataset.companyName = fileResult.companyName;
        mergedDataset.summary += (mergedDataset.summary ? " | " : "") + fileResult.summary;
        if (fileResult.metrics) {
          mergedDataset.metrics = [...mergedDataset.metrics, ...fileResult.metrics];
        }

        newUploadedFiles.push({
          name: fileName,
          type: fileType,
          size: formatBytes(fileSize),
          kindText: getFileKindText(fileType, fileName),
          metricsCount: fileResult.metrics?.length || 0,
          metrics: fileResult.metrics || []
        });
      }

      setUploadedFiles(prev => [...prev, ...newUploadedFiles]);

      // Deduplicate metrics if necessary (same year, same metric name)
      const uniqueMetrics = mergedDataset.metrics.reduce((acc: ESGMetric[], current) => {
        const x = acc.find(item => item.year === current.year && item.metric === current.metric);
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
          : `Successfully ingested data from **${files.length}** file(s) for **${mergedDataset.companyName}**. I have extracted ${mergedDataset.metrics.length} unique metrics across Environmental, Social, and Governance categories. How can I help you analyze this data today?`
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
              <h2 className="text-xl font-bold mb-2 text-white">Upload ESG Reports</h2>
              <p className="text-sm text-zinc-500 mb-8 max-w-xs">
                Upload PDFs, Word, Excel, CSV, text, or any other sustainability reports to begin your analysis.
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
                      <div key={idx} className="bg-transparent border border-zinc-800/80 p-4 rounded-2xl flex flex-col gap-3 hover:bg-zinc-900/30 hover:border-zinc-700/80 transition-all">
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
                                Holds {file.metricsCount} ESG metrics
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
                              Extracted ESG Metrics
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
                                        {m.metric}
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
        </div>

      </main>
    </div>
  );
}
