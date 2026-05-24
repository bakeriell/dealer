
import React, { useEffect, useState, useMemo, useCallback } from 'react';
// FIX: Added BrainCircuit icon to the import from lucide-react to resolve an undefined variable error.
import { LayoutDashboard, Table as TableIcon, Upload, FileSpreadsheet, Users, RefreshCw, AlertCircle, Check, X, DownloadCloud, Filter, TrendingUp, DollarSign, Award, MessageSquare, MapPin, Euro, AlertTriangle, ShoppingCart, Globe, ChevronDown, ChevronUp, PieChart as PieIcon, BarChart3, Tag, Car, Download, Phone, Megaphone, BrainCircuit } from 'lucide-react';
import { DealerData, ColumnDefinition, AppView, LeadFilterType } from './types';
import { generateMockData } from './services/mockData';
import { PivotTable } from './components/PivotTable';
import { read, utils, writeFile } from 'xlsx';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, 
  PieChart, Pie, Cell, Legend, Label, LabelList
} from 'recharts';

// Specific colors from the design
const CHART_COLORS = {
  positive: '#16a34a', // green-600
  negative: '#dc2626', // red-600
  neutral: '#94a3b8',  // slate-400
  info: '#3b82f6',     // blue-500
  warning: '#f59e0b',  // amber-500
  purple: '#8b5cf6',   // violet-500
  teal: '#0d9488',     // teal-600
  models: ['#ef4444', '#374151', '#9ca3af', '#16a34a', '#f59e0b', '#3b82f6'],
  // Specific colors for the 4 purchase answers
  purchase: {
    byd: '#16a34a',      // Green - Si, BYD
    indeciso: '#f59e0b', // Amber - Indeciso
    rimandato: '#3b82f6',// Blue - Rimandato
    competitor: '#ef4444'// Red - Altro Brand
  },
  // Colors for the new Donut Charts
  funnel: {
    si: '#4f46e5',    // Indigo (Blueish)
    no: '#db2777',    // Pink/Magenta
    other: '#0d9488', // Teal
    grey: '#94a3b8'
  },
  // Source Analysis Breakdown
  source: {
    si: '#0d9488',    // Teal (Yes)
    no: '#ef4444',    // Red (No)
    dontRemember: '#64748b', // Slate (Don't Remember)
    noData: '#e2e8f0' // Light Grey (No Survey Data / Unmanaged)
  }
};

// The provided SharePoint link modified for direct download if possible
const SHAREPOINT_URL = "https://byd2-my.sharepoint.com/personal/simona_divita_byd_com/_layouts/15/download.aspx?share=IQCs4PjI9hYYQqriEUyATqLUAW40eMTFsA9hpgq87jHDrWc";

// Helper to check for positive responses (English & Italian)
const isPositiveResponse = (val: any): boolean => {
  if (!val) return false;
  const s = String(val).trim().toLowerCase();
  return ['yes', 'y', 'true', 'si', 'sì', '1', 'pass', 'ok', 'certainly'].some(t => s.includes(t) && !s.includes('non') && !s.includes('no,'));
};

// Helper for Excel Download
const downloadExcel = (data: any[], filename: string) => {
  if (!data || data.length === 0) return;

  const worksheet = utils.json_to_sheet(data);

  // Auto-calculate column widths based on the content
  const maxWidths = Object.keys(data[0]).map(key => {
    // Header length
    let max = key.length;
    // Data length (check up to 500 rows to save perf if large dataset)
    const checkRows = data.slice(0, 500); 
    checkRows.forEach(row => {
      const val = row[key];
      const len = val ? String(val).length : 0;
      if (len > max) max = len;
    });
    // Cap width at 50 chars to prevent massive columns
    return Math.min(max + 2, 50); 
  });

  worksheet['!cols'] = maxWidths.map(w => ({ wch: w }));

  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, worksheet, "Report");
  writeFile(workbook, `${filename}.xlsx`);
};

export default function App() {
  const [view, setView] = useState<AppView>(AppView.DASHBOARD);
  const [data, setData] = useState<DealerData[]>([]);
  const [assignedLeadsMap, setAssignedLeadsMap] = useState<Record<string, number>>({});
  const [columns, setColumns] = useState<ColumnDefinition[]>([]);
  const [loading, setLoading] = useState(true);
  const [fileName, setFileName] = useState<string>("Connecting to SharePoint...");
  const [error, setError] = useState<string | null>(null);
  
  // Filters
  const [selectedDealer, setSelectedDealer] = useState<string>('All Dealers');
  const [selectedModel, setSelectedModel] = useState<string>('All Models');
  const [selectedSource, setSelectedSource] = useState<string>('All Sources');
  const [leadFilter, setLeadFilter] = useState<LeadFilterType>('CONTACTED'); 

  // State for sorting the Malus Table
  const [malusSort, setMalusSort] = useState<{ key: string; dir: 'asc' | 'desc' }>({ key: 'totalMalus', dir: 'desc' });

  // Key mappings to normalize data access
  const [keyMap, setKeyMap] = useState({
    dealer: 'Dealer Name',
    source: 'Source',
    rememberRequest: 'Remember Request',
    userResponded: 'User Responded',
    responded: 'Response Question',
    visited: 'Visited Dealership',
    model: 'Model',
    quoteGiven: 'Quote Given',
    competitors: 'Other Brands',
    consistency: 'Quote Consistency',
    finalized: 'Finalized Purchase'
  });

  // Find the best matching key in the dataset for a given concept
  const findBestKey = (keys: string[], keywords: string[]): string => {
    const lowerKeys = keys.map(k => k.toLowerCase());
    for (const k of keys) {
      const lowerK = k.toLowerCase();
      if (keywords.every(word => lowerK.includes(word.toLowerCase()))) {
        return k;
      }
    }
    return '';
  };

  const processData = useCallback((buffer: ArrayBuffer, name: string) => {
    try {
      const workbook = read(buffer, { type: 'array', cellDates: true });
      
      // 1. Process Main Sheet (Data)
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = utils.sheet_to_json(worksheet, { defval: "" });
      
      if (jsonData.length === 0) {
         setLoading(false);
         return;
      }

      // 2. Process Second Sheet (Assigned Leads) if available
      const newAssignedMap: Record<string, number> = {};
      if (workbook.SheetNames.length > 1) {
         const secondSheet = workbook.Sheets[workbook.SheetNames[1]];
         // Assume columns like "Dealer" and "Total Leads"
         const secondData = utils.sheet_to_json<any>(secondSheet);
         secondData.forEach(row => {
            // Try to find the dealer and number columns vaguely
            const vals = Object.values(row);
            const dealerName = vals.find(v => typeof v === 'string' && v.length > 2) as string;
            const num = vals.find(v => typeof v === 'number') as number;
            if (dealerName && num !== undefined) {
               newAssignedMap[dealerName.trim()] = num;
            }
         });
         console.log("Loaded Assigned Leads from Sheet 2:", newAssignedMap);
      }
      setAssignedLeadsMap(newAssignedMap);

      const sampleKeys = Object.keys(jsonData[0]);
      const getColumnByIndex = (idx: number) => (sampleKeys.length > idx ? sampleKeys[idx] : '');

      // Enhanced mapping logic
      let sourceKey = findBestKey(sampleKeys, ['Fonte', 'Source', 'Channel']);
      if (!sourceKey && sampleKeys.length > 0) {
         // Fallback: Check the very last column if it looks like a Source (usually last added)
         // or if user specifically added it at the end.
         sourceKey = sampleKeys[sampleKeys.length - 1];
      }

      const newKeyMap = {
        dealer: findBestKey(sampleKeys, ['Dealer']) || 'Dealer Name', 
        model: findBestKey(sampleKeys, ['Model', 'Modello']) || 'Model',
        source: sourceKey || 'Fonte',
        
        // Q: Did user respond? (The funnel entry)
        userResponded: findBestKey(sampleKeys, ['utente', 'risposto']) || findBestKey(sampleKeys, ['User', 'Responded']) || '',

        // Q1: Remember Request? (Usually before contact)
        rememberRequest: findBestKey(sampleKeys, ['effettuato', 'richiesta']) || findBestKey(sampleKeys, ['Remember', 'Request']) || getColumnByIndex(6),

        // Q2: Contacted? (Column N - Index 13) - The quality check question
        responded: findBestKey(sampleKeys, ['Dopo', 'lasciato', 'contattato']) || findBestKey(sampleKeys, ['Responded', 'Risposto']) || getColumnByIndex(13),
        
        // Q3: Visited? (Column K - Index 10)
        visited: findBestKey(sampleKeys, ['Visited', 'Visitato', 'recato']) || getColumnByIndex(10),
        
        quoteGiven: findBestKey(sampleKeys, ['preventivo', 'fornito']) || findBestKey(sampleKeys, ['Quote']) || '',
        competitors: findBestKey(sampleKeys, ['valutando', 'altri']) || findBestKey(sampleKeys, ['Other', 'Brand']) || '',
        consistency: findBestKey(sampleKeys, ['preventivo', 'coerente']) || findBestKey(sampleKeys, ['Consistent']) || '',
        finalized: findBestKey(sampleKeys, ['finalizzato', 'acquisto']) || findBestKey(sampleKeys, ['Finalized', 'Purchase']) || ''
      };
      setKeyMap(newKeyMap);

      const processedData: DealerData[] = (jsonData as any[]).map((row, index) => {
        let dealerName = row[newKeyMap.dealer];
        // Hard fallback to Column H (index 7) if key mapping failed or returned empty
        if ((!dealerName || dealerName === '') && Object.keys(row).length > 7) {
             const keys = Object.keys(row);
             dealerName = row[keys[7]]; // Index 7 is H
        }
        dealerName = dealerName || 'Unknown Dealer';

        const cleanRow: any = { ...row, id: row.id || `ROW-${index}` };
        cleanRow['Dealer Name'] = dealerName; 
        return cleanRow;
      });

      const firstRow = processedData[0];
      const newColumns: ColumnDefinition[] = Object.keys(firstRow)
        .filter(k => k !== 'id' && k !== 'Dealer Name')
        .map(key => ({
        key,
        label: key,
        type: typeof firstRow[key] === 'number' ? 'number' : 'text'
      }));
      newColumns.unshift({ key: 'Dealer Name', label: 'Dealer Name', type: 'text' });

      setData(processedData);
      setColumns(newColumns);
      setFileName(name);
      setLoading(false);
      setError(null);
      setSelectedDealer('All Dealers');
      setSelectedModel('All Models');
      setSelectedSource('All Sources');
      setLeadFilter('CONTACTED'); 

    } catch (err) {
      console.error("Parse error:", err);
      alert("Failed to parse file.");
      setLoading(false);
    }
  }, []);

  const loadFromSharePoint = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      
      const timestamp = new Date().getTime();
      const response = await fetch(`${SHAREPOINT_URL}&t=${timestamp}`, {
        cache: 'no-store',
        headers: {
          'Pragma': 'no-cache',
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Expires': '0'
        }
      });
      
      if (!response.ok) {
         throw new Error(`Server returned ${response.status} ${response.statusText}`);
      }

      const arrayBuffer = await response.arrayBuffer();
      processData(arrayBuffer, "SharePoint Data (CSV)");
    } catch (err) {
      console.warn("Direct fetch failed (likely CORS), falling back to empty state.", err);
      setFileName("No Data Loaded");
      setLoading(false);
      setError("Could not auto-fetch file (Browser Security/CORS). Please upload the downloaded CSV manually.");
    }
  }, [processData]);

  useEffect(() => {
    loadFromSharePoint();
  }, [loadFromSharePoint]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);
    setError(null);

    try {
      const arrayBuffer = await file.arrayBuffer();
      processData(arrayBuffer, file.name);
    } catch (err) {
      console.error("Upload error:", err);
      setLoading(false);
      setError("Failed to read file.");
    }
  };

  const loadMockData = () => {
    setLoading(true);
    setTimeout(() => {
      const { data: mockData, columns: mockCols, assignedMap } = generateMockData(30);
      setData(mockData);
      setColumns(mockCols);
      setAssignedLeadsMap(assignedMap);
      setFileName("Mock Data Source");
      setLoading(false);
      setError(null);
      setSelectedDealer('All Dealers');
      setSelectedModel('All Models');
      setSelectedSource('All Sources');
      setLeadFilter('CONTACTED');
      
      // Mock key map
      setKeyMap({
        dealer: 'Dealer Name',
        model: 'Model',
        source: 'Fonte',
        userResponded: 'L’utente ha risposto',
        rememberRequest: 'Ha effettuato una richiesta per un test drive?',
        responded: 'Dopo aver lasciato i suoi dati, è stato contattato da un concessionario BYD?',
        visited: 'Visited Dealership',
        quoteGiven: 'Il concessionario le ha fornito un preventivo per il modello di suo interesse?',
        competitors: 'Sta valutando altri Brand/Modelli?',
        consistency: 'Il preventivo che le è stato offerto dal concessionario è coerente con quanto comunicato da BYD?',
        finalized: 'Ha finalizzato l’acquisto dell\'auto di suo interesse?'
      });
    }, 800);
  };

  // --- FILTERS & AGGREGATION ---
  
  const uniqueDealers = useMemo(() => {
    const dealers = new Set<string>();
    data.forEach(d => {
      if (d['Dealer Name']) dealers.add(String(d['Dealer Name']));
    });
    return Array.from(dealers).sort();
  }, [data]);

  const uniqueModels = useMemo(() => {
    const models = new Set<string>();
    data.forEach(d => {
      const modelName = d[keyMap.model];
      if (modelName) models.add(String(modelName).trim());
    });
    return Array.from(models).sort();
  }, [data, keyMap]);

  const uniqueSources = useMemo(() => {
    const sources = new Set<string>();
    data.forEach(d => {
      const s = d[keyMap.source];
      if (s) sources.add(String(s).trim());
    });
    return Array.from(sources).sort();
  }, [data, keyMap]);

  // 1. Full Data (Filtered by Dealer/Model/Source for context)
  const fullData = useMemo(() => {
     let res = data;
     if (selectedDealer !== 'All Dealers') {
        res = res.filter(d => String(d['Dealer Name']) === selectedDealer);
     }
     if (selectedModel !== 'All Models') {
        res = res.filter(d => String(d[keyMap.model]).trim() === selectedModel);
     }
     if (selectedSource !== 'All Sources') {
        res = res.filter(d => String(d[keyMap.source]).trim() === selectedSource);
     }
     return res;
  }, [data, selectedDealer, selectedModel, selectedSource, keyMap]);

  // 2. Filtered Data (Based on Lead Status)
  const filteredData = useMemo(() => {
    let res = fullData;
    // Apply Lead Status Filter
    if (leadFilter === 'CONTACTED') {
       // "CONTACTED" implies the user responded "Sì" to taking the survey
       res = res.filter(d => {
         const resp = String(d[keyMap.userResponded] || '').toLowerCase();
         return resp === 'sì' || resp === 'si' || resp === 'yes';
       });
    } else if (leadFilter === 'NOT_CONTACTED') {
       res = res.filter(d => {
         const resp = String(d[keyMap.userResponded] || '').toLowerCase();
         return !resp.includes('sì') && !resp.includes('si') && !resp.includes('yes');
       });
    }
    return res;
  }, [fullData, leadFilter, keyMap]);

  // 4. Malus Metrics (Based on FILTERED data now, as requested)
  const malusMetrics = useMemo(() => {
     const dealerAgg: Record<string, { total: number, contacted: number, dontRemember: number }> = {};
     
     // Iterate filteredData instead of fullData to respect the filter
     filteredData.forEach(row => {
        const dealer = String(row['Dealer Name'] || 'Unknown');
        if (!dealerAgg[dealer]) dealerAgg[dealer] = { total: 0, contacted: 0, dontRemember: 0 };
        
        dealerAgg[dealer].total++;

        const responseVal = String(row[keyMap.responded] || '').toLowerCase();

        // Use the contact question (Q5 in screenshot)
        if (isPositiveResponse(row[keyMap.responded])) {
           dealerAgg[dealer].contacted++;
        }

        // Count "Don't Remember" to exclude from denominator
        if (responseVal.includes('non ricordo') || responseVal.includes('dont remember')) {
            dealerAgg[dealer].dontRemember++;
        }
     });

     let networkMalus = 0;
     const metrics = Object.keys(dealerAgg).map(d => {
        const s = dealerAgg[d];
        const assigned = assignedLeadsMap[d] || s.total; // Use External Assign Count if avail, else survey count
        
        // ADJUSTMENT: Exclude "Don't Remember" from the Base calculation for both rate and penalty
        const adjustedBase = Math.max(0, assigned - s.dontRemember);

        const responseRate = adjustedBase > 0 ? (s.contacted / adjustedBase) * 100 : 0;
        
        let malusPerLead = 0;
        let malusTier = 'Compliant (>90%)';
        let malusTierColor = 'text-green-600 bg-green-50';
        
        if (responseRate < 70) {
          malusPerLead = 8;
          malusTier = 'Critical (<70%)';
          malusTierColor = 'text-red-700 bg-red-50';
        } else if (responseRate < 90) {
          malusPerLead = 4;
          malusTier = 'Warning (70-90%)';
          malusTierColor = 'text-amber-700 bg-amber-50';
        }
        
        // Calculate penalty based on the ADJUSTED base (excluding those who don't remember)
        const totalMalus = Math.round(adjustedBase * malusPerLead); 
        networkMalus += totalMalus;

        return {
           dealer: d,
           total: s.total, // Survey takers (filtered)
           assigned: assigned, // Original Assigned
           adjustedBase, // New Adjusted Base
           contacted: s.contacted,
           responseRate: Math.round(responseRate),
           malusTier,
           malusTierColor,
           malusPerLead,
           totalMalus
        };
     }).sort((a,b) => b.totalMalus - a.totalMalus);
     
     return { metrics, networkMalus };
  }, [filteredData, keyMap, assignedLeadsMap]); 

  // 5. Chart Data
  const chartData = useMemo(() => {
    if (fullData.length === 0) return null;

    // Aggregators
    const dealerAgg: Record<string, any> = {};
    const competitorCounts: Record<string, number> = {};
    
    // Stats for Source Chart (Single Chart for Contact/Channel Performance)
    const sourceStats: Record<string, { total: number, si: number, no: number, dontRemember: number, noData: number }> = {};
    
    // NEW KPI BUCKETS
    const responseFunnel = { si: 0, no: 0, other: 0 };
    const contactFunnel = { si: 0, no: 0, autonomy: 0, dontRemember: 0 };
    
    const networkStats = {
      totalLeads: 0,
      remembered: 0,
      responded: 0,
      visited: 0,
      quoteYes: 0,
      consistYes: 0,
      consistTotal: 0,
      purchasedBYD: 0,
      fin_indeciso: 0,
      fin_rimandato: 0,
      fin_competitor: 0,
      finalizedTotal: 0,
    };

    // 5a. Analyze Full Data for the "Did User Respond?" KPI (Funnel Entry)
    fullData.forEach(row => {
      const resp = String(row[keyMap.userResponded] || '').toLowerCase();
      if (resp === 'sì' || resp === 'si' || resp === 'yes') responseFunnel.si++;
      else if (resp.includes('non vuole')) responseFunnel.no++; 
      else responseFunnel.other++; 
    });

    // 5b. Analyze Filtered Data for Performance & Source Analysis
    // We use filteredData here to ensure Source Analysis reflects the same dataset as the other performance charts
    filteredData.forEach(row => {
      const dealer = String(row['Dealer Name'] || 'Unknown');

      // --- SOURCE ANALYSIS START ---
      const source = String(row[keyMap.source] || 'Unknown');
      
      // Contact Value Analysis (Column N logic - "Dopo aver lasciato...")
      const contactResp = String(row[keyMap.responded] || '').toLowerCase();
      const isContactDontRemember = contactResp.includes('non ricordo') || contactResp.includes('dont remember');
      
      // Single Chart: Channel Performance based on Contact/Response Quality
      if (!sourceStats[source]) sourceStats[source] = { total: 0, si: 0, no: 0, dontRemember: 0, noData: 0 };
      sourceStats[source].total++;
      
      if (isPositiveResponse(contactResp)) {
        sourceStats[source].si++;
      } else if (isContactDontRemember) {
        sourceStats[source].dontRemember++;
      } else if (contactResp.length > 0 && !contactResp.includes('autonomia')) {
        sourceStats[source].no++;
      } else {
        sourceStats[source].noData++;
      }
      // --- SOURCE ANALYSIS END ---


      if (!dealerAgg[dealer]) {
        dealerAgg[dealer] = { 
          totalFiltered: 0, 
          remembered: 0, responded: 0, visited: 0, 
          quoteYes: 0, quoteNo: 0, 
          consistYes: 0, consistNo: 0, 
          fin_byd: 0, fin_indeciso: 0, fin_rimandato: 0, fin_competitor: 0,
          competitors: {} as Record<string, number>
        };
      }

      dealerAgg[dealer].totalFiltered++;
      networkStats.totalLeads++;

      // Contact Funnel (Detailed) - Robust logic
      const contactVal = String(row[keyMap.responded] || '').toLowerCase();
      if (contactVal.includes('autonomia')) {
        contactFunnel.autonomy++;
      } else if (contactVal.includes('ricordo')) {
        contactFunnel.dontRemember++;
      } else if (isPositiveResponse(contactVal)) {
        contactFunnel.si++;
      } else if (contactVal.trim().length > 0) { // Catches "No" and other non-empty, non-positive responses
        contactFunnel.no++;
      }

      // Q1: Remember Request
      // FIXED LOGIC: Only count as remembered if Q1 is Positive AND Q5 is NOT "Don't Remember"
      if (isPositiveResponse(row[keyMap.rememberRequest]) && !isContactDontRemember) {
         dealerAgg[dealer].remembered++;
         networkStats.remembered++;
      }

      // Q2: Response Rate
      if (isPositiveResponse(row[keyMap.responded])) {
        dealerAgg[dealer].responded++;
        networkStats.responded++;
      }

      // Q3: Visited
      if (isPositiveResponse(row[keyMap.visited])) {
        dealerAgg[dealer].visited++;
        networkStats.visited++;
      }

      // Quote Given
      if (isPositiveResponse(row[keyMap.quoteGiven])) {
        dealerAgg[dealer].quoteYes++;
        networkStats.quoteYes++;
      } else {
        dealerAgg[dealer].quoteNo++;
      }

      // Consistency
      const cons = row[keyMap.consistency];
      const sCons = cons ? String(cons).trim() : '';
      if (sCons.length > 0 && sCons !== '.' && sCons !== '-') {
         networkStats.consistTotal++;
         if (isPositiveResponse(sCons)) {
            dealerAgg[dealer].consistYes++;
            networkStats.consistYes++;
         } else {
            dealerAgg[dealer].consistNo++; 
         }
      }

      // Finalized Purchase
      const fin = String(row[keyMap.finalized] || '').toLowerCase();
      let finalizedAnswered = false;
      if (fin.includes('si, byd')) {
        dealerAgg[dealer].fin_byd++;
        networkStats.purchasedBYD++;
        finalizedAnswered = true;
      } else if (fin.includes('indeciso')) {
        dealerAgg[dealer].fin_indeciso++;
        networkStats.fin_indeciso++;
        finalizedAnswered = true;
      } else if (fin.includes('rimandare') || fin.includes('rimandato')) {
        dealerAgg[dealer].fin_rimandato++;
        networkStats.fin_rimandato++;
        finalizedAnswered = true;
      } else if (fin.includes('altro brand') || fin.includes('altro modello')) {
        dealerAgg[dealer].fin_competitor++;
        networkStats.fin_competitor++;
        finalizedAnswered = true;
      }
      
      if (finalizedAnswered) {
        networkStats.finalizedTotal++;
      }

      // Competitors
      const comps = row[keyMap.competitors];
      if (comps && typeof comps === 'string') {
        const parts = comps.split(/[,;\n+]/).map(s => s.trim());
        parts.forEach(rawC => {
          let c = rawC.trim();
          if (!c) return;
          c = c.replace(/^["']|["']$/g, '');
          const match = c.match(/^(?:si|sì|yes)\s*[-:,.]+\s*(.+)/i);
          if (match && match[1]) c = match[1].trim();
          c = c.split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
          competitorCounts[c] = (competitorCounts[c] || 0) + 1;
          if (dealerAgg[dealer]) dealerAgg[dealer].competitors[c] = (dealerAgg[dealer].competitors[c] || 0) + 1;
        });
      }
    });

    const dealerMetrics = Object.keys(dealerAgg).map(d => {
      const fStats = dealerAgg[d];
      const totalF = fStats.totalFiltered;
      
      const rememberedRate = totalF > 0 ? (fStats.remembered / totalF) * 100 : 0;
      const visitedRate = totalF > 0 ? (fStats.visited / totalF) * 100 : 0;
      // Quote Rate based on Visited
      const quoteRate = fStats.visited > 0 ? (fStats.quoteYes / fStats.visited) * 100 : 0;
      const consistTotal = fStats.consistYes + fStats.consistNo;
      const consistRate = consistTotal > 0 ? (fStats.consistYes / consistTotal) * 100 : 0; 
      const contactRate = totalF > 0 ? (fStats.responded / totalF) * 100 : 0;

      let topComp = '-';
      let maxComp = 0;
      Object.entries(fStats.competitors as Record<string, number>).forEach(([comp, count]) => {
          if (count > maxComp) {
            maxComp = count;
            topComp = comp;
          }
      });

      let status = 'Bad';
      let statusColor = 'bg-red-100 text-red-700';
      if (contactRate >= 70) {
        status = 'Great';
        statusColor = 'bg-green-100 text-green-700';
      } else if (contactRate >= 50) {
        status = 'Good';
        statusColor = 'bg-blue-100 text-blue-700';
      } else if (contactRate >= 30) {
        status = 'Needs Improvement';
        statusColor = 'bg-amber-100 text-amber-700';
      }

      return {
        name: d,
        dealer: d,
        total: totalF,
        remembered: fStats.remembered,
        notRemembered: totalF - fStats.remembered,
        responded: fStats.responded,
        notResponded: totalF - fStats.responded,
        visited: fStats.visited,
        'Quote Provided': fStats.quoteYes,
        'No Quote': fStats.quoteNo,
        'Consistent': fStats.consistYes,
        'Inconsistent': fStats.consistNo,
        'fin_byd': fStats.fin_byd,
        'fin_indeciso': fStats.fin_indeciso,
        'fin_rimandato': fStats.fin_rimandato,
        'fin_competitor': fStats.fin_competitor,
        rememberedRate: Math.round(rememberedRate),
        visitedRate: Math.round(visitedRate),
        quoteRate: Math.round(quoteRate),
        consistRate: Math.round(consistRate),
        consistTotal: consistTotal,
        topCompetitor: topComp,
        contactRate: Math.round(contactRate),
        status,
        statusColor
      };
    }).sort((a,b) => b.visitedRate - a.visitedRate); 

    const competitorMetrics = Object.keys(competitorCounts)
      .map(k => {
          const count = competitorCounts[k];
          // Calculate percentage based on total filtered leads (Share of Voice)
          const share = networkStats.totalLeads > 0 ? (count / networkStats.totalLeads) * 100 : 0;
          return { name: k, count, share };
      })
      .sort((a,b) => b.count - a.count)
      .slice(0, 50); 
    
    // Transform Source Stats for Stacked Chart (Single)
    const sourceMetrics = Object.keys(sourceStats).map(s => {
       const stats = sourceStats[s];
       return {
          name: s,
          total: stats.total,
          si: stats.si,
          no: stats.no,
          dontRemember: stats.dontRemember,
          noData: stats.noData
       };
    }).sort((a, b) => b.total - a.total);

    // Transform Funnels for Donuts
    const responseDonutData = [
      { name: 'Sì', value: responseFunnel.si, fill: CHART_COLORS.funnel.si },
      { name: 'No, non vuole rispondere', value: responseFunnel.no, fill: CHART_COLORS.funnel.no },
      { name: 'No', value: responseFunnel.other, fill: CHART_COLORS.funnel.other },
    ].filter(d => d.value > 0);

    const contactDonutData = [
      { name: 'Sì', value: contactFunnel.si, fill: CHART_COLORS.funnel.si },
      { name: 'No', value: contactFunnel.no, fill: CHART_COLORS.funnel.no },
      { name: 'No, recato in autonomia', value: contactFunnel.autonomy, fill: CHART_COLORS.funnel.other },
      { name: 'Non ricordo richiesta', value: contactFunnel.dontRemember, fill: CHART_COLORS.purple },
    ].filter(d => d.value > 0);

    return { dealerMetrics, networkStats, competitorMetrics, sourceMetrics, responseDonutData, contactDonutData };
  }, [filteredData, fullData, keyMap]);

  // Derived sorted data for Malus Table
  const sortedMalusData = useMemo(() => {
    if (!malusMetrics?.metrics) return [];
    return [...malusMetrics.metrics].sort((a: any, b: any) => {
      const valA = a[malusSort.key];
      const valB = b[malusSort.key];
      if (valA === valB) return 0;
      if (valA === undefined) return 1;
      if (valB === undefined) return -1;
      let comparison = 0;
      if (typeof valA === 'string') comparison = valA.localeCompare(valB);
      else comparison = valA < valB ? -1 : 1;
      return malusSort.dir === 'asc' ? comparison : -comparison;
    });
  }, [malusMetrics, malusSort]);

  // -- Download Handlers --
  const handleDownloadMalus = () => {
    if (!malusMetrics) return;
    const exportData = sortedMalusData.map(r => ({
      'Dealer Name': r.dealer,
      'Managed Leads (%)': r.responseRate,
      'Contacted Count': r.contacted,
      'Adjusted Base': r.adjustedBase,
      'Total Assigned (Call Center)': r.assigned,
      'Malus Tier': r.malusTier,
      'Penalty Per Lead (€)': r.malusPerLead,
      'Total Malus (€)': r.totalMalus
    }));
    downloadExcel(exportData, `Malus_Assessment_${new Date().toISOString().split('T')[0]}`);
  };

  const handleDownloadDealerSummary = () => {
    if (!chartData) return;
    const exportData = chartData.dealerMetrics.map(r => ({
      'Dealer Name': r.dealer,
      'Total (Filtered)': r.total,
      'Contact Rate %': r.contactRate,
      'Visit Rate %': r.visitedRate,
      'Quote Rate % (of Visits)': r.quoteRate,
      'Consistency Rate %': r.consistRate,
      'Purchased BYD': r.fin_byd,
      'Top Competitor': r.topCompetitor
    }));
    downloadExcel(exportData, `Dealer_Performance_Summary_${leadFilter}_${new Date().toISOString().split('T')[0]}`);
  };

  const toggleMalusSort = (key: string) => {
    setMalusSort(prev => ({
      key,
      dir: prev.key === key && prev.dir === 'desc' ? 'asc' : 'desc'
    }));
  };

  const MalusHeader = ({ label, sortKey, align = "center" }: { label: string, sortKey: string, align?: string }) => (
    <th 
      className={`p-3 font-semibold text-${align} cursor-pointer hover:bg-slate-100 transition-colors select-none ${align === 'left' ? 'pl-6' : ''} ${align === 'right' ? 'pr-6' : ''}`}
      onClick={() => toggleMalusSort(sortKey)}
    >
      <div className={`flex items-center gap-1 ${align === 'center' ? 'justify-center' : align === 'right' ? 'justify-end' : 'justify-start'}`}>
        {label}
        {malusSort.key === sortKey ? (
           malusSort.dir === 'asc' ? <ChevronUp size={14} className="text-blue-600" /> : <ChevronDown size={14} className="text-blue-600" />
        ) : (
           <div className="w-3.5 h-3.5" /> 
        )}
      </div>
    </th>
  );

  const CustomPieLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percent, index }: any) => {
    const RADIAN = Math.PI / 180;
    const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
    const x = cx + radius * Math.cos(-midAngle * RADIAN);
    const y = cy + radius * Math.sin(-midAngle * RADIAN);
    if (percent < 0.05) return null; // Don't show label for small slices
    return (
      <text x={x} y={y} fill="white" textAnchor={x > cx ? 'start' : 'end'} dominantBaseline="central" fontSize="12" fontWeight="bold">
        {`${(percent * 100).toFixed(0)}%`}
      </text>
    );
  };

  const ReusableDonut = ({ data, title }: { data: any[], title: string }) => (
    <div className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 h-[350px] flex flex-col">
       <div className="mb-2">
         <h3 className="font-bold text-slate-800">{title}</h3>
       </div>
       <div className="flex-1 flex items-center justify-center">
         <ResponsiveContainer width="100%" height="100%">
            <PieChart>
               <Pie
                 data={data}
                 cx="50%"
                 cy="50%"
                 innerRadius={60}
                 outerRadius={90}
                 paddingAngle={2}
                 dataKey="value"
               >
                 {data.map((entry, index) => (
                   <Cell key={`cell-${index}`} fill={entry.fill} />
                 ))}
                 <Label value={data.reduce((a,b) => a + b.value, 0)} position="center" className='font-bold text-2xl fill-slate-700' />
               </Pie>
               <Legend verticalAlign="middle" align="right" layout="vertical" wrapperStyle={{fontSize: '11px'}} />
               <RechartsTooltip />
            </PieChart>
         </ResponsiveContainer>
       </div>
    </div>
  );

  const ProgressBar = ({ value, max, color, label }: { value: number, max: number, color: string, label: string }) => {
    const pct = max > 0 ? (value / max) * 100 : 0;
    return (
      <div className="mb-3">
        <div className="flex justify-between items-center text-xs mb-1">
          <span className="font-medium text-slate-700">{label}</span>
          <span className="font-bold text-slate-900">{pct.toFixed(1)}%</span>
        </div>
        <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
          <div className="h-full rounded-full transition-all duration-500" style={{ width: `${pct}%`, backgroundColor: color }}></div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col md:flex-row font-sans text-slate-900">
      
      {/* Sidebar */}
      <aside className="w-full md:w-64 bg-slate-900 text-white flex-shrink-0 flex flex-col shadow-xl z-20">
        <div className="p-6 border-b border-slate-700 flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold">D</div>
          <span className="text-lg font-bold tracking-tight">DealerInsight</span>
        </div>
        
        <nav className="flex-1 py-6 px-3 space-y-1">
          <button 
            onClick={() => setView(AppView.DASHBOARD)}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-all ${view === AppView.DASHBOARD ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/50' : 'text-slate-400 hover:bg-slate-800 hover:text-white'}`}
          >
            <LayoutDashboard size={20} />
            <span className="font-medium">Overview</span>
          </button>
          
          <button 
            onClick={() => setView(AppView.DATA_GRID)}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-all ${view === AppView.DATA_GRID ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/50' : 'text-slate-400 hover:bg-slate-800 hover:text-white'}`}
          >
            <TableIcon size={20} />
            <span className="font-medium">Data Grid</span>
          </button>
        </nav>

        <div className="p-4 border-t border-slate-800">
          <div className="mb-4 px-2">
            <p className="text-xs text-slate-500 mb-1">Current Source:</p>
            <p className="text-xs font-mono text-slate-300 truncate" title={fileName}>{fileName}</p>
          </div>
          <label className="flex items-center gap-3 px-4 py-3 bg-slate-800 hover:bg-slate-700 rounded-lg cursor-pointer transition-colors text-slate-300 hover:text-white border border-slate-700 mb-2">
            <Upload size={18} />
            <span className="text-sm font-medium">Upload File</span>
            <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleFileUpload} />
          </label>
          <button 
            onClick={loadMockData}
            className="w-full flex items-center gap-3 px-4 py-2 text-slate-500 hover:text-white text-xs hover:bg-slate-800 rounded-lg transition-colors"
          >
             <RefreshCw size={14} />
             Reset to Mock Data
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col h-screen overflow-hidden">
        {/* Header */}
        <header className="bg-white border-b border-slate-200 py-4 px-8 flex flex-col xl:flex-row justify-between items-center shadow-sm z-10 gap-4">
          <div>
            <h1 className="text-2xl font-bold text-slate-800">
              {view === AppView.DASHBOARD && 'Dashboard Overview'}
              {view === AppView.DATA_GRID && 'Data Grid'}
            </h1>
            <p className="text-sm text-slate-500 mt-1">
              Data source: <span className="font-medium text-blue-600">{fileName}</span>
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
             {/* Lead Status Filter (New) */}
             <div className="relative">
                <Phone className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                <select 
                  value={leadFilter}
                  onChange={(e) => setLeadFilter(e.target.value as LeadFilterType)}
                  className="pl-10 pr-8 py-2 bg-indigo-50 border border-indigo-200 rounded-lg text-sm font-bold text-indigo-700 focus:ring-2 focus:ring-indigo-500 outline-none appearance-none min-w-[200px]"
                >
                  <option value="CONTACTED">Managed Leads (Sì)</option>
                  <option value="NOT_CONTACTED">Unmanaged Leads (No)</option>
                  <option value="ALL">All Leads</option>
                </select>
             </div>

             {/* Source Filter (NEW) */}
             <div className="relative">
                <Megaphone className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                <select 
                  value={selectedSource}
                  onChange={(e) => setSelectedSource(e.target.value)}
                  className="pl-10 pr-8 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-medium text-slate-700 focus:ring-2 focus:ring-blue-500 outline-none appearance-none min-w-[150px]"
                >
                  <option value="All Sources">All Sources</option>
                  {uniqueSources.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
             </div>

             {/* Dealer Filter */}
             <div className="relative">
                <Filter className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                <select 
                  value={selectedDealer}
                  onChange={(e) => setSelectedDealer(e.target.value)}
                  className="pl-10 pr-8 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-medium text-slate-700 focus:ring-2 focus:ring-blue-500 outline-none appearance-none min-w-[180px]"
                >
                  <option value="All Dealers">All Dealers</option>
                  {uniqueDealers.map(d => (
                    <option key={d} value={d}>{d}</option>
                  ))}
                </select>
             </div>

             {/* Refresh Button */}
            <button 
              onClick={loadFromSharePoint}
              className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 hover:bg-slate-50 hover:border-slate-400 text-slate-700 rounded-md text-sm font-medium transition-colors"
              title="Refresh data from SharePoint"
            >
              <DownloadCloud size={16} />
              <span className="hidden sm:inline">Refresh</span>
            </button>
          </div>
        </header>

        {/* Content Area */}
        <div className="flex-1 overflow-auto bg-slate-50 p-6 md:p-8">
          
          {loading ? (
             <div className="h-full flex flex-col items-center justify-center text-slate-400">
               <div className="w-16 h-16 border-4 border-blue-600 border-t-transparent rounded-full animate-spin mb-4"></div>
               <p className="font-medium text-slate-600">Processing Data...</p>
               {fileName.includes("Connecting") && <p className="text-xs text-slate-400 mt-2">Connecting to SharePoint...</p>}
             </div>
          ) : (
            <>
              {error && (
                 <div className="mb-6 p-4 bg-amber-50 border border-amber-200 rounded-xl flex items-center gap-3 text-amber-800">
                    <AlertCircle className="flex-shrink-0" size={20} />
                    <span className="text-sm">{error}</span>
                 </div>
              )}

              {/* VIEW: DASHBOARD */}
              {view === AppView.DASHBOARD && (
                chartData && malusMetrics ? (
                  <div className="space-y-6 max-w-[1600px] mx-auto">

                    {/* 0. MALUS ASSESSMENT OVERVIEW */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                      <div className="p-5 bg-slate-50 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-4">
                        <div className="flex items-center gap-3">
                           <div className="p-2 bg-rose-50 text-rose-600 rounded-lg">
                             <AlertTriangle size={20} />
                           </div>
                           <div>
                              <h3 className="text-lg font-bold text-slate-800">Lead Management Malus Overview</h3>
                              <p className="text-xs text-slate-500 mt-0.5">Calculated based on <span className="font-bold">Total Assigned Leads</span> (from Call Center)</p>
                           </div>
                        </div>
                        <div className="flex items-center gap-3">
                           <button 
                              onClick={handleDownloadMalus}
                              className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-200 text-slate-600 hover:text-blue-600 hover:border-blue-200 hover:bg-blue-50 rounded-lg transition-colors shadow-sm"
                              title="Download Malus Report (Excel)"
                           >
                             <Download size={18} />
                             <span className="text-sm font-medium">Export .xlsx</span>
                           </button>

                           <div className="flex items-center gap-3 px-4 py-2 bg-rose-50 border border-rose-100 rounded-lg">
                              <div className="flex flex-col text-right">
                                <span className="text-[10px] font-bold text-rose-400 uppercase tracking-wider">Total Network Malus</span>
                                <span className="text-xl font-bold text-rose-700">€{malusMetrics.networkMalus.toLocaleString()}</span>
                              </div>
                              <Euro className="text-rose-500" size={24} />
                           </div>
                        </div>
                      </div>

                      <div className="p-5 grid grid-cols-1 md:grid-cols-3 gap-6 bg-white border-b border-slate-100 text-sm">
                        <div className="flex items-start gap-3 p-3 bg-red-50 rounded-lg border border-red-100">
                           <div className="w-2 h-full bg-red-500 rounded-full"></div>
                           <div>
                             <h4 className="font-bold text-red-900 mb-1">CRITICAL (&lt;70%)</h4>
                             <p className="text-red-700 text-xs">Dealers responding to fewer than 70% of leads.</p>
                           </div>
                        </div>
                        <div className="flex items-start gap-3 p-3 bg-amber-50 rounded-lg border border-amber-100">
                           <div className="w-2 h-full bg-amber-500 rounded-full"></div>
                           <div>
                             <h4 className="font-bold text-amber-900 mb-1">WARNING (70% - 90%)</h4>
                             <p className="text-amber-700 text-xs">Dealers responding to between 70% and 90% of leads.</p>
                           </div>
                        </div>
                        <div className="flex items-start gap-3 p-3 bg-green-50 rounded-lg border border-green-100">
                           <div className="w-2 h-full bg-green-500 rounded-full"></div>
                           <div>
                             <h4 className="font-bold text-green-900 mb-1">COMPLIANT (&gt;90%)</h4>
                             <p className="text-green-700 text-xs">Dealers responding to more than 90% of leads.</p>
                           </div>
                        </div>
                      </div>
                      
                      <div className="overflow-x-auto max-h-[400px] overflow-y-auto custom-scrollbar">
                        <table className="w-full text-left text-sm relative">
                           <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500 border-b border-slate-200 sticky top-0 z-10 shadow-sm">
                              <tr>
                                <MalusHeader label="Dealer" sortKey="dealer" align="left" />
                                <MalusHeader label="Contacted / Total Assigned" sortKey="responseRate" />
                                <MalusHeader label="Mgmt %" sortKey="responseRate" />
                                <MalusHeader label="Malus Tier" sortKey="malusTier" />
                                <MalusHeader label="Penalty / Lead" sortKey="malusPerLead" />
                                <MalusHeader label="Total Penalty" sortKey="totalMalus" align="right" />
                              </tr>
                           </thead>
                           <tbody className="divide-y divide-slate-100">
                             {sortedMalusData.length > 0 ? sortedMalusData.map((row, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors">
                                   <td className="p-3 pl-6 font-medium text-slate-900">{row.dealer}</td>
                                   <td className="p-3 text-center text-slate-600 font-mono">
                                     {row.contacted} / {row.adjustedBase}
                                   </td>
                                   <td className="p-3 text-center">
                                      <span className={`px-2 py-0.5 rounded text-xs font-bold ${row.responseRate < 70 ? 'bg-red-100 text-red-700' : row.responseRate < 90 ? 'bg-amber-100 text-amber-700' : 'bg-green-100 text-green-700'}`}>
                                        {row.responseRate}%
                                      </span>
                                   </td>
                                   <td className="p-3 text-center">
                                      <span className={`text-xs font-semibold ${row.malusTierColor.split(' ')[0]}`}>
                                        {row.malusTier}
                                      </span>
                                   </td>
                                   <td className="p-3 text-center font-mono text-slate-600">
                                      {row.malusPerLead > 0 ? `€${row.malusPerLead}.00` : '-'}
                                   </td>
                                   <td className="p-3 pr-6 text-right font-bold font-mono text-slate-800">
                                      {row.totalMalus > 0 ? `€${row.totalMalus.toLocaleString()}` : <span className="text-green-500 flex items-center justify-end gap-1"><Check size={14} /> OK</span>}
                                   </td>
                                </tr>
                             )) : (
                               <tr><td colSpan={6} className="p-6 text-center text-slate-400 italic">No dealers match the criteria</td></tr>
                             )}
                           </tbody>
                        </table>
                      </div>
                    </div>

                    {/* NEW SECTION: DONUT CHARTS (KPIS) */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                       <ReusableDonut 
                          title="L'utente ha risposto al sondaggio?" 
                          data={chartData.responseDonutData} 
                       />
                       <ReusableDonut 
                          title="È stato contattato da un concessionario BYD?" 
                          data={chartData.contactDonutData} 
                       />
                    </div>
                    
                    {/* RESTORED SECTION: NETWORK SUMMARY (Filtered) */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 border-l-4 border-l-indigo-500">
                       <div className="flex items-center justify-between mb-5 border-b border-slate-100 pb-3">
                          <div className="flex items-center gap-2">
                            <Globe size={20} className="text-indigo-600" />
                            <div>
                               <h3 className="font-bold text-slate-800 text-lg">Network Performance ({leadFilter === 'CONTACTED' ? 'Managed Leads' : leadFilter === 'ALL' ? 'All Leads' : 'Unmanaged Leads'})</h3>
                               <p className="text-xs text-slate-500">Aggregate metrics based on the current filter selection</p>
                            </div>
                          </div>
                          <div className="flex items-center gap-3">
                             <div className="px-3 py-1 bg-slate-100 rounded text-xs font-medium text-slate-600">
                               Analyzed Leads: <span className="text-slate-900 font-bold">{chartData.networkStats.totalLeads}</span>
                             </div>
                          </div>
                       </div>
                       
                       <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-8">
                          
                          {/* Col 1: Funnel Stats */}
                          <div className="space-y-4">
                             <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                               <Users size={16} className="text-blue-500" />
                               Qualitative Funnel
                             </h4>
                             <div className="space-y-4 p-4 bg-slate-50 rounded-lg border border-slate-100">
                                <ProgressBar 
                                   label="Remember Request" 
                                   value={chartData.networkStats.remembered} 
                                   max={chartData.networkStats.totalLeads} 
                                   color={CHART_COLORS.purple} 
                                />
                                <ProgressBar 
                                   label="Visit Rate" 
                                   value={chartData.networkStats.visited} 
                                   max={chartData.networkStats.totalLeads} 
                                   color={CHART_COLORS.purple} 
                                />
                                <ProgressBar 
                                   label="Quote Rate" 
                                   value={chartData.networkStats.quoteYes} 
                                   max={chartData.networkStats.totalLeads} 
                                   color={CHART_COLORS.teal} 
                                />
                             </div>
                          </div>

                          {/* Col 2: Quote Quality */}
                          <div className="space-y-4">
                             <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                               <Check size={16} className="text-green-500" />
                               Quote Consistency
                             </h4>
                             <div className="p-4 bg-slate-50 rounded-lg border border-slate-100 flex flex-col justify-between gap-4 h-full">
                                
                                <div className="flex items-center justify-between">
                                  <div className="space-y-1">
                                     <div className="flex flex-col">
                                        <span className="text-xs text-slate-500 uppercase tracking-wider mb-1">Consistency</span>
                                        <span className="text-2xl font-bold text-green-600">
                                            {chartData.networkStats.consistTotal > 0 ? Math.round((chartData.networkStats.consistYes/chartData.networkStats.consistTotal)*100) : 0}%
                                        </span>
                                        <span className="text-[10px] text-slate-400">
                                            {chartData.networkStats.consistYes} consistent / {chartData.networkStats.consistTotal} answered
                                        </span>
                                     </div>
                                  </div>
                                  <div className="w-16 h-16 rounded-full border-4 border-slate-200 flex items-center justify-center relative">
                                       <div className="absolute inset-0 rounded-full border-4 border-green-500" style={{ clipPath: `polygon(0 0, 100% 0, 100% ${chartData.networkStats.consistTotal > 0 ? (chartData.networkStats.consistYes/chartData.networkStats.consistTotal)*100 : 0}%, 0 100%)` }}></div>
                                       <span className="text-xs font-bold text-slate-700">
                                         {chartData.networkStats.consistTotal > 0 ? Math.round((chartData.networkStats.consistYes/chartData.networkStats.consistTotal)*100) : 0}%
                                       </span>
                                   </div>
                                </div>
                             </div>
                          </div>

                          {/* Col 3: Purchase Outcomes */}
                          <div className="space-y-4">
                            <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                              <PieIcon size={16} className="text-orange-500" />
                              Purchase Outcomes
                            </h4>
                            <div className="p-4 bg-slate-50 rounded-lg border border-slate-100 space-y-3">
                              {(() => {
                                const { purchasedBYD, fin_indeciso, fin_rimandato, fin_competitor, finalizedTotal } = chartData.networkStats;
                                const calculatePercent = (n, d) => (d > 0 ? (n / d) * 100 : 0);
                                const outcomes = [
                                  { label: 'Sì, BYD', value: purchasedBYD, color: CHART_COLORS.purchase.byd },
                                  { label: 'Indeciso', value: fin_indeciso, color: CHART_COLORS.purchase.indeciso },
                                  { label: 'Rimandato', value: fin_rimandato, color: CHART_COLORS.purchase.rimandato },
                                  { label: 'Competitor', value: fin_competitor, color: CHART_COLORS.purchase.competitor },
                                ];

                                return outcomes.map((outcome, index) => (
                                  <div key={outcome.label} className={index > 0 ? 'pt-1' : ''}>
                                    <div className="flex items-center justify-between text-xs">
                                      <span className="flex items-center gap-2 text-slate-600">
                                        <div className="w-2 h-2 rounded-full" style={{ background: outcome.color }}></div>
                                        {outcome.label}
                                      </span>
                                      <span className="font-bold text-slate-800">
                                        {outcome.value} ({Math.round(calculatePercent(outcome.value, finalizedTotal))}%)
                                      </span>
                                    </div>
                                    <div className="w-full h-1.5 bg-slate-200 rounded-full overflow-hidden mt-1">
                                      <div className="h-full" style={{ width: `${calculatePercent(outcome.value, finalizedTotal)}%`, background: outcome.color }}></div>
                                    </div>
                                  </div>
                                ));
                              })()}
                            </div>
                          </div>

                       </div>
                    </div>

                    {/* NEW SECTION: SOURCE QUALITY - SINGLE CHART */}
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                        <div className="flex items-center justify-between mb-4">
                           <div>
                              <h3 className="text-lg font-bold text-slate-800">Channel Performance: Dealer Contact Quality</h3>
                              <p className="text-xs text-slate-500">Evaluating leads based on their response to "Did the dealer contact you?" (Includes 'Non ricordo')</p>
                           </div>
                        </div>
                        <div className="h-[350px] w-full">
                           <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={chartData.sourceMetrics} margin={{top: 20, right: 30, left: 0, bottom: 20}}>
                                 <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                                 <XAxis 
                                    dataKey="name" 
                                    axisLine={false} 
                                    tickLine={false} 
                                    tick={({ x, y, payload }) => {
                                      const item = chartData.sourceMetrics.find(i => i.name === payload.value);
                                      return (
                                        <g transform={`translate(${x},${y})`}>
                                          <text x={0} y={0} dy={16} textAnchor="middle" fill="#64748b" fontSize={11}>
                                            {payload.value}
                                          </text>
                                          {item && (
                                            <text x={0} y={0} dy={30} textAnchor="middle" fill="#94a3b8" fontSize={10}>
                                              ({item.total})
                                            </text>
                                          )}
                                        </g>
                                      );
                                    }}
                                    interval={0}
                                    height={50}
                                 />
                                 <YAxis axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 11}} />
                                 <RechartsTooltip cursor={{fill: '#f1f5f9'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                 <Legend wrapperStyle={{fontSize: '11px', paddingTop: '10px'}} />
                                 <Bar dataKey="si" name="Sì (Contacted)" stackId="a" fill={CHART_COLORS.source.si} barSize={40} />
                                 <Bar dataKey="no" name="No" stackId="a" fill={CHART_COLORS.source.no} barSize={40} />
                                 <Bar dataKey="dontRemember" name="Non Ricordo" stackId="a" fill={CHART_COLORS.source.dontRemember} barSize={40} />
                                 <Bar dataKey="noData" name="No Survey Data" stackId="a" fill={CHART_COLORS.source.noData} barSize={40} />
                              </BarChart>
                           </ResponsiveContainer>
                        </div>
                    </div>

                    {/* 1. DETAILED TABLE (Filtered) */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mb-8">
                      <div className="p-4 border-b border-slate-100 flex items-center justify-between">
                         <div className="flex items-center gap-3">
                           <div className="p-2 bg-blue-50 text-blue-600 rounded-lg">
                             <TableIcon size={18} />
                           </div>
                           <div>
                              <h3 className="text-base font-bold text-slate-800">Dealer Performance ({leadFilter})</h3>
                           </div>
                         </div>
                         <div className="flex items-center gap-3">
                            <button 
                                onClick={handleDownloadDealerSummary}
                                className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-md transition-colors text-xs font-medium"
                                title="Download Dealer Summary (Excel)"
                             >
                               <Download size={16} />
                               Export .xlsx
                            </button>
                            <div className="text-xs text-slate-400">
                               Showing {chartData.dealerMetrics.length} dealers
                            </div>
                         </div>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left text-xs md:text-sm">
                          <thead className="bg-slate-50/70 border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500">
                            <tr>
                              <th className="p-3 font-semibold">Dealer</th>
                              
                              <th className="p-3 font-semibold text-center">
                                Leads
                                <div className="text-[10px] text-slate-400 font-normal mt-0.5 normal-case">Filtered Count</div>
                              </th>

                              <th className="p-3 font-semibold text-center">
                                Contact %
                                <div className="text-[10px] text-slate-400 font-normal mt-0.5 normal-case">Of Filtered</div>
                              </th>

                              <th className="p-3 font-semibold text-center">
                                Visit %
                                <div className="text-[10px] text-slate-400 font-normal mt-0.5 normal-case">Of Filtered</div>
                              </th>
                              <th className="p-3 font-semibold text-center">
                                Quote %
                                <div className="text-[10px] text-slate-400 font-normal mt-0.5 normal-case">Of Visits</div>
                              </th>
                              <th className="p-3 font-semibold text-center">
                                Consistent %
                                <div className="text-[10px] text-slate-400 font-normal mt-0.5 normal-case">Of Answers</div>
                              </th>
                              <th className="p-3 font-semibold text-center">
                                Purchase %
                                <div className="text-[10px] text-slate-400 font-normal mt-0.5 normal-case">Finalized BYD</div>
                              </th>
                              <th className="p-3 font-semibold text-right">Status</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {[...chartData.dealerMetrics].sort((a,b) => b.visitedRate - a.visitedRate).map((item, idx) => (
                              <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                <td className="p-3 font-semibold text-slate-900">{item.dealer}</td>
                                
                                {/* Leads Count (Filtered) */}
                                <td className="p-3 text-center text-slate-600">{item.total}</td>

                                {/* Contact Rate (Filtered Context) */}
                                <td className="p-3 text-center">
                                   <div className="flex flex-col items-center">
                                    <span className={`font-bold ${item.contactRate > 60 ? 'text-blue-600' : 'text-slate-600'}`}>
                                      {item.contactRate}%
                                    </span>
                                    <span className="text-[10px] text-slate-400">({item.responded}/{item.total})</span>
                                  </div>
                                </td>

                                <td className="p-3 text-center">
                                  <div className="flex flex-col items-center">
                                    <span className="text-slate-600 font-medium">{item.visitedRate}%</span>
                                    <span className="text-[10px] text-slate-400">({item.visited}/{item.total})</span>
                                  </div>
                                </td>

                                <td className="p-3 text-center">
                                   <div className="flex flex-col items-center">
                                    <span className={`font-medium ${item.quoteRate > 50 ? 'text-indigo-600' : 'text-slate-600'}`}>
                                      {item.quoteRate.toFixed(0)}%
                                    </span>
                                    <span className="text-[10px] text-slate-400">({item['Quote Provided']}/{item.visited})</span>
                                  </div>
                                </td>

                                <td className="p-3 text-center">
                                   <div className="flex flex-col items-center">
                                    <span className={`font-medium ${item.consistRate > 80 ? 'text-green-600' : 'text-amber-600'}`}>
                                      {item.consistRate}%
                                    </span>
                                    {item.consistTotal > 0 && (
                                      <span className="text-[10px] text-slate-400">({item.Consistent}/{item.consistTotal})</span>
                                    )}
                                  </div>
                                </td>

                                <td className="p-3 text-center">
                                   <div className="flex flex-col items-center">
                                    <span className={`font-medium ${item.fin_byd > 0 ? 'text-teal-600' : 'text-slate-600'}`}>
                                      {Math.round((item.fin_byd / item.total) * 100)}%
                                    </span>
                                  </div>
                                </td>
                                
                                <td className="p-3 text-right">
                                  <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider ${item.statusColor}`}>
                                    {item.status}
                                  </span>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* RESTORED: DETAILED QUESTION CHARTS */}
                      <div className="grid grid-cols-1 gap-8 pb-10">
                        {[
                          {
                            title: 'Did you remember making a request?',
                            icon: BrainCircuit,
                            desc: '"Ha effettuato una richiesta...?"',
                            color: 'text-purple-500',
                            dataKey1: 'remembered', name1: 'Yes', fill1: CHART_COLORS.purple,
                            dataKey2: 'notRemembered', name2: 'No', fill2: CHART_COLORS.neutral
                          },
                          {
                            title: 'Dealer Contact Response',
                            icon: MessageSquare,
                            desc: '"Dopo aver lasciato i suoi dati..." (Filtered View)',
                             color: 'text-blue-500',
                            dataKey1: 'responded', name1: 'Yes', fill1: CHART_COLORS.positive,
                            dataKey2: 'notResponded', name2: 'No', fill2: CHART_COLORS.negative
                          },
                          {
                            title: 'Quote Provided',
                            icon: TrendingUp,
                            desc: '"Il concessionario le ha fornito un preventivo...?"',
                             color: 'text-indigo-500',
                            dataKey1: 'Quote Provided', name1: 'Yes', fill1: CHART_COLORS.info,
                            dataKey2: 'No Quote', name2: 'No', fill2: CHART_COLORS.neutral
                          },
                          {
                            title: 'Quote Consistency',
                            icon: Check,
                            desc: '"Il preventivo... è coerente...?"',
                             color: 'text-green-500',
                            dataKey1: 'Consistent', name1: 'Yes', fill1: CHART_COLORS.positive,
                            dataKey2: 'Inconsistent', name2: 'No', fill2: CHART_COLORS.warning
                          },
                          {
                            title: 'Purchase Finalized (Detail)',
                             icon: ShoppingCart,
                             desc: '"Ha finalizzato l’acquisto dell\'auto...?"',
                             color: 'text-teal-500',
                             isStacked: true,
                             keys: [
                               { key: 'fin_byd', name: 'Sì, BYD', fill: CHART_COLORS.purchase.byd },
                               { key: 'fin_indeciso', name: 'No, Indeciso', fill: CHART_COLORS.purchase.indeciso },
                               { key: 'fin_rimandato', name: 'No, Rimandato', fill: CHART_COLORS.purchase.rimandato },
                               { key: 'fin_competitor', name: 'Sì, Altro Brand', fill: CHART_COLORS.purchase.competitor }
                             ]
                          },
                          {
                            title: 'Are they looking at other models?',
                            icon: Award,
                            desc: '"Sta valutando altri Brand/Modelli?" (Share of mentions vs Total Leads)',
                             color: 'text-purple-500',
                            isCompetitor: true
                          }
                        ].map((config, i) => (
                           <div key={i} className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 transition-all hover:shadow-md">
                             <div className="mb-6 flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
                                <div>
                                   <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                                     <config.icon size={18} className={config.color} />
                                     {config.title}
                                   </h3>
                                   <p className="text-xs text-slate-500 mt-1 italic">{config.desc}</p>
                                </div>
                             </div>
                             
                             <div className="h-[300px] w-full">
                               <ResponsiveContainer width="100%" height="100%">
                                  {config.isCompetitor ? (
                                    <BarChart layout="vertical" data={chartData.competitorMetrics} margin={{top: 5, right: 30, left: 40, bottom: 5}}>
                                       <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                                       <XAxis type="number" hide />
                                       <YAxis dataKey="name" type="category" width={100} tick={{fontSize: 11, fill: '#64748b'}} />
                                       <RechartsTooltip cursor={{fill: '#f8fafc'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                       <Bar dataKey="count" name="Mentions" fill="#8b5cf6" radius={[0, 4, 4, 0]} barSize={20} />
                                    </BarChart>
                                  ) : config.isStacked ? (
                                     <BarChart data={chartData.dealerMetrics} margin={{top: 20, right: 30, left: 0, bottom: 20}}>
                                       <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                                       <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 11, fill: '#64748b'}} interval={0} height={50} angle={-45} textAnchor="end" />
                                       <YAxis axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 11}} />
                                       <RechartsTooltip cursor={{fill: '#f1f5f9'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                       <Legend wrapperStyle={{fontSize: '11px', paddingTop: '10px'}} />
                                       {config.keys && config.keys.map((k, idx) => (
                                          <Bar key={idx} dataKey={k.key} name={k.name} stackId="a" fill={k.fill} />
                                       ))}
                                     </BarChart>
                                  ) : (
                                     <BarChart data={chartData.dealerMetrics} margin={{top: 20, right: 30, left: 0, bottom: 20}}>
                                       <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                                       <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 11, fill: '#64748b'}} interval={0} height={50} angle={-45} textAnchor="end" />
                                       <YAxis axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 11}} />
                                       <RechartsTooltip cursor={{fill: '#f1f5f9'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                       <Legend wrapperStyle={{fontSize: '11px', paddingTop: '10px'}} />
                                       <Bar dataKey={config.dataKey1!} name={config.name1} fill={config.fill1} radius={[4, 4, 0, 0]} />
                                       <Bar dataKey={config.dataKey2!} name={config.name2} fill={config.fill2} radius={[4, 4, 0, 0]} />
                                     </BarChart>
                                  )}
                               </ResponsiveContainer>
                             </div>
                           </div>
                        ))}
                      </div>

                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full text-slate-400">
                    <p>No data available for the dashboard.</p>
                  </div>
                )
              )}

              {/* VIEW: DATA GRID */}
              {view === AppView.DATA_GRID && (
                <div className="h-full">
                  <PivotTable data={fullData} columns={columns} />
                </div>
              )}

            </>
          )}
        </div>
      </main>
    </div>
  );
}
