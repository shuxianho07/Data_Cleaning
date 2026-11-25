import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Upload, 
  FileText, 
  Download, 
  Activity, 
  AlertTriangle, 
  BarChart2, 
  Settings, 
  Trash2, 
  CheckCircle,
  Search,
  Filter
} from 'lucide-react';

/* --- EXTERNAL LIBRARIES ---
  Note: This app uses SheetJS (XLSX) loaded via CDN for Excel parsing.
  In a production environment, you might prefer `npm install xlsx` and import it directly.
*/
const LoadScripts = () => {
  useEffect(() => {
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js";
      script.async = true;
      document.body.appendChild(script);
    }
  }, []);
  return null;
};

/* --- HELPER FUNCTIONS --- */

// Keep a missing or invalid close distinct from a genuine price.
const normalizeClose = (raw) => {
  if (raw == null || (typeof raw === 'string' && raw.trim() === '')) {
    return { close: null, closeIssue: 'Missing close' };
  }
  const close = typeof raw === 'number' || typeof raw === 'string' ? Number(raw) : NaN;
  if (!Number.isFinite(close) || close <= 0) {
    return { close: null, closeIssue: 'Invalid close (expected a positive number)' };
  }
  return { close, closeIssue: null };
};

const displayClose = (close) => Number.isFinite(close) ? close.toFixed(2) : '—';

// Parse CSV Text
const parseCSV = (text) => {
  const lines = text.split('\n').filter(l => l.trim());
  if (!lines.length) return [];
  const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
  // Respect an explicit Close column, even when that cell is blank.
  const closeIndex = headers.includes('close') ? headers.indexOf('close') : headers.indexOf('price');
  
  return lines.slice(1).map((line, index) => {
    const values = line.split(',');
    const entry = {};
    headers.forEach((h, i) => {
      let val = values[i]?.trim();
      if (h === 'date' || h === 'time' || h === 'timestamp') {
        entry.date = val;
      } else {
        // Try to convert to number
        const num = parseFloat(val);
        entry[h] = isNaN(num) ? val : num;
      }
    });
    // Normalization
    return {
      id: index, // unique id for React keys
      date: entry.date || entry.timestamp || new Date().toISOString(),
      open: entry.open || entry.price || 0,
      high: entry.high || entry.open || 0,
      low: entry.low || entry.open || 0,
      ...normalizeClose(closeIndex === -1 ? undefined : values[closeIndex]?.trim().replace(/^"(.*)"$/, '$1')),
      volume: entry.volume || 0,
      original: entry
    };
  });
};

// Generate Sample Data for Demo
const generateSampleData = () => {
  const data = [];
  let price = 150;
  const now = new Date();
  for (let i = 0; i < 200; i++) {
    const date = new Date(now);
    date.setDate(date.getDate() - (200 - i));
    
    // Random walk
    const change = (Math.random() - 0.5) * 5;
    price += change;
    
    // Add some noise/outliers
    let open = price;
    let close = price + (Math.random() - 0.5) * 2;
    let high = Math.max(open, close) + Math.random() * 2;
    let low = Math.min(open, close) - Math.random() * 2;
    let volume = 1000000 + Math.random() * 500000;

    // Inject Black Swan
    if (i === 150) {
      close = open * 0.85; // 15% drop
      low = close;
      volume = 5000000; // Huge volume
    }
    // Inject Random Spike
    if (i === 50) {
      high = high * 1.10; // 10% spike up
    }

    data.push({
      id: i,
      date: date.toISOString().split('T')[0],
      open, high, low, close, volume
    });
  }
  return data;
};

// Stats Calculations
const calculateStats = (data, key = 'close') => {
  const values = data.map(d => d[key]).filter(Number.isFinite).sort((a, b) => a - b);
  if (!values.length) return null;
  const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
  const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
  const stdDev = Math.sqrt(variance);

  const q1 = values[Math.floor(values.length * 0.25)];
  const q3 = values[Math.floor(values.length * 0.75)];
  const iqr = q3 - q1;

  return { mean, stdDev, q1, q3, iqr };
};

/* --- COMPONENTS --- */

const FileUploader = ({ onDataLoaded }) => {
  const fileInputRef = useRef(null);
  const [dragActive, setDragActive] = useState(false);
  const [error, setError] = useState(null);

  const processFile = (file) => {
    setError(null);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        let parsedData = [];

        if (file.name.endsWith('.csv')) {
          parsedData = parseCSV(data);
        } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
          if (window.XLSX) {
            const workbook = window.XLSX.read(data, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = window.XLSX.utils.sheet_to_json(sheet);
            // Normalize
            parsedData = jsonData.map((entry, index) => ({
              id: index,
              date: entry.Date || entry.date || 'N/A',
              open: entry.Open || entry.open || 0,
              high: entry.High || entry.high || 0,
              low: entry.Low || entry.low || 0,
              ...normalizeClose(entry.Close ?? entry.close),
              volume: entry.Volume || entry.volume || 0,
            }));
          } else {
            throw new Error("Excel parser not loaded yet. Please wait a moment or use CSV.");
          }
        } else {
          throw new Error("Unsupported file type. Please use CSV or Excel.");
        }

        if (parsedData.length === 0) throw new Error("No data found in file.");
        onDataLoaded(parsedData, file.name);

      } catch (err) {
        setError(err.message);
      }
    };

    if (file.name.endsWith('.csv')) {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  return (
    <div className="w-full">
      <div 
        className={`border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer
          ${dragActive ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'}
        `}
        onDragEnter={() => setDragActive(true)}
        onDragLeave={() => setDragActive(false)}
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current && fileInputRef.current.click()}
      >
        <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
        <h3 className="text-lg font-medium text-gray-700">Drop your Stock Data (CSV/Excel)</h3>
        <p className="text-sm text-gray-500 mt-2">or click to browse</p>
        <input 
          ref={fileInputRef}
          type="file" 
          accept=".csv,.xlsx,.xls" 
          className="hidden" 
          onChange={(e) => e.target.files[0] && processFile(e.target.files[0])}
        />
      </div>
      
      {error && (
        <div className="mt-4 p-3 bg-red-100 text-red-700 rounded-lg flex items-center gap-2">
          <AlertTriangle size={18} />
          {error}
        </div>
      )}

      <div className="mt-4 flex justify-center">
        <button 
          onClick={() => onDataLoaded(generateSampleData(), "Sample_S&P500.csv")}
          className="text-blue-600 text-sm hover:underline font-medium"
        >
          Use Sample Data (S&P 500 Simulation)
        </button>
      </div>
    </div>
  );
};

// Custom SVG Chart Component
const StockChart = ({ data, outliers }) => {
  const containerRef = useRef(null);
  const [dimensions, setDimensions] = useState({ width: 0, height: 300 });

  useEffect(() => {
    if (containerRef.current) {
      setDimensions({
        width: containerRef.current.offsetWidth,
        height: 300
      });
    }
    const handleResize = () => {
      if(containerRef.current) setDimensions({ width: containerRef.current.offsetWidth, height: 300 });
    }
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  if (!data || data.length === 0) return <div className="h-64 bg-gray-100 rounded flex items-center justify-center text-gray-400">No Data for Chart</div>;

  // Scales
  let minPrice = Infinity;
  let maxPrice = -Infinity;
  for (const { close } of data) {
    if (!Number.isFinite(close)) continue;
    minPrice = Math.min(minPrice, close);
    maxPrice = Math.max(maxPrice, close);
  }
  if (!Number.isFinite(minPrice)) {
    return <div className="h-64 rounded flex items-center justify-center text-gray-400">No valid close prices to chart.</div>;
  }
  const lower = minPrice * 0.95;
  const upper = maxPrice * 1.05;
  const padding = 40;
  // Ensure chartWidth and chartHeight are positive
  const chartWidth = Math.max(0, dimensions.width - padding * 2);
  const chartHeight = Math.max(0, dimensions.height - padding * 2);

  const getX = (index) => padding + (data.length === 1 ? 0.5 : index / (data.length - 1)) * chartWidth;
  const getY = (price) => padding + chartHeight - ((price - lower) / (upper - lower)) * chartHeight;

  // Start a new line segment after a missing close; never draw through the gap.
  let pathD = '';
  let previousValid = false;
  data.forEach((d, i) => {
    const valid = Number.isFinite(d.close);
    if (valid) pathD += ` ${previousValid ? 'L' : 'M'} ${getX(i)} ${getY(d.close)}`;
    previousValid = valid;
  });

  return (
    <div ref={containerRef} className="w-full relative select-none" style={{ minHeight: '300px' }}>
      <svg width={dimensions.width} height={dimensions.height} className="overflow-visible">
        {/* Grid Lines */}
        {[0, 0.25, 0.5, 0.75, 1].map((tick, i) => (
          <line 
            key={i} 
            x1={padding} 
            y1={padding + chartHeight * tick} 
            x2={dimensions.width - padding} 
            y2={padding + chartHeight * tick} 
            stroke="#e5e7eb" 
            strokeDasharray="4 4" 
          />
        ))}

        {/* Main Line */}
        <path d={pathD} fill="none" stroke="#2563eb" strokeWidth="2" />

        {/* Outlier Points */}
        {data.map((d, i) => {
          const isOutlier = outliers[i] && outliers[i].length > 0;
          if (isOutlier && Number.isFinite(d.close)) {
            return (
              <g key={d.id || i} className="group cursor-pointer">
                <circle 
                  cx={getX(i)} 
                  cy={getY(d.close)} 
                  r="4" 
                  fill="#ef4444" 
                  stroke="white" 
                  strokeWidth="2" 
                />
                <title>{`Date: ${d.date}\nPrice: ${d.close.toFixed(2)}\nOutliers: ${outliers[i].join(', ')}`}</title>
              </g>
            );
          }
          return null;
        })}
      </svg>
      
      {/* Axis Labels */}
      <div className="absolute left-0 top-0 h-full flex flex-col justify-between py-10 text-xs text-gray-500 px-1 pointer-events-none">
        <span>{upper.toFixed(0)}</span>
        <span>{lower.toFixed(0)}</span>
      </div>
      <div className="absolute bottom-0 left-0 w-full flex justify-between px-10 text-xs text-gray-500 pointer-events-none">
        <span>{data[0].date}</span>
        <span>{data[data.length - 1].date}</span>
      </div>
    </div>
  );
};

/* --- MAIN APP --- */

export default function App() {
  // State
  const [data, setData] = useState(null);
  const [filename, setFilename] = useState("");
  const [view, setView] = useState('dashboard'); // 'dashboard', 'data'
  
  // Configuration
  const [config, setConfig] = useState({
    zScoreThreshold: 2.5,
    iqrMultiplier: 1.5,
    blackSwanThreshold: 0.10, // 10% drop
    enableZScore: true,
    enableIQR: true,
    enableBlackSwan: true,
    movingAverageWindow: 5
  });

  // Derived State (useMemo to prevent null render cycles)
  const { processedData, outlierMap, stats } = useMemo(() => {
    if (!data || data.length === 0) {
      return { processedData: [], outlierMap: {}, stats: null };
    }

    const calculatedStats = calculateStats(data, 'close');
    const newOutliers = {};
    
    const processed = data.map((entry, index) => {
      const labels = [];
      const flags = [];

      const validClose = Number.isFinite(entry.close);

      // 1. Z-Score Analysis
      if (validClose && config.enableZScore && calculatedStats?.stdDev > 0) {
        const zScore = (entry.close - calculatedStats.mean) / calculatedStats.stdDev;
        if (Math.abs(zScore) > config.zScoreThreshold) {
          labels.push(`Z-Score (${zScore.toFixed(2)})`);
          flags.push('statistical');
        }
      }

      // 2. IQR Analysis
      if (validClose && config.enableIQR) {
        const lowerBound = calculatedStats.q1 - (config.iqrMultiplier * calculatedStats.iqr);
        const upperBound = calculatedStats.q3 + (config.iqrMultiplier * calculatedStats.iqr);
        if (entry.close < lowerBound || entry.close > upperBound) {
          labels.push(`IQR Outlier`);
          flags.push('statistical');
        }
      }

      // 3. Black Swan / Crash Detection (Day-over-Day change)
      const prevClose = data[index - 1]?.close;
      if (config.enableBlackSwan && validClose && Number.isFinite(prevClose) && prevClose > 0) {
        const change = (entry.close - prevClose) / prevClose;
        if (change < -config.blackSwanThreshold) {
          labels.push(`Black Swan Drop (${(change * 100).toFixed(1)}%)`);
          flags.push('critical');
        } else if (change > config.blackSwanThreshold) {
          labels.push(`Extreme Surge (+${(change * 100).toFixed(1)}%)`);
          flags.push('critical');
        }
      }

      if (labels.length > 0) {
        newOutliers[index] = labels;
      }

      return {
        ...entry,
        isOutlier: labels.length > 0,
        outlierLabels: labels,
        outlierFlags: flags
      };
    });

    return { 
      processedData: processed, 
      outlierMap: newOutliers, 
      stats: calculatedStats 
    };
  }, [data, config]);


  const handleDataLoaded = (rawData, name) => {
    setData(rawData);
    setFilename(name);
  };

  const downloadCleanedData = (includeOutliers = false) => {
    if (!processedData.length) return;
    
    // Filter logic
    const exportData = includeOutliers
      ? processedData
      : processedData.filter(d => !d.isOutlier && !d.closeIssue);

    // Convert to CSV
    const headers = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume', 'Is_Outlier', 'Outlier_Reason', 'Validation_Issue'];
    const csvContent = [
      headers.join(','),
      ...exportData.map(d => [
        d.date, d.open, d.high, d.low, d.close, d.volume, 
        d.isOutlier ? 'TRUE' : 'FALSE',
        d.isOutlier ? `"${d.outlierLabels.join('; ')}"` : '',
        d.closeIssue || ''
      ].join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const baseName = filename.replace(/\.[^.]+$/, '');
    link.setAttribute('download', `${includeOutliers ? 'Flagged' : 'Unflagged'}_${baseName}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-900">
      <LoadScripts />
      
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-20">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between h-16 items-center">
            <div className="flex items-center gap-2">
              <Activity className="h-8 w-8 text-blue-600" />
              <div>
                <h1 className="text-xl font-bold text-gray-900">Market<span className="text-blue-600">Cleaner</span> AI</h1>
                <p className="text-xs text-gray-500">Outlier Detection & Data Hygiene System</p>
              </div>
            </div>
            {data && (
              <div className="flex items-center gap-4">
                 <div className="hidden md:flex text-sm text-gray-500 bg-gray-100 px-3 py-1 rounded-full items-center gap-2">
                    <FileText size={14} />
                    {filename}
                 </div>
                 <button 
                  onClick={() => setData(null)} 
                  className="text-sm text-red-600 hover:text-red-700 font-medium"
                >
                  Reset
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {/* Render Logic: Check processedData length, not just raw data to be safe */}
        {!data || !processedData || processedData.length === 0 ? (
          /* Empty State / Upload */
          <div className="max-w-2xl mx-auto mt-10">
            <div className="bg-white rounded-2xl shadow-sm border border-gray-200 p-8">
              <h2 className="text-2xl font-bold text-center mb-2">Import Market Data</h2>
              <p className="text-center text-gray-500 mb-8">Upload historical OHLCV data to begin analysis. Supports .csv and .xlsx</p>
              <FileUploader onDataLoaded={handleDataLoaded} />
              
              <div className="mt-10 grid grid-cols-1 md:grid-cols-3 gap-4">
                {[
                  { title: "Z-Score Analysis", desc: "Identify statistical anomalies deviation from mean." },
                  { title: "Black Swan Detector", desc: "Flag extreme market crashes (e.g. -10%)." },
                  { title: "Clean & Export", desc: "Generate ML-ready datasets instantly." }
                ].map((f, i) => (
                  <div key={i} className="bg-gray-50 p-4 rounded-lg text-center">
                    <h3 className="font-semibold text-sm text-gray-900">{f.title}</h3>
                    <p className="text-xs text-gray-500 mt-1">{f.desc}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        ) : (
          /* Dashboard */
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
            
            {/* Left Column: Controls & Stats */}
            <div className="space-y-6">
              
              {/* Stats Card */}
              <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
                <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
                  <BarChart2 size={20} className="text-gray-400" />
                  Data Summary
                </h3>
                <div className="grid grid-cols-2 gap-4">
                  <div className="p-3 bg-blue-50 rounded-lg">
                    <p className="text-xs text-blue-600 uppercase font-bold tracking-wider">Total Rows</p>
                    <p className="text-2xl font-bold text-blue-900">{processedData.length}</p>
                  </div>
                  <div className="p-3 bg-red-50 rounded-lg">
                    <p className="text-xs text-red-600 uppercase font-bold tracking-wider">Outliers Found</p>
                    <p className="text-2xl font-bold text-red-900">{Object.keys(outlierMap).length}</p>
                  </div>
                  <div className="col-span-2 p-3 bg-amber-50 rounded-lg">
                    <p className="text-xs text-amber-700 uppercase font-bold tracking-wider">Close Price Validation Issues</p>
                    <p className="text-2xl font-bold text-amber-900">{processedData.filter(row => row.closeIssue).length}</p>
                  </div>
                </div>
                {stats && (
                  <div className="mt-4 pt-4 border-t border-gray-100 text-sm space-y-2 text-gray-600">
                    <div className="flex justify-between">
                      <span>Mean Price:</span>
                      <span className="font-mono">{stats.mean.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between">
                      <span>Std Dev:</span>
                      <span className="font-mono">{stats.stdDev.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between">
                      <span>IQR:</span>
                      <span className="font-mono">{stats.iqr.toFixed(2)}</span>
                    </div>
                  </div>
                )}
              </div>

              {/* Controls Card */}
              <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
                <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
                  <Settings size={20} className="text-gray-400" />
                  Detection Settings
                </h3>
                
                <div className="space-y-6">
                  {/* Z-Score Control */}
                  <div>
                    <div className="flex items-center justify-between mb-2">
                      <label className="text-sm font-medium flex items-center gap-2">
                        <input 
                          type="checkbox" 
                          checked={config.enableZScore}
                          onChange={e => setConfig({...config, enableZScore: e.target.checked})}
                          className="rounded text-blue-600 focus:ring-blue-500"
                        />
                        Z-Score Threshold
                      </label>
                      <span className="text-xs font-mono bg-gray-100 px-2 py-1 rounded">{config.zScoreThreshold} σ</span>
                    </div>
                    <input 
                      type="range" 
                      min="1" max="5" step="0.1"
                      value={config.zScoreThreshold}
                      onChange={e => setConfig({...config, zScoreThreshold: parseFloat(e.target.value)})}
                      className="w-full h-2 bg-gray-200 rounded-lg appearance-none cursor-pointer"
                      disabled={!config.enableZScore}
                    />
                  </div>

                  {/* IQR Control */}
                  <div>
                    <div className="flex items-center justify-between mb-2">
                      <label className="text-sm font-medium flex items-center gap-2">
                         <input 
                          type="checkbox" 
                          checked={config.enableIQR}
                          onChange={e => setConfig({...config, enableIQR: e.target.checked})}
                          className="rounded text-blue-600 focus:ring-blue-500"
                        />
                        IQR Multiplier
                      </label>
                      <span className="text-xs font-mono bg-gray-100 px-2 py-1 rounded">{config.iqrMultiplier}x</span>
                    </div>
                    <input 
                      type="range" 
                      min="0.5" max="3" step="0.1"
                      value={config.iqrMultiplier}
                      onChange={e => setConfig({...config, iqrMultiplier: parseFloat(e.target.value)})}
                      className="w-full h-2 bg-gray-200 rounded-lg appearance-none cursor-pointer"
                      disabled={!config.enableIQR}
                    />
                  </div>

                  {/* Black Swan Control */}
                  <div>
                    <div className="flex items-center justify-between mb-2">
                      <label className="text-sm font-medium flex items-center gap-2">
                         <input 
                          type="checkbox" 
                          checked={config.enableBlackSwan}
                          onChange={e => setConfig({...config, enableBlackSwan: e.target.checked})}
                          className="rounded text-blue-600 focus:ring-blue-500"
                        />
                        Crash Threshold
                      </label>
                      <span className="text-xs font-mono bg-gray-100 px-2 py-1 rounded">{Math.round(config.blackSwanThreshold * 100)}%</span>
                    </div>
                    <input 
                      type="range" 
                      min="0.01" max="0.5" step="0.01"
                      value={config.blackSwanThreshold}
                      onChange={e => setConfig({...config, blackSwanThreshold: parseFloat(e.target.value)})}
                      className="w-full h-2 bg-gray-200 rounded-lg appearance-none cursor-pointer"
                      disabled={!config.enableBlackSwan}
                    />
                  </div>
                </div>
              </div>

              {/* Actions */}
              <div className="space-y-3">
                <button 
                  onClick={() => downloadCleanedData(true)}
                  className="w-full flex items-center justify-center gap-2 bg-blue-600 text-white px-4 py-3 rounded-xl hover:bg-blue-700 transition font-medium shadow-sm"
                >
                  <Download size={18} />
                  Export with Flags (All Rows)
                </button>
                <button 
                  onClick={() => downloadCleanedData(false)}
                  className="w-full flex items-center justify-center gap-2 bg-white border border-gray-300 text-gray-700 px-4 py-3 rounded-xl hover:bg-gray-50 transition font-medium"
                >
                  <FileText size={18} />
                  Export Unflagged Rows Only
                </button>
              </div>

            </div>

            {/* Right Column: Visualization & Details */}
            <div className="lg:col-span-2 space-y-6">
              
              {/* Tabs */}
              <div className="bg-white rounded-t-xl border-b border-gray-200 px-6 pt-4 flex gap-6">
                <button 
                  onClick={() => setView('dashboard')}
                  className={`pb-4 text-sm font-medium border-b-2 transition ${view === 'dashboard' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
                >
                  Visualization
                </button>
                <button 
                  onClick={() => setView('data')}
                  className={`pb-4 text-sm font-medium border-b-2 transition ${view === 'data' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
                >
                  Data Table ({processedData.length})
                </button>
              </div>

              {/* View Content */}
              <div className="bg-white rounded-b-xl shadow-sm border border-gray-200 p-6 min-h-[500px]">
                
                {view === 'dashboard' ? (
                  <>
                    <div className="flex justify-between items-center mb-6">
                      <h3 className="font-bold text-gray-800">Price History & Anomalies</h3>
                      <div className="flex items-center gap-2 text-xs text-gray-500">
                        <span className="w-3 h-3 bg-red-500 rounded-full"></span> Outlier Detected
                        <span className="w-3 h-3 bg-blue-500 rounded-full ml-2"></span> Normal Price
                      </div>
                    </div>
                    
                    <StockChart data={processedData} outliers={outlierMap} />
                    
                    <div className="mt-8">
                       <h4 className="font-semibold text-gray-800 mb-4">Price Review Log</h4>
                       <div className="max-h-64 overflow-y-auto pr-2 space-y-2 custom-scrollbar">
                        {processedData.filter(row => row.isOutlier || row.closeIssue).map(item => {
                          const reasons = item.closeIssue ? [item.closeIssue] : item.outlierLabels;
                          return (
                            <div key={item.id} className="flex items-start gap-3 p-3 rounded-lg border border-red-100 bg-red-50">
                              <AlertTriangle className="text-red-500 shrink-0 mt-0.5" size={16} />
                              <div>
                                <div className="flex items-center gap-2 mb-1">
                                  <span className="text-sm font-bold text-gray-900">{item.date}</span>
                                  <span className="text-xs bg-white px-2 py-0.5 rounded border border-red-200 text-red-700 font-mono">
                                    {displayClose(item.close)}
                                  </span>
                                </div>
                                <div className="flex flex-wrap gap-2">
                                  {reasons.map((r, i) => (
                                    <span key={i} className="text-xs text-red-600 bg-red-100 px-1.5 py-0.5 rounded">{r}</span>
                                  ))}
                                </div>
                              </div>
                            </div>
                          );
                        })}
                        {!processedData.some(row => row.isOutlier || row.closeIssue) && (
                          <div className="text-center py-8 text-gray-400">
                            <CheckCircle className="mx-auto mb-2 text-green-500" size={24} />
                            <p>No outliers or close price issues detected.</p>
                          </div>
                        )}
                       </div>
                    </div>
                  </>
                ) : (
                  /* Data Table View */
                  <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-gray-200">
                      <thead className="bg-gray-50">
                        <tr>
                          {['Date', 'Open', 'High', 'Low', 'Close', 'Status'].map(h => (
                            <th key={h} className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="bg-white divide-y divide-gray-200">
                        {processedData.slice(0, 100).map((row, i) => (
                          <tr key={i} className={row.closeIssue ? 'bg-amber-50' : row.isOutlier ? 'bg-red-50' : 'hover:bg-gray-50'}>
                            <td className="px-3 py-2 whitespace-nowrap text-sm text-gray-900">{row.date}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-sm text-gray-500">{row.open.toFixed(2)}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-sm text-gray-500">{row.high.toFixed(2)}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-sm text-gray-500">{row.low.toFixed(2)}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-sm text-gray-900 font-medium">{displayClose(row.close)}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-sm">
                              {row.closeIssue ? (
                                <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-amber-100 text-amber-900">
                                  {row.closeIssue}
                                </span>
                              ) : row.isOutlier ? (
                                <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-red-100 text-red-800">
                                  Outlier
                                </span>
                              ) : (
                                <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                                  Clean
                                </span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {processedData.length > 100 && (
                      <div className="p-4 text-center text-xs text-gray-500 border-t border-gray-200">
                        Showing first 100 of {processedData.length} rows. Export to see full dataset.
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
