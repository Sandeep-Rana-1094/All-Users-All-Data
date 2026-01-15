
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { 
  LayoutDashboard, 
  Search, 
  Calendar, 
  Filter, 
  CheckCircle2, 
  Clock, 
  AlertCircle, 
  RefreshCcw,
  FileText,
  ChevronLeft,
  ChevronRight,
  Percent,
  Timer,
  Download,
  RotateCcw,
  Activity,
  User,
  Cpu,
  ArrowUpRight,
  ChevronDown,
  ChevronUp, // Added ChevronUp for sorting
  X,
  LayoutGrid
} from 'lucide-react';

// --- Types & Constants ---

interface Task {
  id: string;
  task: string;
  planned: Date | null;
  actual: Date | null;
  systemType: string;
  name: string;
  status: 'Completed' | 'Delayed' | 'Pending';
  delayInHours: number;
  // progress: number; // Removed progress field
}

const GOOGLE_SHEET_ID = '1XqXgsTYWv6OAL2TmeRzAPZX9jajyJOGjNmgmxqNhIkA';
const SHEET_NAME = 'All Combined Data';
const CSV_URL = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(SHEET_NAME)}`;

// --- Helper Functions ---

const parseFlexibleDate = (dateStr: string): Date | null => {
  if (!dateStr || dateStr.trim() === '' || dateStr === '—' || dateStr === '-') return null;
  const cleanStr = dateStr.replace(/"/g, '').trim();
  const dmyMatch = cleanStr.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(?:[\s,]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmyMatch) {
    const [_, d, m, y, hh, mm, ss] = dmyMatch;
    const date = new Date(
      parseInt(y), 
      parseInt(m) - 1, 
      parseInt(d), 
      hh ? parseInt(hh) : 0, 
      mm ? parseInt(mm) : 0, 
      ss ? parseInt(ss) : 0
    );
    if (!isNaN(date.getTime())) return date;
  }
  const fallback = new Date(cleanStr);
  return isNaN(fallback.getTime()) ? null : fallback;
};

const parseCSV = (text: string): any[] => {
  const result = [];
  let row: string[] = [];
  let cell = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];
    if (inQuotes) {
      if (char === '"' && nextChar === '"') { cell += '"'; i++; }
      else if (char === '"') inQuotes = false;
      else cell += char;
    } else {
      if (char === '"') inQuotes = true;
      else if (char === ',') { row.push(cell.trim()); cell = ''; }
      else if (char === '\n' || char === '\r') {
        row.push(cell.trim());
        if (row.some(c => c !== '')) result.push(row);
        row = []; cell = '';
        if (char === '\r' && nextChar === '\n') i++; 
      } else cell += char;
    }
  }
  if (cell || row.length > 0) {
    row.push(cell.trim());
    if (row.some(c => c !== '')) result.push(row);
  }
  const headers = result[0] || [];
  const data = [];
  for (let i = 1; i < result.length; i++) {
    const rowValues = result[i];
    const obj: any = {};
    headers.forEach((header, index) => {
      const val = rowValues[index] || '';
      obj[header] = val;
      obj[`col_${index}`] = val;
    });
    data.push(obj);
  }
  return data;
};

const formatDate = (date: Date | null): string => {
  if (!date || isNaN(date.getTime())) return '—';
  const day = date.getDate().toString().padStart(2, '0');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[date.getMonth()];
  const year = date.getFullYear();
  const time = date.toTimeString().split(' ')[0].substring(0, 5);
  return `${day} ${month} ${year}, ${time}`;
};

interface ParsedSearchQuery {
  exactPhrases: string[];
  keywords: string[];
}

const parseSearchQuery = (query: string): ParsedSearchQuery => {
  const exactPhrases: string[] = [];
  let keywords: string[]; 

  let remainingQuery = query;

  // Extract quoted phrases
  const quotedRegex = /"([^"]*)"/g;
  let match;
  // Use a temporary string to store matches and replace them to avoid re-matching
  let tempRemainingQuery = remainingQuery;
  while ((match = quotedRegex.exec(tempRemainingQuery)) !== null) {
    const phrase = match[1].trim();
    if (phrase) {
      exactPhrases.push(phrase.toLowerCase());
    }
  }
  // Remove quoted parts from the original query to process remaining keywords
  remainingQuery = remainingQuery.replace(quotedRegex, '').trim();

  // Split remaining query by commas and handle individual keywords
  keywords = remainingQuery.split(',').map(part => part.trim().toLowerCase()).filter(Boolean);

  return { exactPhrases, keywords };
};

// --- Sub-Components ---

const StatusBadge = ({ status }: { status: Task['status'] }) => {
  const config = {
    Completed: { icon: CheckCircle2, class: 'status-completed' },
    Delayed: { icon: Clock, class: 'status-delayed' },
    Pending: { icon: Timer, class: 'status-pending' }
  };
  const { icon: Icon, class: className } = config[status];
  return (
    <span className={`status-pill ${className}`}>
      <Icon size={12} strokeWidth={2.5} />
      {status}
    </span>
  );
};

const KPICard = ({ title, value, icon: Icon, color, subValue, trend, delay, onClick, active, isFilterable }: { 
  title: string; value: string | number; icon: any; color: string; subValue?: string; trend?: 'up' | 'down' | string; delay?: string; onClick?: () => void; active?: boolean; isFilterable?: boolean 
}) => {
  const colorMap: Record<string, string> = {
    indigo: 'indigo',
    emerald: 'emerald',
    amber: 'amber',
    rose: 'rose',
    slate: 'slate'
  };
  const accentColor = colorMap[color] || color;

  return (
    <div 
      onClick={onClick}
      className={`glass-panel glass-card-hover p-7 rounded-[2.5rem] border cursor-pointer group transition-all duration-500 animate-in ${delay} 
        ${isFilterable ? 'animate-attention-pulse' : ''} 
        ${active ? 'ring-2 ring-indigo-500 ring-offset-4 ring-offset-white scale-[1.02]' : ''}`}
    >
      <div className="flex items-start justify-between h-full">
        <div className="flex flex-col justify-between h-full space-y-4">
          <div>
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.2em] mb-1.5">{title}</p>
            <h3 className="text-4xl font-black text-slate-900 tracking-tighter tabular-nums">{value}</h3>
          </div>
          {subValue && (
            <div className="pt-1">
              <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-xl text-[10px] font-black uppercase tracking-widest bg-slate-50 text-slate-500 border border-slate-100 group-hover:bg-white transition-colors`}>
                {trend === 'up' && <ArrowUpRight size={10} strokeWidth={3} />}
                {subValue}
              </span>
            </div>
          )}
        </div>
        <div className={`w-14 h-14 rounded-2xl flex items-center justify-center bg-${accentColor}-50/50 border border-${accentColor}-100/50 group-hover:rotate-6 transition-transform shadow-sm`}>
          <Icon size={22} className={`text-${accentColor}-600`} strokeWidth={2} />
        </div>
      </div>
    </div>
  );
};

const TaskDetailModal = ({ task, onClose }: { task: Task; onClose: () => void }) => {
  return (
    <div 
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in"
      role="dialog"
      aria-modal="true"
      aria-labelledby="task-detail-title"
    >
      <div className="glass-panel relative w-full max-w-2xl rounded-[2.5rem] p-8 md:p-12 border">
        <button 
          onClick={onClose} 
          className="absolute top-6 right-6 p-2 rounded-full bg-slate-50 text-slate-500 hover:bg-slate-100 transition-colors shadow-sm"
          aria-label="Close task details"
        >
          <X size={20} />
        </button>

        <h3 id="task-detail-title" className="text-3xl font-black text-slate-900 mb-6 tracking-tight">Task Details</h3>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-x-8 gap-y-6 text-sm">
          <div className="detail-item">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Task ID</p>
            <p className="font-bold text-slate-800 bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-100 inline-block text-xs font-mono">{task.id}</p>
          </div>
          <div className="detail-item">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Status</p>
            <StatusBadge status={task.status} />
          </div>
          
          <div className="detail-item md:col-span-2">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Task Description</p>
            <p className="font-semibold text-slate-700 leading-relaxed">{task.task}</p>
          </div>

          <div className="detail-item">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Doer</p>
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 bg-white border border-slate-100 rounded-xl flex items-center justify-center text-slate-500 font-black text-[10px]" aria-hidden="true">
                {task.name.substring(0, 2).toUpperCase()}
              </div>
              <span className="text-sm font-semibold text-slate-700">{task.name}</span>
            </div>
          </div>
          <div className="detail-item">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">System Type</p>
            <p className="font-semibold text-slate-700">{task.systemType}</p>
          </div>

          <div className="detail-item">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Planned Date</p>
            <p className="font-semibold text-slate-700">{formatDate(task.planned)}</p>
          </div>
          <div className="detail-item">
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Actual Date</p>
            <p className="font-semibold text-slate-700">{formatDate(task.actual)}</p>
          </div>
          
          {task.delayInHours > 0 && (
            <div className="detail-item md:col-span-2">
              <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.15em] mb-1">Delay</p>
              <p className="font-semibold text-amber-600">{task.delayInHours.toFixed(1)} hours</p>
            </div>
          )}

          {/* Progress bar removed */}
        </div>
      </div>
    </div>
  );
};


// --- Column Definitions ---
interface ColumnDefinition {
  id: keyof Task | 'taskWithDot'; // 'taskWithDot' is a custom ID for rendering
  displayName: string;
  minWidth?: string;
  defaultVisible: boolean;
  render: (task: Task) => React.ReactNode;
}

const allColumnDefinitions: ColumnDefinition[] = [
  {
    id: 'id',
    displayName: 'Task ID',
    minWidth: '100px',
    defaultVisible: true,
    render: (task) => (
      <span className="text-[10px] font-black text-slate-400 font-mono tracking-wider bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-100">
        {task.id}
      </span>
    ),
  },
  {
    id: 'taskWithDot', // Custom ID for combined task + dot display
    displayName: 'Task',
    minWidth: '250px',
    defaultVisible: true,
    render: (task) => (
      <div className="flex items-center gap-2">
        <span 
          className={`w-2 h-2 rounded-full ${task.actual !== null ? 'bg-emerald-500' : 'bg-slate-300'}`}
          title={task.actual !== null ? "Task Completed" : "Task Not Yet Completed"}
          aria-hidden="true"
        ></span>
        <p className="text-sm font-bold text-slate-800 line-clamp-1">{task.task}</p>
      </div>
    ),
  },
  {
    id: 'systemType',
    displayName: 'System Type',
    minWidth: '120px',
    defaultVisible: true,
    render: (task) => (
      <p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest">{task.systemType}</p>
    ),
  },
  {
    id: 'name',
    displayName: 'Doer',
    minWidth: '150px',
    defaultVisible: true,
    render: (task) => (
      <div className="flex items-center gap-3">
        <div className="w-8 h-8 bg-white border border-slate-100 rounded-xl flex items-center justify-center text-slate-500 font-black text-[10px]" aria-hidden="true">
          {task.name.substring(0, 2).toUpperCase()}
        </div>
        <span className="text-sm font-semibold text-slate-700">{task.name}</span>
      </div>
    ),
  },
  {
    id: 'planned',
    displayName: 'Planned Date',
    minWidth: '150px',
    defaultVisible: true,
    render: (task) => (
      <span className="text-[12px] font-bold text-slate-500 tabular-nums">{formatDate(task.planned)}</span>
    ),
  },
  {
    id: 'actual',
    displayName: 'Actual Date',
    minWidth: '150px',
    defaultVisible: true,
    render: (task) => (
      <span className="text-[12px] font-bold text-slate-500 tabular-nums">{formatDate(task.actual)}</span>
    ),
  },
  {
    id: 'status',
    displayName: 'Status',
    minWidth: '120px',
    defaultVisible: true,
    render: (task) => <StatusBadge status={task.status} />,
  },
];


// --- App Component ---

const App = () => {
  const [data, setData] = useState<Task[]>([]);
  const [loading, setLoading] = useState(true); // For initial app load
  const [isRefreshing, setIsRefreshing] = useState(false); // For background fetches
  const [error, setError] = useState<string | null>(null);
  const [lastSynced, setLastSynced] = useState<Date>(new Date());
  
  const [selectedName, setSelectedName] = useState('All');
  const [selectedSystem, setSelectedSystem] = useState('All');
  const [dateRange, setDateRange] = useState({ start: '', end: '' });
  const [showDelayedOnly, setShowDelayedOnly] = useState(false);
  const [showNotDoneOnly, setShowNotDoneOnly] = useState(false);
  const [searchQuery, setSearchQuery] = useState(''); // New state for search query

  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(100);

  const [selectedTask, setSelectedTask] = useState<Task | null>(null); // State for selected task for modal

  // Column visibility state and persistence
  const [visibleColumns, setVisibleColumns] = useState<Record<string, boolean>>(() => {
    try {
      const storedVisibility = localStorage.getItem('tableColumnVisibility');
      if (storedVisibility) {
        return JSON.parse(storedVisibility);
      }
    } catch (e) {
      console.error("Failed to parse stored column visibility", e);
    }
    // Default visibility if nothing is stored or parsing fails
    return allColumnDefinitions.reduce((acc, col) => ({ ...acc, [col.id]: col.defaultVisible }), {});
  });
  const [showColumnDropdown, setShowColumnDropdown] = useState(false);
  const columnDropdownRef = useRef<HTMLDivElement>(null);

  // Sorting state
  const [sortColumn, setSortColumn] = useState<keyof Task | 'taskWithDot' | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  useEffect(() => {
    localStorage.setItem('tableColumnVisibility', JSON.stringify(visibleColumns));
  }, [visibleColumns]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (columnDropdownRef.current && !columnDropdownRef.current.contains(event.target as Node)) {
        setShowColumnDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const toggleColumnVisibility = (columnId: string) => {
    setVisibleColumns(prev => ({
      ...prev,
      [columnId]: !prev[columnId],
    }));
  };

  const resetColumnVisibility = () => {
    const defaultVisibility = allColumnDefinitions.reduce((acc, col) => ({ ...acc, [col.id]: col.defaultVisible }), {});
    setVisibleColumns(defaultVisibility);
  };

  const visibleColumnDefinitions = useMemo(() => {
    return allColumnDefinitions.filter(col => visibleColumns[col.id]);
  }, [visibleColumns]);


  const fetchData = async (initialLoad = false) => {
    try {
      if (initialLoad) {
        setLoading(true);
      } else {
        setIsRefreshing(true);
      }
      const response = await fetch(CSV_URL);
      if (!response.ok) throw new Error('Network response was not ok');
      const csvText = await response.text();
      const rawData = parseCSV(csvText);

      const now = new Date();
      const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());

      const mappedData: Task[] = rawData
        .filter(row => {
          const id = row['col_0'] || row['Unique Id'] || row['Task ID'];
          const task = row['col_1'] || row['Task'];
          return id && id.trim() !== '' && task && task.trim() !== '';
        })
        .map((row, idx) => {
          const findCol = (keys: string[], fallbackIdx: number, excludeKeywords: string[] = []) => {
            const allKeys = Object.keys(row);
            const exactKey = allKeys.find(k => keys.some(key => k.toLowerCase() === key.toLowerCase()));
            if (exactKey) return row[exactKey];
            const partialKey = allKeys.find(k => 
              keys.some(key => {
                const lowerK = k.toLowerCase();
                const isExcluded = excludeKeywords.some(ex => lowerK.includes(ex.toLowerCase()));
                return lowerK.includes(key.toLowerCase()) && !isExcluded;
              })
            );
            return row[partialKey || `col_${fallbackIdx}`];
          };

          const idStr = findCol(['Unique Id', 'Task ID'], 0);
          const taskStr = findCol(['Task'], 1, ['id', 'unique']); 
          const plannedStr = findCol(['Planned'], 2);
          const actualStr = findCol(['Actual'], 3);
          const nameStr = row['col_9'] || findCol(['Final Name', 'Name'], 9); 
          const systemStr = findCol(['System type', 'System'], 7);

          const plannedDate = parseFlexibleDate(plannedStr);
          const actualDate = parseFlexibleDate(actualStr);
          
          let status: 'Completed' | 'Delayed' | 'Pending' = 'Pending';
          let delay = 0;
          const isDone = actualDate !== null;
          const isPlannedValid = plannedDate !== null;

          if (isPlannedValid) {
            const pDayStart = new Date(plannedDate!.getFullYear(), plannedDate!.getMonth(), plannedDate!.getDate());
            if (isDone) {
              const aDayStart = new Date(actualDate!.getFullYear(), actualDate!.getMonth(), actualDate!.getDate());
              status = aDayStart <= pDayStart ? 'Completed' : 'Delayed';
              delay = Math.max(0, (actualDate!.getTime() - plannedDate!.getTime()) / (1000 * 60 * 60));
            } else {
              status = pDayStart < todayStart ? 'Delayed' : 'Pending';
            }
          } else if (isDone) {
            status = 'Completed';
          }

          // Progress calculation removed
          // const progress = isDone ? 100 : 0;

          return {
            id: idStr || `T-${idx}`,
            task: taskStr || 'No Task Description',
            planned: plannedDate,
            actual: actualDate,
            systemType: systemStr || 'General',
            name: nameStr || 'Unassigned',
            status,
            delayInHours: delay,
            // progress, // Progress assignment removed
          };
        });

      setData(mappedData);
      setLastSynced(new Date());
      setError(null); // Clear any previous error
    } catch (err) {
      console.error(err);
      setError('Failed to fetch dashboard data. Ensure the spreadsheet is shared correctly.');
    } finally {
      if (initialLoad) {
        setLoading(false);
      } else {
        setIsRefreshing(false);
      }
    }
  };

  useEffect(() => {
    fetchData(true); // Initial fetch with full loading screen
    const refreshInterval = setInterval(() => {
      fetchData(false); // Subsequent fetches, uses subtle refresh indicator
    }, 60000); // 60 seconds

    return () => clearInterval(refreshInterval); // Cleanup on component unmount
  }, []); // Empty dependency array means this effect runs once on mount and cleans up on unmount

  const resetFilters = () => {
    setSelectedName('All');
    setSelectedSystem('All');
    setDateRange({ start: '', end: '' });
    setShowDelayedOnly(false);
    setShowNotDoneOnly(false);
    setSearchQuery(''); // Reset search query
    setCurrentPage(1); // Reset pagination to first page when filters are cleared
  };

  const hasActiveFilters = useMemo(() => {
    return selectedName !== 'All' || 
           selectedSystem !== 'All' || 
           dateRange.start !== '' || 
           dateRange.end !== '' || 
           showDelayedOnly || 
           showNotDoneOnly ||
           searchQuery !== ''; // Added searchQuery to active filters check
  }, [selectedName, selectedSystem, dateRange, showDelayedOnly, showNotDoneOnly, searchQuery]);

  const activeFilterCount = useMemo(() => {
    let count = 0;
    if (selectedName !== 'All') count++;
    if (selectedSystem !== 'All') count++;
    if (dateRange.start !== '') count++;
    if (dateRange.end !== '') count++;
    if (showDelayedOnly) count++;
    if (showNotDoneOnly) count++;
    if (searchQuery !== '') count++;
    return count;
  }, [selectedName, selectedSystem, dateRange, showDelayedOnly, showNotDoneOnly, searchQuery]);

  const filteredData = useMemo(() => {
    const { exactPhrases, keywords } = parseSearchQuery(searchQuery);

    return data.filter(item => {
      const matchesName = selectedName === 'All' || item.name === selectedName;
      const matchesSystem = selectedSystem === 'All' || item.systemType === selectedSystem;
      let matchesDate = true;
      if (dateRange.start && item.planned) {
        matchesDate = item.planned >= new Date(dateRange.start);
      }
      if (dateRange.end && item.planned && matchesDate) {
        const endDate = new Date(dateRange.end);
        endDate.setHours(23, 59, 59);
        matchesDate = item.planned <= endDate;
      }
      const matchesDelayed = !showDelayedOnly || item.status === 'Delayed';
      const matchesNotDone = !showNotDoneOnly || item.actual === null;
      
      let matchesSearch = true;
      if (exactPhrases.length > 0 || keywords.length > 0) {
        const itemTaskLower = item.task.toLowerCase();
        const itemIdLower = item.id.toLowerCase();

        const hasExactPhraseMatch = exactPhrases.some(phrase => itemTaskLower.includes(phrase) || itemIdLower.includes(phrase));
        const hasKeywordMatch = keywords.some(keyword => itemTaskLower.includes(keyword) || itemIdLower.includes(keyword));

        matchesSearch = hasExactPhraseMatch || hasKeywordMatch;
      }

      return matchesName && matchesSystem && matchesDate && matchesDelayed && matchesNotDone && matchesSearch;
    });
  }, [data, selectedName, selectedSystem, dateRange, showDelayedOnly, showNotDoneOnly, searchQuery]);

  // Sorting logic
  const handleSort = (columnId: keyof Task | 'taskWithDot') => {
    if (sortColumn === columnId) {
      setSortDirection(prev => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortColumn(columnId);
      setSortDirection('asc');
    }
    setCurrentPage(1); // Reset to first page on new sort
  };

  const sortedData = useMemo(() => {
    if (!sortColumn) {
      return filteredData;
    }

    const collator = new Intl.Collator('en', { numeric: true, sensitivity: 'base' });

    return [...filteredData].sort((a, b) => {
      let aValue: any;
      let bValue: any;

      if (sortColumn === 'taskWithDot') {
        aValue = a.task;
        bValue = b.task;
      } else {
        aValue = a[sortColumn];
        bValue = b[sortColumn];
      }

      // Handle null/Date types for planned/actual dates
      if (sortColumn === 'planned' || sortColumn === 'actual') {
        const dateA = aValue as Date | null;
        const dateB = bValue as Date | null;

        if (dateA === null && dateB === null) return 0;
        if (dateA === null) return sortDirection === 'asc' ? 1 : -1; // nulls last
        if (dateB === null) return sortDirection === 'asc' ? -1 : 1; // nulls last

        const timeA = dateA.getTime();
        const timeB = dateB.getTime();

        return sortDirection === 'asc' ? timeA - timeB : timeB - timeA;
      }

      // Handle numbers (no 'progress' anymore, but kept for other potential number fields)
      if (typeof aValue === 'number' && typeof bValue === 'number') {
        return sortDirection === 'asc' ? aValue - bValue : bValue - aValue;
      }

      // Handle strings (case-insensitive, numeric-aware)
      const comparison = collator.compare(String(aValue || ''), String(bValue || ''));
      return sortDirection === 'asc' ? comparison : -comparison;
    });
  }, [filteredData, sortColumn, sortDirection]);

  const totalPages = Math.ceil(sortedData.length / rowsPerPage);
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return sortedData.slice(start, start + rowsPerPage);
  }, [sortedData, currentPage, rowsPerPage]);

  const stats = useMemo(() => {
    const total = filteredData.length;
    const actualDone = filteredData.filter(d => d.actual !== null).length;
    const delayedCount = filteredData.filter(d => d.status === 'Delayed').length;
    const notDoneCount = filteredData.filter(d => d.actual === null).length;
    const delayedRate = total > 0 ? (delayedCount / total) * 100 : 0;
    return { total, completed: actualDone, delayed: delayedCount, notDone: notDoneCount, delayedRate };
  }, [filteredData]);

  const uniqueNames = useMemo(() => Array.from(new Set(data.map(d => d.name).filter(Boolean))).sort(), [data]);
  const uniqueSystems = useMemo(() => Array.from(new Set(data.map(d => d.systemType).filter(Boolean))).sort(), [data]);

  const handleExport = () => {
    if (filteredData.length === 0) return;
    const headers = ['Unique Id', 'Task', 'Name', 'Planned', 'Actual', 'Status', 'System'];
    const csvRows = [
      headers.join(','),
      ...filteredData.map(t => [t.id, `"${t.task}"`, `"${t.name}"`, formatDate(t.planned), formatDate(t.actual), t.status, t.systemType].join(','))
    ];
    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'operational_report.csv'; a.click();
  };

  if (loading && data.length === 0) { // Only show full-screen loader if no data loaded yet
    return (
      <div className="h-screen w-full flex flex-col items-center justify-center">
        <div className="relative">
          <div className="w-20 h-20 border-4 border-indigo-100 border-t-indigo-500 rounded-full animate-spin"></div>
          <div className="absolute inset-0 flex items-center justify-center text-indigo-500">
            <Activity size={28} className="animate-pulse" />
          </div>
        </div>
        <div className="mt-8 text-center">
           <p className="text-slate-800 font-bold tracking-[0.2em] uppercase text-xs">Loading Data...</p>
           <p className="text-slate-400 text-[10px] mt-2">Fetching real-time data stream</p>
        </div>
      </div>
    );
  }

  if (error && data.length === 0) { // Only show full-screen error if no data was ever loaded
    return (
      <div className="h-screen w-full flex items-center justify-center p-6">
        <div className="glass-panel p-10 rounded-[2.5rem] border-rose-100 text-center max-w-lg">
          <AlertCircle className="text-rose-500 mx-auto mb-6" size={50} />
          <h2 className="text-2xl font-black text-slate-800 mb-3">Sync Interrupted</h2>
          <p className="text-slate-500 mb-8 leading-relaxed">{error}</p>
          <button onClick={() => fetchData(true)} className="w-full bg-slate-900 text-white font-bold py-4 rounded-2xl hover:bg-slate-800 transition-all">
            Retry Connection
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen p-4 lg:p-12 flex justify-center">
      <div className="w-full max-w-[1500px] flex flex-col gap-10">
        
        {/* Header Section */}
        <header className="flex flex-col xl:flex-row xl:items-center justify-between gap-8 animate-in">
          <div className="flex flex-col gap-1">
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 bg-white rounded-2xl shadow-sm border border-slate-100 flex items-center justify-center text-indigo-600">
                <LayoutDashboard size={28} strokeWidth={2.5} />
              </div>
              <h1 className="text-4xl font-black text-slate-900 tracking-tight">All Tasks System <span className="text-indigo-500/80"></span></h1>
            </div>
            <div className="flex items-center gap-2 mt-2 px-1">
              <span className="w-2 h-2 rounded-full bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]"></span>
              <p className="text-slate-500 font-bold text-[11px] uppercase tracking-widest">
                Live Data • {isRefreshing ? (
                  <span className="inline-flex items-center gap-1">
                    <Activity size={12} className="animate-spin text-indigo-500" /> Refreshing...
                  </span>
                ) : (
                  `Last Sync: ${lastSynced.toLocaleTimeString()}`
                )}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            
            {/* Export Button - Exactly as pictured */}
            <button onClick={handleExport} className="btn-export" aria-label="Export Data">
              <Download size={18} className="text-slate-600 mb-0.5" />
              <span className="text-[10px] font-bold text-slate-500 uppercase">Export</span>
            </button>

            {/* Refresh Button - Exactly as pictured */}
            <button 
              onClick={() => fetchData(false)} 
              className="btn-refresh" 
              aria-label="Refresh Data"
              disabled={isRefreshing}
            >
              <RefreshCcw size={22} strokeWidth={2.5} className={isRefreshing ? 'animate-spin' : ''} />
            </button>
          </div>
        </header>

        {/* Filters Bar */}
        <section className="glass-panel p-3 rounded-3xl flex flex-wrap items-center gap-4 border animate-in stagger-1">
          <div className="flex items-center gap-2 px-5 py-2.5 bg-slate-900 rounded-2xl text-white text-[10px] font-black uppercase tracking-widest shadow-lg shadow-slate-200">
            <Filter size={14} /> Controls
          </div>

          <div className="h-8 w-px bg-slate-200 mx-1 hidden lg:block"></div>

          {/* All Doers Filter */}
          <div className="flex items-center bg-white/50 border border-slate-200 rounded-2xl px-4 py-2 shadow-sm relative">
            <User size={16} className="text-slate-400 mr-3" />
            <select 
              className="bg-transparent appearance-none border-none outline-none text-sm font-bold text-slate-700 w-full cursor-pointer pr-6" // pr-6 for arrow space
              value={selectedName}
              onChange={(e) => setSelectedName(e.target.value)}
              aria-label="Select Doer"
            >
              <option value="All">All Doers</option>
              {uniqueNames.map(n => <option key={n} value={n}>{n}</option>)}
            </select>
            <ChevronDown size={14} className="text-slate-400 absolute right-3 pointer-events-none" />
          </div>

          {/* All Systems Data Filter */}
          <div className="flex items-center bg-white/50 border border-slate-200 rounded-2xl px-4 py-2 shadow-sm relative">
            <Cpu size={16} className="text-slate-400 mr-3" />
            <select 
              className="bg-transparent appearance-none border-none outline-none text-sm font-bold text-slate-700 w-full cursor-pointer pr-6" // pr-6 for arrow space
              value={selectedSystem}
              onChange={(e) => setSelectedSystem(e.target.value)}
              aria-label="Select System Type"
            >
              <option value="All">All Systems Data</option>
              {uniqueSystems.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            <ChevronDown size={14} className="text-slate-400 absolute right-3 pointer-events-none" />
          </div>

          {/* Date Range Filter */}
          <div className="flex items-center bg-white/50 border border-slate-200 rounded-2xl px-4 py-2 shadow-sm">
            <Calendar size={14} className="text-slate-400 mr-3" />
            <input type="date" className="bg-transparent border-none text-xs font-bold outline-none cursor-pointer text-slate-700" value={dateRange.start} onChange={(e) => setDateRange(p => ({...p, start: e.target.value}))} aria-label="Start Date"/>
            <span className="mx-3 text-slate-300 text-[10px] font-black">TO</span>
            <input type="date" className="bg-transparent border-none text-xs font-bold outline-none cursor-pointer text-slate-700" value={dateRange.end} onChange={(e) => setDateRange(p => ({...p, end: e.target.value}))} aria-label="End Date"/>
          </div>

          {/* NEW Search Input */}
          <div className="flex items-center flex-grow bg-white/50 border border-slate-200 rounded-2xl px-4 py-2 shadow-sm max-w-sm"> {/* Added flex-grow and max-w-sm for better layout */}
            <Search size={16} className="text-slate-400 mr-3" />
            <input
              type="text"
              placeholder="Search tasks by ID or keyword..."
              className="bg-transparent border-none outline-none text-sm font-bold text-slate-700 w-full"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              aria-label="Search tasks by ID or keyword"
            />
            {searchQuery && (
              <button 
                onClick={() => setSearchQuery('')} 
                className="text-slate-400 hover:text-rose-500 transition-colors p-1"
                aria-label="Clear search"
              >
                <X size={16} />
              </button>
            )}
          </div>

          {hasActiveFilters && (
            <button onClick={resetFilters} className="p-3 text-rose-500 bg-rose-50 rounded-2xl hover:bg-rose-100 transition-all border border-rose-100" title="Clear Filters" aria-label="Clear All Filters">
              <RotateCcw size={18} strokeWidth={3} />
            </button>
          )}
        </section>

        {/* Stats Row */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 animate-in stagger-2">
          <KPICard 
            title="Total Planned" 
            value={stats.total.toLocaleString()} 
            icon={FileText} 
            color="indigo" 
            subValue="All Tasks in Scope"
            delay="stagger-1"
          />
          <KPICard 
            title="Delayed" 
            value={stats.delayed.toLocaleString()} 
            icon={Clock} 
            color="amber" 
            subValue="Tasks Exceeding Deadline"
            delay="stagger-2"
            onClick={() => setShowDelayedOnly(!showDelayedOnly)}
            active={showDelayedOnly}
            isFilterable={true}
          />
          <KPICard 
            title="Remaining" 
            value={stats.notDone.toLocaleString()} 
            icon={Timer} 
            color="rose" 
            subValue="Tasks Yet to Complete"
            delay="stagger-3"
            onClick={() => setShowNotDoneOnly(!showNotDoneOnly)}
            active={showNotDoneOnly}
            isFilterable={true}
          />
          <KPICard 
            title="Negative Score %" 
            value={`${stats.delayedRate.toFixed(1)}%`} 
            icon={Percent} 
            color="slate" 
            subValue="Overall Delay Impact"
            delay="stagger-4"
          />
        </section>

        {/* Table Section */}
        <section className="glass-panel rounded-[3rem] overflow-hidden border animate-in stagger-3">
          <div className="p-8 border-b border-slate-100/60 flex items-center justify-between">
            <div className="flex items-center gap-4">
               <h4 className="text-xl font-black text-slate-900 tracking-tight">ALL Data Table</h4> {/* Renamed table title */}
               {activeFilterCount > 0 && (
                 <span className="px-3 py-1 bg-rose-50 text-rose-600 rounded-xl text-[10px] font-black uppercase tracking-[0.2em] border border-rose-100 animate-in stagger-1">
                   {activeFilterCount} filters active
                 </span>
               )}
            </div>
            <div className="flex items-center gap-4">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Stream Volume</span>
              <select 
                className="bg-white border border-slate-100 rounded-xl px-4 py-1.5 text-xs font-bold text-slate-700 outline-none cursor-pointer shadow-sm"
                value={rowsPerPage}
                onChange={(e) => { setRowsPerPage(Number(e.target.value)); setCurrentPage(1); }}
                aria-label="Rows per page"
              >
                {[50, 100, 250, 500].map(v => <option key={v} value={v}>{v}</option>)}
              </select>

              {/* Column Visibility Dropdown */}
              <div className="relative" ref={columnDropdownRef}>
                <button 
                  onClick={() => setShowColumnDropdown(prev => !prev)} 
                  className="flex items-center gap-2 bg-white border border-slate-100 rounded-xl px-4 py-1.5 text-xs font-bold text-slate-700 outline-none cursor-pointer shadow-sm hover:bg-slate-50 transition-colors"
                  aria-haspopup="true"
                  aria-expanded={showColumnDropdown}
                  aria-label="Toggle column visibility"
                >
                  <LayoutGrid size={14} className="text-slate-400" />
                  Columns
                  <ChevronDown size={12} className={`text-slate-400 transition-transform ${showColumnDropdown ? 'rotate-180' : ''}`} />
                </button>
                {showColumnDropdown && (
                  <div className="absolute right-0 mt-2 w-48 bg-white border border-slate-200 rounded-xl shadow-lg z-20 animate-in fade-in slide-in-from-top-1">
                    <ul className="py-2 text-sm text-slate-700">
                      {allColumnDefinitions.map(col => (
                        <li key={col.id} className="px-4 py-2 hover:bg-slate-50 cursor-pointer flex items-center justify-between">
                          <label className="flex items-center cursor-pointer flex-grow">
                            <input
                              type="checkbox"
                              checked={!!visibleColumns[col.id]}
                              onChange={() => toggleColumnVisibility(col.id)}
                              className="form-checkbox h-4 w-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500"
                            />
                            <span className="ml-3 text-sm font-medium">{col.displayName}</span>
                          </label>
                        </li>
                      ))}
                      <li className="border-t border-slate-100 mt-2 pt-2">
                        <button 
                          onClick={resetColumnVisibility}
                          className="w-full text-left px-4 py-2 text-rose-600 hover:bg-rose-50 flex items-center gap-2"
                        >
                          <RotateCcw size={14} /> Reset to Default
                        </button>
                      </li>
                    </ul>
                  </div>
                )}
              </div>
            </div>
          </div>

          <div className="overflow-x-auto custom-scrollbar max-h-[700px]">
            <table className="w-full text-left modern-table">
              <thead className="sticky top-0 bg-white/90 backdrop-blur-md z-10 border-b border-slate-100">
                <tr>
                  {visibleColumnDefinitions.map(col => (
                    <th
                      key={col.id}
                      className={`cursor-pointer group ${col.id === 'id' || col.id === 'status' ? 'px-10' : 'px-8'}`}
                      style={{ minWidth: col.minWidth }}
                      onClick={() => handleSort(col.id)}
                      aria-sort={sortColumn === col.id ? (sortDirection === 'asc' ? 'ascending' : 'descending') : 'none'}
                    >
                      <div className="flex items-center gap-1">
                        {col.displayName}
                        {sortColumn === col.id && (
                          sortDirection === 'asc' ? (
                            <ChevronUp size={14} className="text-slate-500 transition-transform group-hover:text-slate-700" />
                          ) : (
                            <ChevronDown size={14} className="text-slate-500 transition-transform group-hover:text-slate-700" />
                          )
                        )}
                        {sortColumn !== col.id && (
                          <ChevronDown size={14} className="text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity" />
                        )}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100/30">
                {paginatedData.length > 0 ? (
                  paginatedData.map((task, idx) => (
                    <tr 
                      key={task.id + idx} 
                      className={`${task.status === 'Delayed' ? 'table-row-delayed' : ''} cursor-pointer`}
                      onClick={() => setSelectedTask(task)} // Open modal on row click
                      tabIndex={0} // Make row focusable
                      role="button" // Indicate it's an interactive element
                      aria-label={`View details for task ${task.id}`}
                    >
                      {visibleColumnDefinitions.map(col => (
                        <td key={col.id} className={`${col.id === 'id' || col.id === 'status' ? 'px-10' : 'px-8'} py-6`}>
                          {col.render(task)}
                        </td>
                      ))}
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={visibleColumnDefinitions.length} className="py-40 text-center">
                      <div className="flex flex-col items-center justify-center opacity-40">
                        <Search size={64} className="mb-6 text-slate-300" strokeWidth={1} aria-hidden="true" />
                        <h5 className="text-xl font-black text-slate-400 uppercase tracking-[0.2em]">Zero Nodes Located</h5>
                        <p className="text-slate-400 text-sm mt-2 font-medium">Verify your filter parameters and retry search</p>
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="p-8 bg-slate-50/40 flex items-center justify-between border-t border-slate-100">
            <div className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.15em]">
              Showing <span className="text-slate-900">{((currentPage - 1) * rowsPerPage + 1).toLocaleString()}</span> - 
              <span className="text-slate-900">{Math.min(currentPage * rowsPerPage, sortedData.length).toLocaleString()}</span> of 
              <span className="text-slate-900"> {sortedData.length.toLocaleString()}</span> entries
            </div>
            
            <div className="flex items-center gap-4">
              <button 
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-3 rounded-2xl border border-slate-200 bg-white text-slate-400 hover:text-indigo-600 hover:border-indigo-200 disabled:opacity-20 transition-all shadow-sm active:scale-95"
                aria-label="Previous page"
              >
                <ChevronLeft size={20} strokeWidth={3} />
              </button>
              
              <div className="flex items-center gap-2">
                {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                  let pageNum = i + 1;
                  if (totalPages > 5 && currentPage > 3) pageNum = currentPage - 2 + i;
                  if (pageNum > totalPages) pageNum = totalPages - 4 + i; // Ensure page numbers stay within bounds near the end
                  if (pageNum < 1) return null; // Ensure page numbers don't go below 1
                  if (pageNum > totalPages) return null;
                  
                  return (
                    <button 
                      key={pageNum}
                      onClick={() => setCurrentPage(pageNum)}
                      className={`w-10 h-10 rounded-2xl text-[11px] font-black transition-all ${currentPage === pageNum ? 'bg-slate-900 text-white shadow-xl shadow-slate-200' : 'bg-white border border-slate-100 text-slate-400 hover:bg-slate-50'}`}
                      aria-current={currentPage === pageNum ? "page" : undefined}
                      aria-label={`Page ${pageNum}`}
                    >
                      {pageNum}
                    </button>
                  );
                })}
              </div>

              <button 
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages || totalPages === 0}
                className="p-3 rounded-2xl border border-slate-200 bg-white text-slate-400 hover:text-indigo-600 hover:border-indigo-200 disabled:opacity-20 transition-all shadow-sm active:scale-95"
                aria-label="Next page"
              >
                <ChevronRight size={20} strokeWidth={3} />
              </button>
            </div>
          </div>
        </section>

        <footer className="mt-4 flex flex-col md:flex-row justify-between items-center gap-6 text-[10px] font-black text-slate-400 tracking-[0.2em] uppercase px-4 pb-16">
           <div className="flex items-center gap-4">
              <span className="px-3 py-1.5 bg-white rounded-xl border border-slate-100 shadow-sm text-slate-500">Node: {SHEET_NAME}</span>
           </div>
           <div className="flex items-center gap-8">
              <span className="flex items-center gap-3">
                <span className="w-2 h-2 rounded-full bg-emerald-500 shadow-lg shadow-emerald-200"></span> Nominal Ops
              </span>
              <span className="flex items-center gap-3">
                <span className="w-2 h-2 rounded-full bg-amber-500 shadow-lg shadow-amber-200"></span> Drift Detected
              </span>
           </div>
        </footer>

        {selectedTask && <TaskDetailModal task={selectedTask} onClose={() => setSelectedTask(null)} />}
      </div>
    </div>
  );
};

const root = createRoot(document.getElementById('root')!);
root.render(<App />);
