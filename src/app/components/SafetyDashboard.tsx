import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Activity,
  AlertTriangle,
  Calendar as CalendarIcon,
  CheckCircle2,
  Edit,
  Flame,
  Image as ImageIcon,
  Lock,
  Plus,
  Save,
  Shield,
  Target,
  Trash2,
  Unlock,
  Upload,
} from 'lucide-react';

import nhkLogo from '@/assets/nhk-logo.png';

import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from '@/app/components/ui/resizable';

import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/app/components/ui/dialog';

import { ScrollArea } from '@/app/components/ui/scroll-area';

type DayStatus = 'safe' | 'near_miss' | 'accident' | null;
interface DailyStatistic { day: number; status: DayStatus }
interface MonthlyData { month: number; year: number; days: DailyStatistic[] }
interface Announcement { id: string; text: string }
interface SafetyMetric { id: string; label: string; value: string; unit?: string }

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DAY_HEADERS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

// Keys are ISO dates: YYYY-MM-DD. Values are holiday notes.
const HOLIDAY_NOTES_2026: Record<string, string> = {
  '2026-01-01': "NEW YEAR'S DAY",
  '2026-01-02': "NEW YEAR'S DAY",
  '2026-02-22': 'LABOUR UNION MEETING',
  '2026-03-02': 'วันหยุดพิเศษ',
  '2026-03-03': 'MAKHABUCHA DAY',
  '2026-04-13': 'SONGKRAN DAY',
  '2026-04-14': 'SONGKRAN DAY',
  '2026-04-15': 'SONGKRAN DAY',
  '2026-04-16': 'วันหยุดพิเศษ',
  '2026-04-17': 'วันหยุดพิเศษ',
  '2026-05-01': 'NATIONAL LABOUR DAY',
  '2026-05-04': 'CORONATION DAY',
  '2026-06-01': 'SUBSTITUTE VISAKHABUCHA DAY',
  '2026-06-02': 'วันหยุดพิเศษ',
  '2026-06-03': "H.M. THE QUEEN'S BIRTHDAY",
  '2026-07-27': 'วันหยุดพิเศษ',
  '2026-07-28': "H.M. THE KING'S BIRTHDAY",
  '2026-07-29': 'A-SARNHA BUCHA DAY',
  '2026-08-12': "MOTHER'S DAY",
  '2026-10-12': 'วันหยุดพิเศษ',
  '2026-10-13': 'RAMA IX MEMORIAL DAY',
  '2026-10-23': 'KING CHULALONGKORN DAY',
  '2026-12-07': 'SUBSTITUTE NATION DAY',
  '2026-12-20': 'Party',
  '2026-12-28': 'วันหยุดพิเศษ',
  '2026-12-29': 'วันหยุดพิเศษ',
  '2026-12-30': 'วันหยุดพิเศษ',
  '2026-12-31': "NEW YEAR'S EVE",
};

const AUTO_SAFE_HOUR = 16;

const DEFAULT_ANNOUNCEMENTS: Announcement[] = [
  { id: '1', text: 'PPE Audit ประจำสัปดาห์ทุกวันพฤหัสบดี เวลา 09:00 น.' },
  { id: '2', text: 'Emergency Drill ไตรมาสนี้กำหนดวันที่ 28 มีนาคม 2026' },
];

const DEFAULT_METRICS: SafetyMetric[] = [
  { id: 'm1', label: 'First Aid', value: '0', unit: 'case' },
  { id: 'm2', label: 'Non-Absent', value: '0', unit: 'case' },
  { id: 'm3', label: 'Absent', value: '0', unit: 'case' },
  { id: 'm4', label: 'Fire', value: '0', unit: 'case' },
  { id: 'm5', label: 'IFR', value: '0', unit: '' },
  { id: 'm6', label: 'ISR', value: '1.2', unit: '' },
];

const DEFAULT_POSTER_TOP = '/company-policy-poster.png';
const DEFAULT_POSTER_BOTTOM = '/safety-culture.png';
const DEFAULT_POLICY_IMAGES: string[] = ['/policy-vp.png'];

type MetricToneKey = 'firstAid' | 'nonAbsent' | 'absent' | 'fire' | 'neutral';

function metricToneKey(label: string): MetricToneKey {
  const l = label.trim().toLowerCase();
  if (l === 'first aid' || l.includes('first aid')) return 'firstAid';
  if (l === 'non-absent' || l.includes('non absent') || l.includes('non-absent')) return 'nonAbsent';
  if (l === 'absent' || (l.includes('absent') && !l.includes('non'))) return 'absent';
  if (l === 'fire' || l.includes('fire')) return 'fire';
  return 'neutral';
}

function metricAccentStyle(tone: MetricToneKey) {
  const varName = tone === 'neutral' ? 'var(--tone-cyan-600)' : `var(--color-${tone})`;
  return {
    borderColor: `color-mix(in oklab, ${varName} 35%, white)`,
    background:
      tone === 'neutral'
        ? 'linear-gradient(135deg, rgba(255,255,255,0.95), rgba(236,254,255,0.9))'
        : `linear-gradient(135deg, rgba(255,255,255,0.95), color-mix(in oklab, ${varName} 18%, white))`,
  } as React.CSSProperties;
}

function clamp(n: number, min: number, max: number) { return Math.min(max, Math.max(min, n)); }
function uid(prefix='id'){ return `${prefix}-${Math.random().toString(16).slice(2)}-${Date.now()}`; }

function toISODate(d: Date) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function nextDayStatus(status: DayStatus): DayStatus {
  if (status === null) return 'safe';
  if (status === 'safe') return 'near_miss';
  if (status === 'near_miss') return 'accident';
  return null;
}

function createYearData(year:number): MonthlyData[] {
  return Array.from({length:12}, (_,m)=> ({
    month:m, year,
    days: Array.from({length:new Date(year,m+1,0).getDate()},(_,i)=>({day:i+1,status:null}))
  }));
}

function isValidMonthlyData(data: unknown, year:number): data is MonthlyData[] {
  return Array.isArray(data) && data.length===12 && data.every((m:any, idx)=>
    m && m.month===idx && m.year===year && Array.isArray(m.days) && m.days.length===new Date(year, idx+1, 0).getDate()
  );
}

function applyAutoSafe(prev: MonthlyData[], now: Date, year: number): MonthlyData[] {
  if (now.getFullYear() !== year) return prev;
  const todayStart = new Date(year, now.getMonth(), now.getDate());
  const afterCutoff = (now.getHours() > AUTO_SAFE_HOUR) || (now.getHours() === AUTO_SAFE_HOUR && now.getMinutes() >= 0);

  let next: MonthlyData[] | null = null;
  const ensureNext = () => {
    if (!next) next = prev.map((mm) => ({ ...mm, days: mm.days.map((dd) => ({ ...dd })) }));
    return next;
  };

  let changed = false;

  for (let m = 0; m < 12; m++) {
    const month = prev[m];
    if (!month) continue;
    for (let i = 0; i < month.days.length; i++) {
      const dd = month.days[i];
      if (dd.status !== null) continue;
      const dt = new Date(year, m, dd.day);
      if (dt < todayStart) {
        const tgt = ensureNext()[m].days[i];
        tgt.status = 'safe';
        changed = true;
        continue;
      }
      if (afterCutoff && m === now.getMonth() && dd.day === now.getDate()) {
        const tgt = ensureNext()[m].days[i];
        tgt.status = 'safe';
        changed = true;
      }
    }
  }

  return changed && next ? next : prev;
}

const BASE_VIEWPORT = { width: 1920, height: 1080 };
function rootFontSize(w:number,h:number){
  const scale = Math.min(w/BASE_VIEWPORT.width, h/BASE_VIEWPORT.height);
  return clamp(16 * Math.pow(Math.max(scale,0.35), 0.45), 14, 24);
}

function panelScaleFromSize(w:number,h:number){
  const ratio = Math.min(w/560, h/360);
  return clamp(Math.pow(Math.max(ratio, 0.35), 0.42), 0.68, 1.18);
}

function scaledPx(base:number, panelScale:number, min?:number, max?:number){
  const px = clamp(base * panelScale, min ?? base*0.8, max ?? base*1.35);
  return `${px/16}rem`;
}

function HeaderClock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const t = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(t);
  }, []);
  const dateLabel = now.toLocaleDateString([], { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
  const timeLabel = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  return (
    <div className="flex flex-col items-center tabular-nums leading-none">
      <div className="font-extrabold text-slate-800 header-date">{dateLabel}</div>
      <div className="font-black text-slate-900 header-time mt-1">{timeLabel}</div>
    </div>
  );
}

function Card({
  title,
  icon,
  actions,
  children,
  className='',
  tone='sky',
  panelScale=1,
}:{
  title:string;
  icon:React.ReactNode;
  actions?: React.ReactNode;
  children:React.ReactNode;
  className?:string;
  tone?: 'sky'|'amber'|'green'|'blue'|'teal';
  panelScale?: number;
}) {
  const toneMap = {
    sky: {
      outer: 'border-sky-200 bg-gradient-to-b from-sky-50/70 to-white',
      header: 'from-sky-100 via-white to-sky-50 border-sky-200',
      body: 'bg-gradient-to-b from-white to-sky-50/25',
    },
    amber: {
      outer: 'border-amber-200 bg-gradient-to-b from-amber-50/80 to-white',
      header: 'from-amber-100 via-white to-yellow-50 border-amber-200',
      body: 'bg-gradient-to-b from-white to-amber-50/20',
    },
    green: {
      outer: 'border-emerald-200 bg-gradient-to-b from-emerald-50/70 to-white',
      header: 'from-emerald-100 via-white to-lime-50 border-emerald-200',
      body: 'bg-gradient-to-b from-white to-emerald-50/20',
    },
    blue: {
      outer: 'border-blue-200 bg-gradient-to-b from-blue-50/70 to-white',
      header: 'from-blue-100 via-white to-cyan-50 border-blue-200',
      body: 'bg-gradient-to-b from-white to-blue-50/20',
    },
    teal: {
      outer: 'border-cyan-200 bg-gradient-to-b from-cyan-50/70 to-white',
      header: 'from-cyan-100 via-white to-teal-50 border-cyan-200',
      body: 'bg-gradient-to-b from-white to-cyan-50/20',
    },
  } as const;
  const toneCls = toneMap[tone];
  return (
    <section className={`rounded-2xl border shadow-sm min-h-0 flex flex-col overflow-hidden ${toneCls.outer} ${className}`}>
      <div className={`relative px-4 py-3 border-b bg-gradient-to-r ${toneCls.header} flex items-center gap-2 text-slate-800 font-semibold`}>
        <div className="h-7 w-7 rounded-lg bg-white/80 border border-white shadow-sm flex items-center justify-center shrink-0">
          {icon}
        </div>
        <h2 className="truncate font-extrabold tracking-tight" style={{ fontSize: scaledPx(17, panelScale, 14, 22) }}>{title}</h2>
        {actions ? (
          <div className="ml-auto flex items-center gap-1">
            {actions}
          </div>
        ) : null}
      </div>
      <div className={`p-3 min-h-0 flex-1 overflow-hidden flex flex-col ${toneCls.body}`} style={{ fontSize: scaledPx(14, 0.95 + (panelScale-1)*0.35, 11, 16) }}>{children}</div>
    </section>
  );
}

function PanelWrap({
  children,
}: {
  children: (panelScale: number) => React.ReactNode;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  useEffect(() => {
    const el = ref.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    let raf = 0;
    const ro = new ResizeObserver((entries) => {
      const rect = entries[0]?.contentRect;
      if (!rect) return;
      cancelAnimationFrame(raf);
      raf = window.requestAnimationFrame(() => {
        const next = panelScaleFromSize(rect.width, rect.height);
        setScale((prev) => (Math.abs(prev - next) > 0.02 ? next : prev));
      });
    });
    ro.observe(el);
    return () => {
      cancelAnimationFrame(raf);
      ro.disconnect();
    };
  }, []);
  return (
    <div ref={ref} className="h-full min-h-0">
      {children(scale)}
    </div>
  );
}

function safeEval(expr: string, vars: Record<string, number>): number | null {
  const s = (expr || '').trim();
  if (!s) return null;
  // allow only numbers, identifiers, operators and parentheses
  if (!/^[0-9a-zA-Z_+\-*/().\s]*$/.test(s)) return null;
  try {
    // eslint-disable-next-line no-new-func
    const fn = new Function('vars', `with(vars){ return (${s}); }`) as (v: Record<string, number>) => unknown;
    const out = fn(vars);
    if (typeof out === 'number' && Number.isFinite(out)) return out;
    return null;
  } catch {
    return null;
  }
}

function formatNumber(n: number | null | undefined) {
  if (n === null || n === undefined || !Number.isFinite(n)) return '-';
  return new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 }).format(n);
}

type TargetVars = {
  manpower: number;
  daysPerWeek: number;
  hoursPerDay: number;
  workingDaysYear: number;
  workingDaysSoFar: number;
};

type TargetFormulas = {
  totalExpr: string;
  toDateExpr: string;
};

const DEFAULT_TARGET_VARS: TargetVars = {
  manpower: 675,
  daysPerWeek: 6,
  hoursPerDay: 10,
  workingDaysYear: 250,
  workingDaysSoFar: 0,
};

const DEFAULT_TARGET_FORMULAS: TargetFormulas = {
  totalExpr: 'manpower * daysPerWeek * hoursPerDay * workingDaysYear',
  toDateExpr: 'manpower * daysPerWeek * hoursPerDay * workingDaysSoFar',
};

export function SafetyDashboard() {
  const now = new Date();
  const [displayMonth, setDisplayMonth] = useState(now.getMonth());
  const [currentYear, setCurrentYear] = useState(now.getFullYear());
  const [monthlyData, setMonthlyData] = useState<MonthlyData[]>(() => createYearData(now.getFullYear()));

  // Posters (left column)
  const [posterTop, setPosterTop] = useState<string | null>(DEFAULT_POSTER_TOP);
  const [posterBottom, setPosterBottom] = useState<string | null>(DEFAULT_POSTER_BOTTOM);
  const [posterTopZoom, setPosterTopZoom] = useState(1);
  const [posterBottomZoom, setPosterBottomZoom] = useState(1);

  // Safety Policy (top-middle): up to 2 images
  const [policyImages, setPolicyImages] = useState<string[]>(DEFAULT_POLICY_IMAGES);
  const [policyZoom, setPolicyZoom] = useState(1);

  // Target + metrics (center-bottom)
  const [targetVars, setTargetVars] = useState<TargetVars>(DEFAULT_TARGET_VARS);
  const [targetFormulas, setTargetFormulas] = useState<TargetFormulas>(DEFAULT_TARGET_FORMULAS);
  const [bestRecord, setBestRecord] = useState<number>(0);
  const [lossTimeAccidents, setLossTimeAccidents] = useState<number>(0);
  const [lastUpdateIso, setLastUpdateIso] = useState<string>(() => new Date().toISOString());

  const [metrics, setMetrics] = useState<SafetyMetric[]>(DEFAULT_METRICS);

  // Announcement ticker
  const [announcements, setAnnouncements] = useState<Announcement[]>(DEFAULT_ANNOUNCEMENTS);

  const [layoutLocked, setLayoutLocked] = useState(true);
  const [uiScale, setUiScale] = useState<number>(() => {
    try {
      const raw = localStorage.getItem('safety-dashboard-ui-scale');
      const v = raw ? Number(raw) : 1;
      return Number.isFinite(v) ? clamp(v, 0.8, 1.4) : 1;
    } catch {
      return 1;
    }
  });

  // dialogs
  const [editMetrics, setEditMetrics] = useState(false);
  const [metricsDraft, setMetricsDraft] = useState<SafetyMetric[]>([]);

  const [editTarget, setEditTarget] = useState(false);
  const [targetVarsDraft, setTargetVarsDraft] = useState<TargetVars>(DEFAULT_TARGET_VARS);
  const [targetFormulasDraft, setTargetFormulasDraft] = useState<TargetFormulas>(DEFAULT_TARGET_FORMULAS);
  const [bestRecordDraft, setBestRecordDraft] = useState('0');
  const [lossTimeAccidentsDraft, setLossTimeAccidentsDraft] = useState('0');

  const [editTicker, setEditTicker] = useState(false);
  const [tickerDraft, setTickerDraft] = useState('');

  // file inputs
  const posterTopInputRef = useRef<HTMLInputElement>(null);
  const posterBottomInputRef = useRef<HTMLInputElement>(null);
  const policyReplaceInputRef = useRef<HTMLInputElement>(null);
  const policyAddInputRef = useRef<HTMLInputElement>(null);

  const storageKey = `safety-dashboard-${currentYear}`;

  // Auto SAFE scheduler (16:00) + backfill
  useEffect(() => {
    const now = new Date();
    if (now.getFullYear() !== currentYear) return;

    let timer: number | undefined;

    const scheduleNext = () => {
      const current = new Date();
      const next = new Date(current);
      next.setHours(AUTO_SAFE_HOUR, 0, 0, 0);
      if (current.getTime() >= next.getTime()) next.setDate(next.getDate() + 1);

      const ms = Math.max(250, next.getTime() - current.getTime());
      timer = window.setTimeout(() => {
        const fireNow = new Date();
        setMonthlyData((prev) => applyAutoSafe(prev, fireNow, currentYear));
        scheduleNext();
      }, ms);
    };

    scheduleNext();
    return () => {
      if (timer) window.clearTimeout(timer);
    };
  }, [currentYear]);

  // Root font size scaled for TV
  useEffect(() => {
    const onResize = () => {
      const root = document.documentElement;
      const base = rootFontSize(window.innerWidth, window.innerHeight);
      root.style.setProperty('--font-size', `${base * uiScale}px`);
    };
    onResize();
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, [uiScale]);

  useEffect(() => {
    try { localStorage.setItem('safety-dashboard-ui-scale', String(uiScale)); } catch {}
  }, [uiScale]);

  // Load persisted state
  useEffect(() => {
    try {
      const raw = localStorage.getItem(storageKey);
      if (!raw) {
        setMonthlyData(applyAutoSafe(createYearData(currentYear), new Date(), currentYear));
        return;
      }
      const parsed = JSON.parse(raw);
      const loadedMonthly = isValidMonthlyData(parsed.monthlyData, currentYear) ? parsed.monthlyData : createYearData(currentYear);
      setMonthlyData(applyAutoSafe(loadedMonthly, new Date(), currentYear));

      setAnnouncements(Array.isArray(parsed.announcements) && parsed.announcements.length ? parsed.announcements : DEFAULT_ANNOUNCEMENTS);

      // Posters
      setPosterTop(typeof parsed.posterTop === 'string' ? parsed.posterTop : DEFAULT_POSTER_TOP);
      setPosterBottom(typeof parsed.posterBottom === 'string' ? parsed.posterBottom : DEFAULT_POSTER_BOTTOM);
      setPosterTopZoom(typeof parsed.posterTopZoom === 'number' ? clamp(parsed.posterTopZoom, 0.5, 2.5) : 1);
      setPosterBottomZoom(typeof parsed.posterBottomZoom === 'number' ? clamp(parsed.posterBottomZoom, 0.5, 2.5) : 1);

      // Safety Policy images
      if (Array.isArray(parsed.policyImages) && parsed.policyImages.length) {
        setPolicyImages(parsed.policyImages.slice(0, 2));
      } else {
        setPolicyImages(DEFAULT_POLICY_IMAGES);
      }
      setPolicyZoom(typeof parsed.policyZoom === 'number' ? clamp(parsed.policyZoom, 0.5, 2.5) : 1);

      // Target
      setTargetVars(parsed.targetVars && typeof parsed.targetVars === 'object' ? { ...DEFAULT_TARGET_VARS, ...parsed.targetVars } : DEFAULT_TARGET_VARS);
      setTargetFormulas(parsed.targetFormulas && typeof parsed.targetFormulas === 'object' ? { ...DEFAULT_TARGET_FORMULAS, ...parsed.targetFormulas } : DEFAULT_TARGET_FORMULAS);
      setBestRecord(Number.isFinite(Number(parsed.bestRecord)) ? Number(parsed.bestRecord) : 0);
      setLossTimeAccidents(Number.isFinite(Number(parsed.lossTimeAccidents)) ? Number(parsed.lossTimeAccidents) : 0);
      setLastUpdateIso(typeof parsed.lastUpdateIso === 'string' ? parsed.lastUpdateIso : new Date().toISOString());

      // Metrics
      setMetrics(Array.isArray(parsed.metrics) && parsed.metrics.length ? parsed.metrics : DEFAULT_METRICS);

    } catch {
      setMonthlyData(applyAutoSafe(createYearData(currentYear), new Date(), currentYear));
    }
  }, [storageKey, currentYear]);

  // Persist state (debounced)
  useEffect(() => {
    const payload = {
      monthlyData,
      announcements,
      posterTop,
      posterBottom,
      posterTopZoom,
      posterBottomZoom,
      policyImages,
      policyZoom,
      targetVars,
      targetFormulas,
      bestRecord,
      lossTimeAccidents,
      lastUpdateIso,
      metrics,
    };
    const t = window.setTimeout(() => {
      try { localStorage.setItem(storageKey, JSON.stringify(payload)); } catch {}
    }, 450);
    return () => window.clearTimeout(t);
  }, [
    storageKey,
    monthlyData,
    announcements,
    posterTop,
    posterBottom,
    posterTopZoom,
    posterBottomZoom,
    policyImages,
    policyZoom,
    targetVars,
    targetFormulas,
    bestRecord,
    lossTimeAccidents,
    lastUpdateIso,
    metrics,
  ]);

  const displayMonthData = monthlyData[displayMonth];

  const safetyStreak = useMemo(() => {
    const nowDt = new Date();
    const y = nowDt.getFullYear();
    if (y !== currentYear) return 0;

    const todayM = nowDt.getMonth();
    const todayD = nowDt.getDate();
    const todayStatus = monthlyData[todayM]?.days?.[todayD - 1]?.status ?? null;

    // If today's status isn't set yet (e.g., before 16:00), count streak up to yesterday.
    const end = (todayStatus === 'safe' || todayStatus === 'near_miss')
      ? new Date(y, todayM, todayD)
      : new Date(y, todayM, todayD - 1);

    if (end.getFullYear() !== y) return 0;

    let streak = 0;
    for (let dt = new Date(end); ; ) {
      const m = dt.getMonth();
      const d = dt.getDate();
      const st = monthlyData[m]?.days?.[d - 1]?.status ?? null;

      // Count SAFE and NEAR MISS as streak days; break on ACCIDENT or NOT SET.
      if (st === 'safe' || st === 'near_miss') streak += 1;
      else break;

      dt.setDate(dt.getDate() - 1);
      if (dt.getFullYear() !== y) break;
    }
    return streak;
  }, [monthlyData, currentYear]);

  const monthSummary = useMemo(() => {
    if (!displayMonthData) return { safe: 0, near: 0, accident: 0 };
    let safe = 0, near = 0, accident = 0;
    for (const d of displayMonthData.days) {
      if (d.status === 'safe') safe += 1;
      if (d.status === 'near_miss') near += 1;
      if (d.status === 'accident') accident += 1;
    }
    return { safe, near, accident };
  }, [displayMonthData]);

  const firstDayOffset = useMemo(() => new Date(currentYear, displayMonth, 1).getDay(), [currentYear, displayMonth]);
  const daysInMonth = useMemo(() => new Date(currentYear, displayMonth + 1, 0).getDate(), [currentYear, displayMonth]);
  const gridCells = useMemo(() => {
    const cells: Array<{ day: number | null }> = [];
    for (let i = 0; i < firstDayOffset; i++) cells.push({ day: null });
    for (let d = 1; d <= daysInMonth; d++) cells.push({ day: d });
    while (cells.length % 7 !== 0) cells.push({ day: null });
    return cells;
  }, [firstDayOffset, daysInMonth]);

  const setDayStatus = (day: number, status: DayStatus) => {
    setMonthlyData((prev) => {
      const next = prev.map((mm) => ({ ...mm, days: mm.days.map((dd) => ({ ...dd })) }));
      const month = next[displayMonth];
      if (!month) return prev;
      const target = month.days[day - 1];
      if (!target) return prev;
      target.status = status;
      return next;
    });
  };

  const cycleDayStatus = (day: number) => {
    const current = displayMonthData?.days?.[day - 1]?.status ?? null;
    setDayStatus(day, nextDayStatus(current));
  };

  const resetLayout = () => {
    try {
      localStorage.removeItem('react-resizable-panels:nhk-safety-layout-v2-structure');
    } catch {}
    try { window.location.reload(); } catch {}
  };

  // ---- Upload helpers
  const readFileToDataUrl = (file: File, cb: (url: string) => void) => {
    const reader = new FileReader();
    reader.onload = () => cb(String(reader.result));
    reader.readAsDataURL(file);
  };

  const onPosterTopSelected = (file?: File | null) => {
    if (!file) return;
    readFileToDataUrl(file, (url) => {
      setPosterTop(url);
      setPosterTopZoom(1);
      setLastUpdateIso(new Date().toISOString());
    });
  };

  const onPosterBottomSelected = (file?: File | null) => {
    if (!file) return;
    readFileToDataUrl(file, (url) => {
      setPosterBottom(url);
      setPosterBottomZoom(1);
      setLastUpdateIso(new Date().toISOString());
    });
  };

  const onPolicyReplaceSelected = (file?: File | null) => {
    if (!file) return;
    readFileToDataUrl(file, (url) => {
      setPolicyImages((p) => {
        const next = [...p];
        next[0] = url;
        return next.slice(0, 2);
      });
      setPolicyZoom(1);
      setLastUpdateIso(new Date().toISOString());
    });
  };

  const onPolicyAddSelected = (file?: File | null) => {
    if (!file) return;
    readFileToDataUrl(file, (url) => {
      setPolicyImages((p) => {
        const next = [...p];
        if (next.length < 2) next.push(url);
        else next[1] = url;
        return next.slice(0, 2);
      });
      setPolicyZoom(1);
      setLastUpdateIso(new Date().toISOString());
    });
  };

  // ---- Editors
  const openMetricsEditor = useCallback(() => {
    setMetricsDraft(metrics.map((m) => ({ ...m })));
    setEditMetrics(true);
  }, [metrics]);

  const addMetricDraft = useCallback(() => {
    setMetricsDraft((p) => [{ id: uid('m'), label: 'New Metric', value: '0', unit: '' }, ...p]);
  }, []);

  const updateMetricDraft = useCallback((id: string, patch: Partial<SafetyMetric>) => {
    setMetricsDraft((p) => p.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  }, []);

  const deleteMetricDraft = useCallback((id: string) => {
    setMetricsDraft((p) => p.filter((m) => m.id !== id));
  }, []);

  const saveMetricsEditor = useCallback(() => {
    setMetrics(metricsDraft.map((m) => ({ ...m })));
    setEditMetrics(false);
    setLastUpdateIso(new Date().toISOString());
  }, [metricsDraft]);

  const openTargetEditor = useCallback(() => {
    setTargetVarsDraft({ ...targetVars });
    setTargetFormulasDraft({ ...targetFormulas });
    setBestRecordDraft(String(bestRecord));
    setLossTimeAccidentsDraft(String(lossTimeAccidents));
    setEditTarget(true);
  }, [targetVars, targetFormulas, bestRecord, lossTimeAccidents]);

  const saveTargetEditor = useCallback(() => {
    setTargetVars({ ...targetVarsDraft });
    setTargetFormulas({ ...targetFormulasDraft });
    setBestRecord(Number(bestRecordDraft) || 0);
    setLossTimeAccidents(Number(lossTimeAccidentsDraft) || 0);
    setEditTarget(false);
    setLastUpdateIso(new Date().toISOString());
  }, [targetVarsDraft, targetFormulasDraft, bestRecordDraft, lossTimeAccidentsDraft]);

  const openTickerEditor = useCallback(() => {
    setTickerDraft(announcements.map((a) => a.text).join('\n'));
    setEditTicker(true);
  }, [announcements]);

  const saveTickerEditor = useCallback(() => {
    const lines = tickerDraft
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    const next = lines.length
      ? lines.map((text, idx) => ({ id: String(idx + 1), text }))
      : DEFAULT_ANNOUNCEMENTS;
    setAnnouncements(next);
    setEditTicker(false);
    setLastUpdateIso(new Date().toISOString());
  }, [tickerDraft]);

  // ---- Derived target numbers
  const totalManHours = useMemo(() => safeEval(targetFormulas.totalExpr, targetVars), [targetFormulas, targetVars]);
  const toDateManHours = useMemo(() => safeEval(targetFormulas.toDateExpr, targetVars), [targetFormulas, targetVars]);

  const lastUpdateLabel = useMemo(() => {
    try {
      const d = new Date(lastUpdateIso);
      return d.toLocaleDateString([], { day: '2-digit', month: 'short', year: 'numeric' });
    } catch {
      return '';
    }
  }, [lastUpdateIso]);

  const tickerText = useMemo(() => {
    const txt = announcements.map((a) => a.text).filter(Boolean).join('   •   ');
    return txt || 'No announcements.';
  }, [announcements]);

  // ticker speed based on length (keep calm)
  const tickerSeconds = useMemo(() => clamp(Math.round(tickerText.length / 7), 18, 48), [tickerText]);

  return (
    <div
      className="h-screen max-h-screen max-w-[100vw] overflow-hidden flex flex-col bg-[radial-gradient(circle_at_top_left,_#dbeafe_0%,_#f0f9ff_30%,_#ffffff_55%,_#fefce8_80%,_#ecfdf5_100%)] text-slate-900"
      style={{
        fontSize: 'var(--font-size)',
        ['--color-firstAid' as any]: '#16a34a', // green
        ['--color-nonAbsent' as any]: '#2563eb', // blue
        ['--color-absent' as any]: '#f59e0b', // orange
        ['--color-fire' as any]: '#ef4444', // red
      }}
    >
      <style>
        {`
          .header-date { font-size: clamp(1.05rem, 1.1vw, 1.5rem); }
          .header-time { font-size: clamp(1.2rem, 1.35vw, 1.9rem); }
          .ticker-wrap { height: clamp(48px, 6vh, 64px); }
          @keyframes marqueeLTR { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }
        `}
      </style>

      {/* HEADER */}
      <header className="px-6 py-4 flex items-center gap-4 rounded-b-3xl border-b border-white/70 bg-white/70 backdrop-blur supports-[backdrop-filter]:bg-white/60 shadow-sm">
        <div className="flex items-center gap-3 shrink-0">
          <div className="h-12 flex items-center rounded-2xl bg-white/80 border border-slate-200 shadow-sm px-3 overflow-hidden">
            <img src={nhkLogo} alt="NHK SPRING (THAILAND)" className="h-9 w-auto" />
          </div>
          <div className="text-2xl font-extrabold">Safety Dashboard</div>
        </div>

        <div className="flex-1 flex justify-center">
          <HeaderClock />
        </div>

        <div className="flex items-center gap-3 shrink-0">
          <div className="flex items-center gap-1 rounded-2xl border border-slate-200 bg-white px-3 py-2 shadow-sm">
            <button
              type="button"
              onClick={() => setUiScale((s) => clamp(Number((s - 0.05).toFixed(2)), 0.8, 1.4))}
              className="h-8 w-10 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50"
              aria-label="Decrease font size"
              title="ลดขนาดตัวอักษร"
            >
              A-
            </button>
            <div className="w-14 text-center text-xs font-extrabold text-slate-600">{Math.round(uiScale * 100)}%</div>
            <button
              type="button"
              onClick={() => setUiScale((s) => clamp(Number((s + 0.05).toFixed(2)), 0.8, 1.4))}
              className="h-8 w-10 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50"
              aria-label="Increase font size"
              title="เพิ่มขนาดตัวอักษร"
            >
              A+
            </button>
            <button
              type="button"
              onClick={() => setUiScale(1)}
              className="h-8 px-3 rounded-xl border border-slate-200 bg-white text-xs font-extrabold hover:bg-slate-50"
              aria-label="Reset font size"
              title="รีเซ็ตขนาดตัวอักษร"
            >
              Reset
            </button>
          </div>

          <button
            type="button"
            onClick={() => setLayoutLocked((v) => !v)}
            className={`px-4 py-2 rounded-2xl border font-extrabold flex items-center gap-2 ${layoutLocked ? 'border-slate-200 bg-white hover:bg-slate-50' : 'border-sky-200 bg-sky-50 hover:bg-sky-100'}`}
            title={layoutLocked ? 'Unlock layout to resize panels' : 'Lock layout'}
          >
            {layoutLocked ? <Lock className="h-4 w-4" /> : <Unlock className="h-4 w-4" />}
            {layoutLocked ? 'LOCKED' : 'UNLOCKED'}
          </button>

          <button type="button" onClick={resetLayout} className="px-4 py-2 rounded-2xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50" title="Reset layout">
            Reset Layout
          </button>
        </div>
      </header>

      {/* MAIN (leave space for ticker) */}
      <main className="px-4 pt-2 pb-[max(18px,6vh)] flex-1 min-h-0 overflow-hidden">
        <div className="h-full min-h-0 rounded-3xl overflow-hidden">
          {/* New structure based on your sketch (V2 split):
              Left: Poster(top) + Poster(bottom)
              Center: Safety Policy (images) + Safety & Environment Target
              Right: Streak + Calendar
          */}
          <ResizablePanelGroup direction="horizontal" className="h-full" autoSaveId="nhk-safety-layout-v2-structure">
            {/* LEFT */}
            <ResizablePanel defaultSize={23} minSize={18}>
              <ResizablePanelGroup direction="vertical" className="h-full">
                <ResizablePanel defaultSize={50} minSize={22}>
                  <div className="h-full min-h-0 p-1">
                    <PanelWrap>
                      {(panelScale) => (
                        <Card
                          title="Poster (Policy)"
                          icon={<ImageIcon className="h-5 w-5 text-amber-700" />}
                          tone="amber"
                          panelScale={panelScale}
                          actions={
                            <>
                              <input ref={posterTopInputRef} type="file" accept="image/*" className="hidden" onChange={(e) => onPosterTopSelected(e.target.files?.[0])} />
                              <button onClick={() => posterTopInputRef.current?.click()} className="p-2 rounded-lg hover:bg-amber-50" title="Upload">
                                <Upload className="h-4 w-4 text-amber-700" />
                              </button>
                              {posterTop ? (
                                <>
                                  <button onClick={() => setPosterTopZoom((z) => clamp(z - 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white border border-transparent hover:border-amber-200" title="Zoom out">
                                    <span className="text-amber-800 font-bold">−</span>
                                  </button>
                                  <button onClick={() => setPosterTopZoom(1)} className="px-2 py-1.5 rounded-lg text-xs font-bold border border-amber-200 bg-white hover:bg-amber-50" title="Reset size">
                                    {Math.round(posterTopZoom * 100)}%
                                  </button>
                                  <button onClick={() => setPosterTopZoom((z) => clamp(z + 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white border border-transparent hover:border-amber-200" title="Zoom in">
                                    <span className="text-amber-800 font-bold">+</span>
                                  </button>
                                  <button onClick={() => setPosterTop(null)} className="p-2 rounded-lg hover:bg-rose-50" title="Remove">
                                    <Trash2 className="h-4 w-4 text-rose-600" />
                                  </button>
                                </>
                              ) : null}
                            </>
                          }
                        >
                          {!posterTop ? (
                            <div className="h-full flex flex-col items-center justify-center text-center gap-3 rounded-2xl border-2 border-dashed border-slate-200 bg-slate-50 p-6">
                              <Upload className="h-8 w-8 text-slate-500" />
                              <div className="font-bold text-slate-700">Upload Poster</div>
                              <div className="text-sm text-slate-500">รองรับรูปแนวตั้ง/แนวนอน</div>
                            </div>
                          ) : (
                            <div className="h-full flex flex-col min-h-0">
                              <div className="w-full flex-1 min-h-0 rounded-2xl border border-slate-200 bg-white overflow-hidden flex items-center justify-center relative">
                                <img
                                  src={posterTop}
                                  alt="Poster"
                                  className="max-h-full max-w-full object-contain select-none"
                                  style={{ transform: `scale(${posterTopZoom})`, transformOrigin: 'center center' }}
                                />
                              </div>
                            </div>
                          )}
                        </Card>
                      )}
                    </PanelWrap>
                  </div>
                </ResizablePanel>

                <ResizableHandle withHandle disabled={layoutLocked} className="bg-sky-200/70 hover:bg-sky-300 data-[disabled]:opacity-30" />

                <ResizablePanel defaultSize={50} minSize={22}>
                  <div className="h-full min-h-0 p-1">
                    <PanelWrap>
                      {(panelScale) => (
                        <Card
                          title="Poster"
                          icon={<ImageIcon className="h-5 w-5 text-amber-700" />}
                          tone="amber"
                          panelScale={panelScale}
                          actions={
                            <>
                              <input ref={posterBottomInputRef} type="file" accept="image/*" className="hidden" onChange={(e) => onPosterBottomSelected(e.target.files?.[0])} />
                              <button onClick={() => posterBottomInputRef.current?.click()} className="p-2 rounded-lg hover:bg-amber-50" title="Upload">
                                <Upload className="h-4 w-4 text-amber-700" />
                              </button>
                              {posterBottom ? (
                                <>
                                  <button onClick={() => setPosterBottomZoom((z) => clamp(z - 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white border border-transparent hover:border-amber-200" title="Zoom out">
                                    <span className="text-amber-800 font-bold">−</span>
                                  </button>
                                  <button onClick={() => setPosterBottomZoom(1)} className="px-2 py-1.5 rounded-lg text-xs font-bold border border-amber-200 bg-white hover:bg-amber-50" title="Reset size">
                                    {Math.round(posterBottomZoom * 100)}%
                                  </button>
                                  <button onClick={() => setPosterBottomZoom((z) => clamp(z + 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white border border-transparent hover:border-amber-200" title="Zoom in">
                                    <span className="text-amber-800 font-bold">+</span>
                                  </button>
                                  <button onClick={() => setPosterBottom(null)} className="p-2 rounded-lg hover:bg-rose-50" title="Remove">
                                    <Trash2 className="h-4 w-4 text-rose-600" />
                                  </button>
                                </>
                              ) : null}
                            </>
                          }
                        >
                          {!posterBottom ? (
                            <div className="h-full flex flex-col items-center justify-center text-center gap-3 rounded-2xl border-2 border-dashed border-slate-200 bg-slate-50 p-6">
                              <Upload className="h-8 w-8 text-slate-500" />
                              <div className="font-bold text-slate-700">Upload Poster</div>
                              <div className="text-sm text-slate-500">แนะนำอัตราส่วน A4 / แนวตั้ง</div>
                            </div>
                          ) : (
                            <div className="h-full flex flex-col min-h-0">
                              <div className="w-full flex-1 min-h-0 rounded-2xl border border-slate-200 bg-white overflow-hidden flex items-center justify-center relative">
                                <img
                                  src={posterBottom}
                                  alt="Poster"
                                  className="max-h-full max-w-full object-contain select-none"
                                  style={{ transform: `scale(${posterBottomZoom})`, transformOrigin: 'center center' }}
                                />
                              </div>
                            </div>
                          )}
                        </Card>
                      )}
                    </PanelWrap>
                  </div>
                </ResizablePanel>
              </ResizablePanelGroup>
            </ResizablePanel>

            <ResizableHandle withHandle disabled={layoutLocked} className="bg-sky-200/70 hover:bg-sky-300 data-[disabled]:opacity-30" />

            {/* CENTER */}
            <ResizablePanel defaultSize={52} minSize={34}>
              <ResizablePanelGroup direction="vertical" className="h-full">
                <ResizablePanel defaultSize={28} minSize={18}>
                  <div className="h-full min-h-0 p-1">
                    <PanelWrap>
                      {(panelScale) => (
                        <Card
                          title="Safety Policy"
                          icon={<Shield className="h-5 w-5 text-sky-700" />}
                          tone="sky"
                          panelScale={panelScale}
                          actions={
                            <>
                              <input ref={policyReplaceInputRef} type="file" accept="image/*" className="hidden" onChange={(e) => onPolicyReplaceSelected(e.target.files?.[0])} />
                              <input ref={policyAddInputRef} type="file" accept="image/*" className="hidden" onChange={(e) => onPolicyAddSelected(e.target.files?.[0])} />

                              <button onClick={() => policyReplaceInputRef.current?.click()} className="p-2 rounded-lg hover:bg-white/70" title="Replace image">
                                <Upload className="h-4 w-4 text-slate-600" />
                              </button>
                              <button
                                onClick={() => policyAddInputRef.current?.click()}
                                disabled={policyImages.length >= 2}
                                className="p-2 rounded-lg hover:bg-white/70 disabled:opacity-40"
                                title="Add second image (max 2)"
                              >
                                <Plus className="h-4 w-4 text-slate-600" />
                              </button>
                              {policyImages.length > 1 ? (
                                <button
                                  onClick={() => setPolicyImages((p) => p.slice(0, 1))}
                                  className="p-2 rounded-lg hover:bg-rose-50"
                                  title="Remove second image"
                                >
                                  <Trash2 className="h-4 w-4 text-rose-600" />
                                </button>
                              ) : null}
                              <button onClick={() => setPolicyZoom((z) => clamp(z - 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white/70" title="Zoom out">
                                <span className="text-slate-700 font-bold">−</span>
                              </button>
                              <button onClick={() => setPolicyZoom(1)} className="px-2 py-1.5 rounded-lg text-xs font-bold border border-slate-200 bg-white hover:bg-slate-50" title="Reset size">
                                {Math.round(policyZoom * 100)}%
                              </button>
                              <button onClick={() => setPolicyZoom((z) => clamp(z + 0.1, 0.5, 2.5))} className="p-2 rounded-lg hover:bg-white/70" title="Zoom in">
                                <span className="text-slate-700 font-bold">+</span>
                              </button>
                            </>
                          }
                        >
                          <div className="h-full min-h-0 rounded-2xl border border-slate-200 bg-white overflow-hidden flex items-center justify-center">
                            {policyImages.length <= 1 ? (
                              <img
                                src={policyImages[0]}
                                alt="Safety Policy"
                                className="max-h-full max-w-full object-contain select-none"
                                style={{ transform: `scale(${policyZoom})`, transformOrigin: 'center center' }}
                              />
                            ) : (
                              <div className="h-full w-full grid grid-cols-2 gap-2 p-2">
                                {policyImages.slice(0, 2).map((src, idx) => (
                                  <div key={idx} className="rounded-xl border border-slate-200 bg-white overflow-hidden flex items-center justify-center">
                                    <img
                                      src={src}
                                      alt={`Policy ${idx + 1}`}
                                      className="max-h-full max-w-full object-contain select-none"
                                      style={{ transform: `scale(${policyZoom})`, transformOrigin: 'center center' }}
                                    />
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        </Card>
                      )}
                    </PanelWrap>
                  </div>
                </ResizablePanel>

                <ResizableHandle withHandle disabled={layoutLocked} className="bg-sky-200/70 hover:bg-sky-300 data-[disabled]:opacity-30" />

                <ResizablePanel defaultSize={72} minSize={38}>
                  <div className="h-full min-h-0 p-1">
                    <PanelWrap>
                      {(panelScale) => {
                        const metricCols = panelScale < 0.82 ? 1 : 2;
                        const metricRows = Math.max(1, Math.ceil(metrics.length / metricCols));
                        const densityScale = clamp((6 / Math.max(metrics.length, 1)) ** 0.28, 0.78, 1);
                        const cardScale = clamp(panelScale * densityScale * (metricRows >= 4 ? 0.92 : 1), 0.68, 1.08);

                        return (
                          <Card
                            title="Safety & Environment Target"
                            icon={<Target className="h-5 w-5 text-cyan-700" />}
                            tone="teal"
                            panelScale={panelScale}
                            actions={
                              <>
                                <button type="button" onClick={openTargetEditor} className="p-2 rounded-lg hover:bg-white/70" title="Edit Target / Formula" aria-label="Edit Target / Formula">
                                  <Edit className="h-4 w-4 text-slate-600" />
                                </button>
                                <button type="button" onClick={openMetricsEditor} className="p-2 rounded-lg hover:bg-white/70" title="Edit Metrics" aria-label="Edit Metrics">
                                  <Activity className="h-4 w-4 text-slate-600" />
                                </button>
                              </>
                            }
                          >
                            <div className="h-full min-h-0 flex flex-col overflow-hidden">
                              {/* Target text block */}
                              <div className="rounded-2xl border border-cyan-100 bg-white/70 p-3">
                                <div className="font-extrabold text-slate-900" style={{ fontSize: scaledPx(16, panelScale, 12, 20) }}>
                                  Safety & Environment Target
                                </div>
                                <div className="mt-1 text-slate-800 font-medium leading-relaxed" style={{ fontSize: scaledPx(13, panelScale, 11, 16) }}>
                                  <div>
                                    <span className="font-bold">Total Working Time</span> = ({targetVars.manpower} man × {targetVars.daysPerWeek} Day/Week × {targetVars.hoursPerDay} Hr. × {targetVars.workingDaysYear} Day)
                                    {' '}= <span className="font-extrabold">{formatNumber(totalManHours)}</span> man-hours
                                  </div>
                                  <div>
                                    <span className="font-bold">To Date Record</span> = ({targetVars.manpower} man × {targetVars.daysPerWeek} Day/Week × {targetVars.hoursPerDay} Hr. × {targetVars.workingDaysSoFar} Day)
                                    {' '}= <span className="font-extrabold">{formatNumber(toDateManHours)}</span> man-hours
                                  </div>
                                  <div>
                                    <span className="font-bold">Best Record</span> = <span className="font-extrabold">{formatNumber(bestRecord)}</span> man-hours
                                  </div>
                                  <div>
                                    <span className="font-bold">Number of loss time accident in this year</span> = <span className="font-extrabold">{formatNumber(lossTimeAccidents)}</span> Time
                                  </div>
                                </div>
                              </div>

                              <div className="my-3 h-px bg-cyan-200/70" />

                              {/* Metrics grid (no graph) */}
                              <div className="min-h-0 flex-1 overflow-hidden">
                                <div
                                  className="grid gap-2 h-full min-h-0"
                                  style={{
                                    gridTemplateColumns: `repeat(${metricCols}, minmax(0, 1fr))`,
                                    gridTemplateRows: `repeat(${metricRows}, minmax(0, 1fr))`,
                                  }}
                                >
                                  {metrics.map((m) => (
                                    <div
                                      key={m.id}
                                      className="rounded-2xl border p-3 flex flex-col justify-between shadow-[0_1px_0_rgba(2,132,199,0.05)] min-h-0 overflow-hidden"
                                      style={metricAccentStyle(metricToneKey(m.label))}
                                    >
                                      <div className="font-semibold text-slate-700 leading-tight break-words line-clamp-2" style={{ fontSize: scaledPx(14, cardScale, 10, 16) }}>
                                        <span className="inline-flex items-center gap-2">
                                          <span className="h-2.5 w-2.5 rounded-full" style={{ background: metricToneKey(m.label) === 'neutral' ? 'var(--tone-cyan-600)' : `var(--color-${metricToneKey(m.label)})` }} />
                                          <span className="min-w-0 truncate">{m.label}</span>
                                        </span>
                                      </div>
                                      <div className="mt-1 flex items-end gap-1 min-w-0">
                                        <div className="font-extrabold text-slate-900 leading-none truncate" style={{ fontSize: scaledPx(38, cardScale, 18, 48) }}>
                                          {m.value}
                                        </div>
                                        <div className="font-semibold text-slate-500 pb-[0.12rem] truncate" style={{ fontSize: scaledPx(16, cardScale, 9, 18) }}>
                                          {m.unit}
                                        </div>
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              </div>

                              <div className="mt-2 flex justify-end text-slate-600 font-semibold" style={{ fontSize: scaledPx(12, panelScale, 10, 14) }}>
                                {lastUpdateLabel ? `Date Update: ${lastUpdateLabel}` : ''}
                              </div>
                            </div>
                          </Card>
                        );
                      }}
                    </PanelWrap>
                  </div>
                </ResizablePanel>
              </ResizablePanelGroup>
            </ResizablePanel>

            <ResizableHandle withHandle disabled={layoutLocked} className="bg-sky-200/70 hover:bg-sky-300 data-[disabled]:opacity-30" />

            {/* RIGHT */}
            <ResizablePanel defaultSize={25} minSize={20}>
              <ResizablePanelGroup direction="vertical" className="h-full">
                <ResizablePanel defaultSize={34} minSize={22}>
                  <div className="h-full min-h-0 p-1">
                    <PanelWrap>
                      {(panelScale) => (
                        <Card title="Safety Streak" icon={<Flame className="h-5 w-5 text-emerald-700" />} tone="green" panelScale={panelScale}>
                          <div className="h-full flex flex-col items-center justify-center text-center">
                            <div className="font-bold text-slate-600" style={{ fontSize: scaledPx(14, panelScale, 12, 18) }}>Zero Accident Days</div>
                            <div className="mt-2 font-extrabold text-emerald-700 leading-none" style={{ fontSize: scaledPx(84, panelScale, 46, 120) }}>{safetyStreak}</div>
                            <div className="mt-2 font-semibold text-slate-700" style={{ fontSize: scaledPx(16, panelScale, 12, 22) }}>days</div>
                            <div className="mt-4 inline-flex items-center gap-2 rounded-2xl bg-sky-50 border border-sky-100 px-4 py-2">
                              <Shield className="h-4 w-4 text-sky-700" />
                              <span className="text-sm font-bold text-slate-700">Zero Accident Workplace</span>
                            </div>
                          </div>
                        </Card>
                      )}
                    </PanelWrap>
                  </div>
                </ResizablePanel>

                <ResizableHandle withHandle disabled={layoutLocked} className="bg-sky-200/70 hover:bg-sky-300 data-[disabled]:opacity-30" />

                <ResizablePanel defaultSize={66} minSize={30}>
                  <div className="h-full min-h-0 p-1">
                    <PanelWrap>
                      {(panelScale) => {
                        const weekRows = Math.max(4, Math.ceil(gridCells.length / 7));
                        const compactScale = clamp(panelScale * (weekRows >= 6 ? 0.9 : 1), 0.62, 1.08);
                        return (
                          <Card title="Safety Calendar" icon={<CalendarIcon className="h-5 w-5 text-sky-600" />} tone="sky" panelScale={panelScale}>
                            <div className="h-full flex flex-col min-h-0">
                              <div className="flex items-center justify-between gap-2 mb-2 shrink-0">
                                <div className="flex items-center gap-2 min-w-0">
                                  <button type="button" className="px-2 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50" onClick={() => setDisplayMonth((m) => (m + 11) % 12)} title="Prev">‹</button>
                                  <div className="text-base md:text-lg font-extrabold text-slate-900 truncate">{MONTHS[displayMonth]} {currentYear}</div>
                                  <button type="button" className="px-2 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50" onClick={() => setDisplayMonth((m) => (m + 1) % 12)} title="Next">›</button>
                                </div>
                                <div className="grid grid-cols-3 gap-1 text-[10px] md:text-xs font-bold shrink-0">
                                  <div className="px-2 py-1 rounded-lg bg-emerald-50 border border-emerald-100 text-emerald-700">{monthSummary.safe} SAFE</div>
                                  <div className="px-2 py-1 rounded-lg bg-amber-50 border border-amber-100 text-amber-800">{monthSummary.near} NM</div>
                                  <div className="px-2 py-1 rounded-lg bg-rose-50 border border-rose-100 text-rose-700">{monthSummary.accident} ACC</div>
                                </div>
                              </div>

                              <div className="grid grid-cols-7 gap-1 mb-1 shrink-0">
                                {DAY_HEADERS.map((d) => (
                                  <div key={d} className="text-[10px] md:text-xs font-bold text-slate-500 text-center">{d}</div>
                                ))}
                              </div>
                              <div className="mb-1 rounded-lg border border-sky-100 bg-sky-50/80 px-2 py-1 text-[10px] md:text-xs font-semibold text-sky-800 shrink-0 leading-tight">
                                คลิกวันเพื่อเปลี่ยนสถานะ: ว่าง → SAFE → NEAR MISS → ACCIDENT → ว่าง
                              </div>

                              <div className="grid grid-cols-7 gap-1 h-full min-h-0" style={{ gridTemplateRows: `repeat(${weekRows}, minmax(0, 1fr))` }}>
                                {gridCells.map((c, idx) => {
                                  if (!c.day) return <div key={idx} className="rounded-xl bg-transparent min-h-0" />;
                                  const cellDate = new Date(currentYear, displayMonth, c.day);
                                  const iso = toISODate(cellDate);
                                  const isWeekend = cellDate.getDay() === 0 || cellDate.getDay() === 6;
                                  const holidayNote = HOLIDAY_NOTES_2026[iso];
                                  const isHoliday = isWeekend || !!holidayNote;
                                  const st = displayMonthData?.days?.[c.day - 1]?.status ?? null;
                                  const cls = st === 'safe' ? 'border-emerald-200 bg-emerald-50'
                                    : st === 'near_miss' ? 'border-amber-200 bg-amber-50'
                                    : st === 'accident' ? 'border-rose-200 bg-rose-50'
                                    : isHoliday ? 'border-slate-200 bg-slate-100/80'
                                    : 'border-slate-200 bg-white';
                                  const statusText = st === 'safe' ? 'SAFE' : st === 'near_miss' ? 'NEAR MISS' : st === 'accident' ? 'ACCIDENT' : 'NOT SET';
                                  const statusTone = st === 'safe'
                                    ? 'text-emerald-700 bg-emerald-100/80 border-emerald-200'
                                    : st === 'near_miss'
                                    ? 'text-amber-800 bg-amber-100/80 border-amber-200'
                                    : st === 'accident'
                                    ? 'text-rose-700 bg-rose-100/80 border-rose-200'
                                    : isHoliday
                                    ? 'text-slate-600 bg-slate-200/60 border-slate-300'
                                    : 'text-slate-500 bg-slate-50 border-slate-200';

                                  return (
                                    <button
                                      key={idx}
                                      type="button"
                                      onClick={() => cycleDayStatus(c.day!)}
                                      className={`rounded-xl border p-1 md:p-1.5 flex flex-col gap-1 text-left cursor-pointer hover:shadow-sm transition-shadow min-h-0 overflow-hidden ${cls}`}
                                      title={holidayNote ? `${holidayNote} • คลิกเพื่อเปลี่ยนสถานะ` : isWeekend ? `วันหยุดสุดสัปดาห์ • คลิกเพื่อเปลี่ยนสถานะ` : 'คลิกเพื่อเปลี่ยนสถานะ'}
                                    >
                                      <div className="flex items-center justify-between gap-1 shrink-0">
                                        <div className={`font-extrabold ${(!st && isHoliday) ? 'text-slate-600' : 'text-slate-900'}`} style={{ fontSize: scaledPx(13, compactScale, 10, 16) }}>{c.day}</div>
                                        {st === 'safe' && <CheckCircle2 className="h-3.5 w-3.5 text-emerald-700 shrink-0" />}
                                        {st === 'near_miss' && <AlertTriangle className="h-3.5 w-3.5 text-amber-700 shrink-0" />}
                                        {st === 'accident' && <AlertTriangle className="h-3.5 w-3.5 text-rose-700 shrink-0" />}
                                      </div>
                                      <div className={`mt-auto rounded-md border px-1 py-0.5 font-bold text-center leading-tight ${statusTone}`} style={{ fontSize: scaledPx(10, compactScale, 8, 12) }}>
                                        {statusText}
                                      </div>
                                    </button>
                                  );
                                })}
                              </div>
                            </div>
                          </Card>
                        );
                      }}
                    </PanelWrap>
                  </div>
                </ResizablePanel>
              </ResizablePanelGroup>
            </ResizablePanel>
          </ResizablePanelGroup>
        </div>
      </main>

      {/* TICKER (bottom) */}
      <div className="ticker-wrap fixed bottom-0 left-0 right-0 z-20 border-t border-slate-200 bg-white/85 backdrop-blur supports-[backdrop-filter]:bg-white/70">
        <div className="h-full flex items-center gap-3 px-4">
          <div className="shrink-0 font-extrabold text-slate-800">Announcement</div>
          <div className="flex-1 overflow-hidden">
            <div
              className="whitespace-nowrap font-bold text-slate-700"
              style={{
                animation: `marqueeLTR ${tickerSeconds}s linear infinite`,
                willChange: 'transform',
              }}
            >
              {tickerText}
            </div>
          </div>
          <button
            type="button"
            onClick={openTickerEditor}
            className="shrink-0 h-9 px-3 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50 inline-flex items-center gap-2"
            title="Edit announcements"
          >
            <Edit className="h-4 w-4" /> Edit
          </button>
        </div>
      </div>

      {/* Metrics dialog */}
      <Dialog open={editMetrics} onOpenChange={(open) => { if (!open) setEditMetrics(false); }}>
        <DialogContent className="w-[min(980px,96vw)] max-w-none">
          <DialogHeader>
            <DialogTitle className="text-slate-900 font-extrabold">Safety Metrics Settings</DialogTitle>
            <div className="text-sm text-slate-600 font-semibold">เพิ่ม/ลบ/แก้ไขหัวข้อได้สะดวก • กด Save เพื่อบันทึก</div>
          </DialogHeader>

          <div className="flex items-center justify-between gap-2">
            <button
              type="button"
              onClick={addMetricDraft}
              className="px-4 py-2 rounded-xl bg-cyan-600 text-white font-extrabold hover:bg-cyan-700 inline-flex items-center gap-2"
            >
              <Plus className="h-4 w-4" /> เพิ่มหัวข้อ
            </button>
            <div className="text-xs font-bold text-slate-500">รายการทั้งหมด: {metricsDraft.length}</div>
          </div>

          <ScrollArea className="h-[60vh] rounded-2xl border border-slate-200 bg-slate-50/40 p-3">
            <div className="space-y-2">
              {metricsDraft.map((m, idx) => (
                <div key={m.id} className="rounded-2xl border border-slate-200 bg-white p-3 shadow-sm">
                  <div className="grid gap-2 items-end" style={{ gridTemplateColumns: 'minmax(220px,1.6fr) minmax(110px,0.5fr) minmax(90px,0.45fr) auto' }}>
                    <div className="min-w-0">
                      <label className="block text-xs font-extrabold text-slate-500 mb-1">หัวข้อ #{idx + 1}</label>
                      <input
                        value={m.label}
                        onChange={(e) => updateMetricDraft(m.id, { label: e.target.value })}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-extrabold text-slate-500 mb-1">ค่า</label>
                      <input
                        value={m.value}
                        onChange={(e) => updateMetricDraft(m.id, { value: e.target.value })}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 text-right bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                        inputMode="decimal"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-extrabold text-slate-500 mb-1">หน่วย</label>
                      <input
                        value={m.unit || ''}
                        onChange={(e) => updateMetricDraft(m.id, { unit: e.target.value })}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                      />
                    </div>
                    <div className="flex items-end justify-end">
                      <button
                        type="button"
                        onClick={() => deleteMetricDraft(m.id)}
                        className="h-10 px-3 rounded-xl border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-700 font-extrabold inline-flex items-center gap-1"
                        title="ลบหัวข้อ"
                      >
                        <Trash2 className="h-4 w-4" /> ลบ
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </ScrollArea>

          <DialogFooter className="gap-2">
            <button type="button" onClick={() => setEditMetrics(false)} className="px-4 py-2 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50">Cancel</button>
            <button type="button" onClick={saveMetricsEditor} className="px-4 py-2 rounded-xl bg-emerald-600 text-white font-extrabold hover:bg-emerald-700 inline-flex items-center gap-2">
              <Save className="h-4 w-4" /> Save
            </button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Target dialog */}
      <Dialog open={editTarget} onOpenChange={(open) => { if (!open) setEditTarget(false); }}>
        <DialogContent className="w-[min(980px,96vw)] max-w-none">
          <DialogHeader>
            <DialogTitle className="text-slate-900 font-extrabold">Safety & Environment Target Settings</DialogTitle>
            <div className="text-sm text-slate-600 font-semibold">แก้ไขสูตร (Formula) และตัวแปร รวมถึงค่า Best record / Loss time accident</div>
          </DialogHeader>

          <ScrollArea className="h-[60vh] rounded-2xl border border-slate-200 bg-slate-50/40 p-3">
            <div className="space-y-4">
              <div className="rounded-2xl border border-slate-200 bg-white p-3">
                <div className="font-extrabold text-slate-800 mb-2">Variables</div>
                <div className="grid gap-3" style={{ gridTemplateColumns: 'repeat(2, minmax(0,1fr))' }}>
                  {([
                    ['manpower','Manpower (man)'],
                    ['daysPerWeek','Day/Week'],
                    ['hoursPerDay','Hr/Day'],
                    ['workingDaysYear','Working Days/Year'],
                    ['workingDaysSoFar','Working Days (To Date)'],
                  ] as Array<[keyof TargetVars, string]>).map(([k, label]) => (
                    <div key={String(k)}>
                      <label className="block text-xs font-extrabold text-slate-500 mb-1">{label}</label>
                      <input
                        value={String(targetVarsDraft[k])}
                        onChange={(e) => setTargetVarsDraft((p) => ({ ...p, [k]: Number(e.target.value) || 0 }))}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                        inputMode="decimal"
                      />
                    </div>
                  ))}
                </div>
              </div>

              <div className="rounded-2xl border border-slate-200 bg-white p-3">
                <div className="font-extrabold text-slate-800 mb-2">Formulas</div>
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-extrabold text-slate-500 mb-1">Total Working Time (expression)</label>
                    <input
                      value={targetFormulasDraft.totalExpr}
                      onChange={(e) => setTargetFormulasDraft((p) => ({ ...p, totalExpr: e.target.value }))}
                      className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                      placeholder="manpower * daysPerWeek * hoursPerDay * workingDaysYear"
                    />
                    <div className="mt-1 text-xs font-semibold text-slate-500">Result: {formatNumber(safeEval(targetFormulasDraft.totalExpr, targetVarsDraft))} man-hours</div>
                  </div>
                  <div>
                    <label className="block text-xs font-extrabold text-slate-500 mb-1">To Date Record (expression)</label>
                    <input
                      value={targetFormulasDraft.toDateExpr}
                      onChange={(e) => setTargetFormulasDraft((p) => ({ ...p, toDateExpr: e.target.value }))}
                      className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                      placeholder="manpower * daysPerWeek * hoursPerDay * workingDaysSoFar"
                    />
                    <div className="mt-1 text-xs font-semibold text-slate-500">Result: {formatNumber(safeEval(targetFormulasDraft.toDateExpr, targetVarsDraft))} man-hours</div>
                  </div>
                </div>
              </div>

              <div className="rounded-2xl border border-slate-200 bg-white p-3">
                <div className="font-extrabold text-slate-800 mb-2">Values</div>
                <div className="grid gap-3" style={{ gridTemplateColumns: 'repeat(2, minmax(0,1fr))' }}>
                  <div>
                    <label className="block text-xs font-extrabold text-slate-500 mb-1">Best Record (man-hours)</label>
                    <input
                      value={bestRecordDraft}
                      onChange={(e) => setBestRecordDraft(e.target.value)}
                      className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                      inputMode="decimal"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-extrabold text-slate-500 mb-1">Loss time accident (Time)</label>
                    <input
                      value={lossTimeAccidentsDraft}
                      onChange={(e) => setLossTimeAccidentsDraft(e.target.value)}
                      className="w-full rounded-xl border border-slate-200 px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-cyan-200"
                      inputMode="decimal"
                    />
                  </div>
                </div>
              </div>
            </div>
          </ScrollArea>

          <DialogFooter className="gap-2">
            <button type="button" onClick={() => setEditTarget(false)} className="px-4 py-2 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50">Cancel</button>
            <button type="button" onClick={saveTargetEditor} className="px-4 py-2 rounded-xl bg-emerald-600 text-white font-extrabold hover:bg-emerald-700 inline-flex items-center gap-2">
              <Save className="h-4 w-4" /> Save
            </button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Ticker edit dialog */}
      <Dialog open={editTicker} onOpenChange={(open) => { if (!open) setEditTicker(false); }}>
        <DialogContent className="w-[min(980px,96vw)] max-w-none">
          <DialogHeader>
            <DialogTitle className="text-slate-900 font-extrabold">Announcement Ticker</DialogTitle>
            <div className="text-sm text-slate-600 font-semibold">พิมพ์ 1 บรรทัด = 1 ข้อความ • ระบบจะนำไปวิ่งเป็นแถบด้านล่าง</div>
          </DialogHeader>

          <textarea
            value={tickerDraft}
            onChange={(e) => setTickerDraft(e.target.value)}
            className="w-full h-[55vh] rounded-2xl border border-slate-200 p-4 font-semibold bg-white focus:outline-none focus:ring-2 focus:ring-sky-200"
            placeholder="พิมพ์ข้อความ..."
          />

          <DialogFooter className="gap-2">
            <button type="button" onClick={() => setEditTicker(false)} className="px-4 py-2 rounded-xl border border-slate-200 bg-white font-extrabold hover:bg-slate-50">Cancel</button>
            <button type="button" onClick={saveTickerEditor} className="px-4 py-2 rounded-xl bg-emerald-600 text-white font-extrabold hover:bg-emerald-700 inline-flex items-center gap-2">
              <Save className="h-4 w-4" /> Save
            </button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
