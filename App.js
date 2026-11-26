import {jsxs as _jsxs, jsx as _jsx, Fragment as _Fragment} from "react/jsx-runtime";import { useEffect, useMemo, useRef, useState } from 'react';
import { buildWorkDocx, saveDocx, saveFile, safeFileName } from './docx.js?v=20261126-1614';
import { buildStatementImage } from './wage-image.js?v=20261126-1614';

const APP_VERSION = '1.0.0';

const KEY = 'gongri.v1';
const PROJECTS_KEY = 'gongri.projects.v1';
const ACTIVE_KEY = 'gongri.active.v1';
const RATES_KEY = 'gongri.rates.v1';
const LIVING_KEY = 'gongri.migrate.living.v1';
const APP_BUILD_KEY = 'gongri.web.build.v1';

const pad = (value) => String(value).padStart(2, '0');
const iso = (date) => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
const fromISO = (text) => {
  const [year, month, day] = text.split('-').map(Number);
  return new Date(year, month - 1, day);
};
const todayISO = () => iso(new Date());

const round2 = (value) => Math.round((value + Number.EPSILON) * 100) / 100;
/// 2.5 记成 "2.5"，2 记成 "2"。
const trim = (value) => String(round2(value));

const amountFmt = new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 });
const moneyFmt = new Intl.NumberFormat('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const yuan = (value) => `¥${moneyFmt.format(round2(value))}`;

/// 一天按几小时折时薪、加班几小时算一个工日，都可以在项目设置里改。
const DEFAULT_NORM_HOURS = 9;
const DEFAULT_OT_HOURS = 6;

const monthOf = (date) => `${date.getFullYear()}-${pad(date.getMonth() + 1)}`;
const shiftMonthKey = (key, delta) => {
  const [year, month] = key.split('-').map(Number);
  return monthOf(new Date(year, month - 1 + delta, 1));
};

const monthLabel = (key) => {
  const [year, month] = String(key).split('-');
  return `${year}年${Number(month)}月`;
};

const dayLabel = (text) => {
  const date = fromISO(text);
  return `${date.getMonth() + 1}月${date.getDate()}日`;
};

/// 农历新年（正月初一）的公历日期，用来划分「本年度」：从今年春节到明年春节前一天。
const LUNAR_NEW_YEAR = {
  2015: '2015-02-19', 2016: '2016-02-08', 2017: '2017-01-28', 2018: '2018-02-16',
  2019: '2019-02-05', 2020: '2020-01-25', 2021: '2021-02-12', 2022: '2022-02-01',
  2023: '2023-01-22', 2024: '2024-02-10', 2025: '2025-01-29', 2026: '2026-02-17',
  2027: '2027-02-06', 2028: '2028-01-26', 2029: '2029-02-13', 2030: '2030-02-03',
  2031: '2031-01-23', 2032: '2032-02-11', 2033: '2033-01-31', 2034: '2034-02-19',
  2035: '2035-02-08', 2036: '2036-01-28', 2037: '2037-02-15', 2038: '2038-02-04',
  2039: '2039-01-24', 2040: '2040-02-12', 2041: '2041-02-01', 2042: '2042-01-22',
  2043: '2043-02-10', 2044: '2044-01-30', 2045: '2045-02-17',
};

function lunarYearRange() {
  const today = todayISO();
  const starts = Object.values(LUNAR_NEW_YEAR).sort();
  const start = [...starts].reverse().find((day) => day <= today);
  const next = starts.find((day) => day > today);
  if (!start || !next) {
    const year = today.slice(0, 4);
    return { start: `${year}-01-01`, end: `${year}-12-31` };
  }
  const endDate = fromISO(next);
  endDate.setDate(endDate.getDate() - 1);
  return { start, end: iso(endDate) };
}

const describeMissingDays = (days) => {
  if (days.length === 1) return dayLabel(days[0]);
  const first = days[0];
  const last = days[days.length - 1];
  const span = (fromISO(last) - fromISO(first)) / 86400000 + 1;
  if (span === days.length) return `${dayLabel(first)}—${dayLabel(last)}`;
  return `${days.slice(0, 3).map(dayLabel).join('、')} 等`;
};

const dateLabel = (text) => {
  if (!text) return '';
  const date = fromISO(text);
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
};

/// 加班小时数：结束早于开始就当作跨到第二天。
function overtimeHours(overtime) {
  if (!overtime || !overtime.start || !overtime.end) return 0;
  const [startHour, startMinute] = overtime.start.split(':').map(Number);
  const [endHour, endMinute] = overtime.end.split(':').map(Number);
  let minutes = endHour * 60 + endMinute - (startHour * 60 + startMinute);
  if (minutes < 0) minutes += 24 * 60;
  return round2(minutes / 60);
}

/// 只有收款（发生活费）、没有工日和加班的记录。
const isPaymentOnly = (record) =>
  (Number(record.paid) || 0) > 0 &&
  (Number(record.days) || 0) === 0 &&
  overtimeHours(record.overtime) === 0;

/// 一条记录的工资：工日 × 日薪 + 加班小时 × 加班时薪，按分取整。
function recordWage(record, daily, otHourly) {
  const days = Number(record.days) || 0;
  return round2(days * daily + overtimeHours(record.overtime) * otHourly);
}

/// 一个项目的日薪、工时和两个时薪：时薪按正常工时折，加班时薪按加班折算小时折。
function projectRates(project) {
  const daily = Number(project && project.daily) > 0 ? Number(project.daily) : 0;
  const normHours =
    Number(project && project.normHours) > 0 ? Number(project.normHours) : DEFAULT_NORM_HOURS;
  const otHours = Number(project && project.otHours) > 0 ? Number(project.otHours) : DEFAULT_OT_HOURS;
  return {
    daily,
    normHours,
    otHours,
    hourly: round2(daily / normHours),
    otHourly: round2(daily / otHours),
  };
}

/// 项目名太长时拆成两行：「××建设项目（××院区）」显示成主名 + 括号里的说明。
function splitProjectName(name) {
  const text = (name || '').trim();
  const matched = text.match(/^([^（(]+?)\s*([（(].+[）)])$/);
  if (matched) return { main: matched[1].trim(), sub: matched[2].trim() };
  return { main: text, sub: '' };
}

function normalizeProject(project, index) {
  const source = project && typeof project === 'object' ? project : {};
  return {
    id: typeof source.id === 'string' && source.id ? source.id : `p${index + 1}`,
    name: typeof source.name === 'string' ? source.name : '',
    daily: Number(source.daily) > 0 ? Number(source.daily) : 0,
    normHours: Number(source.normHours) > 0 ? Number(source.normHours) : DEFAULT_NORM_HOURS,
    otHours: Number(source.otHours) > 0 ? Number(source.otHours) : DEFAULT_OT_HOURS,
  };
}

function readRecords() {
  try {
    const saved = JSON.parse(localStorage.getItem(KEY));
    if (Array.isArray(saved)) {
      return saved.filter((item) => item && typeof item.date === 'string');
    }
  } catch (e) {}
  return [];
}

/// 旧版只有一份项目和一份设置。这里把它们迁成第一条项目，旧记录补上 projectId。
function loadInitialData() {
  const records = readRecords();
  let projects = [];
  let migrated = false;
  try {
    const saved = JSON.parse(localStorage.getItem(PROJECTS_KEY));
    if (Array.isArray(saved) && saved.length > 0) {
      projects = saved.map(normalizeProject);
      migrated = true;
    }
  } catch (e4) {}

  if (!migrated) {
    let legacy = {};
    try {
      legacy = JSON.parse(localStorage.getItem(RATES_KEY)) || {};
    } catch (e5) {}
    const daily = Number(legacy.daily) > 0 ? Number(legacy.daily) : 0;
    let otHours = DEFAULT_OT_HOURS;
    // 旧版可能手填过加班时薪，换算成「几小时一个工」，金额保持不变。
    const legacyHourly = Number(legacy.hourly);
    if (daily > 0 && legacyHourly > 0) {
      const derived = round2(daily / legacyHourly);
      if (derived >= 1 && derived <= 24) otHours = derived;
    }
    projects = [
      normalizeProject(
        {
          id: 'p1',
          name: typeof legacy.project === 'string' ? legacy.project : '',
          daily,
          normHours: DEFAULT_NORM_HOURS,
          otHours,
        },
        0
      ),
    ];
  }

  const firstId = projects[0].id;
  let stamped = false;
  const withProject = records.map((record) => {
    if (record.projectId && projects.some((project) => project.id === record.projectId)) {
      return record;
    }
    stamped = true;
    return { ...record, projectId: firstId };
  });

  const savedActive = (() => {
    try {
      return localStorage.getItem(ACTIVE_KEY);
    } catch (e6) {
      return null;
    }
  })();
  const activeId = projects.some((project) => project.id === savedActive) ? savedActive : firstId;

  if (!migrated || stamped) {
    localStorage.setItem(PROJECTS_KEY, JSON.stringify(projects));
  }
  if (stamped) localStorage.setItem(KEY, JSON.stringify(withProject));
  localStorage.setItem(ACTIVE_KEY, activeId);

  return { projects, records: withProject, activeId };
}

/// 发工资、生活费统一成生活费：旧记录的 note 一次性改掉。
function applyLivingMigration(data) {
  try {
    if (localStorage.getItem(LIVING_KEY)) return data;
    const records = data.records.map((record) =>
      record.note === 'wage' ? { ...record, note: 'living' } : record
    );
    localStorage.setItem(KEY, JSON.stringify(records));
    localStorage.setItem(LIVING_KEY, '1');
    return { ...data, records };
  } catch (e8) {
    return data;
  }
}

let BOOT = loadInitialData();
BOOT = applyLivingMigration(BOOT);

const newestFirst = (a, b) => {
  if (a.date !== b.date) return a.date < b.date ? 1 : -1;
  return String(a.id) < String(b.id) ? 1 : -1;
};

const inRange = (date, start, end) => date >= start && date <= end;

const dayShift = (text, delta) => {
  const date = fromISO(text);
  date.setDate(date.getDate() + delta);
  return iso(date);
};

/// 某个项目在这个月里、指定日期之前没有记录的日子（升序）。
function missingDays(records, projectId, monthKey, beforeISO) {
  const [year, month] = monthKey.split('-').map(Number);
  const total = new Date(year, month, 0).getDate();
  const taken = new Set(
    records
      .filter((record) => record.projectId === projectId && record.date.slice(0, 7) === monthKey)
      .map((record) => record.date)
  );
  const list = [];
  for (let day = 1; day <= total; day += 1) {
    const key = `${monthKey}-${pad(day)}`;
    if (key >= beforeISO) break;
    if (!taken.has(key)) list.push(key);
  }
  return list;
}

export default function App() {
  const [projects, setProjects] = useState(BOOT.projects);
  const [records, setRecords] = useState(BOOT.records);
  const [activeId, setActiveId] = useState(BOOT.activeId);
  const [cursor, setCursor] = useState(() => {
    const now = new Date();
    return { y: now.getFullYear(), m: now.getMonth() };
  });
  const [view, setView] = useState('home');
  const [editing, setEditing] = useState(null);
  const [editingProject, setEditingProject] = useState(null);
  const [menuProject, setMenuProject] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [batch, setBatch] = useState(null);
  const [booting, setBooting] = useState(true);
  const pendingUpdate = useRef(null);
  const importRef = useRef(null);

  /// 网页版：发现线上有新构建时，提示更新（小程序里取不到 version.json，会静默跳过）。
  useEffect(() => {
    let alive = true;
    const check = async () => {
      try {
        const response = await fetch(`./version.json?t=${Date.now()}`, { cache: 'no-store' });
        if (!response.ok) return;
        const data = await response.json();
        const build = String((data && data.build) || '');
        if (!build || !alive) return;
        let stored = '';
        try {
          stored = localStorage.getItem(APP_BUILD_KEY) || '';
        } catch (error) {}
        if (!stored) {
          try {
            localStorage.setItem(APP_BUILD_KEY, build);
          } catch (error) {}
          return;
        }
        if (stored !== build) {
          try {
            localStorage.setItem(APP_BUILD_KEY, build);
          } catch (error) {}
          // 正在填表/弹层时先记下来，等关掉再刷新，免得丢掉没保存的输入。
          if (document.querySelector('.sheet')) {
            pendingUpdate.current = build;
          } else {
            location.replace(`${location.pathname}?v=${encodeURIComponent(build)}`);
          }
        }
      } catch (error) {}
    };
    const timer = setTimeout(check, 1500);
    const onVisible = () => {
      if (document.visibilityState === 'visible') check();
    };
    document.addEventListener('visibilitychange', onVisible);
    return () => {
      alive = false;
      clearTimeout(timer);
      document.removeEventListener('visibilitychange', onVisible);
    };
  }, []);

  useEffect(() => {
    if (!pendingUpdate.current) return;
    if (document.querySelector('.sheet')) return;
    location.replace(`${location.pathname}?v=${encodeURIComponent(pendingUpdate.current)}`);
  }, [editing, editingProject, menuProject, exporting, batch]);

  useEffect(() => {
    localStorage.setItem(KEY, JSON.stringify(records));
  }, [records]);

  useEffect(() => {
    localStorage.setItem(PROJECTS_KEY, JSON.stringify(projects));
  }, [projects]);

  useEffect(() => {
    localStorage.setItem(ACTIVE_KEY, activeId);
  }, [activeId]);

  const activeProject = projects.find((project) => project.id === activeId) || projects[0];

  /// 在职 = 最近一条真正上工（工日 > 0）的项目，且不超过 30 天；生活费、休息、离职记录都不算。
  const inServiceId = useMemo(() => {
    const today = todayISO();
    let best = null;
    for (const project of projects) {
      let last = null;
      for (const record of records) {
        if (record.projectId !== project.id) continue;
        if ((Number(record.days) || 0) <= 0) continue;
        if (!last || record.date > last) last = record.date;
      }
      if (!last) continue;
      if (!best || last > best.date) best = { id: project.id, date: last };
    }
    // 一条上工记录都还没有时，先按当前打开的项目算在职。
    if (!best) return activeId;
    const gap = (fromISO(today) - fromISO(best.date)) / 86400000;
    return gap <= 30 ? best.id : null;
  }, [projects, records]);
  const rates = projectRates(activeProject);

  const now = new Date();
  const isThisMonth = cursor.y === now.getFullYear() && cursor.m === now.getMonth();
  const monthKey = `${cursor.y}-${pad(cursor.m + 1)}`;

  const activeRecords = useMemo(
    () => records.filter((record) => record.projectId === activeProject.id),
    [records, activeProject.id]
  );

  const monthRecords = useMemo(
    () => activeRecords.filter((record) => record.date.slice(0, 7) === monthKey).sort(newestFirst),
    [activeRecords, monthKey]
  );

  const totalDays = monthRecords.reduce((sum, record) => sum + (Number(record.days) || 0), 0);
  const totalHours = monthRecords.reduce((sum, record) => sum + overtimeHours(record.overtime), 0);
  const totalWage = monthRecords.reduce(
    (sum, record) => sum + recordWage(record, rates.daily, rates.otHourly),
    0
  );
  // 支账 = 本月发出去的工资 + 生活费，一起算。
  const totalPaid = monthRecords.reduce((sum, record) => sum + (Number(record.paid) || 0), 0);
  const monthBalance = round2(totalWage - totalPaid);

  /// 有记录的月份，翻页时跳过其余空月份；但当月及以后可以逐月往前走，方便开新的一月。
  const monthKeys = useMemo(() => {
    const set = new Set(activeRecords.map((record) => record.date.slice(0, 7)));
    return [...set].sort();
  }, [activeRecords]);

  const viewingMonth = `${cursor.y}-${pad(cursor.m + 1)}`;
  const todayMonth = monthOf(now);
  const stepPrev = shiftMonthKey(viewingMonth, -1);
  const stepNext = shiftMonthKey(viewingMonth, 1);

  const prevCandidates = monthKeys.filter((key) => key < viewingMonth);
  const nextCandidates = [...new Set([...monthKeys, todayMonth])]
    .filter((key) => key > viewingMonth)
    .sort();

  const targetPrev = stepPrev >= todayMonth ? stepPrev : prevCandidates[prevCandidates.length - 1];
  const targetNext = stepNext >= todayMonth ? stepNext : nextCandidates[0];
  const hasPrevMonth = Boolean(targetPrev);
  const hasNextMonth = Boolean(targetNext);

  const shift = (delta) => {
    const target = delta < 0 ? targetPrev : targetNext;
    if (!target) return;
    const [year, month] = target.split('-').map(Number);
    setCursor({ y: year, m: month - 1 });
  };

  const backToNow = () => setCursor({ y: now.getFullYear(), m: now.getMonth() });

  const save = (record) => {
    // 同一个项目里一天只留一条：同一天的旧记录被这条替换。
    setRecords((list) => [
      ...list.filter(
        (item) =>
          item.id !== record.id &&
          !(item.projectId === record.projectId && item.date === record.date)
      ),
      record,
    ]);
    const [y, m] = record.date.split('-').map(Number);
    setCursor({ y, m: m - 1 });
    setEditing(null);
  };

  const remove = (id) => {
    setRecords((list) => list.filter((item) => item.id !== id));
    setEditing(null);
  };

  const saveProject = (draft) => {
    if (draft.id) {
      setProjects((list) =>
        list.map((project) => (project.id === draft.id ? { ...project, ...draft } : project))
      );
      setActiveId(draft.id);
    } else {
      const id = `p-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 5)}`;
      setProjects((list) => [
        ...list,
        {
          id,
          name: draft.name,
          daily: draft.daily,
          normHours: draft.normHours,
          otHours: draft.otHours,
        },
      ]);
      setActiveId(id);
      setView('project');
    }
    setEditingProject(null);
  };

  const askDeleteProject = (project) => {
    setMenuProject(null);
    if (projects.length <= 1) return;
    const count = records.filter((record) => record.projectId === project.id).length;
    const name = project.name || '未命名项目';
    if (!confirm(`删除「${name}」？\n将一并删除 ${count} 条记录，不能恢复。`)) return;
    deleteProject(project);
  };

  const deleteProject = (project) => {
    if (projects.length <= 1) return;
    const next = projects.filter((item) => item.id !== project.id);
    setProjects(next);
    setRecords((list) => list.filter((record) => record.projectId !== project.id));
    if (activeId === project.id) setActiveId(next[0].id);
    setEditingProject(null);
    setView('home');
  };

  /// 把所有项目和记录导出成 JSON，换手机或清数据前备份用。
  const exportData = async () => {
    const payload = {
      app: 'gongri',
      version: 1,
      exportedAt: new Date().toISOString(),
      projects,
      records,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    await saveFile(blob, `记工数据_${todayISO()}.json`, 'application/json');
  };

  /// 从备份文件恢复，会覆盖当前所有项目和记录。
  const importData = async (event) => {
    const input = event.target;
    const file = input.files && input.files[0];
    input.value = '';
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (!data || !Array.isArray(data.projects) || !Array.isArray(data.records)) {
        throw new Error('bad file');
      }
      const nextProjects = data.projects.map(normalizeProject);
      if (nextProjects.length === 0) throw new Error('no project');
      const ids = new Set(nextProjects.map((project) => project.id));
      const firstId = nextProjects[0].id;
      const nextRecords = data.records
        .filter((record) => record && typeof record.date === 'string')
        .map((record) => {
          const note = record.note === 'wage' ? 'living' : record.note;
          return ids.has(record.projectId)
            ? { ...record, note }
            : { ...record, projectId: firstId, note };
        });
      if (
        !confirm(
          `导入「${file.name}」？\n会覆盖当前 ${projects.length} 个项目、${records.length} 条记录，不能撤销。`
        )
      ) {
        return;
      }
      localStorage.setItem(PROJECTS_KEY, JSON.stringify(nextProjects));
      localStorage.setItem(KEY, JSON.stringify(nextRecords));
      localStorage.setItem(ACTIVE_KEY, firstId);
      location.reload();
    } catch (e9) {
      alert('这个文件不是「记工」导出的数据，换一个再试。');
    }
  };

  /// 「记一笔」永远按实际时间带上今天的日期（翻到别的月份也一样）。
  const openAdd = () => setEditing({ projectId: activeProject.id, date: todayISO() });

  /// 首页「今天」的一键记工。

  /// 批量补记的默认区间：优先用这个月到昨天为止的空档，没有就用最近 7 天。
  const openBatch = () => {
    const today = todayISO();
    const gaps = missingDays(records, activeProject.id, monthKey, today);
    setBatch(
      gaps.length > 0
        ? { start: gaps[0], end: gaps[gaps.length - 1] }
        : { start: dayShift(today, -7), end: dayShift(today, -1) }
    );
  };

  /// 批量补记保存：同一项目同一天只留一条。
  const saveBatch = (list) => {
    if (list.length > 0) {
      const keys = new Set(list.map((record) => `${record.projectId}|${record.date}`));
      setRecords((prev) => [
        ...prev.filter((record) => !keys.has(`${record.projectId}|${record.date}`)),
        ...list,
      ]);
      const [year, month] = list[0].date.split('-').map(Number);
      setCursor({ y: year, m: month - 1 });
    }
    setBatch(null);
  };

  /// 生成当前项目、当前月份的工资单图片。
  /// 按导出面板选的项目和时间，生成一张工资单 / 汇总图片。
  const runImageExport = async ({ projectIds, start, end, label, fields }) => {
    const chosen = projects.filter((project) => projectIds.includes(project.id));
    const sections = [];
    const totals = { days: 0, hours: 0, wage: 0, paid: 0 };
    for (const project of chosen) {
      const rates = projectRates(project);
      const own = records.filter((record) => record.projectId === project.id);
      const picked = own.filter((record) => inRange(record.date, start, end));
      if (picked.length === 0) continue;
      const wage = round2(
        picked.reduce((sum, record) => sum + recordWage(record, rates.daily, rates.otHourly), 0)
      );
      const paid = round2(picked.reduce((sum, record) => sum + (Number(record.paid) || 0), 0));
      const days = round2(picked.reduce((sum, record) => sum + (Number(record.days) || 0), 0));
      const hours = round2(
        picked.reduce((sum, record) => sum + overtimeHours(record.overtime), 0)
      );
      const allWage = round2(
        own.reduce((sum, record) => sum + recordWage(record, rates.daily, rates.otHourly), 0)
      );
      const allPaid = round2(own.reduce((sum, record) => sum + (Number(record.paid) || 0), 0));
      sections.push({
        name: project.name || '未命名项目',
        days,
        hours,
        wage,
        paid,
        allWage,
        allPaid,
        worked: picked.filter((record) => record.status === 'work').length,
        rested: picked.filter((record) => record.status === 'rest' && !isPaymentOnly(record))
          .length,
      });
      totals.days += days;
      totals.hours += hours;
      totals.wage += wage;
      totals.paid += paid;
    }
    if (sections.length === 0) {
      alert('这段时间没有记录，换一个时间或项目再导出。');
      return false;
    }
    const now = new Date();
    try {
      const blob = await buildStatementImage({
        title: sections.length === 1 ? sections[0].name : `${sections.length} 个项目`,
        subtitle: label,
        fields,
        sections,
        totals: {
          days: round2(totals.days),
          hours: round2(totals.hours),
          wage: round2(totals.wage),
          paid: round2(totals.paid),
        },
        generatedAt: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`,
      });
      const who =
        sections.length === 1 ? safeFileName(sections[0].name) : `${sections.length}个项目`;
      await saveFile(blob, `记工_${who}_${safeFileName(label)}.png`, 'image/png');
      return true;
    } catch (e11) {
      alert('图片生成失败，请再试一次。');
      return false;
    }
  };

  const runExport = ({ projectIds, start, end, label, fields }) => {
    const now = new Date();
    const generatedAt =
      `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ` +
      `${pad(now.getHours())}:${pad(now.getMinutes())}`;

    const chosen = projects.filter((project) => projectIds.includes(project.id));
    const sections = [];
    let grandWage = 0;
    let grandPaid = 0;
    let grandPeriodWage = 0;
    let grandPeriodPaid = 0;

    for (const project of chosen) {
      const projectRate = projectRates(project);
      const picked = records
        .filter(
          (record) => record.projectId === project.id && inRange(record.date, start, end)
        )
        .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
      if (picked.length === 0) continue;

      const wage = picked.reduce(
        (sum, record) => sum + recordWage(record, projectRate.daily, projectRate.otHourly),
        0
      );
      const paid = picked.reduce((sum, record) => sum + (Number(record.paid) || 0), 0);

      const own = records.filter((record) => record.projectId === project.id);
      const allWage = own.reduce(
        (sum, record) => sum + recordWage(record, projectRate.daily, projectRate.otHourly),
        0
      );
      const allPaid = own.reduce((sum, record) => sum + (Number(record.paid) || 0), 0);
      const balance = round2(allWage - allPaid);
      grandWage += allWage;
      grandPaid += allPaid;
      grandPeriodWage += wage;
      grandPeriodPaid += paid;

      // 这段时间跨了年份，日期就带上年，免得同一天重复看不出是哪年。
      const multiYear = new Set(picked.map((record) => record.date.slice(0, 4))).size > 1;

      sections.push({
        project: project.name || '未命名项目',
        periodLabel: label,
        rows: picked.map((record) => {
          const date = fromISO(record.date);
          const hours = overtimeHours(record.overtime);
          const note =
            record.note === 'am'
              ? '上午上班'
              : record.note === 'pm'
                ? '下午上班'
                : record.note === 'wage' || record.note === 'living'
                  ? '生活费'
                  : '无';
          return {
            date: `${
              multiYear ? `${date.getFullYear()}年` : ''
            }${date.getMonth() + 1}月${date.getDate()}日`,
            // 只有收款的哪天不写「休息」，状态留空。
            status:
              record.status === 'rest' && isPaymentOnly(record)
                ? '—'
                : record.status === 'work'
                  ? '上班'
                  : record.status === 'leave'
                    ? '离职'
                    : '休息',
            days: trim(Number(record.days) || 0),
            hours: hours > 0 ? trim(hours) : '',
            overtime: record.overtime
              ? `${record.overtime.start}–${
                  record.overtime.end < record.overtime.start ? '次日 ' : ''
                }${record.overtime.end}`
              : '',
            wage: yuan(recordWage(record, projectRate.daily, projectRate.otHourly)),
            note,
            paid: Number(record.paid) > 0 ? yuan(Number(record.paid)) : '',
          };
        }),
        month: {
          days: trim(picked.reduce((sum, record) => sum + (Number(record.days) || 0), 0)),
          hours: trim(
            picked.reduce((sum, record) => sum + overtimeHours(record.overtime), 0)
          ),
          wage: yuan(wage),
          paid: yuan(paid),
          worked: picked.filter((record) => record.status === 'work').length,
          rested: picked.filter((record) => record.status === 'rest' && !isPaymentOnly(record))
            .length,
          leaves: picked.filter((record) => record.status === 'leave').length,
        },
        totals: {
          wage: yuan(allWage),
          paid: yuan(allPaid),
          balance: yuan(balance >= 0 ? Math.abs(balance) : Math.abs(allPaid)),
          balanceLabel: balance >= 0 ? '工资余额' : '支账',
          negative: balance < 0,
        },
      });
    }

    if (sections.length === 0) {
      alert('这段时间没有记录，换一个时间或项目再导出。');
      return false;
    }

    const grandBalance = round2(grandWage - grandPaid);
    const blob = buildWorkDocx({
      title: sections.length === 1 ? sections[0].project : '记工汇总',
      subtitle:
        sections.length === 1 ? `${label} 记工明细` : `${label} · ${sections.length} 个项目`,
      periodWord: /月$/.test(label) ? '本月' : '本期',
      generatedAt,
      iso: now.toISOString(),
      fields,
      sections,
      grand:
        sections.length > 1
          ? {
              periodWage: yuan(grandPeriodWage),
              periodPaid: yuan(grandPeriodPaid),
              wage: yuan(grandWage),
              paid: yuan(grandPaid),
              balance: yuan(grandBalance >= 0 ? Math.abs(grandBalance) : Math.abs(grandPaid)),
              balanceLabel: grandBalance >= 0 ? '合计余额' : '支账',
              negative: grandBalance < 0,
            }
          : null,
    });

    const who =
      sections.length === 1 ? safeFileName(sections[0].project) : `${sections.length}个项目`;
    const filename = `记工_${who}_${safeFileName(label)}.docx`;

    saveDocx(blob, filename)
      .then((result) => {
        if (result === 'cancelled') return;
        setExporting(false);
      })
      .catch(() => alert('导出没有成功，请再试一次。'));
    return true;
  };

  const projectCount = projects.length;
  const lunarRange = lunarYearRange();

  /// 进项目时，如果当前翻到的月份在这个项目里没有记录，就跳到它最近有记录的月份。
  const openProject = (id) => {
    const own = records.filter((record) => record.projectId === id);
    const viewing = `${cursor.y}-${pad(cursor.m + 1)}`;
    if (own.length > 0 && !own.some((record) => record.date.slice(0, 7) === viewing)) {
      // 当前月份没这个项目的记录，就跳到它最近有工日的月份。
      const worked = own
        .filter((record) => (Number(record.days) || 0) > 0 || overtimeHours(record.overtime) > 0)
        .map((record) => record.date.slice(0, 7))
        .sort();
      const months = worked.length > 0 ? worked : own.map((record) => record.date.slice(0, 7)).sort();
      const latest = months[months.length - 1];
      if (latest) {
        const [year, month] = latest.split('-').map(Number);
        setCursor({ y: year, m: month - 1 });
      }
    }
    setActiveId(id);
    setView('project');
  };

  return (
    _jsxs('main', { className: "page", children: [
      _jsx('div', { className: "view", children: 
      view === 'home' ? (
        _jsxs(_Fragment, { children: [
          _jsx(HomeSummary, {
            projects: projects,
            records: records,
            yearStart: lunarRange.start,
            yearEnd: lunarRange.end,}
          )

          , _jsxs('div', { className: "cards", children: [
            projects.map((project) => (
              _jsx(SwipeRow, {
                onDelete: () => askDeleteProject(project),
                disabled: projects.length <= 1,
                children: _jsx(ProjectCard, {
                  project: project,
                  records: records,
                  active: project.id === inServiceId,
                  showState: true,
                  onOpen: () => openProject(project.id),
                  onMenu: () => setMenuProject(project),})}, project.id
              )
            ))

            , _jsxs('button', { className: "add-project", onClick: () => setEditingProject({ project: null }), children: [
              _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                _jsx('path', { d: "M12 5v14M5 12h14"  , stroke: "currentColor", strokeWidth: "2.2", strokeLinecap: "round", fill: "none",} )
              }), "新建项目"

            ]})
          ]})

        ]})
      ) : (
        _jsxs(_Fragment, { children: [
          _jsxs('header', { className: "projecttop", children: [
            _jsxs('button', { type: "button", className: "back", onClick: () => setView('home'), children: [
              _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                _jsx('path', { d: "M14.5 5.5 8 12l6.5 6.5"    , fill: "none", stroke: "currentColor", strokeWidth: "2.2", strokeLinecap: "round", strokeLinejoin: "round",} )
              }), "首页"

            ]})

            , _jsxs('div', { className: "month-nav", children: [
              _jsx('button', {
                className: "month-arrow",
                'aria-label': "上个月",
                onClick: () => shift(-1),
                disabled: !hasPrevMonth,
 children: 
                _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                  _jsx('path', { d: "M14.5 5.5 8 12l6.5 6.5"    , fill: "none", stroke: "currentColor", strokeWidth: "2.2", strokeLinecap: "round", strokeLinejoin: "round",} )
                })
              })
              , _jsxs('span', { className: "month-name", children: [
                cursor.y, "年", cursor.m + 1, "月"
              ]})
              , _jsx('button', {
                className: "month-arrow",
                'aria-label': "下个月",
                onClick: () => shift(1),
                disabled: !hasNextMonth,
 children: 
                _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                  _jsx('path', { d: "M9.5 5.5 16 12l-6.5 6.5"    , fill: "none", stroke: "currentColor", strokeWidth: "2.2", strokeLinecap: "round", strokeLinejoin: "round",} )
                })
              })
              , !isThisMonth && (
                _jsx('button', { className: "back-now", onClick: backToNow, children: "回本月"})
              )
            ]})

            , _jsxs('button', { className: "export", onClick: () => setExporting(true), 'aria-label': "导出记录", children: [
              _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                _jsx('path', { d: "M12 4v11m0 0 4.5-4.5M12 15l-4.5-4.5M5 19.5h14"     , fill: "none", stroke: "currentColor", strokeWidth: "2", strokeLinecap: "round", strokeLinejoin: "round",} )
              }), "导出"

            ]})
          ]})

          , _jsx(ProjectCard, {
            project: activeProject,
            records: records,
            onEdit: () => setEditingProject({ project: activeProject }),}
          )

          , _jsx('section', { className: "statcard", children: 
            _jsxs('div', { className: "board-stats", children: [
              _jsxs('div', { className: "board-head", children: [
                _jsx('span', { className: "board-label", children: "本月统计"})
                , _jsxs('button', { type: "button", className: "board-action", onClick: openBatch, children: [
                  _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                    _jsx('path', { d: "M8 3v3M16 3v3M4 9h16M6 5h12a1 1 0 0 1 1 1v13a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1Z" , fill: "none", stroke: "currentColor", strokeWidth: "1.8", strokeLinecap: "round", strokeLinejoin: "round",} )
                  }), "补记"

                ]})
              ]})
              , _jsx('p', { className: "board-value", children: 
                `${trim(totalDays)} 个工日 · 加班 ${trim(totalHours)} 小时`
              })

              , _jsxs('p', { className: "stats-metrics", children: [
                _jsxs('span', { children: ["工资 " , _jsx('b', { children: yuan(totalWage)})]})
                , monthBalance >= 0 && _jsxs('span', { children: ["支账 " , _jsx('b', { className: "pay", children: yuan(totalPaid)})]})
              ]})

              , _jsxs('div', { className: "stats-balance", children: [
                _jsx('span', { children: monthBalance >= 0 ? '本月应得余额' : '本月支账'})
                , _jsx('b', { children: yuan(monthBalance >= 0 ? monthBalance : totalPaid)})
              ]})
            ]})
          })

          , _jsxs('section', { className: "board", children: [
            _jsxs('div', { className: "thead", children: [
              _jsx('span', { children: "日期"})
              , _jsx('span', { children: "状态"})
              , _jsx('span', { className: "num", children: "工日"})
              , _jsx('span', { className: "num", children: "加班"})
              , _jsx('span', { className: "num", children: "工资"})
              , _jsx('span', { className: "num", children: "备注"})
            ]})

            , monthRecords.length > 0 ? (
              _jsx('ul', { className: "rows", children: 
                monthRecords.map((record) => (
                  _jsx('li', { children: 
                    _jsx(Row, {
                      record: record,
                      daily: rates.daily,
                      otHourly: rates.otHourly,
                      onOpen: () => setEditing(record),}
                    )
                  }, record.id)
                ))
              }, monthKey)
            ) : (
              _jsxs('div', { className: "empty", children: [
                _jsx('span', { className: "empty-icon", 'aria-hidden': "true", children: "📒"})
                , _jsx('p', { children: "这个月还没有记录"})
                , _jsx('p', { className: "empty-hint", children: 
                  rates.daily > 0
                    ? '点下面的“记一笔”，把上班或休息记上。'
                    : '先点顶部设置项目名称和日薪，再记工，工资会自动算。'
                })
              ]})
            )
          ]})

          , _jsxs('div', { className: "bottombar", children: [
            _jsxs('button', { className: "add", onClick: openAdd, children: [
              _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
                _jsx('path', { d: "M12 5v14M5 12h14"  , stroke: "currentColor", strokeWidth: "2.4", strokeLinecap: "round", fill: "none",} )
              }), "记一笔"

            ]})
          ]})
        ]})
      )
      }, view)

      , editing && (
        _jsx(Editor, {
          record: editing,
          records: records,
          daily: rates.daily,
          otHourly: rates.otHourly,
          onSave: save,
          onDelete: remove,
          onClose: () => setEditing(null),}
        )
      )

      , menuProject && (
        _jsx(ProjectMenu, {
          project: menuProject,
          canDelete: projectCount > 1,
          onEdit: () => {
            const target = menuProject;
            setMenuProject(null);
            setEditingProject({ project: target });
          },
          onDelete: () => askDeleteProject(menuProject),
          onClose: () => setMenuProject(null),}
        )
      )

      , editingProject && (
        _jsx(ProjectSheet, {
          project: editingProject.project,
          canDelete: projectCount > 1,
          onSave: saveProject,
          onDelete: deleteProject,
          onClose: () => setEditingProject(null),}
        )
      )

      , exporting && (
        _jsx(ExportSheet, {
          projects: projects,
          records: records,
          defaultProjectId: activeProject.id,
          defaultMonth: monthKey,
          onExport: runExport,
          onExportData: exportData,
          onImport: () => importRef.current && importRef.current.click(),
          onExportImage: runImageExport,
          onClose: () => setExporting(false),}
        )
      )

      , batch && (
        _jsx(BatchSheet, {
          projectId: activeProject.id,
          projectName: activeProject.name,
          defaultStart: batch.start,
          defaultEnd: batch.end,
          records: records,
          onSave: saveBatch,
          onClose: () => setBatch(null),}
        )
      )

      , _jsx('input', {
        ref: importRef,
        type: "file",
        accept: ".json,application/json",
        hidden: true,
        onChange: importData,}
      )

      , booting && _jsx(Splash, { version: APP_VERSION, onDone: () => setBooting(false),} )
    ]})
  );
}

/// 首页顶部：所有项目按农历年算的应得总金额。
function HomeSummary({ projects, records, yearStart, yearEnd }) {
  let wage = 0;
  let paid = 0;

  for (const project of projects) {
    const rates = projectRates(project);
    const own = records.filter(
      (record) =>
        record.projectId === project.id && record.date >= yearStart && record.date <= yearEnd
    );
    wage += own.reduce((sum, record) => sum + recordWage(record, rates.daily, rates.otHourly), 0);
    paid += own.reduce((sum, record) => sum + (Number(record.paid) || 0), 0);
  }

  const totalWage = round2(wage);
  const totalPaid = round2(paid);
  // 应得 = 本年度工资 − 支账（已发的生活费）。
  const balance = round2(totalWage - totalPaid);
  const negative = balance < 0;
  const shownText = moneyFmt.format(Math.abs(negative ? totalPaid : balance));
  const shownSize = Math.min(60, Math.floor(535 / Math.max(shownText.length, 1)));

  return (
    _jsxs('section', { className: "summary homesum" , children: [
      _jsx('p', { className: "summary-label", children: negative ? '支账' : '本年度总应得余额'})
      , _jsx('div', { className: "summary-number", children: 
        _jsxs('b', { style: { fontSize: `min(14vw, ${shownSize}px)` }, children: [
          _jsx('span', { className: "cur", children: "¥"})
          , shownText
          , _jsx('span', { className: "unit", children: "元"})
        ]})
      })
      , _jsx('p', { className: "homesum-split", children: 
        `${projects.length} 个项目 · 总工资 ¥${moneyFmt.format(totalWage)} · 支账 ¥${moneyFmt.format(
          totalPaid
        )}`
      })
    ]})
  );
}

/// 数字变化时从当前值缓动到目标值，用来做余额的滚动动画。
function useAnimatedNumber(target, duration = 820) {
  const [display, setDisplay] = useState(target);
  const current = useRef(target);

  useEffect(() => {
    const from = current.current;
    if (from === target) return;
    const start = performance.now();
    let raf = 0;
    const tick = (now) => {
      const t = Math.min(1, (now - start) / duration);
      // easeInOutCubic：起步和收尾都轻，不会一下子窜出去
      const eased = t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
      const next = from + (target - from) * eased;
      current.current = next;
      setDisplay(next);
      if (t < 1) {
        raf = requestAnimationFrame(tick);
      } else {
        current.current = target;
        setDisplay(target);
      }
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [target, duration]);

  return display;
}

/// 开屏：图标、名称、版本，约 1.5 秒后淡出到首页。
function Splash({ version, onDone }) {
  const [leaving, setLeaving] = useState(false);

  useEffect(() => {
    const fade = setTimeout(() => setLeaving(true), 1050);
    const done = setTimeout(onDone, 1500);
    return () => {
      clearTimeout(fade);
      clearTimeout(done);
    };
  }, [onDone]);

  return (
    _jsxs('div', {
      className: `splash${leaving ? ' leaving' : ''}`,
      onClick: onDone,
      role: "presentation",
 children: [
      _jsx('span', { className: "splash-mark", 'aria-hidden': "true", children: "📒"})
      , _jsx('p', { className: "splash-name", children: "记工"})
      , _jsx('p', { className: "splash-version", children: `v${version}`})
    ]})
  );
}

/// 长按（约 450ms）触发菜单；按在按钮上、或者手指滑动时不触发。
function useLongPress(onLongPress, delay = 450) {
  const timer = useRef(null);
  const origin = useRef(null);
  const suppressed = useRef(false);

  const clear = () => {
    if (timer.current) {
      clearTimeout(timer.current);
      timer.current = null;
    }
  };

  useEffect(() => clear, []);

  if (!onLongPress) return { suppressed, pressProps: {} };

  const pressProps = {
    onPointerDown: (event) => {
      if (event.target.closest('button')) return;
      suppressed.current = false;
      origin.current = { x: event.clientX, y: event.clientY };
      clear();
      timer.current = setTimeout(() => {
        timer.current = null;
        suppressed.current = true;
        onLongPress();
      }, delay);
    },
    onPointerMove: (event) => {
      if (!timer.current || !origin.current) return;
      const dx = event.clientX - origin.current.x;
      const dy = event.clientY - origin.current.y;
      if (Math.hypot(dx, dy) > 10) clear();
    },
    onPointerUp: clear,
    onPointerCancel: clear,
    onPointerLeave: clear,
    onContextMenu: (event) => {
      event.preventDefault();
      suppressed.current = true;
      onLongPress();
    },
  };

  return { suppressed, pressProps };
}

/// 首页项目卡：左右滑动露出删除。
function SwipeRow({ children, onDelete, disabled }) {
  const [dx, setDx] = useState(0);
  const [dragging, setDragging] = useState(false);
  const start = useRef(null);
  const moved = useRef(false);
  const opened = useRef(0);

  const down = (event) => {
    if (disabled) return;
    if (event.target.closest('button')) return;
    start.current = { x: event.clientX, y: event.clientY, base: opened.current * 92 };
    moved.current = false;
  };
  const move = (event) => {
    if (!start.current) return;
    const mx = event.clientX - start.current.x;
    const my = event.clientY - start.current.y;
    if (!moved.current) {
      if (Math.abs(mx) < 12 || Math.abs(mx) < Math.abs(my)) return;
      moved.current = true;
      setDragging(true);
    }
    setDx(Math.max(-120, Math.min(120, start.current.base + mx)));
  };
  const up = () => {
    if (!start.current) return;
    start.current = null;
    if (!moved.current) return;
    const next = dx > 56 ? 1 : dx < -56 ? -1 : 0;
    opened.current = next;
    setDx(next * 92);
    setDragging(false);
  };
  const clickCapture = (event) => {
    if (event.target.closest('.swipe-del')) return;
    if (moved.current) {
      event.stopPropagation();
      event.preventDefault();
      moved.current = false;
      return;
    }
    if (opened.current !== 0) {
      event.stopPropagation();
      event.preventDefault();
      opened.current = 0;
      setDx(0);
    }
  };

  return (
    _jsxs('div', { className: "swipe-row", onPointerDown: down, onPointerMove: move, onPointerUp: up, onPointerCancel: up, onPointerLeave: up, onClickCapture: clickCapture, children: [
      _jsx('button', { type: "button", className: "swipe-del", onClick: onDelete, children: [
        _jsx('span', { children: "删除"})
        , _jsx('span', { children: "删除"})
      ]})
      , _jsx('div', { className: `swipe-content${dragging ? ' dragging' : ''}`, style: { transform: `translateX(${dx}px)` }, children: children })
    ]})
  );
}

function ProjectCard({ project, records, onOpen, onEdit, onMenu, active, showState }) {
  const { suppressed, pressProps } = useLongPress(onMenu);

  const rates = projectRates(project);
  const own = useMemo(
    () => records.filter((record) => record.projectId === project.id),
    [records, project.id]
  );

  const { main, sub } = splitProjectName(project.name);
  const rateText =
    rates.daily > 0
      ? `日薪 ¥${amountFmt.format(rates.daily)} · 时薪 ¥${amountFmt.format(
          rates.hourly
        )} · 加班时薪 ¥${amountFmt.format(rates.otHourly)}`
      : '设置日薪和工时';
  const allWage = own.reduce((sum, record) => sum + recordWage(record, rates.daily, rates.otHourly), 0);
  const allPaid = own.reduce((sum, record) => sum + (Number(record.paid) || 0), 0);
  const allDays = own.reduce((sum, record) => sum + (Number(record.days) || 0), 0);
  const allHours = own.reduce((sum, record) => sum + overtimeHours(record.overtime), 0);
  const balance = round2(allWage - allPaid);
  // 余额为负时不再显示「多付工资」，改为显示这段时间的支账（已发）总额。
  const targetBalance = balance >= 0 ? Math.abs(balance) : Math.abs(allPaid);
  const animatedBalance = useAnimatedNumber(targetBalance);
  const displayText = moneyFmt.format(animatedBalance);
  const targetText = moneyFmt.format(targetBalance);
  const displaySize = Math.min(60, Math.floor(535 / Math.max(targetText.length, 1)));

  return (
    _jsxs('article', {
      className: `summary card${onOpen ? ' openable' : ''}`,
      ...pressProps,
      onClick: () => {
        if (suppressed.current) {
          suppressed.current = false;
          return;
        }
        if (onOpen) onOpen();
      },
 children: [
      showState && (
        _jsx('span', { className: `project-state corner${active ? ' on' : ''}`, children: active ? '在职' : '离职'})
      )

      , _jsxs('div', { className: `project-block${project.name ? '' : ' unset'}`, children: [
        _jsx('span', { className: "name", children: project.name ? main : '设置项目名称'})
        , sub && _jsx('span', { className: "sub", children: sub})
      ]})

      , _jsx('div', { className: "summary-divider", 'aria-hidden': "true",} )

      , _jsx('p', { className: "summary-total", children: `本项目累计 ${trim(allDays)} 个工日 · 加班 ${trim(allHours)} 小时`})

      , onEdit ? (
        _jsxs('button', {
          type: "button",
          className: "rate-line",
          onClick: (event) => {
            event.stopPropagation();
            onEdit();
          },
 children: [
          rateText
          , _jsx('svg', { viewBox: "0 0 24 24"   , 'aria-hidden': "true", children: 
            _jsx('path', { d: "M9 6l6 6-6 6"   , fill: "none", stroke: "currentColor", strokeWidth: "2.6", strokeLinecap: "round", strokeLinejoin: "round",} )
          })
        ]})
      ) : (
        _jsx('p', { className: "rate-line", children: rateText})
      )

      , _jsx('p', { className: "summary-split", children: 
        `本项目总工资 ¥${moneyFmt.format(round2(allWage))} · 支账 ¥${moneyFmt.format(
          round2(allPaid)
        )}`
      })

      , _jsx('p', { className: "summary-label", children: balance >= 0 ? '应得工资余额' : '支账'})
      , _jsx('div', { className: "summary-number", children: 
        _jsxs('b', { style: { fontSize: `min(14vw, ${displaySize}px)` }, children: [
          _jsx('span', { className: "cur", children: "¥"})
          , displayText
          , _jsx('span', { className: "unit", children: "元"})
        ]})
      })

    ]})
  );
}

function Row({ record, daily, otHourly, onOpen }) {
  const date = fromISO(record.date);
  const isWork = record.status === 'work';
  const isLeave = record.status === 'leave';
  const isPayment = record.note === 'wage' || record.note === 'living';
  const note =
    record.note === 'am'
      ? '上午上班'
      : record.note === 'pm'
        ? '下午上班'
        : record.note === 'wage' || record.note === 'living'
          ? '生活费'
          : '无';
  const paid = Number(record.paid) || 0;
  const hours = overtimeHours(record.overtime);

  return (
    _jsxs('button', { className: "row", onClick: onOpen, children: [
      _jsxs('span', { className: "cell-date", children: [date.getDate(), "日"]})
      , _jsx('span', { className: "cell-status", children: 
        _jsx('span', { className: `badge ${isWork ? 'work' : isLeave ? 'leave' : 'rest'}`, children: 
          isWork ? '上班' : isLeave ? '离职' : '休息'
        })
      })
      , _jsx('span', { className: "cell-days", children: trim(Number(record.days) || 0)})
      , _jsx('span', { className: `cell-ot${hours > 0 ? ' on' : ''}`, children: 
        hours > 0 ? `${trim(hours)}h` : '—'
      })
      , _jsx('span', { className: "cell-wage", children: yuan(recordWage(record, daily, otHourly))})
      , _jsxs('span', { className: "cell-note", children: [
        note
        , isPayment && paid > 0 && _jsxs('small', { children: ["¥", amountFmt.format(paid)]})
      ]})
    ]})
  );
}

function Editor({ record, records, daily, otHourly, onSave, onDelete, onClose }) {
  const isNew = !record.id;
  const initialStatus =
    record.status === 'rest' ? 'rest' : record.status === 'leave' ? 'leave' : 'work';
  const [date, setDate] = useState(record.date || todayISO());
  const [status, setStatus] = useState(initialStatus);
  const [days, setDays] = useState(
    initialStatus === 'work' ? (record.days === 0.5 ? 0.5 : 1) : 0
  );
  const [note, setNote] = useState(record.note || '');
  const [paidText, setPaidText] = useState(record.paid > 0 ? String(record.paid) : '');
  const [overtime, setOvertime] = useState(Boolean(record.overtime && record.overtime.start));
  const [start, setStart] = useState(record.overtime && record.overtime.start ? record.overtime.start : '18:00');
  const [end, setEnd] = useState(record.overtime && record.overtime.end ? record.overtime.end : '20:00');
  const [shown, setShown] = useState(false);

  useEffect(() => {
    const frame = requestAnimationFrame(() => setShown(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  /// 新建时提示：这个月在这天之前，还有哪几天没记。
  const missingDays = useMemo(() => {
    if (!isNew || !date) return [];
    const [year, month, day] = date.split('-').map(Number);
    if (!year || !month || !day) return [];
    const prefix = `${year}-${pad(month)}-`;
    const taken = new Set(
      records
        .filter((item) => item.projectId === record.projectId && item.date.startsWith(prefix))
        .map((item) => item.date)
    );
    const today = todayISO();
    const list = [];
    for (let index = 1; index < day; index += 1) {
      const key = `${prefix}${pad(index)}`;
      // 只提示已经过去的日子，选到未来的日期时不要把还没到的算成漏记。
      if (key < today && !taken.has(key)) list.push(key);
    }
    return list;
  }, [date, records, record.projectId, isNew]);

  const isPayment = note === 'wage' || note === 'living';
  const paid = Math.max(0, Number(paidText) || 0);
  const hours = overtime ? overtimeHours({ start, end }) : 0;
  const wage = days * daily + hours * otHourly;
  const valid =
    Boolean(date) &&
    (!overtime || (Boolean(start) && Boolean(end))) &&
    (!isPayment || paid > 0);
  const collision = records.find(
    (item) =>
      item.projectId === record.projectId && item.date === date && item.id !== record.id
  );

  /// 状态、备注一变，工日跟着给一个合理的数，之后还能在 0.5 / 1 里选。
  const changeStatus = (next) => {
    setStatus(next);
    if (next === 'rest') {
      setDays(0);
      if (note === 'am' || note === 'pm') setNote('');
    } else if (next === 'leave') {
      // 离职当天不记工日，默认只记一笔发工资。
      setDays(0);
      if (note !== 'wage' && note !== 'living') setNote('living');
    } else {
      setDays(isPayment ? 1 : note ? 0.5 : 1);
    }
  };

  const changeNote = (next) => {
    const payment = next === 'wage' || next === 'living';
    setNote(next);
    if (status === 'work') setDays(payment ? 1 : next ? 0.5 : 1);
    if (!payment) setPaidText('');
  };

  const submit = (event) => {
    event.preventDefault();
    if (!valid) return;
    if (collision && !confirm('这一天已有记录，保存会覆盖原来那条。继续？')) return;
    onSave({
      id: record.id || `${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      projectId: record.projectId,
      date,
      status,
      days: status === 'work' ? days : 0,
      note: status === 'rest' && (note === 'am' || note === 'pm') ? '' : note,
      paid: isPayment ? round2(paid) : 0,
      overtime: overtime ? { start, end } : null,
    });
  };

  const askDelete = () => {
    if (confirm('删除这条记录？')) onDelete(record.id);
  };

  return (
    _jsx('div', { className: `sheet-backdrop${shown ? ' shown' : ''}`, onClick: onClose, children: 
      _jsxs('form', { className: "sheet", onClick: (event) => event.stopPropagation(), onSubmit: submit, children: [
        _jsx('div', { className: "sheet-grabber", 'aria-hidden': "true",} )
        , _jsxs('header', { className: "sheet-bar", children: [
          _jsx('button', { type: "button", className: "link", onClick: onClose, children: "取消"})
          , _jsx('strong', { children: isNew ? '记一笔' : '编辑记录'})
          , _jsx('button', { type: "submit", className: "link strong" , disabled: !valid, children: "保存"})
        ]})

        , missingDays.length > 0 && (
          _jsx('p', { className: "sheet-note", children: 
            `前面还有 ${missingDays.length} 天没记：${describeMissingDays(missingDays)}`
          })
        )

        , _jsxs('div', { className: "group", children: [
          _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "日期"})
            , _jsx('input', {
              type: "date",
              value: date,
              onChange: (event) => setDate(event.target.value),
              required: true,}
            )
          ]})

          , _jsxs('div', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "状态"})
            , _jsx('span', { className: `select-wrap ${status}`, children: 
              _jsxs('select', { 'aria-label': "状态", value: status, onChange: (event) => changeStatus(event.target.value), children: [
                _jsx('option', { value: "work", children: "上班"})
                , _jsx('option', { value: "rest", children: "休息"})
                , _jsx('option', { value: "leave", children: "离职"})
              ]})
            })
          ]})

          , _jsxs('div', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "工日"})
            , status !== 'work' ? (
              _jsx('span', { className: "field-value", children: "0"})
            ) : (
              _jsxs('div', { className: "segmented", role: "group", 'aria-label': "工日", children: [
                _jsx('button', { type: "button", className: days === 0.5 ? 'on' : '', onClick: () => setDays(0.5), children: "0.5"})
                , _jsx('button', { type: "button", className: days === 1 ? 'on' : '', onClick: () => setDays(1), children: "1"})
              ]})
            )
          ]})

          , _jsxs('div', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "备注"})
            , _jsx('span', { className: "select-wrap note" , children: 
              _jsxs('select', {
                'aria-label': "备注",
                value: note,
                onChange: (event) => changeNote(event.target.value),
 children: [
                _jsx('option', { value: "", children: "无"})
                , _jsx('option', { value: "am", disabled: status !== 'work', children: "上午上班"})
                , _jsx('option', { value: "pm", disabled: status !== 'work', children: "下午上班"})
                , _jsx('option', { value: "living", children: "生活费"})
              ]})
            })
          ]})

          , isPayment && (
            _jsxs('label', { className: "field", children: [
              _jsx('span', { className: "field-label", children: "金额"})
              , _jsx('input', {
                type: "number",
                inputMode: "decimal",
                min: "0",
                step: "0.01",
                value: paidText,
                onChange: (event) => setPaidText(event.target.value),
                placeholder: "0",}
              )
              , _jsx('span', { className: "unit", children: "元"})
            ]})
          )
        ]})

        , collision && (
          _jsx('p', { className: "caption warn" , children: "这一天已有记录，保存会覆盖原来那条。"})
        )

        , _jsxs('div', { className: "group", children: [
          _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "加班"})
            , _jsx('span', { className: "switch", children: 
              _jsx('input', {
                type: "checkbox",
                checked: overtime,
                onChange: (event) => setOvertime(event.target.checked),}
              )
            })
          ]})

          , overtime && (
            _jsxs(_Fragment, { children: [
              _jsxs('label', { className: "field", children: [
                _jsx('span', { className: "field-label", children: "开始"})
                , _jsx('input', {
                  type: "time",
                  value: start,
                  onChange: (event) => setStart(event.target.value),
                  required: true,}
                )
              ]})
              , _jsxs('label', { className: "field", children: [
                _jsx('span', { className: "field-label", children: "结束"})
                , _jsx('input', {
                  type: "time",
                  value: end,
                  onChange: (event) => setEnd(event.target.value),
                  required: true,}
                )
              ]})
              , _jsxs('div', { className: "field", children: [
                _jsx('span', { className: "field-label", children: "小时数"})
                , _jsxs('span', { className: "field-value", children: [trim(hours), " 小时" ]})
              ]})
            ]})
          )
        ]})

        , _jsxs('div', { className: "wage-preview", children: [
          _jsxs('div', { className: "wage-copy", children: [
            _jsx('span', { children: "本条工资"})
            , _jsxs('small', { children: ["日薪 ¥"
               , amountFmt.format(daily), " · 加班时薪 ¥"   , amountFmt.format(otHourly)
            ]})
            , isPayment && (
              _jsxs('small', { className: "pay-line", children: ["另记支账 ¥" , amountFmt.format(paid)]})
            )
          ]})
          , _jsx('b', { children: yuan(wage)})
        ]})

        , _jsx('p', { className: "caption", children: "上午/下午上班自动记 0.5 工日；工日只能选 0.5 或 1。离职当天不记工日，只记生活费。生活费要填金额，单独算支账。"

        })

        , !isNew && (
          _jsx('button', { type: "button", className: "remove", onClick: askDelete, children: "删除这条"})
        )
      ]})
    })
  );
}

function ProjectSheet({ project, canDelete, onSave, onDelete, onClose }) {
  const isNew = !project;
  const [name, setName] = useState(project ? project.name : '');
  const [dailyText, setDailyText] = useState(project && project.daily > 0 ? String(project.daily) : '');
  const [normText, setNormText] = useState(
    project ? String(project.normHours) : String(DEFAULT_NORM_HOURS)
  );
  const [otText, setOtText] = useState(project ? String(project.otHours) : String(DEFAULT_OT_HOURS));
  const [shown, setShown] = useState(false);

  useEffect(() => {
    const frame = requestAnimationFrame(() => setShown(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  const dailyValue = dailyText.trim() === '' ? 0 : Number(dailyText);
  const normValue = normText.trim() === '' ? DEFAULT_NORM_HOURS : Number(normText);
  const otValue = otText.trim() === '' ? DEFAULT_OT_HOURS : Number(otText);
  const valid =
    Number.isFinite(dailyValue) &&
    dailyValue >= 0 &&
    Number.isFinite(normValue) &&
    normValue > 0 &&
    normValue <= 24 &&
    Number.isFinite(otValue) &&
    otValue > 0 &&
    otValue <= 24;
  const hourly = round2(Math.max(0, dailyValue) / (normValue > 0 ? normValue : DEFAULT_NORM_HOURS));
  const otHourly = round2(Math.max(0, dailyValue) / (otValue > 0 ? otValue : DEFAULT_OT_HOURS));

  const submit = (event) => {
    event.preventDefault();
    if (!valid) return;
    onSave({
      id: project ? project.id : null,
      name: name.trim(),
      daily: dailyValue,
      normHours: normValue,
      otHours: otValue,
    });
  };

  const askDelete = () => {
    if (confirm('删除这个项目和它的全部记录？删了就找不回来了。')) onDelete(project);
  };

  return (
    _jsx('div', { className: `sheet-backdrop${shown ? ' shown' : ''}`, onClick: onClose, children: 
      _jsxs('form', { className: "sheet", onClick: (event) => event.stopPropagation(), onSubmit: submit, children: [
        _jsx('div', { className: "sheet-grabber", 'aria-hidden': "true",} )
        , _jsxs('header', { className: "sheet-bar", children: [
          _jsx('button', { type: "button", className: "link", onClick: onClose, children: "取消"})
          , _jsx('strong', { children: isNew ? '新建项目' : '项目设置'})
          , _jsx('button', { type: "submit", className: "link strong" , disabled: !valid, children: "保存"})
        ]})

        , _jsxs('div', { className: "group", children: [
          _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "项目"})
            , _jsx('input', {
              type: "text",
              value: name,
              onChange: (event) => setName(event.target.value),
              placeholder: "例如：××工地",
              maxLength: 80,}
            )
          ]})

          , _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "日薪"})
            , _jsx('input', {
              type: "number",
              inputMode: "decimal",
              min: "0",
              step: "0.01",
              value: dailyText,
              onChange: (event) => setDailyText(event.target.value),
              placeholder: "0",}
            )
            , _jsx('span', { className: "unit", children: "元/天"})
          ]})

          , _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "正常工时"})
            , _jsx('input', {
              type: "number",
              inputMode: "decimal",
              min: "1",
              max: "24",
              step: "0.5",
              value: normText,
              onChange: (event) => setNormText(event.target.value),}
            )
            , _jsx('span', { className: "unit", children: "小时/天"})
          ]})

          , _jsxs('div', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "时薪"})
            , _jsx('span', { className: "field-value", children: `¥${amountFmt.format(hourly)} 元/时`})
          ]})

          , _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "加班折算"})
            , _jsx('input', {
              type: "number",
              inputMode: "decimal",
              min: "1",
              max: "24",
              step: "0.5",
              value: otText,
              onChange: (event) => setOtText(event.target.value),}
            )
            , _jsx('span', { className: "unit", children: "小时/工"})
          ]})

          , _jsxs('div', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "加班时薪"})
            , _jsx('span', { className: "field-value", children: `¥${amountFmt.format(otHourly)} 元/时`})
          ]})
        ]})

        , _jsx('p', { className: "caption", children: 
          `时薪 = 日薪 ÷ 正常工时（${trim(normValue)} 小时），加班时薪 = 日薪 ÷ 加班折算（${trim(
            otValue
          )} 小时，加班 ${trim(otValue)} 小时算一个工日）。改日薪或这两个小时数，上面两个时薪和所有月份的工资都会跟着重算。`
        })

        , !isNew && canDelete && (
          _jsx('button', { type: "button", className: "remove", onClick: askDelete, children: "删除项目"})
        )
      ]})
    })
  );
}

function ExportSheet({
  projects,
  records,
  defaultProjectId,
  defaultMonth,
  onExport,
  onExportData,
  onImport,
  onExportImage,
  onClose,
}) {
  const activeProject =
    projects.find((project) => project.id === defaultProjectId) || projects[0];
  const activeName = activeProject ? activeProject.name || '未命名项目' : '当前项目';

  const [who, setWho] = useState('current');
  const [picked, setPicked] = useState([defaultProjectId]);
  const [period, setPeriod] = useState('pick');
  const [month, setMonth] = useState(defaultMonth);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [activeEnd, setActiveEnd] = useState('from');
  const [calMonth, setCalMonth] = useState('');
  const [format, setFormat] = useState('docx');
  const [fields, setFields] = useState({
    days: true,
    overtime: true,
    wage: true,
    paid: true,
    balance: true,
    note: true,
  });
  const [shown, setShown] = useState(false);

  useEffect(() => {
    const frame = requestAnimationFrame(() => setShown(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  // 当前：当前项目 + 当前月份；全部：当前项目 + 全部时间；自选：项目和时间都自己挑。
  const chosenIds = who === 'pick' ? picked : [defaultProjectId];
  const chosenKey = chosenIds.join('|');

  /// 选中的项目里，哪些日子真有记录——只有这些日子能选。
  const dateSet = useMemo(() => {
    const ids = new Set(chosenKey.split('|'));
    const set = new Set();
    records.forEach((record) => {
      if (ids.has(record.projectId)) set.add(record.date);
    });
    return set;
  }, [records, chosenKey]);

  const dateKeys = useMemo(() => [...dateSet].sort(), [dateSet]);
  const monthsWithRecords = useMemo(
    () => [...new Set(dateKeys.map((date) => date.slice(0, 7)))].sort(),
    [dateKeys]
  );

  const firstDate = dateKeys[0] || '';
  const lastDate = dateKeys[dateKeys.length - 1] || '';
  // 自定义默认先框住当前看的那一个月；那个月没记录就用全部范围。
  const monthDates = dateKeys.filter((date) => date.slice(0, 7) === defaultMonth);
  const defaultFrom = monthDates[0] || firstDate;
  const defaultTo = monthDates[monthDates.length - 1] || lastDate;
  // 换项目后原来的起止可能已经没有记录，这时退回默认范围。
  const fromValue = dateSet.has(from) ? from : defaultFrom;
  const toValue = dateSet.has(to) ? to : defaultTo;
  const rangeLo = fromValue && toValue && fromValue > toValue ? toValue : fromValue;
  const rangeHi = fromValue && toValue && fromValue > toValue ? fromValue : toValue;

  // 「全部」按当前查看项目的首末记录标出时间跨度。
  const ownDates = useMemo(
    () =>
      records
        .filter((record) => record.projectId === defaultProjectId)
        .map((record) => record.date)
        .sort(),
    [records, defaultProjectId]
  );
  const spanLabel = ownDates.length
    ? `${dateLabel(ownDates[0])}${
        ownDates[ownDates.length - 1] !== ownDates[0]
          ? `–${dateLabel(ownDates[ownDates.length - 1])}`
          : ''
      }`
    : '全部时间';

  // 自选「全部时间」时，用所选项目的首末记录标出跨度。
  const chosenSpanDates = useMemo(
    () =>
      records
        .filter((record) => chosenIds.includes(record.projectId))
        .map((record) => record.date)
        .sort(),
    [records, chosenKey]
  );
  const chosenSpan =
    chosenSpanDates.length > 0
      ? `${dateLabel(chosenSpanDates[0])}${
          chosenSpanDates[chosenSpanDates.length - 1] !== chosenSpanDates[0]
            ? `–${dateLabel(chosenSpanDates[chosenSpanDates.length - 1])}`
            : ''
        }`
      : '全部时间';

  const monthValue = monthsWithRecords.includes(month)
    ? month
    : monthsWithRecords[0] || defaultMonth;

  const pickStart =
    period === 'custom' ? rangeLo : period === 'all' ? '0000-01-01' : `${monthValue}-01`;
  const pickEnd =
    period === 'custom' ? rangeHi : period === 'all' ? '9999-12-31' : `${monthValue}-31`;
  const pickLabel =
    period === 'all'
      ? chosenSpan
      : period === 'custom'
        ? !rangeLo || !rangeHi
          ? '自定义时间'
          : rangeLo === rangeHi
            ? dateLabel(rangeLo)
            : `${dateLabel(rangeLo)}–${dateLabel(rangeHi)}`
        : monthLabel(monthValue);

  const start =
    who === 'current' ? `${defaultMonth}-01` : who === 'all' ? '0000-01-01' : pickStart;
  const end = who === 'current' ? `${defaultMonth}-31` : who === 'all' ? '9999-12-31' : pickEnd;
  const label =
    who === 'current' ? monthLabel(defaultMonth) : who === 'all' ? spanLabel : pickLabel;

  const chosenProjects = projects.filter((project) => chosenIds.includes(project.id));
  const counts = chosenProjects.map((project) => ({
    project,
    count: records.filter(
      (record) => record.projectId === project.id && inRange(record.date, start, end)
    ).length,
  }));
  const withData = counts.filter((item) => item.count > 0);
  const count = counts.reduce((sum, item) => sum + item.count, 0);

  // 自定义区间的小日历：只亮有记录的日子。
  const calAnchor = (activeEnd === 'from' ? fromValue : toValue).slice(0, 7);
  const calMonthValue = monthsWithRecords.includes(calMonth)
    ? calMonth
    : monthsWithRecords.includes(calAnchor)
      ? calAnchor
      : monthsWithRecords[0] || '';
  const calIndex = monthsWithRecords.indexOf(calMonthValue);

  const calCells = useMemo(() => {
    if (!calMonthValue) return [];
    const [year, monthNumber] = calMonthValue.split('-').map(Number);
    const total = new Date(year, monthNumber, 0).getDate();
    const lead = (new Date(year, monthNumber - 1, 1).getDay() + 6) % 7;
    const cells = [];
    for (let index = 0; index < lead; index += 1) cells.push(null);
    for (let day = 1; day <= total; day += 1) {
      cells.push(`${year}-${pad(monthNumber)}-${pad(day)}`);
    }
    return cells;
  }, [calMonthValue]);

  /// 有记录才可选；正在改开始，就不能晚于结束；正在改结束，就不能早于开始。
  const isSelectable = (day) => {
    if (!dateSet.has(day)) return false;
    if (activeEnd === 'from' && toValue && day > toValue) return false;
    if (activeEnd === 'to' && fromValue && day < fromValue) return false;
    return true;
  };

  const pickDay = (day) => {
    if (!isSelectable(day)) return;
    if (activeEnd === 'from') {
      setFrom(day);
      if (toValue && day > toValue) setTo(day);
      setActiveEnd('to');
    } else {
      setTo(day);
    }
  };

  const focusEnd = (which) => {
    setActiveEnd(which);
    const target = which === 'from' ? fromValue : toValue;
    if (target) setCalMonth(target.slice(0, 7));
  };

  const stepMonth = (delta) => {
    const next = monthsWithRecords[calIndex + delta];
    if (next) setCalMonth(next);
  };

  const toggle = (id) => {
    setPicked((list) => (list.includes(id) ? list.filter((item) => item !== id) : [...list, id]));
  };

  const toggleField = (key) => setFields((prev) => ({ ...prev, [key]: !prev[key] }));

  const submit = () => {
    if (count === 0) return;
    const payload = { projectIds: chosenIds, start, end, label, fields };
    if (format === 'image') {
      Promise.resolve(onExportImage(payload)).then((ok) => {
        if (ok) onClose();
      });
      return;
    }
    onExport(payload);
  };

  return (
    _jsx('div', { className: `sheet-backdrop${shown ? ' shown' : ''}`, onClick: onClose, children: 
      _jsxs('div', { className: "sheet", onClick: (event) => event.stopPropagation(), children: [
        _jsx('div', { className: "sheet-grabber", 'aria-hidden': "true",} )
        , _jsxs('header', { className: "sheet-bar", children: [
          _jsx('button', { type: "button", className: "link", onClick: onClose, children: "取消"})
          , _jsx('strong', { children: "导出"})
          , _jsx('button', { type: "button", className: "link strong" , disabled: count === 0, onClick: submit, children: "导出"

          })
        ]})

        , _jsxs('div', { className: "group", children: [
          _jsxs('div', { className: "field choices-field" , children: [
            _jsx('span', { className: "field-label", children: "范围"})
            , _jsxs('span', { className: "choices", children: [
              _jsx('button', {
                type: "button",
                className: who === 'current' ? 'on' : '',
                onClick: () => setWho('current'),
 children: "当前"

              })
              , _jsx('button', {
                type: "button",
                className: who === 'all' ? 'on' : '',
                onClick: () => setWho('all'),
 children: "全部"

              })
              , _jsx('button', {
                type: "button",
                className: who === 'pick' ? 'on' : '',
                onClick: () => setWho('pick'),
 children: "自选"

              })
            ]})
          ]})

          , _jsxs('div', { className: "field choices-field", children: [
            _jsx('span', { className: "field-label", children: "格式"})
            , _jsxs('span', { className: "choices", children: [
              _jsx('button', { type: "button", className: format === 'docx' ? 'on' : '', onClick: () => setFormat('docx'), children: "Word"})
              , _jsx('button', { type: "button", className: format === 'image' ? 'on' : '', onClick: () => setFormat('image'), children: "图片"})
            ]})
          ]})

          , _jsxs('div', { className: "field choices-field", children: [
            _jsx('span', { className: "field-label", children: "内容"})
            , _jsxs('span', { className: "choices compact", children: [
              _jsx('button', { type: "button", className: fields.days ? 'on' : '', onClick: () => toggleField('days'), children: "工日"})
              , _jsx('button', { type: "button", className: fields.overtime ? 'on' : '', onClick: () => toggleField('overtime'), children: "加班"})
              , _jsx('button', { type: "button", className: fields.wage ? 'on' : '', onClick: () => toggleField('wage'), children: "工资"})
              , _jsx('button', { type: "button", className: fields.paid ? 'on' : '', onClick: () => toggleField('paid'), children: "支账"})
              , _jsx('button', { type: "button", className: fields.balance ? 'on' : '', onClick: () => toggleField('balance'), children: "余额"})
              , _jsx('button', { type: "button", className: fields.note ? 'on' : '', onClick: () => toggleField('note'), children: "备注"})
            ]})
          ]})

          , who === 'pick' && (
            _jsx('div', { className: "checks", children: 
              projects.map((project) => (
                _jsxs('label', { className: "check-row", children: [
                  _jsx('input', {
                    type: "checkbox",
                    checked: picked.includes(project.id),
                    onChange: () => toggle(project.id),}
                  )
                  , _jsx('span', { className: "check-name", children: project.name || '未命名项目'})
                  , _jsx('span', { className: "check-rate", children: 
                    project.daily > 0 ? `¥${amountFmt.format(project.daily)}/天` : '没设日薪'
                  })
                ]}, project.id)
              ))
            })
          )
        ]})

        , _jsx('div', { className: "group", children: 
          who === 'pick' ? (
            _jsxs(_Fragment, { children: [
              _jsxs('div', { className: "field choices-field" , children: [
                _jsx('span', { className: "field-label", children: "时间"})
                , _jsxs('span', { className: "choices", children: [
                  _jsx('button', {
                    type: "button",
                    className: period === 'pick' ? 'on' : '',
                    onClick: () => setPeriod('pick'),
 children: "选月份"

                  })
                  , _jsx('button', {
                    type: "button",
                    className: period === 'custom' ? 'on' : '',
                    onClick: () => setPeriod('custom'),
 children: "自定义"

                  })
                  , _jsx('button', {
                    type: "button",
                    className: period === 'all' ? 'on' : '',
                    onClick: () => setPeriod('all'),
 children: "全部时间"

                  })
                ]})
              ]})

              , period === 'pick' && (
                _jsxs('div', { className: "field", children: [
                  _jsx('span', { className: "field-label", children: "月份"})
                  , _jsx('span', { className: "select-wrap note" , children: 
                    _jsx('select', {
                      value: monthValue,
                      onChange: (event) => setMonth(event.target.value),
                      disabled: monthsWithRecords.length === 0,
 children: 
                      monthsWithRecords.map((key) => (
                        _jsx('option', { value: key, children: monthLabel(key)}, key)
                      ))
                    })
                  })
                ]})
              )

              , period === 'custom' && (
                _jsxs(_Fragment, { children: [
                  _jsxs('div', { className: "field", children: [
                    _jsx('span', { className: "field-label", children: "开始"})
                    , _jsx('button', {
                      type: "button",
                      className: `date-chip${activeEnd === 'from' ? ' on' : ''}`,
                      onClick: () => focusEnd('from'),
 children: 
                      fromValue ? dateLabel(fromValue) : '没有可选日期'
                    })
                  ]})
                  , _jsxs('div', { className: "field", children: [
                    _jsx('span', { className: "field-label", children: "结束"})
                    , _jsx('button', {
                      type: "button",
                      className: `date-chip${activeEnd === 'to' ? ' on' : ''}`,
                      onClick: () => focusEnd('to'),
 children: 
                      toValue ? dateLabel(toValue) : '没有可选日期'
                    })
                  ]})

                  , calCells.length > 0 && (
                    _jsxs('div', { className: "calendar", children: [
                      _jsxs('div', { className: "cal-head", children: [
                        _jsx('button', {
                          type: "button",
                          className: "cal-nav",
                          'aria-label': "上个月",
                          disabled: calIndex <= 0,
                          onClick: () => stepMonth(-1),
 children: "‹"

                        })
                        , _jsx('strong', { children: monthLabel(calMonthValue)})
                        , _jsx('button', {
                          type: "button",
                          className: "cal-nav",
                          'aria-label': "下个月",
                          disabled: calIndex >= monthsWithRecords.length - 1,
                          onClick: () => stepMonth(1),
 children: "›"

                        })
                      ]})
                      , _jsxs('div', { className: "cal-grid", children: [
                        ['一', '二', '三', '四', '五', '六', '日'].map((dow) => (
                          _jsx('span', { className: "cal-dow", children: dow}, dow)
                        ))
                        , calCells.map((day, index) =>
                          day ? (
                            _jsx('button', {

                              type: "button",
                              className: `cal-day${
                                day === rangeLo || day === rangeHi ? ' on' : ''
                              }${day > rangeLo && day < rangeHi ? ' in' : ''}`,
                              disabled: !isSelectable(day),
                              onClick: () => pickDay(day),
 children: 
                              Number(day.slice(8))
                            }, day)
                          ) : (
                            _jsx('span', { className: "cal-pad",}, `pad-${index}` )
                          )
                        )
                      ]})
                    ]})
                  )
                ]})
              )
            ]})
          ) : (
            /* 当前 / 全部：时间由上面的选择决定，这里只显示、不能点。 */
            _jsxs('div', { className: "field off" , 'aria-disabled': "true", children: [
              _jsx('span', { className: "field-label", children: "时间"})
              , _jsx('span', { className: "field-value", children: 
                who === 'current' ? monthLabel(defaultMonth) : '全部时间'
              })
            ]})
          )
        })

        , _jsx('p', { className: "caption", children: 
          count === 0
            ? who === 'pick' && picked.length === 0
              ? '先勾选要导出的项目。'
              : dateKeys.length === 0
                ? who === 'pick'
                  ? '勾选的项目还没有记录，换一个项目。'
                  : '这个项目还没有记录。'
                : '这段时间没有记录，换个时间或项目。'
            : format === 'image'
              ? who === 'current'
                ? `将生成「${activeName}」${label} 的工资单图片（${count} 条记录）。`
                : who === 'all'
                  ? `将生成「${activeName}」全部时间的工资单图片（${count} 条记录）。`
                  : `将生成 ${withData.length} 个项目、${label} 的工资单图片（${count} 条记录）。`
              : who === 'current'
                ? `将导出「${activeName}」${label} 的 ${count} 条记录。`
                : who === 'all'
                  ? `将导出「${activeName}」的全部 ${count} 条记录（${spanLabel}）。`
                  : `将导出 ${withData.length} 个项目、${count} 条记录（${label}）。${
                      withData.length > 1 ? '每个项目一节，最后带合计。' : ''
                    }`
        })

        , _jsxs('div', { className: "datacard", children: [
          _jsxs('div', { className: "datacard-head", children: [
            _jsx('span', { className: "datacard-title", children: "数据备份"})
            , _jsx('span', { className: "datacard-hint", children: "换手机或清数据前先导出一份"})
          ]})
          , _jsxs('div', { className: "datacard-actions", children: [
            _jsx('button', { type: "button", onClick: onExportData, children: "导出数据"})
            , _jsx('button', { type: "button", onClick: onImport, children: "导入数据"})
          ]})
        ]})
      ]})
    })
  );
}

function ProjectMenu({ project, canDelete, onEdit, onDelete, onClose }) {
  const [shown, setShown] = useState(false);

  useEffect(() => {
    const frame = requestAnimationFrame(() => setShown(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  return (
    _jsx('div', { className: `sheet-backdrop${shown ? ' shown' : ''}`, onClick: onClose, children: 
      _jsxs('div', { className: "sheet", onClick: (event) => event.stopPropagation(), children: [
        _jsx('div', { className: "sheet-grabber", 'aria-hidden': "true",} )
        , _jsx('p', { className: "action-title", children: project.name || '未命名项目'})

        , _jsxs('div', { className: "group", children: [
          _jsx('button', { type: "button", className: "action-row", onClick: onEdit, children: "编辑项目"})
          , _jsx('button', {
            type: "button",
            className: "action-row danger" ,
            onClick: onDelete,
            disabled: !canDelete,
 children: "删除项目"

          })
        ]})

        , _jsx('div', { className: "group action-cancel" , children: 
          _jsx('button', { type: "button", className: "action-row", onClick: onClose, children: "取消"})
        })

        , !canDelete && _jsx('p', { className: "caption", children: "至少要留一个项目。"})
      ]})
    })
  );
}

/// 首页「今天」：一眼看到今天记没记，点一下就记。

/// 批量补记：选一段日期，一次填上上班或休息。
function BatchSheet({ projectId, projectName, defaultStart, defaultEnd, records, onSave, onClose }) {
  const [from, setFrom] = useState(defaultStart);
  const [to, setTo] = useState(defaultEnd);
  const [fill, setFill] = useState('work');
  const [skip, setSkip] = useState(true);
  const [shown, setShown] = useState(false);

  useEffect(() => {
    const frame = requestAnimationFrame(() => setShown(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  const days = useMemo(() => {
    if (!from || !to || from > to) return [];
    const list = [];
    let cursor = from;
    for (let index = 0; index < 366 && cursor <= to; index += 1) {
      list.push(cursor);
      cursor = dayShift(cursor, 1);
    }
    return list;
  }, [from, to]);

  const existing = useMemo(
    () =>
      new Set(
        records.filter((record) => record.projectId === projectId).map((record) => record.date)
      ),
    [records, projectId]
  );

  const writable = useMemo(
    () => (skip ? days.filter((day) => !existing.has(day)) : days),
    [days, existing, skip]
  );
  const skippedCount = days.length - (skip ? writable.length : 0);
  const valid = writable.length > 0;

  const submit = (event) => {
    event.preventDefault();
    if (!valid) return;
    onSave(
      writable.map((day) => ({
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}-${day}`,
        projectId,
        date: day,
        status: fill === 'rest' ? 'rest' : 'work',
        days: fill === 'rest' ? 0 : 1,
        note: '',
        paid: 0,
        overtime: null,
      }))
    );
  };

  return (
    _jsx('div', { className: `sheet-backdrop${shown ? ' shown' : ''}`, onClick: onClose, children: 
      _jsxs('form', { className: "sheet", onClick: (event) => event.stopPropagation(), onSubmit: submit, children: [
        _jsx('div', { className: "sheet-grabber", 'aria-hidden': "true",} )
        , _jsxs('header', { className: "sheet-bar", children: [
          _jsx('button', { type: "button", className: "link", onClick: onClose, children: "取消"})
          , _jsx('strong', { children: "补记"})
          , _jsx('button', { type: "submit", className: "link strong" , disabled: !valid, children: "补记"})
        ]})

        , _jsxs('div', { className: "group", children: [
          _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "开始"})
            , _jsx('input', { type: "date", value: from, onChange: (event) => setFrom(event.target.value),} )
          ]})
          , _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "结束"})
            , _jsx('input', { type: "date", value: to, onChange: (event) => setTo(event.target.value),} )
          ]})
          , _jsxs('div', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "默认"})
            , _jsxs('span', { className: "choices", children: [
              _jsx('button', { type: "button", className: fill === 'work' ? 'on' : '', onClick: () => setFill('work'), children: "上班"  })
              , _jsx('button', { type: "button", className: fill === 'rest' ? 'on' : '', onClick: () => setFill('rest'), children: "休息"})
            ]})
          ]})
          , _jsxs('label', { className: "field", children: [
            _jsx('span', { className: "field-label", children: "跳过已有"})
            , _jsx('span', { className: "switch", children: 
              _jsx('input', { type: "checkbox", checked: skip, onChange: (event) => setSkip(event.target.checked),} )
            })
          ]})
        ]})

        , _jsx('p', { className: "caption", children: 
          from && to && from > to
            ? '结束日期不能早于开始日期。'
            : valid
              ? `将补记 ${writable.length} 天：${
                  from === to ? dayLabel(from) : `${dayLabel(from)}—${dayLabel(to)}`
                }。${skippedCount > 0 ? `其中 ${skippedCount} 天已有记录，会跳过。` : ''}`
              : '这段日期没有需要补记的日子。'
        })
      ]})
    })
  );
}


/// 首页「今天」的项目切换。
