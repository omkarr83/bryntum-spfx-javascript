/**
 * Transform Dataverse tasks to Bryntum Gantt format (ported from server/utils/dataTransformer).
 * Read-only: used for loading data into the Gantt.
 */

export interface DataverseTask {
  eppm_projecttaskid?: string;
  eppm_projectid?: string;
  eppm_name?: string;
  eppm_startdate?: string;
  eppm_finishdate?: string;
  eppm_taskduration?: number;
  eppm_taskindex?: number;
  eppm_pocpercentage?: number;
  eppm_taskwork?: number;
  eppm_predecessor?: string;
  eppm_successors?: string;
  eppm_notes?: string;
  eppm_parenttaskid?: string;
  eppm_calendarname?: string;
  eppm_ignoreresourcecalendar?: boolean;
  eppm_rollup?: boolean;
  eppm_inactive?: boolean;
  eppm_manuallyscheduled?: boolean;
  eppm_projectborder?: string;
  eppm_constrainttype?: number;
  eppm_constraintdate?: string;
  [key: string]: unknown;
}

export interface BryntumTask {
  id?: number | string;
  projectId?: string;
  name?: string;
  startDate?: string;
  endDate?: string;
  duration?: number;
  durationUnit?: string;
  percentDone?: number;
  effort?: number;
  effortUnit?: string;
  parentId?: number | string;
  expanded?: boolean;
  children?: BryntumTask[];
  constraintType?: string;
  constraintDate?: string;
  note?: string;
  successors?: string;
  calendar?: string;
  ignoreResourceCalendar?: boolean;
  schedulingMode?: string;
  effortDriven?: boolean;
  rollup?: boolean;
  inactive?: boolean;
  manuallyScheduled?: boolean;
  projectBorder?: string;
  rawStartDate?: string;
  rawFinishDate?: string;
  rawDuration?: number;
  taskIndex?: number;
  [key: string]: unknown;
}

export interface BryntumDependency {
  id?: number | string;
  fromTask?: number | string;
  toTask?: number | string;
  type?: number;
  lag?: number;
}

function formatDateToYYYYMMDD(dateString: string | null | undefined): string | undefined {
  if (!dateString) return undefined;
  try {
    const trimmed = String(dateString).trim();
    const dateMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
    if (dateMatch) return dateMatch[1];
    const date = new Date(trimmed);
    if (isNaN(date.getTime())) return undefined;
    const y = date.getUTCFullYear();
    const m = ('0' + (date.getUTCMonth() + 1)).slice(-2);
    const d = ('0' + date.getUTCDate()).slice(-2);
    return y + '-' + m + '-' + d;
  } catch {
    return undefined;
  }
}

function addDaysToDateString(dateStr: string, days: number): string {
  if (!dateStr || typeof dateStr !== 'string') return '';
  const date = new Date(dateStr + 'T12:00:00Z');
  if (isNaN(date.getTime())) return '';
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().split('T')[0];
}

const CONSTRAINT_TYPE_MAP: Record<number, string> = {
  100000000: 'assoonaspossible',
  100000001: 'aslateaspossible',
  100000002: 'muststarton',
  100000003: 'mustfinishon',
  100000004: 'startnoearlierthan',
  100000005: 'startnolaterthan',
  100000006: 'finishnoearlierthan',
  100000007: 'finishnolaterthan'
};

function getConstraintTypeName(constraintType: number): string {
  return CONSTRAINT_TYPE_MAP[constraintType] || 'assoonaspossible';
}

export function dataverseToBryntumTask(dataverseTask: DataverseTask): BryntumTask {
  const startDateStr = formatDateToYYYYMMDD(dataverseTask.eppm_startdate);
  const finishDateStr = formatDateToYYYYMMDD(dataverseTask.eppm_finishdate);
  const endDateForBryntum = finishDateStr ? addDaysToDateString(finishDateStr, 1) : undefined;

  const bryntumTask: BryntumTask = {
    id: dataverseTask.eppm_projecttaskid || undefined,
    projectId: dataverseTask.eppm_projectid ?? undefined,
    name: dataverseTask.eppm_name || 'Unnamed Task',
    startDate: startDateStr,
    endDate: endDateForBryntum,
    durationUnit: 'day',
    taskIndex: dataverseTask.eppm_taskindex ?? undefined,
    percentDone: dataverseTask.eppm_pocpercentage ?? undefined,
    effort: dataverseTask.eppm_taskwork ?? undefined,
    effortUnit: 'hour',
    note: dataverseTask.eppm_notes ?? undefined,
    successors: dataverseTask.eppm_successors ?? undefined,
    calendar: dataverseTask.eppm_calendarname ?? undefined,
    ignoreResourceCalendar: dataverseTask.eppm_ignoreresourcecalendar ?? undefined,
    schedulingMode: 'Normal',
    effortDriven: false,
    rollup: dataverseTask.eppm_rollup ?? undefined,
    inactive: dataverseTask.eppm_inactive ?? undefined,
    manuallyScheduled: true,
    projectBorder: dataverseTask.eppm_projectborder ?? undefined,
    rawStartDate: startDateStr,
    rawFinishDate: finishDateStr,
    rawDuration: dataverseTask.eppm_taskduration ?? undefined
  };

  if (dataverseTask.eppm_constrainttype !== undefined && dataverseTask.eppm_constrainttype !== null) {
    bryntumTask.constraintType = getConstraintTypeName(dataverseTask.eppm_constrainttype);
  }
  if (dataverseTask.eppm_constraintdate) {
    bryntumTask.constraintDate = formatDateToYYYYMMDD(dataverseTask.eppm_constraintdate);
  }
  if (dataverseTask.eppm_parenttaskid) {
    bryntumTask.parentId = dataverseTask.eppm_parenttaskid;
  }

  return bryntumTask;
}

export function buildTaskHierarchy(tasks: DataverseTask[]): BryntumTask[] {
  const taskMap = new Map<string, BryntumTask>();
  const rootTasks: BryntumTask[] = [];

  if (!tasks || tasks.length === 0) return rootTasks;

  tasks.forEach(task => {
    if (!task.eppm_projecttaskid) return;
    const bryntumTask = dataverseToBryntumTask(task);
    if (bryntumTask.id) taskMap.set(String(bryntumTask.id), bryntumTask);
  });

  tasks.forEach(task => {
    if (!task.eppm_projecttaskid) return;
    const bryntumTask = taskMap.get(String(task.eppm_projecttaskid));
    if (!bryntumTask) return;

    if (task.eppm_parenttaskid) {
      const parent = taskMap.get(String(task.eppm_parenttaskid));
      if (parent) {
        if (!parent.children) parent.children = [];
        parent.children.push(bryntumTask);
      } else {
        rootTasks.push(bryntumTask);
      }
    } else {
      rootTasks.push(bryntumTask);
    }
  });

  const sortTasks = (taskList: BryntumTask[]): BryntumTask[] => {
    return taskList.sort((a, b) => {
      const indexA = (a as BryntumTask).taskIndex;
      const indexB = (b as BryntumTask).taskIndex;
      if (indexA !== undefined && indexB !== undefined) return Number(indexA) - Number(indexB);
      if (indexA !== undefined) return -1;
      if (indexB !== undefined) return 1;
      return (a.startDate || '').localeCompare(b.startDate || '');
    }).map(task => {
      if (task.children && task.children.length > 0) task.children = sortTasks(task.children);
      return task;
    });
  };

  return sortTasks(rootTasks);
}

// Dependency parsing (from server tasks.routes.ts)
const DEPENDENCY_TYPE_MAP: Record<string, number> = { SS: 0, SF: 1, FS: 2, FF: 3 };

function abbrToType(abbr: string): number {
  return DEPENDENCY_TYPE_MAP[abbr.toUpperCase()] ?? 2;
}

function normalizePredecessorString(value: unknown): string {
  if (typeof value !== 'string') return '';
  return value.split(/[;,]+/g).map(s => s.trim()).filter(Boolean).join(';');
}

function normalizeSuccessorString(value: unknown): string {
  if (typeof value !== 'string') return '';
  return value.split(/[;,]+/g).map(s => s.trim()).filter(Boolean).join(';');
}

function parsePredecessorString(str: string, toTaskId: string): BryntumDependency[] {
  const cleaned = normalizePredecessorString(str);
  if (!cleaned) return [];
  const parts = cleaned.split(';').map(s => s.trim()).filter(Boolean);
  const deps: BryntumDependency[] = [];
  parts.forEach((part, idx) => {
    const m = part.match(/^(.*?)(FS|SS|FF|SF)([+-]\d+(?:\.\d+)?[a-zA-Z]?)?$/i);
    if (!m) return;
    const fromTask = m[1]?.trim();
    const abbr = (m[2] || 'FS').toUpperCase();
    const lagPart = m[3];
    if (!fromTask) return;
    let lag: number | undefined;
    if (lagPart) {
      const n = parseFloat(lagPart);
      if (!isNaN(n) && n !== 0) lag = n;
    }
    deps.push({
      id: `${toTaskId}_${fromTask}_${abbr}_${idx}`,
      fromTask,
      toTask: toTaskId,
      type: abbrToType(abbr),
      ...(lag !== undefined ? { lag } : {})
    });
  });
  return deps;
}

function parseSuccessorString(str: string, fromTaskId: string): BryntumDependency[] {
  const cleaned = normalizeSuccessorString(str);
  if (!cleaned) return [];
  const parts = cleaned.split(';').map(s => s.trim()).filter(Boolean);
  const deps: BryntumDependency[] = [];
  parts.forEach((part, idx) => {
    const m = part.match(/^(.*?)(FS|SS|FF|SF)([+-]\d+(?:\.\d+)?[a-zA-Z]?)?$/i);
    if (!m) return;
    const toTask = m[1]?.trim();
    const abbr = (m[2] || 'FS').toUpperCase();
    const lagPart = m[3];
    if (!toTask) return;
    let lag: number | undefined;
    if (lagPart) {
      const n = parseFloat(lagPart);
      if (!isNaN(n) && n !== 0) lag = n;
    }
    deps.push({
      id: `${fromTaskId}_${toTask}_${abbr}_succ_${idx}`,
      fromTask: fromTaskId,
      toTask,
      type: abbrToType(abbr),
      ...(lag !== undefined ? { lag } : {})
    });
  });
  return deps;
}

export function buildDependencies(tasks: DataverseTask[]): BryntumDependency[] {
  const dependencyRows: BryntumDependency[] = [];
  const depKeySet = new Set<string>();

  const pushUnique = (d: BryntumDependency): void => {
    const key = `${d.fromTask ?? ''}->${d.toTask ?? ''}:${d.type ?? ''}:${d.lag ?? ''}`;
    if (!depKeySet.has(key)) {
      depKeySet.add(key);
      dependencyRows.push(d);
    }
  };

  for (const t of tasks) {
    const taskId = t.eppm_projecttaskid;
    if (!taskId) continue;
    const pred = t.eppm_predecessor;
    if (typeof pred === 'string' && pred.trim()) {
      parsePredecessorString(pred, String(taskId)).forEach(pushUnique);
    }
    const succ = t.eppm_successors;
    if (typeof succ === 'string' && succ.trim()) {
      parseSuccessorString(succ, String(taskId)).forEach(pushUnique);
    }
  }

  return dependencyRows.filter(d => d && d.fromTask && d.toTask);
}
