/**
 * MSPDI import/export in the browser (SPFx only, no Node server).
 * Uses DOMParser for XML and fetch to Dataverse for create/update.
 */

// ---------- Import types (same shape as server mspdiImporter) ----------
export interface ImportedTask {
  uid: number;
  id: number;
  name: string;
  startDate?: string;
  finishDate?: string;
  duration?: number;
  percentComplete?: number;
  effort?: number;
  outlineLevel: number;
  parentUid?: number;
  notes?: string;
  wbs?: string;
  isSummary?: boolean;
  isMilestone?: boolean;
  dataverseTaskId?: string;
}

export interface ImportedResource {
  uid: number;
  id: number;
  name: string;
  email?: string;
}

export interface ImportedAssignment {
  uid: number;
  taskUid: number;
  resourceUid: number;
  units?: number;
  startDate?: string;
  finishDate?: string;
  dataverseAssignmentId?: string;
}

export interface ImportedDependency {
  fromTaskUid: number;
  toTaskUid: number;
  type: number;
  lag?: number;
}

export interface ImportedProjectData {
  projectName?: string;
  startDate?: string;
  tasks: ImportedTask[];
  resources: ImportedResource[];
  assignments: ImportedAssignment[];
  dependencies: ImportedDependency[];
}

export interface BryntumImportData {
  tasks: Array<Record<string, unknown>>;
  resources: Array<Record<string, unknown>>;
  assignments: Array<Record<string, unknown>>;
  dependencies: Array<Record<string, unknown>>;
}

// ---------- Helpers ----------
function parseMspdiDate(dateStr: string | undefined): string | undefined {
  if (!dateStr) return undefined;
  const trimmed = dateStr.trim();
  const m = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  const d = new Date(trimmed);
  if (isNaN(d.getTime())) return undefined;
  const y = d.getUTCFullYear();
  const mo = (d.getUTCMonth() + 1) < 10 ? '0' + (d.getUTCMonth() + 1) : String(d.getUTCMonth() + 1);
  const day = d.getUTCDate() < 10 ? '0' + d.getUTCDate() : String(d.getUTCDate());
  return y + '-' + mo + '-' + day;
}

function parseMspdiDuration(duration: string | undefined): number | undefined {
  if (!duration) return undefined;
  const m = duration.match(/PT(\d+)H(\d+)M(\d+)S/);
  if (!m) return undefined;
  const hours = parseInt(m[1], 10) || 0;
  const minutes = parseInt(m[2], 10) || 0;
  return (hours + minutes / 60) / 8;
}

function parseMspdiWork(work: string | undefined): number | undefined {
  if (!work) return undefined;
  const m = work.match(/PT(\d+)H(\d+)M(\d+)S/);
  if (!m) return undefined;
  const hours = parseInt(m[1], 10) || 0;
  const minutes = parseInt(m[2], 10) || 0;
  return hours + minutes / 60;
}

function toSafeISOString(value: string | Date | number | null | undefined): string | undefined {
  if (value == null || value === '') return undefined;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    const dateOnly = trimmed.match(/^(\d{4}-\d{2}-\d{2})$/);
    if (dateOnly) return `${dateOnly[1]}T12:00:00Z`;
    const dateTime = trimmed.match(/^(\d{4}-\d{2}-\d{2})T/);
    if (dateTime) return `${dateTime[1]}T12:00:00Z`;
  }
  const date = value instanceof Date ? value : new Date(value as string | number);
  if (isNaN(date.getTime())) return undefined;
  const y = date.getUTCFullYear();
  const mo = (date.getUTCMonth() + 1) < 10 ? '0' + (date.getUTCMonth() + 1) : String(date.getUTCMonth() + 1);
  const day = date.getUTCDate() < 10 ? '0' + date.getUTCDate() : String(date.getUTCDate());
  return y + '-' + mo + '-' + day + 'T12:00:00Z';
}

// ---------- DOM parser helpers ----------
function elText(el: Element | null): string | undefined {
  if (!el) return undefined;
  const t = el.textContent;
  return t != null ? t.trim() || undefined : undefined;
}

function elInt(el: Element | null): number | undefined {
  const t = elText(el);
  if (t === undefined) return undefined;
  const n = parseInt(t, 10);
  return isNaN(n) ? undefined : n;
}

function elNum(el: Element | null): number | undefined {
  const t = elText(el);
  if (t === undefined) return undefined;
  const n = parseFloat(t);
  return isNaN(n) ? undefined : n;
}

const FIELD_ID_DATAVERSE_TASK_ID = 188743731;
const FIELD_ID_DATAVERSE_ASSIGNMENT_ID = 188743734;

function getExtendedAttributeValue(parent: Element, fieldId: number): string | undefined {
  const extAttrs = parent.getElementsByTagName('ExtendedAttribute');
  for (let i = 0; i < extAttrs.length; i++) {
    const fid = elInt(extAttrs[i].getElementsByTagName('FieldID')[0]);
    if (fid === fieldId) {
      return elText(extAttrs[i].getElementsByTagName('Value')[0]);
    }
  }
  return undefined;
}

/**
 * Parse MSPDI XML in the browser using DOMParser.
 */
export function parseMspdiXmlBrowser(xmlContent: string): ImportedProjectData {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlContent, 'text/xml');
  const project = doc.documentElement;
  if (!project || project.tagName !== 'Project') {
    throw new Error('Invalid MSPDI XML: Missing or wrong root Project element');
  }

  const projectName = elText(project.getElementsByTagName('Name')[0]) || elText(project.getElementsByTagName('Title')[0]) || 'Imported Project';
  const startDate = parseMspdiDate(elText(project.getElementsByTagName('StartDate')[0]));

  const tasks: ImportedTask[] = [];
  const tasksRoot = project.getElementsByTagName('Tasks')[0];
  if (tasksRoot) {
    const taskEls = tasksRoot.getElementsByTagName('Task');
    for (let i = 0; i < taskEls.length; i++) {
      const t = taskEls[i];
      const uid = elInt(t.getElementsByTagName('UID')[0]);
      const id = elInt(t.getElementsByTagName('ID')[0]);
      if (uid === 0 || uid === undefined || id === undefined) continue;
      const outlineLevel = elInt(t.getElementsByTagName('OutlineLevel')[0]) || 1;
      const isSummary = elText(t.getElementsByTagName('Summary')[0]) === '1';
      const isMilestone = elText(t.getElementsByTagName('Milestone')[0]) === '1';
      const dataverseTaskId = getExtendedAttributeValue(t, FIELD_ID_DATAVERSE_TASK_ID);
      tasks.push({
        uid,
        id,
        name: elText(t.getElementsByTagName('Name')[0]) || 'Unnamed Task',
        startDate: parseMspdiDate(elText(t.getElementsByTagName('Start')[0])),
        finishDate: parseMspdiDate(elText(t.getElementsByTagName('Finish')[0])),
        duration: parseMspdiDuration(elText(t.getElementsByTagName('Duration')[0])),
        percentComplete: elNum(t.getElementsByTagName('PercentComplete')[0]),
        effort: parseMspdiWork(elText(t.getElementsByTagName('Work')[0])),
        outlineLevel,
        notes: elText(t.getElementsByTagName('Notes')[0]),
        wbs: elText(t.getElementsByTagName('WBS')[0]),
        isSummary,
        isMilestone,
        dataverseTaskId
      });
    }
  }

  // Parent from outline level
  const taskStack: ImportedTask[] = [];
  for (const task of tasks) {
    while (taskStack.length > 0 && taskStack[taskStack.length - 1].outlineLevel >= task.outlineLevel) taskStack.pop();
    if (taskStack.length > 0) task.parentUid = taskStack[taskStack.length - 1].uid;
    taskStack.push(task);
  }

  const dependencies: ImportedDependency[] = [];
  if (tasksRoot) {
    const taskEls = tasksRoot.getElementsByTagName('Task');
    for (let i = 0; i < taskEls.length; i++) {
      const t = taskEls[i];
      const toTaskUid = elInt(t.getElementsByTagName('UID')[0]);
      if (toTaskUid === undefined || toTaskUid === 0) continue;
      const predLinks = t.getElementsByTagName('PredecessorLink');
      for (let j = 0; j < predLinks.length; j++) {
        const link = predLinks[j];
        const fromTaskUid = elInt(link.getElementsByTagName('PredecessorUID')[0]);
        if (fromTaskUid === undefined || fromTaskUid === 0) continue;
        const type = elInt(link.getElementsByTagName('Type')[0]) ?? 1;
        const linkLag = elInt(link.getElementsByTagName('LinkLag')[0]) || 0;
        const lagDays = linkLag / 4800;
        dependencies.push({
          fromTaskUid,
          toTaskUid,
          type,
          lag: lagDays !== 0 ? lagDays : undefined
        });
      }
    }
  }

  const resources: ImportedResource[] = [];
  const resourcesRoot = project.getElementsByTagName('Resources')[0];
  if (resourcesRoot) {
    const resEls = resourcesRoot.getElementsByTagName('Resource');
    for (let i = 0; i < resEls.length; i++) {
      const r = resEls[i];
      const uid = elInt(r.getElementsByTagName('UID')[0]);
      const id = elInt(r.getElementsByTagName('ID')[0]);
      if (uid === undefined || id === undefined || uid === 0) continue;
      const resourceType = elInt(r.getElementsByTagName('Type')[0]);
      if (resourceType !== undefined && resourceType !== 1) continue;
      resources.push({
        uid,
        id,
        name: elText(r.getElementsByTagName('Name')[0]) || 'Unknown Resource',
        email: elText(r.getElementsByTagName('EmailAddress')[0])
      });
    }
  }

  const assignments: ImportedAssignment[] = [];
  const assignRoot = project.getElementsByTagName('Assignments')[0];
  if (assignRoot) {
    const assignEls = assignRoot.getElementsByTagName('Assignment');
    for (let i = 0; i < assignEls.length; i++) {
      const a = assignEls[i];
      const uid = elInt(a.getElementsByTagName('UID')[0]);
      const taskUid = elInt(a.getElementsByTagName('TaskUID')[0]);
      const resourceUid = elInt(a.getElementsByTagName('ResourceUID')[0]);
      if (uid === undefined || taskUid === undefined || resourceUid === undefined || taskUid === 0 || resourceUid === 0) continue;
      const unitsRaw = elNum(a.getElementsByTagName('Units')[0]);
      let units = 100;
      if (unitsRaw !== undefined && isFinite(unitsRaw)) {
        units = Math.round(unitsRaw * 100);
        if (units < 0) units = 0;
        if (units > 1000) units = 1000;
      }
      const dataverseAssignmentId = getExtendedAttributeValue(a, FIELD_ID_DATAVERSE_ASSIGNMENT_ID);
      assignments.push({
        uid,
        taskUid,
        resourceUid,
        units,
        startDate: elText(a.getElementsByTagName('Start')[0]),
        finishDate: elText(a.getElementsByTagName('Finish')[0]),
        dataverseAssignmentId
      });
    }
  }

  return { projectName, startDate, tasks, resources, assignments, dependencies };
}

/**
 * Convert imported data to Bryntum format (same as server).
 */
export function convertImportedDataToBryntum(data: ImportedProjectData): BryntumImportData {
  const uidToTask = new Map<number, ImportedTask>();
  data.tasks.forEach(task => uidToTask.set(task.uid, task));
  const uidToResource = new Map<number, ImportedResource>();
  data.resources.forEach(resource => uidToResource.set(resource.uid, resource));

  const tasks = data.tasks.map(task => ({
    id: `import_${task.uid}`,
    name: task.name,
    startDate: task.startDate,
    endDate: task.finishDate,
    duration: task.duration,
    percentDone: task.percentComplete || 0,
    effort: task.effort,
    note: task.notes,
    parentId: task.parentUid ? `import_${task.parentUid}` : undefined,
    manuallyScheduled: true,
    effortDriven: false,
    taskIndex: task.id,
    _importUid: task.uid,
    _outlineLevel: task.outlineLevel,
    _dataverseTaskId: task.dataverseTaskId,
    _importId: `import_${task.uid}`,
    _parentImportId: task.parentUid ? `import_${task.parentUid}` : undefined
  }));

  const resources = data.resources.map(resource => {
    const identifier = resource.email || resource.name || `resource_${resource.uid}`;
    return {
      id: identifier.toLowerCase(),
      name: resource.name,
      email: identifier.toLowerCase(),
      _importUid: resource.uid
    };
  });

  const resourceUidToId = new Map<number, string>();
  data.resources.forEach(resource => {
    const identifier = resource.email || resource.name || `resource_${resource.uid}`;
    resourceUidToId.set(resource.uid, identifier.toLowerCase());
  });

  const assignments = data.assignments
    .filter(function (a) { return uidToTask.has(a.taskUid) && uidToResource.has(a.resourceUid); })
    .map(function (assignment) {
      const u = assignment.units;
      const unitsVal = (typeof u === 'number' && (u === u) && u !== Infinity && u !== -Infinity) ? u : 100;
      return {
      id: 'assignment_' + assignment.uid,
      event: 'import_' + assignment.taskUid,
      resource: resourceUidToId.get(assignment.resourceUid) || '',
      units: unitsVal,
      startDate: assignment.startDate,
      finishDate: assignment.finishDate,
      _dataverseAssignmentId: assignment.dataverseAssignmentId,
      taskImportId: 'import_' + assignment.taskUid,
      resourceEmail: resourceUidToId.get(assignment.resourceUid) || ''
    };
    });

  const typeMap: Record<number, number> = { 0: 3, 1: 2, 2: 1, 3: 0 };
  const dependencies = data.dependencies
    .filter(dep => uidToTask.has(dep.fromTaskUid) && uidToTask.has(dep.toTaskUid))
    .map((dep, index) => ({
      id: `dep_${index}`,
      fromTask: `import_${dep.fromTaskUid}`,
      toTask: `import_${dep.toTaskUid}`,
      type: typeMap[dep.type] ?? 2,
      lag: dep.lag
    }));

  return { tasks, resources, assignments, dependencies };
}

/**
 * Convert to Dataverse payload shape (for create/update).
 */
export function convertImportedDataToDataverse(
  data: BryntumImportData,
  projectId?: string
): { tasks: Array<Record<string, unknown>>; resources: Array<Record<string, unknown>>; assignments: Array<Record<string, unknown>> } {
  const tasks = data.tasks.map(function (task) {
    const row: Record<string, unknown> = {
      eppm_name: task.name,
      eppm_startdate: toSafeISOString(task.startDate as string | undefined),
      eppm_finishdate: toSafeISOString((task.endDate || task.finishDate) as string | undefined),
      eppm_taskduration: task.duration,
      eppm_pocpercentage: task.percentDone,
      eppm_taskwork: task.effort,
      eppm_notes: task.note,
      eppm_taskindex: task.taskIndex || task._importUid,
      _importId: task.id,
      _parentImportId: task.parentId,
      _dataverseTaskId: task._dataverseTaskId
    };
    if (projectId) row.eppm_projectid = projectId;
    return row;
  });

  const resources = data.resources.map(resource => ({
    email: resource.email || resource.id,
    name: resource.name,
    _importUid: resource._importUid
  }));

  const assignments = data.assignments.map(function (assignment) {
    const u = assignment.units;
    const unitsVal = (typeof u === 'number' && (u === u) && u !== Infinity && u !== -Infinity) ? u : 100;
    return {
    taskImportId: assignment.event,
    resourceEmail: assignment.resource,
    units: unitsVal,
    startDate: assignment.startDate,
    finishDate: assignment.finishDate,
    _dataverseAssignmentId: assignment._dataverseAssignmentId
  };
  });

  return { tasks, resources, assignments };
}

const TYPE_TO_STR: Record<number, string> = { 0: 'SS', 1: 'SF', 2: 'FS', 3: 'FF' };

export function buildPredecessorStringForTask(
  taskImportId: string,
  dependencies: Array<Record<string, unknown>>,
  importIdToDataverseId: Map<string, string>
): string {
  const predecessors = dependencies.filter(d => d.toTask === taskImportId);
  if (predecessors.length === 0) return '';
  return predecessors.map(function (dep) {
    const fromId = importIdToDataverseId.get(dep.fromTask as string) || dep.fromTask;
    const typeNum = dep.type as number;
    const typeStr = TYPE_TO_STR[typeNum] || 'FS';
    let token = String(fromId) + typeStr;
    const lagVal = dep.lag as number | undefined | null;
    if (lagVal != null && lagVal !== 0) token += (lagVal > 0 ? '+' : '') + lagVal + 'd';
    return token;
  }).join(';');
}

export function buildSuccessorStringForTask(
  taskImportId: string,
  dependencies: Array<Record<string, unknown>>,
  importIdToDataverseId: Map<string, string>
): string {
  const successors = dependencies.filter(d => d.fromTask === taskImportId);
  if (successors.length === 0) return '';
  return successors.map(function (dep) {
    const toId = importIdToDataverseId.get(dep.toTask as string) || dep.toTask;
    const typeNum = dep.type as number;
    const typeStr = TYPE_TO_STR[typeNum] || 'FS';
    let token = String(toId) + typeStr;
    const lagVal = dep.lag as number | undefined | null;
    if (lagVal != null && lagVal !== 0) token += (lagVal > 0 ? '+' : '') + lagVal + 'd';
    return token;
  }).join(';');
}

// ---------- Export to MSPDI XML (browser-only) ----------
export interface MspdiTaskExport {
  id: string;
  uid: number;
  name: string;
  startDate?: string;
  finishDate?: string;
  duration?: number;
  percentComplete?: number;
  effort?: number;
  parentId?: string;
  outlineLevel?: number;
  predecessors?: Array<{ predecessorUid: number; type: number; lag?: number }>;
  notes?: string;
  dataverseTaskId?: string;
}

export interface MspdiResourceExport {
  id: string;
  uid: number;
  name: string;
  email?: string;
}

export interface MspdiAssignmentExport {
  taskUid: number;
  resourceUid: number;
  units?: number;
  dataverseAssignmentId?: string;
}

export interface MspdiProjectDataExport {
  projectName?: string;
  startDate?: string;
  tasks: MspdiTaskExport[];
  resources: MspdiResourceExport[];
  assignments: MspdiAssignmentExport[];
}

function escapeXml(str: string | null | undefined): string {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function formatMspdiDateExport(dateStr: string | null | undefined, defaultTime: 'start' | 'finish' = 'start'): string {
  if (!dateStr) return '';
  const m = String(dateStr).match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1] + (defaultTime === 'finish' ? 'T17:00:00' : 'T08:00:00');
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return '';
  const y = d.getUTCFullYear();
  const mo = (d.getUTCMonth() + 1) < 10 ? '0' + (d.getUTCMonth() + 1) : String(d.getUTCMonth() + 1);
  const day = d.getUTCDate() < 10 ? '0' + d.getUTCDate() : String(d.getUTCDate());
  return y + '-' + mo + '-' + day + (defaultTime === 'finish' ? 'T17:00:00' : 'T08:00:00');
}

function formatMspdiDurationExport(days: number | null | undefined): string {
  if (days == null || isNaN(days)) return 'PT0H0M0S';
  return 'PT' + Math.round(days * 8) + 'H0M0S';
}

function formatMspdiWorkExport(hours: number | null | undefined): string {
  if (hours == null || isNaN(hours)) return 'PT0H0M0S';
  return 'PT' + Math.round(hours) + 'H0M0S';
}

/** Build MspdiProjectData from flat Dataverse tasks + resources + assignments + dependencies (from buildDependencies). */
export function convertToMspdiFormatFromDataverse(
  tasks: Array<Record<string, unknown>>,
  resources: Array<{ id: string; name: string; email?: string }>,
  assignments: Array<{ event: string; resource: string; units?: number; id?: string }>,
  dependencies: Array<{ fromTask?: string; toTask?: string; type?: number; lag?: number }>,
  projectName?: string
): MspdiProjectDataExport {
  const taskIdToUid = new Map<string, number>();
  const flatTasks: MspdiTaskExport[] = [];
  tasks.forEach((task, index) => {
    const id = String(task.eppm_projecttaskid || task.id || '');
    if (!id) return;
    const uid = index + 1;
    taskIdToUid.set(id, uid);
    const startDate = (task.rawStartDate || task.eppm_startdate || task.startDate) as string | undefined;
    const finishDate = (task.rawFinishDate || task.eppm_finishdate || task.finishDate) as string | undefined;
    flatTasks.push({
      id,
      uid,
      name: (task.eppm_name || task.name || 'Unnamed Task') as string,
      startDate: startDate ? String(startDate).split('T')[0] : undefined,
      finishDate: finishDate ? String(finishDate).split('T')[0] : undefined,
      duration: (task.eppm_taskduration ?? task.rawDuration ?? task.duration) as number | undefined,
      percentComplete: (task.eppm_pocpercentage ?? task.percentDone) as number | undefined,
      effort: (task.eppm_taskwork ?? task.effort) as number | undefined,
      parentId: task.eppm_parenttaskid as string | undefined,
      outlineLevel: 1,
      predecessors: [],
      notes: (task.eppm_notes || task.note) as string | undefined,
      dataverseTaskId: id
    });
  });

  const bryntumToMspdiType: Record<number, number> = { 0: 3, 1: 2, 2: 1, 3: 0 };
  dependencies.forEach(dep => {
    const toTaskId = String(dep.toTask || '');
    const fromTaskId = String(dep.fromTask || '');
    if (!toTaskId || !fromTaskId) return;
    const toUid = taskIdToUid.get(toTaskId);
    const fromUid = taskIdToUid.get(fromTaskId);
    if (toUid == null || fromUid == null) return;
    let target: MspdiTaskExport | undefined;
    for (let ti = 0; ti < flatTasks.length; ti++) {
      if (flatTasks[ti].id === toTaskId) { target = flatTasks[ti]; break; }
    }
    if (target && target.predecessors) {
      target.predecessors.push({
        predecessorUid: fromUid,
        type: bryntumToMspdiType[dep.type ?? 2] ?? 1,
        lag: dep.lag
      });
    }
  });

  const mspdiResources: MspdiResourceExport[] = resources.map((r, i) => ({
    id: r.id || r.email || '',
    uid: i + 1,
    name: r.name || r.email || 'Unknown',
    email: r.email
  }));
  const resourceIdToUid = new Map<string, number>();
  mspdiResources.forEach(res => resourceIdToUid.set(res.id, res.uid));

  const mspdiAssignments: MspdiAssignmentExport[] = [];
  assignments.forEach(a => {
    const taskUid = taskIdToUid.get(String(a.event));
    const resourceUid = resourceIdToUid.get(String(a.resource));
    if (taskUid != null && resourceUid != null) {
      mspdiAssignments.push({
        taskUid,
        resourceUid,
        units: a.units ?? 100
      });
    }
  });

  let projectStartDate: string | undefined;
  if (flatTasks.length > 0) {
    const dates = flatTasks.map(t => t.startDate).filter(Boolean) as string[];
    if (dates.length > 0) {
      const parsed = dates.map(d => new Date(d).getTime()).filter(n => !isNaN(n));
      if (parsed.length > 0) projectStartDate = new Date(Math.min(...parsed)).toISOString().split('T')[0];
    }
  }
  if (!projectStartDate) projectStartDate = new Date().toISOString().split('T')[0];

  return {
    projectName: projectName || 'Exported Project',
    startDate: projectStartDate,
    tasks: flatTasks,
    resources: mspdiResources,
    assignments: mspdiAssignments
  };
}

function calculateOutlineLevelsExport(tasks: MspdiTaskExport[]): Map<string, number> {
  const levelMap = new Map<string, number>();
  const childToParent = new Map<string, string>();
  tasks.forEach(t => { if (t.parentId) childToParent.set(t.id, t.parentId); });
  const getLevel = (taskId: string): number => {
    if (levelMap.has(taskId)) return levelMap.get(taskId)!;
    const parentId = childToParent.get(taskId);
    if (!parentId) { levelMap.set(taskId, 1); return 1; }
    const level = getLevel(parentId) + 1;
    levelMap.set(taskId, level);
    return level;
  };
  tasks.forEach(t => getLevel(t.id));
  return levelMap;
}

/** Generate MSPDI XML string for download. */
export function generateMspdiXmlBrowser(data: MspdiProjectDataExport): string {
  const { projectName = 'Exported Project', startDate, tasks, resources, assignments } = data;
  const outlineLevels = calculateOutlineLevelsExport(tasks);
  tasks.forEach((t, i) => {
    t.uid = i + 1;
    t.outlineLevel = outlineLevels.get(t.id) || 1;
  });
  resources.forEach((r, i) => { r.uid = i + 1; });

  const taskWorkHoursMap = new Map<number, number>();
  assignments.forEach(function (a) {
    let task: MspdiTaskExport | undefined;
    for (let ti = 0; ti < tasks.length; ti++) {
      if (tasks[ti].uid === a.taskUid) { task = tasks[ti]; break; }
    }
    if (!task) return;
    const units = (a.units ?? 100) / 100;
    const hours = (task.effort && task.effort > 0) ? task.effort * units : (task.duration ?? 0) * 8 * units;
    taskWorkHoursMap.set(task.uid, (taskWorkHoursMap.get(task.uid) || 0) + hours);
  });

  let projectFinishDate = startDate || '';
  if (tasks.length > 0) {
    const finishDates = tasks.map(t => t.finishDate).filter(Boolean) as string[];
    if (finishDates.length > 0)
      projectFinishDate = new Date(Math.max(...finishDates.map(d => new Date(d).getTime()))).toISOString().split('T')[0];
  }
  if (!projectFinishDate) projectFinishDate = startDate || new Date().toISOString().split('T')[0];

  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Project xmlns="http://schemas.microsoft.com/project">
    <SaveVersion>14</SaveVersion>
    <Name>${escapeXml(projectName)}</Name>
    <Title>${escapeXml(projectName)}</Title>
    <ScheduleFromStart>1</ScheduleFromStart>
    <StartDate>${formatMspdiDateExport(startDate, 'start')}</StartDate>
    <FinishDate>${formatMspdiDateExport(projectFinishDate, 'finish')}</FinishDate>
    <FYStartDate>1</FYStartDate>
    <DefaultStartTime>08:00:00</DefaultStartTime>
    <DefaultFinishTime>17:00:00</DefaultFinishTime>
    <MinutesPerDay>480</MinutesPerDay>
    <MinutesPerWeek>2400</MinutesPerWeek>
    <ExtendedAttributes>
        <ExtendedAttribute><FieldID>188743731</FieldID><FieldName>Text1</FieldName><Alias>DataverseTaskID</Alias></ExtendedAttribute>
        <ExtendedAttribute><FieldID>188743734</FieldID><FieldName>Text2</FieldName><Alias>DataverseAssignmentID</Alias></ExtendedAttribute>
    </ExtendedAttributes>
    <Tasks>
        <Task>
            <UID>0</UID>
            <ID>0</ID>
            <Name>${escapeXml(projectName)}</Name>
            <Type>1</Type>
            <OutlineLevel>0</OutlineLevel>
            <Start>${formatMspdiDateExport(startDate, 'start')}</Start>
            <Finish>${formatMspdiDateExport(projectFinishDate, 'finish')}</Finish>
            <Duration>PT0H0M0S</Duration>
            <Summary>1</Summary>
        </Task>
`;

  tasks.forEach(function (task, index) {
    const taskId = index + 1;
    const outlineLevel = task.outlineLevel ?? 1;
    let hasChildren = false;
    for (let si = 0; si < tasks.length; si++) {
      if (tasks[si].parentId === task.id) { hasChildren = true; break; }
    }
    const isSummary = hasChildren ? 1 : 0;
    const isMilestone = (task.duration === 0 || !task.duration) && !hasChildren ? 1 : 0;
    const taskWorkHours = (task.effort && task.effort > 0) ? task.effort : (taskWorkHoursMap.get(task.uid) || 0);
    const taskWork = formatMspdiWorkExport(taskWorkHours);

    xml += `        <Task>
            <UID>${task.uid}</UID>
            <ID>${taskId}</ID>
            <Name>${escapeXml(task.name)}</Name>
            <Type>1</Type>
            <OutlineLevel>${outlineLevel}</OutlineLevel>
            <Start>${formatMspdiDateExport(task.startDate, 'start')}</Start>
            <Finish>${formatMspdiDateExport(task.finishDate, 'finish')}</Finish>
            <Duration>${formatMspdiDurationExport(task.duration)}</Duration>
            <Work>${taskWork}</Work>
            <PercentComplete>${task.percentComplete ?? 0}</PercentComplete>
            <Milestone>${isMilestone}</Milestone>
            <Summary>${isSummary}</Summary>
`;
    if (task.notes) xml += `            <Notes>${escapeXml(task.notes)}</Notes>
`;
    if (task.dataverseTaskId || task.id) {
      xml += `            <ExtendedAttribute><FieldID>188743731</FieldID><Value>${escapeXml(task.dataverseTaskId || task.id)}</Value></ExtendedAttribute>
`;
    }
    if (task.predecessors && task.predecessors.length > 0) {
      task.predecessors.forEach(pred => {
        xml += `            <PredecessorLink><PredecessorUID>${pred.predecessorUid}</PredecessorUID><Type>${pred.type}</Type><LinkLag>${(pred.lag || 0) * 4800}</LinkLag></PredecessorLink>
`;
      });
    }
    xml += `        </Task>
`;
  });

  xml += `    </Tasks>
    <Resources>
`;
  resources.forEach((res, i) => {
    xml += `        <Resource>
            <UID>${res.uid}</UID>
            <ID>${i + 1}</ID>
            <Name>${escapeXml(res.name)}</Name>
            <Type>1</Type>
            <MaxUnits>1</MaxUnits>
`;
    if (res.email) xml += `            <EmailAddress>${escapeXml(res.email)}</EmailAddress>
`;
    xml += `        </Resource>
`;
  });

  xml += `    </Resources>
    <Assignments>
`;
  const taskByUid = new Map<number, MspdiTaskExport>();
  tasks.forEach(t => taskByUid.set(t.uid, t));
  let assignUid = 1;
  assignments.forEach(a => {
    const task = taskByUid.get(a.taskUid);
    const workHours = task ? ((task.effort && task.effort > 0) ? task.effort * (a.units ?? 100) / 100 : (task.duration ?? 0) * 8 * (a.units ?? 100) / 100) : 0;
    const workStr = formatMspdiWorkExport(workHours);
    const startStr = task ? formatMspdiDateExport(task.startDate, 'start') : formatMspdiDateExport(new Date().toISOString());
    const finishStr = task ? formatMspdiDateExport(task.finishDate, 'finish') : formatMspdiDateExport(new Date().toISOString());
    xml += `        <Assignment>
            <UID>${assignUid++}</UID>
            <TaskUID>${a.taskUid}</TaskUID>
            <ResourceUID>${a.resourceUid}</ResourceUID>
            <Units>${(a.units ?? 100) / 100}</Units>
            <Start>${startStr}</Start>
            <Finish>${finishStr}</Finish>
            <Work>${workStr}</Work>
        </Assignment>
`;
  });

  xml += `    </Assignments>
</Project>`;
  return xml;
}
