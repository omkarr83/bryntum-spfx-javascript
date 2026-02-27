/**
 * Fetch project tasks directly from Dataverse OData API (no Node server).
 */
import { dataverseConfig } from './dataverseConfig';
import { DataverseTask } from './dataTransformer';

const SELECT_FIELDS = [
  'eppm_projecttaskid',
  'eppm_projectid',
  'eppm_name',
  'eppm_startdate',
  'eppm_finishdate',
  'eppm_taskduration',
  'eppm_taskindex',
  'eppm_pocpercentage',
  'eppm_taskwork',
  'eppm_predecessor',
  'eppm_successors',
  'eppm_notes',
  'eppm_parenttaskid',
  'eppm_calendarname',
  'eppm_ignoreresourcecalendar',
  'eppm_rollup',
  'eppm_inactive',
  'eppm_manuallyscheduled',
  'eppm_projectborder',
  'eppm_constrainttype',
  'eppm_constraintdate'
].join(',');

export async function fetchTasksFromDataverse(
  accessToken: string,
  projectId?: string
): Promise<DataverseTask[]> {
  const baseUrl = dataverseConfig.environmentUrl.replace(/\/$/, '');
  const url = projectId && projectId.trim()
    ? `${baseUrl}/api/data/v9.2/${dataverseConfig.tableName}?$select=${SELECT_FIELDS}&$filter=${encodeURIComponent("eppm_projectid eq '" + projectId + "'")}`
    : `${baseUrl}/api/data/v9.2/${dataverseConfig.tableName}?$select=${SELECT_FIELDS}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'Accept': 'application/json',
      'Authorization': `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    const text = await response.text();
    let message = `Dataverse request failed: ${response.status}`;
    try {
      const json = JSON.parse(text);
      message = (json.error && json.error.message) || message;
    } catch {
      if (text) message = text.slice(0, 200);
    }
    throw new Error(message);
  }

  const data = await response.json();
  if (data && data.value && Array.isArray(data.value)) {
    return data.value as DataverseTask[];
  }
  if (Array.isArray(data)) {
    return data as DataverseTask[];
  }
  return [];
}

export interface DataverseAssignment {
  eppm_taskassignmentsid?: string;
  eppm_taskid?: string;
  _eppm_taskid_value?: string;
  eppm_projecttaskid?: string;
  _eppm_projecttaskid_value?: string;
  eppm_resourceemail?: string;
  eppm_maxunits?: number;
  [key: string]: unknown;
}

export interface BryntumResource {
  id: string;
  name: string;
  email?: string;
}

export interface BryntumAssignment {
  id: string;
  event: string;
  resource: string;
  units: number;
}

function isGuid(value: any): boolean {
  return typeof value === 'string' && /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/i.test(value);
}

function extractTaskIdFromAssignment(row: DataverseAssignment): string | undefined {
  if (row.eppm_taskid && isGuid(row.eppm_taskid)) return row.eppm_taskid;
  if (row._eppm_taskid_value && isGuid(row._eppm_taskid_value)) return row._eppm_taskid_value;
  if (row.eppm_projecttaskid && isGuid(row.eppm_projecttaskid)) return row.eppm_projecttaskid;
  if (row._eppm_projecttaskid_value && isGuid(row._eppm_projecttaskid_value)) return row._eppm_projecttaskid_value;
  return undefined;
}

function extractAssignmentId(row: DataverseAssignment): string | undefined {
  if (row.eppm_taskassignmentsid && isGuid(row.eppm_taskassignmentsid)) return row.eppm_taskassignmentsid;
  return undefined;
}

function safeParseUnits(value: any): number {
  if (typeof value === 'number' && !isNaN(value) && isFinite(value)) return value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    // Avoid String.prototype.startsWith to stay ES5-compatible
    if (trimmed.length > 0 && trimmed.charAt(0) === '=') return 100;
    const parsed = parseFloat(trimmed);
    if (!isNaN(parsed) && isFinite(parsed)) return parsed;
  }
  return 100;
}

export async function fetchAssignmentsFromDataverse(
  accessToken: string,
  projectTaskIds: Set<string>
): Promise<{ resources: BryntumResource[]; assignments: BryntumAssignment[] }> {
  const TASK_ASSIGNMENTS_TABLE = 'eppm_taskassignmentses';
  const baseUrl = dataverseConfig.environmentUrl.replace(/\/$/, '');
  const url = `${baseUrl}/api/data/v9.2/${TASK_ASSIGNMENTS_TABLE}`;

  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'Accept': 'application/json',
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!response.ok) {
      console.warn('[fetchAssignments] Failed to fetch assignments:', response.status);
      return { resources: [], assignments: [] };
    }

    const data = await response.json();
    const assignmentRows: DataverseAssignment[] = (data && data.value && Array.isArray(data.value)) ? data.value : (Array.isArray(data) ? data : []);

    const resourceMap = new Map<string, BryntumResource>();
    const assignments: BryntumAssignment[] = [];

    for (const row of assignmentRows) {
      const taskId = extractTaskIdFromAssignment(row);
      if (!taskId || !projectTaskIds.has(taskId)) continue;

      const resourceName = row.eppm_resourceemail;
      if (typeof resourceName !== 'string' || !resourceName.trim()) continue;

      const nameTrimmed = resourceName.trim();
      const resourceId = nameTrimmed;

      if (!resourceMap.has(resourceId)) {
        resourceMap.set(resourceId, {
          id: resourceId,
          name: nameTrimmed,
          email: nameTrimmed
        });
      }

      const assignmentId = extractAssignmentId(row) || (taskId + '_' + resourceId);
      assignments.push({
        id: assignmentId,
        event: taskId,
        resource: resourceId,
        units: safeParseUnits(row.eppm_maxunits)
      });
    }

    const resources: BryntumResource[] = [];
    resourceMap.forEach(resource => {
      resources.push(resource);
    });

    return {
      resources: resources,
      assignments: assignments
    };
  } catch (error: any) {
    console.warn('[fetchAssignments] Error fetching assignments:', error.message || error);
    return { resources: [], assignments: [] };
  }
}
