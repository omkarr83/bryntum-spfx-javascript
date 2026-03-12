/**
 * Dataverse create/update (CRUD) from SPFx using fetch only.
 * No Node server: all calls go to Dataverse OData from the browser.
 */
import { dataverseConfig } from './dataverseConfig';

export function isGuid(value: unknown): value is string {
  return typeof value === 'string' && /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value);
}

const TASKS_TABLE = dataverseConfig.tableName;
const TASK_ASSIGNMENTS_TABLE = 'eppm_taskassignmentses';
const ASSIGNMENT_UNITS_FIELD = 'eppm_maxunits';

const BASE_URL = dataverseConfig.environmentUrl.replace(/\/$/, '');
const ODATA_BASE = `${BASE_URL}/api/data/v9.2`;

function getHeaders(accessToken: string): Record<string, string> {
  return {
    'Content-Type': 'application/json; charset=utf-8',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    'Accept': 'application/json',
    'Authorization': `Bearer ${accessToken}`
  };
}

async function handleResponse<T>(response: Response, operation: string): Promise<T> {
  if (!response.ok) {
    const text = await response.text();
    let message = `${operation} failed: ${response.status}`;
    try {
      const json = JSON.parse(text);
      message = (json.error && json.error.message) || message;
    } catch {
      if (text) message = text.slice(0, 300);
    }
    throw new Error(message);
  }
  if (response.status === 204) return undefined as unknown as T;
  return response.json();
}

/**
 * Create a task in Dataverse. Returns the created entity (with eppm_projecttaskid).
 */
export async function createTask(
  accessToken: string,
  payload: Record<string, unknown>
): Promise<{ eppm_projecttaskid?: string }> {
  const url = `${ODATA_BASE}/${TASKS_TABLE}`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      ...getHeaders(accessToken),
      'Prefer': 'return=representation'
    },
    body: JSON.stringify(payload)
  });
  return handleResponse(response, 'Create task');
}

/**
 * Update a task in Dataverse (PATCH).
 */
export async function updateTask(
  accessToken: string,
  taskId: string,
  payload: Record<string, unknown>
): Promise<void> {
  const url = `${ODATA_BASE}/${TASKS_TABLE}(${taskId})`;
  const response = await fetch(url, {
    method: 'PATCH',
    headers: {
      ...getHeaders(accessToken),
      'If-Match': '*'
    },
    body: JSON.stringify(payload)
  });
  await handleResponse(response, 'Update task');
}

/**
 * Create an assignment in Dataverse. Payload should include eppm_taskid (GUID), eppm_resourceemail, eppm_maxunits, eppm_projectid.
 */
export async function createAssignment(
  accessToken: string,
  payload: Record<string, unknown>
): Promise<{ eppm_taskassignmentsid?: string }> {
  const url = `${ODATA_BASE}/${TASK_ASSIGNMENTS_TABLE}`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      ...getHeaders(accessToken),
      'Prefer': 'return=representation'
    },
    body: JSON.stringify(payload)
  });
  return handleResponse(response, 'Create assignment');
}

/**
 * Update an assignment in Dataverse (PATCH).
 */
export async function updateAssignment(
  accessToken: string,
  assignmentId: string,
  payload: Record<string, unknown>
): Promise<void> {
  const url = `${ODATA_BASE}/${TASK_ASSIGNMENTS_TABLE}(${assignmentId})`;
  const response = await fetch(url, {
    method: 'PATCH',
    headers: {
      ...getHeaders(accessToken),
      'If-Match': '*'
    },
    body: JSON.stringify(payload)
  });
  await handleResponse(response, 'Update assignment');
}

/**
 * Delete an assignment in Dataverse.
 */
export async function deleteAssignment(
  accessToken: string,
  assignmentId: string
): Promise<void> {
  const url = `${ODATA_BASE}/${TASK_ASSIGNMENTS_TABLE}(${assignmentId})`;
  const response = await fetch(url, {
    method: 'DELETE',
    headers: {
      ...getHeaders(accessToken),
      'If-Match': '*'
    }
  });
  await handleResponse(response, 'Delete assignment');
}

/**
 * Fetch all tasks for a project (used to build task name -> ID map for assignments).
 */
export async function getTasksForProject(
  accessToken: string,
  projectId: string
): Promise<Array<{ eppm_projecttaskid?: string; eppm_name?: string }>> {
  const filter = encodeURIComponent("eppm_projectid eq '" + projectId + "'");
  const url = `${ODATA_BASE}/${TASKS_TABLE}?$select=eppm_projecttaskid,eppm_name&$filter=${filter}`;
  const response = await fetch(url, {
    method: 'GET',
    headers: getHeaders(accessToken)
  });
  const data = await handleResponse<{ value?: unknown[] }>(response, 'Get tasks');
  const value = data && (data as { value?: unknown[] }).value;
  return Array.isArray(value) ? value as Array<{ eppm_projecttaskid?: string; eppm_name?: string }> : [];
}

/** Build assignment payload for create/update (SPFx format). */
export function buildAssignmentPayload(
  projectId: string,
  taskDataverseId: string,
  resourceEmail: string,
  units: number,
  startDate?: string,
  finishDate?: string
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    eppm_resourceemail: resourceEmail,
    [ASSIGNMENT_UNITS_FIELD]: units,
    eppm_projectid: projectId,
    eppm_taskid: taskDataverseId
  };
  if (startDate) payload.eppm_startdate = startDate;
  if (finishDate) payload.eppm_finishdate = finishDate;
  return payload;
}
