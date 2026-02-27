import express, { Request, Response } from 'express';
import multer from 'multer';

import { DataverseService } from '../services/dataverse.service.js';
import {
    buildTaskHierarchy,
    flattenTaskHierarchy,
    bryntumToDataverseTask,
    BryntumProjectData
} from '../utils/dataTransformer.js';
import { generateMspdiXml, convertToMspdiFormat } from '../utils/mspdiExporter.js';
import {
    parseMspdiXml,
    convertImportedDataToBryntum,
    convertImportedDataToDataverse,
    buildPredecessorStringForTask,
    buildSuccessorStringForTask
} from '../utils/mspdiImporter.js';

const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 50 * 1024 * 1024, // 50MB limit
    },
    fileFilter: (_req: any, file: Express.Multer.File, cb: multer.FileFilterCallback) => {
        // Accept XML files and MPP files
        const allowedMimeTypes = [
            'application/xml',
            'text/xml',
            'application/vnd.ms-project',
            'application/x-ms-project'
        ];
        const allowedExtensions = ['.xml', '.mpp'];
        const ext = file.originalname.toLowerCase().slice(file.originalname.lastIndexOf('.'));

        if (allowedMimeTypes.includes(file.mimetype) || allowedExtensions.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Invalid file type. Only XML and MPP files are allowed.') as any, false);
        }
    }
});
const router = express.Router();

const TASK_ASSIGNMENTS_TABLE = process.env.DATAVERSE_TASK_ASSIGNMENTS_TABLE || 'eppm_taskassignmentses';
const TASKS_TABLE = process.env.DATAVERSE_TABLE_NAME || 'eppm_projecttasks';
// IMPORTANT:
// Dataverse lookups must be set using the *navigation property* via <nav>@odata.bind.
// In some environments, the lookup on eppm_taskassignmentses is not named "eppm_taskid"
// (and using the wrong property with @odata.bind causes a 400 like:
// "ODataPrimitiveValue was instantiated with ... ODataEntityReferenceLink").
//
// If you know the correct lookup navigation property, set it with DATAVERSE_ASSIGNMENT_TASK_LOOKUP.
// Otherwise we will try common candidates automatically.

// const PROJECT_ID_FILTER = process.env.DATAVERSE_PROJECT_ID || '78c1d16b-223c-4414-8896-399e3ae950c7';
// const PROJECT_ID_FILTER = process.env.DATAVERSE_PROJECT_ID || 'dda7b87a-46d1-f011-8544-7ced8d281f53';
const PROJECT_ID_FILTER = process.env.DATAVERSE_PROJECT_ID || '35d56841-1466-47b7-bd18-0e697474fdfd';
const TASK_FILTER = PROJECT_ID_FILTER ? `eppm_projectid eq '${PROJECT_ID_FILTER}'` : undefined;
const PROJECTS_TABLE = process.env.DATAVERSE_PROJECTS_TABLE || 'eppm_projects';

const ASSIGNMENT_TASK_LOOKUP_OVERRIDE = process.env.DATAVERSE_ASSIGNMENT_TASK_LOOKUP;
// Store units in this assignment column (per your requirement)
const ASSIGNMENT_UNITS_FIELD = process.env.DATAVERSE_ASSIGNMENT_UNITS_FIELD || 'eppm_maxunits';

// Store resource name and units on eppm_projecttasks
const TASK_RESOURCES_FIELD = process.env.DATAVERSE_TASK_RESOURCES_FIELD || 'eppm_resources';

function getAssignmentTaskLookupCandidates(): string[] {
    const override = typeof ASSIGNMENT_TASK_LOOKUP_OVERRIDE === 'string' ? ASSIGNMENT_TASK_LOOKUP_OVERRIDE.trim() : '';
    const candidates = [
        ...(override ? [override] : []),
        // Common lookup names we have seen for task assignment -> task
        // eppm_taskid is the primary lookup field for assignments
        'eppm_taskid',
        'eppm_projecttaskid'
    ];

    // De-dupe while keeping order
    return Array.from(new Set(candidates.filter(Boolean)));
}

/** Build lookup payload using _value format (raw GUID) - use when @odata.bind causes ODataEntityReferenceLink error */
function buildTaskLookupAsValue(lookupName: string, taskId: string): Record<string, string> {
    return { [`_${lookupName}_value`]: taskId };
}

function isGuid(value: any): value is string {
    return typeof value === 'string' && /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value);
}

function emailToDisplayName(email: string): string {
    const local = email.split('@')[0] || email;
    return local
        .split(/[._-]+/g)
        .filter(Boolean)
        .map(part => part.charAt(0).toUpperCase() + part.slice(1))
        .join(' ');
}

/**
 * Safely parse units value to ensure it's a valid number for Bryntum.
 * Prevents "Unknown formula for `units`" error by:
 * - Handling numbers directly
 * - Parsing string numbers
 * - Rejecting formula-like strings (starting with "=")
 * - Defaulting to 100 for invalid values
 */
function safeParseUnits(value: any): number {
    // Already a valid number
    if (typeof value === 'number' && Number.isFinite(value)) {
        return value;
    }

    // Handle string values
    if (typeof value === 'string') {
        const trimmed = value.trim();

        // Reject formula-like strings (Bryntum would try to parse "=..." as formula)
        if (trimmed.startsWith('=')) {
            console.warn(`[safeParseUnits] Rejected formula-like value: "${trimmed}"`);
            return 100;
        }

        // Try to parse as number
        const parsed = parseFloat(trimmed);
        if (Number.isFinite(parsed)) {
            return parsed;
        }
    }

    // Default for null, undefined, or unparseable values
    return 100;
}

function extractAssignmentId(row: any): string | undefined {
    // Common Dataverse primary key naming convention: <logicalname>id
    if (isGuid(row.eppm_taskassignmentsid)) return row.eppm_taskassignmentsid;
    if (isGuid(row.eppm_taskassignmentid)) return row.eppm_taskassignmentid;

    // Fallback: find any key that looks like task assignment id
    const keys = Object.keys(row || {});
    for (const k of keys) {
        if (!/id$/i.test(k)) continue;
        if (!/taskassign/i.test(k)) continue;
        const v = row[k];
        if (isGuid(v)) return v;
    }
    return undefined;
}

function extractTaskIdFromAssignment(row: any): string | undefined {
    // Your environment uses eppm_taskid (lookup to task)
    const taskIdDirect = row.eppm_taskid;
    if (isGuid(taskIdDirect)) return taskIdDirect;

    const taskIdLookup = row._eppm_taskid_value;
    if (isGuid(taskIdLookup)) return taskIdLookup;

    const direct = row.eppm_projecttaskid;
    if (isGuid(direct)) return direct;

    // Dataverse lookup raw guid form usually appears as _<lookup>_value
    const lookup = row._eppm_projecttaskid_value;
    if (isGuid(lookup)) return lookup;

    // Fallback: search
    const keys = Object.keys(row || {});
    for (const k of keys) {
        if (!/(projecttaskid|taskid)/i.test(k)) continue;
        const v = row[k];
        if (isGuid(v)) return v;
    }
    return undefined;
}

function extractProjectIdFromTaskRow(row: any): string | undefined {
    if (!row || typeof row !== 'object') return undefined;

    const direct = row.eppm_projectid;
    if (typeof direct === 'string' && direct.trim()) return direct.trim();

    const lookup = row._eppm_projectid_value;
    if (typeof lookup === 'string' && lookup.trim()) return lookup.trim();

    return undefined;
}

function ensureProjectId(payload: any, projectId?: string): void {
    const targetProjectId = projectId || PROJECT_ID_FILTER;
    if (!targetProjectId) {
        console.warn('[Tasks Route] No project ID available to set on task');
        return;
    }

    // If payload already has eppm_projectid, don't override
    if (payload.eppm_projectid) {
        return;
    }

    // Try direct GUID assignment first (most common)
    payload.eppm_projectid = targetProjectId;

    // Note: If direct assignment fails, you may need to use lookup binding format:
    // payload[`eppm_projectid@odata.bind`] = `/${PROJECTS_TABLE}(${targetProjectId})`;
    // delete payload.eppm_projectid;
}

function extractTaskId(row: any): string | undefined {
    if (!row || typeof row !== 'object') return undefined;
    if (isGuid(row.eppm_projecttaskid)) return row.eppm_projecttaskid;

    // Fallback: find any key that looks like a task PK
    const keys = Object.keys(row);
    for (const k of keys) {
        if (!/projecttaskid$/i.test(k)) continue;
        const v = (row as any)[k];
        if (isGuid(v)) return v;
    }
    return undefined;
}

function extractAssignmentsIdFromPayload(value: any): string | undefined {
    if (isGuid(value)) return value;
    if (typeof value === 'string' && isGuid(value)) return value;
    if (value && typeof value === 'object') {
        if (isGuid(value.id)) return value.id;
        if (isGuid(value.$PhantomId)) return value.$PhantomId;
    }
    return undefined;
}

// -------------------------
// Predecessor <-> Dependency helpers
// -------------------------
const dependencyTypeMap: Record<string, number> = {
    SS: 0,
    SF: 1,
    FS: 2,
    FF: 3
};

function typeToAbbr(type: any): string {
    if (typeof type === 'string') {
        const upper = type.toUpperCase();
        if (dependencyTypeMap[upper] !== undefined) return upper;
        if (upper.includes('FINISH') && upper.includes('START')) return 'FS';
        if (upper.includes('START') && upper.includes('START')) return 'SS';
        if (upper.includes('FINISH') && upper.includes('FINISH')) return 'FF';
        if (upper.includes('START') && upper.includes('FINISH')) return 'SF';
        return 'FS';
    }

    if (typeof type === 'number') {
        // Bryntum mapping: 0=SS,1=SF,2=FS,3=FF
        return Object.keys(dependencyTypeMap).find(k => dependencyTypeMap[k] === type) || 'FS';
    }

    // Default dependency type is Finish-to-Start
    return 'FS';
}

function abbrToType(abbr: string): number {
    const upper = abbr.toUpperCase();
    return dependencyTypeMap[upper] ?? 2;
}

function dependencyToToken(dep: any): string | null {
    const from = dep?.fromTask ?? dep?.from ?? dep?.fromEvent ?? dep?.fromTaskId ?? dep?.fromId;
    if (!from) return null;

    const abbr = typeToAbbr(dep?.type);
    const lag = dep?.lag;
    const lagUnit = dep?.lagUnit || dep?.lagunit;

    // Keep token format consistent with Bryntum predecessor column ("417FS")
    let token = `${from}${abbr}`;

    if (typeof lag === 'number' && lag !== 0) {
        const unit = typeof lagUnit === 'string' ? lagUnit[0] : 'd';
        token += `${lag > 0 ? '+' : ''}${lag}${unit}`;
    }

    return token;
}

function successorToToken(dep: any): string | null {
    // For successors, store token against the "from" task, pointing to its successor (to task)
    const to = dep?.toTask ?? dep?.to ?? dep?.toEvent ?? dep?.toTaskId ?? dep?.toId;
    if (!to) return null;

    const abbr = typeToAbbr(dep?.type);
    const lag = dep?.lag;
    const lagUnit = dep?.lagUnit || dep?.lagunit;

    let token = `${to}${abbr}`;
    if (typeof lag === 'number' && lag !== 0) {
        const unit = typeof lagUnit === 'string' ? lagUnit[0] : 'd';
        token += `${lag > 0 ? '+' : ''}${lag}${unit}`;
    }

    return token;
}

function normalizePredecessorString(value: any): string {
    if (typeof value !== 'string') return '';
    return value
        .split(/[;,]+/g)
        .map(s => s.trim())
        .filter(Boolean)
        .join(';');
}

function buildPredecessorStringFromArray(predecessors: any[]): string {
    if (!Array.isArray(predecessors)) return '';
    const tokens: string[] = [];
    for (const p of predecessors) {
        const token = dependencyToToken(p);
        if (token) tokens.push(token);
    }
    return tokens.join(';');
}

function normalizeSuccessorString(value: any): string {
    if (typeof value !== 'string') return '';
    return value
        .split(/[;,]+/g)
        .map(s => s.trim())
        .filter(Boolean)
        .join(';');
}

function buildSuccessorStringFromArray(successors: any[]): string {
    if (!Array.isArray(successors)) return '';
    const tokens: string[] = [];
    for (const s of successors) {
        const token = successorToToken(s);
        if (token) tokens.push(token);
    }
    return tokens.join(';');
}

function parsePredecessorString(str: string, toTaskId: string): any[] {
    const cleaned = normalizePredecessorString(str);
    if (!cleaned) return [];

    const parts = cleaned.split(';').map(s => s.trim()).filter(Boolean);
    const deps: any[] = [];

    for (const [idx, part] of parts.entries()) {
        // Try to match "...(FS|SS|FF|SF)(+/-lag?)"
        const m = part.match(/^(.*?)(FS|SS|FF|SF)([+-]\d+(?:\.\d+)?[a-zA-Z]?)?$/i);
        if (!m) continue;

        const fromTask = m[1]?.trim();
        const abbr = m[2].toUpperCase();
        const lagPart = m[3];

        // Skip if fromTask is empty or invalid
        if (!fromTask) continue;

        const type = abbrToType(abbr);
        let lag: number | undefined = undefined;
        if (lagPart) {
            const n = parseFloat(lagPart);
            if (!Number.isNaN(n) && n !== 0) lag = n;
        }

        deps.push({
            id: `${toTaskId}_${fromTask}_${abbr}_${idx}`,
            fromTask,
            toTask: toTaskId,
            type,
            ...(lag !== undefined ? { lag } : {})
        });
    }

    return deps;
}

function parseSuccessorString(str: string, fromTaskId: string): any[] {
    const cleaned = normalizeSuccessorString(str);
    if (!cleaned) return [];

    const parts = cleaned.split(';').map(s => s.trim()).filter(Boolean);
    const deps: any[] = [];

    for (const [idx, part] of parts.entries()) {
        const m = part.match(/^(.*?)(FS|SS|FF|SF)([+-]\d+(?:\.\d+)?[a-zA-Z]?)?$/i);
        if (!m) continue;

        const toTask = m[1]?.trim();
        const abbr = m[2].toUpperCase();
        const lagPart = m[3];

        // Skip if toTask is empty or invalid
        if (!toTask) continue;

        const type = abbrToType(abbr);
        let lag: number | undefined = undefined;
        if (lagPart) {
            const n = parseFloat(lagPart);
            if (!Number.isNaN(n) && n !== 0) lag = n;
        }

        deps.push({
            id: `${fromTaskId}_${toTask}_${abbr}_succ_${idx}`,
            fromTask: fromTaskId,
            toTask,
            type,
            ...(lag !== undefined ? { lag } : {})
        });
    }

    return deps;
}

/**
 * Middleware to log incoming requests and set response headers
 */
router.use((req: Request, res: Response, next) => {
    // console.log(`\n[${new Date().toISOString()}] ===== ${req.method} ${req.path} =====`);
    // console.log('[Request Headers - Raw]:', JSON.stringify(req.headers, null, 2));
    // Log Authorization header specifically
    const authHeader = req.headers.authorization || req.headers.Authorization || req.headers['authorization'] || req.headers['Authorization'];
    const authHeaderStr = Array.isArray(authHeader) ? authHeader[0] : authHeader;

    const authLowercase = req.headers['authorization'];
    const authLowercaseStr = Array.isArray(authLowercase) ? authLowercase[0] : authLowercase;

    const authHeaderPreview = authHeaderStr && typeof authHeaderStr === 'string'
        ? authHeaderStr.substring(0, 30) + '...'
        : (authHeaderStr ? 'Present' : 'MISSING');

    console.log('[Request Headers - Summary]:', {
        authorization: authHeaderPreview,
        'authorization-lowercase': authLowercaseStr ? 'Present' : 'Missing',
        'content-type': req.headers['content-type'] || 'Not set',
        'accept': req.headers.accept || 'Not set',
        origin: req.headers.origin || 'Not set',
        'all-header-keys': Object.keys(req.headers).join(', ')
    });

    if (authHeaderStr && typeof authHeaderStr === 'string') {
        console.log('[Authorization Header Found]:', authHeaderStr.substring(0, 50) + '...');
    } else {
        console.error('[Authorization Header]: NOT FOUND');
        console.error('[All Headers]:', Object.keys(req.headers));
    }

    // Set response headers for all routes
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
    res.setHeader('Access-Control-Allow-Credentials', 'true');

    next();
});

/**
 * Middleware to extract access token from request
 */
function getAccessToken(req: Request): string | null {
    console.log('[getAccessToken] Checking for token in request...');
    console.log('[getAccessToken] req.headers.authorization:', req.headers.authorization ? 'exists' : 'missing');
    console.log('[getAccessToken] req.headers.Authorization:', req.headers.Authorization ? 'exists' : 'missing');

    // Check Authorization header (case-insensitive) - try multiple variations
    const authHeader = req.headers.authorization
        || req.headers.Authorization
        || (req.headers as any)['authorization']
        || (req.headers as any)['Authorization'];

    if (authHeader) {
        console.log('[getAccessToken] Auth header found, type:', typeof authHeader);
        console.log('[getAccessToken] Auth header value:', typeof authHeader === 'string' ? authHeader.substring(0, 50) + '...' : authHeader);

        if (typeof authHeader === 'string') {
            if (authHeader.startsWith('Bearer ')) {
                const token = authHeader.substring(7);
                console.log('[getAccessToken] ✓ Token extracted, length:', token.length);
                return token;
            } else {
                console.warn('[getAccessToken] Auth header does not start with "Bearer "');
                console.warn('[getAccessToken] Auth header starts with:', authHeader.substring(0, 20));
            }
        } else if (Array.isArray(authHeader) && authHeader.length > 0) {
            const headerValue = authHeader[0];
            if (typeof headerValue === 'string' && headerValue.startsWith('Bearer ')) {
                const token = headerValue.substring(7);
                console.log('[getAccessToken] ✓ Token extracted from array, length:', token.length);
                return token;
            }
        }
    }

    // Also check if token is passed as a query parameter (fallback)
    const tokenParam = req.query.token as string;
    if (tokenParam) {
        console.log('[getAccessToken] ✓ Token found in query parameter');
        return tokenParam;
    }

    console.error('[getAccessToken] ✗ No token found in request');
    return null;
}

/**
 * Create DataverseService instance with token from request
 */
function createDataverseService(req: Request): DataverseService {
    const token = getAccessToken(req);
    if (!token) {
        throw new Error('No access token provided');
    }
    return new DataverseService(token);
}

/**
 * Helper function to set response headers consistently
 */
function setResponseHeaders(res: Response, req: Request): void {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
}

/**
 * GET /api/tasks - Get all tasks in Bryntum format
 */
router.get('/', async (req: Request, res: Response) => {
    try {
        // Check for token first
        const token = getAccessToken(req);
        if (!token) {
            console.error('[Tasks Route] No access token provided in request headers');
            console.error('[Tasks Route] Request headers:', JSON.stringify(req.headers, null, 2));
            console.error('[Tasks Route] Authorization header:', req.headers.authorization || req.headers.Authorization || 'Not found');
            console.error('[Tasks Route] Query params:', JSON.stringify(req.query, null, 2));
            setResponseHeaders(res, req);
            return res.status(401).json({
                success: false,
                error: 'No access token provided. Please authenticate first.',
                debug: {
                    hasAuthHeader: !!(req.headers.authorization || req.headers.Authorization),
                    headerKeys: Object.keys(req.headers),
                    origin: req.headers.origin || 'Not set'
                }
            });
        }

        // Validate token format (should start with Bearer or be a JWT)
        if (token.length < 50) {
            console.error('[Tasks Route] Token appears to be invalid (too short)');
            setResponseHeaders(res, req);
            return res.status(401).json({
                success: false,
                error: 'Invalid access token format.'
            });
        }

        console.log('[Tasks Route] Token received (length:', token.length, '), fetching tasks from Dataverse...');
        console.log('[Tasks Route] Token preview:', token.substring(0, 20) + '...');

        // Get projectId from query parameter (required - comes from dropdown selection)
        const projectId = req.query.projectId as string | undefined;
        if (!projectId || projectId.trim() === '') {
            console.log('[Tasks Route] No projectId provided - returning empty data');
            setResponseHeaders(res, req);
            return res.json({
                success: true,
                tasks: { rows: [] },
                dependencies: { rows: [] },
                resources: { rows: [] },
                assignments: { rows: [] },
                calendars: {
                    rows: [{
                        id: 'general',
                        name: 'General',
                        intervals: [{ recurrentStartDate: 'on Sat', recurrentEndDate: 'on Mon', isWorking: false }],
                        expanded: true
                    }]
                }
            });
        }
        const taskFilter = `eppm_projectid eq '${projectId}'`;
        console.log('[Tasks Route] Using projectId:', projectId);
        console.log('[Tasks Route] Task filter:', taskFilter);

        const dataverseService = createDataverseService(req);
        const dataverseTasks = await dataverseService.getAllTasks(taskFilter);
        console.log(`[Tasks Route] Fetched ${dataverseTasks?.length || 0} tasks from Dataverse`);

        if (dataverseTasks && dataverseTasks.length > 0) {
            console.log('[Tasks Route] Sample task data (first 3):');
            dataverseTasks.slice(0, 3).forEach((t, i) => {
                console.log(`  Task ${i + 1}: name="${t.eppm_name}", startDate="${t.eppm_startdate}", finishDate="${t.eppm_finishdate}", duration=${t.eppm_taskduration}`);
            });
        }

        if (!dataverseTasks || dataverseTasks.length === 0) {
            setResponseHeaders(res, req);
            return res.json({
                success: true,
                project: {
                    calendar: 'general',
                    startDate: new Date().toISOString().split('T')[0],
                    hoursPerDay: 8,
                    daysPerWeek: 5,
                    daysPerMonth: 20
                },
                tasks: {
                    rows: []
                },
                dependencies: {
                    rows: []
                },
                resources: {
                    rows: []
                },
                assignments: {
                    rows: []
                },
                calendars: {
                    rows: [
                        {
                            id: 'general',
                            name: 'General',
                            intervals: [
                                {
                                    recurrentStartDate: 'on Sat',
                                    recurrentEndDate: 'on Mon',
                                    isWorking: false
                                }
                            ],
                            expanded: true
                        }
                    ]
                }
            });
        }

        const hierarchicalTasks = buildTaskHierarchy(dataverseTasks);

        if (hierarchicalTasks && hierarchicalTasks.length > 0) {
            console.log('[Tasks Route] Sample Bryntum task data (first 5):');
            const logTask = (t: any, indent: string = '  ') => {
                console.log(`${indent}Task: name="${t.name}", startDate="${t.startDate}", endDate="${t.endDate}", duration=${t.duration}, children=${t.children?.length || 0}`);
                if (t.children) {
                    t.children.slice(0, 2).forEach((child: any) => logTask(child, indent + '  '));
                }
            };
            hierarchicalTasks.slice(0, 5).forEach((t) => logTask(t));
        }
        // Build dependency rows from stored predecessor/successor strings
        const dependencyRows: any[] = [];
        const depKeySet = new Set<string>();

        const pushUnique = (d: any) => {
            const key = `${d?.fromTask ?? ''}->${d?.toTask ?? ''}:${d?.type ?? ''}:${d?.lag ?? ''}`;
            if (!depKeySet.has(key)) {
                depKeySet.add(key);
                dependencyRows.push(d);
            }
        };

        for (const t of dataverseTasks) {
            const taskId = t.eppm_projecttaskid;
            if (!taskId) continue;

            const pred = (t as any).eppm_predecessor;
            if (typeof pred === 'string' && pred.trim()) {
                for (const d of parsePredecessorString(pred, String(taskId))) {
                    pushUnique(d);
                }
            }

            const succ = (t as any).eppm_successors;
            if (typeof succ === 'string' && succ.trim()) {
                for (const d of parseSuccessorString(succ, String(taskId))) {
                    pushUnique(d);
                }
            }
        }

        // Fetch resources from eppm_taskassignmentses (eppm_resourceemail = resource name, eppm_maxunits = units)
        // Filter assignments to only include those for tasks in the selected project
        let resourcesRows: any[] = [];
        let assignmentsRows: any[] = [];
        try {
            // Build a set of task IDs for the current project to filter assignments
            const projectTaskIds = new Set<string>();
            for (const t of dataverseTasks) {
                const taskId = t.eppm_projecttaskid;
                if (taskId) {
                    projectTaskIds.add(taskId);
                }
            }
            console.log(`[Tasks Route] Project has ${projectTaskIds.size} tasks for filtering assignments`);

            const assignmentRows = await dataverseService.getTableRows<any>(TASK_ASSIGNMENTS_TABLE);
            console.log(`[Tasks Route] Fetched ${assignmentRows?.length || 0} total assignment rows`);

            const resourceMap = new Map<string, any>();
            for (const row of assignmentRows) {
                // First check if this assignment belongs to a task in the current project
                const taskId = extractTaskIdFromAssignment(row);
                if (!taskId || !projectTaskIds.has(taskId)) {
                    // Skip assignments that don't belong to tasks in this project
                    continue;
                }

                const resourceName = row.eppm_resourceemail;
                if (typeof resourceName !== 'string' || !resourceName.trim()) continue;

                const nameTrimmed = resourceName.trim();
                // Use resource name as id for Bryntum (eppm_resourceemail stores the name)
                const resourceId = nameTrimmed;
                if (!resourceMap.has(resourceId)) {
                    resourceMap.set(resourceId, {
                        id: resourceId,
                        name: nameTrimmed,
                        email: nameTrimmed
                    });
                }

                const assignmentId = extractAssignmentId(row) || `${taskId}_${resourceId}`;
                // Always include units with a valid numeric value to prevent "Unknown formula for `units`" error
                // Bryntum's scheduling engine requires units to be present
                assignmentsRows.push({
                    id: assignmentId,
                    event: taskId,
                    resource: resourceId,
                    units: 100  // Always provide a valid numeric value
                });
            }

            resourcesRows = Array.from(resourceMap.values());
            console.log(`[Tasks Route] Filtered to ${assignmentsRows.length} assignments and ${resourcesRows.length} resources for project`);
        } catch (e: any) {
            console.warn('[Tasks Route] Failed to fetch resources/assignments:', e?.message || e);
            resourcesRows = [];
            assignmentsRows = [];
        }

        // Find the earliest start date for project startDate
        let projectStartDate = new Date().toISOString().split('T')[0];
        let projectEndDate = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]; // Default to 30 days from now

        // Collect all dates from tasks (including children)
        const collectDates = (tasks: any[]): { starts: string[], ends: string[] } => {
            const starts: string[] = [];
            const ends: string[] = [];
            for (const task of tasks) {
                if (task.startDate) starts.push(task.startDate);
                if (task.endDate) ends.push(task.endDate);
                if (task.children) {
                    const childDates = collectDates(task.children);
                    starts.push(...childDates.starts);
                    ends.push(...childDates.ends);
                }
            }
            return { starts, ends };
        };
        // if (hierarchicalTasks.length > 0) {
        //     const allDates = hierarchicalTasks
        //         .map(task => task.startDate)
        //         .filter(date => date !== undefined && date !== null) as string[];
        //     if (allDates.length > 0) {
        //         projectStartDate = allDates.sort()[0];
        //     }
        // }

        if (hierarchicalTasks.length > 0) {
            const { starts, ends } = collectDates(hierarchicalTasks);
            if (starts.length > 0) {
                projectStartDate = starts.sort()[0];
            }
            if (ends.length > 0) {
                projectEndDate = ends.sort().reverse()[0];
            } else if (starts.length > 0) {
                // If no end dates, set project end date to 60 days after earliest start
                const earliestStart = new Date(starts.sort()[0]);
                projectEndDate = new Date(earliestStart.getTime() + 60 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
            }
        }
        // Filter out any invalid data to prevent "Cannot read properties of undefined" errors
        const validAssignments = assignmentsRows.filter((a: any) => a && a.id && a.event && a.resource);
        const validDependencies = dependencyRows.filter((d: any) => d && d.id && d.fromTask && d.toTask);
        const validResources = resourcesRows.filter((r: any) => r && r.id);

        const response: BryntumProjectData = {
            success: true,
            project: {
                calendar: 'general',
                startDate: projectStartDate,
                hoursPerDay: 8,
                daysPerWeek: 5,
                daysPerMonth: 20
            },
            tasks: {
                rows: hierarchicalTasks
            },
            dependencies: {
                rows: validDependencies
            },
            resources: {
                rows: validResources
            },
            assignments: {
                rows: validAssignments
            },
            calendars: {
                rows: [
                    {
                        id: 'general',
                        name: 'General',
                        intervals: [
                            {
                                recurrentStartDate: 'on Sat',
                                recurrentEndDate: 'on Mon',
                                isWorking: false
                            }
                        ],
                        expanded: true
                    }
                ]
            }
        };

        setResponseHeaders(res, req);
        res.json(response);
    } catch (error: any) {
        console.error('[Tasks Route] Error fetching tasks:', error);
        console.error('[Tasks Route] Error stack:', error.stack);

        // Check if it's a 401 from Dataverse
        if (error.response?.status === 401) {
            console.error('[Tasks Route] Dataverse returned 401 Unauthorized');
            console.error('[Tasks Route] Dataverse error response:', JSON.stringify(error.response?.data, null, 2));
            setResponseHeaders(res, req);
            return res.status(401).json({
                success: false,
                error: 'Authentication failed. Token may be expired or invalid. Please login again.',
                details: process.env.NODE_ENV === 'development' ? {
                    dataverseError: error.response?.data,
                    message: error.message
                } : undefined
            });
        }

        const errorMessage = error.message || 'Failed to fetch tasks';
        const statusCode = error.message?.includes('No access token') || error.response?.status === 401 ? 401 : 500;
        setResponseHeaders(res, req);
        res.status(statusCode).json({
            success: false,
            error: errorMessage,
            details: process.env.NODE_ENV === 'development' ? {
                stack: error.stack,
                response: error.response?.data,
                status: error.response?.status
            } : undefined
        });
    }
});

/**
 * GET /api/tasks/export/mpp - Export tasks, resources, and assignments to MS Project XML format
 * Returns an MSPDI XML file that can be opened by Microsoft Project and saved as .mpp
 * IMPORTANT: This route must be defined before /:id to avoid matching "export" as a task ID
 */
router.get('/export/mpp', async (req: Request, res: Response) => {
    try {
        console.log('[Export MPP] Starting export...');

        // Check for token first
        const token = getAccessToken(req);
        if (!token) {
            console.error('[Export MPP] No access token provided');
            return res.status(401).json({
                success: false,
                error: 'No access token provided. Please authenticate first.'
            });
        }

        // Get projectId from query parameter (required)
        const projectId = req.query.projectId as string | undefined;
        if (!projectId || projectId.trim() === '') {
            console.error('[Export MPP] No projectId provided');
            return res.status(400).json({
                success: false,
                error: 'projectId query parameter is required for export'
            });
        }

        const dataverseService = createDataverseService(req);

        // Fetch tasks filtered by the selected projectId
        const taskFilter = `eppm_projectid eq '${projectId}'`;
        console.log('[Export MPP] Fetching tasks with filter:', taskFilter);
        const dataverseTasks = await dataverseService?.getAllTasks(taskFilter);
        console.log(`[Export MPP] Fetched ${dataverseTasks?.length || 0} tasks for project ${projectId}`);

        // Build hierarchical task structure
        const hierarchicalTasks = dataverseTasks && dataverseTasks.length > 0
            ? buildTaskHierarchy(dataverseTasks)
            : [];

        // Fetch resources and assignments from eppm_taskassignmentses
        // Filter to only include assignments for tasks in the selected project
        let resourcesRows: any[] = [];
        let assignmentsRows: any[] = [];
        const dependencyRows: any[] = [];

        try {
            // Build a set of task IDs for the current project to filter assignments
            const projectTaskIds = new Set<string>();
            for (const t of dataverseTasks || []) {
                const taskId = t.eppm_projecttaskid;
                if (taskId) {
                    projectTaskIds.add(taskId);
                }
            }
            console.log(`[Export MPP] Project has ${projectTaskIds.size} tasks for filtering assignments`);

            const assignmentRows = await dataverseService.getTableRows<any>(TASK_ASSIGNMENTS_TABLE);
            console.log(`[Export MPP] Fetched ${assignmentRows?.length || 0} total assignment rows`);

            const resourceMap = new Map<string, any>();
            for (const row of assignmentRows) {
                // First check if this assignment belongs to a task in the current project
                const taskId = extractTaskIdFromAssignment(row);
                if (!taskId || !projectTaskIds.has(taskId)) {
                    // Skip assignments that don't belong to tasks in this project
                    continue;
                }

                const resourceName = row.eppm_resourceemail;
                if (typeof resourceName !== 'string' || !resourceName.trim()) continue;

                const nameTrimmed = resourceName.trim();
                if (!resourceMap.has(nameTrimmed)) {
                    resourceMap.set(nameTrimmed, {
                        id: nameTrimmed,
                        name: nameTrimmed,
                        email: nameTrimmed
                    });
                }

                // Map assignments (eppm_resourceemail = resource name, eppm_maxunits = units)
                const assignmentId = extractAssignmentId(row) || `${taskId}_${nameTrimmed}`;
                // Always include units with a valid numeric value
                assignmentsRows.push({
                    id: assignmentId,
                    event: taskId,
                    resource: nameTrimmed,
                    units: 100  // Always provide a valid numeric value
                });
            }

            resourcesRows = Array.from(resourceMap.values());
            console.log(`[Export MPP] Filtered to ${assignmentsRows.length} assignments and ${resourcesRows.length} resources for project`);
        } catch (e: any) {
            console.warn('[Export MPP] Failed to fetch resources/assignments:', e?.message || e);
        }

        // Build dependency rows from stored predecessor/successor strings
        const depKeySet = new Set<string>();
        const pushUniqueExport = (d: any) => {
            const key = `${d?.fromTask ?? ''}->${d?.toTask ?? ''}:${d?.type ?? ''}:${d?.lag ?? ''}`;
            if (!depKeySet.has(key)) {
                depKeySet.add(key);
                dependencyRows.push(d);
            }
        };

        for (const t of dataverseTasks || []) {
            const taskId = t.eppm_projecttaskid;
            if (!taskId) continue;

            const pred = (t as any).eppm_predecessor;
            if (typeof pred === 'string' && pred.trim()) {
                for (const d of parsePredecessorString(pred, String(taskId))) {
                    pushUniqueExport(d);
                }
            }

            const succ = (t as any).eppm_successors;
            if (typeof succ === 'string' && succ.trim()) {
                for (const d of parseSuccessorString(succ, String(taskId))) {
                    pushUniqueExport(d);
                }
            }
        }
        console.log(`[Export MPP] Found ${dependencyRows.length} dependencies`);

        // Get project name from query parameter or use default
        const projectName = (req.query.projectName as string) || 'Exported Project';

        // Convert to MSPDI format
        console.log('[Export MPP] Converting to MSPDI format...');
        const mspdiData = convertToMspdiFormat(
            hierarchicalTasks,
            resourcesRows,
            assignmentsRows,
            dependencyRows,
            projectName
        );

        // Generate XML
        console.log('[Export MPP] Generating XML...');
        const xmlContent = generateMspdiXml(mspdiData);

        // Set response headers for file download
        const filename = `${projectName.replace(/[^a-zA-Z0-9_-]/g, '_')}_${new Date().toISOString().split('T')[0]}.xml`;

        res.setHeader('Content-Type', 'application/xml; charset=utf-8');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
        res.setHeader('Access-Control-Allow-Credentials', 'true');
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        console.log(`[Export MPP] Export complete. File: ${filename}`);
        res.send(xmlContent);
    } catch (error: any) {
        console.error('[Export MPP] Error exporting:', error);
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
        res.status(500).json({
            success: false,
            error: error.message || 'Failed to export to MS Project format'
        });
    }
});

/**
 * POST /api/tasks/import/mpp - Import tasks, resources, and assignments from MS Project XML format
 * Accepts MSPDI XML files (exported from Microsoft Project)
 * Saves imported data to Dataverse and streams progress updates in real-time
 *
 * Response format: Newline-delimited JSON (NDJSON) for streaming progress updates
 * Each line is a JSON object with type: 'progress' | 'complete' | 'error'
 */
router.post('/import/mpp', upload.single('file'), async (req: Request, res: Response) => {
    // Helper to send progress update
    const sendProgress = (data: any) => {
        try {
            res.write(JSON.stringify(data) + '\n');
        } catch (e) {
            // Ignore write errors if connection closed
        }
    };

    try {
        console.log('[Import MPP] Starting import...');

        // Check for token first
        const token = getAccessToken(req);
        if (!token) {
            console.error('[Import MPP] No access token provided');
            return res.status(401).json({
                success: false,
                error: 'No access token provided. Please authenticate first.'
            });
        }

        // Get projectId from query parameter (required)
        const projectId = req.query.projectId as string | undefined;
        if (!projectId || projectId.trim() === '') {
            console.error('[Import MPP] No projectId provided');
            return res.status(400).json({
                success: false,
                error: 'projectId query parameter is required for import. Please select a project first.'
            });
        }
        console.log(`[Import MPP] Importing for project: ${projectId}`);

        // Check if file was uploaded
        if (!req.file) {
            return res.status(400).json({
                success: false,
                error: 'No file uploaded. Please select an XML file to import.'
            });
        }

        console.log(`[Import MPP] File received: ${req.file.originalname}, size: ${req.file.size} bytes`);

        // Check file extension
        const ext = req.file.originalname.toLowerCase().slice(req.file.originalname.lastIndexOf('.'));
        if (ext === '.mpp') {
            return res.status(400).json({
                success: false,
                error: 'Native .mpp files are not directly supported. Please export your project from Microsoft Project as XML (File > Save As > XML Format) and upload the XML file instead.'
            });
        }

        // Set up streaming response headers
        res.setHeader('Content-Type', 'application/x-ndjson');
        res.setHeader('Transfer-Encoding', 'chunked');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');
        res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
        res.setHeader('Access-Control-Allow-Credentials', 'true');

        // Send initial progress
        sendProgress({
            type: 'progress',
            stage: 'parsing',
            message: 'Parsing XML file...',
            progress: { tasksCreated: 0, tasksUpdated: 0, tasksFailed: 0, assignmentsCreated: 0, assignmentsUpdated: 0, assignmentsFailed: 0, dependenciesProcessed: 0 }
        });

        // Parse the XML content
        const xmlContent = req.file.buffer.toString('utf-8');
        console.log('[Import MPP] Parsing XML content...');

        let importedData;
        try {
            importedData = await parseMspdiXml(xmlContent);
        } catch (parseError: any) {
            console.error('[Import MPP] XML parsing error:', parseError);
            sendProgress({
                type: 'error',
                error: `Failed to parse XML file: ${parseError.message}. Please ensure the file is a valid MS Project XML (MSPDI) format.`
            });
            return res.end();
        }

        console.log(`[Import MPP] Parsed: ${importedData.tasks.length} tasks, ${importedData.resources.length} resources, ${importedData.assignments.length} assignments, ${importedData.dependencies.length} dependencies`);

        // Send parsing complete progress
        sendProgress({
            type: 'progress',
            stage: 'parsed',
            message: `Parsed ${importedData.tasks.length} tasks, ${importedData.assignments.length} assignments`,
            totalTasks: importedData.tasks.length,
            totalAssignments: importedData.assignments.length,
            progress: { tasksCreated: 0, tasksUpdated: 0, tasksFailed: 0, assignmentsCreated: 0, assignmentsUpdated: 0, assignmentsFailed: 0, dependenciesProcessed: 0 }
        });

        // Convert to Bryntum format for display
        const bryntumData = convertImportedDataToBryntum(importedData);
        console.log(`[Import MPP] After Bryntum conversion: ${bryntumData.tasks.length} tasks, ${bryntumData.resources.length} resources, ${bryntumData.assignments.length} assignments`);

        // Log sample assignments for debugging
        if (bryntumData.assignments.length > 0) {
            console.log('[Import MPP] Sample assignments (first 3):');
            bryntumData.assignments.slice(0, 3).forEach((a, i) => {
                console.log(`  Assignment ${i + 1}: event=${a.event}, resource=${a.resource}, units=${a.units}`);
            });
        } else {
            console.warn('[Import MPP] WARNING: No assignments after Bryntum conversion!');
            console.log('[Import MPP] Imported resources:', importedData.resources.map(r => ({ uid: r.uid, name: r.name, email: r.email })));
            console.log('[Import MPP] Imported assignments:', importedData.assignments.slice(0, 5));
        }

        // Convert to Dataverse format for saving (using projectId from validation above)
        const dataverseData = convertImportedDataToDataverse(bryntumData, projectId);
        console.log(`[Import MPP] After Dataverse conversion: ${dataverseData.tasks.length} tasks, ${dataverseData.assignments.length} assignments`);

        // Save to Dataverse
        const dataverseService = createDataverseService(req);
        console.log('[Import MPP] Saving tasks to Dataverse...');

        // Map import IDs to Dataverse IDs
        const importIdToDataverseId = new Map<string, string>();
        const processedTasks: any[] = [];
        const taskFailures: Array<{ name: string; error: string; action: string }> = [];
        let tasksCreated = 0;
        let tasksUpdated = 0;

        // Sort tasks by outline level to ensure parents are processed first
        const sortedTasks = [...dataverseData.tasks].sort((a, b) => {
            const levelA = bryntumData.tasks.find(t => t.id === a._importId)?._outlineLevel || 0;
            const levelB = bryntumData.tasks.find(t => t.id === b._importId)?._outlineLevel || 0;
            return levelA - levelB;
        });

        const totalTasks = sortedTasks.length;
        const totalAssignments = dataverseData.assignments.length;

        // Process tasks - update existing or create new
        for (const task of sortedTasks) {
            try {
                // Prepare task payload
                const taskPayload: any = {
                    eppm_name: task.eppm_name,
                    eppm_startdate: task.eppm_startdate,
                    eppm_finishdate: task.eppm_finishdate,
                    eppm_taskduration: task.eppm_taskduration,
                    eppm_pocpercentage: task.eppm_pocpercentage,
                    eppm_taskwork: task.eppm_taskwork,
                    eppm_notes: task.eppm_notes,
                    // Task index for sorting (ID column)
                    eppm_taskindex: task.eppm_taskindex
                };

                // Ensure eppm_projectid is set (required field in Dataverse)
                ensureProjectId(taskPayload, task.eppm_projectid);

                // Check if task has existing Dataverse ID (round-trip update)
                const existingTaskId = task._dataverseTaskId;
                const isExistingTask = existingTaskId && isGuid(existingTaskId);

                if (isExistingTask) {
                    // UPDATE existing task
                    console.log(`[Import MPP] Updating existing task: ${task.eppm_name} (${existingTaskId})`);

                    // Set parent if exists and was already processed
                    if (task._parentImportId) {
                        const parentDataverseId = importIdToDataverseId.get(task._parentImportId);
                        if (parentDataverseId) {
                            taskPayload.eppm_parenttaskid = parentDataverseId;
                        }
                    }

                    try {
                        await dataverseService.updateTask(existingTaskId, taskPayload);
                        importIdToDataverseId.set(task._importId, existingTaskId);
                        processedTasks.push({
                            ...task,
                            eppm_projecttaskid: existingTaskId,
                            _dataverseId: existingTaskId,
                            _action: 'updated'
                        });
                        tasksUpdated++;
                        console.log(`[Import MPP] Updated task: ${task.eppm_name} -> ${existingTaskId}`);
                    } catch (updateError: any) {
                        // If update fails (task may have been deleted), try to create new
                        console.warn(`[Import MPP] Update failed for ${task.eppm_name}, attempting to create new: ${updateError?.message}`);

                        // Set parent for new task
                        if (task._parentImportId) {
                            const parentDataverseId = importIdToDataverseId.get(task._parentImportId);
                            if (parentDataverseId) {
                                taskPayload.eppm_parenttaskid = parentDataverseId;
                            }
                        }

                        const created = await dataverseService.createRow<any>(TASKS_TABLE, taskPayload);
                        const createdId = created?.eppm_projecttaskid;

                        if (createdId) {
                            importIdToDataverseId.set(task._importId, createdId);
                            processedTasks.push({
                                ...task,
                                eppm_projecttaskid: createdId,
                                _dataverseId: createdId,
                                _action: 'created'
                            });
                            tasksCreated++;
                            console.log(`[Import MPP] Created task (after failed update): ${task.eppm_name} -> ${createdId}`);
                        } else {
                            taskFailures.push({ name: task.eppm_name, error: 'No ID returned from Dataverse', action: 'create' });
                        }
                    }
                } else {
                    // CREATE new task
                    // Set parent if exists and was already created
                    if (task._parentImportId) {
                        const parentDataverseId = importIdToDataverseId.get(task._parentImportId);
                        if (parentDataverseId) {
                            taskPayload.eppm_parenttaskid = parentDataverseId;
                        }
                    }

                    const created = await dataverseService.createRow<any>(TASKS_TABLE, taskPayload);
                    const createdId = created?.eppm_projecttaskid;

                    if (createdId) {
                        importIdToDataverseId.set(task._importId, createdId);
                        processedTasks.push({
                            ...task,
                            eppm_projecttaskid: createdId,
                            _dataverseId: createdId,
                            _action: 'created'
                        });
                        tasksCreated++;
                        console.log(`[Import MPP] Created task: ${task.eppm_name} -> ${createdId}`);
                    } else {
                        taskFailures.push({ name: task.eppm_name, error: 'No ID returned from Dataverse', action: 'create' });
                    }
                }
            } catch (error: any) {
                console.error(`[Import MPP] Failed to process task ${task.eppm_name}:`, error?.message || error);
                taskFailures.push({ name: task.eppm_name, error: error?.message || 'Unknown error', action: 'process' });
            }

            // Send progress update after each task
            sendProgress({
                type: 'progress',
                stage: 'tasks',
                message: `Processing tasks... (${tasksCreated + tasksUpdated + taskFailures.length}/${totalTasks})`,
                currentTask: task.eppm_name,
                totalTasks,
                totalAssignments,
                progress: {
                    tasksCreated,
                    tasksUpdated,
                    tasksFailed: taskFailures.length,
                    assignmentsCreated: 0,
                    assignmentsUpdated: 0,
                    assignmentsFailed: 0,
                    dependenciesProcessed: 0
                }
            });
        }

        console.log(`[Import MPP] Tasks processed: ${tasksCreated} created, ${tasksUpdated} updated, ${taskFailures.length} failures`);

        // Send tasks complete progress
        sendProgress({
            type: 'progress',
            stage: 'tasks_complete',
            message: 'Tasks processed. Updating dependencies...',
            totalTasks,
            totalAssignments,
            progress: {
                tasksCreated,
                tasksUpdated,
                tasksFailed: taskFailures.length,
                assignmentsCreated: 0,
                assignmentsUpdated: 0,
                assignmentsFailed: 0,
                dependenciesProcessed: 0
            }
        });

        // Update tasks with predecessor/successor strings
        console.log('[Import MPP] Updating dependencies...');
        let dependenciesProcessed = 0;
        const totalDependencies = bryntumData.dependencies.length;

        for (const task of processedTasks) {
            const dataverseId = importIdToDataverseId.get(task._importId);
            if (!dataverseId) continue;

            const predecessorStr = buildPredecessorStringForTask(task._importId, bryntumData.dependencies, importIdToDataverseId);
            const successorStr = buildSuccessorStringForTask(task._importId, bryntumData.dependencies, importIdToDataverseId);

            if (predecessorStr || successorStr) {
                try {
                    const updatePayload: any = {};
                    if (predecessorStr) updatePayload.eppm_predecessor = predecessorStr;
                    if (successorStr) updatePayload.eppm_successors = successorStr;

                    await dataverseService.updateTask(dataverseId, updatePayload);
                    // Count dependencies for this task (approximate)
                    const predCount = predecessorStr ? predecessorStr.split(';').length : 0;
                    dependenciesProcessed += predCount;
                    console.log(`[Import MPP] Updated dependencies for task ${task.eppm_name}`);
                } catch (error: any) {
                    console.warn(`[Import MPP] Failed to update dependencies for ${task.eppm_name}:`, error?.message);
                }
            }
        }

        // Send dependencies complete progress
        sendProgress({
            type: 'progress',
            stage: 'dependencies_complete',
            message: 'Dependencies processed. Processing assignments...',
            totalTasks,
            totalAssignments,
            progress: {
                tasksCreated,
                tasksUpdated,
                tasksFailed: taskFailures.length,
                assignmentsCreated: 0,
                assignmentsUpdated: 0,
                assignmentsFailed: 0,
                dependenciesProcessed: totalDependencies
            }
        });

        // PASS 2: Process assignments
        // First, fetch fresh task data from Dataverse to get actual task IDs
        console.log('[Import MPP] Fetching tasks from Dataverse to get actual task IDs...');
        sendProgress({
            type: 'progress',
            stage: 'fetching_tasks',
            message: 'Fetching tasks from Dataverse for assignment linking...',
            totalTasks,
            totalAssignments,
            progress: {
                tasksCreated,
                tasksUpdated,
                tasksFailed: taskFailures.length,
                assignmentsCreated: 0,
                assignmentsUpdated: 0,
                assignmentsFailed: 0,
                dependenciesProcessed: totalDependencies
            }
        });

        // Fetch all tasks for this project from Dataverse
        const taskFilterForAssignments = `eppm_projectid eq '${projectId}'`;
        const freshDataverseTasks = await dataverseService.getAllTasks(taskFilterForAssignments);
        console.log(`[Import MPP] Fetched ${freshDataverseTasks?.length || 0} tasks from Dataverse for project ${projectId}`);

        // Build a map of task name -> Dataverse task ID for assignment linking
        const taskNameToDataverseId = new Map<string, string>();
        for (const task of freshDataverseTasks || []) {
            const taskId = task.eppm_projecttaskid;
            const taskName = task.eppm_name;
            if (taskId && taskName) {
                taskNameToDataverseId.set(taskName.toLowerCase().trim(), taskId);
            }
        }
        console.log(`[Import MPP] Built task name mapping with ${taskNameToDataverseId.size} entries`);

        // Process assignments - update existing or create new
        console.log('[Import MPP] Processing assignments...');
        const processedAssignments: any[] = [];
        const assignmentFailures: Array<{ taskName: string; resource: string; error: string; action: string }> = [];
        let assignmentsCreated = 0;
        let assignmentsUpdated = 0;

        for (const assignment of dataverseData.assignments) {
            // Get task name from import data
            const importTask = dataverseData.tasks.find(t => t._importId === assignment.taskImportId);
            const taskName = importTask?.eppm_name || 'Unknown';

            // Look up actual Dataverse task ID by task name
            let taskDataverseId = importIdToDataverseId.get(assignment.taskImportId);

            // If not found in import mapping, try by task name
            if (!taskDataverseId && taskName && taskName !== 'Unknown') {
                taskDataverseId = taskNameToDataverseId.get(taskName.toLowerCase().trim());
                if (taskDataverseId) {
                    console.log(`[Import MPP] Found task ID by name: "${taskName}" -> ${taskDataverseId}`);
                }
            }

            // Validate task ID - must not be blank
            if (!taskDataverseId || !isGuid(taskDataverseId)) {
                const errorMsg = `Task ID is blank or invalid for task "${taskName}" (importId: ${assignment.taskImportId})`;
                console.error(`[Import MPP] FAILURE: ${errorMsg}`);
                assignmentFailures.push({
                    taskName,
                    resource: assignment.resourceEmail,
                    error: errorMsg,
                    action: 'skip'
                });
                continue;
            }

            console.log(`[Import MPP] Using eppm_projecttaskid: ${taskDataverseId} for task "${taskName}"`);

            const existingAssignmentId = assignment._dataverseAssignmentId;
            const isExistingAssignment = existingAssignmentId && isGuid(existingAssignmentId);

            // Build assignment payload for eppm_taskassignmentses:
            // resource name -> eppm_resourceemail, units -> eppm_maxunits
            const assignmentPayload: any = {
                eppm_resourceemail: assignment.resourceEmail,
                [ASSIGNMENT_UNITS_FIELD]: safeParseUnits(assignment.units),
                eppm_projectid: projectId,
                eppm_taskid: taskDataverseId, // Text field - direct value
            };
            if (assignment.startDate) {
                assignmentPayload.eppm_startdate = assignment.startDate;
            }
            if (assignment.finishDate) {
                assignmentPayload.eppm_finishdate = assignment.finishDate;
            }

            console.log(`[Import MPP] Assignment payload for "${taskName}":`, JSON.stringify(assignmentPayload));

            try {
                if (isExistingAssignment) {
                    // UPDATE existing assignment
                    console.log(`[Import MPP] Updating existing assignment: ${taskName} -> ${assignment.resourceEmail} (${existingAssignmentId})`);
                    console.log(`[Import MPP] Update payload:`, JSON.stringify(assignmentPayload));

                    try {
                        await dataverseService.patchRow(TASK_ASSIGNMENTS_TABLE, existingAssignmentId, assignmentPayload);
                        processedAssignments.push({
                            task: taskName,
                            resource: assignment.resourceEmail,
                            units: assignment.units,
                            action: 'updated'
                        });
                        assignmentsUpdated++;
                        console.log(`[Import MPP] Updated assignment: ${taskName} -> ${assignment.resourceEmail}`);
                    } catch (updateError: any) {
                        // If update fails, try to create new
                        const errMsg = updateError?.response?.data?.error?.message || updateError?.message || String(updateError);
                        console.warn(`[Import MPP] Update failed for assignment: ${errMsg}`);

                        try {
                            await dataverseService.createRow(TASK_ASSIGNMENTS_TABLE, assignmentPayload);
                            processedAssignments.push({
                                task: taskName,
                                resource: assignment.resourceEmail,
                                units: assignment.units,
                                action: 'created'
                            });
                            assignmentsCreated++;
                            console.log(`[Import MPP] Created assignment (after failed update): ${taskName} -> ${assignment.resourceEmail}`);
                        } catch (createError: any) {
                            const createErrMsg = createError?.response?.data?.error?.message || createError?.message || String(createError);
                            console.error(`[Import MPP] FAILURE creating assignment: ${createErrMsg}`);
                            assignmentFailures.push({
                                taskName,
                                resource: assignment.resourceEmail,
                                error: createErrMsg,
                                action: 'create'
                            });
                        }
                    }
                } else {
                    // CREATE new assignment
                    console.log(`[Import MPP] Creating new assignment: ${taskName} -> ${assignment.resourceEmail}`);
                    console.log(`[Import MPP] Create payload:`, JSON.stringify(assignmentPayload));

                    try {
                        await dataverseService.createRow(TASK_ASSIGNMENTS_TABLE, assignmentPayload);
                        processedAssignments.push({
                            task: taskName,
                            resource: assignment.resourceEmail,
                            units: assignment.units,
                            action: 'created'
                        });
                        assignmentsCreated++;
                        console.log(`[Import MPP] Created assignment: ${taskName} -> ${assignment.resourceEmail}`);
                    } catch (createError: any) {
                        const errMsg = createError?.response?.data?.error?.message || createError?.message || String(createError);
                        console.error(`[Import MPP] FAILURE creating assignment: ${errMsg}`);
                        assignmentFailures.push({
                            taskName,
                            resource: assignment.resourceEmail,
                            error: errMsg,
                            action: 'create'
                        });
                    }
                }
            } catch (error: any) {
                const errMsg = error?.response?.data?.error?.message || error?.message || String(error);
                console.error(`[Import MPP] Failed to process assignment for ${taskName}: ${errMsg}`);
                assignmentFailures.push({
                    taskName,
                    resource: assignment.resourceEmail,
                    error: errMsg,
                    action: 'process'
                });
            }

            // Send progress update after each assignment
            sendProgress({
                type: 'progress',
                stage: 'assignments',
                message: `Processing assignments... (${assignmentsCreated + assignmentsUpdated + assignmentFailures.length}/${totalAssignments})`,
                currentAssignment: `${taskName} -> ${assignment.resourceEmail}`,
                totalTasks,
                totalAssignments,
                progress: {
                    tasksCreated,
                    tasksUpdated,
                    tasksFailed: taskFailures.length,
                    assignmentsCreated,
                    assignmentsUpdated,
                    assignmentsFailed: assignmentFailures.length,
                    dependenciesProcessed: totalDependencies
                }
            });
        }

        console.log(`[Import MPP] Assignments processed: ${assignmentsCreated} created, ${assignmentsUpdated} updated, ${assignmentFailures.length} failures`);

        // Send final completion message
        sendProgress({
            type: 'complete',
            success: taskFailures.length === 0 && assignmentFailures.length === 0,
            message: `Import completed! Processed ${tasksCreated + tasksUpdated} tasks (${tasksCreated} created, ${tasksUpdated} updated), ${assignmentsCreated + assignmentsUpdated} assignments`,
            summary: {
                tasksCreated,
                tasksUpdated,
                tasksFailed: taskFailures.length,
                assignmentsCreated,
                assignmentsUpdated,
                assignmentsFailed: assignmentFailures.length,
                dependenciesProcessed: totalDependencies
            },
            ...(taskFailures.length > 0 && { taskFailures }),
            ...(assignmentFailures.length > 0 && { assignmentFailures }),
            importIdMapping: Object.fromEntries(importIdToDataverseId)
        });

        console.log('[Import MPP] Import complete');
        res.end();
    } catch (error: any) {
        console.error('[Import MPP] Error importing:', error);
        // Check if headers already sent (streaming started)
        if (res.headersSent) {
            sendProgress({
                type: 'error',
                error: error.message || 'Failed to import MS Project file'
            });
            res.end();
        } else {
            setResponseHeaders(res, req);
            res.status(500).json({
                success: false,
                error: error.message || 'Failed to import MS Project file'
            });
        }
    }
});

/**
 * GET /api/tasks/:id - Get a single task
 */
router.get('/:id', async (req: Request, res: Response) => {
    try {
        const dataverseService = createDataverseService(req);
        const taskId = req.params.id;
        const dataverseTask = await dataverseService.getTaskById(taskId);

        const { buildTaskHierarchy } = await import('../utils/dataTransformer.js');
        const bryntumTask = buildTaskHierarchy([dataverseTask])[0];

        setResponseHeaders(res, req);
        res.json({ success: true, data: bryntumTask });
    } catch (error: any) {
        console.error('Error fetching task:', error);
        setResponseHeaders(res, req);
        res.status(500).json({
            success: false,
            error: error.message || 'Failed to fetch task'
        });
    }
});

/**
 * POST /api/tasks - Create a new task
 */
router.post('/', async (req: Request, res: Response) => {
    try {
        const dataverseService = createDataverseService(req);
        const bryntumTask = req.body;
        const dataverseTask = bryntumToDataverseTask(bryntumTask);

        ensureProjectId(dataverseTask);

        const createdTask = await dataverseService.createTask(dataverseTask);

        const { buildTaskHierarchy } = await import('../utils/dataTransformer.js');
        const bryntumResponse = buildTaskHierarchy([createdTask])[0];

        setResponseHeaders(res, req);
        res.status(201).json({ success: true, data: bryntumResponse });
    } catch (error: any) {
        console.error('Error creating task:', error);
        setResponseHeaders(res, req);
        res.status(500).json({
            success: false,
            error: error.message || 'Failed to create task'
        });
    }
});

/**
 * PUT /api/tasks/:id - Update an existing task
 */
router.put('/:id', async (req: Request, res: Response) => {
    try {
        const dataverseService = createDataverseService(req);
        const taskId = req.params.id;
        const bryntumTask = req.body;
        const dataverseTask = bryntumToDataverseTask(bryntumTask);

        await dataverseService.updateTask(taskId, dataverseTask);

        setResponseHeaders(res, req);
        res.json({ success: true, message: 'Task updated successfully' });
    } catch (error: any) {
        console.error('Error updating task:', error);
        setResponseHeaders(res, req);
        res.status(500).json({
            success: false,
            error: error.message || 'Failed to update task'
        });
    }
});

/**
 * DELETE /api/tasks/:id - Delete a task
 */
router.delete('/:id', async (req: Request, res: Response) => {
    try {
        const dataverseService = createDataverseService(req);
        const taskId = req.params.id;
        await dataverseService.deleteTask(taskId);

        setResponseHeaders(res, req);
        res.json({ success: true, message: 'Task deleted successfully' });
    } catch (error: any) {
        console.error('Error deleting task:', error);
        setResponseHeaders(res, req);
        res.status(500).json({
            success: false,
            error: error.message || 'Failed to delete task'
        });
    }
});

/**
 * POST /api/tasks/sync - Sync multiple tasks (for batch operations)
 */
router.post('/sync', async (req: Request, res: Response) => {
    try {
        const dataverseService = createDataverseService(req);
        /**
         * Bryntum CrudManager can send sync payload in multiple shapes:
         * 1) JSON body: { type:"sync", requestId, tasks:{ updated:[...], added:[...], removed:[...] } }
         * 2) urlencoded body: { data: "<json-string>" }
         * 3) legacy/custom: { tasks: BryntumTask[] }
         *
         * For now we only persist these Dataverse fields:
         * - eppm_name
         * - eppm_startdate
         * - eppm_finishdate
         * - eppm_taskduration
         */
        const payload: any =
            typeof req.body?.data === 'string'
                ? JSON.parse(req.body.data)
                : (req.body || {});

        const requestId = payload?.requestId;

        // Legacy format: { tasks: BryntumTask[] }
        if (Array.isArray(payload?.tasks)) {
            const flatTasks = flattenTaskHierarchy(payload.tasks);
            for (const t of flatTasks) {
                if (!t.eppm_projecttaskid) continue;
                // Only update fields that bryntumToDataverseTask maps (name/start/end/duration/parentId)
                const patch = bryntumToDataverseTask({
                    id: t.eppm_projecttaskid,
                    name: t.eppm_name,
                    startDate: t.eppm_startdate,
                    endDate: t.eppm_finishdate,
                    duration: t.eppm_taskduration,
                    parentId: t.eppm_parenttaskid
                } as any);
                await dataverseService.updateTask(t.eppm_projecttaskid, patch);
            }

            setResponseHeaders(res, req);
            return res.json({ success: true, requestId });
        }

        // Standard Bryntum sync format - support multiple payload shapes
        const taskChanges = payload?.tasks || payload?.changes?.tasks || payload?.stores?.tasks;
        const added: any[] = taskChanges?.added || [];
        const updated: any[] = taskChanges?.updated || [];
        const removed: any[] = taskChanges?.removed || [];
        const dependenciesChanges =
            payload?.dependencies ||
            payload?.stores?.dependencies ||
            payload?.changes?.dependencies;
        const dependenciesAdded: any[] = dependenciesChanges?.added || [];
        const dependenciesUpdated: any[] = dependenciesChanges?.updated || [];
        const dependenciesRemoved: any[] = dependenciesChanges?.removed || [];
        const resourcesChanges = payload?.resources;
        const resourcesUpdated: any[] = resourcesChanges?.updated || [];
        const assignmentsChanges =
            payload?.assignments ||
            payload?.stores?.assignments ||
            payload?.changes?.assignments ||
            payload?.project?.assignments;
        const assignmentsAdded: any[] = assignmentsChanges?.added || [];
        const assignmentsUpdated: any[] = assignmentsChanges?.updated || [];
        const assignmentsRemoved: any[] = assignmentsChanges?.removed || [];

        // Apply updates
        const rows: any[] = [];
        const predecessorPatched = new Set<string>();
        const successorPatched = new Set<string>();
        const taskRemovedIds: string[] = [];
        const taskFailures: Array<{ action: string; id?: string; error?: string }> = [];

        // When creating tasks (copy/paste), Bryntum uses temporary ids ($PhantomId).
        // We must map them to real Dataverse GUIDs and return the mapping in the sync response.
        const taskIdMap = new Map<string, string>();
        const taskProjectIdMap = new Map<string, string>();
        const taskProjectIdCache = new Map<string, string>();
        const resolveTaskId = (raw: any): string => {
            const s = String(raw ?? '').trim();
            if (!s) return '';
            if (isGuid(s)) return s;
            return taskIdMap.get(s) || s;
        };
        const getTaskProjectId = async (taskId: string): Promise<string | undefined> => {
            const id = String(taskId || '').trim();
            if (!isGuid(id)) return undefined;

            if (taskProjectIdCache.has(id)) return taskProjectIdCache.get(id);
            if (taskProjectIdMap.has(id)) return taskProjectIdMap.get(id);

            try {
                const row = await dataverseService.getTaskById(id);
                const pid = extractProjectIdFromTaskRow(row);
                if (pid) {
                    taskProjectIdCache.set(id, pid);
                    return pid;
                }
            } catch (e: any) {
                // ignore
            }
            return undefined;
        };

        // Handle task deletions (single or multiple)
        // Bryntum usually sends tasks.removed as array of ids or objects containing { id }
        const removedTaskIds = (Array.isArray(removed) ? removed : [])
            .map((r: any) => typeof r === 'string' ? r : (r?.id ?? r?.$PhantomId))
            .map((v: any) => String(v || ''))
            .filter((id: string) => isGuid(id));

        if (removedTaskIds.length) {
            const removedSet = new Set<string>(removedTaskIds);

            // Best-effort: delete related task assignment rows first (avoid delete failures due to relationships)
            try {
                const assignmentRows = await dataverseService.getTableRows<any>(TASK_ASSIGNMENTS_TABLE);
                const toDeleteAssignmentIds = new Set<string>();

                for (const row of assignmentRows) {
                    const taskId = extractTaskIdFromAssignment(row);
                    if (!taskId || !removedSet.has(taskId)) continue;

                    const assignmentId = extractAssignmentId(row);
                    if (assignmentId) toDeleteAssignmentIds.add(assignmentId);
                }

                for (const assignmentId of toDeleteAssignmentIds) {
                    try {
                        await dataverseService.deleteRow(TASK_ASSIGNMENTS_TABLE, assignmentId);
                    } catch (e: any) {
                        console.warn('[Tasks Route] Failed to delete related assignment for task removal:', e?.message || e);
                    }
                }
            } catch (e: any) {
                console.warn('[Tasks Route] Failed to fetch/delete assignments during task removal:', e?.message || e);
            }

            for (const id of removedTaskIds) {
                try {
                    await dataverseService.deleteTask(id);
                    taskRemovedIds.push(id);
                } catch (e: any) {
                    const msg = e?.response?.data?.error?.message || e?.message || String(e || 'Unknown error');
                    taskFailures.push({ action: 'remove', id, error: msg });
                }
            }
        }

        // Handle task creations (copy/paste/cut+paste) - create new rows in eppm_projecttasks
        const addedTasks: any[] = Array.isArray(added) ? added : [];
        if (addedTasks.length) {
            const pending = [...addedTasks];

            const getTempId = (t: any) => String(t?.$PhantomId ?? t?.id ?? '').trim();
            const getParentRaw = (t: any) => t?.parentId ?? t?.parent ?? t?.parent_id ?? t?.parentID;

            let safety = pending.length * 3 + 10;
            while (pending.length && safety-- > 0) {
                let progressed = false;

                for (let i = 0; i < pending.length; i++) {
                    const t = pending[i];
                    const phantomId = getTempId(t);
                    const parentRaw = getParentRaw(t);
                    const parentIdResolved = resolveTaskId(parentRaw);

                    const canCreate =
                        !parentRaw ||
                        isGuid(String(parentRaw)) ||
                        (parentIdResolved && isGuid(parentIdResolved));

                    if (!canCreate) continue;

                    // Build create payload. Do NOT send the task id to Dataverse.
                    const createPayload: Record<string, any> = {
                        ...(bryntumToDataverseTask(t) as any)
                    };

                    delete (createPayload as any).eppm_projecttaskid;

                    // Parent mapping: if parent is known and GUID, persist; else omit
                    if (parentIdResolved && isGuid(parentIdResolved)) {
                        (createPayload as any).eppm_parenttaskid = parentIdResolved;
                    } else {
                        delete (createPayload as any).eppm_parenttaskid;
                    }

                    // Cross-project paste support:
                    // If pasting under an existing task (or an already-created parent), inherit that parent's project id.
                    if (parentIdResolved && isGuid(parentIdResolved)) {
                        const parentProjectId =
                            taskProjectIdMap.get(parentIdResolved) ||
                            (await getTaskProjectId(parentIdResolved));
                        if (parentProjectId) {
                            (createPayload as any).eppm_projectid = parentProjectId;
                        }
                    }

                    // Ensure project ID is always set (required for add above/below, paste, etc.)
                    ensureProjectId(createPayload);

                    // Add default dates for new tasks if missing (Dataverse often requires them; add above/below may omit)
                    const todayIso = new Date().toISOString().split('T')[0] + 'T12:00:00.000Z';
                    if (!(createPayload as any).eppm_startdate) {
                        (createPayload as any).eppm_startdate = todayIso;
                    }
                    if (!(createPayload as any).eppm_finishdate) {
                        (createPayload as any).eppm_finishdate = todayIso;
                    }
                    if ((createPayload as any).eppm_taskduration === undefined || (createPayload as any).eppm_taskduration === null) {
                        (createPayload as any).eppm_taskduration = 1;
                    }

                    // Avoid persisting predecessor/successor strings from pasted task data directly
                    // (they might reference old ids). Dependencies will be handled separately.
                    delete (createPayload as any).eppm_predecessor;
                    delete (createPayload as any).eppm_successors;

                    // Ensure there is at least a name
                    if (!(createPayload as any).eppm_name) {
                        (createPayload as any).eppm_name = typeof t?.name === 'string' && t.name.trim() ? t.name.trim() : 'New task';
                    }

                    try {
                        const created = await dataverseService.createRow<any>(TASKS_TABLE, createPayload);
                        const createdId = extractTaskId(created);

                        if (!createdId) {
                            taskFailures.push({ action: 'create', id: phantomId || undefined, error: 'Created task row but could not extract eppm_projecttaskid from Dataverse response' });
                            pending.splice(i, 1);
                            i--;
                            progressed = true;
                            continue;
                        }

                        if (phantomId) taskIdMap.set(phantomId, createdId);

                        const createdProjectId =
                            (createPayload as any).eppm_projectid
                                ? String((createPayload as any).eppm_projectid)
                                : extractProjectIdFromTaskRow(created);
                        if (createdProjectId) {
                            taskProjectIdMap.set(createdId, createdProjectId);
                            if (phantomId) taskProjectIdMap.set(phantomId, createdProjectId);
                        }

                        rows.push({
                            ...(phantomId ? { $PhantomId: phantomId } : {}),
                            id: createdId,
                            ...(parentIdResolved && isGuid(parentIdResolved) ? { parentId: parentIdResolved } : {})
                        });

                        pending.splice(i, 1);
                        i--;
                        progressed = true;
                    } catch (e: any) {
                        const msg = e?.response?.data?.error?.message || e?.message || String(e || 'Unknown error');
                        taskFailures.push({ action: 'create', id: phantomId || undefined, error: msg });
                        pending.splice(i, 1);
                        i--;
                        progressed = true;
                    }
                }

                if (!progressed) break;
            }

            // If something remains (likely parent references we couldn't resolve), create them as root tasks (best effort)
            for (const t of pending) {
                const phantomId = getTempId(t);
                const createPayload: Record<string, any> = { ...(bryntumToDataverseTask(t) as any) };
                delete (createPayload as any).eppm_projecttaskid;
                delete (createPayload as any).eppm_parenttaskid;
                delete (createPayload as any).eppm_predecessor;
                delete (createPayload as any).eppm_successors;

                if (!(createPayload as any).eppm_name) {
                    const n = (t as any)?.name ?? (t as any)?.taskName ?? (t as any)?.text;
                    (createPayload as any).eppm_name = typeof n === 'string' && n.trim() ? n.trim() : 'New task';
                }

                ensureProjectId(createPayload);

                // Add default dates for new tasks if missing (add above/below often omit)
                const todayIso = new Date().toISOString().split('T')[0] + 'T12:00:00.000Z';
                if (!(createPayload as any).eppm_startdate) {
                    (createPayload as any).eppm_startdate = todayIso;
                }
                if (!(createPayload as any).eppm_finishdate) {
                    (createPayload as any).eppm_finishdate = todayIso;
                }
                if ((createPayload as any).eppm_taskduration === undefined || (createPayload as any).eppm_taskduration === null) {
                    (createPayload as any).eppm_taskduration = 1;
                }

                try {
                    const created = await dataverseService.createRow<any>(TASKS_TABLE, createPayload);
                    const createdId = extractTaskId(created);

                    if (createdId) {
                        if (phantomId) taskIdMap.set(phantomId, createdId);

                        const createdProjectId =
                            (createPayload as any).eppm_projectid
                                ? String((createPayload as any).eppm_projectid)
                                : extractProjectIdFromTaskRow(created);
                        if (createdProjectId) {
                            taskProjectIdMap.set(createdId, createdProjectId);
                            if (phantomId) taskProjectIdMap.set(phantomId, createdProjectId);
                        }
                        rows.push({
                            ...(phantomId ? { $PhantomId: phantomId } : {}),
                            id: createdId
                        });
                    } else {
                        taskFailures.push({ action: 'create', id: phantomId || undefined, error: 'Created task row but could not extract eppm_projecttaskid from Dataverse response' });
                    }
                } catch (e: any) {
                    const msg = e?.response?.data?.error?.message || e?.message || String(e || 'Unknown error');
                    taskFailures.push({ action: 'create', id: phantomId || undefined, error: msg });
                }
            }
        }

        for (const t of updated) {
            const id = resolveTaskId(t.id || t.$PhantomId || '');
            if (!id || !isGuid(id)) continue;

            // Normalize task: Bryntum may send flat object or nested (record.data)
            const taskForPatch = (t && typeof t === 'object' && t.data) ? { ...t, ...t.data } : t;

            // This will include: eppm_name, eppm_startdate, eppm_finishdate, eppm_taskduration, etc.
            const patch = bryntumToDataverseTask(taskForPatch);

            // Cross-project cut/paste (move):
            // If record is moved under a different parent, inherit that parent's project id unless payload explicitly set it.
            if (!(patch as any).eppm_projectid) {
                const parentRaw = (taskForPatch as any).parentId ?? (taskForPatch as any).parent ?? (taskForPatch as any).parent_id ?? (taskForPatch as any).parentID;
                const parentIdResolved = resolveTaskId(parentRaw);
                if (parentIdResolved && isGuid(parentIdResolved)) {
                    const parentProjectId =
                        taskProjectIdMap.get(parentIdResolved) ||
                        (await getTaskProjectId(parentIdResolved));
                    if (parentProjectId) {
                        (patch as any).eppm_projectid = parentProjectId;
                    }
                }
            }

            // Store predecessor string in eppm_predecessor.
            // Bryntum often includes `predecessors` array on task update when edited via Task Editor -> Predecessors tab.
            if (Array.isArray(taskForPatch.predecessors)) {
                const predStr = buildPredecessorStringFromArray(taskForPatch.predecessors);
                (patch as any).eppm_predecessor = predStr;
                predecessorPatched.add(id);
            } else if (typeof taskForPatch.predecessor === 'string') {
                (patch as any).eppm_predecessor = normalizePredecessorString(taskForPatch.predecessor);
                predecessorPatched.add(id);
            }

            // Store successor string in eppm_successors.
            // Bryntum may include `successors` array on task update when edited via Task Editor -> Successors tab.
            if (Array.isArray((taskForPatch as any).successors)) {
                const succStr = buildSuccessorStringFromArray((taskForPatch as any).successors);
                (patch as any).eppm_successors = succStr || null;
                successorPatched.add(id);
            } else if (typeof (taskForPatch as any).successor === 'string') {
                (patch as any).eppm_successors = normalizeSuccessorString((taskForPatch as any).successor) || null;
                successorPatched.add(id);
            } else if (typeof (taskForPatch as any).successors === 'string') {
                (patch as any).eppm_successors = normalizeSuccessorString((taskForPatch as any).successors) || null;
                successorPatched.add(id);
            }

            // Filter out undefined values - only send defined fields to Dataverse
            const cleanPatch: Record<string, any> = {};
            for (const [k, v] of Object.entries(patch)) {
                if (v !== undefined) cleanPatch[k] = v;
            }
            if (Object.keys(cleanPatch).length > 0) {
                await dataverseService.updateTask(id, cleanPatch);
            }
            // Only push valid task rows with a valid id
            if (id && isGuid(id)) {
                rows.push({ id });
            }
        }

        // If dependencies changed but the task update did not include `predecessors`,
        // update the stored eppm_predecessor string incrementally.
        const affected = new Map<string, { add: Set<string>; remove: Set<string> }>();
        const upsert = (toTask: string, token: string, isRemove: boolean) => {
            if (!affected.has(toTask)) {
                affected.set(toTask, { add: new Set(), remove: new Set() });
            }
            const entry = affected.get(toTask)!;
            if (isRemove) entry.remove.add(token);
            else entry.add.add(token);
        };

        // Mirror dependencies into successors on the FROM task
        const affectedSucc = new Map<string, { add: Set<string>; remove: Set<string> }>();
        const upsertSucc = (fromTask: string, token: string, isRemove: boolean) => {
            if (!affectedSucc.has(fromTask)) {
                affectedSucc.set(fromTask, { add: new Set(), remove: new Set() });
            }
            const entry = affectedSucc.get(fromTask)!;
            if (isRemove) entry.remove.add(token);
            else entry.add.add(token);
        };

        const resolveDep = (dep: any) => {
            const fromRaw = dep?.fromTask ?? dep?.from ?? dep?.fromEvent ?? dep?.fromTaskId ?? dep?.fromId;
            const toRaw = dep?.toTask ?? dep?.to ?? dep?.toEvent ?? dep?.toTaskId ?? dep?.toId;
            const fromTask = resolveTaskId(fromRaw);
            const toTask = resolveTaskId(toRaw);

            return {
                ...dep,
                fromTask,
                toTask,
                from: fromTask,
                to: toTask
            };
        };

        for (const dep of [...dependenciesAdded, ...dependenciesUpdated]) {
            const depResolved = resolveDep(dep);
            const toTask = depResolved?.toTask ?? depResolved?.to;
            const token = dependencyToToken(depResolved);
            if (!toTask || !isGuid(String(toTask)) || !token) continue;
            upsert(String(toTask), token, false);

            const fromTask = depResolved?.fromTask ?? depResolved?.from;
            const succToken = successorToToken(depResolved);
            if (fromTask && isGuid(String(fromTask)) && succToken) upsertSucc(String(fromTask), succToken, false);
        }
        for (const dep of dependenciesRemoved) {
            // Removed may be just an id; only process if it includes from/to
            const depResolved = resolveDep(dep);
            const toTask = depResolved?.toTask ?? depResolved?.to;
            const token = dependencyToToken(depResolved);
            if (!toTask || !isGuid(String(toTask)) || !token) continue;
            upsert(String(toTask), token, true);

            const fromTask = depResolved?.fromTask ?? depResolved?.from;
            const succToken = successorToToken(depResolved);
            if (fromTask && isGuid(String(fromTask)) && succToken) upsertSucc(String(fromTask), succToken, true);
        }

        for (const [toTaskId, { add, remove }] of affected.entries()) {
            if (predecessorPatched.has(toTaskId)) continue;
            if (add.size === 0 && remove.size === 0) continue;

            try {
                const current = await dataverseService.getTaskById(toTaskId);
                const currentStr = (current as any).eppm_predecessor || '';
                const tokens = new Set(normalizePredecessorString(currentStr).split(';').filter(Boolean));

                for (const t of add) tokens.add(t);
                for (const t of remove) tokens.delete(t);

                const nextStr = Array.from(tokens).join(';');
                await dataverseService.updateTask(toTaskId, { eppm_predecessor: nextStr || null } as any);
            } catch (e: any) {
                console.warn('[Tasks Route] Failed to update predecessor string from dependencies:', e?.message || e);
            }
        }

        for (const [fromTaskId, { add, remove }] of affectedSucc.entries()) {
            if (successorPatched.has(fromTaskId)) continue;
            if (add.size === 0 && remove.size === 0) continue;

            try {
                const current = await dataverseService.getTaskById(fromTaskId);
                const currentStr = (current as any).eppm_successors || '';
                const tokens = new Set(normalizeSuccessorString(currentStr).split(';').filter(Boolean));

                for (const t of add) tokens.add(t);
                for (const t of remove) tokens.delete(t);

                const nextStr = Array.from(tokens).join(';');
                await dataverseService.updateTask(fromTaskId, { eppm_successors: nextStr || null } as any);
            } catch (e: any) {
                console.warn('[Tasks Route] Failed to update successors string from dependencies:', e?.message || e);
            }
        }

        // If user edits a resource name, update eppm_resourceemail (stores resource name) in eppm_taskassignmentses
        for (const r of resourcesUpdated) {
            const oldName = typeof r.id === 'string' ? r.id.trim() : '';
            const newNameRaw = r.name ?? r.email ?? r.resourceemail;
            const newName = typeof newNameRaw === 'string' ? newNameRaw.trim() : '';

            if (!oldName || !newName || oldName === newName) continue;

            try {
                const filter = `$filter=eppm_resourceemail eq '${oldName.replace(/'/g, "''")}'`;
                const assignmentRows = await dataverseService.getTableRows<any>(TASK_ASSIGNMENTS_TABLE, filter);

                for (const row of assignmentRows) {
                    const rowId = extractAssignmentId(row);
                    if (!rowId) continue;
                    await dataverseService.patchRow(TASK_ASSIGNMENTS_TABLE, rowId, { eppm_resourceemail: newName });
                }
            } catch (e: any) {
                console.warn('[Tasks Route] Failed to update resource name in assignments:', e?.message || e);
            }
        }

        // Persist task->resource assignments into eppm_taskassignmentses
        // Each task can have multiple assignments (multiple rows).
        const assignmentResponseRows: any[] = [];
        const assignmentRemovedIds: string[] = [];
        const assignmentFailures: Array<{ action: string; id?: string; phantomId?: string; taskId?: string; resourceEmail?: string; error?: string; }> = [];
        const affectedTaskIdsForResources = new Set<string>();

        // Build resource id -> name map from sync payload (eppm_resourceemail stores resource NAME)
        const resourceIdToName = new Map<string, string>();
        const resourcesRowsFromPayload = payload?.resources?.rows || payload?.stores?.resources?.rows || [];
        for (const r of resourcesRowsFromPayload) {
            const id = (r.id ?? r.email ?? '').toString().trim();
            const name = (r.name ?? r.email ?? id).toString().trim();
            if (id) {
                resourceIdToName.set(id, name);
                resourceIdToName.set(id.toLowerCase(), name);
            }
        }

        // Create assignments (store resource NAME in eppm_resourceemail, units in eppm_maxunits)
        for (const a of assignmentsAdded) {
            const phantomId = a.$PhantomId || a.id;
            const taskId = resolveTaskId(a.event ?? a.taskId ?? a.task ?? a.eventId ?? '');
            const resourceIdRaw = a.resource ?? a.resourceId ?? a.resource;
            const resourceId = typeof resourceIdRaw === 'string'
                ? resourceIdRaw.trim()
                : (typeof resourceIdRaw?.id === 'string' ? resourceIdRaw.id.trim() : '');
            const resourceName = resourceIdToName.get(resourceId) || resourceIdToName.get(resourceId.toLowerCase()) || (a.resourceName ?? a.name) || resourceId;
            // Use safeParseUnits to ensure valid number
            const units = safeParseUnits(a.units);

            if (!taskId || !resourceName) continue;

            const basePayload: Record<string, any> = {
                eppm_resourceemail: resourceName,
                [ASSIGNMENT_UNITS_FIELD]: Number.isFinite(units) ? units : 100
            };

            try {
                let created: any | null = null;

                if (isGuid(taskId)) {
                    const candidates = getAssignmentTaskLookupCandidates();
                    let lastErr: any = null;

                    // 1) Try @odata.bind (standard Dataverse format)
                    for (const lookupName of candidates) {
                        const createPayload: Record<string, any> = {
                            ...basePayload,
                            [`${lookupName}@odata.bind`]: `/${TASKS_TABLE}(${taskId})`
                        };
                        try {
                            created = await dataverseService.createRow<any>(TASK_ASSIGNMENTS_TABLE, createPayload);
                            break;
                        } catch (e: any) {
                            lastErr = e;
                        }
                    }

                    // 2) If @odata.bind fails with ODataEntityReferenceLink error, try _value format (raw GUID)
                    const errMsg = String(lastErr?.response?.data?.error?.message || lastErr?.message || '');
                    if (!created && errMsg.includes('ODataEntityReferenceLink')) {
                        for (const lookupName of candidates) {
                            const createPayload: Record<string, any> = {
                                ...basePayload,
                                ...buildTaskLookupAsValue(lookupName, taskId)
                            };
                            try {
                                created = await dataverseService.createRow<any>(TASK_ASSIGNMENTS_TABLE, createPayload);
                                break;
                            } catch (e: any) {
                                lastErr = e;
                            }
                        }
                    }

                    // 3) Last resort: create without task link (assignment may still be valid if task is optional)
                    if (!created) {
                        try {
                            created = await dataverseService.createRow<any>(TASK_ASSIGNMENTS_TABLE, basePayload);
                            if (created) {
                                // PATCH to set task link using _value format
                                const createdId = extractAssignmentId(created);
                                if (createdId) {
                                    for (const lookupName of candidates) {
                                        try {
                                            await dataverseService.patchRow(TASK_ASSIGNMENTS_TABLE, createdId, buildTaskLookupAsValue(lookupName, taskId));
                                            break;
                                        } catch {
                                            /* try next */
                                        }
                                    }
                                }
                            }
                        } catch (e: any) {
                            lastErr = e;
                        }
                    }

                    if (!created) {
                        const msg = lastErr?.response?.data?.error?.message || lastErr?.message || String(lastErr || 'Unknown error');
                        assignmentFailures.push({ action: 'create', phantomId: String(phantomId || ''), taskId, resourceEmail: resourceName, error: msg });
                        continue;
                    }
                } else {
                    // If your environment uses a non-GUID task key, you may need a different column.
                    created = await dataverseService.createRow<any>(TASK_ASSIGNMENTS_TABLE, basePayload);
                }

                const createdId = extractAssignmentId(created);

                if (createdId) {
                    if (isGuid(taskId)) affectedTaskIdsForResources.add(taskId);
                    // Always include units with a valid numeric value
                    assignmentResponseRows.push({
                        ...(phantomId ? { $PhantomId: phantomId } : {}),
                        id: createdId,
                        event: taskId,
                        resource: resourceName,
                        units: 100  // Always provide a valid numeric value
                    });
                } else {
                    assignmentFailures.push({ action: 'create', phantomId: String(phantomId || ''), taskId, resourceEmail: resourceName, error: 'Created assignment row but could not extract id from Dataverse response' });
                }
            } catch (e: any) {
                const msg = e?.response?.data?.error?.message || e?.message || String(e || 'Unknown error');
                assignmentFailures.push({ action: 'create', phantomId: String(phantomId || ''), taskId, resourceEmail: resourceName, error: msg });
            }
        }

        // Update assignments (store resource NAME in eppm_resourceemail, units in eppm_maxunits)
        for (const a of assignmentsUpdated) {
            const assignmentId = String(a.id || '');
            if (!isGuid(assignmentId)) continue;

            const taskId = resolveTaskId(a.event ?? a.taskId ?? a.task ?? a.eventId ?? '');
            const resourceIdRaw = a.resource ?? a.resourceId ?? a.resource;
            const resourceId = typeof resourceIdRaw === 'string' ? resourceIdRaw.trim() : (typeof resourceIdRaw?.id === 'string' ? resourceIdRaw.id.trim() : '');
            const resourceName = resourceIdToName.get(resourceId) || resourceIdToName.get(resourceId?.toLowerCase?.() || '') || (a.resourceName ?? a.name) || resourceId;
            // Use safeParseUnits to ensure valid number
            const units = safeParseUnits(a.units);

            const patch: Record<string, any> = {};
            if (resourceName) patch.eppm_resourceemail = resourceName;
            if (units !== undefined && Number.isFinite(units)) patch[ASSIGNMENT_UNITS_FIELD] = units;

            // Task bind: try candidate lookup navigation properties (see create logic above)
            if (isGuid(taskId)) {
                const candidates = getAssignmentTaskLookupCandidates();

                // We'll attempt patching with each candidate until one succeeds if we need to bind.
                // If none succeeds, we record a failure (but still allow other updates to proceed).
                let patched = false;
                let lastErr: any = null;

                for (const lookupName of candidates) {
                    const tryPatch = {
                        ...patch,
                        [`${lookupName}@odata.bind`]: `/${TASKS_TABLE}(${taskId})`
                    };

                    try {
                        await dataverseService.patchRow(TASK_ASSIGNMENTS_TABLE, assignmentId, tryPatch);
                        if (isGuid(taskId)) affectedTaskIdsForResources.add(taskId);
                        // Always include units with a valid numeric value
                        assignmentResponseRows.push({
                            id: assignmentId,
                            ...(taskId ? { event: taskId } : {}),
                            ...(resourceName ? { resource: resourceName } : {}),
                            units: 100  // Always provide a valid numeric value
                        });
                        patched = true;
                        break;
                    } catch (e: any) {
                        lastErr = e;
                    }
                }

                if (!patched) {
                    const msg = lastErr?.response?.data?.error?.message || lastErr?.message || String(lastErr || 'Unknown error');
                    assignmentFailures.push({ action: 'update', id: assignmentId, taskId, resourceEmail: resourceName, error: msg });
                }

                continue;
            }

            if (Object.keys(patch).length === 0) continue;

            try {
                await dataverseService.patchRow(TASK_ASSIGNMENTS_TABLE, assignmentId, patch);
                if (isGuid(taskId)) affectedTaskIdsForResources.add(taskId);
                // Always include units with a valid numeric value
                assignmentResponseRows.push({
                    id: assignmentId,
                    ...(taskId ? { event: taskId } : {}),
                    ...(resourceName ? { resource: resourceName } : {}),
                    units: 100  // Always provide a valid numeric value
                });
            } catch (e: any) {
                const msg = e?.response?.data?.error?.message || e?.message || String(e || 'Unknown error');
                assignmentFailures.push({ action: 'update', id: assignmentId, taskId, resourceEmail: resourceName, error: msg });
            }
        }

        // Delete assignments (fetch taskId before delete for eppm_resources updates)
        for (const a of assignmentsRemoved) {
            const assignmentId = typeof a === 'string' ? a : (a?.id || a);
            const id = String(assignmentId || '');
            if (!isGuid(id)) continue;

            try {
                const assignmentRow = await dataverseService.getRowById<any>(TASK_ASSIGNMENTS_TABLE, id);
                const taskIdForRemoved = assignmentRow ? extractTaskIdFromAssignment(assignmentRow) : undefined;
                if (taskIdForRemoved && isGuid(taskIdForRemoved)) {
                    affectedTaskIdsForResources.add(taskIdForRemoved);
                }
                await dataverseService.deleteRow(TASK_ASSIGNMENTS_TABLE, id);
                assignmentRemovedIds.push(id);
            } catch (e: any) {
                const msg = e?.response?.data?.error?.message || e?.message || String(e || 'Unknown error');
                assignmentFailures.push({ action: 'remove', id, error: msg });
            }
        }

        // Persist resource name and units to eppm_resources on eppm_projecttasks
        if (TASK_RESOURCES_FIELD && affectedTaskIdsForResources.size > 0) {
            try {
                const assignmentRows = await dataverseService.getTableRows<any>(TASK_ASSIGNMENTS_TABLE);
                const assignmentsByTask = new Map<string, Array<{ name: string; units: number }>>();
                for (const row of assignmentRows) {
                    const taskId = extractTaskIdFromAssignment(row);
                    if (!taskId || !isGuid(taskId)) continue;
                    const resourceName = row.eppm_resourceemail;
                    if (typeof resourceName !== 'string' || !resourceName.trim()) continue;
                    const units = safeParseUnits(row[ASSIGNMENT_UNITS_FIELD]);
                    const list = assignmentsByTask.get(taskId) || [];
                    list.push({ name: resourceName.trim(), units });
                    assignmentsByTask.set(taskId, list);
                }
                for (const taskId of affectedTaskIdsForResources) {
                    const assignments = assignmentsByTask.get(taskId) || [];
                    const summary = assignments.map(({ name, units }) => ({ name, units }));
                    try {
                        await dataverseService.updateTask(taskId, { [TASK_RESOURCES_FIELD]: JSON.stringify(summary) } as any);
                    } catch (patchErr: any) {
                        console.warn(`[Tasks Route] Failed to update eppm_resources for task ${taskId}:`, patchErr?.message || patchErr);
                    }
                }
            } catch (e: any) {
                console.warn('[Tasks Route] Failed to update task eppm_resources:', e?.message || e);
            }
        }

        // Respond in a Bryntum-friendly format
        // Filter out any undefined/null values from response arrays to prevent "Cannot read properties of undefined" errors
        const validTaskRows = rows.filter((r: any) => r && r.id);
        const validTaskRemovedIds = taskRemovedIds.filter((id: string) => id && isGuid(id));
        const validAssignmentRows = assignmentResponseRows.filter((r: any) => r && r.id);
        const validAssignmentRemovedIds = assignmentRemovedIds.filter((id: string) => id && isGuid(id));

        // Build response object
        const syncResponse: any = {
            success: assignmentFailures.length === 0 && taskFailures.length === 0,
            requestId,
            tasks: {
                rows: validTaskRows
            }
        };

        // Only include removed tasks if there are valid removed IDs
        if (validTaskRemovedIds.length > 0) {
            syncResponse.tasks.removed = validTaskRemovedIds;
        }

        // Only include error details if there are failures
        if (assignmentFailures.length || taskFailures.length) {
            syncResponse.error = 'One or more changes failed to persist in Dataverse';
            syncResponse.details = [
                ...(assignmentFailures.length ? assignmentFailures.map(f => ({ type: 'assignment', ...f })) : []),
                ...(taskFailures.length ? taskFailures.map(f => ({ type: 'task', ...f })) : [])
            ];
        }

        // Only include assignments section if there are actual assignment rows or removals
        if (validAssignmentRows.length > 0 || validAssignmentRemovedIds.length > 0) {
            syncResponse.assignments = {
                rows: validAssignmentRows
            };
            if (validAssignmentRemovedIds.length > 0) {
                syncResponse.assignments.removed = validAssignmentRemovedIds;
            }
        }

        setResponseHeaders(res, req);
        res.json(syncResponse);
    } catch (error: any) {
        console.error('Error syncing tasks:', error);
        setResponseHeaders(res, req);
        res.status(500).json({
            success: false,
            error: error.message || 'Failed to sync tasks'
        });
    }
});

export default router;
