/**
 * MSPDI (Microsoft Project Data Interchange) XML Importer
 *
 * Parses XML files in Microsoft Project's MSPDI format and converts
 * them to Bryntum/Dataverse compatible format.
 *
 * Supports:
 * - Tasks with hierarchy (parent-child relationships)
 * - Resources
 * - Assignments (task-resource mappings)
 * - Dependencies (predecessor links)
 */

import { parseString } from 'xml2js';

export interface ImportedTask {
    uid: number;
    id: number;
    name: string;
    startDate?: string;
    finishDate?: string;
    duration?: number; // in days
    percentComplete?: number;
    effort?: number; // in hours
    outlineLevel: number;
    parentUid?: number;
    notes?: string;
    wbs?: string;
    isSummary?: boolean;
    isMilestone?: boolean;
    // Dataverse ID for round-trip support (stored in ExtendedAttribute Text1)
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
    units?: number; // decimal (1 = 100%)
    startDate?: string;
    finishDate?: string;
    // Dataverse ID for round-trip support (stored in ExtendedAttribute Text2)
    dataverseAssignmentId?: string;
}

export interface ImportedDependency {
    fromTaskUid: number;
    toTaskUid: number;
    type: number; // 0=FF, 1=FS, 2=SF, 3=SS in MSPDI
    lag?: number; // in days
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
    tasks: any[];
    resources: any[];
    assignments: any[];
    dependencies: any[];
}

/**
 * Parse MSPDI duration format (PT8H0M0S) to days
 * Assumes 8 hours per day
 */
function parseMspdiDuration(duration: string | undefined): number | undefined {
    if (!duration) return undefined;

    // Format: PT{hours}H{minutes}M{seconds}S
    const match = duration.match(/PT(\d+)H(\d+)M(\d+)S/);
    if (!match) return undefined;

    const hours = parseInt(match[1], 10) || 0;
    const minutes = parseInt(match[2], 10) || 0;

    // Convert to days (8 hours per day)
    return (hours + minutes / 60) / 8;
}

/**
 * Parse MSPDI work format to hours
 */
function parseMspdiWork(work: string | undefined): number | undefined {
    if (!work) return undefined;

    const match = work.match(/PT(\d+)H(\d+)M(\d+)S/);
    if (!match) return undefined;

    const hours = parseInt(match[1], 10) || 0;
    const minutes = parseInt(match[2], 10) || 0;

    return hours + minutes / 60;
}

/**
 * Parse MSPDI date format to ISO date string (YYYY-MM-DD).
 * Preserves the exact date from the XML without timezone conversion.
 * MS Project XML uses ISO 8601 format: "2024-02-15T08:00:00"
 */
function parseMspdiDate(dateStr: string | undefined): string | undefined {
    if (!dateStr) return undefined;

    try {
        // Extract date portion directly from the string to avoid timezone issues
        // MSPDI format: "2024-02-15T08:00:00" or "2024-02-15"
        const trimmed = dateStr.trim();

        // Try to extract YYYY-MM-DD from the start of the string
        const dateMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
        if (dateMatch) {
            return dateMatch[1]; // Return just the date part
        }

        // Fallback: try parsing but avoid timezone conversion by extracting components
        // This handles formats like "02/15/2024" if present
        const date = new Date(trimmed);
        if (isNaN(date.getTime())) return undefined;

        // Use UTC methods to avoid timezone shift
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        const day = String(date.getUTCDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    } catch {
        return undefined;
    }
}

/**
 * Safely convert a date value to ISO string for Dataverse.
 * Preserves the exact date without timezone conversion.
 * Returns format: "YYYY-MM-DDT12:00:00Z" (noon UTC to avoid day boundary issues).
 * Accepts: string, Date, number (timestamp), null, undefined.
 */
function toSafeISOString(value: string | Date | number | null | undefined): string | undefined {
    if (value == null || value === '') return undefined;
    try {
        // If it's already a date string in YYYY-MM-DD format, preserve it exactly
        if (typeof value === 'string') {
            const trimmed = value.trim();

            // Check for YYYY-MM-DD format (already parsed by parseMspdiDate)
            const dateOnlyMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})$/);
            if (dateOnlyMatch) {
                // Return as noon UTC to avoid day boundary issues
                return `${dateOnlyMatch[1]}T12:00:00Z`;
            }

            // Check for YYYY-MM-DDTHH:MM:SS format
            const dateTimeMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})T/);
            if (dateTimeMatch) {
                // Return as noon UTC to avoid day boundary issues
                return `${dateTimeMatch[1]}T12:00:00Z`;
            }
        }

        // Fallback for Date objects or other formats
        const date = value instanceof Date ? value : new Date(value as string | number);
        if (isNaN(date.getTime())) return undefined;

        // Extract date components and format as noon UTC to avoid day boundary issues
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        const day = String(date.getUTCDate()).padStart(2, '0');
        return `${year}-${month}-${day}T12:00:00Z`;
    } catch {
        return undefined;
    }
}

/**
 * Get text value from XML element (handles both direct value and array)
 */
function getTextValue(element: any): string | undefined {
    if (element === undefined || element === null) return undefined;
    if (typeof element === 'string') return element;
    if (Array.isArray(element)) {
        const first = element[0];
        if (typeof first === 'string') return first;
        if (first && typeof first === 'object' && '_' in first) return first._;
        return first?.toString();
    }
    if (typeof element === 'object' && '_' in element) return element._;
    return element.toString();
}

/**
 * Get numeric value from XML element
 */
function getNumericValue(element: any): number | undefined {
    const text = getTextValue(element);
    if (text === undefined) return undefined;
    const num = parseFloat(text);
    return isNaN(num) ? undefined : num;
}

/**
 * Get integer value from XML element
 */
function getIntValue(element: any): number | undefined {
    const text = getTextValue(element);
    if (text === undefined) return undefined;
    const num = parseInt(text, 10);
    return isNaN(num) ? undefined : num;
}

/**
 * Extract ExtendedAttribute value by FieldID
 * Used to retrieve Dataverse IDs stored in custom fields
 */
function getExtendedAttributeValue(element: any, fieldId: number): string | undefined {
    if (!element?.ExtendedAttribute) return undefined;

    const extAttrs = Array.isArray(element.ExtendedAttribute)
        ? element.ExtendedAttribute
        : [element.ExtendedAttribute];

    for (const attr of extAttrs) {
        const attrFieldId = getIntValue(attr.FieldID);
        if (attrFieldId === fieldId) {
            return getTextValue(attr.Value);
        }
    }

    return undefined;
}

// ExtendedAttribute FieldIDs used for storing Dataverse IDs
const FIELD_ID_DATAVERSE_TASK_ID = 188743731;      // Text1 field
const FIELD_ID_DATAVERSE_ASSIGNMENT_ID = 188743734; // Text2 field

/**
 * Parse MSPDI XML content and extract project data
 */
export async function parseMspdiXml(xmlContent: string): Promise<ImportedProjectData> {
    return new Promise((resolve, reject) => {
        parseString(xmlContent, { explicitArray: false, ignoreAttrs: false }, (err, result) => {
            if (err) {
                reject(new Error(`Failed to parse XML: ${err.message}`));
                return;
            }

            try {
                const project = result.Project;
                if (!project) {
                    reject(new Error('Invalid MSPDI XML: Missing Project element'));
                    return;
                }

                const projectName = getTextValue(project.Name) || getTextValue(project.Title) || 'Imported Project';
                const startDate = parseMspdiDate(getTextValue(project.StartDate));

                // Parse tasks
                const tasks: ImportedTask[] = [];
                const tasksElement = project.Tasks?.Task;
                const taskArray = Array.isArray(tasksElement) ? tasksElement : (tasksElement ? [tasksElement] : []);

                for (const task of taskArray) {
                    const uid = getIntValue(task.UID);
                    const id = getIntValue(task.ID);

                    // Skip project summary task (UID 0)
                    if (uid === 0) continue;
                    if (uid === undefined || id === undefined) continue;

                    const outlineLevel = getIntValue(task.OutlineLevel) || 1;
                    const isSummary = getTextValue(task.Summary) === '1';
                    const isMilestone = getTextValue(task.Milestone) === '1';

                    // Extract Dataverse Task ID from ExtendedAttribute (Text1 field)
                    const dataverseTaskId = getExtendedAttributeValue(task, FIELD_ID_DATAVERSE_TASK_ID);

                    tasks.push({
                        uid,
                        id,
                        name: getTextValue(task.Name) || 'Unnamed Task',
                        startDate: parseMspdiDate(getTextValue(task.Start)),
                        finishDate: parseMspdiDate(getTextValue(task.Finish)),
                        duration: parseMspdiDuration(getTextValue(task.Duration)),
                        percentComplete: getNumericValue(task.PercentComplete),
                        effort: parseMspdiWork(getTextValue(task.Work)),
                        outlineLevel,
                        notes: getTextValue(task.Notes),
                        wbs: getTextValue(task.WBS),
                        isSummary,
                        isMilestone,
                        dataverseTaskId
                    });
                }

                // Build parent-child relationships based on outline levels
                // Tasks are ordered by ID, and outline level determines hierarchy
                const taskStack: ImportedTask[] = [];
                for (const task of tasks) {
                    // Pop tasks from stack until we find a parent with lower outline level
                    while (taskStack.length > 0 && taskStack[taskStack.length - 1].outlineLevel >= task.outlineLevel) {
                        taskStack.pop();
                    }

                    // If stack is not empty, the top task is the parent
                    if (taskStack.length > 0) {
                        task.parentUid = taskStack[taskStack.length - 1].uid;
                    }

                    // Push current task to stack
                    taskStack.push(task);
                }

                // Parse dependencies from predecessor links within tasks
                const dependencies: ImportedDependency[] = [];
                for (const task of taskArray) {
                    const toTaskUid = getIntValue(task.UID);
                    if (toTaskUid === undefined || toTaskUid === 0) continue;

                    const predecessorLinks = task.PredecessorLink;
                    const linkArray = Array.isArray(predecessorLinks) ? predecessorLinks : (predecessorLinks ? [predecessorLinks] : []);

                    for (const link of linkArray) {
                        const fromTaskUid = getIntValue(link.PredecessorUID);
                        if (fromTaskUid === undefined || fromTaskUid === 0) continue;

                        const type = getIntValue(link.Type) ?? 1; // Default to FS
                        const linkLag = getIntValue(link.LinkLag) || 0;
                        // LinkLag is in tenths of minutes, convert to days (assuming 8h day = 480 minutes)
                        const lagDays = linkLag / 4800;

                        dependencies.push({
                            fromTaskUid,
                            toTaskUid,
                            type,
                            lag: lagDays !== 0 ? lagDays : undefined
                        });
                    }
                }

                // Parse resources
                const resources: ImportedResource[] = [];
                const resourcesElement = project.Resources?.Resource;
                const resourceArray = Array.isArray(resourcesElement) ? resourcesElement : (resourcesElement ? [resourcesElement] : []);

                for (const resource of resourceArray) {
                    const uid = getIntValue(resource.UID);
                    const id = getIntValue(resource.ID);

                    if (uid === undefined || id === undefined) continue;
                    // Skip null resources (UID 0 is often unassigned)
                    if (uid === 0) continue;

                    const resourceType = getIntValue(resource.Type);
                    // Type 1 = Work resource, Type 0 = Material, Type 2 = Cost
                    // We primarily want work resources
                    if (resourceType !== undefined && resourceType !== 1) continue;

                    resources.push({
                        uid,
                        id,
                        name: getTextValue(resource.Name) || 'Unknown Resource',
                        email: getTextValue(resource.EmailAddress)
                    });
                }

                // Parse assignments
                const assignments: ImportedAssignment[] = [];
                const assignmentsElement = project.Assignments?.Assignment;
                const assignmentArray = Array.isArray(assignmentsElement) ? assignmentsElement : (assignmentsElement ? [assignmentsElement] : []);

                for (const assignment of assignmentArray) {
                    const uid = getIntValue(assignment.UID);
                    const taskUid = getIntValue(assignment.TaskUID);
                    const resourceUid = getIntValue(assignment.ResourceUID);

                    if (uid === undefined || taskUid === undefined || resourceUid === undefined) continue;
                    // Skip if task or resource is 0 (unassigned)
                    if (taskUid === 0 || resourceUid === 0) continue;

                    const unitsRaw = getNumericValue(assignment.Units); // Decimal (1 = 100%)
                    // Convert to percentage and ensure it's a valid number (default 100)
                    let units = 100;
                    if (unitsRaw !== undefined && Number.isFinite(unitsRaw)) {
                        units = Math.round(unitsRaw * 100); // Convert decimal to percentage
                        // Clamp to reasonable range (0-1000%)
                        if (units < 0) units = 0;
                        if (units > 1000) units = 1000;
                    }

                    // Extract start and finish dates
                    const startDate = getTextValue(assignment.Start);
                    const finishDate = getTextValue(assignment.Finish);

                    // Extract Dataverse Assignment ID from ExtendedAttribute (Text2 field)
                    const dataverseAssignmentId = getExtendedAttributeValue(assignment, FIELD_ID_DATAVERSE_ASSIGNMENT_ID);

                    assignments.push({
                        uid,
                        taskUid,
                        resourceUid,
                        units, // Already converted to percentage
                        startDate,
                        finishDate,
                        dataverseAssignmentId
                    });
                }

                resolve({
                    projectName,
                    startDate,
                    tasks,
                    resources,
                    assignments,
                    dependencies
                });
            } catch (parseError: any) {
                reject(new Error(`Failed to extract data from XML: ${parseError.message}`));
            }
        });
    });
}

/**
 * Convert imported project data to Bryntum/Dataverse format
 */
export function convertImportedDataToBryntum(data: ImportedProjectData): BryntumImportData {
    // Build UID to task mapping
    const uidToTask = new Map<number, ImportedTask>();
    data.tasks.forEach(task => uidToTask.set(task.uid, task));

    // Build UID to resource mapping
    const uidToResource = new Map<number, ImportedResource>();
    data.resources.forEach(resource => uidToResource.set(resource.uid, resource));

    // Convert tasks to Bryntum format (flat list with parentId)
    const tasks = data.tasks.map(task => ({
        // Use UID as temporary ID (will be replaced by Dataverse GUID)
        id: `import_${task.uid}`,
        name: task.name,
        startDate: task.startDate,
        endDate: task.finishDate,
        duration: task.duration,
        percentDone: task.percentComplete || 0,
        effort: task.effort,
        note: task.notes,
        parentId: task.parentUid ? `import_${task.parentUid}` : undefined,
        // Prevent Bryntum from recalculating dates/duration and avoid formula errors
        manuallyScheduled: true,
        effortDriven: false,
        // Task index for sorting - use the actual ID from MPP (the task number shown in MS Project UI)
        taskIndex: task.id,
        // Store original UID for reference
        _importUid: task.uid,
        _outlineLevel: task.outlineLevel,
        // Dataverse ID for round-trip support (update existing tasks)
        _dataverseTaskId: task.dataverseTaskId
    }));

    // Convert resources to Bryntum format
    // Use email as ID if available, otherwise use the resource name
    const resources = data.resources.map(resource => {
        const identifier = resource.email || resource.name || `resource_${resource.uid}`;
        return {
            id: identifier.toLowerCase(),
            name: resource.name,
            email: identifier.toLowerCase(),
            _importUid: resource.uid
        };
    });

    // Build resource UID to ID mapping
    const resourceUidToId = new Map<number, string>();
    data.resources.forEach(resource => {
        const identifier = resource.email || resource.name || `resource_${resource.uid}`;
        resourceUidToId.set(resource.uid, identifier.toLowerCase());
    });

    // Convert assignments to Bryntum format
    let filteredOutCount = 0;
    const assignments = data.assignments
        .filter(assignment => {
            // Only include assignments where both task and resource exist
            const hasTask = uidToTask.has(assignment.taskUid);
            const hasResource = uidToResource.has(assignment.resourceUid);
            if (!hasTask || !hasResource) {
                filteredOutCount++;
                console.log(`[MSPDI Import] Filtering out assignment: taskUid=${assignment.taskUid} (exists=${hasTask}), resourceUid=${assignment.resourceUid} (exists=${hasResource})`);
            }
            return hasTask && hasResource;
        })
        .map(assignment => ({
            id: `assignment_${assignment.uid}`,
            event: `import_${assignment.taskUid}`,
            resource: resourceUidToId.get(assignment.resourceUid) || '',
            // Always include units with a valid numeric value to prevent "Unknown formula for `units`" error
            units: typeof assignment.units === 'number' && Number.isFinite(assignment.units) ? assignment.units : 100,
            startDate: assignment.startDate,
            finishDate: assignment.finishDate,
            // Dataverse ID for round-trip support (update existing assignments)
            _dataverseAssignmentId: assignment.dataverseAssignmentId
        }));

    // Convert dependencies to Bryntum format
    // Map MSPDI type to Bryntum type
    // MSPDI: 0=FF, 1=FS, 2=SF, 3=SS
    // Bryntum: 0=SS, 1=SF, 2=FS, 3=FF
    const typeMap: Record<number, number> = {
        0: 3, // FF
        1: 2, // FS
        2: 1, // SF
        3: 0  // SS
    };

    const dependencies = data.dependencies
        .filter(dep => {
            // Only include dependencies where both tasks exist
            return uidToTask.has(dep.fromTaskUid) && uidToTask.has(dep.toTaskUid);
        })
        .map((dep, index) => ({
            id: `dep_${index}`,
            fromTask: `import_${dep.fromTaskUid}`,
            toTask: `import_${dep.toTaskUid}`,
            type: typeMap[dep.type] ?? 2, // Default to FS
            lag: dep.lag
        }));

    if (filteredOutCount > 0) {
        console.log(`[MSPDI Import] WARNING: ${filteredOutCount} assignments were filtered out due to missing task/resource references`);
    }
    console.log(`[MSPDI Import] Converted: ${tasks.length} tasks, ${resources.length} resources, ${assignments.length} assignments, ${dependencies.length} dependencies`);

    return {
        tasks,
        resources,
        assignments,
        dependencies
    };
}

/**
 * Convert imported data to Dataverse format for saving
 */
export function convertImportedDataToDataverse(data: BryntumImportData, projectId?: string): {
    tasks: any[];
    resources: any[];
    assignments: any[];
} {
    // Map import IDs to indices for parent references
    const importIdToIndex = new Map<string, number>();
    data.tasks.forEach((task, index) => {
        importIdToIndex.set(task.id, index);
    });

    // Convert tasks to Dataverse format
    const tasks = data.tasks.map(task => {
        const dataverseTask: any = {
            eppm_name: task.name,
            eppm_startdate: toSafeISOString(task.startDate),
            eppm_finishdate: toSafeISOString(task.endDate || task.finishDate),
            eppm_taskduration: task.duration,
            eppm_pocpercentage: task.percentDone,
            eppm_taskwork: task.effort,
            eppm_notes: task.note,
            // Task index for sorting (from import UID)
            eppm_taskindex: task.taskIndex || task._importUid,
            // Store import ID for reference during creation
            _importId: task.id,
            _parentImportId: task.parentId,
            // Dataverse task ID for round-trip support (update existing tasks)
            _dataverseTaskId: task._dataverseTaskId
        };

        if (projectId) {
            dataverseTask.eppm_projectid = projectId;
        }

        return dataverseTask;
    });

    // Resources (for creating in assignment table)
    const resources = data.resources.map(resource => ({
        email: resource.email || resource.id,
        name: resource.name,
        _importUid: resource._importUid
    }));

    // Assignments (will be created after tasks are saved)
    const assignments = data.assignments.map(assignment => ({
        taskImportId: assignment.event,
        resourceEmail: assignment.resource,
        // Always include units with a valid numeric value
        units: typeof assignment.units === 'number' && Number.isFinite(assignment.units) ? assignment.units : 100,
        startDate: assignment.startDate,
        finishDate: assignment.finishDate,
        // Dataverse assignment ID for round-trip support (update existing assignments)
        _dataverseAssignmentId: assignment._dataverseAssignmentId
    }));

    return {
        tasks,
        resources,
        assignments
    };
}

/**
 * Build predecessor string from dependencies for a specific task
 */
export function buildPredecessorStringForTask(
    taskImportId: string,
    dependencies: any[],
    importIdToDataverseId: Map<string, string>
): string {
    const predecessors = dependencies.filter(dep => dep.toTask === taskImportId);

    if (predecessors.length === 0) return '';

    // Map Bryntum type to string
    const typeToString: Record<number, string> = {
        0: 'SS',
        1: 'SF',
        2: 'FS',
        3: 'FF'
    };

    const tokens = predecessors.map(dep => {
        const fromId = importIdToDataverseId.get(dep.fromTask) || dep.fromTask;
        const typeStr = typeToString[dep.type] || 'FS';
        let token = `${fromId}${typeStr}`;

        if (dep.lag && dep.lag !== 0) {
            const lagStr = dep.lag > 0 ? `+${dep.lag}d` : `${dep.lag}d`;
            token += lagStr;
        }

        return token;
    });

    return tokens.join(';');
}

/**
 * Build successor string from dependencies for a specific task
 */
export function buildSuccessorStringForTask(
    taskImportId: string,
    dependencies: any[],
    importIdToDataverseId: Map<string, string>
): string {
    const successors = dependencies.filter(dep => dep.fromTask === taskImportId);

    if (successors.length === 0) return '';

    // Map Bryntum type to string
    const typeToString: Record<number, string> = {
        0: 'SS',
        1: 'SF',
        2: 'FS',
        3: 'FF'
    };

    const tokens = successors.map(dep => {
        const toId = importIdToDataverseId.get(dep.toTask) || dep.toTask;
        const typeStr = typeToString[dep.type] || 'FS';
        let token = `${toId}${typeStr}`;

        if (dep.lag && dep.lag !== 0) {
            const lagStr = dep.lag > 0 ? `+${dep.lag}d` : `${dep.lag}d`;
            token += lagStr;
        }

        return token;
    });

    return tokens.join(';');
}
