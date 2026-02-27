/**
 * MSPDI (Microsoft Project Data Interchange) XML Exporter
 *
 * Generates XML in Microsoft Project's MSPDI format that can be:
 * - Opened directly by MS Project
 * - Imported into MS Project and saved as .mpp
 *
 * Reference: https://docs.microsoft.com/en-us/office-project-server-sdk/schema-reference
 */

export interface MspdiTask {
    id: string;
    uid: number;
    name: string;
    startDate?: string;
    finishDate?: string;
    duration?: number; // in days
    percentComplete?: number;
    effort?: number; // in hours
    parentId?: string;
    outlineLevel?: number;
    predecessors?: MspdiPredecessor[];
    notes?: string;
    // Dataverse ID for round-trip support
    dataverseTaskId?: string;
}

export interface MspdiResource {
    id: string;
    uid: number;
    name: string;
    email?: string;
}

export interface MspdiAssignment {
    taskUid: number;
    resourceUid: number;
    units?: number; // percentage (100 = 100%)
    // Dataverse assignment ID for round-trip support
    dataverseAssignmentId?: string;
}

export interface MspdiPredecessor {
    predecessorUid: number;
    type: number; // 0=FF, 1=FS, 2=SF, 3=SS
    lag?: number; // in days
}

export interface MspdiProjectData {
    projectName?: string;
    startDate?: string;
    tasks: MspdiTask[];
    resources: MspdiResource[];
    assignments: MspdiAssignment[];
}

/**
 * Escape XML special characters
 */
function escapeXml(str: string | null | undefined): string {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Format date to ISO 8601 format required by MSPDI.
 * Returns empty string if date is invalid.
 * @param dateStr - date string (YYYY-MM-DD or ISO format)
 * @param defaultTime - time to use for date-only strings: 'start' (08:00:00) or 'finish' (17:00:00)
 */
function formatMspdiDate(dateStr: string | null | undefined, defaultTime: 'start' | 'finish' = 'start'): string {
    if (!dateStr) return '';
    try {
        // Extract YYYY-MM-DD portion directly to avoid timezone issues
        const dateMatch = dateStr.match(/^(\d{4}-\d{2}-\d{2})/);
        if (dateMatch) {
            const time = defaultTime === 'finish' ? 'T17:00:00' : 'T08:00:00';
            return dateMatch[1] + time;
        }

        // Fallback: parse and use UTC components
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return '';
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        const day = String(date.getUTCDate()).padStart(2, '0');
        const time = defaultTime === 'finish' ? 'T17:00:00' : 'T08:00:00';
        return `${year}-${month}-${day}${time}`;
    } catch {
        return '';
    }
}

/**
 * Convert duration in days to MSPDI duration format (PT8H0M0S for 1 day = 8 hours)
 */
function formatMspdiDuration(days: number | null | undefined): string {
    if (days === null || days === undefined || isNaN(days)) return 'PT0H0M0S';
    const hours = Math.round(days * 8); // Assuming 8 hours per day
    return `PT${hours}H0M0S`;
}

/**
 * Convert effort in hours to MSPDI work format
 */
function formatMspdiWork(hours: number | null | undefined): string {
    if (hours === null || hours === undefined || isNaN(hours)) return 'PT0H0M0S';
    return `PT${Math.round(hours)}H0M0S`;
}

/**
 * Calculate outline level for hierarchical tasks
 */
function calculateOutlineLevels(tasks: MspdiTask[]): Map<string, number> {
    const levelMap = new Map<string, number>();
    const childToParent = new Map<string, string>();

    // Build parent-child relationships
    tasks.forEach(task => {
        if (task.parentId) {
            childToParent.set(task.id, task.parentId);
        }
    });

    // Calculate levels
    const getLevel = (taskId: string): number => {
        if (levelMap.has(taskId)) {
            return levelMap.get(taskId)!;
        }
        const parentId = childToParent.get(taskId);
        if (!parentId) {
            levelMap.set(taskId, 1);
            return 1;
        }
        const parentLevel = getLevel(parentId);
        const level = parentLevel + 1;
        levelMap.set(taskId, level);
        return level;
    };

    tasks.forEach(task => getLevel(task.id));
    return levelMap;
}

/**
 * Build task ID to UID mapping
 */
function buildTaskUidMap(tasks: MspdiTask[]): Map<string, number> {
    const uidMap = new Map<string, number>();
    tasks.forEach((task, index) => {
        // UID starts from 1 (0 is reserved for project summary task)
        uidMap.set(task.id, index + 1);
    });
    return uidMap;
}

/**
 * Generate MSPDI XML content
 */
export function generateMspdiXml(data: MspdiProjectData): string {
    const { projectName = 'Exported Project', startDate, tasks, resources, assignments } = data;

    // Build UID maps
    const taskUidMap = buildTaskUidMap(tasks);
    const outlineLevels = calculateOutlineLevels(tasks);

    // Assign UIDs to tasks
    tasks.forEach((task, index) => {
        task.uid = index + 1;
        task.outlineLevel = outlineLevels.get(task.id) || 1;
    });

    // Assign UIDs to resources
    resources.forEach((resource, index) => {
        resource.uid = index + 1;
    });

    // Build resource ID to UID map
    const resourceUidMap = new Map<string, number>();
    resources.forEach(resource => {
        resourceUidMap.set(resource.id, resource.uid);
    });

    // Calculate project dates
    // Use the provided startDate if available, otherwise calculate from tasks
    let projectStartDate = startDate;
    let projectFinishDate = startDate;

    if (tasks.length > 0) {
        const startDates = tasks
            .map(t => t.startDate)
            .filter(d => d)
            .map(d => {
                try {
                    const date = new Date(d!);
                    return isNaN(date.getTime()) ? null : date.getTime();
                } catch {
                    return null;
                }
            })
            .filter(d => d !== null) as number[];

        const finishDates = tasks
            .map(t => t.finishDate)
            .filter(d => d)
            .map(d => {
                try {
                    const date = new Date(d!);
                    return isNaN(date.getTime()) ? null : date.getTime();
                } catch {
                    return null;
                }
            })
            .filter(d => d !== null) as number[];

        if (startDates.length > 0) {
            const minStart = new Date(Math.min(...startDates));
            // Only override if we don't have a startDate or if calculated date is earlier
            if (!projectStartDate || minStart.getTime() < new Date(projectStartDate).getTime()) {
                projectStartDate = minStart.toISOString().split('T')[0];
            }
        }

        if (finishDates.length > 0) {
            projectFinishDate = new Date(Math.max(...finishDates)).toISOString().split('T')[0];
        } else if (startDates.length > 0) {
            // If no finish dates, use latest start date
            projectFinishDate = new Date(Math.max(...startDates)).toISOString().split('T')[0];
        }
    }

    // Ensure projectStartDate is set (use today as absolute fallback)
    if (!projectStartDate) {
        projectStartDate = new Date().toISOString().split('T')[0];
    }
    if (!projectFinishDate) {
        projectFinishDate = projectStartDate;
    }

    // Generate XML
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Project xmlns="http://schemas.microsoft.com/project">
    <SaveVersion>14</SaveVersion>
    <Name>${escapeXml(projectName)}</Name>
    <Title>${escapeXml(projectName)}</Title>
    <ScheduleFromStart>1</ScheduleFromStart>
    <StartDate>${formatMspdiDate(projectStartDate, 'start')}</StartDate>
    <FinishDate>${formatMspdiDate(projectFinishDate, 'finish')}</FinishDate>
    <FYStartDate>1</FYStartDate>
    <CriticalSlackLimit>0</CriticalSlackLimit>
    <CurrencyDigits>2</CurrencyDigits>
    <CurrencySymbol>$</CurrencySymbol>
    <CurrencyCode>USD</CurrencyCode>
    <CurrencySymbolPosition>0</CurrencySymbolPosition>
    <CalendarUID>1</CalendarUID>
    <DefaultStartTime>08:00:00</DefaultStartTime>
    <DefaultFinishTime>17:00:00</DefaultFinishTime>
    <MinutesPerDay>480</MinutesPerDay>
    <MinutesPerWeek>2400</MinutesPerWeek>
    <DaysPerMonth>20</DaysPerMonth>
    <DefaultTaskType>1</DefaultTaskType>
    <DefaultFixedCostAccrual>2</DefaultFixedCostAccrual>
    <DefaultStandardRate>0</DefaultStandardRate>
    <DefaultOvertimeRate>0</DefaultOvertimeRate>
    <DurationFormat>7</DurationFormat>
    <WorkFormat>2</WorkFormat>
    <EditableActualCosts>0</EditableActualCosts>
    <HonorConstraints>1</HonorConstraints>
    <InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>
    <MultipleCriticalPaths>0</MultipleCriticalPaths>
    <NewTasksEffortDriven>1</NewTasksEffortDriven>
    <NewTasksEstimated>1</NewTasksEstimated>
    <SplitsInProgressTasks>1</SplitsInProgressTasks>
    <SpreadActualCost>0</SpreadActualCost>
    <SpreadPercentComplete>0</SpreadPercentComplete>
    <TaskUpdatesResource>1</TaskUpdatesResource>
    <FiscalYearStart>0</FiscalYearStart>
    <WeekStartDay>0</WeekStartDay>
    <MoveCompletedEndsBack>0</MoveCompletedEndsBack>
    <MoveRemainingStartsBack>0</MoveRemainingStartsBack>
    <MoveRemainingStartsForward>0</MoveRemainingStartsForward>
    <MoveCompletedEndsForward>0</MoveCompletedEndsForward>
    <BaselineForEarnedValue>0</BaselineForEarnedValue>
    <AutoAddNewResourcesAndTasks>1</AutoAddNewResourcesAndTasks>
    <CurrentDate>${formatMspdiDate(new Date().toISOString())}</CurrentDate>
    <MicrosoftProjectServerURL>1</MicrosoftProjectServerURL>
    <Autolink>1</Autolink>
    <NewTaskStartDate>0</NewTaskStartDate>
    <NewTasksAreManual>0</NewTasksAreManual>
    <DefaultTaskEVMethod>0</DefaultTaskEVMethod>
    <ProjectExternallyEdited>0</ProjectExternallyEdited>
    <ExtendedCreationDate>${formatMspdiDate(new Date().toISOString())}</ExtendedCreationDate>
    <ActualsInSync>1</ActualsInSync>
    <RemoveFileProperties>0</RemoveFileProperties>
    <AdminProject>0</AdminProject>
    <ExtendedAttributes>
        <ExtendedAttribute>
            <FieldID>188743731</FieldID>
            <FieldName>Text1</FieldName>
            <Alias>DataverseTaskID</Alias>
            <Ltuid>DataverseTaskID</Ltuid>
        </ExtendedAttribute>
        <ExtendedAttribute>
            <FieldID>188743734</FieldID>
            <FieldName>Text2</FieldName>
            <Alias>DataverseAssignmentID</Alias>
            <Ltuid>DataverseAssignmentID</Ltuid>
        </ExtendedAttribute>
    </ExtendedAttributes>
`;

    // Add Calendar
    xml += `    <Calendars>
        <Calendar>
            <UID>1</UID>
            <Name>Standard</Name>
            <IsBaseCalendar>1</IsBaseCalendar>
            <IsBaselineCalendar>0</IsBaselineCalendar>
            <WeekDays>
                <WeekDay>
                    <DayType>1</DayType>
                    <DayWorking>0</DayWorking>
                </WeekDay>
                <WeekDay>
                    <DayType>2</DayType>
                    <DayWorking>1</DayWorking>
                    <WorkingTimes>
                        <WorkingTime>
                            <FromTime>08:00:00</FromTime>
                            <ToTime>12:00:00</ToTime>
                        </WorkingTime>
                        <WorkingTime>
                            <FromTime>13:00:00</FromTime>
                            <ToTime>17:00:00</ToTime>
                        </WorkingTime>
                    </WorkingTimes>
                </WeekDay>
                <WeekDay>
                    <DayType>3</DayType>
                    <DayWorking>1</DayWorking>
                    <WorkingTimes>
                        <WorkingTime>
                            <FromTime>08:00:00</FromTime>
                            <ToTime>12:00:00</ToTime>
                        </WorkingTime>
                        <WorkingTime>
                            <FromTime>13:00:00</FromTime>
                            <ToTime>17:00:00</ToTime>
                        </WorkingTime>
                    </WorkingTimes>
                </WeekDay>
                <WeekDay>
                    <DayType>4</DayType>
                    <DayWorking>1</DayWorking>
                    <WorkingTimes>
                        <WorkingTime>
                            <FromTime>08:00:00</FromTime>
                            <ToTime>12:00:00</ToTime>
                        </WorkingTime>
                        <WorkingTime>
                            <FromTime>13:00:00</FromTime>
                            <ToTime>17:00:00</ToTime>
                        </WorkingTime>
                    </WorkingTimes>
                </WeekDay>
                <WeekDay>
                    <DayType>5</DayType>
                    <DayWorking>1</DayWorking>
                    <WorkingTimes>
                        <WorkingTime>
                            <FromTime>08:00:00</FromTime>
                            <ToTime>12:00:00</ToTime>
                        </WorkingTime>
                        <WorkingTime>
                            <FromTime>13:00:00</FromTime>
                            <ToTime>17:00:00</ToTime>
                        </WorkingTime>
                    </WorkingTimes>
                </WeekDay>
                <WeekDay>
                    <DayType>6</DayType>
                    <DayWorking>1</DayWorking>
                    <WorkingTimes>
                        <WorkingTime>
                            <FromTime>08:00:00</FromTime>
                            <ToTime>12:00:00</ToTime>
                        </WorkingTime>
                        <WorkingTime>
                            <FromTime>13:00:00</FromTime>
                            <ToTime>17:00:00</ToTime>
                        </WorkingTime>
                    </WorkingTimes>
                </WeekDay>
                <WeekDay>
                    <DayType>7</DayType>
                    <DayWorking>0</DayWorking>
                </WeekDay>
            </WeekDays>
        </Calendar>
    </Calendars>
`;

    // Add Tasks
    xml += `    <Tasks>
`;

    // Add project summary task (UID 0)
    xml += `        <Task>
            <UID>0</UID>
            <ID>0</ID>
            <Name>${escapeXml(projectName)}</Name>
            <Type>1</Type>
            <IsNull>0</IsNull>
            <CreateDate>${formatMspdiDate(new Date().toISOString())}</CreateDate>
            <WBS>0</WBS>
            <OutlineNumber>0</OutlineNumber>
            <OutlineLevel>0</OutlineLevel>
            <Priority>500</Priority>
            <Start>${formatMspdiDate(projectStartDate, 'start')}</Start>
            <Finish>${formatMspdiDate(projectFinishDate, 'finish')}</Finish>
            <Duration>${formatMspdiDuration(0)}</Duration>
            <DurationFormat>7</DurationFormat>
            <Work>${formatMspdiWork(0)}</Work>
            <ResumeValid>0</ResumeValid>
            <EffortDriven>0</EffortDriven>
            <Recurring>0</Recurring>
            <OverAllocated>0</OverAllocated>
            <Estimated>1</Estimated>
            <Milestone>0</Milestone>
            <Summary>1</Summary>
            <Critical>0</Critical>
            <IsSubproject>0</IsSubproject>
            <IsSubprojectReadOnly>0</IsSubprojectReadOnly>
            <ExternalTask>0</ExternalTask>
            <EarlyStart>${formatMspdiDate(projectStartDate, 'start')}</EarlyStart>
            <EarlyFinish>${formatMspdiDate(projectFinishDate, 'finish')}</EarlyFinish>
            <LateStart>${formatMspdiDate(projectStartDate, 'start')}</LateStart>
            <LateFinish>${formatMspdiDate(projectFinishDate, 'finish')}</LateFinish>
            <StartVariance>0</StartVariance>
            <FinishVariance>0</FinishVariance>
            <WorkVariance>0</WorkVariance>
            <FreeSlack>0</FreeSlack>
            <TotalSlack>0</TotalSlack>
            <FixedCost>0</FixedCost>
            <FixedCostAccrual>3</FixedCostAccrual>
            <PercentComplete>0</PercentComplete>
            <PercentWorkComplete>0</PercentWorkComplete>
            <Cost>0</Cost>
            <OvertimeCost>0</OvertimeCost>
            <OvertimeWork>PT0H0M0S</OvertimeWork>
            <ActualStart>${formatMspdiDate(projectStartDate, 'start')}</ActualStart>
            <ActualDuration>PT0H0M0S</ActualDuration>
            <ActualCost>0</ActualCost>
            <ActualOvertimeCost>0</ActualOvertimeCost>
            <ActualWork>PT0H0M0S</ActualWork>
            <ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>
            <RegularWork>PT0H0M0S</RegularWork>
            <RemainingDuration>${formatMspdiDuration(0)}</RemainingDuration>
            <RemainingCost>0</RemainingCost>
            <RemainingWork>PT0H0M0S</RemainingWork>
            <RemainingOvertimeCost>0</RemainingOvertimeCost>
            <RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>
            <ACWP>0</ACWP>
            <CV>0</CV>
            <ConstraintType>0</ConstraintType>
            <CalendarUID>-1</CalendarUID>
            <LevelAssignments>1</LevelAssignments>
            <LevelingCanSplit>1</LevelingCanSplit>
            <LevelingDelay>0</LevelingDelay>
            <LevelingDelayFormat>8</LevelingDelayFormat>
            <IgnoreResourceCalendar>0</IgnoreResourceCalendar>
            <HideBar>0</HideBar>
            <Rollup>0</Rollup>
            <BCWS>0</BCWS>
            <BCWP>0</BCWP>
            <PhysicalPercentComplete>0</PhysicalPercentComplete>
            <EarnedValueMethod>0</EarnedValueMethod>
            <IsPublished>1</IsPublished>
            <CommitmentType>0</CommitmentType>
        </Task>
`;

    // Pre-compute total work hours per task for consistent task/assignment work values
    // MS Project expects: Task Work = sum of Assignment Work
    const taskWorkHoursMap = new Map<number, number>();
    assignments.forEach(assignment => {
        const task = tasks.find(t => t.uid === assignment.taskUid);
        if (!task) return;
        const units = (assignment.units || 100) / 100;
        let workHours = 0;
        if (task.effort && task.effort > 0) {
            // Task has explicit effort — each assignment's share is computed later,
            // but total task work = task.effort
            // We only need to mark that effort exists; actual value is task.effort itself
        } else if (task.duration && task.duration > 0) {
            workHours = task.duration * 8 * units;
        }
        taskWorkHoursMap.set(task.uid, (taskWorkHoursMap.get(task.uid) || 0) + workHours);
    });

    // Add project tasks
    tasks.forEach((task, index) => {
        const taskId = index + 1;
        const outlineLevel = task.outlineLevel || 1;
        const hasChildren = tasks.some(t => t.parentId === task.id);
        const isSummary = hasChildren ? 1 : 0;
        const isMilestone = (task.duration === 0 || !task.duration) && !hasChildren ? 1 : 0;

        // Calculate task work: use explicit effort if available, otherwise sum from assignments
        let taskWorkHours = 0;
        if (task.effort && task.effort > 0) {
            taskWorkHours = task.effort;
        } else {
            taskWorkHours = taskWorkHoursMap.get(task.uid) || 0;
        }
        const taskWork = formatMspdiWork(taskWorkHours);

        xml += `        <Task>
            <UID>${task.uid}</UID>
            <ID>${taskId}</ID>
            <Name>${escapeXml(task.name)}</Name>
            <Type>1</Type>
            <IsNull>0</IsNull>
            <CreateDate>${formatMspdiDate(new Date().toISOString())}</CreateDate>
            <WBS>${taskId}</WBS>
            <OutlineNumber>${taskId}</OutlineNumber>
            <OutlineLevel>${outlineLevel}</OutlineLevel>
            <Priority>500</Priority>
            <Start>${formatMspdiDate(task.startDate, 'start')}</Start>
            <Finish>${formatMspdiDate(task.finishDate, 'finish')}</Finish>
            <Duration>${formatMspdiDuration(task.duration)}</Duration>
            <DurationFormat>7</DurationFormat>
            <Work>${taskWork}</Work>
            <ResumeValid>0</ResumeValid>
            <EffortDriven>1</EffortDriven>
            <Recurring>0</Recurring>
            <OverAllocated>0</OverAllocated>
            <Estimated>0</Estimated>
            <Milestone>${isMilestone}</Milestone>
            <Summary>${isSummary}</Summary>
            <Critical>0</Critical>
            <IsSubproject>0</IsSubproject>
            <IsSubprojectReadOnly>0</IsSubprojectReadOnly>
            <ExternalTask>0</ExternalTask>
            <EarlyStart>${formatMspdiDate(task.startDate, 'start')}</EarlyStart>
            <EarlyFinish>${formatMspdiDate(task.finishDate, 'finish')}</EarlyFinish>
            <LateStart>${formatMspdiDate(task.startDate, 'start')}</LateStart>
            <LateFinish>${formatMspdiDate(task.finishDate, 'finish')}</LateFinish>
            <StartVariance>0</StartVariance>
            <FinishVariance>0</FinishVariance>
            <WorkVariance>0</WorkVariance>
            <FreeSlack>0</FreeSlack>
            <TotalSlack>0</TotalSlack>
            <FixedCost>0</FixedCost>
            <FixedCostAccrual>3</FixedCostAccrual>
            <PercentComplete>${task.percentComplete || 0}</PercentComplete>
            <PercentWorkComplete>${task.percentComplete || 0}</PercentWorkComplete>
            <Cost>0</Cost>
            <OvertimeCost>0</OvertimeCost>
            <OvertimeWork>PT0H0M0S</OvertimeWork>
            <ActualDuration>PT0H0M0S</ActualDuration>
            <ActualCost>0</ActualCost>
            <ActualOvertimeCost>0</ActualOvertimeCost>
            <ActualWork>PT0H0M0S</ActualWork>
            <ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>
            <RegularWork>${taskWork}</RegularWork>
            <RemainingDuration>${formatMspdiDuration(task.duration)}</RemainingDuration>
            <RemainingCost>0</RemainingCost>
            <RemainingWork>${taskWork}</RemainingWork>
            <RemainingOvertimeCost>0</RemainingOvertimeCost>
            <RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>
            <ACWP>0</ACWP>
            <CV>0</CV>
            <ConstraintType>0</ConstraintType>
            <CalendarUID>-1</CalendarUID>
            <LevelAssignments>1</LevelAssignments>
            <LevelingCanSplit>1</LevelingCanSplit>
            <LevelingDelay>0</LevelingDelay>
            <LevelingDelayFormat>8</LevelingDelayFormat>
            <IgnoreResourceCalendar>0</IgnoreResourceCalendar>
            <HideBar>0</HideBar>
            <Rollup>0</Rollup>
            <BCWS>0</BCWS>
            <BCWP>0</BCWP>
            <PhysicalPercentComplete>0</PhysicalPercentComplete>
            <EarnedValueMethod>0</EarnedValueMethod>
            <IsPublished>1</IsPublished>
            <CommitmentType>0</CommitmentType>
`;

        // Add notes if present
        if (task.notes) {
            xml += `            <Notes>${escapeXml(task.notes)}</Notes>
`;
        }

        // Add Dataverse Task ID as ExtendedAttribute (Text1 field)
        // This enables round-trip editing - imported files will update existing tasks
        if (task.dataverseTaskId || task.id) {
            const dataverseId = task.dataverseTaskId || task.id;
            xml += `            <ExtendedAttribute>
                <FieldID>188743731</FieldID>
                <Value>${escapeXml(dataverseId)}</Value>
            </ExtendedAttribute>
`;
        }

        // Add predecessor links
        if (task.predecessors && task.predecessors.length > 0) {
            task.predecessors.forEach(pred => {
                xml += `            <PredecessorLink>
                <PredecessorUID>${pred.predecessorUid}</PredecessorUID>
                <Type>${pred.type}</Type>
                <CrossProject>0</CrossProject>
                <LinkLag>${(pred.lag || 0) * 4800}</LinkLag>
                <LagFormat>7</LagFormat>
            </PredecessorLink>
`;
            });
        }

        xml += `        </Task>
`;
    });

    xml += `    </Tasks>
`;

    // Add Resources
    xml += `    <Resources>
`;

    resources.forEach((resource, index) => {
        xml += `        <Resource>
            <UID>${resource.uid}</UID>
            <ID>${index + 1}</ID>
            <Name>${escapeXml(resource.name)}</Name>
            <Type>1</Type>
            <IsNull>0</IsNull>
            <MaxUnits>1</MaxUnits>
            <PeakUnits>1</PeakUnits>
            <OverAllocated>0</OverAllocated>
            <CanLevel>1</CanLevel>
            <AccrueAt>3</AccrueAt>
            <Work>PT0H0M0S</Work>
            <RegularWork>PT0H0M0S</RegularWork>
            <OvertimeWork>PT0H0M0S</OvertimeWork>
            <ActualWork>PT0H0M0S</ActualWork>
            <RemainingWork>PT0H0M0S</RemainingWork>
            <ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>
            <RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>
            <PercentWorkComplete>0</PercentWorkComplete>
            <StandardRate>0</StandardRate>
            <StandardRateFormat>2</StandardRateFormat>
            <Cost>0</Cost>
            <OvertimeRate>0</OvertimeRate>
            <OvertimeRateFormat>2</OvertimeRateFormat>
            <OvertimeCost>0</OvertimeCost>
            <CostPerUse>0</CostPerUse>
            <ActualCost>0</ActualCost>
            <ActualOvertimeCost>0</ActualOvertimeCost>
            <RemainingCost>0</RemainingCost>
            <RemainingOvertimeCost>0</RemainingOvertimeCost>
            <WorkVariance>0</WorkVariance>
            <CostVariance>0</CostVariance>
            <SV>0</SV>
            <CV>0</CV>
            <ACWP>0</ACWP>
            <CalendarUID>1</CalendarUID>
            <BCWS>0</BCWS>
            <BCWP>0</BCWP>
            <IsGeneric>0</IsGeneric>
            <IsInactive>0</IsInactive>
            <IsEnterprise>0</IsEnterprise>
            <BookingType>0</BookingType>
            <IsCostResource>0</IsCostResource>
`;

        if (resource.email) {
            xml += `            <EmailAddress>${escapeXml(resource.email)}</EmailAddress>
`;
        }

        xml += `        </Resource>
`;
    });

    xml += `    </Resources>
`;

    // Add Assignments
    xml += `    <Assignments>
`;

    // Build taskUid → task map for assignment work calculation
    const taskByUid = new Map<number, MspdiTask>();
    tasks.forEach(task => {
        taskByUid.set(task.uid, task);
    });

    // Group assignments by taskUid to distribute work among multiple resources
    const assignmentsByTask = new Map<number, MspdiAssignment[]>();
    assignments.forEach(assignment => {
        const list = assignmentsByTask.get(assignment.taskUid) || [];
        list.push(assignment);
        assignmentsByTask.set(assignment.taskUid, list);
    });

    let assignmentUid = 1;
    assignments.forEach(assignment => {
        const taskUid = assignment.taskUid;
        const resourceUid = assignment.resourceUid;
        const units = (assignment.units || 100) / 100; // Convert percentage to decimal

        // Calculate assignment work from the task's effort or duration
        const task = taskByUid.get(taskUid);
        let assignmentWorkHours = 0;
        if (task) {
            if (task.effort && task.effort > 0) {
                // Task has explicit effort (total work hours) — distribute proportionally
                const taskAssignments = assignmentsByTask.get(taskUid) || [];
                const totalUnits = taskAssignments.reduce((sum, a) => sum + ((a.units || 100) / 100), 0);
                assignmentWorkHours = totalUnits > 0 ? (task.effort * units / totalUnits) : task.effort;
            } else if (task.duration && task.duration > 0) {
                // No explicit effort — derive from duration: work = duration(days) * 8(hrs/day) * units
                assignmentWorkHours = task.duration * 8 * units;
            }
        }
        const assignmentWork = formatMspdiWork(assignmentWorkHours);
        const assignmentStart = task ? formatMspdiDate(task.startDate, 'start') : formatMspdiDate(new Date().toISOString());
        const assignmentFinish = task ? formatMspdiDate(task.finishDate, 'finish') : formatMspdiDate(new Date().toISOString());

        xml += `        <Assignment>
            <UID>${assignmentUid++}</UID>
            <TaskUID>${taskUid}</TaskUID>
            <ResourceUID>${resourceUid}</ResourceUID>
            <Units>${units}</Units>
            <PercentWorkComplete>0</PercentWorkComplete>
            <ActualCost>0</ActualCost>
            <ActualOvertimeCost>0</ActualOvertimeCost>
            <ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>
            <ActualWork>PT0H0M0S</ActualWork>
            <ACWP>0</ACWP>
            <Confirmed>0</Confirmed>
            <Cost>0</Cost>
            <CostRateTable>0</CostRateTable>
            <CostVariance>0</CostVariance>
            <CV>0</CV>
            <Delay>0</Delay>
            <Finish>${assignmentFinish}</Finish>
            <FinishVariance>0</FinishVariance>
            <WorkVariance>0</WorkVariance>
            <HasFixedRateUnits>1</HasFixedRateUnits>
            <FixedMaterial>0</FixedMaterial>
            <LevelingDelay>0</LevelingDelay>
            <LevelingDelayFormat>8</LevelingDelayFormat>
            <LinkedFields>0</LinkedFields>
            <Milestone>0</Milestone>
            <Overallocated>0</Overallocated>
            <OvertimeCost>0</OvertimeCost>
            <OvertimeWork>PT0H0M0S</OvertimeWork>
            <RegularWork>${assignmentWork}</RegularWork>
            <RemainingCost>0</RemainingCost>
            <RemainingOvertimeCost>0</RemainingOvertimeCost>
            <RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>
            <RemainingWork>${assignmentWork}</RemainingWork>
            <ResponsePending>0</ResponsePending>
            <Start>${assignmentStart}</Start>
            <StartVariance>0</StartVariance>
            <SV>0</SV>
            <Work>${assignmentWork}</Work>
            <WorkContour>0</WorkContour>
            <BCWS>0</BCWS>
            <BCWP>0</BCWP>
            <BookingType>0</BookingType>
            <CreationDate>${formatMspdiDate(new Date().toISOString())}</CreationDate>
            <BudgetCost>0</BudgetCost>
            <BudgetWork>PT0H0M0S</BudgetWork>
`;
        // Add Dataverse Assignment ID as ExtendedAttribute (Text2 field)
        // This enables round-trip editing - imported files will update existing assignments
        if (assignment.dataverseAssignmentId) {
            xml += `            <ExtendedAttribute>
                <FieldID>188743734</FieldID>
                <Value>${escapeXml(assignment.dataverseAssignmentId)}</Value>
            </ExtendedAttribute>
`;
        }

        xml += `        </Assignment>
`;
    });

    xml += `    </Assignments>
`;

    xml += `</Project>`;

    return xml;
}

/**
 * Convert Bryntum/Dataverse data to MSPDI format
 */
export function convertToMspdiFormat(
    tasks: any[],
    resources: any[],
    assignments: any[],
    dependencies: any[],
    projectName?: string
): MspdiProjectData {
    // Build task ID to UID map
    const taskIdToUid = new Map<string, number>();
    const flatTasks: MspdiTask[] = [];

    // Helper function to extract start date from task
    // Prefers rawStartDate (exact backend value) to avoid any transformation issues
    const extractStartDate = (task: any): string | undefined => {
        // Prefer raw backend value (exact Dataverse value, no transformation)
        if (task.rawStartDate && typeof task.rawStartDate === 'string') {
            return task.rawStartDate;
        }
        // Fallback: Bryntum format (startDate)
        if (task.startDate && typeof task.startDate === 'string') {
            return task.startDate.split('T')[0];
        }
        // Fallback: Dataverse format (eppm_startdate)
        if (task.eppm_startdate && typeof task.eppm_startdate === 'string') {
            return task.eppm_startdate.split('T')[0];
        }
        return undefined;
    };

    // Helper function to extract finish date from task
    // Prefers rawFinishDate (exact backend value) — no +1/-1 day conversion needed
    const extractFinishDate = (task: any): string | undefined => {
        // Prefer raw backend value (exact inclusive finish date from Dataverse)
        if (task.rawFinishDate && typeof task.rawFinishDate === 'string') {
            return task.rawFinishDate;
        }
        // Fallback: Bryntum endDate (exclusive, subtract 1 day using UTC to avoid timezone bugs)
        if (task.endDate && typeof task.endDate === 'string') {
            try {
                const dateStr = task.endDate.split('T')[0];
                const date = new Date(dateStr + 'T12:00:00Z');
                if (!isNaN(date.getTime())) {
                    date.setUTCDate(date.getUTCDate() - 1);
                    return date.toISOString().split('T')[0];
                }
            } catch {
                // Fall through to other options
            }
        }
        if (task.finishDate && typeof task.finishDate === 'string') {
            return task.finishDate.split('T')[0];
        }
        // Fallback: Dataverse format (eppm_finishdate)
        if (task.eppm_finishdate && typeof task.eppm_finishdate === 'string') {
            return task.eppm_finishdate.split('T')[0];
        }
        return undefined;
    };

    // Flatten hierarchical tasks and collect all tasks
    const flattenTasks = (taskList: any[], parentId?: string) => {
        taskList.forEach(task => {
            const taskId = String(task.id || task.eppm_projecttaskid || '');
            if (!taskId) return;

            const uid = flatTasks.length + 1;
            taskIdToUid.set(taskId, uid);

            flatTasks.push({
                id: taskId,
                uid,
                name: task.name || task.eppm_name || 'Unnamed Task',
                startDate: extractStartDate(task),
                finishDate: extractFinishDate(task),
                // Prefer rawDuration (exact backend value) over Bryntum's calculated duration
                duration: task.rawDuration ?? task.duration ?? task.eppm_taskduration,
                percentComplete: task.percentDone || task.eppm_pocpercentage || 0,
                effort: task.effort || task.eppm_taskwork,
                parentId: parentId || task.parentId || task.eppm_parenttaskid,
                notes: task.note || task.notes || task.eppm_notes,
                predecessors: [],
                // Store Dataverse task ID for round-trip support
                dataverseTaskId: task.eppm_projecttaskid || task.id || taskId
            });

            // Process children recursively
            if (task.children && Array.isArray(task.children)) {
                flattenTasks(task.children, taskId);
            }
        });
    };

    flattenTasks(tasks);

    // Process dependencies and add to tasks
    dependencies.forEach(dep => {
        const fromTaskId = String(dep.fromTask || dep.from || '');
        const toTaskId = String(dep.toTask || dep.to || '');

        if (!fromTaskId || !toTaskId) return;

        const fromUid = taskIdToUid.get(fromTaskId);
        const toUid = taskIdToUid.get(toTaskId);

        if (!fromUid || !toUid) return;

        // Find the target task and add predecessor
        const targetTask = flatTasks.find(t => t.id === toTaskId);
        if (targetTask) {
            // Map dependency type: Bryntum uses 0=SS, 1=SF, 2=FS, 3=FF
            // MSPDI uses 0=FF, 1=FS, 2=SF, 3=SS
            const bryntumType = dep.type ?? 2; // Default to FS
            let mspdiType = 1; // Default to FS
            switch (bryntumType) {
                case 0: mspdiType = 3; break; // SS
                case 1: mspdiType = 2; break; // SF
                case 2: mspdiType = 1; break; // FS
                case 3: mspdiType = 0; break; // FF
            }

            targetTask.predecessors!.push({
                predecessorUid: fromUid,
                type: mspdiType,
                lag: dep.lag
            });
        }
    });

    // Convert resources
    const mspdiResources: MspdiResource[] = resources.map((resource, index) => ({
        id: String(resource.id || resource.email || ''),
        uid: index + 1,
        name: resource.name || resource.email || 'Unknown Resource',
        email: resource.email
    }));

    // Build resource ID to UID map
    const resourceIdToUid = new Map<string, number>();
    mspdiResources.forEach(resource => {
        resourceIdToUid.set(resource.id, resource.uid);
    });

    // Convert assignments
    const mspdiAssignments: MspdiAssignment[] = [];
    assignments.forEach(assignment => {
        const taskId = String(assignment.event || assignment.taskId || assignment.task || '');
        const resourceId = String(assignment.resource || assignment.resourceId || '');

        const taskUid = taskIdToUid.get(taskId);
        const resourceUid = resourceIdToUid.get(resourceId);

        if (taskUid && resourceUid) {
            mspdiAssignments.push({
                taskUid,
                resourceUid,
                units: assignment.units || 100,
                // Store Dataverse assignment ID for round-trip support
                dataverseAssignmentId: assignment.id || assignment.assignmentId
            });
        }
    });

    // Calculate project start date (earliest task start date)
    let projectStartDate: string | undefined;
    if (flatTasks.length > 0) {
        const dates = flatTasks
            .map(t => t.startDate)
            .filter(d => d)
            .map(d => {
                try {
                    const date = new Date(d!);
                    return isNaN(date.getTime()) ? null : date.getTime();
                } catch {
                    return null;
                }
            })
            .filter(d => d !== null) as number[];

        if (dates.length > 0) {
            projectStartDate = new Date(Math.min(...dates)).toISOString().split('T')[0];
        }
    }

    // If project start date is still undefined, use today's date as fallback
    if (!projectStartDate) {
        projectStartDate = new Date().toISOString().split('T')[0];
    }

    // Ensure all tasks have valid start dates (use project start date as fallback)
    flatTasks.forEach(task => {
        if (!task.startDate) {
            task.startDate = projectStartDate;
        }
        // Ensure finish date is set if start date exists but finish date doesn't
        if (task.startDate && !task.finishDate && task.duration) {
            try {
                const startDate = new Date(task.startDate);
                if (!isNaN(startDate.getTime())) {
                    // Add duration days to start date
                    startDate.setDate(startDate.getDate() + Math.round(task.duration));
                    task.finishDate = startDate.toISOString().split('T')[0];
                }
            } catch {
                // If calculation fails, use project start date + duration
                task.finishDate = projectStartDate;
            }
        }
        // If finish date is still missing, use start date
        if (task.startDate && !task.finishDate) {
            task.finishDate = task.startDate;
        }
    });

    return {
        projectName: projectName || 'Exported Project',
        startDate: projectStartDate,
        tasks: flatTasks,
        resources: mspdiResources,
        assignments: mspdiAssignments
    };
}
