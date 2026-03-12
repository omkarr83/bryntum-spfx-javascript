import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import { Gantt } from '@bryntum/gantt';

require('@bryntum/gantt/fontawesome/css/fontawesome.css');
require('@bryntum/gantt/fontawesome/css/solid.css');
require('@bryntum/gantt/gantt.css');
require('@bryntum/gantt/svalbard-light.css');

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { getAccessToken, ensureLoginOrRedirect, isAuthenticated, processRedirectOnLoad } from './dataverse/auth.service';
import { fetchTasksFromDataverse, fetchAssignmentsFromDataverse } from './dataverse/dataverseApi';
import { buildTaskHierarchy, buildDependencies, bryntumToDataverseTask } from './dataverse/dataTransformer';
import {
  createTask,
  updateTask,
  createAssignment,
  updateAssignment,
  deleteAssignment,
  getTasksForProject,
  buildAssignmentPayload,
  isGuid
} from './dataverse/dataverseCrud';
import {
  parseMspdiXmlBrowser,
  convertImportedDataToBryntum,
  convertImportedDataToDataverse,
  buildPredecessorStringForTask,
  buildSuccessorStringForTask,
  convertToMspdiFormatFromDataverse,
  generateMspdiXmlBrowser
} from './dataverse/mspdiBrowser';

const PROJECT_ID_OPTIONS = [
  { value: '', text: '-- Select Project --' },
  { value: '35d56841-1466-daad18-0e697474fdfd', text: 'OCV-SS TT' },
  { value: '35d56841-1466-47b7-bd18-0e697474fdfd', text: 'Design of SLS Dispenser' },
  { value: '613dde01-0492-f011-b41c-6045bdc5e503', text: 'Sample - Test' }
];

export interface IHelloWorldWebPartProps {
  description: string;
  /** Default Dataverse project ID (GUID) to load tasks for */
  defaultProjectId?: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _gantt: Gantt | null = null;

  public render(): void {
    if (this._gantt) {
      this._gantt.destroy();
      this._gantt = null;
    }

    const containerId = `bryntum-gantt-${this.context.instanceId}`;
    const defaultProjectId = this.properties.defaultProjectId || '';
    const projectOptionsHtml = PROJECT_ID_OPTIONS.map(function (opt) {
      return '<option value="' + (opt.value || '').replace(/"/g, '&quot;') + '"' + (opt.value === defaultProjectId ? ' selected' : '') + '>' + (opt.text || opt.value || '').replace(/</g, '&lt;') + '</option>';
    }).join('');

    this.domElement.innerHTML = `
      <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.toolbar}">
          <div class="${styles.toolbarLeft}">
            <button type="button" class="${styles.toolbarLink}" id="btn-export-${this.context.instanceId}">Export to MS Project</button>
            <button type="button" class="${styles.toolbarLink}" id="btn-import-${this.context.instanceId}">Import from MS Project</button>
          </div>
          <div class="${styles.toolbarFill}"></div>
          <div class="${styles.toolbarRight}">
            <label class="${styles.label}" for="projectId-${this.context.instanceId}">ProjectId</label>
            <select class="${styles.input}" id="projectId-${this.context.instanceId}">${projectOptionsHtml}</select>
          </div>
        </div>
        <span class="${styles.status}" id="status-${this.context.instanceId}">Signing you in...</span>
        <div class="${styles.ganttContainer}">
          <div id="${containerId}" class="${styles.ganttHost}"></div>
        </div>
      </section>`;

    const exportBtn = this.domElement.querySelector(`#btn-export-${this.context.instanceId}`) as HTMLButtonElement;
    const importBtn = this.domElement.querySelector(`#btn-import-${this.context.instanceId}`) as HTMLButtonElement;
    const projectIdSelect = this.domElement.querySelector(`#projectId-${this.context.instanceId}`) as HTMLSelectElement;
    const statusEl = this.domElement.querySelector(`#status-${this.context.instanceId}`);
    const container = this.domElement.querySelector(`#${containerId}`) as HTMLElement;

    const setStatus = (msg: string, isError?: boolean): void => {
      if (statusEl) {
        statusEl.textContent = msg;
        (statusEl as HTMLElement).style.color = isError ? '#a4262c' : '#323130';
      }
    };

    const getSelectedProjectId = (): string => {
      return (projectIdSelect && projectIdSelect.value) ? projectIdSelect.value.trim() : '';
    };

    const formatDateDisplay = (dateStr: string): string => {
      if (!dateStr) return '';
      const parts = dateStr.split('-');
      if (parts.length !== 3) return dateStr;
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10);
      const day = parseInt(parts[2], 10);
      if (isNaN(year) || isNaN(month) || isNaN(day)) return dateStr;
      const d = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
      return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
    };

    let isLoading = false;

    const buildGanttConfig = (tasksData: any[], dependenciesData: any[]): any => ({
      appendTo: container,
      columns: [
        { field: 'taskIndex', text: 'ID', width: 60, align: 'center' as const, sortable: true, editor: false },
        { type: 'wbs', text: 'WBS', width: 80, editor: false },
        { type: 'name', field: 'name', text: 'Name', width: 250 },
        { text: 'Start Date', field: 'rawStartDate', width: 120, editor: false, renderer: function (data: any) { return formatDateDisplay(data.record?.rawStartDate || data.record?.data?.rawStartDate); } },
        { text: 'Finish Date', field: 'rawFinishDate', width: 120, editor: false, renderer: function (data: any) { return formatDateDisplay(data.record?.rawFinishDate || data.record?.data?.rawFinishDate); } },
        { text: 'Duration', field: 'rawDuration', width: 120, editor: false, renderer: function (data: any) { var v = data.record?.rawDuration ?? data.record?.data?.rawDuration; return v == null ? '' : v + ' days'; } },
        { type: 'effort', text: 'Effort', width: 120 },
        { type: 'percentdone', text: 'Percent Done', width: 120 },
        { type: 'resourceassignment', text: 'Resources', width: 160, showAvatars: true }
      ],
      project: {
        autoSetConstraints: true,
        taskStore: {
          fields: [
            { name: 'taskIndex', type: 'number' },
            { name: 'rawStartDate', type: 'string' },
            { name: 'rawFinishDate', type: 'string' },
            { name: 'rawDuration', type: 'number' }
          ]
        },
        resourceStore: { data: [] },
        assignmentStore: { data: [] },
        tasks: tasksData,
        dependencies: dependenciesData
      },
      viewPreset: 'weekAndDayLetter'
    });

    if (container) {
      this._gantt = new Gantt(buildGanttConfig([], []) as any);
      const project = (this._gantt as any).project;
      const taskStore = project && project.taskStore;
      if (taskStore && typeof taskStore.on === 'function') {
        taskStore.on('update', async (args: any) => {
          try {
            if (isLoading) return;
            const record = args && (args.record || args.records && args.records[0]);
            if (!record) return;
            let id = record.id || record.data && (record.data.id || record.data.eppm_projecttaskid);
            if (!id || !isGuid(String(id))) return;
            const patch = bryntumToDataverseTask(record.data || record);
            let token = await getAccessToken();
            if (!token) {
              token = await ensureLoginOrRedirect();
            }
            if (!token) return;
            await updateTask(token, String(id), patch);
          } catch (e) {
            // swallow errors for now
          }
        });
      }

      const dependencyStore = project && project.dependencyStore;
      if (dependencyStore && typeof dependencyStore.on === 'function' && taskStore && typeof taskStore.forEach === 'function') {
        dependencyStore.on('change', async () => {
          try {
            if (isLoading) return;
            let token = await getAccessToken();
            if (!token) {
              token = await ensureLoginOrRedirect();
            }
            if (!token) return;

            const deps: any[] = [];
            const affectedTaskIds = new Set<string>();
            dependencyStore.forEach(function (rec: any) {
              const d = rec && (rec.data || rec);
              if (!d) return;
              const fromId = d.fromTask != null ? String(d.fromTask) : '';
              const toId = d.toTask != null ? String(d.toTask) : '';
              const typeVal = typeof d.type === 'number' ? d.type : d.type == null ? 2 : Number(d.type);
              const lagVal = d.lag != null ? Number(d.lag) : 0;
              deps.push({ fromTask: fromId, toTask: toId, type: typeVal, lag: lagVal });
              if (fromId) affectedTaskIds.add(fromId);
              if (toId) affectedTaskIds.add(toId);
            });

            if (deps.length === 0 || affectedTaskIds.size === 0) {
              return;
            }

            const TYPE_TO_STR: Record<number, string> = { 0: 'SS', 1: 'SF', 2: 'FS', 3: 'FF' };

            const buildPred = function (taskId: string): string {
              const preds: any[] = [];
              for (let i = 0; i < deps.length; i++) {
                const d = deps[i];
                if (d.toTask === taskId) preds.push(d);
              }
              if (preds.length === 0) return '';
              const tokens: string[] = [];
              for (let i = 0; i < preds.length; i++) {
                const dep = preds[i];
                const fromId = dep.fromTask;
                const typeStr = TYPE_TO_STR[dep.type] || 'FS';
                let token = String(fromId) + typeStr;
                const lag = dep.lag;
                if (lag != null && lag !== 0 && !isNaN(lag)) {
                  token += (lag > 0 ? '+' : '') + lag + 'd';
                }
                tokens.push(token);
              }
              return tokens.join(';');
            };

            const buildSucc = function (taskId: string): string {
              const succs: any[] = [];
              for (let i = 0; i < deps.length; i++) {
                const d = deps[i];
                if (d.fromTask === taskId) succs.push(d);
              }
              if (succs.length === 0) return '';
              const tokens: string[] = [];
              for (let i = 0; i < succs.length; i++) {
                const dep = succs[i];
                const toId = dep.toTask;
                const typeStr = TYPE_TO_STR[dep.type] || 'FS';
                let token = String(toId) + typeStr;
                const lag = dep.lag;
                if (lag != null && lag !== 0 && !isNaN(lag)) {
                  token += (lag > 0 ? '+' : '') + lag + 'd';
                }
                tokens.push(token);
              }
              return tokens.join(';');
            };

            const tasksToUpdate: Array<{ id: string; payload: Record<string, unknown> }> = [];
            taskStore.forEach(function (rec: any) {
              const r = rec && (rec.data || rec);
              if (!r) return;
              const rawId = rec.id || r.id || r.eppm_projecttaskid;
              const idStr = rawId != null ? String(rawId) : '';
              if (!idStr || !isGuid(idStr) || !affectedTaskIds.has(idStr)) return;
              const predStr = buildPred(idStr);
              const succStr = buildSucc(idStr);
              const payload: Record<string, unknown> = {
                eppm_predecessor: predStr ? predStr : null,
                eppm_successors: succStr ? succStr : null
              };
              tasksToUpdate.push({ id: idStr, payload: payload });
            });

            for (let i = 0; i < tasksToUpdate.length; i++) {
              const t = tasksToUpdate[i];
              await updateTask(token, t.id, t.payload);
            }
          } catch (e) {
            // swallow errors for now
          }
        });
      }

      const assignmentStore = project && project.assignmentStore;
      if (assignmentStore && typeof assignmentStore.on === 'function') {
        const syncAssignmentRecords = async function (recordsArg: any, action: 'add' | 'update' | 'remove'): Promise<void> {
          if (isLoading) return;
          const records = Array.isArray(recordsArg) ? recordsArg : (recordsArg ? [recordsArg] : []);
          if (!records.length) return;
          let token = await getAccessToken();
          if (!token) {
            token = await ensureLoginOrRedirect();
          }
          if (!token) return;
          const projectIdForAssignments = getSelectedProjectId().trim();
          if (!projectIdForAssignments) return;
          for (let i = 0; i < records.length; i++) {
            const rec = records[i];
            const r = rec && (rec.data || rec);
            if (!r) continue;
            const taskId = r.event != null ? String(r.event) : '';
            if (!taskId || !isGuid(taskId)) continue;
            const resourceId = r.resource != null ? String(r.resource) : '';
            if (!resourceId) continue;
            const unitsVal = r.units != null ? Number(r.units) : 100;
            if (action === 'remove') {
              const assignId = rec.id != null ? String(rec.id) : '';
              if (assignId && isGuid(assignId)) {
                try {
                  await deleteAssignment(token, assignId);
                } catch (e) {
                  // ignore delete errors
                }
              }
              continue;
            }
            const payload = buildAssignmentPayload(projectIdForAssignments, taskId, resourceId, unitsVal);
            const existingId = rec.id != null ? String(rec.id) : '';
            if (existingId && isGuid(existingId)) {
              try {
                await updateAssignment(token, existingId, payload);
              } catch (updateErr) {
                const msg = updateErr instanceof Error ? updateErr.message || '' : String(updateErr);
                if (msg.indexOf('Does Not Exist') !== -1 || msg.indexOf('does not exist') !== -1) {
                  const created = await createAssignment(token, payload);
                  const createdId = created && created.eppm_taskassignmentsid;
                  if (createdId && typeof rec.set === 'function') {
                    rec.set('id', createdId);
                  }
                } else {
                  // rethrow unexpected errors
                  throw updateErr;
                }
              }
            } else {
              const createdOnAdd = await createAssignment(token, payload);
              const createdIdOnAdd = createdOnAdd && createdOnAdd.eppm_taskassignmentsid;
              if (createdIdOnAdd && typeof rec.set === 'function') {
                rec.set('id', createdIdOnAdd);
              }
            }
          }
        };

        assignmentStore.on('add', function (args: any) {
          const recs = args && (args.records || args.record);
          syncAssignmentRecords(recs, 'add');
        });
        assignmentStore.on('update', function (args: any) {
          const rec = args && (args.record || args.records && args.records[0]);
          syncAssignmentRecords(rec, 'update');
        });
        assignmentStore.on('remove', function (args: any) {
          const recs = args && (args.records || args.record);
          syncAssignmentRecords(recs, 'remove');
        });
      }
    }

    const loadFromDataverse = async (projectIdArg?: string): Promise<void> => {
      const pid = (projectIdArg !== undefined ? projectIdArg : getSelectedProjectId()).trim();
      if (exportBtn) exportBtn.disabled = true;
      if (importBtn) importBtn.disabled = true;
      if (projectIdSelect) projectIdSelect.disabled = true;
      setStatus(pid ? 'Loading...' : '');
      try {
        isLoading = true;
        if (!pid) {
          if (this._gantt && this._gantt.project) {
            const project = this._gantt.project;
            if (project.taskStore && typeof project.taskStore.removeAll === 'function') {
              project.taskStore.removeAll(true);
            }
            if (project.dependencyStore && typeof project.dependencyStore.removeAll === 'function') {
              project.dependencyStore.removeAll(true);
            }
            if (project.resourceStore && typeof project.resourceStore.removeAll === 'function') {
              project.resourceStore.removeAll(true);
            }
            if (project.assignmentStore && typeof project.assignmentStore.removeAll === 'function') {
              project.assignmentStore.removeAll(true);
            }
          }
          setStatus('');
          return;
        }
        let token = await getAccessToken();
        if (!token) {
          token = await ensureLoginOrRedirect();
          if (!token) {
            setStatus('Redirecting to sign in...', false);
            return;
          }
        }
        const tasks = await fetchTasksFromDataverse(token, pid);
        const hierarchicalTasks = buildTaskHierarchy(tasks);
        const dependencies = buildDependencies(tasks);

        const projectTaskIds = new Set<string>();
        tasks.forEach(function (t) {
          if (t.eppm_projecttaskid) projectTaskIds.add(t.eppm_projecttaskid);
        });

        const { resources, assignments } = await fetchAssignmentsFromDataverse(token, projectTaskIds);

        if (this._gantt && this._gantt.project) {
          const project = this._gantt.project;
          // Let Bryntum ProjectModel handle updating its own stores
          if (typeof (project as any).loadInlineData === 'function') {
            await (project as any).loadInlineData({
              tasksData: hierarchicalTasks,
              dependenciesData: dependencies,
              resourcesData: resources,
              assignmentsData: assignments
            });
          }
          if (this._gantt && typeof (this._gantt as any).zoomToFit === 'function') {
            (this._gantt as any).zoomToFit();
          }
          setStatus('Loaded ' + tasks.length + ' task(s), ' + resources.length + ' resource(s), ' + assignments.length + ' assignment(s).');
        } else {
          setStatus('Gantt not ready.', true);
        }
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        setStatus(message, true);
      } finally {
        isLoading = false;
        if (exportBtn) exportBtn.disabled = false;
        if (importBtn) importBtn.disabled = false;
        if (projectIdSelect) projectIdSelect.disabled = false;
      }
    };

    const handleExport = async (): Promise<void> => {
      const selectedProjectId = getSelectedProjectId();
      if (!selectedProjectId) {
        alert('Please select a project first before exporting.');
        return;
      }
      try {
        let token = await getAccessToken();
        if (!token) {
          token = await ensureLoginOrRedirect();
          if (!token) {
            alert('Please sign in first to export data.');
            return;
          }
        }
        if (exportBtn) {
          exportBtn.textContent = 'Exporting...';
          exportBtn.disabled = true;
        }
        const tasks = await fetchTasksFromDataverse(token, selectedProjectId);
        const projectTaskIds = new Set<string>();
        tasks.forEach(function (t: { eppm_projecttaskid?: string }) {
          if (t.eppm_projecttaskid) projectTaskIds.add(t.eppm_projecttaskid);
        });
        const { resources, assignments } = await fetchAssignmentsFromDataverse(token, projectTaskIds);
        const dependencies = buildDependencies(tasks);
        const depsForExport = dependencies.map(function (d) {
          return {
            fromTask: d.fromTask != null ? String(d.fromTask) : undefined,
            toTask: d.toTask != null ? String(d.toTask) : undefined,
            type: d.type,
            lag: d.lag
          };
        });
        const mspdiData = convertToMspdiFormatFromDataverse(
          tasks as unknown as Array<Record<string, unknown>>,
          resources,
          assignments,
          depsForExport,
          'Exported Project'
        );
        const xmlContent = generateMspdiXmlBrowser(mspdiData);
        const utcNow = new Date().toISOString().replace(/[:.]/g, '-');
        const exportFilename = selectedProjectId + '-' + utcNow + '.xml';
        const blob = new Blob([xmlContent], { type: 'application/xml' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = exportFilename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        alert('Export completed! File: ' + exportFilename + '\n\nOpen this XML file with Microsoft Project to view and save as .mpp');
      } catch (err: any) {
        alert('Export failed: ' + (err && err.message ? err.message : String(err)));
      } finally {
        if (exportBtn) {
          exportBtn.textContent = 'Export to MS Project';
          exportBtn.disabled = false;
        }
      }
    };

    const handleImport = (): void => {
      const selectedProjectId = getSelectedProjectId();
      if (!selectedProjectId) {
        alert('Please select a project first before importing.');
        return;
      }
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = '.xml';
      fileInput.style.display = 'none';
      fileInput.onchange = async function (event: Event) {
        const target = event.target as HTMLInputElement;
        const file = target.files && target.files[0];
        if (!file) return;
        const ok = confirm('Import "' + file.name + '" into project?\n\nThis will add/update tasks for the selected project. Proceed?');
        if (!ok) return;
        try {
          let token = await getAccessToken();
          if (!token) {
            token = await ensureLoginOrRedirect();
            if (!token) {
              alert('Please sign in first to import data.');
              return;
            }
          }
          if (importBtn) importBtn.disabled = true;
          setStatus('Importing...');
          const xmlContent = await file.text();
          const importedData = parseMspdiXmlBrowser(xmlContent);
          const bryntumData = convertImportedDataToBryntum(importedData);
          const dataverseData = convertImportedDataToDataverse(bryntumData, selectedProjectId);
          const importIdToDataverseId = new Map<string, string>();
          const getLevel = (importId: unknown): number => {
            for (let i = 0; i < bryntumData.tasks.length; i++) {
              if (bryntumData.tasks[i].id === importId) return (bryntumData.tasks[i]._outlineLevel as number) || 0;
            }
            return 0;
          };
          const sortedTasks = dataverseData.tasks.slice().sort(function (a, b) {
            return getLevel(a._importId) - getLevel(b._importId);
          });
          for (const task of sortedTasks) {
            const taskPayload: Record<string, unknown> = {
              eppm_name: task.eppm_name,
              eppm_startdate: task.eppm_startdate,
              eppm_finishdate: task.eppm_finishdate,
              eppm_taskduration: task.eppm_taskduration,
              eppm_pocpercentage: task.eppm_pocpercentage,
              eppm_taskwork: task.eppm_taskwork,
              eppm_notes: task.eppm_notes,
              eppm_projectid: selectedProjectId,
              eppm_taskindex: task.eppm_taskindex
            };
            if (task._parentImportId) {
              const parentId = importIdToDataverseId.get(task._parentImportId as string);
              if (parentId) taskPayload.eppm_parenttaskid = parentId;
            }
            const existingTaskId = task._dataverseTaskId as string | undefined;
            if (existingTaskId && isGuid(existingTaskId)) {
              try {
                await updateTask(token, existingTaskId, taskPayload);
                importIdToDataverseId.set(task._importId as string, existingTaskId);
              } catch (updateErr) {
                const msg = updateErr instanceof Error ? updateErr.message || '' : String(updateErr);
                if (msg.indexOf('Does Not Exist') !== -1 || msg.indexOf('does not exist') !== -1) {
                  const createdOnUpdateFail = await createTask(token, taskPayload);
                  const createdIdOnUpdateFail = createdOnUpdateFail && createdOnUpdateFail.eppm_projecttaskid;
                  if (createdIdOnUpdateFail) {
                    importIdToDataverseId.set(task._importId as string, createdIdOnUpdateFail);
                  }
                } else {
                  throw updateErr;
                }
              }
            } else {
              const created = await createTask(token, taskPayload);
              const createdId = created && created.eppm_projecttaskid;
              if (createdId) importIdToDataverseId.set(task._importId as string, createdId);
            }
          }
          for (const task of sortedTasks) {
            const dataverseId = importIdToDataverseId.get(task._importId as string);
            if (!dataverseId) continue;
            const predStr = buildPredecessorStringForTask(task._importId as string, bryntumData.dependencies, importIdToDataverseId);
            const succStr = buildSuccessorStringForTask(task._importId as string, bryntumData.dependencies, importIdToDataverseId);
            if (predStr || succStr) {
              const updatePayload: Record<string, unknown> = {};
              if (predStr) updatePayload.eppm_predecessor = predStr;
              if (succStr) updatePayload.eppm_successors = succStr;
              await updateTask(token, dataverseId, updatePayload);
            }
          }
          const taskNameToDataverseId = new Map<string, string>();
          const freshTasks = await getTasksForProject(token, selectedProjectId);
          freshTasks.forEach(function (t) {
            if (t.eppm_projecttaskid && t.eppm_name) taskNameToDataverseId.set(t.eppm_name.toLowerCase().trim(), t.eppm_projecttaskid);
          });
          for (const assignment of dataverseData.assignments) {
            let taskNameForAssign = '';
            for (let ti = 0; ti < dataverseData.tasks.length; ti++) {
              if (dataverseData.tasks[ti]._importId === assignment.taskImportId) {
                taskNameForAssign = (dataverseData.tasks[ti].eppm_name as string) || '';
                break;
              }
            }
            const taskDataverseId = importIdToDataverseId.get(assignment.taskImportId as string) ||
              taskNameToDataverseId.get(taskNameForAssign.toLowerCase().trim());
            if (!taskDataverseId || !isGuid(taskDataverseId)) continue;
            const units = (assignment.units as number) ?? 100;
            const payload = buildAssignmentPayload(selectedProjectId, taskDataverseId, assignment.resourceEmail as string, units, assignment.startDate as string, assignment.finishDate as string);
            const existingAssignId = assignment._dataverseAssignmentId as string | undefined;
            if (existingAssignId && isGuid(existingAssignId)) {
              try {
                await updateAssignment(token, existingAssignId, payload);
              } catch (updateAssignErr) {
                const msg = updateAssignErr instanceof Error ? updateAssignErr.message || '' : String(updateAssignErr);
                if (msg.indexOf('Does Not Exist') !== -1 || msg.indexOf('does not exist') !== -1) {
                  await createAssignment(token, payload);
                } else {
                  throw updateAssignErr;
                }
              }
            } else {
              await createAssignment(token, payload);
            }
          }
          setStatus('Import completed. Reloading...');
          await loadFromDataverse(selectedProjectId);
        } catch (err: any) {
          const message = err && err.message ? err.message : String(err);
          alert('Import failed: ' + message);
          setStatus(message, true);
        } finally {
          if (importBtn) importBtn.disabled = false;
          target.value = '';
        }
      };
      document.body.appendChild(fileInput);
      fileInput.click();
      document.body.removeChild(fileInput);
    };

    if (exportBtn) exportBtn.addEventListener('click', function () { handleExport(); });
    if (importBtn) importBtn.addEventListener('click', function () { handleImport(); });
    if (projectIdSelect) {
      projectIdSelect.addEventListener('change', function () {
        loadFromDataverse(projectIdSelect.value.trim());
      });
    }

    // Process MSAL redirect first (when returning from Microsoft login), then check auth.
    processRedirectOnLoad().then(function () {
      const isAuth = isAuthenticated();
      if (!isAuth) {
        ensureLoginOrRedirect().then(function (token) {
          if (token) {
            setStatus('');
            const pid = defaultProjectId.trim();
            if (pid) loadFromDataverse(pid);
          } else {
            setStatus('Redirecting to sign in...');
          }
        }).catch(function (err: unknown) {
          const msg = err instanceof Error ? err.message : String(err);
          setStatus('Sign-in error: ' + msg, true);
        });
      } else {
        setStatus('');
        const pid = defaultProjectId.trim();
        if (pid) loadFromDataverse(pid);
      }
    });
  }

  public dispose(): void {
    if (this._gantt) {
      this._gantt.destroy();
      this._gantt = null;
    }
    super.dispose();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', { label: strings.DescriptionFieldLabel }),
                PropertyPaneTextField('defaultProjectId', {
                  label: 'Default Project ID (Dataverse)',
                  description: 'GUID of the project to load tasks for (e.g. 35d56841-1466-47b7-bd18-0e697474fdfd)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
