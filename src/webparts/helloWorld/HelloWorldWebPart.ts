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
import { getAccessToken, ensureLoginOrRedirect, isAuthenticated } from './dataverse/auth.service';
import { fetchTasksFromDataverse, fetchAssignmentsFromDataverse } from './dataverse/dataverseApi';
import { buildTaskHierarchy, buildDependencies } from './dataverse/dataTransformer';
import { dataverseConfig } from './dataverse/dataverseConfig';

const PROJECT_ID_OPTIONS = [
  { value: '', text: '-- Select Project --' },
  { value: '35d56841-1466-daad18-0e697474fdfd', text: 'OCV-SS TT' },
  { value: '35d56841-1466-47b7-bd18-0e697474fdfd', text: 'Design of SLS Dispenser' },
  { value: '613dde01-0492-f011-b41c-6045bdc5e503', text: 'Sample - Test' }
];

const API_URL = dataverseConfig.apiBaseUrl || 'http://localhost:3001/api';

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
    const isAuth = isAuthenticated();
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
        <span class="${styles.status}" id="status-${this.context.instanceId}">${!isAuth ? 'Signing you in...' : ''}</span>
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
    }

    const loadFromDataverse = async (projectIdArg?: string): Promise<void> => {
      const pid = (projectIdArg !== undefined ? projectIdArg : getSelectedProjectId()).trim();
      if (exportBtn) exportBtn.disabled = true;
      if (importBtn) importBtn.disabled = true;
      if (projectIdSelect) projectIdSelect.disabled = true;
      setStatus(pid ? 'Loading...' : '');
      try {
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
        const utcNow = new Date().toISOString().replace(/[:.]/g, '-');
        const exportFilename = selectedProjectId + '-' + utcNow + '.xml';
        const response = await fetch(API_URL + '/tasks/export/mpp?projectId=' + encodeURIComponent(selectedProjectId) + '&projectName=' + encodeURIComponent(exportFilename), {
          method: 'GET',
          headers: { Authorization: 'Bearer ' + token, Accept: 'application/xml' }
        });
        if (!response.ok) {
          const errorData = await response.json().catch(function () { return { error: 'Export failed' }; });
          throw new Error(errorData.error || 'Export failed with status ' + response.status);
        }
        const xmlContent = await response.text();
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
          const formData = new FormData();
          formData.append('file', file);
          formData.append('projectId', selectedProjectId);
          const response = await fetch(API_URL + '/tasks/import/mpp?projectId=' + encodeURIComponent(selectedProjectId), {
            method: 'POST',
            headers: { Authorization: 'Bearer ' + token },
            body: formData
          });
          if (!response.ok) {
            const errorText = await response.text();
            let errorMessage = 'Import failed with status ' + response.status;
            try {
              const errorJson = JSON.parse(errorText);
              errorMessage = errorJson.error || errorMessage;
            } catch (e) { /* ignore */ }
            throw new Error(errorMessage);
          }
          setStatus('Import completed. Reloading...');
          await loadFromDataverse(selectedProjectId);
        } catch (err: any) {
          alert('Import failed: ' + (err && err.message ? err.message : String(err)));
          setStatus(err && err.message ? err.message : 'Import failed', true);
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

    if (!isAuth) {
      ensureLoginOrRedirect().then(function (token) {
        if (token) {
          setStatus('');
          const pid = defaultProjectId.trim();
          if (pid) loadFromDataverse(pid);
        }
      });
    } else {
      setStatus('');
      const pid = defaultProjectId.trim();
      if (pid) loadFromDataverse(pid);
    }
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
