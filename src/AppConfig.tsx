import { getAccessToken } from "./services/auth.service";

/** Gantt config (vanilla @bryntum/gantt) - same shape as React wrapper props */
export type GanttProps = Record<string, unknown>;

const API_URL =
  (import.meta as any).env?.VITE_API_URL || "http://localhost:3001/api";

interface ImportProgressState {
  isOpen: boolean;
  status: 'idle' | 'uploading' | 'processing' | 'completed' | 'error';
  message: string;
  progress: {
    tasksCreated: number;
    tasksUpdated: number;
    tasksFailed: number;
    assignmentsCreated: number;
    assignmentsUpdated: number;
    assignmentsFailed: number;
    dependenciesProcessed: number;
  };
  error?: string;
}

function createImportProgressModal(): {
  show: () => void;
  hide: () => void;
  update: (state: Partial<ImportProgressState>) => void;
  element: HTMLDivElement;
} {
  // Create modal container
  const overlay = document.createElement('div');
  overlay.id = 'import-progress-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: none;
    justify-content: center;
    align-items: center;
    z-index: 10000;
  `;

  const modal = document.createElement('div');
  modal.id = 'import-progress-modal';
  modal.style.cssText = `
    background: white;
    border-radius: 8px;
    padding: 24px;
    min-width: 400px;
    max-width: 500px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  `;

  overlay.appendChild(modal);

  let currentState: ImportProgressState = {
    isOpen: false,
    status: 'idle',
    message: '',
    progress: {
      tasksCreated: 0,
      tasksUpdated: 0,
      tasksFailed: 0,
      assignmentsCreated: 0,
      assignmentsUpdated: 0,
      assignmentsFailed: 0,
      dependenciesProcessed: 0,
    }
  };

  const renderModal = () => {
    const { status, message, progress, error } = currentState;

    let statusIcon = '';
    let statusColor = '#666';
    let showSpinner = false;

    switch (status) {
      case 'uploading':
        statusIcon = '📤';
        statusColor = '#2196F3';
        showSpinner = true;
        break;
      case 'processing':
        statusIcon = '⚙️';
        statusColor = '#FF9800';
        showSpinner = true;
        break;
      case 'completed':
        statusIcon = '✅';
        statusColor = '#4CAF50';
        break;
      case 'error':
        statusIcon = '❌';
        statusColor = '#f44336';
        break;
      default:
        statusIcon = '📁';
    }

    const spinnerHtml = showSpinner ? `
      <div style="
        width: 24px;
        height: 24px;
        border: 3px solid #f3f3f3;
        border-top: 3px solid ${statusColor};
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin-right: 12px;
      "></div>
    ` : '';

    const totalTasks = progress.tasksCreated + progress.tasksUpdated;
    const totalAssignments = progress.assignmentsCreated + progress.assignmentsUpdated;

    modal.innerHTML = `
      <style>
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        .progress-row {
          display: flex;
          justify-content: space-between;
          padding: 8px 0;
          border-bottom: 1px solid #eee;
        }
        .progress-row:last-child {
          border-bottom: none;
        }
        .progress-label {
          color: #666;
        }
        .progress-value {
          font-weight: 600;
          color: #333;
        }
        .progress-value.success {
          color: #4CAF50;
        }
        .progress-value.warning {
          color: #FF9800;
        }
        .progress-value.error {
          color: #f44336;
        }
        .close-btn {
          background: #2196F3;
          color: white;
          border: none;
          padding: 10px 24px;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          margin-top: 16px;
        }
        .close-btn:hover {
          background: #1976D2;
        }
        .close-btn:disabled {
          background: #ccc;
          cursor: not-allowed;
        }
      </style>
      <div style="display: flex; align-items: center; margin-bottom: 20px;">
        ${spinnerHtml}
        <div>
          <h2 style="margin: 0 0 4px 0; font-size: 18px; color: #333;">
            ${statusIcon} Import Progress
          </h2>
          <p style="margin: 0; color: ${statusColor}; font-size: 14px;">${message}</p>
        </div>
      </div>

      ${status === 'processing' || status === 'completed' ? `
        <div style="background: #f5f5f5; border-radius: 4px; padding: 16px; margin-bottom: 16px;">
          <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #333;">Tasks</h3>
          <div class="progress-row">
            <span class="progress-label">Created:</span>
            <span class="progress-value success">${progress.tasksCreated}</span>
          </div>
          <div class="progress-row">
            <span class="progress-label">Updated:</span>
            <span class="progress-value success">${progress.tasksUpdated}</span>
          </div>
          ${progress.tasksFailed > 0 ? `
            <div class="progress-row">
              <span class="progress-label">Failed:</span>
              <span class="progress-value error">${progress.tasksFailed}</span>
            </div>
          ` : ''}
          <div class="progress-row" style="font-weight: 600;">
            <span class="progress-label">Total Processed:</span>
            <span class="progress-value">${totalTasks}</span>
          </div>
        </div>

        <div style="background: #f5f5f5; border-radius: 4px; padding: 16px; margin-bottom: 16px;">
          <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #333;">Assignments</h3>
          <div class="progress-row">
            <span class="progress-label">Created:</span>
            <span class="progress-value success">${progress.assignmentsCreated}</span>
          </div>
          <div class="progress-row">
            <span class="progress-label">Updated:</span>
            <span class="progress-value success">${progress.assignmentsUpdated}</span>
          </div>
          ${progress.assignmentsFailed > 0 ? `
            <div class="progress-row">
              <span class="progress-label">Failed:</span>
              <span class="progress-value error">${progress.assignmentsFailed}</span>
            </div>
          ` : ''}
          <div class="progress-row" style="font-weight: 600;">
            <span class="progress-label">Total Processed:</span>
            <span class="progress-value">${totalAssignments}</span>
          </div>
        </div>

        <div style="background: #f5f5f5; border-radius: 4px; padding: 16px;">
          <div class="progress-row">
            <span class="progress-label">Dependencies Processed:</span>
            <span class="progress-value">${progress.dependenciesProcessed}</span>
          </div>
        </div>
      ` : ''}

      ${error ? `
        <div style="background: #ffebee; border: 1px solid #f44336; border-radius: 4px; padding: 12px; margin-top: 16px;">
          <p style="margin: 0; color: #c62828; font-size: 13px;">${error}</p>
        </div>
      ` : ''}

      <div style="text-align: center;">
        <button class="close-btn" id="import-modal-close-btn" ${status === 'uploading' || status === 'processing' ? 'disabled' : ''}>
          ${status === 'completed' ? 'Close & Refresh' : status === 'error' ? 'Close' : 'Cancel'}
        </button>
      </div>
    `;

    // Add close button handler
    const closeBtn = modal.querySelector('#import-modal-close-btn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        if (status !== 'uploading' && status !== 'processing') {
          overlay.style.display = 'none';
          currentState.isOpen = false;
        }
      });
    }
  };

  return {
    show: () => {
      if (!document.body.contains(overlay)) {
        document.body.appendChild(overlay);
      }
      overlay.style.display = 'flex';
      currentState.isOpen = true;
      renderModal();
    },
    hide: () => {
      overlay.style.display = 'none';
      currentState.isOpen = false;
    },
    update: (state: Partial<ImportProgressState>) => {
      currentState = { ...currentState, ...state };
      if (state.progress) {
        currentState.progress = { ...currentState.progress, ...state.progress };
      }
      renderModal();
    },
    element: overlay,
  };
}

// Store token globally for synchronous access
let cachedToken: string | null = null;

// Project ID options for dropdown
const PROJECT_ID_OPTIONS = [
  { value: '', text: '-- Select Project --' },
  { value: '35d56841-1466-daad18-0e697474fdfd', text: 'OCV-SS TT' },
  { value: '35d56841-1466-47b7-bd18-0e697474fdfd', text: 'Design of SLS Dispenser' },
  { value: '613dde01-0492-f011-b41c-6045bdc5e503', text: 'Sample - Test' }
];

// Store selected project ID globally (empty by default - no auto load)
let selectedProjectId: string = '';

// Export getter/setter for project ID
export const getSelectedProjectId = () => selectedProjectId;
export const setSelectedProjectId = (id: string) => { selectedProjectId = id; };

// Function to get token from localStorage (synchronous)
function getTokenFromStorage(): string | null {
  try {
    const token = localStorage.getItem("dataverse_access_token");
    const expiryTime = localStorage.getItem("dataverse_token_expiry");

    if (!token) {
      return null;
    }

    // Check if token is expired
    if (expiryTime) {
      const expiry = parseInt(expiryTime, 10);
      if (Date.now() >= expiry) {
        // Token expired, remove it
        localStorage.removeItem("dataverse_access_token");
        localStorage.removeItem("dataverse_token_expiry");
        return null;
      }
    }

    return token;
  } catch (error) {
    console.error("Error getting token from storage:", error);
    return null;
  }
}

// Function to update cached token
export const updateCachedToken = async () => {
  cachedToken = await getAccessToken();
  return cachedToken;
};

// Initialize cached token from localStorage on module load
cachedToken = getTokenFromStorage();

// Function to ensure token is available (called before Gantt loads)
export const ensureTokenAvailable = async (): Promise<string | null> => {
  console.log("[AppConfig] Ensuring token is available...");

  // First check cached token
  let token = cachedToken || getTokenFromStorage();

  // Validate token is not expired
  if (token) {
    const expiryTime = localStorage.getItem("dataverse_token_expiry");
    if (expiryTime) {
      const expiry = parseInt(expiryTime, 10);
      const now = Date.now();
      const timeUntilExpiry = expiry - now;

      // If token expires in less than 5 minutes, refresh it
      if (timeUntilExpiry < 5 * 60 * 1000) {
        console.log("[AppConfig] Token expires soon, refreshing...");
        token = null; // Force refresh
      } else {
        console.log(`[AppConfig] Token is valid, expires in ${Math.floor(timeUntilExpiry / 1000)} seconds`);
        return token;
      }
    }
  }

  // If no token or expired, fetch from MSAL
  if (!token) {
    console.log("[AppConfig] Token not available or expired, fetching from MSAL...");
    try {
      token = await getAccessToken();
      if (token) {
        cachedToken = token;
        console.log("[AppConfig] Token retrieved and cached successfully");
      } else {
        console.error("[AppConfig] Failed to retrieve token from MSAL");
      }
    } catch (error: any) {
      console.error("[AppConfig] Failed to get token:", error);
      console.error("[AppConfig] Error details:", {
        errorCode: error?.errorCode,
        errorMessage: error?.message,
        stack: error?.stack
      });
    }
  }

  if (!token) {
    console.error("[AppConfig] No valid token available!");
  }

  return token;
};

export const useGanttProps = (_handleEditClick: Function): GanttProps => {
  // Build headers with Authorization if token is available
  // This function gets the token synchronously from localStorage
  const getAuthHeaders = () => {
    const currentToken = getTokenFromStorage();
    const headers: Record<string, string> = {
      "Content-Type": "application/json; charset=utf-8",
      "Accept": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    };

    if (currentToken) {
      headers["Authorization"] = `Bearer ${currentToken}`;
      console.log("[AppConfig] ✓ Authorization header will be sent with token");
    } else {
      console.warn("[AppConfig] ⚠ No token available for headers");
    }

    return headers;
  };

  return {
    // Enable all features properly - don't use spread pattern which can override defaults
    features: {
      // Enable tree expand/collapse (should be on by default but explicitly enable)
      tree: true,
      // Enable critical path highlighting by default
      criticalPaths: true,
      // Enable cell editing
      cellEdit: true,
      // Enable task context menu (right-click)
      taskMenu: true,
      // Enable cell context menu
      cellMenu: true,
      // Enable task editing dialog
      taskEdit: {
        // Advanced tab fields are automatically available when task data includes:
        // - calendar, ignoreResourceCalendar, schedulingMode, effortDriven
        // - constraintType, constraintDate, rollup, inactive, manuallyScheduled, projectBorder
      },
      // Enable Ctrl+C / Ctrl+V / Ctrl+X for tasks (copy/paste/cut)
      taskCopyPaste: {
        listeners: {
          // After paste, force pasted/moved tasks to inherit projectId from paste target
          paste: ({ targetCell, modifiedRecords }: any) => {
            const targetRecord = targetCell?.record;
            const targetProjectId =
              targetRecord?.projectId ||
              targetRecord?.data?.projectId ||
              targetRecord?.data?.eppm_projectid;

            if (!targetProjectId || !Array.isArray(modifiedRecords)) return;

            const applyProjectId = (rec: any) => {
              try {
                // Support both model.set() and plain assignments
                if (typeof rec?.set === "function") {
                  rec.set("projectId", targetProjectId);
                  rec.set("eppm_projectid", targetProjectId);
                } else {
                  rec.projectId = targetProjectId;
                  rec.eppm_projectid = targetProjectId;
                }

                const children = rec?.children || rec?.data?.children;
                if (Array.isArray(children)) {
                  children.forEach(applyProjectId);
                }
              } catch (e) {
                // ignore
              }
            };

            modifiedRecords.forEach(applyProjectId);
          },
        },
      },
    },
    // Allow selecting multiple rows for multi-delete
    selectionMode: {
      row: true,
      multiSelect: true,
    },
    tbar: [
      //   {
      //     type: "button",
      //     text: "Delete selected",
      //     icon: "b-fa b-fa-trash",
      //     onAction: ({ source }: any) => {
      //       const gantt = source?.up?.("gantt");
      //       const selected = gantt?.selectedRecords || [];

      //       if (!selected.length) {
      //         alert("Please select one or more tasks to delete.");
      //         return;
      //       }

      //       const ok = confirm(`Delete ${selected.length} task(s)? This will also delete from Dataverse.`);
      //       if (!ok) return;

      //       // Removing from store will trigger CrudManager sync (autoSync=true)
      //       gantt.taskStore.remove(selected);
      //     },
      //   },

      {
        type: "button",
        text: "Export to MS Project",
        icon: "b-fa b-fa-file-export",
        onAction: async () => {
          try {
            console.log("[Export] Starting MS Project export...");

            // Check if a project is selected
            if (!selectedProjectId || selectedProjectId.trim() === '') {
              alert("Please select a project first before exporting.");
              return;
            }

            // Get token from localStorage
            const token = localStorage.getItem("dataverse_access_token");
            if (!token) {
              alert("Please login first to export data.");
              return;
            }

            // Show loading indicator
            const exportBtn = document.querySelector('[data-ref="exportMppButton"]') as HTMLButtonElement;
            if (exportBtn) {
              exportBtn.textContent = "Exporting...";
              exportBtn.disabled = true;
            }

            // Generate filename: projectid-utcnow.xml
            const utcNow = new Date().toISOString().replace(/[:.]/g, '-');
            const exportFilename = `${selectedProjectId}-${utcNow}.xml`;

            // Make request to export endpoint with projectId
            const response = await fetch(`${API_URL}/tasks/export/mpp?projectId=${selectedProjectId}&projectName=${encodeURIComponent(exportFilename)}`, {
              method: "GET",
              headers: {
                Authorization: `Bearer ${token}`,
                Accept: "application/xml",
              },
            });

            if (!response.ok) {
              const errorData = await response.json().catch(() => ({ error: "Export failed" }));
              throw new Error(errorData.error || `Export failed with status ${response.status}`);
            }

            // Get the XML content
            const xmlContent = await response.text();

            // Create a blob and download
            const blob = new Blob([xmlContent], { type: "application/xml" });
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = url;
            link.download = exportFilename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.URL.revokeObjectURL(url);

            console.log("[Export] Export completed successfully:", exportFilename);
            alert(`Export completed! File: ${exportFilename}\n\nOpen this XML file with Microsoft Project to view and save as .mpp`);
          } catch (error: any) {
            console.error("[Export] Export failed:", error);
            alert(`Export failed: ${error.message}`);
          } finally {
            // Restore button state
            const exportBtn = document.querySelector('[data-ref="exportMppButton"]') as HTMLButtonElement;
            if (exportBtn) {
              exportBtn.textContent = "Export to MS Project";
              exportBtn.disabled = false;
            }
          }
        },
        ref: "exportMppButton",
      },
      {
        type: "button",
        text: "Import from MS Project",
        icon: "b-fa b-fa-file-import",
        onAction: async ({ source }: any) => {
          // Check if a project is selected
          if (!selectedProjectId || selectedProjectId.trim() === '') {
            alert("Please select a project first before importing.");
            return;
          }

          // Create progress modal
          const progressModal = createImportProgressModal();

          try {
            console.log("[Import] Starting MS Project import for project:", selectedProjectId);

            // Get token from localStorage
            const token = localStorage.getItem("dataverse_access_token");
            if (!token) {
              alert("Please login first to import data.");
              return;
            }

            // Create file input element
            const fileInput = document.createElement("input");
            fileInput.type = "file";
            fileInput.accept = ".xml";
            fileInput.style.display = "none";

            // Handle file selection
            fileInput.onchange = async (event: Event) => {
              const target = event.target as HTMLInputElement;
              const file = target.files?.[0];

              if (!file) {
                console.log("[Import] No file selected");
                return;
              }

              console.log(`[Import] File selected: ${file.name}, size: ${file.size} bytes`);

              // Confirm import
              const confirmImport = confirm(
                `Import "${file.name}" into project ${selectedProjectId}?\n\nThis will add/update tasks, resources, and assignments for the selected project.\n\n- Existing tasks (with matching IDs) will be UPDATED\n- New tasks will be CREATED with the selected project ID\n\nProceed with import?`
              );

              if (!confirmImport) {
                console.log("[Import] Import cancelled by user");
                return;
              }

              // Show progress modal
              progressModal.show();
              progressModal.update({
                status: 'uploading',
                message: `Uploading "${file.name}"...`,
              });

              // Disable import button
              const importBtn = document.querySelector('[data-ref="importMppButton"]') as HTMLButtonElement;
              if (importBtn) {
                importBtn.disabled = true;
              }

              try {
                // Create FormData and append file and projectId
                const formData = new FormData();
                formData.append("file", file);
                formData.append("projectId", selectedProjectId);

                // Update status to processing
                progressModal.update({
                  status: 'processing',
                  message: 'Uploading and processing...',
                });

                // Upload to server and read streaming response
                const response = await fetch(`${API_URL}/tasks/import/mpp?projectId=${selectedProjectId}`, {
                  method: "POST",
                  headers: {
                    Authorization: `Bearer ${token}`,
                  },
                  body: formData,
                });

                if (!response.ok) {
                  // Try to parse error as JSON
                  const errorText = await response.text();
                  let errorMessage = `Import failed with status ${response.status}`;
                  try {
                    const errorJson = JSON.parse(errorText);
                    errorMessage = errorJson.error || errorMessage;
                  } catch {
                    // Use default error message
                  }
                  throw new Error(errorMessage);
                }

                // Read streaming NDJSON response
                const reader = response.body?.getReader();
                if (!reader) {
                  throw new Error('Failed to read response stream');
                }

                const decoder = new TextDecoder();
                let buffer = '';
                let finalResult: any = null;

                while (true) {
                  const { done, value } = await reader.read();

                  if (done) {
                    break;
                  }

                  // Decode chunk and add to buffer
                  buffer += decoder.decode(value, { stream: true });

                  // Process complete lines (NDJSON format)
                  const lines = buffer.split('\n');
                  buffer = lines.pop() || ''; // Keep incomplete line in buffer

                  for (const line of lines) {
                    if (!line.trim()) continue;

                    try {
                      const data = JSON.parse(line);
                      console.log("[Import] Progress update:", data);

                      if (data.type === 'progress') {
                        progressModal.update({
                          status: 'processing',
                          message: data.message || 'Processing...',
                          progress: data.progress || {
                            tasksCreated: 0,
                            tasksUpdated: 0,
                            tasksFailed: 0,
                            assignmentsCreated: 0,
                            assignmentsUpdated: 0,
                            assignmentsFailed: 0,
                            dependenciesProcessed: 0,
                          },
                        });
                      } else if (data.type === 'complete') {
                        finalResult = data;
                        progressModal.update({
                          status: 'completed',
                          message: data.message || 'Import completed successfully!',
                          progress: data.summary || data.progress || {
                            tasksCreated: 0,
                            tasksUpdated: 0,
                            tasksFailed: 0,
                            assignmentsCreated: 0,
                            assignmentsUpdated: 0,
                            assignmentsFailed: 0,
                            dependenciesProcessed: 0,
                          },
                        });
                      } else if (data.type === 'error') {
                        throw new Error(data.error || 'Import failed');
                      }
                    } catch (parseError) {
                      console.warn("[Import] Failed to parse progress line:", line, parseError);
                    }
                  }
                }

                // Process any remaining buffer
                if (buffer.trim()) {
                  try {
                    const data = JSON.parse(buffer);
                    if (data.type === 'complete') {
                      finalResult = data;
                      progressModal.update({
                        status: 'completed',
                        message: data.message || 'Import completed successfully!',
                        progress: data.summary || data.progress || {
                          tasksCreated: 0,
                          tasksUpdated: 0,
                          tasksFailed: 0,
                          assignmentsCreated: 0,
                          assignmentsUpdated: 0,
                          assignmentsFailed: 0,
                          dependenciesProcessed: 0,
                        },
                      });
                    } else if (data.type === 'error') {
                      throw new Error(data.error || 'Import failed');
                    }
                  } catch (parseError) {
                    console.warn("[Import] Failed to parse final buffer:", buffer, parseError);
                  }
                }

                console.log("[Import] Import result:", finalResult);

                // Wait for user to close modal, then refresh
                const closeBtn = document.querySelector('#import-modal-close-btn');
                if (closeBtn) {
                  closeBtn.addEventListener('click', async () => {
                    // Reload the Gantt chart to show imported data
                    const gantt = source?.up?.("gantt");
                    if (gantt?.project) {
                      console.log("[Import] Reloading Gantt data...");
                      try {
                        await gantt.project.load();
                        console.log("[Import] Gantt data reloaded");
                      } catch (loadError) {
                        console.error("[Import] Failed to reload Gantt:", loadError);
                        window.location.reload();
                      }
                    } else {
                      // Fallback: reload the page
                      console.log("[Import] Gantt not found, reloading page...");
                      window.location.reload();
                    }
                  }, { once: true });
                }
              } catch (error: any) {
                console.error("[Import] Import failed:", error);
                progressModal.update({
                  status: 'error',
                  message: 'Import failed',
                  error: error.message,
                });
              } finally {
                // Restore button state
                const importBtn = document.querySelector('[data-ref="importMppButton"]') as HTMLButtonElement;
                if (importBtn) {
                  importBtn.disabled = false;
                }
              }
            };

            // Trigger file picker
            document.body.appendChild(fileInput);
            fileInput.click();
            document.body.removeChild(fileInput);
          } catch (error: any) {
            console.error("[Import] Error:", error);
            progressModal.update({
              status: 'error',
              message: 'Import error',
              error: error.message,
            });
          }
        },
        ref: "importMppButton",
      },
      // Spacer to push the dropdown to the right
      { type: 'widget', cls: 'b-toolbar-fill' },
      // Project ID dropdown
      {
        type: 'combo',
        ref: 'projectIdCombo',
        label: 'ProjectId',
        labelPosition: 'before',
        width: 280,
        editable: false,
        value: selectedProjectId,
        items: PROJECT_ID_OPTIONS,
        onChange: async ({ value, source }: any) => {
          console.log('[ProjectId] Selection changed to:', value);
          selectedProjectId = value || '';

          // Get the gantt instance
          const gantt = source?.up?.('gantt');
          if (!gantt?.project) return;

          const project = gantt.project;

          // 1. Suspend autoSync to prevent sync (DELETE) requests for cleared data
          //    Without this, removeAll triggers autoSync which sends DELETE for all
          //    tasks from the previous project, causing data loss and race conditions.
          project.suspendAutoSync?.();

          try {
            // 2. Clear existing data from all stores silently
            project.taskStore?.removeAll?.(true);
            project.dependencyStore?.removeAll?.(true);
            project.resourceStore?.removeAll?.(true);
            project.assignmentStore?.removeAll?.(true);

            // 3. Accept/clear all pending CRUD changes so the removed records
            //    are not queued for sync (DELETE) when autoSync resumes
            project.acceptChanges?.();

            // Only load data if a valid project is selected (not empty placeholder)
            if (value && value.trim() !== '') {
              console.log('[ProjectId] Loading Gantt data for project:', value);

              // Update the transport load URL with new projectId
              if (project.transport?.load) {
                project.transport.load.url = `${API_URL}/tasks?projectId=${value}`;
                console.log('[ProjectId] Updated transport URL:', project.transport.load.url);
              }

              // Load fresh data for the selected project
              await project.load();

              // Wait for the scheduling engine to finish processing all loaded data
              await project.commitAsync?.();

              // Accept loaded data as committed state (not pending changes to sync)
              project.acceptChanges?.();

              console.log('[ProjectId] Gantt data loaded successfully');
            } else {
              console.log('[ProjectId] No project selected - grid cleared');
            }
          } catch (error) {
            console.error('[ProjectId] Failed to load Gantt:', error);
          } finally {
            // 4. Always resume autoSync regardless of success/failure
            project.resumeAutoSync?.();
          }
        }
      },
    ],
    project: {
      autoSetConstraints: false, // Don't auto-set constraints - display as-is
      autoLoad: false, // Don't auto-load - wait for user to select a project
      // Do NOT set loadUrl/syncUrl - use transport configuration instead
      // Enable auto-sync on changes
      autoSync: true,
      // Disable effort-driven scheduling to prevent formula calculations
      effortDriven: false,
      // Use "Normal" scheduling mode (not "FixedDuration" or "FixedEffort" which require formula calculations)
      schedulingMode: 'Normal',
      // Configure assignmentStore fields (simple numeric units, no formulas)
      assignmentStore: {
        fields: [
          // Override units to be a simple number field without formula
          { name: 'units', type: 'number', defaultValue: 100, persist: true }
        ]
      },
      // Configure taskStore to ensure new tasks get proper projectId when added via default features
      taskStore: {
        // Define custom fields for tasks
        fields: [
          { name: 'taskIndex', type: 'number' },
          { name: 'projectId', type: 'string' },
          { name: 'rawStartDate', type: 'string' },
          { name: 'rawFinishDate', type: 'string' },
          { name: 'rawDuration', type: 'number' }
        ],
        listeners: {
          // When a task is added via default "Add Task Above/Below" features
          add: ({ records, source }: any) => {
            if (!Array.isArray(records) || records.length === 0) return;

            // Get the gantt instance from the store's owner
            const gantt = source?.owner?.gantt || source?.gantt;
            if (!gantt) return;

            // For each newly added task, ensure it has the correct projectId from its parent/context
            records.forEach((record: any) => {
              if (!record) return;

              // If task already has projectId, skip
              if (record.projectId || record.data?.projectId || record.data?.eppm_projectid) return;

              // Try to get projectId from parent
              const parentId = record.parentId || record.data?.parentId;
              if (parentId) {
                const parent = source.getById(parentId);
                if (parent) {
                  const parentProjectId =
                    parent.projectId ||
                    parent.data?.projectId ||
                    parent.data?.eppm_projectid;
                  if (parentProjectId) {
                    if (typeof record.set === 'function') {
                      record.set('projectId', parentProjectId);
                      record.set('eppm_projectid', parentProjectId);
                    } else {
                      record.projectId = parentProjectId;
                      record.eppm_projectid = parentProjectId;
                      if (record.data) {
                        record.data.projectId = parentProjectId;
                        record.data.eppm_projectid = parentProjectId;
                      }
                    }
                  }
                }
              } else {
                // Root task - try to get projectId from selected task
                const selected = gantt.selectedRecords || [];
                if (selected.length > 0) {
                  const selectedTask = selected[0];
                  const projectId =
                    selectedTask.projectId ||
                    selectedTask.data?.projectId ||
                    selectedTask.data?.eppm_projectid;
                  if (projectId) {
                    if (typeof record.set === 'function') {
                      record.set('projectId', projectId);
                      record.set('eppm_projectid', projectId);
                    } else {
                      record.projectId = projectId;
                      record.eppm_projectid = projectId;
                      if (record.data) {
                        record.data.projectId = projectId;
                        record.data.eppm_projectid = projectId;
                      }
                    }
                  }
                }
              }
            });
          },
        },
      },
      // This config enables response validation and dumping of found errors to the browser console.
      // It's meant to be used as a development stage helper only so please set it to false for production systems.
      validateResponse: true,
      // Custom transport with headers and fetch function to add authentication headers
      // Using both headers (for XMLHttpRequest) and fetch (for fetch API) to cover all cases
      transport: {
        load: {
          url: `${API_URL}/tasks?projectId=${selectedProjectId}`,
          method: "GET",
          // Add headers directly - works with both XMLHttpRequest and fetch
          headers: getAuthHeaders(),
          // Custom fetch function to add auth headers (as backup)
          // This intercepts all load requests and adds Authorization header
          fetch: async (url: string, options: any = {}) => {
            console.log("[Bryntum Load] ===== Starting Load Request =====");
            console.log("[Bryntum Load] Request URL received:", url);
            console.log("[Bryntum Load] Options:", options);
            console.log("[Bryntum Load] Current selectedProjectId:", selectedProjectId);

            // If no project selected, return empty response
            if (!selectedProjectId || selectedProjectId.trim() === '') {
              console.log("[Bryntum Load] No project selected - returning empty data");
              return new Response(JSON.stringify({
                success: true,
                tasks: { rows: [] },
                dependencies: { rows: [] },
                resources: { rows: [] },
                assignments: { rows: [] }
              }), {
                status: 200,
                headers: { 'Content-Type': 'application/json' }
              });
            }

            // Build URL with selected projectId
            let requestUrl = `${API_URL}/tasks?projectId=${selectedProjectId}`;
            console.log("[Bryntum Load] Using URL with projectId:", requestUrl);

            // Get token directly from localStorage first (like reference implementation)
            console.log("[Bryntum Load] Getting access token...");
            let token: string | null = localStorage.getItem('dataverse_access_token');

            // Check if token exists and is valid
            if (token) {
              const expiryTime = localStorage.getItem("dataverse_token_expiry");
              if (expiryTime) {
                const expiry = parseInt(expiryTime, 10);
                if (Date.now() >= expiry) {
                  console.log("[Bryntum Load] Token expired, refreshing...");
                  token = null;
                  localStorage.removeItem('dataverse_access_token');
                  localStorage.removeItem('dataverse_token_expiry');
                } else {
                  const timeUntilExpiry = expiry - Date.now();
                  console.log(`[Bryntum Load] ✓ Token from localStorage, expires in ${Math.floor(timeUntilExpiry / 1000)} seconds`);
                  cachedToken = token;
                }
              } else {
                console.log("[Bryntum Load] ✓ Token from localStorage (no expiry info)");
                cachedToken = token;
              }
            }

            // If no token in localStorage, get from MSAL
            if (!token) {
              console.log("[Bryntum Load] Token not in localStorage, fetching from MSAL...");
              try {
                token = await getAccessToken();
                if (token) {
                  cachedToken = token;
                  console.log("[Bryntum Load] ✓ Token retrieved from MSAL, length:", token.length);
                }
              } catch (error: any) {
                console.warn("[Bryntum Load] Failed to get token from MSAL:", error?.errorCode || error?.message);
              }
            }

            // Final fallback - try cached token
            if (!token) {
              token = cachedToken || getTokenFromStorage();
              if (token) {
                console.log("[Bryntum Load] ✓ Using cached token");
              }
            }

            // Build headers (matching reference implementation)
            const headers: Record<string, string> = {
              "Content-Type": "application/json; charset=utf-8",
              "Accept": "application/json",
              "X-Requested-With": "XMLHttpRequest",
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
            };
            console.log('[AppConfig] Token:', token);
            if (token) {
              headers["Authorization"] = `Bearer ${token}`;
              console.log("[Bryntum Load] ✓ Authorization header added");
              console.log("[Bryntum Load] Authorization header value:", `Bearer ${token.substring(0, 30)}...`);
            } else {
              console.error("[Bryntum Load] ✗ No access token available!");
              console.error("[Bryntum Load] Request URL:", requestUrl);
              console.error("[Bryntum Load] Cached token:", cachedToken ? "exists" : "null");
              console.error("[Bryntum Load] localStorage token:", localStorage.getItem("dataverse_access_token") ? "exists" : "null");
              console.error("[Bryntum Load] This will cause authentication errors.");
              // Return a rejected promise to prevent the request
              return Promise.reject(
                new Error(
                  "No access token available. Please authenticate first."
                )
              );
            }

            console.log("[Bryntum Load] Final request URL:", requestUrl);
            console.log("[Bryntum Load] Request headers:", {
              "Content-Type": headers["Content-Type"],
              "Accept": headers["Accept"],
              "Authorization": headers["Authorization"] ? "Bearer [TOKEN]" : "Missing"
            });

            // Build fetch options - ensure Authorization header is always included
            const fetchOptions: RequestInit = {
              method: "GET",
              headers: {
                ...headers,
                // Ensure Authorization is not overridden by options
                ...(options?.headers || {}),
                // Re-apply Authorization to ensure it's not lost
                "Authorization": headers["Authorization"],
              },
              credentials: "include", // Include credentials for CORS
            };

            // Merge any additional options (but preserve Authorization)
            if (options) {
              // Don't override headers completely, merge them
              if (options.headers) {
                fetchOptions.headers = {
                  ...headers,
                  ...options.headers,
                  // Always ensure Authorization is present
                  "Authorization": headers["Authorization"],
                };
              }
              // Copy other options but don't override headers (ES5-safe)
              const { headers: _, ...otherOptions } = options;
              for (const key in otherOptions) {
                if (Object.prototype.hasOwnProperty.call(otherOptions, key)) {
                  (fetchOptions as any)[key] = (otherOptions as any)[key];
                }
              }
            }

            console.log("[Bryntum Load] Making fetch request...");
            const response = await fetch(requestUrl, fetchOptions);
            console.log("[Bryntum Load] Response status:", response.status);

            return response;
          },
        },
        sync: {
          url: `${API_URL}/tasks/sync`,
          method: "POST",
          // Add headers directly - works with both XMLHttpRequest and fetch
          headers: getAuthHeaders(),
          // Custom fetch function to add auth headers (as backup)
          fetch: async (url: string, options: any = {}) => {
            console.log("[Bryntum Sync] ===== Starting Sync Request =====");
            console.log("[Bryntum Sync] Request URL received:", url);
            console.log("[Bryntum Sync] Options:", options);

            // Always use the correct absolute URL
            let requestUrl = url;
            if (
              !requestUrl ||
              requestUrl === "/" ||
              requestUrl.indexOf("http") !== 0
            ) {
              requestUrl = `${API_URL}/tasks/sync`;
              console.log(
                "[Bryntum Sync] URL was invalid/relative, using:",
                requestUrl
              );
            }

            // Ensure it's a full URL
            if (requestUrl.indexOf("/") === 0) {
              requestUrl = `${API_URL}${requestUrl}`;
            }

            // Get token directly from localStorage first (like reference implementation)
            console.log("[Bryntum Sync] Getting access token...");
            let token: string | null = localStorage.getItem('dataverse_access_token');

            // Check if token exists and is valid
            if (token) {
              const expiryTime = localStorage.getItem("dataverse_token_expiry");
              if (expiryTime) {
                const expiry = parseInt(expiryTime, 10);
                if (Date.now() >= expiry) {
                  console.log("[Bryntum Sync] Token expired, refreshing...");
                  token = null;
                  localStorage.removeItem('dataverse_access_token');
                  localStorage.removeItem('dataverse_token_expiry');
                } else {
                  console.log("[Bryntum Sync] ✓ Token from localStorage");
                  cachedToken = token;
                }
              } else {
                console.log("[Bryntum Sync] ✓ Token from localStorage (no expiry info)");
                cachedToken = token;
              }
            }

            // If no token in localStorage, get from MSAL
            if (!token) {
              console.log("[Bryntum Sync] Token not in localStorage, fetching from MSAL...");
              try {
                token = await getAccessToken();
                if (token) {
                  cachedToken = token;
                  console.log("[Bryntum Sync] ✓ Token retrieved from MSAL");
                }
              } catch (error: any) {
                console.warn("[Bryntum Sync] Failed to get token from MSAL:", error?.errorCode || error?.message);
              }
            }

            // Final fallback
            if (!token) {
              token = cachedToken || getTokenFromStorage();
            }

            const headers: Record<string, string> = {
              "Content-Type": "application/json; charset=utf-8",
              "Accept": "application/json",
              "X-Requested-With": "XMLHttpRequest",
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
            };

            if (token) {
              headers["Authorization"] = `Bearer ${token}`;
              console.log("[Bryntum Sync] ✓ Authorization header added");
            } else {
              console.error("[Bryntum Sync] ✗ No access token available!");
              return Promise.reject(
                new Error("No access token available. Please authenticate first.")
              );
            }

            // Build fetch options - ensure Authorization header is always included
            const fetchOptions: RequestInit = {
              method: "POST",
              headers: {
                ...headers,
                // Ensure Authorization is not overridden by options
                ...(options?.headers || {}),
                // Re-apply Authorization to ensure it's not lost
                "Authorization": headers["Authorization"],
              },
              credentials: "include",
              body: options?.body || JSON.stringify(options?.data || {}),
            };

            // Merge any additional options (but preserve Authorization)
            if (options) {
              // Don't override headers completely, merge them
              if (options.headers) {
                fetchOptions.headers = {
                  ...headers,
                  ...options.headers,
                  // Always ensure Authorization is present
                  "Authorization": headers["Authorization"],
                };
              }
              // Copy other options but don't override headers (ES5-safe)
              const { headers: _, body: __, ...otherOptions } = options;
              for (const key in otherOptions) {
                if (Object.prototype.hasOwnProperty.call(otherOptions, key)) {
                  (fetchOptions as any)[key] = (otherOptions as any)[key];
                }
              }
              // Preserve body if provided
              if (options.body) {
                fetchOptions.body = options.body;
              } else if (options.data) {
                fetchOptions.body = JSON.stringify(options.data);
              }
            }

            console.log("[Bryntum Sync] Making fetch request...");
            return fetch(requestUrl, fetchOptions);
          },
        },
      },
    },
    columns: [
      {
        field: "taskIndex",
        text: "ID",
        width: 60,
        align: "center",
        sortable: true,
        editor: false,  // Read-only - ID should not be edited
        renderer: ({ value }: any) => value ?? '',  // Display the value or empty string
      },
      {
        type: "wbs",
        text: "WBS",
        width: 80,
        editor: false,
      },
      {
        type: "name",
        field: "name",
        width: 250,
      },
      {
        text: "Start Date",
        field: "rawStartDate",
        width: 120,
        editor: false,
        // Display the raw start date from backend (exactly as stored in Dataverse)
        renderer: ({ record }: any) => {
          const dateStr = record?.rawStartDate || record?.data?.rawStartDate;
          if (!dateStr) return "";
          const parts = dateStr.split('-');
          if (parts.length !== 3) return dateStr;
          const year = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10);
          const day = parseInt(parts[2], 10);
          if (isNaN(year) || isNaN(month) || isNaN(day)) return dateStr;
          const displayDate = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
          return displayDate.toLocaleDateString(undefined, {
            year: "numeric",
            month: "short",
            day: "numeric",
          });
        },
      },
      {
        text: "Finish Date",
        field: "rawFinishDate",
        width: 120,
        editor: false,
        // Display the raw finish date from backend (exactly as stored in Dataverse - inclusive)
        renderer: ({ record }: any) => {
          const dateStr = record?.rawFinishDate || record?.data?.rawFinishDate;
          if (!dateStr) return "";
          const parts = dateStr.split('-');
          if (parts.length !== 3) return dateStr;
          const year = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10);
          const day = parseInt(parts[2], 10);
          if (isNaN(year) || isNaN(month) || isNaN(day)) return dateStr;
          const displayDate = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
          return displayDate.toLocaleDateString(undefined, {
            year: "numeric",
            month: "short",
            day: "numeric",
          });
        },
      },
      {
        text: "Duration",
        field: "rawDuration",
        width: 120,
        editor: false,
        // Display the raw duration from backend (exactly as stored in Dataverse)
        renderer: ({ record }: any) => {
          const duration = record?.rawDuration ?? record?.data?.rawDuration;
          if (duration == null) return "";
          return `${duration} days`;
        },
      },
      { type: "effort", text: "Effort", width: 120 },
      { type: "percentdone", text: "Percent Done", width: 120 },
      {
        type: "resourceassignment",
        text: "Resources",
        width: 160,
        showAvatars: true,
      },
      // {
      //   text: 'Edit<div class="small-text">(React component)</div>',
      //   htmlEncodeHeaderText: false,
      //   width: 120,
      //   editor: false,
      //   align: "center",
      //   type: "column",
      //   // Using custom React component
      //   renderer: ({ record, grid }: any) =>
      //     record.isLeaf ? (
      //       <DemoButton
      //         text={"Edit"}
      //         onClick={() => handleEditClick(record, grid)}
      //       />
      //     ) : (
      //       ""
      //     ),
      // },
      // {
      //   field: "draggable",
      //   text: 'Draggable<div class="small-text">(React editor)</div>',
      //   htmlEncodeHeaderText: false,
      //   width: 120,
      //   align: "center",
      //   type: "column",
      //   editor: (ref) => <DemoEditor ref={ref} />,
      //   renderer: ({ value }: any) => (value ? "Yes" : "No"),
      // },
    ],
    viewPreset: "weekAndDayLetter",
    barMargin: 10,
  };
};
