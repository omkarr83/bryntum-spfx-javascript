import axios, { AxiosInstance, AxiosError } from "axios";
import * as https from "https";
import { dataverseConfig } from "../config/dataverse.config.js";

// Retry configuration
const RETRY_CONFIG = {
  maxRetries: 3,
  baseDelayMs: 1000,
  maxDelayMs: 10000,
  retryableErrors: ['ETIMEDOUT', 'ECONNRESET', 'ECONNREFUSED', 'ENOTFOUND', 'EAI_AGAIN'],
  retryableStatusCodes: [408, 429, 500, 502, 503, 504],
};

/**
 * Sleep for a specified duration
 */
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Check if an error is retryable
 */
function isRetryableError(error: any): boolean {
  // Network errors (ETIMEDOUT, ECONNRESET, etc.)
  if (error.code && RETRY_CONFIG.retryableErrors.includes(error.code)) {
    return true;
  }

  // HTTP status codes that are retryable
  if (error.response?.status && RETRY_CONFIG.retryableStatusCodes.includes(error.response.status)) {
    return true;
  }

  // Axios timeout
  if (error.code === 'ECONNABORTED' || error.message?.includes('timeout')) {
    return true;
  }

  return false;
}

/**
 * Execute a function with retry logic and exponential backoff
 */
async function withRetry<T>(
  fn: () => Promise<T>,
  operationName: string,
  maxRetries: number = RETRY_CONFIG.maxRetries
): Promise<T> {
  let lastError: any;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (error: any) {
      lastError = error;

      if (attempt < maxRetries && isRetryableError(error)) {
        // Calculate delay with exponential backoff and jitter
        const baseDelay = RETRY_CONFIG.baseDelayMs * Math.pow(2, attempt);
        const jitter = Math.random() * 1000;
        const delay = Math.min(baseDelay + jitter, RETRY_CONFIG.maxDelayMs);

        console.warn(
          `[DataverseService] ${operationName} failed (attempt ${attempt + 1}/${maxRetries + 1}): ${error.code || error.message}. Retrying in ${Math.round(delay)}ms...`
        );

        await sleep(delay);
      } else {
        // Non-retryable error or max retries exceeded
        break;
      }
    }
  }

  throw lastError;
}

export interface DataverseTask {
  eppm_projecttaskid?: string;
  eppm_projectid?: string;
  eppm_name?: string;
  eppm_startdate?: string;
  eppm_finishdate?: string;
  eppm_taskduration?: number;
  eppm_pocpercentage?: number;
  eppm_taskwork?: number;
  eppm_predecessor?: string;
  eppm_successors?: string;
  eppm_notes?: string;
  eppm_parenttaskid?: string;
  // Advanced features fields
  eppm_calendarname?: string; // Calendar ID or name
  eppm_ignoreresourcecalendar?: boolean; // Ignore resource calendar toggle
  eppm_schedulingmode?: number; // Option Set: 100000000='Normal', 100000001='Fixed Duration', 100000002='Fixed Units', 100000003='Fixed Efforts'
  eppm_effortdriven?: boolean; // Effort driven toggle
  eppm_constrainttype?: number; // Option Set: 100000000='As soon as possible', 100000001='As late as possible', 100000002='Must start on', 100000003='Must finish on', 100000004='Start no earlier than', 100000005='Start no later than', 100000006='Finish no earlier than', 100000007='Finish no later than'
  eppm_constraintdate?: string; // Constraint date
  eppm_rollup?: boolean; // Rollup toggle
  eppm_inactive?: boolean; // Inactive toggle
  eppm_manuallyscheduled?: boolean; // Manually scheduled toggle
  eppm_projectborder?: string; // 'honor' | 'ignore' | 'askuser'
  eppm_resources?: string; // JSON array of {name, units} stored when assignments change
  [key: string]: any;
}

// export interface DataverseDependency {
//     eppm_projectdependencyid?: string;
//     eppm_fromtaskid?: string;
//     eppm_totaskid?: string;
//     eppm_lag?: number;
// }

export class DataverseService {
  private apiUrl: string;
  private axiosInstance: AxiosInstance;
  private accessToken: string | null = null;

  constructor(accessToken?: string) {
    // Construct the API URL
    // Remove trailing slash and construct full API URL
    const orgUrl = dataverseConfig.environmentUrl.replace(/\/$/, "");
    const tableName = dataverseConfig.tableName;
    this.apiUrl = `${orgUrl}/api/data/v9.2/${tableName}`;
    this.accessToken = accessToken || null;

    // If you're on a corporate network with TLS inspection, Node may reject the
    // proxy's certificate chain as "self-signed". You can enable a dev-only
    // bypass with DATAVERSE_ALLOW_SELF_SIGNED_CERT=true.
    //
    // Recommended (more secure) alternative: install your corporate root CA and
    // point Node to it via NODE_EXTRA_CA_CERTS.
    const allowSelfSignedCert =
      process.env.DATAVERSE_ALLOW_SELF_SIGNED_CERT === "true";

    if (allowSelfSignedCert) {
      console.warn(
        "[DataverseService] WARNING: TLS certificate verification is DISABLED for Dataverse requests (DATAVERSE_ALLOW_SELF_SIGNED_CERT=true). Do not use this in production."
      );
    }

    // Create axios instance with default config
    this.axiosInstance = axios.create({
      baseURL: `${orgUrl}/api/data/v9.2`,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Accept: "application/json",
      },
      // Timeout settings - 30 seconds for connection, 60 seconds for response
      timeout: 60000,
      // Dev-only escape hatch for environments that inject a self-signed cert
      httpsAgent: allowSelfSignedCert
        ? new https.Agent({ rejectUnauthorized: false, keepAlive: true })
        : new https.Agent({ keepAlive: true }),
    });

    // Add request interceptor to inject access token and ensure headers are set
    this.axiosInstance.interceptors.request.use((config) => {
      // Ensure Authorization header is set
      if (this.accessToken) {
        config.headers.Authorization = `Bearer ${this.accessToken}`;
      }
      // Ensure all required headers are present
      config.headers["OData-MaxVersion"] = "4.0";
      config.headers["OData-Version"] = "4.0";
      config.headers["Accept"] = "application/json";
      // Set Content-Type with charset for requests with body
      if (config.data && !config.headers["Content-Type"]) {
        config.headers["Content-Type"] = "application/json; charset=utf-8";
      }
      return config;
    });
  }

  /**
   * Set access token (for MSAL authentication from frontend)
   */
  setAccessToken(token: string) {
    this.accessToken = token;
  }

  /**
   * Get headers for Dataverse API requests
   */
  private getHeaders(): Record<string, string> {
    const headers: Record<string, string> = {
      Authorization: this.accessToken ? `Bearer ${this.accessToken}` : "",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
    };

    // if (this.accessToken) {
    //   headers["Authorization"] = `Bearer ${this.accessToken}`;
    // }

    return headers;
  }

  /**
   * Get all tasks from Dataverse (with retry logic)
   */
  async getAllTasks(TASK_FILTER?: string | undefined): Promise<DataverseTask[]> {
    const tableName = dataverseConfig.tableName;
    const selectFields = [
      "eppm_projecttaskid",
      "eppm_projectid",
      "eppm_name",
      "eppm_startdate",
      "eppm_finishdate",
      "eppm_taskduration",
      "eppm_taskindex",
      "eppm_pocpercentage",
      "eppm_taskwork",
      "eppm_predecessor",
      "eppm_successors",
      "eppm_notes",
      "eppm_parenttaskid",
      "eppm_calendarname",
      "eppm_ignoreresourcecalendar",
      "eppm_schedulingmode",
      "eppm_effortdriven",
      "eppm_constrainttype",
      "eppm_constraintdate",
      "eppm_rollup",
      "eppm_inactive",
      "eppm_manuallyscheduled",
      "eppm_projectborder",
    ].join(",");

    let url = `/${tableName}?$select=${selectFields}`;
    if (TASK_FILTER && TASK_FILTER.trim()) {
      url += `&$filter=${encodeURIComponent(TASK_FILTER)}`;
    }

    return withRetry(
      async () => {
        const response = await this.axiosInstance.get(url, {
          headers: this.getHeaders(),
        });

        // Handle OData response format
        if (response.data && response.data.value) {
          return response.data.value;
        } else if (Array.isArray(response.data)) {
          return response.data;
        } else {
          console.warn("Unexpected response format:", response.data);
          return [];
        }
      },
      `getAllTasks(${TASK_FILTER || 'no filter'})`
    ).catch((error: any) => {
      console.error(
        "[DataverseService] Error fetching tasks from Dataverse:",
        error.response?.data || error.message
      );
      console.error("[DataverseService] Error status:", error.response?.status);
      console.error("[DataverseService] Error headers:", error.response?.headers);
      console.error("[DataverseService] Request URL:", this.apiUrl);
      console.error("[DataverseService] Token present:", !!this.accessToken);
      console.error("[DataverseService] Token length:", this.accessToken?.length || 0);

      if (error.response?.status === 401) {
        console.error("[DataverseService] 401 Unauthorized - Token may be expired or invalid");
        console.error("[DataverseService] Dataverse error details:", JSON.stringify(error.response?.data, null, 2));
        throw new Error(
          "Authentication failed. Please check your access token. Token may be expired or invalid."
        );
      }
      throw new Error(`Failed to fetch tasks: ${error.message}`);
    });
  }

  /**
   * Generic: fetch any Dataverse table/entityset (with retry logic)
   * Example entitySet: "eppm_taskassignmentses"
   */
  async getTableRows<T = any>(
    entitySet: string,
    queryString?: string
  ): Promise<T[]> {
    const url = queryString ? `/${entitySet}?${queryString}` : `/${entitySet}`;

    return withRetry(
      async () => {
        const response = await this.axiosInstance.get(url, {
          headers: this.getHeaders(),
        });

        if (response.data && response.data.value) return response.data.value as T[];
        if (Array.isArray(response.data)) return response.data as T[];
        return [];
      },
      `getTableRows(${entitySet})`
    ).catch((error: any) => {
      console.error(
        `[DataverseService] Error fetching table '${entitySet}':`,
        error.response?.data || error.message
      );
      throw new Error(`Failed to fetch '${entitySet}': ${error.message}`);
    });
  }

  /**
   * Generic: fetch a single row by id from any Dataverse entitySet
   */
  async getRowById<T = any>(entitySet: string, rowId: string): Promise<T | null> {
    try {
      const response = await this.axiosInstance.get(`/${entitySet}(${rowId})`, {
        headers: this.getHeaders(),
      });
      return response.data as T;
    } catch (error: any) {
      if (error.response?.status === 404) return null;
      console.error(
        `[DataverseService] Error fetching row from '${entitySet}':`,
        error.response?.data || error.message
      );
      throw error;
    }
  }

  /**
   * Generic: patch a row in any Dataverse entitySet by id (with retry logic)
   */
  async patchRow(entitySet: string, rowId: string, patch: Record<string, any>): Promise<void> {
    await withRetry(
      async () => {
        await this.axiosInstance.patch(`/${entitySet}(${rowId})`, patch, {
          headers: {
            ...this.getHeaders(),
            "If-Match": "*",
          },
        });
      },
      `patchRow(${entitySet}, ${rowId})`
    );
  }

  /**
   * Generic: create a row in any Dataverse entitySet (with retry logic).
   * Returns the created representation (when Dataverse honors Prefer header).
   */
  async createRow<T = any>(entitySet: string, data: Record<string, any>): Promise<T> {
    return withRetry(
      async () => {
        const response = await this.axiosInstance.post(`/${entitySet}`, data, {
          headers: {
            ...this.getHeaders(),
            Prefer: "return=representation",
          },
        });
        return response.data as T;
      },
      `createRow(${entitySet})`
    ).catch((error: any) => {
      console.error(
        `[DataverseService] Error creating row in '${entitySet}':`,
        error.response?.data || error.message
      );
      throw error;
    });
  }

  /**
   * Generic: delete a row in any Dataverse entitySet by id
   */
  async deleteRow(entitySet: string, rowId: string): Promise<void> {
    await this.axiosInstance.delete(`/${entitySet}(${rowId})`, {
      headers: {
        ...this.getHeaders(),
        "If-Match": "*",
      },
    });
  }

  /**
   * Get a single task by ID
   */
  async getTaskById(taskId: string): Promise<DataverseTask> {
    try {
      const tableName = dataverseConfig.tableName;
      const response = await this.axiosInstance.get(
        `/${tableName}(${taskId})`,
        {
          headers: this.getHeaders(),
        }
      );
      return response.data;
    } catch (error: any) {
      console.error(
        "Error fetching task from Dataverse:",
        error.response?.data || error.message
      );
      if (error.response?.status === 401) {
        throw new Error(
          "Authentication failed. Please check your access token."
        );
      }
      throw new Error(`Failed to fetch task: ${error.message}`);
    }
  }

  /**
   * Create a new task in Dataverse (with retry logic)
   */
  async createTask(task: Partial<DataverseTask>): Promise<DataverseTask> {
    const tableName = dataverseConfig.tableName;

    return withRetry(
      async () => {
        const response = await this.axiosInstance.post(`/${tableName}`, task, {
          headers: this.getHeaders(),
        });
        return response.data;
      },
      `createTask(${task.eppm_name || 'unnamed'})`
    ).catch((error: any) => {
      console.error(
        "Error creating task in Dataverse:",
        error.response?.data || error.message
      );
      throw new Error(`Failed to create task: ${error.message}`);
    });
  }

  /**
   * Update an existing task in Dataverse (with retry logic)
   */
  async updateTask(
    taskId: string,
    task: Partial<DataverseTask>
  ): Promise<void> {
    const tableName = dataverseConfig.tableName;

    await withRetry(
      async () => {
        await this.axiosInstance.patch(`/${tableName}(${taskId})`, task, {
          headers: {
            ...this.getHeaders(),
            // Avoid needing to fetch ETag first
            "If-Match": "*",
          },
        });
      },
      `updateTask(${taskId})`
    ).catch((error: any) => {
      console.error(
        "Error updating task in Dataverse:",
        error.response?.data || error.message
      );
      throw new Error(`Failed to update task: ${error.message}`);
    });
  }

  /**
   * Delete a task from Dataverse
   */
  async deleteTask(taskId: string): Promise<void> {
    try {
      const tableName = dataverseConfig.tableName;
      await this.axiosInstance.delete(`/${tableName}(${taskId})`, {
        headers: {
          ...this.getHeaders(),
          // Avoid needing to fetch ETag first
          "If-Match": "*",
        },
      });
    } catch (error: any) {
      console.error(
        "Error deleting task from Dataverse:",
        error.response?.data || error.message
      );
      throw new Error(`Failed to delete task: ${error.message}`);
    }
  }

  /**
   * Batch create/update/delete operations
   */
  async batchOperation(
    operations: Array<{
      method: "POST" | "PATCH" | "DELETE";
      url: string;
      data?: any;
    }>
  ): Promise<any[]> {
    try {
      const batchId = `batch_${Date.now()}`;
      const changesetId = `changeset_${Date.now()}`;

      const boundary = `batch_${batchId}`;
      const changesetBoundary = `changeset_${changesetId}`;

      let batchBody = "";

      operations.forEach((operation, index) => {
        batchBody += `--${boundary}\r\n`;
        batchBody += `Content-Type: multipart/mixed; boundary="${changesetBoundary}"\r\n\r\n`;
        batchBody += `--${changesetBoundary}\r\n`;
        batchBody += `Content-Type: application/http\r\n`;
        batchBody += `Content-Transfer-Encoding: binary\r\n\r\n`;
        batchBody += `${operation.method} ${this.apiUrl}${operation.url} HTTP/1.1\r\n`;
        batchBody += `Content-Type: application/json\r\n\r\n`;

        if (operation.data) {
          batchBody += JSON.stringify(operation.data) + "\r\n";
        }

        batchBody += `--${changesetBoundary}--\r\n`;
      });

      batchBody += `--${boundary}--\r\n`;

      const batchHeaders = this.getHeaders();
      batchHeaders["Content-Type"] = `multipart/mixed; boundary=${boundary}`;

      const response = await this.axiosInstance.post("/$batch", batchBody, {
        headers: batchHeaders,
      });

      return response.data;
    } catch (error: any) {
      console.error(
        "Error in batch operation:",
        error.response?.data || error.message
      );
      throw new Error(`Failed to execute batch operation: ${error.message}`);
    }
  }
}
