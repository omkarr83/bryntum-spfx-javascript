import { getAccessToken } from './auth.service';

// const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001/api';
const API_BASE_URL = 'http://localhost:3001/api';

export interface ApiResponse<T> {
    success: boolean;
    data?: T;
    error?: string;
    message?: string;
}

/**
 * Get headers with authentication token and all required headers
 * Similar to the reference implementation - gets token directly from localStorage
 */
async function getAuthHeaders(): Promise<HeadersInit> {
    // Get token directly from localStorage (like reference code)
    let token = localStorage.getItem('dataverse_access_token');
    
    // If not found, try to get from MSAL
    if (!token) {
        console.log('[API Service] Token not in localStorage, fetching from MSAL...');
        token = await getAccessToken();
    }
    
    // Validate token expiry
    if (token) {
        const expiryTime = localStorage.getItem('dataverse_token_expiry');
        if (expiryTime) {
            const expiry = parseInt(expiryTime, 10);
            if (Date.now() >= expiry) {
                console.log('[API Service] Token expired, refreshing...');
                token = await getAccessToken();
            }
        }
    }
    
    const headers: HeadersInit = {
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'X-Requested-With': 'XMLHttpRequest',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
    };
    
    if (token) {
        headers['Authorization'] = `Bearer ${token}`;
        console.log('[API Service] ✓ Token included in headers, length:', token.length);
    } else {
        console.error('[API Service] ✗ No token available!');
        console.error('[API Service] This will cause authentication errors. Please ensure user is logged in.');
    }
    
    return headers;
}

export class ApiService {
    /**
     * Fetch all tasks from the backend
     */
    static async getTasks(): Promise<any> {
        try {
            const headers = await getAuthHeaders();
            console.log('[API Service] Fetching tasks from:', `${API_BASE_URL}/tasks`);
            console.log('[API Service] Request headers:', Object.keys(headers));
            
            const response = await fetch(`${API_BASE_URL}/tasks`, {
                method: 'GET',
                headers,
                credentials: 'include', // Include credentials for CORS
            });
            
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({ error: 'Unknown error' }));
                console.error('API Error Response:', errorData);
                throw new Error(`HTTP error! status: ${response.status}, message: ${errorData.error || 'Unknown error'}`);
            }
            return await response.json();
        } catch (error) {
            console.error('Error fetching tasks:', error);
            throw error;
        }
    }

    /**
     * Create a new task
     */
    static async createTask(task: any): Promise<ApiResponse<any>> {
        try {
            const headers = await getAuthHeaders();
            console.log('[API Service] Creating task');
            const response = await fetch(`${API_BASE_URL}/tasks`, {
                method: 'POST',
                headers,
                body: JSON.stringify(task),
                credentials: 'include',
            });
            return await response.json();
        } catch (error) {
            console.error('Error creating task:', error);
            throw error;
        }
    }

    /**
     * Update an existing task
     */
    static async updateTask(taskId: string, task: any): Promise<ApiResponse<any>> {
        try {
            const headers = await getAuthHeaders();
            console.log('[API Service] Updating task:', taskId);
            const response = await fetch(`${API_BASE_URL}/tasks/${taskId}`, {
                method: 'PUT',
                headers,
                body: JSON.stringify(task),
                credentials: 'include',
            });
            return await response.json();
        } catch (error) {
            console.error('Error updating task:', error);
            throw error;
        }
    }

    /**
     * Delete a task
     */
    static async deleteTask(taskId: string): Promise<ApiResponse<any>> {
        try {
            const headers = await getAuthHeaders();
            console.log('[API Service] Deleting task:', taskId);
            const response = await fetch(`${API_BASE_URL}/tasks/${taskId}`, {
                method: 'DELETE',
                headers,
                credentials: 'include',
            });
            return await response.json();
        } catch (error) {
            console.error('Error deleting task:', error);
            throw error;
        }
    }

    /**
     * Sync all tasks (batch operation)
     */
    static async syncTasks(tasks: any[]): Promise<ApiResponse<any>> {
        try {
            const headers = await getAuthHeaders();
            console.log('[API Service] Syncing tasks:', tasks.length);
            const response = await fetch(`${API_BASE_URL}/tasks/sync`, {
                method: 'POST',
                headers,
                body: JSON.stringify({ tasks }),
                credentials: 'include',
            });
            return await response.json();
        } catch (error) {
            console.error('Error syncing tasks:', error);
            throw error;
        }
    }
}
