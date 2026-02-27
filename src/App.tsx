import * as React from 'react';
import { FunctionComponent, useState, useEffect, useRef } from 'react';
import { Gantt, TaskModel } from '@bryntum/gantt';
import { useGanttProps, ensureTokenAvailable } from './AppConfig';
import './App.scss';
import { getAccessToken, isAuthenticated } from './services/auth.service';

const App: FunctionComponent = () => {
    const [ganttProps, setGanttProps] = useState<Record<string, unknown> | null>(null);
    const [isTokenReady, setIsTokenReady] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const ganttRef = useRef<HTMLDivElement>(null);
    const ganttInstanceRef = useRef<Gantt | null>(null);

    useEffect(() => {
        const initializeGantt = async () => {
            try {
                console.log('[App] Initializing Gantt component...');

                // First, ensure user is authenticated
                const authenticated = await isAuthenticated();
                if (!authenticated) {
                    console.error('[App] User is not authenticated');
                    setError('User is not authenticated. Please login.');
                    return;
                }

                // Ensure token is available before creating Gantt props
                console.log('[App] Ensuring token is available...');
                let token = await ensureTokenAvailable();

                if (!token) {
                    console.error('[App] Failed to get access token, retrying...');
                    // Retry once
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    token = await ensureTokenAvailable();
                }

                if (!token) {
                    console.error('[App] Failed to get access token after retry');
                    setError('Failed to get access token. Please try logging in again.');
                    return;
                }

                console.log('[App] Token is available, length:', token.length);
                console.log('[App] Token preview:', token.substring(0, 30) + '...');
                console.log('[App] Verifying token is cached...');

                // Verify token is in cache/localStorage
                const cachedToken = localStorage.getItem('dataverse_access_token');
                if (!cachedToken || cachedToken !== token) {
                    console.warn('[App] Token not properly cached, updating cache...');
                    // Token should already be cached by ensureTokenAvailable, but ensure it
                    await new Promise(resolve => setTimeout(resolve, 100));
                }

                console.log('[App] Creating Gantt props...');

                // Create Gantt props with token ready
                const handleEditClick = (record: TaskModel, grid: Gantt) => {
                    grid.editTask(record);
                };

                const props = useGanttProps(handleEditClick);
                setGanttProps(props);

                // Small delay to ensure everything is ready
                await new Promise(resolve => setTimeout(resolve, 200));

                setIsTokenReady(true);
                console.log('[App] ✓ Gantt component ready with token');
            } catch (err: any) {
                console.error('[App] Error initializing Gantt:', err);
                setError(`Failed to initialize: ${err.message || 'Unknown error'}`);
            }
        };

        initializeGantt();
    }, []);

    useEffect(() => {
        if (!ganttProps || !ganttRef.current) return;
        if (ganttInstanceRef.current) {
            ganttInstanceRef.current.destroy();
            ganttInstanceRef.current = null;
        }
        ganttInstanceRef.current = new Gantt({
            ...ganttProps,
            appendTo: ganttRef.current
        } as any);
        return () => {
            if (ganttInstanceRef.current) {
                ganttInstanceRef.current.destroy();
                ganttInstanceRef.current = null;
            }
        };
    }, [ganttProps]);

    if (error) {
        return (
            <>
                {/* <BryntumDemoHeader /> */}
                <div style={{
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    height: '80vh',
                    flexDirection: 'column',
                    gap: '20px',
                    padding: '20px'
                }}>
                    <div style={{ color: 'red', fontSize: '18px', fontWeight: 'bold' }}>
                        {error}
                    </div>
                    <button
                        onClick={async () => {
                            setError(null);
                            setIsTokenReady(false);
                            try {
                                const token = await getAccessToken();
                                if (token) {
                                    window.location.reload();
                                } else {
                                    setError('Failed to get token. Please check console for details.');
                                }
                            } catch (err: any) {
                                setError(`Error: ${err.message}`);
                            }
                        }}
                        style={{
                            padding: '10px 20px',
                            fontSize: '16px',
                            cursor: 'pointer'
                        }}
                    >
                        Retry
                    </button>
                </div>
            </>
        );
    }

    if (!isTokenReady || !ganttProps) {
        return (
            <>
                {/* <BryntumDemoHeader /> */}
                <div style={{
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    height: '80vh',
                    flexDirection: 'column',
                    gap: '20px'
                }}>
                    <div>Loading Gantt Chart...</div>
                    <div style={{ fontSize: '14px', color: '#666' }}>
                        Ensuring authentication token is ready...
                    </div>
                </div>
            </>
        );
    }

    return (
        <>
            <div
                ref={ganttRef}
                style={{ height: '100%', minHeight: '400px' }}
            />
        </>
    );
};

export default App;