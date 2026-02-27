import * as React from 'react';
import { FunctionComponent, ReactNode, useEffect, useState } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from '../config/msalConfig';
import { getCurrentAccount, getAccessToken, login, setMsalInstance } from '../services/auth.service';
import { updateCachedToken } from '../AppConfig';

interface AuthProviderProps {
    children: ReactNode;
}

export const AuthProvider: FunctionComponent<AuthProviderProps> = ({ children }) => {
    const [msalInstance] = useState<PublicClientApplication>(() => {
        const instance = new PublicClientApplication(msalConfig);
        return instance;
    });
    const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
    const [isLoading, setIsLoading] = useState<boolean>(true);

    useEffect(() => {
        const checkAuth = async () => {
            try {
                console.log('Starting authentication check...');
                // Initialize MSAL instance first
                console.log('Initializing MSAL instance...');
                await msalInstance.initialize();
                console.log('MSAL instance initialized successfully');
                // Share the instance with auth.service
                setMsalInstance(msalInstance);
                console.log('MSAL instance shared with auth.service');
                
                // Check for existing account
                let account = await getCurrentAccount();
                
                // Check localStorage first for token (synchronous, doesn't need MSAL)
                const storedToken = localStorage.getItem('dataverse_access_token');
                const expiryTime = localStorage.getItem('dataverse_token_expiry');
                
                if (storedToken) {
                    // Check if token is expired
                    let isValid = true;
                    if (expiryTime) {
                        const expiry = parseInt(expiryTime, 10);
                        if (Date.now() >= expiry) {
                            isValid = false;
                            localStorage.removeItem('dataverse_access_token');
                            localStorage.removeItem('dataverse_token_expiry');
                        }
                    }
                    
                    if (isValid) {
                        console.log('Found valid stored token');
                        await updateCachedToken();
                        setIsAuthenticated(true);
                        setIsLoading(false);
                        return;
                    }
                }
                
                if (account) {
                    // Try to get token silently
                    const token = await getAccessToken();
                    if (token) {
                        // Update cached token
                        await updateCachedToken();
                        setIsAuthenticated(true);
                        setIsLoading(false);
                        return;
                    }
                }
                
                // No account or silent token failed - auto-login
                console.log('No account found or token expired. Attempting auto-login...');
                try {
                    const loginResult = await login();
                    if (loginResult && loginResult.account) {
                        console.log('Login successful, getting access token...');
                        const token = await getAccessToken();
                        if (token) {
                            console.log('Access token obtained successfully');
                            // Update cached token (token is already stored in localStorage by getAccessToken)
                            await updateCachedToken();
                            setIsAuthenticated(true);
                        } else {
                            console.error('Failed to get access token after login');
                            setIsAuthenticated(false);
                        }
                    } else {
                        console.error('Login failed - no account returned');
                        setIsAuthenticated(false);
                    }
                } catch (loginError: any) {
                    console.error('Login attempt failed:', loginError);
                    console.error('Error details:', {
                        errorCode: loginError.errorCode,
                        errorMessage: loginError.message,
                        stack: loginError.stack
                    });
                    setIsAuthenticated(false);
                }
            } catch (error: any) {
                console.error('Auth check error:', error);
                console.error('Error details:', {
                    errorCode: error.errorCode,
                    errorMessage: error.message,
                    stack: error.stack
                });
                // Even if auto-login fails, try to show the UI
                setIsAuthenticated(false);
            } finally {
                setIsLoading(false);
            }
        };

        checkAuth();
    }, []);

    if (isLoading) {
        return (
            <div style={{ 
                display: 'flex', 
                justifyContent: 'center', 
                alignItems: 'center', 
                height: '100vh',
                flexDirection: 'column',
                gap: '20px'
            }}>
                <div>Loading...</div>
                <div>Authenticating with Microsoft...</div>
            </div>
        );
    }

    if (!isAuthenticated) {
        return (
            <div style={{ 
                display: 'flex', 
                justifyContent: 'center', 
                alignItems: 'center', 
                height: '100vh',
                flexDirection: 'column',
                gap: '20px'
            }}>
                <div>Authentication failed. Please try again.</div>
                <div style={{ fontSize: '12px', color: '#666', maxWidth: '500px', textAlign: 'center', marginBottom: '10px' }}>
                    Check browser console (F12) for detailed error messages.
                    <br />
                    Common issues: Popup blocked, redirect URI mismatch, or expired token.
                    <br />
                    <button 
                        onClick={() => {
                            console.log('=== Authentication Diagnostics ===');
                            console.log('Redirect URI:', window.location.origin);
                            console.log('Stored token exists:', !!localStorage.getItem('dataverse_access_token'));
                            console.log('Token expiry:', localStorage.getItem('dataverse_token_expiry'));
                            console.log('MSAL instance:', msalInstance ? 'Exists' : 'Missing');
                            console.log('Current URL:', window.location.href);
                            alert('Diagnostics logged to console. Press F12 to view.');
                        }}
                        style={{
                            padding: '5px 10px',
                            fontSize: '12px',
                            marginTop: '10px',
                            cursor: 'pointer'
                        }}
                    >
                        Run Diagnostics
                    </button>
                </div>
                <button 
                    onClick={async () => {
                        setIsLoading(true);
                        try {
                            console.log('Manual login attempt...');
                            const result = await login();
                            if (result && result.account) {
                                console.log('Login successful, getting token...');
                                const token = await getAccessToken();
                                if (token) {
                                    await updateCachedToken();
                                    setIsAuthenticated(true);
                                } else {
                                    console.error('Failed to get token after manual login');
                                    alert('Login successful but failed to get access token. Please check console for details.');
                                }
                            } else {
                                console.error('Login failed - no result or account');
                                alert('Login failed. Please check console for details and ensure popups are allowed.');
                            }
                        } catch (error: any) {
                            console.error('Manual login error:', error);
                            console.error('Error details:', {
                                errorCode: error.errorCode,
                                errorMessage: error.message,
                                name: error.name
                            });
                            alert(`Login error: ${error.message || error.errorCode || 'Unknown error'}. Please check console for details.`);
                        } finally {
                            setIsLoading(false);
                        }
                    }}
                    style={{
                        padding: '10px 20px',
                        fontSize: '16px',
                        cursor: 'pointer'
                    }}
                >
                    Retry Login with Microsoft
                </button>
            </div>
        );
    }

    return <React.Fragment>{children}</React.Fragment>;
};
