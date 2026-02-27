import * as React from 'react';
import { StrictMode } from 'react';
import * as ReactDOM from 'react-dom';
import App from './App';
import { AuthProvider } from './components/AuthProvider';

const root = document.getElementById('root');
if (root) {
    ReactDOM.render(
        <StrictMode>
            <AuthProvider>
                <App />
            </AuthProvider>
        </StrictMode>,
        root
    );
}
