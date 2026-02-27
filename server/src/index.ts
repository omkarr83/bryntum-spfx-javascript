import express from 'express';
import cors from 'cors';
import { serverConfig } from './config/dataverse.config.js';
import tasksRoutes from './routes/tasks.routes.js';

const app = express();

// Middleware
app.use(cors({
    origin: true, // Allow all origins
    credentials: true,
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'OData-MaxVersion', 'OData-Version'],
    exposedHeaders: ['Authorization', 'Content-Type'],
    methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS']
}));

// Middleware to set response headers
app.use((req, res, next) => {
    // Set default response headers
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('X-XSS-Protection', '1; mode=block');
    next();
});

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Handle OPTIONS preflight requests
app.options('*', (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, PATCH, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With, Accept, OData-MaxVersion, OData-Version');
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Max-Age', '86400'); // 24 hours
    res.status(204).end();
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.json({ status: 'ok', message: 'Server is running' });
});

// API routes
app.use('/api/tasks', tasksRoutes);

// Error handling middleware
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
    console.error('Error:', err);
    res.status(err.status || 500).json({
        success: false,
        error: err.message || 'Internal server error'
    });
});

// Start server
const PORT = serverConfig.port;

app.listen(PORT, () => {
    console.log(`🚀 Server is running on http://localhost:${PORT}`);
    console.log(`📊 Dataverse Environment: ${process.env.DATAVERSE_ENVIRONMENT_URL}`);
    console.log(`📋 Table Name: ${process.env.DATAVERSE_TABLE_NAME}`);
});
