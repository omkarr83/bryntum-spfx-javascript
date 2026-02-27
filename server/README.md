# Backend Server - Gantt Dataverse Integration

Node.js/Express backend server for integrating Bryntum Gantt with Microsoft Dataverse.

## Installation

```bash
npm install
```

## Configuration

Create a `.env` file in the `server` directory:

```env
PORT=3001
DATAVERSE_ENVIRONMENT_URL=https://your-org.crm8.dynamics.com
DATAVERSE_TENANT_ID=your-tenant-id
DATAVERSE_CLIENT_ID=your-client-id
DATAVERSE_CLIENT_SECRET=your-client-secret
DATAVERSE_TABLE_NAME=eppm_projecttasks
```

## Running

### Development
```bash
npm run dev
```

### Production
```bash
npm run build
npm start
```

## API Endpoints

- `GET /health` - Health check
- `GET /api/tasks` - Get all tasks
- `GET /api/tasks/:id` - Get single task
- `POST /api/tasks` - Create task
- `PUT /api/tasks/:id` - Update task
- `DELETE /api/tasks/:id` - Delete task
- `POST /api/tasks/sync` - Sync multiple tasks

## Architecture

- **Express**: Web server framework
- **TypeScript**: Type-safe JavaScript
- **@azure/identity**: Azure authentication
- **axios**: HTTP client for Dataverse API
