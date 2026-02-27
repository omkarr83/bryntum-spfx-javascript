import dotenv from 'dotenv';

dotenv.config();

export const dataverseConfig = {
    environmentUrl: process.env.DATAVERSE_ENVIRONMENT_URL || '',
    tenantId: process.env.DATAVERSE_TENANT_ID || '',
    clientId: process.env.DATAVERSE_CLIENT_ID || '',
    clientSecret: process.env.DATAVERSE_CLIENT_SECRET || '',
    tableName: process.env.DATAVERSE_TABLE_NAME || 'eppm_projecttasks'
};

export const serverConfig = {
    port: parseInt(process.env.PORT || '3001', 10)
};
