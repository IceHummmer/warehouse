// PUBLIC configuration: client IDs are non-secret. Leave testPinMode TRUE on test deployment.
// API scope assumes standard Application ID URI; verify under Entra > HH Warehouse API > Expose an API.
window.warehouseConfig = {
 apiUrl: 'https://hh-warehouse-api-test.onrender.com',
 testPinMode: true,
 // Enable only on the approved production frontend when ADMIN_PASSWORD_SCRYPT is set server-side.
 adminLoginEnabled: false,
 tenantId: '02bec730-845f-46db-aec9-f469f4992917',
 clientId: '717030ab-fc22-4364-b945-e4e0173260e3',
 apiScope: 'api://534af606-f406-4a82-ac99-3b17e692f9d9/Warehouse.Access'
};
