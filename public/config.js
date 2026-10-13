// HH Warehouse Register — official deployment configuration.
// This frontend is public; keep only non-secret Microsoft application identifiers.
// Deployment remains in maintenance until persistent storage, backup,
// Microsoft authorization and independent administrator are verified.
window.warehouseConfig = {
 apiUrl: 'https://hh-warehouse-api.onrender.com',
 productionMode: true,
 launchReady: false,
 testPinMode: false,
 microsoftEnabled: true,
 adminLoginEnabled: false,
 tenantId: '02bec730-845f-46db-aec9-f469f4992917',
 clientId: '717030ab-fc22-4364-b945-e4e0173260e3',
 apiScope: 'api://534af606-f406-4a82-ac99-3b17e692f9d9/Warehouse.Access'
};
