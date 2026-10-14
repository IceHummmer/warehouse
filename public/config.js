// HH Warehouse Register — official deployment configuration.
// This frontend is public; keep only non-secret Microsoft application identifiers.
// Official launch enabled after permanent GitHub storage, weekly backup restore,
// Microsoft authorization and separate admin secret checks passed.
window.warehouseConfig = {
 apiUrl: 'https://hh-warehouse-api.onrender.com',
 productionMode: true,
 launchReady: true,
 testPinMode: false,
 microsoftEnabled: true,
 adminLoginEnabled: true,
 tenantId: '02bec730-845f-46db-aec9-f469f4992917',
 clientId: '717030ab-fc22-4364-b945-e4e0173260e3',
 apiScope: 'api://534af606-f406-4a82-ac99-3b17e692f9d9/Warehouse.Access'
};
