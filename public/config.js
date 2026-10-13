// PUBLIC configuration: no passwords or PIN hashes belong here.
window.warehouseConfig = {
 apiUrl: 'https://hh-warehouse-api-test.onrender.com',
 testPinMode: true,
 // Enable only on the approved production frontend when ADMIN_PASSWORD_SCRYPT is set server-side.
 adminLoginEnabled: false,
 tenantId: '',
 clientId: '',
 apiScope: ''
};
