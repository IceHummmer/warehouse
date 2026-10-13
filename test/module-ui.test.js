import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8');
const app=fs.readFileSync(new URL('../src/main.js',import.meta.url),'utf8');
const styles=fs.readFileSync(new URL('../src/style.css',import.meta.url),'utf8');

test('Warehouse header has a same-tab Project Register button, not the old selector',()=>{
 assert.match(html,/id="warehouseModuleButton"/);
 assert.doesNotMatch(html,/id="warehouse-module-select"/);
 assert.match(app,/warehouseModuleButton.*addEventListener\('click'/);
 assert.match(app,/window\.location\.assign\(PROJECT_REGISTER_URL/);
});

test('Warehouse module entry shows a progress view until SSO and data resolve',()=>{
 assert.match(html,/id="moduleLoading"/);
 assert.match(html,/data-module-sso='pending'/);
 assert.match(styles,/html\[data-module-sso="pending"\] \.hh-module-loading/);
 assert.match(app,/document\.documentElement\.removeAttribute\('data-module-sso'\)/);
 assert.match(html,/href="\/hh-logo\.svg"/);
});

test('Warehouse Microsoft bootstrap uses single silent token acquisition',()=>{
 const connect=app.slice(app.indexOf('async function connect(){'),app.indexOf('async function reload(){'));
 assert.equal((connect.match(/acquireTokenSilent\(/g)||[]).length,1);
});
