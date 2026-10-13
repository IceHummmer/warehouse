import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
const read=path=>fs.readFileSync(new URL('../'+path,import.meta.url),'utf8');
test('Warehouse signed-in username is a button opening the sign-out confirmation',()=>{
 const html=read('index.html'),js=read('src/main.js');
 assert.match(html,/id="warehouseAccountButton"/);
 assert.match(html,/id="warehouseSignOutDialog"/);
 assert.doesNotMatch(html,/id="logout"/);
 assert.match(js,/warehouseAccountButton'\)\.addEventListener/);
 assert.match(js,/warehouseSignOutConfirm'\)\.addEventListener/);
});
test('Warehouse user registry has separate display name and Microsoft account columns',()=>{
 const js=read('src/main.js');
 assert.match(js,/'name','User'/);
 assert.match(js,/'email','Microsoft account'/);
 assert.match(js,/field\('name','User'/);
 assert.match(js,/field\('email','Microsoft account'/);
});
test('Warehouse footer and secure admin handoff exist',()=>{
 const html=read('index.html'),js=read('src/main.js');
 assert.match(html,/© Helsingin Hitsaus/);
 assert.match(html,/data-hh-version/);
 assert.match(js,/\/api\/auth\/admin-handoff\/issue/);
 assert.match(js,/\/api\/auth\/admin-handoff\/exchange/);
 assert.match(JSON.parse(read('version.json')).version,/^\d+\.\d+\.\d+$/);
});
