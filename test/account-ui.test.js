import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
const html=readFileSync(new URL('../index.html',import.meta.url),'utf8');
const js=readFileSync(new URL('../src/main.js',import.meta.url),'utf8');
test('Warehouse account is one username button with confirmation dialog',()=>{
 assert.match(html,/id="warehouseAccountButton"/);
 assert.doesNotMatch(html,/id="logout"/);
 assert.match(html,/id="warehouseSignOutDialog"/);
 assert.match(js,/warehouseSignOutConfirm/);
});
test('Warehouse Users & Access shows user and Microsoft account separately',()=>{
 assert.match(js,/\['name','User'\],\['email','Microsoft account'\]/);
 assert.match(js,/member\?\.name/);
 assert.match(js,/admin-handoff\/issue/);
 assert.match(js,/admin-handoff\/exchange/);
});
test('Warehouse displays company copyright and bundled version',()=>{
 assert.match(html,/© Helsingin Hitsaus/);
 assert.match(html,/data-hh-version/);
 assert.match(js,/releaseInfo\.version/);
});
