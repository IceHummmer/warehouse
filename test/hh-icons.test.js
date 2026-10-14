import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
const read=p=>readFileSync(new URL('../'+p,import.meta.url));
test('Warehouse desktop app PWA has same branded HH icon sizes',()=>{
 const manifest=JSON.parse(read('public/manifest.webmanifest').toString('utf8'));
 const icons=manifest.icons.filter(x=>x.type==='image/png');
 assert.deepEqual(icons.map(x=>x.src),['/pwa/hh-icon-192.png','/pwa/hh-icon-512.png']);
 for(const icon of icons){
  const png=read('public'+icon.src);
  assert.equal(png.toString('hex',0,8),'89504e470d0a1a0a');
  assert.equal(png.readUInt32BE(16),Number(icon.sizes.split('x')[0]));
  assert.equal(png.readUInt32BE(20),Number(icon.sizes.split('x')[1]));
 }
 const html=read('index.html').toString('utf8');
 assert.match(html,/href="\/favicon\.svg\?v=hh-ios-3"/);
 assert.match(html,/href="\/manifest\.webmanifest\?v=hh-ios-3"/);
});
