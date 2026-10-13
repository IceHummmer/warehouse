import test from 'node:test';
import assert from 'node:assert/strict';
import {t,setLanguage,getLanguage,localizeError} from '../src/i18n.js';

test('switching between Russian and English translates only interface labels',()=>{
 setLanguage('ru');
 assert.equal(getLanguage(),'ru');
 assert.equal(t('Inventory'),'Склад');
 assert.equal(t('Contractor'),'Подрядчик');
 assert.equal(t('Employee'),'Сотрудник');
 assert.equal(t('Issue date'),'Дата выдачи');
 assert.equal(t('Rental equipment'),'Арендуемое оборудование');
 assert.equal(t('Sign out'),'Выйти');
 assert.equal(t('Critical total quantity'),'Критический остаток (по общему количеству)');
 // Original workbook and user-entered names must remain byte-for-byte identical.
 for(const original of ['FEHU OU','IPPON OU','PPE-007','CMK 3.0','Сварочный аппарат Kemppi','Иван Иванов','Original equipment name']){
  assert.equal(t(original),original);
 }
 assert.equal(localizeError('Invalid test access code'),'Неверный код доступа');
 setLanguage('en');
 assert.equal(getLanguage(),'en');
 assert.equal(t('Inventory'),'Inventory');
 assert.equal(t('Employee'),'Employee');
 assert.equal(t('FEHU OU'),'FEHU OU');
});
