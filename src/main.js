import './style.css';
import {PublicClientApplication} from '@azure/msal-browser';
const $=(sel,root=document)=>root.querySelector(sel);
const $$=(sel,root=document)=>Array.from(root.querySelectorAll(sel));
const cfg=window.warehouseConfig||{};
const pages={
 dashboard:['Dashboard','Inventory and project supply overview.'],
 inventory:['Inventory','Catalogue, total stock, available equipment and critical levels.'],
 receipts:['Stock Receipts','Purchase receipts and incoming items.'],
 issues:['Issued Materials','Consumable materials issued to projects and employees.'],
 rentals:['Rentals','Active rentals, closed rentals and monthly billing periods.'],
 projects:['Projects','Warehouse project directory. Independent from Project Register.'],
 employees:['Employees','Staff and contractor directory.'],
 history:['Change History','All recorded data changes and stock activity.'],
 system:['System messages','Administrative messages and operational warnings.']
};
const icons={inventory:'▦',receipts:'↓',issues:'↑',rentals:'⟳'};
const state={view:'dashboard',demo:false,user:null,token:null,client:null,account:null,data:{},search:'',sort:'sku',descending:false,zoom:100,drag:null};
const today=()=>new Date().toISOString().slice(0,10);
const money=x=>Number(x||0).toLocaleString('en-FI',{style:'currency',currency:'EUR'});
const esc=x=>String(x==null?'':x).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const id=()=>crypto.randomUUID();
const fullDate=x=>x?String(x).slice(0,10):'';
const options=(list,selected)=>list.map(x=>'<option value="'+esc(x.id)+'" '+(x.id===selected?'selected':'')+'>'+esc(x.label||x.name)+'</option>').join('');
function toast(message,bad=false){const n=$('#toast');n.textContent=message;n.className='toast'+(bad?' error':'');setTimeout(()=>n.classList.add('hidden'),4500);}
function fakeData(){
 const projects=['Rauma','Eems_Carrier','Uusikaupunki'].map(name=>({id:id(),name}));
 const employees=[{id:id(),name:'Example Worker A',contractor:'Example Contractor'},{id:id(),name:'Example Worker B',contractor:'Example Contractor'}];
 const presets=[
 ['CMS-101','Cutting discs 125 mm','Consumables','consumable',84,10,0,0],
 ['PPE-101','Protective gloves','Personal protective equipment','consumable',15,20,0,0],
 ['WLD-101','Welding wire 1.0 mm','Welding supplies','consumable',29,8,0,0],
 ['CTR-101','Cordless drill','Small equipment','rental',12,2,3,5.5],
 ['SME-101','Angle grinder','Small equipment','rental',8,3,2,6],
 ['WSP-101','Welding machine','Welding equipment','rental',6,1,2,65],
 ['CMS-102','Grinding discs','Consumables','consumable',145,30,0,0],
 ['SME-102','Extension cable 25 m','Equipment','rental',16,3,1,3.5],
 ['PPE-102','Safety glasses','Personal protective equipment','consumable',8,12,0,0],
 ['CTR-102','Toolbox','Equipment','rental',7,1,0,35]
 ];
 const items=presets.map(a=>({id:id(),sku:a[0],name:a[1],category:a[2],kind:a[3],total_qty:a[4],critical_qty:a[5],rented_qty:a[6],available_qty:a[4]-a[6],rental_rate:a[7],rental_unit:a[0]==='WSP-101'||a[0]==='CTR-102'?'month':'day',replacement_cost:120,revision:1,unit:'pcs'}));
 const rents=items.filter(x=>x.rented_qty).map((x,i)=>({id:id(),item_id:x.id,item_name:x.name,sku:x.sku,project_id:projects[i%3].id,project_name:projects[i%3].name,employee_id:employees[i%2].id,employee_name:employees[i%2].name,quantity:x.rented_qty,started_on:'2026-10-01',rate:x.rental_rate,rate_unit:x.rental_unit,status:'active',billed_amount:0,last_billed_through:null}));
 return {items,projects,employees,rentals:rents,movements:[],history:[],periods:[],messages:[]};
}
async function initMicrosoft(){
 if(!cfg.tenantId||!cfg.clientId||!cfg.apiScope){$('#login-status').textContent='Microsoft login is not configured for this test. Explore the interactive demo.';$('#login-microsoft').disabled=true;return;}
 state.client=new PublicClientApplication({auth:{clientId:cfg.clientId,authority:'https://login.microsoftonline.com/'+cfg.tenantId,redirectUri:location.origin},cache:{cacheLocation:'sessionStorage'}});
 await state.client.initialize();
 const accounts=state.client.getAllAccounts();
 if(accounts.length){state.account=accounts[0];try{await connect();}catch(e){$('#login-status').textContent=e.message;}}
}
async function connect(){
 const result=await state.client.acquireTokenSilent({scopes:[cfg.apiScope],account:state.account});
 state.token=result.accessToken;
 const data=await api('/api/bootstrap');
 if(data.user.role==='admin'){
   try{data.messages=(await api('/api/system-messages')).messages;}catch{data.messages=[];}
 }
 state.user=data.user;state.data=data;state.demo=false;
 $('#landing').classList.add('hidden');$('#shell').classList.remove('hidden');$('#mode-label').textContent='LIVE TEST';$('#mode-label').className='tag live';$('#identity').textContent=state.user.email;$('#demo-warning').classList.add('hidden');
 render();
}
async function api(path,body,method='GET'){
 const token=state.client? (await state.client.acquireTokenSilent({scopes:[cfg.apiScope],account:state.account})).accessToken :state.token;
 const response=await fetch(String(cfg.apiUrl||'http://localhost:8787').replace(/\/$/,'')+path,{method,headers:{Authorization:'Bearer '+token,...(body?{'Content-Type':'application/json'}:{})},body:body?JSON.stringify(body):undefined});
 const data=await response.json().catch(()=>({}));
 if(!response.ok)throw new Error(data.error||'Operation failed');
 return data;
}
async function reload(){if(!state.demo){state.data=await api('/api/bootstrap');if(state.user?.role==='admin'){try{state.data.messages=(await api('/api/system-messages')).messages;}catch{state.data.messages=[];}}}render();}
function beginDemo(){
 state.demo=true;state.user={email:'Demo administrator',role:'admin'};state.data=fakeData();
 $('#landing').classList.add('hidden');$('#shell').classList.remove('hidden');$('#mode-label').textContent='DEMO';$('#mode-label').className='tag';$('#identity').textContent='Demo administrator';$('#demo-warning').classList.remove('hidden');
 render();
}
function cols(){
 const v=state.view;
 const inventory=[['sku','SKU'],['name','Item'],['category','Category'],['kind','Type'],['total_qty','Total'],['rented_qty','Rented'],['available_qty','Available'],['critical_qty','Critical'],['rental_rate','Rental rate'],['status','Stock status'],['_actions','Actions']];
 const map={
 inventory,
 receipts:[['created_at','Date'],['sku','SKU'],['item_name','Item'],['quantity','Qty'],['unit_cost','Unit price'],['note','Document / note'],['actor','User']],
 issues:[['created_at','Date'],['sku','SKU'],['item_name','Item'],['quantity','Qty'],['project_name','Project'],['employee_name','Employee'],['actor','User']],
 rentals:[['sku','SKU'],['item_name','Equipment'],['project_name','Project'],['employee_name','Employee'],['quantity','Qty'],['started_on','Issued'],['rate','Rate'],['rate_unit','Per'],['status','Status'],['billed_amount','Billed'],['_actions','Actions']],
 projects:[['name','Project']],employees:[['name','Employee'],['contractor','Contractor']],
 history:[['created_at','Date & time'],['actor','User'],['entity','Entity'],['entity_id','Record'],['field','Field'],['before_value','Before'],['after_value','After'],['detail','Details']],
 system:[['created_at','Date & time'],['severity','Type'],['subject','Subject'],['details','Details'],['_actions','Actions']]
 };
 return map[v]||[];
}
function rows(){
 const d=state.data;
 switch(state.view){
 case 'inventory':return d.items.map(x=>({...x,status:x.total_qty<x.critical_qty?'low':'ok'}));
 case 'receipts':return d.movements.filter(x=>x.kind==='receipt');
 case 'issues':return d.movements.filter(x=>x.kind==='issue');
 case 'rentals':return d.rentals;
 case 'projects':return d.projects;
 case 'employees':return d.employees;
 case 'history':return d.history;
 case 'system':return d.messages||[];
 default:return [];
 }
}
function canEdit(){return ['admin','operator'].includes(state.user?.role);}
function button(text,action,value='',extra=''){return '<button class="button '+extra+'" data-action="'+esc(action)+'" data-value="'+esc(value)+'">'+esc(text)+'</button>';}
function actions(){
 const v=state.view;if(!canEdit())return '';
 if(v==='inventory')return button('+ Add item','item','','primary')+button('Receive stock','receipt')+button('Issue material','issue')+button('Rent equipment','rent')+(state.user.role==='admin'?button('Adjust stock','adjust'):'');
 if(v==='receipts')return button('+ Receive stock','receipt','','primary');
 if(v==='issues')return button('+ Issue material','issue','','primary');
 if(v==='rentals')return button('+ Rent equipment','rent','','primary')+button('Monthly closing','monthly');
 if(v==='projects')return button('+ Add project','project','','primary');
 if(v==='employees')return button('+ Add employee','employee','','primary');
 return '';
}
function render(){
 if(!state.user)return;
 if(state.view==='system'&&state.user.role!=='admin')state.view='dashboard';
 const d=state.data;
 $('#view-title').textContent=pages[state.view][0];$('#view-description').textContent=pages[state.view][1];
 $$('#tabs .tab').forEach(x=>x.classList.toggle('active',x.dataset.view===state.view));
 $$('.admin-only').forEach(x=>x.classList.toggle('hidden',state.user.role!=='admin'));
 const low=d.items.filter(x=>x.total_qty<x.critical_qty).length;
 const active=d.rentals.filter(x=>x.status==='active').length;
 $('#summary').innerHTML=[['Items',d.items.length],['Low stock',low],['Active rentals',active]].map(([label,value])=>'<div class="stat"><strong>'+esc(value)+'</strong><span>'+label+'</span></div>').join('');
 $('#primary-actions').innerHTML=actions();
 $('#zoom-value').textContent=state.zoom+'%';
 const count=(d.messages||[]).filter(m=>!m.is_read).length;
 $('#unread-count').textContent=count;$('#unread-count').classList.toggle('hidden',!count);
 if(state.view==='dashboard'){$('#page-content').innerHTML=dashboard();return;}
 const columns=cols(), q=state.search.toLowerCase(),found=rows().filter(x=>JSON.stringify(x).toLowerCase().includes(q));
 found.sort((a,b)=>{const x=a[state.sort],y=b[state.sort];const c=typeof x==='number'&&typeof y==='number'?x-y:String(x??'').localeCompare(String(y??''),undefined,{numeric:true});return state.descending?-c:c;});
 const widths=JSON.parse(localStorage.getItem('hh-wh-widths-'+state.view)||'{}');
 const headers=columns.map(([key,label])=>'<th data-sort="'+esc(key)+'" style="width:'+esc(widths[key]||'auto')+'"><div class="heading">'+esc(label)+(state.sort===key?(state.descending?' ↓':' ↑'):'')+'<span class="grip" data-grip="'+esc(key)+'"></span></div></th>').join('');
 const dataRows=found.map(x=>'<tr>'+columns.map(([key])=>'<td>'+cell(x,key)+'</td>').join('')+'</tr>').join('');
 $('#page-content').innerHTML='<div class="table-scroller"><table class="data-table" style="--zoom-font:'+Math.round(12*state.zoom/100)+'px"><thead><tr>'+headers+'</tr></thead><tbody>'+dataRows+'</tbody></table>'+(found.length?'':'<div class="empty"><h3>No records found</h3><p>Use the actions above to add data or change the search.</p></div>')+'</div>';
}
function dashboard(){
 const d=state.data, low=d.items.filter(x=>x.total_qty<x.critical_qty);
 const list=(arr,empty)=>arr.length?'<ul class="dash-list">'+arr.slice(0,5).map(x=>'<li>'+x+'</li>').join('')+'</ul>':'<p>'+empty+'</p>';
 return '<div class="dashboard-cards">'+
 '<div class="dash-card"><h3>Inventory overview</h3><div class="big-number">'+d.items.length+'</div><p>Catalogue entries across '+new Set(d.items.map(x=>x.category)).size+' categories.</p>'+button('Browse inventory','view','inventory')+'</div>'+
 '<div class="dash-card"><h3>Critical inventory</h3><div class="big-number">'+low.length+'</div>'+list(low.map(x=>'<strong>'+esc(x.sku)+'</strong> · '+esc(x.name)+' · total '+x.total_qty),'All items above critical level.')+'</div>'+
 '<div class="dash-card"><h3>Active rentals</h3><div class="big-number">'+d.rentals.filter(x=>x.status==='active').length+'</div>'+list(d.rentals.filter(x=>x.status==='active').map(x=>esc(x.item_name)+' · '+esc(x.employee_name)),'No active rentals.')+button('Manage rentals','view','rentals')+'</div>'+
 '<div class="dash-card"><h3>Warehouse movements</h3><p>Receipts, issues, rental releases and losses are captured as auditable operations.</p>'+button('View receipts','view','receipts')+'</div>'+
 '<div class="dash-card"><h3>Project allocations</h3><div class="big-number">'+d.projects.length+'</div><p>Projects are independent from HH Project Register.</p>'+button('View projects','view','projects')+'</div>'+
 '<div class="dash-card"><h3>Rental policy</h3><p>Daily tariffs include both dates. Monthly tariffs bill every calendar month touched. Transfer and partial return are disabled in v1.</p></div></div>';
}
function cell(x,key){
 const value=x[key];
 if(key==='_actions'){
  if(!canEdit())return '';
  if(state.view==='inventory')return '<div class="row-actions"><button class="mini" data-action="edit" data-value="'+esc(x.id)+'">Edit</button><button class="mini" data-action="receipt" data-value="'+esc(x.id)+'">Receive</button></div>';
  if(state.view==='rentals'&&x.status==='active')return '<button class="mini" data-action="close" data-value="'+esc(x.id)+'">Close rental</button>';
  if(state.view==='system'&&!x.is_read)return '<button class="mini" data-action="read" data-value="'+esc(x.id)+'">Mark read</button>';
  return '';
 }
 if(key==='status'||key==='kind')return '<span class="pill '+esc(String(value).toLowerCase())+'">'+esc(value==='low'?'Critical':value==='ok'?'Available':value)+'</span>';
 if(key==='rental_rate'||key==='rate'||key==='billed_amount'||key==='unit_cost')return money(value);
 if(key==='created_at')return value?new Date(value).toLocaleString('en-FI'):'';
 if(key==='started_on')return esc(fullDate(value));
 return esc(value);
}
const field=(name,label,type='text',value='',extra='')=>'<label class="field"><span>'+esc(label)+'</span><input name="'+esc(name)+'" type="'+esc(type)+'" value="'+esc(value)+'" '+extra+( ['note','contractor'].includes(name)?'':' required')+' /></label>';
const selectField=(name,label,values,selected)=>'<label class="field"><span>'+esc(label)+'</span><select name="'+esc(name)+'" required>'+options(values,selected)+'</select></label>';
const itemOptions=(kind)=>state.data.items.filter(x=>!kind||x.kind===kind).map(x=>({id:x.id,label:x.sku+' · '+x.name+' (available '+(x.total_qty-x.rented_qty)+')'}));
function modal(type,value=''){
 let title='',content='',note='',record=null;
 const item=state.data.items.find(x=>x.id===value),rental=state.data.rentals.find(x=>x.id===value);
 const projectOpt=state.data.projects.map(x=>({id:x.id,label:x.name})), employeeOpt=state.data.employees.map(x=>({id:x.id,label:x.name}));
 if(type==='item'||type==='edit'){
  title=type==='edit'?'Edit item':'Create inventory item';record=type==='edit'?item:null;
  content=field('sku','SKU','text',record?.sku||'',type==='edit'?'readonly':'')+field('name','Item name','text',record?.name||'')+field('category','Category','text',record?.category||'Equipment')+
  (type==='item'?selectField('kind','Kind',[{id:'consumable',label:'Consumable'},{id:'rental',label:'Rental equipment'}],'consumable'):'')+
  (type==='item'?field('total_qty','Initial total quantity','number',0,'min="0" step="1"'):'')+
  field('critical_qty','Critical total quantity','number',record?.critical_qty??0,'min="0" step="1"')+
  selectField('rental_unit','Rental unit',[{id:'day',label:'Per day'},{id:'month',label:'Per month'}],record?.rental_unit||'day')+
  field('rental_rate','Rental rate, EUR','number',record?.rental_rate||0,'min="0" step="0.01"')+
  field('replacement_cost','Replacement cost, EUR','number',record?.replacement_cost||0,'min="0" step="0.01"');
 }
 if(type==='receipt'){title='Receive stock';content=selectField('item_id','Item',itemOptions(),item?.id)+field('quantity','Quantity','number',1,'min="1" step="1"')+field('unit_cost','Purchase price per unit, EUR','number',0,'min="0" step="0.01"')+field('note','Document / note','text','');}
 if(type==='issue'){title='Issue consumable material';content=selectField('item_id','Consumable item',itemOptions('consumable'))+selectField('project_id','Project',projectOpt)+selectField('employee_id','Employee',employeeOpt)+field('quantity','Quantity','number',1,'min="1" step="1"')+field('note','Note','text','');}
 if(type==='rent'){title='Rent equipment';content=selectField('item_id','Rental equipment',itemOptions('rental'))+selectField('project_id','Project',projectOpt)+selectField('employee_id','Employee',employeeOpt)+field('quantity','Quantity','number',1,'min="1" step="1"')+field('started_on','Issue date','date',today());note='Rental rate is captured when equipment is issued. Closing uses the captured historical rate.';}
 if(type==='close'){title='Close rental — '+(rental?.item_name||'');content=field('closed_on','Return / loss date','date',today())+selectField('status','Condition',[{id:'returned',label:'Returned'},{id:'lost',label:'Lost / broken'}],'returned')+field('note','Reason / note','text','');note='Closes the full quantity: '+esc(rental?.quantity)+'. Partial return is not available in v1.';}
 if(type==='monthly'){title='Monthly rental closing';const end=new Date();end.setMonth(end.getMonth()+1,0);const through=end.toISOString().slice(0,10);content=selectField('project_id','Project',projectOpt)+field('through','Closing date (month end)','date',through);note='Bills all active project rentals through the selected last day of a month, without changing their physical stock.';}
 if(type==='adjust'){title='Correct stock (administrator)';content=selectField('item_id','Item',itemOptions(),item?.id)+field('total_qty','New total quantity','number',item?.total_qty??0,'min="0" step="1"')+field('reason','Reason for adjustment','text','');note='Adjustments are recorded in Change History.';}
 if(type==='project'){title='Add project';content=field('name','Project name');}
 if(type==='employee'){title='Add employee';content=field('name','Employee name')+field('contractor','Contractor','text','');}
 if(!content){toast('No action available',true);return;}
 $('#modal-host').innerHTML='<div class="modal-screen" role="presentation"><section class="modal" role="dialog" aria-modal="true" aria-label="'+esc(title)+'"><div class="modal-head"><h3>'+esc(title)+'</h3><button class="icon-button" data-action="dismiss" aria-label="Close">×</button></div><form id="record-form" data-type="'+esc(type)+'" data-record="'+esc(value)+'"><div class="modal-body">'+(note?'<div class="note">'+note+'</div>':'')+'<div class="fields">'+content+'</div></div><div class="modal-footer"><button type="button" class="button" data-action="dismiss">Cancel</button><button type="submit" class="button primary">Save changes</button></div></form></section></div>';
 $('#record-form').addEventListener('submit',save);
}
function dateISO(value){return String(value||'').slice(0,10);}
function amount(kind,rate,qty,start,end){
 const d1=new Date(dateISO(start)+'T00:00:00Z'),d2=new Date(dateISO(end)+'T00:00:00Z');
 if(d2<d1)throw new Error('Closing date precedes issue date');
 const units=kind==='day'?Math.round((d2-d1)/86400000)+1:(d2.getUTCFullYear()-d1.getUTCFullYear())*12+d2.getUTCMonth()-d1.getUTCMonth()+1;
 return Math.round(units*Number(rate)*Number(qty)*100)/100;
}
function demoAudit(entity,record,field,before,after,detail=''){state.data.history.unshift({id:id(),created_at:new Date().toISOString(),actor:state.user.email,entity,entity_id:record,field,before_value:before,after_value:after,detail});}
function demoMove(item,kind,quantity,delta,extra={}){
 state.data.movements.unshift({id:id(),created_at:new Date().toISOString(),item_id:item.id,item_name:item.name,sku:item.sku,kind,quantity,quantity_change:delta,actor:state.user.email,...extra});
}
function demoMutation(type,b,value){
 const d=state.data,item=d.items.find(x=>x.id===b.item_id||x.id===value),count=Number(b.quantity||0);
 if(type==='item'){
  if(d.items.some(x=>x.sku===b.sku.toUpperCase()))throw new Error('SKU already exists');
  const o={...b,id:id(),sku:b.sku.toUpperCase(),total_qty:Number(b.total_qty),rented_qty:0,critical_qty:Number(b.critical_qty),rental_rate:Number(b.rental_rate),replacement_cost:Number(b.replacement_cost),available_qty:Number(b.total_qty),revision:1,unit:'pcs'};
  d.items.push(o);demoAudit('item',o.id,'created',null,o.sku);return;
 }
 if(type==='edit'){
  if(!item)throw new Error('Item unavailable');
  for(const key of ['name','category','critical_qty','rental_unit','rental_rate','replacement_cost']){
   if(key==='rental_unit'&&item.kind==='consumable')continue;
   const next=['critical_qty','rental_rate','replacement_cost'].includes(key)?Number(b[key]):b[key];
   if(String(item[key]??'')!==String(next??'')){demoAudit('item',item.id,key,item[key],next);item[key]=next;}
  }item.revision++;return;
 }
 if(type==='project'){if(d.projects.some(x=>x.name===b.name))throw new Error('Duplicate project');d.projects.push({id:id(),name:b.name});return;}
 if(type==='employee'){if(d.employees.some(x=>x.name===b.name))throw new Error('Duplicate employee');d.employees.push({id:id(),name:b.name,contractor:b.contractor});return;}
 if(type==='monthly'){
  if(new Date(b.through+'T00:00:00Z').getUTCDate()!==new Date(new Date(b.through+'T00:00:00Z').setUTCMonth(new Date(b.through+'T00:00:00Z').getUTCMonth()+1,0)).getUTCDate())throw new Error('Choose the last day of a month');
  let n=0;
  for(const r of d.rentals.filter(x=>x.project_id===b.project_id&&x.status==='active'&&x.started_on<=b.through)){
   const start=r.last_billed_through?nextDate(r.last_billed_through):r.started_on;
   if(start>b.through)continue;
   const charge=amount(r.rate_unit,r.rate,r.quantity,start,b.through);
   d.periods.unshift({rental_id:r.id,start_on:start,end_on:b.through,amount:charge});
   r.billed_amount=Number(r.billed_amount||0)+charge;r.last_billed_through=b.through;n++;
  }
  if(!n)throw new Error('No unbilled active rentals in this period');
  demoAudit('project',b.project_id,'monthly_closing',null,b.through,n+' rentals');return;
 }
 if(type==='close'){
  const r=d.rentals.find(x=>x.id===value);if(!r||r.status!=='active')throw new Error('Rental already closed');
  if(b.closed_on<r.started_on)throw new Error('Closing date precedes issue date');
  const i=d.items.find(x=>x.id===r.item_id),start=r.last_billed_through?nextDate(r.last_billed_through):r.started_on;
  const charge=start<=b.closed_on?amount(r.rate_unit,r.rate,r.quantity,start,b.closed_on):0;
  r.billed_amount=Number(r.billed_amount||0)+charge+(b.status==='lost'?Number(i.replacement_cost)*r.quantity:0);
  r.status=b.status;r.closed_on=b.closed_on;i.rented_qty-=r.quantity;
  if(b.status==='lost'){i.total_qty-=r.quantity;demoAudit('item',i.id,'total_qty',i.total_qty+r.quantity,i.total_qty,'Loss');}
  i.available_qty=i.total_qty-i.rented_qty;
  demoMove(i,b.status==='lost'?'loss':'return',r.quantity,b.status==='lost'?-r.quantity:r.quantity);
  demoAudit('rental',r.id,'status','active',b.status);return;
 }
 if(!item)throw new Error('Choose an item');
 if(type==='receipt'){
  if(count<1)throw new Error('Quantity must be positive');
  item.total_qty+=count;item.available_qty=item.total_qty-item.rented_qty;
  demoMove(item,'receipt',count,count,{unit_cost:b.unit_cost,note:b.note});demoAudit('item',item.id,'total_qty',item.total_qty-count,item.total_qty);return;
 }
 if(type==='issue'){
  if(item.kind!=='consumable'||count>item.total_qty)throw new Error('Insufficient consumable stock');
  item.total_qty-=count;item.available_qty=item.total_qty;
  demoMove(item,'issue',count,-count,{project_name:d.projects.find(x=>x.id===b.project_id)?.name,employee_name:d.employees.find(x=>x.id===b.employee_id)?.name});
  demoAudit('item',item.id,'total_qty',item.total_qty+count,item.total_qty);return;
 }
 if(type==='rent'){
  if(item.kind!=='rental'||count>item.total_qty-item.rented_qty)throw new Error('Insufficient available rental equipment');
  item.rented_qty+=count;item.available_qty=item.total_qty-item.rented_qty;
  const r={id:id(),item_id:item.id,sku:item.sku,item_name:item.name,project_id:b.project_id,project_name:d.projects.find(x=>x.id===b.project_id)?.name,employee_id:b.employee_id,employee_name:d.employees.find(x=>x.id===b.employee_id)?.name,quantity:count,started_on:b.started_on,rate:item.rental_rate,rate_unit:item.rental_unit,status:'active',billed_amount:0};
  d.rentals.unshift(r);demoMove(item,'rent-out',count,-count);demoAudit('rental',r.id,'status',null,'active');return;
 }
 if(type==='adjust'){
  const target=Number(b.total_qty);if(target<item.rented_qty)throw new Error('Total cannot be less than rented quantity');
  const old=item.total_qty;item.total_qty=target;item.available_qty=target-item.rented_qty;
  demoMove(item,'adjust',Math.abs(target-old),target-old,{note:b.reason});demoAudit('item',item.id,'total_qty',old,target,b.reason);return;
 }
 throw new Error('Not implemented');
}
function nextDate(s){const d=new Date(s+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+1);return d.toISOString().slice(0,10);}
async function save(event){
 event.preventDefault();const form=event.currentTarget,type=form.dataset.type,value=form.dataset.record;
 const b=Object.fromEntries(new FormData(form).entries());
 try{
  if(state.demo){demoMutation(type,b,value);}
  else{
   const routes={item:['/api/items','POST'],edit:['/api/items/'+value,'PATCH'],receipt:['/api/stock/receipt','POST'],issue:['/api/stock/issue','POST'],rent:['/api/rentals','POST'],close:['/api/rentals/'+value+'/close','POST'],monthly:['/api/rentals/monthly-close','POST'],adjust:['/api/stock/adjust','POST'],project:['/api/projects','POST'],employee:['/api/employees','POST']};
   if(type==='edit'){
     const original=state.data.items.find(x=>x.id===value);
     b.revision=original?.revision;
     if(original?.kind==='consumable')delete b.rental_unit;
   }
   const route=routes[type];if(!route)throw new Error('Unknown operation');
   await api(route[0],b,route[1]);
  }
  $('#modal-host').innerHTML='';await reload();toast('Changes saved successfully');
 }catch(error){toast(error.message,true);}
}
document.addEventListener('click',async e=>{
 const tab=e.target.closest('[data-view]');if(tab){state.view=tab.dataset.view;state.sort='sku';state.search='';$('#search').value='';render();return;}
 const action=e.target.closest('[data-action]');if(!action)return;
 const t=action.dataset.action,v=action.dataset.value;
 if(t==='dismiss'){$('#modal-host').innerHTML='';return;}
 if(t==='view'){state.view=v;render();return;}
 if(t==='read'){
  if(state.demo){const m=state.data.messages.find(x=>x.id===v);if(m)m.is_read=true;}
  else await api('/api/system-messages/'+v+'/read',{},'PATCH');
  await reload();return;
 }
 if(['item','edit','receipt','issue','rent','close','monthly','adjust','project','employee'].includes(t))modal(t,v);
});
$('#login-demo').addEventListener('click',beginDemo);
$('#login-microsoft').addEventListener('click',async()=>{try{
 const login=await state.client.loginPopup({scopes:[cfg.apiScope]});
 state.account=login.account;await connect();
}catch(e){$('#login-status').textContent=e.message;}});
$('#logout').addEventListener('click',async()=>{if(state.client&&state.account)await state.client.logoutPopup({account:state.account}).catch(()=>{});location.reload();});
$('#search').addEventListener('input',e=>{state.search=e.target.value;render();});
$('#zoom-in').addEventListener('click',()=>changeZoom(10));
$('#zoom-out').addEventListener('click',()=>changeZoom(-10));
$('#zoom-value').addEventListener('click',()=>{state.zoom=100;render();});
$('#fullscreen').addEventListener('click',()=>{document.body.classList.toggle('fullscreen');$('#fullscreen').textContent=document.body.classList.contains('fullscreen')?'⤢':'⛶';});
$('#reset-width').addEventListener('click',()=>{localStorage.removeItem('hh-wh-widths-'+state.view);render();});
function changeZoom(by){state.zoom=Math.max(65,Math.min(150,state.zoom+by));render();}
$('#page-content').addEventListener('wheel',e=>{if(e.ctrlKey){e.preventDefault();changeZoom(e.deltaY<0?5:-5);}},{passive:false});
$('#page-content').addEventListener('click',e=>{
 if(e.target.closest('.grip'))return;
 const th=e.target.closest('th[data-sort]');if(!th||th.dataset.sort==='_actions')return;
 state.descending=state.sort===th.dataset.sort?!state.descending:false;state.sort=th.dataset.sort;render();
});
$('#page-content').addEventListener('pointerdown',e=>{
 const grip=e.target.closest('.grip');if(!grip)return;e.preventDefault();
 const th=grip.closest('th'),key=grip.dataset.grip,origin=e.clientX,start=th.getBoundingClientRect().width;
 const move=ev=>{const w=Math.max(65,Math.round(start+ev.clientX-origin));th.style.width=w+'px';};
 const up=ev=>{const w=Math.max(65,Math.round(start+ev.clientX-origin));const k='hh-wh-widths-'+state.view,all=JSON.parse(localStorage.getItem(k)||'{}');all[key]=w+'px';localStorage.setItem(k,JSON.stringify(all));window.removeEventListener('pointermove',move);window.removeEventListener('pointerup',up);render();};
 window.addEventListener('pointermove',move);window.addEventListener('pointerup',up,{once:true});
});
if('serviceWorker' in navigator)window.addEventListener('load',()=>navigator.serviceWorker.register('/sw.js').catch(()=>{}));
initMicrosoft().catch(e=>{$('#login-status').textContent='Microsoft sign-in not ready: '+e.message;});
