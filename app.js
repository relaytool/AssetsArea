(()=>{
'use strict';
const $=id=>document.getElementById(id), cfg=window.STOCK_AREA_CONFIG||{};
const SHEET=String(cfg.SPREADSHEET_ID||''), BASE=`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(SHEET)}`;
const tabs={legacy:cfg.POSITIONS_TAB||'Area Positions',activity:cfg.ACTIVITY_TAB||'Area Activity',settings:cfg.SETTINGS_TAB||'Area Settings',areas:cfg.AREAS_TAB||'Warehouse Areas',stock:cfg.STOCK_TAB||'Warehouse Stock',blockSettings:cfg.BLOCK_SETTINGS_TAB||'Warehouse Block Settings',blockStock:cfg.BLOCK_STOCK_TAB||'Warehouse Block Stock',blockTransactions:cfg.BLOCK_TRANSACTIONS_TAB||'Block Transactions'};
const C=Number(cfg.BOARD_COLUMNS)||16,MIN_R=Number(cfg.BOARD_MIN_ROWS)||18,MAX_R=Number(cfg.BOARD_MAX_ROWS)||80;
const AREA_HEADERS=['Area ID','Name','Zone','Type','X','Y','Width','Height','Color','Updated At','Active'];
const STOCK_HEADERS=['Area ID','Item','Quantity','Updated At'];
const BLOCK_SETTINGS_HEADERS=['Area ID','Rows','Updated At','Columns'];
function settingsHeaders(rows){const headers=(rows[0]||[]).map(normalizedHeader);if(!rows.length)return BLOCK_SETTINGS_HEADERS.map(normalizedHeader);if(headers[0]!=='areaid'||headers[1]!=='rows'||!headers.includes('updatedat')||!(headers.includes('columns')||headers.length===3))throw Error('Warehouse Block Settings must have Area ID, Rows, Updated At and Columns headers.');return headers;}
function settingsRow(headers,id,rows,columns,stamp){return headers.map(h=>h==='areaid'?id:h==='rows'?rows:h==='columns'?columns:h==='updatedat'?stamp:'');}
function settingsHeaderUpgrade(rows){const h=settingsHeaders(rows);return h.includes('columns')?null:{range:`'${tabs.blockSettings}'!D1`,values:[['Columns']]};}
const BLOCK_STOCK_HEADERS=['Area ID','Block ID','Item','Quantity','Updated At'];
const ACTIVITY_HEADERS=['Timestamp','Action','Item','Quantity','From','To','Note','By'];
const TX_HEADERS=['Timestamp','Action','Area','Block','Asset Type','Quantity','From','To','Reason','By','Transaction ID'];
const PRESETS=['Red sleeves','Wood Pallets','Magnum Bases','Blue Carbags'];
const now=()=>new Date().toISOString(), fmt=n=>Number(n||0).toLocaleString('en-GB'), norm=s=>String(s??'').trim().toLocaleLowerCase('en-GB');
const normalizedHeader=s=>norm(s).replace(/[^a-z0-9]/g,'');
const dateText=s=>{const d=new Date(s);return Number.isFinite(+d)?d.toLocaleString('en-GB',{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}):String(s||'')};
let state={areas:[],stock:[],events:[],blockTransactions:[],historyReady:false,legacy:false,loaded:false,legacyAreas:[],legacyStock:[],blockSettings:[],blockStock:[],blocksReady:false};
let bulkMode=false,bulkSelected=new Set(),findRequest=null,mainView='warehouse',transactionLimit=50,blockMoveDrag=null,blockMoveTimer=0;
let blockView='lanes',blockSelection=new Set(),focusedBlock='',blockAnchor='',blockClipboard=null,lastBlockGroupSize=0;
let drawingTool='select',drawGesture=null,mapZoom=1,focusZoom=1;
let mode='viewer',token='',ownerEmail='',tokenClient=null,gisReady=false,designing=false,selectedId='',selectedBlock='',editingId='',stockAction='add',filterItem='',busy=false,toastTimer=0,loading=false,queuedLoad=false,drag=null,stockPress=null,stockPressTimer=0,stockDrop=null,ignoreClickUntil=0;
const activeAreas=()=>state.areas.filter(a=>a.active);
const getArea=id=>activeAreas().find(a=>norm(a.id)===norm(id));
const looseStockAt=id=>state.stock.filter(s=>norm(s.areaId)===norm(id)&&s.quantity>0);
const blockStockAt=(areaId,blockId)=>state.blockStock.filter(s=>norm(s.areaId)===norm(areaId)&&norm(s.blockId)===norm(blockId)&&s.quantity>0);
const allAreaRows=id=>[...looseStockAt(id),...state.blockStock.filter(s=>norm(s.areaId)===norm(id)&&s.quantity>0)];
const stockAt=id=>{const groups=new Map();for(const s of allAreaRows(id)){const key=norm(s.item),old=groups.get(key);if(old)old.quantity+=s.quantity;else groups.set(key,{areaId:id,item:s.item,quantity:s.quantity});}return [...groups.values()];};
const blockRows=id=>state.blockSettings.find(s=>norm(s.areaId)===norm(id))?.rows||8;
const blockColumns=id=>state.blockSettings.find(s=>norm(s.areaId)===norm(id))?.columns||2;
// Stable Excel-like column names: A...Z, AA...AZ. Blocks retain their ID when a grid expands.
function blockColumnLabel(index){let name='';for(let n=index;n>0;n=Math.floor((n-1)/26))name=String.fromCharCode(65+(n-1)%26)+name;return name;}
function blockColumnIndex(label){let n=0;for(const char of label)n=n*26+char.charCodeAt(0)-64;return n;}
function blockCoordinates(id){const match=/^([A-Z]+)(\d{2})$/.exec(String(id||'').toUpperCase());if(!match)return null;return {col:blockColumnIndex(match[1]),row:Number(match[2])};}
function blockIds(rows,columns){const ids=[];for(let r=1;r<=rows;r++)for(let c=1;c<=columns;c++)ids.push(blockColumnLabel(c)+String(r).padStart(2,'0'));return ids;}
const validBlock=(areaId,blockId)=>{const pos=blockCoordinates(blockId);return Boolean(getArea(areaId)&&pos&&pos.row>=1&&pos.row<=blockRows(areaId)&&pos.col>=1&&pos.col<=blockColumns(areaId));};
function parseBlockSettings(table){const headers=table.headers||[];if(!headers.some(h=>normalizedHeader(h)==='areaid'))throw Error('Warehouse Block Settings headers missing');return (table.rows||[]).map(row=>{const d=rowDict(headers,row);return {areaId:String(d.areaid||'').trim(),rows:safeInt(d.rows,8),columns:safeInt(d.columns,2)||2,updatedAt:String(d.updatedat||'')}}).filter(x=>x.areaId);}
function parseBlockStock(table){const headers=table.headers||[];if(!headers.some(h=>normalizedHeader(h)==='blockid'))throw Error('Warehouse Block Stock headers missing');return (table.rows||[]).map(row=>{const d=rowDict(headers,row);return {areaId:String(d.areaid||'').trim(),blockId:String(d.blockid||'').toUpperCase().trim(),item:String(d.item||'').trim(),quantity:Math.max(0,Number(d.quantity)||0),updatedAt:String(d.updatedat||'')}}).filter(x=>x.areaId&&x.blockId&&x.item&&x.quantity>0);}
const NON_STORAGE_TYPES=new Set(['aisle','bay','exit','occupied','label','door','wall','safety','marking']);
const canHold=a=>Boolean(a&&!NON_STORAGE_TYPES.has(norm(a.type)));
const isMarking=a=>Boolean(a&&NON_STORAGE_TYPES.has(norm(a.type)));
const markerSymbols={aisle:'⇢',bay:'▤',exit:'⇥',occupied:'▧',label:'T',door:'▱',wall:'▰',safety:'⚠',marking:'✦'};
const markerClass=a=>'floor-type-'+(NON_STORAGE_TYPES.has(norm(a?.type))?norm(a.type):'storage');
const isOwner=()=>mode==='owner'&&Boolean(token);
const assetColor=name=>{const k=norm(name);return k==='red sleeves'?'red':k==='wood pallets'?'amber':k==='magnum bases'?'purple':k==='blue carbags'?'blue':''};
function notice(message='',kind=''){const el=$('notice');el.textContent=message;el.className='notice'+(kind?' '+kind:'')+(message?'':' hidden');}
function connection(message,kind=''){const e=$('connection');e.className='connection'+(kind?' '+kind:'');$('connectionText').textContent=message;}
function toast(message){const e=$('toast');e.textContent=message;e.classList.add('show');clearTimeout(toastTimer);toastTimer=setTimeout(()=>e.classList.remove('show'),3100);}
function setError(id,message){const e=$(id);e.textContent=message;e.classList.toggle('hidden',!message);}
function make(tag,className,text){const e=document.createElement(tag);if(className)e.className=className;if(text!==undefined)e.textContent=String(text);return e;}
function rowDict(headers,row){const obj={};headers.forEach((h,i)=>{obj[normalizedHeader(h)]=row?.[i]??''});return obj;}
const safeInt=(n,fallback=0)=>Number.isFinite(Number(n))?Math.trunc(Number(n)):fallback;
function parseAreas(table){const h=table.headers||[];if(!h.some(x=>normalizedHeader(x)==='areaid'))throw new Error('Warehouse Areas header missing');return table.rows.map(row=>{const d=rowDict(h,row);return {id:String(d.areaid||'').trim(),name:String(d.name||'').trim(),zone:String(d.zone||'').trim(),type:String(d.type||'Holding').trim(),x:safeInt(d.x,1),y:safeInt(d.y,1),w:safeInt(d.width,3),h:safeInt(d.height,1),color:String(d.color||'green'),updatedAt:String(d.updatedat||''),active:!['FALSE','NO','0'].includes(String(d.active||'TRUE').toUpperCase())};}).filter(a=>a.id);}
function parseStock(table){const h=table.headers||[];if(!h.some(x=>normalizedHeader(x)==='areaid')||!h.some(x=>normalizedHeader(x)==='quantity'))throw new Error('Warehouse Stock header missing');return table.rows.map(row=>{const d=rowDict(h,row);return {areaId:String(d.areaid||'').trim(),item:String(d.item||'').trim(),quantity:Math.max(0,Number(d.quantity)||0),updatedAt:String(d.updatedat||'')};}).filter(s=>s.areaId&&s.item&&s.quantity>0);}
function parseTransactions(table){
 const h=table.headers||[];
 if(!h.some(v=>normalizedHeader(v)==='timestamp')||!h.some(v=>normalizedHeader(v)==='assettype'))return [];
 return (table.rows||[]).map(row=>{const d=rowDict(h,row);return {timestamp:String(d.timestamp||''),action:String(d.action||''),area:String(d.area||''),block:String(d.block||''),item:String(d.assettype||''),quantity:Number(d.quantity)||0,from:String(d.from||''),to:String(d.to||''),reason:String(d.reason||''),by:String(d.by||''),id:String(d.transactionid||'')};}).filter(e=>e.timestamp);
}
function parseEvents(table){const h=table.headers||[];return table.rows.map(row=>{const d=rowDict(h,row);return {timestamp:String(d.timestamp||''),action:String(d.action||'Update'),item:String(d.item||''),quantity:Number(d.quantity)||0,from:String(d.from||''),to:String(d.to||''),note:String(d.note||''),by:String(d.by||'')};}).filter(e=>e.timestamp);}
function parseLegacy(table,settings){const h=table.headers||[];if(!h.some(x=>normalizedHeader(x)==='position'))throw new Error('Legacy storage header missing');const rows=table.rows||[];const hi=h.map(normalizedHeader);const canonical=['position','row','lane','item','quantity'].every(x=>hi.includes(x));const col=k=>hi.indexOf(k);
 const old=rows.map(v=>{const at=k=>v[canonical?col(k):({position:0,row:1,lane:2,item:3,quantity:4,updatedat:5}[k])]??'';const id=String(at('position')).trim();return {id,row:safeInt(at('row'),Number(id.slice(1))||1),lane:String(at('lane')||id[0]).trim(),item:String(at('item')||'').trim(),quantity:Number(at('quantity'))||0};}).filter(v=>v.id);
 let count=8;if(settings){const sh=settings.headers.map(normalizedHeader);const settingIndex=sh.indexOf('setting'),valueIndex=sh.indexOf('value');const r=settings.rows.find(x=>normalizedHeader(x[settingIndex])==='rowcount');if(r)count=safeInt(r[valueIndex],8);}count=Math.max(2,Math.min(30,count));const used=Math.max(count,...old.map(v=>v.row));
 const oldById=new Map(old.map(p=>[norm(p.id),p]));const areas=[],stock=[];for(let n=1;n<=used;n++)for(const lane of ['A','B']){const id=lane+String(n).padStart(2,'0'),p=oldById.get(norm(id));areas.push({id,name:`Block ${id}`,zone:'Original storage',type:'Holding',x:lane==='A'?1:9,y:n,w:7,h:1,color:lane==='A'?'green':'blue',active:true});if(p&&p.item&&p.quantity>0)stock.push({areaId:id,item:p.item,quantity:p.quantity,updatedAt:''});}
 return {areas,stock};}
function gviz(tab){return new Promise((resolve,reject)=>{const key=`warehouseReply_${Math.random().toString(36).slice(2)}`,script=document.createElement('script');let settled=false;const finish=(error,result)=>{if(settled)return;settled=true;clearTimeout(t);delete window[key];script.remove();error?reject(error):resolve(result);};const t=setTimeout(()=>finish(new Error('Could not load '+tab)),16000);
 window[key]=result=>{if(!result||result.status!=='ok'){finish(new Error('Cannot access '+tab));return;}finish(null,{headers:(result.table?.cols||[]).map(v=>String(v.label||v.id||'').trim()),rows:(result.table?.rows||[]).map(v=>(v.c||[]).map(c=>c?.v??''))});};
 script.src=`https://docs.google.com/spreadsheets/d/${encodeURIComponent(SHEET)}/gviz/tq?sheet=${encodeURIComponent(tab)}&headers=1&tqx=${encodeURIComponent('responseHandler:'+key)}&_=${Date.now()}`;
 script.onerror=()=>finish(new Error('Cannot load '+tab));document.head.appendChild(script);});}
async function loadData(show=true){if(loading){queuedLoad=true;return;}loading=true;if(show)connection('Refreshing…');try{
 const [areas,stock,events,blockSettings,blockStock,transactions]=await Promise.allSettled([gviz(tabs.areas),gviz(tabs.stock),gviz(tabs.activity),gviz(tabs.blockSettings),gviz(tabs.blockStock),gviz(tabs.blockTransactions)]);
 if(areas.status==='fulfilled'&&stock.status==='fulfilled'){
   state.areas=parseAreas(areas.value);state.stock=parseStock(stock.value);state.legacy=false;state.loaded=true;notice();
 }else{
   const [old,settings]=await Promise.allSettled([gviz(tabs.legacy),gviz(tabs.settings)]);
   if(old.status!=='fulfilled')throw new Error('Warehouse Areas / Warehouse Stock could not be read, and the original Area Positions tab is unavailable. Check sheet names and public Viewer access.');
   const legacy=parseLegacy(old.value,settings.status==='fulfilled'?settings.value:null);
   state.areas=legacy.areas;state.stock=legacy.stock;state.legacyAreas=legacy.areas;state.legacyStock=legacy.stock;state.legacy=true;state.loaded=true;
   if(areas.status==='fulfilled'||stock.status==='fulfilled')notice('Warehouse setup is incomplete. The original storage map is shown until the manager finishes setup.');else notice('The original storage area is shown. The manager can activate the new warehouse layout.');
 }
 if(events.status==='fulfilled'){state.events=parseEvents(events.value);}
 if(transactions.status==='fulfilled'){state.blockTransactions=parseTransactions(transactions.value);state.historyReady=true;}else{state.historyReady=false;state.blockTransactions=[];}
   // Parse independently: a missing/empty stock tab must not reset saved area grid dimensions.
   const blockIssues=[];let settingsOK=false,stockOK=false;
   if(blockSettings.status==='fulfilled')try{state.blockSettings=parseBlockSettings(blockSettings.value);settingsOK=true;}catch(e){blockIssues.push(`${tabs.blockSettings}: ${e.message}`);}
   else blockIssues.push(`${tabs.blockSettings}: cannot read tab`);
   if(blockStock.status==='fulfilled')try{state.blockStock=parseBlockStock(blockStock.value);stockOK=true;}catch(e){blockIssues.push(`${tabs.blockStock}: ${e.message}`);}
   else blockIssues.push(`${tabs.blockStock}: cannot read tab`);
   state.blocksReady=settingsOK&&stockOK;
   if(!state.blocksReady&&!state.legacy){notice('Block setup needs repair. '+blockIssues.join(' · ')+'. Manager: choose Enable area blocks to initialise empty tabs. Existing recorded stock is not changed.','error');}

 render();connection(isOwner()?'Manager connected':'Live · view only','live');openFromURL();
 }catch(err){if(state.loaded){notice('Refresh failed: '+(err.message||'Unable to reach Google Sheets')+'. Previously loaded figures may be out of date.','error');connection('Refresh failed','error');}else{notice('Unable to read warehouse data: '+(err.message||'Check Google Sheets permissions and tab names')+'.','error');connection('Not connected','error');if(document.documentElement.classList.contains('area-route')){$('areaLoading').querySelector('strong').textContent='Unable to open this area';$('areaLoading').querySelector('span').textContent='Check the Sheet sharing permissions and reload.';}render();}console.warn('Warehouse refresh:',err);}
 finally{loading=false;if(queuedLoad){queuedLoad=false;loadData(false);}}
}
function render(){const owner=isOwner();$('modePill').textContent=owner?'MANAGER EDIT':'VIEW ONLY';$('modePill').classList.toggle('editing',owner);
 $('ownerButton').textContent=owner?'Lock editing':'Manager edit';$('ownerButton').classList.toggle('btn-dark',!owner);
 $('setupButton').classList.toggle('hidden',!owner||!state.legacy);$('designButton').classList.toggle('hidden',!owner||state.legacy);$('addAreaButton').classList.toggle('hidden',!owner||state.legacy||!designing);$('enableBlocks').classList.toggle('hidden',!owner||state.legacy||state.blocksReady);
 if(!owner){designing=false;drawingTool='select';} $('designButton').textContent=designing?'Finish design':'Design layout';$('designerHint').classList.toggle('hidden',!designing);$('designPalette').classList.toggle('hidden',!designing);syncDrawingPalette();$('mapDescription').textContent=designing?'Draw bays, exits, occupied zones and labels, or select a shape to move or resize it.':owner&&!state.legacy?'Select an area, or long-press a stocked block to move its contents.':'Select an area to view its contents.';
 renderSummary();renderFindChoices();renderFindResult();renderBoard();renderList();renderActivity();renderBlockTransactions();if(selectedId&&$('detailsOverlay').classList.contains('open'))renderDetails(false);if($('mapOverlay').classList.contains('open'))renderFocusMap();
}
function renderSummary(){const groups=new Map(PRESETS.map(name=>[norm(name),{name,qty:0,areas:new Set()}]));for(const s of [...state.stock,...state.blockStock]){if(!getArea(s.areaId))continue;const key=norm(s.item),found=groups.get(key)||{name:s.item,qty:0,areas:new Set()};found.qty+=s.quantity;found.areas.add(norm(s.areaId));groups.set(key,found);}
 const order=[...groups.values()].sort((a,b)=>{const ia=PRESETS.findIndex(x=>norm(x)===norm(a.name)),ib=PRESETS.findIndex(x=>norm(x)===norm(b.name));return (ia<0?999:ia)-(ib<0?999:ib)||a.name.localeCompare(b.name);});const el=$('summary');el.replaceChildren();for(const group of order){const card=make('button','summary-card '+assetColor(group.name));card.type='button';card.setAttribute('aria-pressed',String(norm(group.name)===filterItem));card.setAttribute('aria-label',`${group.name}: ${fmt(group.qty)} in ${group.areas.size} areas`);card.append(make('span','name',group.name),make('strong','',fmt(group.qty)),make('small','',`${group.areas.size} ${group.areas.size===1?'area':'areas'}`));card.onclick=()=>{filterItem=filterItem===norm(group.name)?'':norm(group.name);render();};el.append(card);}$('clearFilter').classList.toggle('hidden',!filterItem);}
const areaFiltered=a=>Boolean(filterItem&&!stockAt(a.id).some(s=>norm(s.item)===filterItem));
// Public retrieval: find and present a route, but never change the sheet.
function renderFindChoices(){
 const select=$('findItem'),previous=select.value,names=new Map();
 for(const name of PRESETS)names.set(norm(name),name);
 for(const item of [...state.stock,...state.blockStock])if(item.quantity>0&&!names.has(norm(item.item)))names.set(norm(item.item),item.item);
 select.replaceChildren();select.append(new Option('Choose an asset',''));
 for(const name of names.values())select.append(new Option(name,name));
 const matched=[...select.options].find(o=>norm(o.value)===norm(previous)&&previous);
 if(matched)select.value=matched.value;
}
function focusFindPick(index,scrollTo='map'){
 const selected=document.querySelectorAll('#findResult .find-mini-cell.is-selected, #findResult .pick-step');
 for(const node of selected)node.classList.toggle('find-focused',Number(node.dataset.pickIndex)===index);
 const target=$('findResult').querySelector(scrollTo==='map'?`.find-mini-cell[data-pick-index="${index}"]`:`.pick-step[data-pick-index="${index}"]`);
 if(target){target.scrollIntoView({block:'center',inline:'center',behavior:'smooth'});target.focus({preventScroll:true});}
}
function renderFindMap(result,panel){
 if(!result.steps.length)return;
 const groups=result.areasUsed||[],areaGroups=new Map(groups.map((group,i)=>[norm(group.areaId),{...group,order:i+1}]));
 const section=make('section','find-visual');section.setAttribute('aria-label','Visual picking map');
 const top=make('div','find-visual-head');top.append(make('div','find-visual-title','Visual picking map'),make('span','find-visual-hint','Follow the numbered steps; orange blocks are clearances'));section.append(top);
 const overview=make('div','find-overview');
 const mapHead=make('div','find-map-heading');mapHead.append(make('strong','','Warehouse overview'),make('span','',`${groups.length} ${groups.length===1?'area':'areas'} · ${result.steps.length} ${result.steps.length===1?'block':'blocks'} on this route`));overview.append(mapHead);
 const mapScroll=make('div','find-floor-scroll'),map=make('div','warehouse-board find-overview-board');
 const visible=activeAreas(),rows=Math.max(4,...visible.map(a=>a.y+a.h+1));
 map.style.setProperty('--cols',C);map.style.setProperty('--rows',Math.min(MAX_R,rows));
 for(const area of visible){
  const group=areaGroups.get(norm(area.id)),tile=make(group?'button':'div',`map-block find-floor-area color-${area.color} ${markerClass(area)}`+(isMarking(area)&&area.w<=2?' marker-narrow':'')+(group?' find-target-area':' find-dim-area'));
  if(group){tile.type='button';tile.setAttribute('aria-label',`Stop ${group.order}: ${area.name}, empty ${group.steps.length} blocks to retrieve ${fmt(group.quantity)} ${result.item}`);tile.onclick=()=>{$(`findArea_${group.order}`)?.scrollIntoView({block:'start',behavior:'smooth'});};}
  tile.style.gridColumn=`${area.x} / span ${area.w}`;tile.style.gridRow=`${area.y} / span ${area.h}`;
  if(isMarking(area))tile.append(make('span','marker-symbol',markerSymbols[norm(area.type)]||'✦'));tile.append(make('span','block-id',area.id),make('span','block-name',area.name));
  if(group){tile.append(make('span','find-floor-badge',`STOP ${group.order} · ${group.steps.length} ${group.steps.length===1?'BLOCK':'BLOCKS'}`));}
  map.append(tile);
 }
 mapScroll.append(map);overview.append(mapScroll);
 overview.append(make('p','find-map-hint','Green = collect requested asset; amber = clear another asset first. Follow the step numbers in order.'));
 section.append(overview);
 const cards=make('div','find-area-maps');
 for(let groupIndex=0;groupIndex<groups.length;groupIndex++){
  const group=groups[groupIndex],area=getArea(group.areaId);if(!area)continue;
  const rows=blockRows(area.id),columns=blockColumns(area.id),selected=new Map(group.steps.map(step=>[norm(step.blockId),step]));
  const occupied=new Map();for(const stock of state.blockStock)if(norm(stock.areaId)===norm(area.id)&&stock.quantity>0){const key=norm(stock.blockId);occupied.set(key,(occupied.get(key)||0)+stock.quantity);}
  const card=make('section','find-area-map');card.id=`findArea_${groupIndex+1}`;card.setAttribute('aria-label',`${group.areaName} block map`);
  const header=make('div','find-area-map-head'),number=make('span','find-area-order',String(groupIndex+1)),copy=make('div','find-area-heading');
  copy.append(make('strong','',`${group.areaName} · ${group.areaId}`),make('span','',`${rows} rows × ${columns} columns · Retrieve ${fmt(group.quantity)} ${result.item}`));
  header.append(number,copy,make('span','find-block-count',`${group.steps.length} ${group.steps.length===1?'block':'blocks'}`));card.append(header);
  const wrap=make('div','find-mini-scroll'),grid=make('div','find-mini-grid');grid.style.gridTemplateColumns=`repeat(${columns},minmax(67px,1fr))`;
  for(let row=1;row<=rows;row++)for(let col=1;col<=columns;col++){
   const id=blockColumnLabel(col)+String(row).padStart(2,'0'),step=selected.get(norm(id)),total=occupied.get(norm(id))||0;
   const cell=make(step?'button':'div',`find-mini-cell${step?' is-selected'+(step.kind==='clear'?' is-clear':''):total?' is-stocked':' is-empty'}`);
   cell.dataset.block=id;
   cell.append(make('span','find-cell-id',id));
   if(step){
    const index=step.number;cell.type='button';cell.dataset.pickIndex=String(index);cell.setAttribute('aria-label',`Step ${index}, ${area.name}, ${step.kind==='clear'?'clear obstructing stock from':'collect '+fmt(step.quantity)+' '+result.item+' from'} block ${id}. Open block contents.`);
    cell.append(make('span','find-cell-pin',`#${index}`),make('span','find-cell-qty',step.kind==='clear'?'CLEAR FIRST':`Take ${fmt(step.quantity)}`));
    cell.onclick=()=>{openDetails(area.id);openBlock(id);};
   }else if(total){cell.append(make('span','find-cell-stock','Occupied'));}
   grid.append(cell);
  }
  wrap.append(grid);card.append(wrap);
  const foot=make('div','find-area-entry');foot.append(make('span','find-entry-arrow','↑'),make('span','',`ACCESS FROM ROW ${rows} (N) · MOVE TOWARDS ROW 1`));card.append(foot);
  const summary=make('div','find-area-stops');
  for(const step of group.steps){const chip=make('button','find-area-stop-chip',`#${step.number} ${step.blockId} · ${step.kind==='clear'?'CLEAR':`Take ${fmt(step.quantity)}`}`);chip.type='button';chip.onclick=()=>focusFindPick(step.number,'map');summary.append(chip);}
  card.append(summary);cards.append(card);
 }
 section.append(cards);panel.append(section);
}
function renderFindResult(){
 const panel=$('findResult');panel.replaceChildren();panel.classList.toggle('hidden',!findRequest||!state.loaded);
 if(!findRequest||!state.loaded)return;
 let result;try{
  result=window.WarehouseRetrieval.plan({areas:state.areas,blockSettings:state.blockSettings,blockStock:state.blockStock,looseStock:state.stock,item:findRequest.item,quantity:findRequest.quantity});
 }catch(err){panel.append(make('p','form-error',err.message));return;}
 const heading=make('div','find-result-head'),title=make('strong','',result.shortfall?`Route finds ${fmt(result.planned)} of ${fmt(result.requested)} ${result.item}`:`${fmt(result.blockCount)} blocks in order · collect ${fmt(result.planned)} ${result.item}`);
 heading.append(title,make('span',result.shortfall?'find-short':'find-complete',result.shortfall?'Insufficient recorded stock':'Request covered'));panel.append(heading);
 if(result.steps.length){
  const stats=make('div','find-route-stats');
  for(const [value,label] of [[result.areaCount,'Areas to visit'],[result.blockCount,'Blocks in route'],[result.planned,'Target units collected']]){const card=make('div','find-route-stat');card.append(make('strong','',fmt(value)),make('span','',label));stats.append(card);}
  panel.append(stats);
  if(result.excess)panel.append(make('p','find-surplus',`Whole-block removal: ${fmt(result.requested)} ${result.item} needed; the marked blocks contain ${fmt(result.planned)}. That is ${fmt(result.excess)} extra. All marked blocks will be emptied completely.`));
  renderFindMap(result,panel);
 }
 if(!result.steps.length)panel.append(make('p','find-message','No matching stock is assigned to storage blocks.'));
 if(result.steps.length){
  const title=make('h4','find-steps-title','Step-by-step removal order');panel.append(title);
  const list=make('div','pick-list');result.steps.forEach((step,i)=>{
   const row=make('button','pick-step'+(step.kind==='clear'?' is-clear':''));row.type='button';row.dataset.pickIndex=String(step.number);row.append(make('span','pick-number',String(i+1)));
   const details=make('div','pick-location');details.append(make('strong','',`${step.areaName||step.areaId} · ${step.areaId} / ${step.blockId}`),make('span','',step.kind==='clear'?`CLEAR FIRST · ${step.otherItems.map(s=>`${fmt(s.quantity)} ${s.item}`).join(', ')} · Opens access to the next block`:`COLLECT · Row ${step.row} · Column ${step.column}${step.mixed?' · Also empty the other assets in this block':''}`));
   row.append(details,make('span','pick-quantity',step.kind==='clear'?'Clear block':`Take ${fmt(step.quantity)}`));
   row.setAttribute('aria-label',`Step ${i+1}: ${step.kind==='clear'?'clear other assets from':`collect ${fmt(step.quantity)} ${result.item} from`} ${step.areaId}, block ${step.blockId}. Show on map.`);
   row.onclick=()=>focusFindPick(step.number,'map');list.append(row);
  });panel.append(list);
 }
 if(result.shortfall)panel.append(make('p','find-warning',`Short by ${fmt(result.shortfall)} ${result.item}. The route includes all ${fmt(result.accessible)} located target units, including those reached by clearing occupied blocks first.`));
 if(result.otherRemoved){
  const extras=new Map();for(const step of result.steps)for(const other of step.otherItems||[]){const key=norm(other.item),previous=extras.get(key)||{item:other.item,quantity:0};previous.quantity+=other.quantity;extras.set(key,previous);}
  panel.append(make('p','find-warning',`The numbered route includes ${[...extras.values()].map(x=>`${fmt(x.quantity)} ${x.item}`).join(', ')} of OTHER assets to clear or relocate. These do not count toward your ${result.item} request.`));
 }
 if(result.clearanceCount)panel.append(make('p','find-footnote',`${result.clearanceCount} clearance step(s) unlock the requested stock behind other assets. Each preceding block must actually be emptied before proceeding deeper.`));
 if(result.unassigned)panel.append(make('p','find-footnote',`${fmt(result.unassigned)} ${result.item} are not assigned to a block and cannot be routed until they have a location.`));
 panel.append(make('p','find-footnote','Follow #1, #2, #3… in order. Once a block is emptied its space is available and the next marked block in that column can be reached. No partial blocks are selected. Find is read-only; update the warehouse inventory separately as each movement actually happens. Verify physical access and where cleared non-target assets will go.'));
}

function findAssets(e){e.preventDefault();
 const item=$('findItem').value,quantity=Number($('findQty').value);
 if(!item||!Number.isSafeInteger(quantity)||quantity<1){$('findResult').classList.remove('hidden');$('findResult').textContent='Choose an asset type and enter a positive whole quantity.';return;}
 findRequest={item,quantity};renderFindResult();$('findResult').scrollIntoView({block:'nearest',behavior:'smooth'});
}
// Warehouse floor plan drawing tools use the existing Warehouse Areas schema.
// Nothing is sent to Sheets until the manager confirms the new marking in the editor.
const DRAW_DEFAULTS={Bay:{prefix:'BAY',name:'Bay',w:3,h:2,color:'blue'},Exit:{prefix:'EXIT',name:'Exit',w:2,h:1,color:'green'},Occupied:{prefix:'OCC',name:'Occupied zone',w:4,h:3,color:'amber'},Label:{prefix:'LBL',name:'Label',w:3,h:1,color:'slate'},Aisle:{prefix:'AISLE',name:'Walkway',w:4,h:1,color:'slate'},Door:{prefix:'DOOR',name:'Door',w:2,h:1,color:'blue'},Wall:{prefix:'WALL',name:'Barrier',w:4,h:1,color:'slate'},Safety:{prefix:'SAFE',name:'Safety zone',w:2,h:2,color:'amber'},Marking:{prefix:'MARK',name:'Marking',w:2,h:1,color:'purple'}};
function syncDrawingPalette(){document.querySelectorAll('[data-draw-tool]').forEach(el=>{const active=el.dataset.drawTool===drawingTool;el.classList.toggle('active',active);el.setAttribute('aria-pressed',String(active));});}
function nextMarkingId(prefix){let n=1;while(state.areas.some(a=>norm(a.id)===norm(`${prefix}-${String(n).padStart(2,'0')}`)))n++;return `${prefix}-${String(n).padStart(2,'0')}`;}
function drawCellFromEvent(e){const board=$('board'),rect=board.getBoundingClientRect(),style=getComputedStyle(board),gap=parseFloat(style.columnGap)||4,rowGap=parseFloat(style.rowGap)||4,padX=parseFloat(style.paddingLeft)||12,padY=parseFloat(style.paddingTop)||12,stepX=(rect.width-padX-(parseFloat(style.paddingRight)||12)-(C-1)*gap)/C+gap,stepY=(parseFloat(style.gridTemplateRows)||44)+rowGap;
 return {x:Math.min(C,Math.max(1,Math.floor((e.clientX-rect.left-padX)/stepX)+1)),y:Math.min(MAX_R,Math.max(1,Math.floor((e.clientY-rect.top-padY)/stepY)+1))};}
function newRect(start,end){return {x:Math.min(start.x,end.x),y:Math.min(start.y,end.y),w:Math.abs(start.x-end.x)+1,h:Math.abs(start.y-end.y)+1};}
function drawPreview(rect){if(!drawGesture)return;const el=drawGesture.preview;el.style.gridColumn=`${rect.x} / span ${rect.w}`;el.style.gridRow=`${rect.y} / span ${rect.h}`;el.classList.toggle('invalid',Boolean(validateRect(rect)));drawGesture.rect=rect;}
function drawStart(e){if(!isOwner()||!designing||drawingTool==='select'||busy||e.button!==0||e.target.closest('.map-block,button'))return;const board=$('board'),start=drawCellFromEvent(e),preview=make('div','draw-preview');preview.append(make('span','',drawingTool));board.append(preview);drawGesture={start,preview,pointerId:e.pointerId,rect:{...start,w:1,h:1}};drawPreview(drawGesture.rect);board.setPointerCapture(e.pointerId);e.preventDefault();}
function drawMove(e){if(!drawGesture||drawGesture.pointerId!==e.pointerId)return;drawPreview(newRect(drawGesture.start,drawCellFromEvent(e)));e.preventDefault();}
function cancelDrawing(){if(drawGesture){drawGesture.preview.remove();drawGesture=null;}}
function drawEnd(e){if(!drawGesture||drawGesture.pointerId!==e.pointerId)return;const gesture=drawGesture,tool=drawingTool;let rect=gesture.rect;const def=DRAW_DEFAULTS[tool];if(rect.w===1&&rect.h===1&&def){const suggested={...rect,w:Math.min(def.w,C-rect.x+1),h:Math.min(def.h,MAX_R-rect.y+1)};if(!validateRect(suggested))rect=suggested;}
 cancelDrawing();ignoreClickUntil=Date.now()+450;const issue=validateRect(rect);if(issue){toast(issue);return;}if(!def)return;openAreaEditor('',{type:tool,id:nextMarkingId(def.prefix),name:`${def.name} ${String(state.areas.filter(a=>norm(a.type)===norm(tool)).length+1)}`,color:def.color,rect});}
function applyMapSize(which){const target=$(which),scroll=which==='board'?$('boardScroll'):$('mapCanvasScroll');if(!target||!scroll)return;const factor=which==='board'?mapZoom:focusZoom,available=Math.max(1,scroll.clientWidth);target.style.width=Math.round(available*factor)+'px';target.style.minWidth='0';target.classList.toggle('map-board-compact',available*factor<540);}
function adjustMapZoom(which,delta){const factor=which==='board'?mapZoom:focusZoom,next=delta===0?1:Math.min(3.5,Math.max(1,Math.round((factor+delta*.35)*100)/100));if(which==='board')mapZoom=next;else focusZoom=next;const scroller=which==='board'?$('boardScroll'):$('mapCanvasScroll'),oldWidth=Math.max(1,$(which).offsetWidth),mid=(scroller.scrollLeft+scroller.clientWidth/2)/oldWidth;applyMapSize(which);scroller.scrollLeft=Math.max(0,$(which).offsetWidth*mid-scroller.clientWidth/2);}
function renderBoard(){const areas=activeAreas(),b=$('board');const rows=Math.max(MIN_R,...areas.map(a=>a.y+a.h+1));b.style.setProperty('--cols',C);b.style.setProperty('--rows',Math.min(MAX_R,rows));b.classList.toggle('designing',designing);b.classList.toggle('drawing-tool',designing&&drawingTool!=='select');b.replaceChildren();applyMapSize('board');
 for(const area of areas){const occupied=stockAt(area.id),special=!canHold(area),el=make('div','map-block '+(occupied.length?'occupied ':'')+(special?'special ':'')+'color-'+area.color+' '+markerClass(area)+(isMarking(area)&&area.w<=2?' marker-narrow':'')+(areaFiltered(area)?' filtered':''));el.tabIndex=0;el.setAttribute('role','button');el.dataset.id=area.id;el.setAttribute('aria-label',isMarking(area)?`${area.type}: ${area.name}`:`${area.id}, ${area.name}: ${occupied.length} item types`);el.title=`${area.type} · ${area.name}`;el.style.gridColumn=`${area.x} / span ${area.w}`;el.style.gridRow=`${area.y} / span ${area.h}`;if(special)el.append(make('span','marker-symbol',markerSymbols[norm(area.type)]||'✦'));el.append(make('span','block-id',area.id),make('span','block-name',area.name),make('span','block-status',special?area.type:occupied.length?`${occupied.length} ${occupied.length===1?'type':'types'} stored`:'Empty'));
 if(designing){const handle=make('button','resize-handle','⤡');handle.type='button';handle.title='Resize area';handle.setAttribute('aria-label',`Resize ${area.id}`);el.append(handle);el.addEventListener('pointerdown',e=>{if(drawingTool!=='select')return;startDrag(e,area,el);});}else if(isOwner()&&!state.legacy&&occupied.length){el.addEventListener('pointerdown',e=>startStockPress(e,area,el));}
 el.addEventListener('click',()=>{if(Date.now()<ignoreClickUntil)return;if(designing){if(drawingTool==='select')openAreaEditor(area.id);}else openDetails(area.id);});el.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();if(designing)openAreaEditor(area.id);else openDetails(area.id);}});b.append(el);
 }
 const storages=areas.filter(canHold).length,markings=areas.length-storages;$('mapCount').textContent=`${storages} storage · ${markings} markings`;$('boardEmpty').classList.toggle('hidden',areas.length>0);
}
function renderList(){const wrap=$('areaList');wrap.replaceChildren();const search=norm($('search').value);const areas=activeAreas().filter(a=>canHold(a)&&(!search||norm([a.id,a.name,a.zone,a.type,...stockAt(a.id).map(s=>s.item)].join(' ')).includes(search))).sort((a,b)=>a.zone.localeCompare(b.zone)||a.id.localeCompare(b.id,undefined,{numeric:true}));
 for(const a of areas){const items=stockAt(a.id),row=make('button','area-list-item'+(areaFiltered(a)?' filtered':''));row.type='button';const left=make('div');left.append(make('div','title',`${a.id} · ${a.name}`),make('div','sub',[a.zone||a.type,items.length?`${items.length} item ${items.length===1?'type':'types'}`:'Empty'].join(' · ')));row.append(left,make('span','arrow','›'));row.onclick=()=>designing?openAreaEditor(a.id):openDetails(a.id);wrap.append(row);}
 if(!areas.length)wrap.append(make('p','none',search?'No matching areas.':'No areas have been created yet.'));
}
function renderActivity(){const els=[...state.events].sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));$('activityCount').textContent=`${els.length} events`;const target=$('activityList');target.replaceChildren();if(!els.length){target.append(make('p','content-empty','No movements recorded yet.'));return;}for(const e of els.slice(0,25)){const row=make('div','event'),action=norm(e.action);row.append(make('div','event-icon',action.includes('move')?'⇄':action.includes('clear')?'×':action.includes('remove')?'−':'+'));const info=make('div');let sentence='';if(action.includes('move'))sentence=`Moved ${fmt(e.quantity)} ${e.item} · ${e.from} → ${e.to}`;else if(action.includes('clear'))sentence=`Cleared ${fmt(e.quantity)} ${e.item} · ${e.from}`;else if(action.includes('remove'))sentence=`Removed ${fmt(e.quantity)} ${e.item} · ${e.from}`;else if(action.includes('add'))sentence=`Added ${fmt(e.quantity)} ${e.item} · ${e.to}`;else sentence=[e.action,e.item,e.from||e.to].filter(Boolean).join(' · ');info.append(make('div','event-main',sentence),make('div','event-meta',[dateText(e.timestamp),e.note].filter(Boolean).join(' · ')));row.append(info);target.append(row);}}
function showMainView(view){
 mainView=view==='transactions'?'transactions':'warehouse';
 const showing=mainView==='warehouse';
 for(const id of ['findPanel','activity'])$(id).classList.toggle('hidden',!showing);
 for(const selector of ['.intro','.summary-section','.workspace']){const element=document.querySelector(selector);if(element)element.classList.toggle('hidden',!showing);}
 $('transactionsPage').classList.toggle('hidden',showing);
 for(const [id,on] of [['warehouseTab',showing],['transactionsTab',!showing]]){$(id).classList.toggle('active',on);if(on)$(id).setAttribute('aria-current','page');else $(id).removeAttribute('aria-current');}
 if(!showing)renderBlockTransactions();else requestAnimationFrame(()=>applyMapSize('board'));
}
function renderBlockTransactions(){
 const list=$('transactionsList');list.replaceChildren();
 const type=$('transactionType').value,search=norm($('transactionSearch').value);
 const history=state.blockTransactions.filter(e=>(type==='all'||(type==='remove'?/remove|clear/i.test(e.action):type==='move'?/move/i.test(e.action):/add/i.test(e.action)))&&(!search||norm([e.timestamp,e.action,e.area,e.block,e.item,e.quantity,e.from,e.to,e.reason,e.by].join(' ')).includes(search))).sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));
 $('transactionCount').textContent=`${fmt(history.length)} ${history.length===1?'transaction':'transactions'}`;
 if(!history.length){list.append(make('p','content-empty',state.historyReady?'No block transactions match your filters.':'The Block Transactions sheet has not been created yet. It is created automatically when the owner next saves a block movement. Earlier movements remain in Recent movements.'));$('transactionsMore').classList.add('hidden');return;}
 for(const e of history.slice(0,transactionLimit)){
  const row=make('article','transaction-row'),left=make('div'),right=make('div','tx-qty');
  const header=make('div','tx-main'),tag=make('span','tx-action '+(/remove|clear/i.test(e.action)?'remove':/move/i.test(e.action)?'move':''),e.action);
  header.append(tag,document.createTextNode(`${e.item} · ${e.area}/${e.block}`));
  left.append(header,make('div','tx-meta',[dateText(e.timestamp),e.from&&'From '+e.from,e.to&&'To '+e.to,e.by].filter(Boolean).join(' · ')));
  if(e.reason)left.append(make('div','tx-reason','Reason: '+e.reason));
  right.textContent=fmt(e.quantity);row.append(left,right);list.append(row);
 }
 $('transactionsMore').classList.toggle('hidden',history.length<=transactionLimit);
}
// Owner-only one-time setup. The existing Area Activity tab is retained.
async function ensureBlockTransactions(){
 const meta=await api('?fields=sheets(properties(title))'),known=new Set((meta.sheets||[]).map(x=>x.properties.title));
 if(!known.has(tabs.blockTransactions))await api(':batchUpdate',{method:'POST',body:JSON.stringify({requests:[{addSheet:{properties:{title:tabs.blockTransactions,gridProperties:{rowCount:1000,columnCount:12}}}}]})});
 let data=await readApiTab(tabs.blockTransactions,'A:K');
 if(!data.length||!data[0]?.some(x=>String(x||'').trim())){await apiWrites([{range:`'${tabs.blockTransactions}'!A1:K1`,values:[TX_HEADERS]}]);data=await readApiTab(tabs.blockTransactions,'A:K');}
 verifyApiHeader(data,TX_HEADERS);
 state.historyReady=true;
}
async function appendBlockTransactions(events){
 if(!events.length)return;
 const payload=events.map((e,i)=>{
  const location=/^([^/]+)\/([^/]+)$/.exec(String(e.from||''))||/^([^/]+)\/([^/]+)$/.exec(String(e.to||''));
  const area=location?.[1]||'',block=location?.[2]||'';
  return [e.timestamp,e.action,area,block,e.item,e.quantity,e.from||'',e.to||'',e.reason||'',ownerEmail,`${e.timestamp}-${i+1}`];
 });
 await api(`/values/${encodeURIComponent("'"+tabs.blockTransactions+"'!A:K")}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,{method:'POST',body:JSON.stringify({values:payload})});
 for(const row of payload)state.blockTransactions.push({timestamp:row[0],action:row[1],area:row[2],block:row[3],item:row[4],quantity:row[5],from:row[6],to:row[7],reason:row[8],by:row[9],id:row[10]});
 renderBlockTransactions();
}
function areaURL(id){const u=new URL(window.location.href);u.searchParams.set('area',id);u.searchParams.delete('block');u.hash='';return u.toString();}
function changeURL(id,replace=false){const u=new URL(location.href);if(id){u.searchParams.set('area',id);u.searchParams.delete('block');}else{u.searchParams.delete('area');u.searchParams.delete('block');}history[replace?'replaceState':'pushState']({},'',u.pathname+u.search+u.hash);}
function focusRoute(on){document.documentElement.classList.toggle('area-route',on);document.body.classList.toggle('area-focus',on);$('areaLoading').classList.toggle('hidden',!on);}
function openFromURL(){const q=new URLSearchParams(location.search),id=q.get('area');if(!id){if(selectedId)closeDetails(false);else focusRoute(false);return;}if(!state.loaded)return;const a=getArea(id);if(a){if(selectedId!==a.id)openDetails(a.id,false);else focusRoute(true);const block=q.get('block');if(block&&validBlock(a.id,block)){if(selectedBlock!==block)openBlock(block,false);}else if(selectedBlock)closeBlock(false);}else{toast(`Area ${id} was not found.`);changeURL('',true);closeDetails(false);focusRoute(false);notice(`Area ${id} was not found. Check the link or select an area from the warehouse overview.`,'error');}}
function openDetails(id,update=true){const a=getArea(id);if(!a)return;showMainView('warehouse');if(norm(selectedId)!==norm(a.id)){bulkMode=false;bulkSelected.clear();blockSelection.clear();focusedBlock='';blockAnchor='';}selectedId=a.id;selectedBlock='';stockAction='add';$('stockForm').reset();$('stockQty').value='1';setError('stockError','');$('linkStatus').textContent='';renderDetails(true);focusRoute(true);$('areaLoading').classList.add('hidden');openOverlay('detailsOverlay');if(update)changeURL(a.id);}
function renderDetails(resetStock=false){const a=getArea(selectedId);if(!a){closeDetails(true);return;}$('areaType').textContent=a.type.toUpperCase();$('areaTitle').textContent=a.name;$('areaSubtitle').textContent=a.id; $('areaViewMode').textContent=isOwner()?'MANAGER EDIT':'VIEW ONLY';$('areaOwnerButton').textContent=isOwner()?'Lock editing':'Manager edit';
 const tags=$('areaMeta');tags.replaceChildren();for(const t of [a.id,a.zone,a.type].filter(Boolean))tags.append(make('span','',t));const items=stockAt(a.id).sort((a,b)=>a.item.localeCompare(b.item));renderBlockGrid(a.id);$('areaItemCount').textContent=`${items.length} ${items.length===1?'type':'types'}`;const wrap=$('contents');wrap.replaceChildren();for(const s of items){const row=make('div','content-row');row.append(make('span','content-name',s.item),make('span','content-qty',fmt(s.quantity)));wrap.append(row);}if(!items.length)wrap.append(make('div','content-empty','Nothing stored in this area.'));
 const enabled=isOwner()&&!state.legacy&&canHold(a)&&looseStockAt(a.id).length>0;$('ownerStockPanel').classList.toggle('hidden',!enabled);$('clearArea').disabled=!items.length;$('transferPanel').classList.toggle('hidden',!looseStockAt(a.id).length);if(items.length){const dest=$('transferTo'),previous=dest.value;dest.replaceChildren();dest.append(new Option('Choose destination area',''));for(const other of activeAreas().filter(other=>norm(other.id)!==norm(a.id)&&canHold(other)).sort((x,y)=>x.id.localeCompare(y.id,undefined,{numeric:true})))dest.append(new Option(`${other.id} · ${other.name}`,other.id));if([...dest.options].some(o=>o.value===previous))dest.value=previous;}
 if(resetStock){$('stockForm').reset();$('stockQty').value='1';setStockAction('add');}else if(stockAction==='remove')setRemoveOptions();}
function closeDetails(update=true){bulkMode=false;bulkSelected.clear();blockSelection.clear();focusedBlock='';blockAnchor='';if(update&&new URLSearchParams(location.search).has('area'))changeURL('');closeOverlay('mapOverlay');closeOverlay('blockOverlay');closeOverlay('detailsOverlay');selectedId='';selectedBlock='';focusRoute(false);}
function openOverlay(id){$(id).classList.add('open');$(id).setAttribute('aria-hidden','false');document.body.style.overflow='hidden';}
function closeOverlay(id){$(id).classList.remove('open');$(id).setAttribute('aria-hidden','true');if(!document.querySelector('.overlay.open'))document.body.style.overflow='';}
function setRemoveOptions(){const chosen=$('itemType').value,items=looseStockAt(selectedId);$('itemType').replaceChildren();$('itemType').append(new Option('Choose an asset',''));for(const s of items)$('itemType').append(new Option(`${s.item} (${fmt(s.quantity)} here)`,s.item));if(items.some(s=>s.item===chosen))$('itemType').value=chosen;else $('itemType').value=items.length===1?items[0].item:'';$('customWrap').classList.add('hidden');$('customItem').required=false;}
function setAddOptions(){const chosen=$('itemType').value;$('itemType').replaceChildren();$('itemType').append(new Option('Choose an asset',''));const items=[...PRESETS];for(const s of looseStockAt(selectedId))if(!items.some(x=>norm(x)===norm(s.item)))items.push(s.item);items.forEach(s=>$('itemType').append(new Option(s,s)));$('itemType').append(new Option('Other / custom','__other__'));if(items.includes(chosen))$('itemType').value=chosen;else $('itemType').value='';toggleCustom();}
function toggleCustom(){const showing=stockAction==='add'&&$('itemType').value==='__other__';$('customWrap').classList.toggle('hidden',!showing);$('customItem').required=showing;}
function setStockAction(next){stockAction=next;document.querySelectorAll('[data-action]').forEach(el=>el.classList.toggle('selected',el.dataset.action===next));if(next==='add')setAddOptions();else setRemoveOptions();$('saveStock').textContent=next==='add'?'Add to area':'Remove from area';setError('stockError','');}
function validateRect(a,ignore=''){if(!Number.isInteger(a.x)||!Number.isInteger(a.y)||!Number.isInteger(a.w)||!Number.isInteger(a.h)||a.x<1||a.y<1||a.w<1||a.h<1||a.x+a.w-1>C||a.y+a.h-1>MAX_R)return 'The position or size is outside the warehouse grid.';for(const b of activeAreas()){if(norm(b.id)===norm(ignore))continue;if(a.x<b.x+b.w&&a.x+a.w>b.x&&a.y<b.y+b.h&&a.y+a.h>b.y)return `Overlaps area ${b.id}. Choose another position.`;}return '';}
function freeSpot(w=4,h=2){for(let y=1;y<=MAX_R-h+1;y++)for(let x=1;x<=C-w+1;x++){const a={x,y,w,h};if(!validateRect(a))return a;}return {x:1,y:1,w,h};}
function openAreaEditor(id='',defaults=null){if(!isOwner()||state.legacy)return;const a=id?getArea(id):null;editingId=a?.id||'';const pos=a||defaults?.rect||freeSpot();const requestedType=a?.type||defaults?.type||'Holding';$('editorTitle').textContent=a?(isMarking(a)?'Edit floor marking':'Edit area'):(isMarking({type:requestedType})?'New floor marking':'New holding area');$('editId').value=a?.id||defaults?.id||'';$('editId').readOnly=Boolean(a);$('editName').value=a?.name||defaults?.name||'';$('editZone').value=a?.zone||'';$('editType').value=requestedType;$('editX').value=pos.x;$('editY').value=pos.y;$('editW').value=pos.w;$('editH').value=pos.h;$('editColor').value=a?.color||defaults?.color||'green';$('editBlockRows').value=a?blockRows(a.id):8;$('editBlockColumns').value=a?blockColumns(a.id):2;syncEditorType();$('archiveArea').classList.toggle('hidden',!a);setError('editorError','');openOverlay('editorOverlay');}
function closeEditor(){closeOverlay('editorOverlay');editingId='';}
function syncEditorType(){const marking=isMarking({type:$('editType').value});$('editorBlockSettings').classList.toggle('hidden',Boolean(editingId)||!state.blocksReady||marking);$('editorTypeHint').textContent=marking?'Map marking only — no inventory or block grid. For a text label, enter the wording in Area name. Drag its edges on the map to resize later.':'Storage area: place in free grid cells. Markings and storage areas cannot overlap.';}
function collectEditor(){return {id:$('editId').value.trim().toUpperCase(),name:$('editName').value.trim(),zone:$('editZone').value.trim(),type:$('editType').value,x:Number($('editX').value),y:Number($('editY').value),w:Number($('editW').value),h:Number($('editH').value),color:$('editColor').value,updatedAt:now(),active:true};}
function startDrag(e,a,el){if(!isOwner()||!designing||busy||e.button!==0)return;const resize=Boolean(e.target.closest('.resize-handle'));const board=$('board'),bounds=board.getBoundingClientRect(),computed=getComputedStyle(board);const gapX=parseFloat(computed.columnGap)||4,gapY=parseFloat(computed.rowGap)||4,paddingX=(parseFloat(computed.paddingLeft)||0)+(parseFloat(computed.paddingRight)||0);const cellX=(bounds.width-paddingX-(C-1)*gapX)/C+gapX;const cellY=(parseFloat(computed.gridTemplateRows)||44)+gapY;drag={id:a.id,resize,initial:{...a},startX:e.clientX,startY:e.clientY,cellX,cellY,element:el,changed:false};el.setPointerCapture(e.pointerId);el.classList.add('dragging');e.preventDefault();}
function dragMove(e){if(!drag)return;const dx=Math.round((e.clientX-drag.startX)/drag.cellX),dy=Math.round((e.clientY-drag.startY)/drag.cellY);const a=drag.initial;const next=drag.resize?{...a,w:Math.min(C-a.x+1,Math.max(1,a.w+dx)),h:Math.min(MAX_R-a.y+1,Math.max(1,a.h+dy))}:{...a,x:Math.min(C-a.w+1,Math.max(1,a.x+dx)),y:Math.min(MAX_R-a.h+1,Math.max(1,a.y+dy))};drag.next=next;drag.changed=next.x!==a.x||next.y!==a.y||next.w!==a.w||next.h!==a.h;drag.element.style.gridColumn=`${next.x} / span ${next.w}`;drag.element.style.gridRow=`${next.y} / span ${next.h}`;drag.element.classList.toggle('invalid',Boolean(validateRect(next,a.id)));}
async function dragEnd(){if(!drag)return;const d=drag;drag=null;d.element.classList.remove('dragging','invalid');if(!d.changed){d.element.style.gridColumn=`${d.initial.x} / span ${d.initial.w}`;d.element.style.gridRow=`${d.initial.y} / span ${d.initial.h}`;return;}ignoreClickUntil=Date.now()+600;const problem=validateRect(d.next,d.id);if(problem){toast(problem);renderBoard();return;}try{await saveAreaRemote({...d.next,updatedAt:now()});toast(`Area ${d.id} updated.`);}catch(err){toast(err.message);render();}}

function clearStockDrag(){clearTimeout(stockPressTimer);stockPressTimer=0;document.querySelectorAll('.stock-dragging,.stock-drop,.stock-invalid').forEach(el=>el.classList.remove('stock-dragging','stock-drop','stock-invalid'));stockPress=null;stockDrop=null;}
function startStockPress(e,area,el){if(!isOwner()||designing||busy||!canHold(area)||e.button!==0)return;clearStockDrag();stockPress={id:area.id,x:e.clientX,y:e.clientY,el,active:false};stockPressTimer=setTimeout(()=>{if(!stockPress||stockPress.id!==area.id)return;stockPress.active=true;el.classList.add('stock-dragging');if(navigator.vibrate)try{navigator.vibrate(20);}catch{}},350);}
function stockPointerMove(e){if(!stockPress)return;const p=stockPress;if(!p.active){if(Math.abs(e.clientX-p.x)>11||Math.abs(e.clientY-p.y)>11)clearStockDrag();return;}e.preventDefault();const target=document.elementFromPoint(e.clientX,e.clientY)?.closest('.map-block');if(stockDrop)stockDrop.classList.remove('stock-drop','stock-invalid');stockDrop=target;if(target){const id=target.dataset.id;target.classList.add(norm(id)!==norm(p.id)&&canHold(getArea(id))?'stock-drop':'stock-invalid');}}
function stockPointerUp(e){if(!stockPress)return;const p=stockPress;if(!p.active){clearStockDrag();return;}const target=document.elementFromPoint(e.clientX,e.clientY)?.closest('.map-block'),to=target?.dataset?.id;clearStockDrag();ignoreClickUntil=Date.now()+600;if(to&&norm(to)!==norm(p.id)&&canHold(getArea(to)))moveContents(p.id,to);else toast('Drop onto a different storage area.');}

function mapStock(items){return new Map(items.map(s=>[norm(s.item),{...s}]));}
function equalStock(a,b){const x=mapStock(a),y=mapStock(b);if(x.size!==a.length||y.size!==b.length||x.size!==y.size)return false;for(const [key,val] of x){if(!y.has(key)||Number(y.get(key).quantity)!==Number(val.quantity))return false;}return true;}
function api(path,options={}){if(!isOwner())throw new Error('Manager sign-in required.');return fetch(BASE+path,{...options,headers:{Authorization:'Bearer '+token,'Content-Type':'application/json',...(options.headers||{})}}).then(async r=>{const text=await r.text();let body={};try{body=JSON.parse(text);}catch{}if(!r.ok)throw new Error(body.error?.message||`Google Sheets error (${r.status})`);return body;});}
const apiRange=(tab,cols)=>`/values/${encodeURIComponent("'"+tab+"'!"+cols)}?valueRenderOption=UNFORMATTED_VALUE`;
const apiWrites=ranges=>api('/values:batchUpdate',{method:'POST',body:JSON.stringify({valueInputOption:'RAW',data:ranges})});
async function readApiTab(tab,cols){const result=await api(apiRange(tab,cols));return result.values||[];}
function verifyApiHeader(rows,expected){if(!rows.length||expected.some((s,i)=>normalizedHeader(rows[0][i])!==normalizedHeader(s)))throw new Error('Worksheet columns were changed. Restore the expected column headers before editing.');}
function tableFromApi(rows){return {headers:rows[0]||[],rows:rows.slice(1)};}
function assertNoDuplicateStock(rows){const seen=new Set();for(const s of rows){const k=norm(s.areaId)+'|'+norm(s.item);if(seen.has(k))throw new Error('Duplicate asset rows exist in the sheet. Resolve these before making changes.');seen.add(k);}}
async function appendEvents(events){if(!events.length)return;await api(`/values/${encodeURIComponent("'"+tabs.activity+"'!A:H")}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,{method:'POST',body:JSON.stringify({values:events.map(e=>[e.timestamp,e.action,e.item,e.quantity,e.from||'',e.to||'',e.note||'',ownerEmail])})});state.events.push(...events);}
async function transactStock(areaIds,modifier,makeEvents){if(!isOwner()||state.legacy)throw new Error('Activate the warehouse layout first.');if(busy)throw new Error('A save is already in progress.');busy=true;connection('Saving…');try{
 const rows=await readApiTab(tabs.stock,'A:D');verifyApiHeader(rows,STOCK_HEADERS);const all=parseStock(tableFromApi(rows));assertNoDuplicateStock(all);
 for(const id of areaIds){if(!equalStock(all.filter(s=>norm(s.areaId)===norm(id)),looseStockAt(id)))throw new Error('Stock changed in another session. Refresh and try again.');}
 const desired=new Map(areaIds.map(id=>[norm(id),mapStock(all.filter(s=>norm(s.areaId)===norm(id)))]));const before=new Map([...desired].map(([id,items])=>[id,new Map(items)]));modifier(desired);
 const targets=[];for(const [id,items] of desired){for(const s of items.values()){if(!Number.isSafeInteger(s.quantity)||s.quantity<0)throw new Error('Quantity must be a non-negative whole number.');}const original=before.get(id);for(const [key,entry] of original){const target=items.get(key);if(!target||!target.quantity)targets.push({id,original:entry,desired:null,key});else if(target.quantity!==entry.quantity)targets.push({id,original:entry,desired:target,key});}for(const [key,entry] of items){if(!original.has(key)&&entry.quantity>0)targets.push({id,original:null,desired:entry,key});}}
 if(!targets.length){toast('Nothing to change.');return;}
 const existing=new Map();const blanks=[];for(let i=1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[0]),item=norm(r[1]);if(id&&item&&Number(r[2])>0)existing.set(id+'|'+item,i+1);else blanks.push(i+1);}let nextRow=rows.length+1;const writes=[],stamp=now();for(const target of targets){const key=target.id+'|'+target.key;let row=existing.get(key);if(!row){row=blanks.length?blanks.shift():nextRow++;}const s=target.desired;writes.push({range:`'${tabs.stock}'!A${row}:D${row}`,values:[[s?s.areaId:'',s?s.item:'',s?s.quantity:0,stamp]]});}
 await apiWrites(writes);const checked=await readApiTab(tabs.stock,'A:D');verifyApiHeader(checked,STOCK_HEADERS);const updated=parseStock(tableFromApi(checked));assertNoDuplicateStock(updated);for(const id of areaIds){const expected=[...desired.get(norm(id)).values()].filter(s=>s.quantity>0);if(!equalStock(expected,updated.filter(s=>norm(s.areaId)===norm(id))))throw new Error('Saved rows could not be verified. Refresh the sheet and check stock.');}
 state.stock=updated;const eventRows=makeEvents(before,desired,stamp);let logged=true;try{await appendEvents(eventRows);}catch(err){logged=false;console.warn('Activity could not be appended',err);}render();connection('Manager connected','live');toast(logged?'Saved to Google Sheets.':'Stock saved; activity log could not be updated.');
 }catch(err){connection('Save failed','error');throw err;}finally{busy=false;}}
async function addOrRemove(e){e.preventDefault();if(!selectedId||busy)return;setError('stockError','');const id=selectedId,qty=Number($('stockQty').value),chosen=$('itemType').value,item=stockAction==='add'&&chosen==='__other__'?$('customItem').value.trim():chosen;if(!Number.isSafeInteger(qty)||qty<1){setError('stockError','Enter a valid whole quantity.');return;}if(!item||item.length>90){setError('stockError','Choose an asset type.');return;}
 $('saveStock').disabled=true;try{await transactStock([id],desired=>{const m=desired.get(norm(id)),key=norm(item),old=m.get(key);if(stockAction==='remove'){if(!old||old.quantity<qty)throw new Error('Not enough of that item in this area.');m.set(key,{...old,quantity:old.quantity-qty});}else{const stored=old?.item||item;m.set(key,{areaId:id,item:stored,quantity:(old?.quantity||0)+qty});}},(_prev,_next,timestamp)=>[{timestamp,action:stockAction==='add'?'Add':'Remove',item,quantity:qty,from:stockAction==='remove'?id:'',to:stockAction==='add'?id:'',note:''}]);renderDetails(true);
 }catch(err){setError('stockError',err.message||'Could not save stock.');}finally{$('saveStock').disabled=false;}}
async function clearArea(){const a=getArea(selectedId);if(!a||!isOwner()||busy)return;const items=stockAt(a.id);if(!items.length)return;const list=items.map(s=>`• ${s.item}: ${fmt(s.quantity)}`).join('\n');if(!confirm(`Empty ALL stock from ${a.name} (${a.id})?\n\n${list}\n\nThis will remove every item type from this area and record the movement.`))return;
 $('clearArea').disabled=true;try{await clearEntireArea(a.id,$('areaClearReason').value.trim());$('areaClearReason').value='';renderDetails(true);}catch(err){toast(err.message||'Could not clear area.');}finally{$('clearArea').disabled=false;}}
async function moveContents(fromId,toId){const a=getArea(fromId),b=getArea(toId);if(!a||!b||!canHold(b)||!looseStockAt(fromId).length)return;if(!confirm(`Move all stock from ${a.id} to ${b.id}? Items already in ${b.id} will be combined by asset type.`))return;
 try{await transactStock([fromId,toId],desired=>{const src=desired.get(norm(fromId)),dst=desired.get(norm(toId));for(const [key,item] of src){const old=dst.get(key);dst.set(key,{areaId:toId,item:old?.item||item.item,quantity:(old?.quantity||0)+item.quantity});}src.clear();},(before,_after,timestamp)=>[...before.get(norm(fromId)).values()].map(s=>({timestamp,action:'Move',item:s.item,quantity:s.quantity,from:fromId,to:toId,note:'All contents moved'})));}catch(err){toast(err.message);}}
async function saveAreaRemote(a){if(!isOwner()||state.legacy)throw new Error('Manager edit required.');if(busy)throw new Error('Another save is in progress.');const problem=validateRect(a,a.id);if(problem)throw new Error(problem);busy=true;connection('Saving layout…');try{
 const rows=await readApiTab(tabs.areas,'A:K');verifyApiHeader(rows,AREA_HEADERS);const all=parseAreas(tableFromApi(rows));const old=all.find(v=>norm(v.id)===norm(a.id));const previously=activeAreas().find(v=>norm(v.id)===norm(a.id));if(previously&&(!old||JSON.stringify([old.x,old.y,old.w,old.h,old.name,old.zone,old.type,old.color])!==JSON.stringify([previously.x,previously.y,previously.w,previously.h,previously.name,previously.zone,previously.type,previously.color])))throw new Error('Layout changed in another session. Refresh before editing.');if(!previously&&old)throw new Error('That area ID is already used.');
 let rowIndex=old?all.findIndex(v=>norm(v.id)===norm(a.id))+2:0;if(!rowIndex){const vacant=rows.findIndex((r,i)=>i>0&&!String(r[0]||'').trim());rowIndex=vacant>=0?vacant+1:rows.length+1;}
 await apiWrites([{range:`'${tabs.areas}'!A${rowIndex}:K${rowIndex}`,values:[[a.id,a.name,a.zone,a.type,a.x,a.y,a.w,a.h,a.color,a.updatedAt,'TRUE']]}]);
 const verified=await readApiTab(tabs.areas,'A:K');const next=parseAreas(tableFromApi(verified));const updated=next.find(v=>norm(v.id)===norm(a.id));if(!updated||!updated.active||updated.x!==a.x||updated.y!==a.y||updated.w!==a.w||updated.h!==a.h||updated.name!==a.name)throw new Error('Area save was not confirmed. Please refresh.');state.areas=next;render();connection('Manager connected','live');
 }catch(err){connection('Save failed','error');throw err;}finally{busy=false;}}
async function setInitialBlockGrid(areaId,rows,columns){
 if(rows===8&&columns===2)return;
 const data=await readApiTab(tabs.blockSettings,'A:D');verifyBlockSettingsHeader(data);
 if(data.some((r,i)=>i>0&&norm(r[0])===norm(areaId)))throw Error('The new area has grid settings already. Open it to check the saved size.');
 const changes=[];const upgrade=settingsHeaderUpgrade(data);if(upgrade)changes.push(upgrade);
 changes.push({range:`'${tabs.blockSettings}'!A${data.length+1}:D${data.length+1}`,values:[settingsRow(settingsHeaders(data),areaId,rows,columns,now())]});
 await apiWrites(changes);
 const verified=await readApiTab(tabs.blockSettings,'A:D');verifyBlockSettingsHeader(verified);
 const next=parseBlockSettings(tableFromApi(verified));if(!next.some(s=>norm(s.areaId)===norm(areaId)&&s.rows===rows&&s.columns===columns))throw Error('Area was created, but the chosen grid size could not be verified. Open the area to review it.');
 state.blockSettings=next;
}
async function submitArea(e){e.preventDefault();setError('editorError','');const a=collectEditor();
 const newArea=!editingId,initialRows=Number($('editBlockRows').value),initialColumns=Number($('editBlockColumns').value);
 if(newArea&&state.blocksReady&&canHold(a)&&(!Number.isInteger(initialRows)||initialRows<2||initialRows>30||!Number.isInteger(initialColumns)||initialColumns<1||initialColumns>52)){setError('editorError','Enter 2–30 block rows and 1–52 block columns.');return;}if(!/^[A-Z0-9][A-Z0-9_-]{0,31}$/.test(a.id)){setError('editorError','Use letters, numbers, hyphens or underscores for the area ID.');return;}if(!a.name){setError('editorError','Enter an area name.');return;}if(!canHold(a)&&stockAt(a.id).length){setError('editorError','Move or empty the stock before changing this into a non-storage marking.');return;}if(!editingId&&state.areas.some(old=>norm(old.id)===norm(a.id))){setError('editorError','This ID has already been used. Choose another.');return;}if(editingId&&norm(editingId)!==norm(a.id)){setError('editorError','An area ID cannot be changed after creation.');return;}const error=validateRect(a,editingId);if(error){setError('editorError',error);return;}$('saveArea').disabled=true;const wasEditing=Boolean(editingId);try{await saveAreaRemote(a);if(newArea&&state.blocksReady&&canHold(a))await setInitialBlockGrid(a.id,initialRows,initialColumns);closeEditor();render();toast(wasEditing?'Floor plan updated.':canHold(a)?`Area created with ${initialRows} rows × ${initialColumns} columns.`:'Floor marking added.');}catch(err){setError('editorError',err.message);}finally{$('saveArea').disabled=false;}}
async function archiveArea(){const a=getArea(editingId);if(!a||!isOwner()||busy)return;if(stockAt(a.id).length){setError('editorError','Empty or move the stock before deleting this area.');return;}if(!confirm(`Delete area ${a.id} from the warehouse layout? Its ID will not be reused.`))return;busy=true;try{const rows=await readApiTab(tabs.areas,'A:K');verifyApiHeader(rows,AREA_HEADERS);const index=rows.findIndex((r,i)=>i>0&&norm(r[0])===norm(a.id));if(index<0)throw new Error('Area not found in Google Sheets.');await apiWrites([{range:`'${tabs.areas}'!K${index+1}`,values:[['FALSE']]}]);const checked=await readApiTab(tabs.areas,'A:K');const after=parseAreas(tableFromApi(checked)).find(x=>norm(x.id)===norm(a.id));if(!after||after.active)throw new Error('Area deletion was not confirmed.');state.areas=state.areas.map(v=>norm(v.id)===norm(a.id)?{...v,active:false}:v);closeEditor();render();toast('Empty area deleted.');}catch(err){setError('editorError',err.message);}finally{busy=false;}}
async function setupWarehouse(){if(!isOwner()||busy)return;const ok=confirm('Activate the warehouse layout?\n\nYour existing positions and stock will be copied into new Warehouse Areas and Warehouse Stock tabs. The original tabs will be kept unchanged.');if(!ok)return;busy=true;$('setupButton').disabled=true;connection('Preparing warehouse…');try{
 const meta=await api('?fields=sheets(properties(sheetId,title))');const found=new Set((meta.sheets||[]).map(v=>v.properties.title));const missing=[tabs.areas,tabs.stock].filter(x=>!found.has(x));if(missing.length)await api(':batchUpdate',{method:'POST',body:JSON.stringify({requests:missing.map(title=>({addSheet:{properties:{title,gridProperties:{rowCount:500,columnCount:12}}}}))})});
 let ar=await readApiTab(tabs.areas,'A:K'),st=await readApiTab(tabs.stock,'A:D');if(ar.length)verifyApiHeader(ar,AREA_HEADERS);if(st.length)verifyApiHeader(st,STOCK_HEADERS);
 if(ar.length>1||st.length>1){if(!ar.length||!st.length){const headers=[];if(!ar.length)headers.push({range:`'${tabs.areas}'!A1:K1`,values:[AREA_HEADERS]});if(!st.length)headers.push({range:`'${tabs.stock}'!A1:D1`,values:[STOCK_HEADERS]});if(headers.length)await apiWrites(headers);}throw new Error('Existing warehouse data was found. No old stock was copied to avoid duplicating records. Refresh to inspect the existing layout.');}
 // Read the original stock through the authenticated API so setup never relies on an old public snapshot.
 const original=await readApiTab(tabs.legacy,'A:F'),settings=await readApiTab(tabs.settings,'A:B');const legacy=parseLegacy(tableFromApi(original),tableFromApi(settings));
 const stamp=now(),areas=legacy.areas.map(a=>[a.id,a.name,a.zone,a.type,a.x,a.y,a.w,a.h,a.color,stamp,'TRUE']),stock=legacy.stock.map(s=>[s.areaId,s.item,s.quantity,stamp]);
 const ranges=[{range:`'${tabs.areas}'!A1:K${areas.length+1}`,values:[AREA_HEADERS,...areas]},{range:`'${tabs.stock}'!A1:D${Math.max(1,stock.length+1)}`,values:[STOCK_HEADERS,...stock]}];await apiWrites(ranges);
 ar=await readApiTab(tabs.areas,'A:K');st=await readApiTab(tabs.stock,'A:D');verifyApiHeader(ar,AREA_HEADERS);verifyApiHeader(st,STOCK_HEADERS);const savedAreas=parseAreas(tableFromApi(ar)),savedStock=parseStock(tableFromApi(st));if(savedAreas.length!==areas.length||!equalStock(savedStock,legacy.stock))throw new Error('The imported inventory could not be fully verified. Check the new tabs before proceeding.');
 state.areas=savedAreas;state.stock=savedStock;state.legacy=false;state.loaded=true;designing=false;notice();render();connection('Manager connected','live');toast('Warehouse layout activated. Original stock preserved.');
 }catch(err){notice(err.message||'Could not set up the warehouse.','error');connection('Setup needs attention','error');}finally{busy=false;$('setupButton').disabled=false;}}
// Bulk selection is local until an owner confirms one verified Sheets batch.
function renderBulk(){
 const permitted=isOwner()&&!state.legacy&&state.blocksReady&&canHold(getArea(selectedId));
 if(!permitted){bulkMode=false;bulkSelected.clear();}
 $('bulkToolbar').classList.toggle('hidden',!permitted);
 $('bulkStart').classList.toggle('hidden',bulkMode);$('bulkSelectAll').classList.toggle('hidden',!bulkMode);$('bulkCancel').classList.toggle('hidden',!bulkMode);
 $('bulkHint').textContent=bulkMode?'Tap to select; drag any selected block onto a destination to move all selected stock. Or remove them together below.':'Select blocks for bulk removal, or drag one block onto another to move its contents.';
 const valid=[...bulkSelected].filter(id=>blockStockAt(selectedId,id).length>0&&validBlock(selectedId,id));
 bulkSelected=new Set(valid);
 $('bulkSummary').classList.toggle('hidden',!permitted||!bulkMode);
 $('bulkSelectedCount').textContent=`${valid.length} ${valid.length===1?'block':'blocks'} selected`;
 const totals=new Map();for(const id of valid)for(const s of blockStockAt(selectedId,id)){
  const key=norm(s.item),entry=totals.get(key)||{item:s.item,quantity:0};entry.quantity+=s.quantity;totals.set(key,entry);
 }
 const box=$('bulkTotals');box.replaceChildren();
 if(!valid.length)box.append(make('p','bulk-empty','Select a stocked block above to see removal totals.'));
 for(const item of [...totals.values()].sort((a,b)=>a.item.localeCompare(b.item))){const row=make('div','bulk-total-row');row.append(make('span','',item.item),make('strong','',fmt(item.quantity)));box.append(row);}
 $('bulkRemove').disabled=!valid.length||busy;
}
function refreshBulkBlockState(){for(const button of $('innerBlocks').querySelectorAll('.inner-block')){const chosen=bulkSelected.has(button.dataset.block);button.classList.toggle('bulk-picked',chosen);button.setAttribute('aria-pressed',String(chosen));}renderBulk();}
function toggleBulkBlock(id){if(!bulkMode||!isOwner()||busy)return;if(!blockStockAt(selectedId,id).length)return;
 if(bulkSelected.has(id))bulkSelected.delete(id);else bulkSelected.add(id);setError('bulkError','');refreshBulkBlockState();}
function startBulk(){if(!isOwner()||!state.blocksReady||busy)return;bulkMode=true;blockSelection.clear();focusedBlock='';blockAnchor='';bulkSelected.clear();setError('bulkError','');renderBlockGrid(selectedId);}
function cancelBulk(){bulkMode=false;bulkSelected.clear();setError('bulkError','');renderBlockGrid(selectedId);}
function selectAllBulk(){if(!bulkMode)return;const ids=blockIds(blockRows(selectedId),blockColumns(selectedId)).filter(id=>blockStockAt(selectedId,id).length);bulkSelected=new Set(ids);refreshBulkBlockState();}
async function removeBulk(){
 if(!isOwner()||!bulkMode||busy)return;
 const area=selectedId,blocks=[...bulkSelected].filter(id=>blockStockAt(area,id).length&&validBlock(area,id));
 if(!blocks.length){setError('bulkError','Select at least one stocked block.');return;}
 const totals=new Map();for(const id of blocks)for(const s of blockStockAt(area,id)){
  const key=norm(s.item),e=totals.get(key)||{item:s.item,quantity:0};e.quantity+=s.quantity;totals.set(key,e);
 }
 const summary=[...totals.values()].map(s=>`• ${s.item}: ${fmt(s.quantity)}`).join('\n'),reason=$('bulkReason').value.trim();
 if(!confirm(`Empty ${blocks.length} selected blocks in ${area}?\n\n${summary}\n\nThe full contents of each selected block will be removed from inventory and logged.`))return;
 $('bulkRemove').disabled=true;setError('bulkError','');
 try{
  await transactBlockStock(blocks.map(id=>[area,id]),current=>{
   for(const id of blocks)current.get(blockKey(area,id)).clear();
  },(before,stamp)=>blocks.flatMap(id=>[...before.get(blockKey(area,id)).values()].map(s=>({timestamp:stamp,action:'Bulk remove',item:s.item,quantity:s.quantity,from:`${area}/${id}`,to:'',note:`Bulk removal of ${blocks.length} blocks${reason?' · Reason: '+reason:''}`,reason}))));
  bulkMode=false;bulkSelected.clear();$('bulkReason').value='';renderDetails(false);
 }catch(err){setError('bulkError',err.message||'Bulk removal failed. Refresh and check stock.');renderBulk();}
 finally{$('bulkRemove').disabled=false;}
}
// Each area keeps its existing real block IDs. The view only changes their arrangement.
function blockGroupSize(){return window.matchMedia('(max-width:650px)').matches?2:(document.documentElement.classList.contains('area-route')&&window.innerWidth>=1100?5:4);}
function updateBlockViewControls(){
 for(const [id,value] of [['blockViewLanes','lanes'],['blockViewCompact','compact']]){
  $(id).classList.toggle('is-active',blockView===value);$(id).setAttribute('aria-pressed',String(blockView===value));
 }
 $('blockViewDesc').textContent=blockView==='lanes'?'Physical columns grouped to fit your screen — no sideways scrolling.':'All blocks in one responsive grid — quickest way to scan or edit.';
}
function switchBlockView(next){if(next===blockView)return;blockView=next;if(selectedId)renderBlockGrid(selectedId);}
function refreshBlockSelection(){
 const enabled=isOwner()&&!state.legacy&&state.blocksReady&&canHold(getArea(selectedId));
 for(const button of $('innerBlocks').querySelectorAll('.inner-block')){
  const id=button.dataset.block,chosen=enabled&&blockSelection.has(id);
  button.classList.toggle('sheet-selected',chosen);
  button.classList.toggle('sheet-focused',chosen&&focusedBlock===id);
  if(!bulkMode)button.setAttribute('aria-pressed',String(chosen));
 }
 $('blockSheetToolbar').classList.toggle('hidden',!enabled||bulkMode);
 const size=blockSelection.size,one=size===1,validFocus=Boolean(focusedBlock&&blockSelection.has(focusedBlock));
 $('blockSelectionLabel').textContent=!size?'No blocks selected':size===1?`Block ${focusedBlock||[...blockSelection][0]}`:`${size} blocks selected · ${focusedBlock} active`;
 $('blockClipboardLabel').textContent=blockClipboard?`${blockClipboard.blocks.length} block${blockClipboard.blocks.length===1?'':'s'} copied from ${blockClipboard.area}`:'Select a destination to paste';
 $('blockCopySelection').disabled=!enabled||!size||busy;
 $('blockPasteSelection').disabled=!enabled||!blockClipboard||!validFocus||busy;
 $('blockEditSelection').disabled=!enabled||!one||busy;
 $('blockDeselect').disabled=!size;
}
function chooseBlock(id,shift=false,additive=false){
 if(!validBlock(selectedId,id))return;
 if(shift&&(blockAnchor||focusedBlock)){
  const a=blockCoordinates(blockAnchor||focusedBlock),b=blockCoordinates(id),ids=[];
  for(let r=Math.min(a.row,b.row);r<=Math.max(a.row,b.row);r++)for(let c=Math.min(a.col,b.col);c<=Math.max(a.col,b.col);c++)ids.push(blockColumnLabel(c)+String(r).padStart(2,'0'));
  blockSelection=new Set(additive?[...blockSelection,...ids]:ids);focusedBlock=id;
 }else if(additive){if(blockSelection.has(id))blockSelection.delete(id);else blockSelection.add(id);focusedBlock=id;blockAnchor=id;}
 else{blockSelection=new Set([id]);focusedBlock=id;blockAnchor=id;}
 if(!blockSelection.has(focusedBlock))focusedBlock=[...blockSelection].at(-1)||'';if(!blockSelection.size)blockAnchor='';
 refreshBlockSelection();
}
function focusBlockNode(id){
 const node=[...$('innerBlocks').querySelectorAll('.inner-block')].find(x=>x.dataset.block===id);if(!node)return;
 node.scrollIntoView({block:'nearest',inline:'nearest',behavior:'smooth'});node.focus({preventScroll:true});
}
function renderBlockGrid(areaId){
 const area=getArea(areaId),grid=$('innerBlocks');grid.replaceChildren();if(!area)return;
 const rows=blockRows(areaId),columns=blockColumns(areaId),compact=blockView==='compact';lastBlockGroupSize=blockGroupSize();
 const byId=new Map();for(const stock of state.blockStock){if(norm(stock.areaId)!==norm(areaId)||stock.quantity<=0)continue;const list=byId.get(stock.blockId)||[];list.push(stock);byId.set(stock.blockId,list);}
 grid.classList.toggle('compact-block-grid',compact);grid.classList.toggle('grouped-block-grid',!compact);
 grid.style.removeProperty('grid-template-columns');grid.style.removeProperty('width');grid.style.removeProperty('min-width');
 $('blockMapCount').textContent=`${rows} rows × ${columns} columns · ${rows*columns} blocks`;
 $('blockAccessCaption').textContent=`Row ${rows} is the access end. Select a block to edit or open it to view its contents.`;
 $('rowSettings').classList.toggle('hidden',!isOwner()||state.legacy||!canHold(area));
 $('blockRowCount').value=rows;$('blockColumnCount').value=columns;const hasStock=stockAt(areaId).length>0;$('clearWholeArea').classList.toggle('hidden',!isOwner()||state.legacy||!hasStock);$('areaClearReasonField').classList.toggle('hidden',!isOwner()||state.legacy||!hasStock);$('clearWholeArea').disabled=!hasStock;
 $('blocksDisabled').classList.toggle('hidden',canHold(area));$('blockWorkspace').classList.toggle('hidden',!canHold(area));
 if(!canHold(area)){renderBulk();return;}
 updateBlockViewControls();
 const createBlock=(id)=>{
  const contents=byId.get(id)||[],lane=blockCoordinates(id).col;
  const el=make('button','inner-block '+(contents.length?'filled ':'empty-block ')+(lane%2===0?'lane-b':'lane-a')+(bulkMode&&bulkSelected.has(id)?' bulk-picked':''));
  el.type='button';el.dataset.block=id;el.setAttribute('aria-label',`${areaId} block ${id}: ${contents.length?contents.map(s=>`${s.item}, ${fmt(s.quantity)}`).join('; '):'empty'}`);
  el.title=contents.length?contents.map(s=>`${s.item}: ${fmt(s.quantity)}`).join('\n'):'Empty block';
  el.append(make('strong','inner-id',id),make('span','inner-note',contents.length?contents.map(s=>`${s.item}: ${fmt(s.quantity)}`).join(' · '):'Empty'));
  el.setAttribute('aria-pressed',String(bulkMode&&bulkSelected.has(id)));if(bulkMode&&!contents.length)el.disabled=true;
  el.onclick=e=>{if(Date.now()<ignoreClickUntil)return;if(bulkMode)toggleBulkBlock(id);else if(isOwner()&&state.blocksReady)chooseBlock(id,e.shiftKey,e.ctrlKey||e.metaKey);else openBlock(id);};
  el.ondblclick=()=>{if(!bulkMode&&isOwner()&&state.blocksReady){blockSelection=new Set([id]);focusedBlock=id;blockAnchor=id;refreshBlockSelection();openBlock(id);}};
  if(isOwner()&&state.blocksReady&&contents.length)el.addEventListener('pointerdown',e=>beginBlockMove(e,areaId,id,el));
  return el;
 };
 if(compact){for(const id of blockIds(rows,columns))grid.append(createBlock(id));}
 else{
  const groupSize=lastBlockGroupSize;
  for(let start=1;start<=columns;start+=groupSize){
   const count=Math.min(groupSize,columns-start+1),end=start+count-1;
   const section=make('section','block-lane-section'),heading=make('div','block-lane-heading');
   heading.append(make('strong','',start===end?`COLUMN ${blockColumnLabel(start)}`:`COLUMNS ${blockColumnLabel(start)}–${blockColumnLabel(end)}`),make('span','',`${count} lane${count===1?'':'s'} · ${rows*count} blocks`));
   const inner=make('div','block-lane-grid');inner.style.setProperty('--lane-columns',count);
   for(let r=1;r<=rows;r++)for(let c=start;c<=end;c++)inner.append(createBlock(blockColumnLabel(c)+String(r).padStart(2,'0')));
   section.append(heading,inner);grid.append(section);
  }
 }
 const loose=looseStockAt(areaId),legacy=$('unassignedPanel');$('assignLoosePanel').classList.toggle('hidden',!isOwner()||!state.blocksReady||!loose.length);const dest=$('assignLooseTo'),prior=dest.value;dest.replaceChildren();dest.append(new Option('Choose target block',''));for(const id of blockIds(rows,columns))dest.append(new Option(`Block ${id}`,id));if([...dest.options].some(o=>o.value===prior))dest.value=prior;legacy.classList.toggle('hidden',!loose.length);
 renderBulk();refreshBlockSelection();
 const wrap=$('unassignedContents');wrap.replaceChildren();for(const s of loose){const item=make('div','content-row');item.append(make('span','content-name',s.item),make('span','content-qty',fmt(s.quantity)));wrap.append(item);}
}

const BLOCK_COPY_KIND='AssetsAreaBlockCopyV1';
function copyBlockSelection(clipboardEvent){
 if(!isOwner()||!state.blocksReady||!selectedId||!blockSelection.size)return;
 const ids=[...blockSelection].filter(id=>validBlock(selectedId,id));if(!ids.length)return;
 const coords=ids.map(blockCoordinates),row=Math.min(...coords.map(v=>v.row)),col=Math.min(...coords.map(v=>v.col));
 blockClipboard={kind:BLOCK_COPY_KIND,area:selectedId,originRow:row,originCol:col,blocks:ids.map((id,i)=>({row:coords[i].row-row,col:coords[i].col-col,from:id,items:blockStockAt(selectedId,id).map(s=>({item:s.item,quantity:s.quantity}))}))};
 const text=JSON.stringify(blockClipboard);
 if(clipboardEvent?.clipboardData){clipboardEvent.clipboardData.setData('text/plain',text);clipboardEvent.preventDefault();}
 else if(navigator.clipboard?.writeText)void navigator.clipboard.writeText(text).catch(()=>{});
 refreshBlockSelection();toast(`Copied ${ids.length} block${ids.length===1?'':'s'}. Select a destination, then paste.`);
}
function isValidClipboard(data){
 return Boolean(data&&data.kind===BLOCK_COPY_KIND&&typeof data.area==='string'&&data.area.length<=40&&Array.isArray(data.blocks)&&data.blocks.length>0&&data.blocks.length<=1560&&data.blocks.every(b=>Number.isInteger(b.row)&&b.row>=0&&b.row<30&&Number.isInteger(b.col)&&b.col>=0&&b.col<52&&typeof b.from==='string'&&Array.isArray(b.items)&&b.items.length<=200&&b.items.every(s=>typeof s.item==='string'&&s.item.length>0&&s.item.length<=90&&Number.isSafeInteger(s.quantity)&&s.quantity>0&&s.quantity<=99999999)));
}
function pasteBlockSelection(raw=''){
 if(!isOwner()||busy||!selectedId||!focusedBlock||!validBlock(selectedId,focusedBlock)||bulkMode)return;
 let clip=blockClipboard;
 if(raw){try{const parsed=JSON.parse(raw);if(isValidClipboard(parsed)){clip=parsed;blockClipboard=parsed;}}catch{}}
 if(!isValidClipboard(clip)){toast('Copy blocks first (Ctrl/Cmd+C), then select where to paste.');return;}
 const area=selectedId,origin=blockCoordinates(focusedBlock),moves=clip.blocks.map(b=>({id:blockColumnLabel(origin.col+b.col)+String(origin.row+b.row).padStart(2,'0'),from:b.from,items:b.items.map(i=>({...i}))}));
 if(moves.some(v=>!validBlock(area,v.id))){toast('The copied range does not fit. Select a destination further up or left.');return;}
 const targets=[...new Set(moves.map(m=>m.id))];if(targets.length!==moves.length){toast('Invalid copied range.');return;}
 const altered=moves.filter(m=>!sameBlockStock(m.items,blockStockAt(area,m.id)));
 if(!altered.length){toast('Destination already matches the copied blocks.');return;}
 const occupied=altered.filter(m=>blockStockAt(area,m.id).length);
 const totalAdded=altered.reduce((n,m)=>n+m.items.reduce((v,x)=>v+x.quantity,0),0);
 const desc=`Paste ${clip.blocks.length} copied block${clip.blocks.length===1?'':'s'} from ${clip.area} starting at ${area}/${focusedBlock}?\n\nThis REPLACES the stock in ${altered.length} destination block${altered.length===1?'':'s'} (it does not move or change the sources).${occupied.length?`\n\n${occupied.length} destination block${occupied.length===1?' already contains':'s already contain'} stock, which will be overwritten.`:''}\n\nCopied quantity: ${fmt(totalAdded)}. Every change will be logged.`;
 if(!confirm(desc))return;
 const sourceArea=clip.area;
 void (async()=>{try{
  await transactBlockStock(altered.map(m=>[area,m.id]),current=>{
   for(const m of altered){const target=current.get(blockKey(area,m.id));target.clear();for(const item of m.items)target.set(norm(item.item),{areaId:area,blockId:m.id,item:item.item,quantity:item.quantity});}
  },(before,stamp)=>altered.flatMap(m=>{
   const old=before.get(blockKey(area,m.id)),newRows=new Map(m.items.map(i=>[norm(i.item),i]));const changes=[];
   for(const k of new Set([...old.keys(),...newRows.keys()])){const prior=old.get(k),next=newRows.get(k),difference=(next?.quantity||0)-(prior?.quantity||0);if(!difference)continue;
    changes.push({timestamp:stamp,action:difference>0?'Add (paste)':'Remove (paste)',item:next?.item||prior.item,quantity:Math.abs(difference),from:difference<0?`${area}/${m.id}`:'',to:difference>0?`${area}/${m.id}`:'',note:`Copied from ${sourceArea}/${m.from} to ${area}/${m.id} · Replaced destination contents`,reason:'Block copy / paste'});
   }return changes;
  }));
  blockSelection=new Set(moves.map(m=>m.id));focusedBlock=moves[0].id;blockAnchor=focusedBlock;refreshBlockSelection();toast(`Pasted ${moves.length} block${moves.length===1?'':'s'}. Source stock unchanged.`);
 }catch(err){toast(err.message||'Paste failed. Refresh and check stock.');}})();
}
async function saveBlockQuantities(e){
 e.preventDefault();if(!isOwner()||busy||!selectedBlock)return;
 const area=selectedId,id=selectedBlock,rows=[...$('blockQuantityRows').querySelectorAll('input[data-item]')],updated=new Map();
 for(const input of rows){const value=Number(input.value),item=input.dataset.item;if(!Number.isSafeInteger(value)||value<0||value>99999999){setError('blockQtyEditError',`Enter a valid quantity for ${item}.`);input.focus();return;}updated.set(norm(item),{item,quantity:value});}
 const before=blockStockAt(area,id),changes=before.map(item=>({...item,after:updated.get(norm(item.item))?.quantity??0})).filter(v=>v.after!==v.quantity);
 if(!changes.length){toast('No quantities changed.');return;}
 const reason=$('blockReason').value.trim();setError('blockQtyEditError','');$('saveBlockQuantities').disabled=true;
 try{
  await transactBlockStock([[area,id]],current=>{
   const values=current.get(blockKey(area,id));for(const [key,updatedItem] of updated){const existing=values.get(key);if(existing)values.set(key,{...existing,quantity:updatedItem.quantity});}
  },(_previous,stamp)=>changes.map(item=>({timestamp:stamp,action:item.after>item.quantity?'Add (edit)':'Remove (edit)',item:item.item,quantity:Math.abs(item.after-item.quantity),from:item.after<item.quantity?`${area}/${id}`:'',to:item.after>item.quantity?`${area}/${id}`:'',note:`Direct quantity edit${reason?' · '+reason:''}`,reason})));
  $('blockReason').value='';renderBlockDetails();
 }catch(err){setError('blockQtyEditError',err.message||'Could not save quantity changes.');}finally{$('saveBlockQuantities').disabled=false;}
}
function blockKeyboard(e){
 if(!isOwner()||!state.blocksReady||!$('detailsOverlay').classList.contains('open')||$('blockOverlay').classList.contains('open')||$('mapOverlay').classList.contains('open')||$('editorOverlay').classList.contains('open')||bulkMode)return;
 if(e.target.closest('input,textarea,select,[contenteditable="true"]'))return;
 if(!focusedBlock)return;
 if(e.key==='Enter'&&blockSelection.size===1){e.preventDefault();openBlock(focusedBlock);return;}
 const movement={ArrowUp:[0,-1],ArrowDown:[0,1],ArrowLeft:[-1,0],ArrowRight:[1,0]}[e.key];if(!movement)return;
 const current=blockCoordinates(focusedBlock),col=current.col+movement[0],row=current.row+movement[1];if(col<1||col>blockColumns(selectedId)||row<1||row>blockRows(selectedId))return;
 e.preventDefault();const id=blockColumnLabel(col)+String(row).padStart(2,'0');chooseBlock(id,e.shiftKey,false);focusBlockNode(id);
}
function clearBlockMove(){
 clearTimeout(blockMoveTimer);blockMoveTimer=0;
 document.querySelectorAll('.inner-block.move-source,.inner-block.move-target').forEach(el=>el.classList.remove('move-source','move-target'));
 document.querySelector('.block-drag-label')?.remove();blockMoveDrag=null;
}
function beginBlockMove(e,area,source,el){
 if(!isOwner()||busy||e.button!==0||!blockStockAt(area,source).length)return;
 clearBlockMove();
 blockMoveDrag={pointer:e.pointerId,area,source,x:e.clientX,y:e.clientY,active:false,armed:e.pointerType!=='touch',el};
 if(e.pointerType==='touch')blockMoveTimer=setTimeout(()=>{if(blockMoveDrag&&blockMoveDrag.pointer===e.pointerId){blockMoveDrag.armed=true;el.classList.add('move-source');if(navigator.vibrate)navigator.vibrate(25);}},350);
}
function moveBlockPointer(e){
 const d=blockMoveDrag;if(!d||d.pointer!==e.pointerId)return;
 const distance=Math.hypot(e.clientX-d.x,e.clientY-d.y);
 if(!d.armed){if(distance>10)clearBlockMove();return;}
 if(!d.active&&distance<9)return;
 d.active=true;e.preventDefault();
 const selected=bulkMode&&bulkSelected.has(d.source)?[...bulkSelected].filter(id=>blockStockAt(d.area,id).length):[d.source];
 d.sources=selected;d.el.classList.add('move-source');
 let ghost=document.querySelector('.block-drag-label');if(!ghost){ghost=make('div','block-drag-label',`${selected.length} block${selected.length===1?'':'s'} → drop on destination`);document.body.append(ghost);}
 ghost.style.left=`${Math.min(window.innerWidth-205,e.clientX+12)}px`;ghost.style.top=`${Math.max(4,e.clientY-45)}px`;
 const target=document.elementFromPoint(e.clientX,e.clientY)?.closest('.inner-block');
 document.querySelectorAll('.inner-block.move-target').forEach(node=>node.classList.remove('move-target'));
 d.destination=target&&$('innerBlocks').contains(target)&&!selected.includes(target.dataset.block)?target.dataset.block:null;
 if(d.destination)target.classList.add('move-target');
}
function finishBlockPointer(e){
 const d=blockMoveDrag;if(!d||d.pointer!==e.pointerId)return;
 const wasActive=d.active,sources=d.sources||[d.source],destination=d.destination,area=d.area;
 clearBlockMove();
 if(!wasActive)return;
 ignoreClickUntil=Date.now()+700;e.preventDefault();
 if(!destination){toast('Drop onto a different block in this area.');return;}
 void movePickedBlocks(area,sources,destination);
}
async function movePickedBlocks(area,sources,destination){
 if(!isOwner()||busy||!validBlock(area,destination))return;
 const from=[...new Set(sources)].filter(id=>id!==destination&&validBlock(area,id)&&blockStockAt(area,id).length);
 if(!from.length)return;
 const amounts=new Map();for(const id of from)for(const asset of blockStockAt(area,id)){const entry=amounts.get(norm(asset.item))||{item:asset.item,quantity:0};entry.quantity+=asset.quantity;amounts.set(norm(asset.item),entry);}
 const summary=[...amounts.values()].map(x=>`${fmt(x.quantity)} ${x.item}`).join(', ');
 if(!confirm(`Move ALL stock from ${from.map(id=>area+'/'+id).join(', ')} to ${area}/${destination}?\n\n${summary}\n\nMatching types in the destination will be merged. This will be logged as a block transaction.`))return;
 try{
  await transactBlockStock([...from.map(id=>[area,id]),[area,destination]],current=>{
   const target=current.get(blockKey(area,destination));
   for(const id of from){const source=current.get(blockKey(area,id));for(const [key,item] of source){const old=target.get(key);target.set(key,{areaId:area,blockId:destination,item:old?.item||item.item,quantity:(old?.quantity||0)+item.quantity});}source.clear();}
  },(before,stamp)=>from.flatMap(id=>[...before.get(blockKey(area,id)).values()].map(item=>({timestamp:stamp,action:from.length>1?'Bulk move':'Move',item:item.item,quantity:item.quantity,from:`${area}/${id}`,to:`${area}/${destination}`,note:`Drag-and-drop transfer from ${from.length} block(s)`,reason:''}))));
  bulkMode=false;bulkSelected.clear();renderDetails(false);
 }catch(err){toast(err.message||'Could not move selected blocks.');}
}
function findBlock(){const id=String($('blockFind').value||'').trim().toUpperCase(),status=$('blockFindStatus');if(!validBlock(selectedId,id)){status.textContent='That block is not in this area.';return;}
 const node=Array.from($('innerBlocks').querySelectorAll('.inner-block')).find(el=>el.dataset.block===id);if(!node)return;node.scrollIntoView({behavior:'smooth',block:'center',inline:'nearest'});node.classList.add('block-found');node.focus({preventScroll:true});if(isOwner()&&!bulkMode){chooseBlock(id);}status.textContent=`Showing ${id}`;setTimeout(()=>node.classList.remove('block-found'),2400);}
function renderFocusMap(){const board=$('focusMapBoard'),areas=activeAreas(),rows=Math.max(MIN_R,...areas.map(a=>a.y+a.h+1));board.style.setProperty('--cols',C);board.style.setProperty('--rows',Math.min(MAX_R,rows));board.replaceChildren();applyMapSize('focusMapBoard');const here=getArea(selectedId);$('mapHereCaption').textContent=here?`${here.name} (${here.id})`:'Choose an area';
 for(const a of areas){const selected=here&&norm(here.id)===norm(a.id),inventory=stockAt(a.id),special=!canHold(a),el=make('button','map-block color-'+a.color+(inventory.length?' occupied':'')+(special?' special':'')+(selected?' map-current':''));el.type='button';el.dataset.id=a.id;el.style.gridColumn=`${a.x} / span ${a.w}`;el.style.gridRow=`${a.y} / span ${a.h}`;el.setAttribute('aria-label',`${a.type}: ${a.name}, ${a.id}${selected?', your current area':''}`);el.title=`${a.type} · ${a.name}`;if(special)el.append(make('span','marker-symbol',markerSymbols[norm(a.type)]||'✦'));el.append(make('span','block-id',a.id),make('span','block-name',a.name));if(selected){const marker=make('span','you-are-here');marker.append(make('span','here-dot'),document.createTextNode('YOU ARE HERE'));el.append(marker);}else el.append(make('span','block-status',special?a.type:inventory.length?`${inventory.length} types stored`:'Empty'));el.onclick=()=>{closeOverlay('mapOverlay');if(selected)return;closeOverlay('blockOverlay');openDetails(a.id);};board.append(el);}
}
function showAreaMap(){if(!selectedId||!getArea(selectedId))return;renderFocusMap();openOverlay('mapOverlay');requestAnimationFrame(()=>{applyMapSize('focusMapBoard');const selected=$('focusMapBoard').querySelector('.map-current'),scroll=$('mapCanvasScroll');if(selected){scroll.scrollLeft=Math.max(0,selected.offsetLeft-scroll.offsetLeft-scroll.clientWidth/2+selected.clientWidth/2);scroll.scrollTop=Math.max(0,selected.offsetTop-scroll.offsetTop-scroll.clientHeight/2+selected.clientHeight/2);}});}
function openWarehouseOverview(){closeOverlay('mapOverlay');closeDetails(true);render();}
function blockLink(areaId,blockId){const u=new URL(location.href);u.searchParams.set('area',areaId);u.searchParams.set('block',blockId);u.hash='';return u.toString();}
function openBlock(id,update=true){
 if(!validBlock(selectedId,id))return;selectedBlock=id;
 $('blockTitle').textContent=`Block ${id}`;$('blockSubtitle').textContent=`${getArea(selectedId).name} · ${selectedId}`;
 $('blockForm').reset();$('blockQty').value=1;setError('blockError','');renderBlockDetails();openOverlay('blockOverlay');
 if(update){const u=new URL(location.href);u.searchParams.set('area',selectedId);u.searchParams.set('block',id);history.pushState({},'',u.pathname+u.search+u.hash);}
}
function closeBlock(update=true){if(!selectedBlock)return;closeOverlay('blockOverlay');selectedBlock='';if(update){const u=new URL(location.href);u.searchParams.delete('block');history.pushState({},'',u.pathname+u.search+u.hash);}}
function renderBlockDetails(){
 if(!selectedBlock)return;const items=blockStockAt(selectedId,selectedBlock).sort((a,b)=>a.item.localeCompare(b.item)),wrap=$('blockContents');wrap.replaceChildren();
 for(const s of items){const el=make('div','content-row');el.append(make('span','content-name',s.item),make('span','content-qty',fmt(s.quantity)));wrap.append(el);}
 if(!items.length)wrap.append(make('div','content-empty','This block is empty.'));
 $('blockItemCount').textContent=`${items.length} ${items.length===1?'asset type':'asset types'}`;
 const enabled=isOwner()&&!state.legacy&&state.blocksReady&&canHold(getArea(selectedId));
 $('blockOwnerPanel').classList.toggle('hidden',!enabled);
 $('blockSetupMessage').classList.toggle('hidden',state.blocksReady||!isOwner());
 $('blockClear').disabled=!items.length;
 $('blockTransfer').classList.toggle('hidden',!items.length);
 const chosen=$('blockItem').value,sel=$('blockItem');sel.replaceChildren();sel.append(new Option('Choose an asset',''));
 const stockList=Array.from(new Set([...PRESETS,...state.blockStock.map(s=>s.item)]));
 if($('blockAction').value==='remove'){
  for(const s of items)sel.append(new Option(`${s.item} (${fmt(s.quantity)} here)`,s.item));
 }else{for(const name of stockList)sel.append(new Option(name,name));sel.append(new Option('Other / custom','__other__'));}
 if([...sel.options].some(o=>o.value===chosen))sel.value=chosen;
 $('blockCustomWrap').classList.toggle('hidden',sel.value!=='__other__'||$('blockAction').value==='remove');
 $('blockCustom').required=sel.value==='__other__'&&$('blockAction').value==='add';
 const editForm=$('blockQuantitiesForm'),rows=$('blockQuantityRows');rows.replaceChildren();
 editForm.classList.toggle('hidden',!enabled||!items.length);setError('blockQtyEditError','');
 for(const stock of items){
  const label=make('label','block-quantity-row'),name=make('span','',stock.item),input=make('input');
  input.type='number';input.min='0';input.max='99999999';input.step='1';input.inputMode='numeric';input.value=String(stock.quantity);input.dataset.item=stock.item;input.setAttribute('aria-label',`Set ${stock.item} quantity`);
  label.append(name,input);rows.append(label);
 }
 const dest=$('blockTo'),prior=dest.value;dest.replaceChildren();dest.append(new Option('Select destination block',''));
 for(const other of blockIds(blockRows(selectedId),blockColumns(selectedId)))if(other!==selectedBlock)dest.append(new Option(`Block ${other}`,other));
 if([...dest.options].some(o=>o.value===prior))dest.value=prior;
}
async function enableBlockStorage(){
 if(!isOwner()||busy||state.legacy)return;busy=true;$('enableBlocks').disabled=true;connection('Setting up storage blocks…');
 try{
  const meta=await api('?fields=sheets(properties(title))'),known=new Set((meta.sheets||[]).map(x=>x.properties.title));
  const missing=[tabs.blockSettings,tabs.blockStock].filter(x=>!known.has(x));
  if(missing.length)await api(':batchUpdate',{method:'POST',body:JSON.stringify({requests:missing.map(title=>({addSheet:{properties:{title,gridProperties:{rowCount:1000,columnCount:8}}}}))})});
  const settings=await readApiTab(tabs.blockSettings,'A:D'),stock=await readApiTab(tabs.blockStock,'A:E');
  if(settings.length&&settings[0].some(v=>String(v||'').trim()))verifyBlockSettingsHeader(settings);
  if(stock.length&&stock[0].some(v=>String(v||'').trim()))verifyApiHeader(stock,BLOCK_STOCK_HEADERS);
  const changes=[];
  if(!settings.length||!settings[0].some(v=>String(v||'').trim()))changes.push({range:`'${tabs.blockSettings}'!A1:D1`,values:[BLOCK_SETTINGS_HEADERS]});
  else {const upgrade=settingsHeaderUpgrade(settings);if(upgrade)changes.push(upgrade);}
  if(!stock.length||!stock[0].some(v=>String(v||'').trim()))changes.push({range:`'${tabs.blockStock}'!A1:E1`,values:[BLOCK_STOCK_HEADERS]});
  if(changes.length)await apiWrites(changes);
  const [s,b]=await Promise.all([readApiTab(tabs.blockSettings,'A:D'),readApiTab(tabs.blockStock,'A:E')]);
  verifyBlockSettingsHeader(s);verifyApiHeader(b,BLOCK_STOCK_HEADERS);
  state.blockSettings=parseBlockSettings(tableFromApi(s));state.blockStock=parseBlockStock(tableFromApi(b));state.blocksReady=true;
  render();toast('Block storage enabled. Set rows and columns independently in each area.');connection('Manager connected','live');
 }catch(err){toast(err.message||'Could not enable storage blocks.');connection('Setup failed','error');}
 finally{busy=false;$('enableBlocks').disabled=false;}
}
// Upgrades the previous 3-column settings tab without altering existing row counts or stock.
function verifyBlockSettingsHeader(rows){if(!rows.length)return;settingsHeaders(rows);}
async function saveBlockGrid(){
 if(!isOwner()||busy||!selectedId||!state.blocksReady)return;
 const id=selectedId,rows=Number($('blockRowCount').value),columns=Number($('blockColumnCount').value),oldRows=blockRows(id),oldColumns=blockColumns(id);
 if(!Number.isInteger(rows)||rows<2||rows>30){toast('Use between 2 and 30 rows.');return;}
 if(!Number.isInteger(columns)||columns<1||columns>52){toast('Use between 1 and 52 columns.');return;}
 if(state.blockStock.some(s=>{if(norm(s.areaId)!==norm(id))return false;const pos=blockCoordinates(s.blockId);return !pos||pos.row>rows||pos.col>columns;})){toast('Move or clear stock from the blocks you are removing first.');return;}
 if(rows===oldRows&&columns===oldColumns)return;
 busy=true;$('saveBlockGrid').disabled=true;
 try{
  let data=await readApiTab(tabs.blockSettings,'A:D');verifyBlockSettingsHeader(data);
  const at=data.findIndex((r,i)=>i>0&&norm(r[0])===norm(id));
  const current=parseBlockSettings(tableFromApi(data)).find(x=>norm(x.areaId)===norm(id));
  if((current?.rows||8)!==oldRows||(current?.columns||2)!==oldColumns)throw Error('Grid changed in another session. Refresh first.');
  // Recheck live inventory before shrinking, including stock added from another session.
  const remoteStock=await readApiTab(tabs.blockStock,'A:E');verifyApiHeader(remoteStock,BLOCK_STOCK_HEADERS);
  if(parseBlockStock(tableFromApi(remoteStock)).some(s=>{if(norm(s.areaId)!==norm(id))return false;const pos=blockCoordinates(s.blockId);return !pos||pos.row>rows||pos.col>columns;}))throw Error('A block being removed contains stock. Refresh and relocate its stock first.');
  const writes=[];
  const upgrade=settingsHeaderUpgrade(data);if(upgrade)writes.push(upgrade);
  const n=at>=0?at+1:data.length+1;
  writes.push({range:`'${tabs.blockSettings}'!A${n}:D${n}`,values:[settingsRow(settingsHeaders(data),id,rows,columns,now())]});
  await apiWrites(writes);
  data=await readApiTab(tabs.blockSettings,'A:D');verifyBlockSettingsHeader(data);
  const after=parseBlockSettings(tableFromApi(data));
  if(!after.some(x=>norm(x.areaId)===norm(id)&&x.rows===rows&&x.columns===columns))throw Error('Block grid could not be verified.');
  state.blockSettings=after;render();toast(`Area now has ${rows*columns} blocks (${rows} × ${columns}).`);
 }catch(err){toast(err.message);}finally{busy=false;$('saveBlockGrid').disabled=false;}
}
const blockKey=(a,b)=>norm(a)+'|'+norm(b);
function validateBlockRows(rows){const seen=new Set();for(const s of rows){const key=blockKey(s.areaId,s.blockId)+'|'+norm(s.item);if(seen.has(key))throw Error('Duplicate stock entries found in Warehouse Block Stock.');seen.add(key);}}
function sameBlockStock(a,b){return equalStock(a,b);}
async function transactBlockStock(locations,modify,eventsFn){
 if(!isOwner()||!state.blocksReady)throw Error('Enable block storage first.');if(busy)throw Error('A save is already in progress.');
 busy=true;connection('Saving block stock…');
 try{
  await ensureBlockTransactions();
  const data=await readApiTab(tabs.blockStock,'A:E');verifyApiHeader(data,BLOCK_STOCK_HEADERS);
  const all=parseBlockStock(tableFromApi(data));validateBlockRows(all);
  for(const [area,block] of locations){
   if(!validBlock(area,block))throw Error('Invalid destination block.');
   const persisted=all.filter(s=>blockKey(s.areaId,s.blockId)===blockKey(area,block));
   if(!sameBlockStock(persisted,blockStockAt(area,block)))throw Error('Block stock changed in another session. Refresh first.');
  }
  const current=new Map(locations.map(([a,b])=>[blockKey(a,b),mapStock(all.filter(s=>blockKey(s.areaId,s.blockId)===blockKey(a,b)))]));
  const before=new Map([...current].map(([key,val])=>[key,new Map(val)]));modify(current);
  const changes=[];
  for(const [k,entries] of current){
   for(const item of entries.values())if(!Number.isSafeInteger(item.quantity)||item.quantity<0)throw Error('Quantity must be a positive whole number.');
   for(const [item,old] of before.get(k)){const next=entries.get(item);if(!next||next.quantity!==old.quantity)changes.push({old,next:next?.quantity>0?next:null});}
   for(const [item,next] of entries)if(!before.get(k).has(item)&&next.quantity>0)changes.push({old:null,next});
  }
  if(!changes.length){toast('Nothing to change.');return;}
  const occupied=new Map(),blank=[];
  for(let i=1;i<data.length;i++){const r=data[i]||[],key=blockKey(r[0],r[1])+'|'+norm(r[2]);if(String(r[0]||'').trim()&&String(r[2]||'').trim()&&Number(r[3])>0)occupied.set(key,i+1);else blank.push(i+1);}
  let n=data.length+1;const stamp=now(),writes=[];
  for(const {old,next} of changes){const s=old||next,key=blockKey(s.areaId,s.blockId)+'|'+norm(s.item),row=occupied.get(key)||blank.shift()||n++;writes.push({range:`'${tabs.blockStock}'!A${row}:E${row}`,values:[[next?next.areaId:'',next?next.blockId:'',next?next.item:'',next?next.quantity:0,stamp]]});}
  await apiWrites(writes);
  const checked=await readApiTab(tabs.blockStock,'A:E');verifyApiHeader(checked,BLOCK_STOCK_HEADERS);const nextRows=parseBlockStock(tableFromApi(checked));validateBlockRows(nextRows);
  for(const [a,b] of locations){const key=blockKey(a,b),want=[...current.get(key).values()].filter(s=>s.quantity>0),actual=nextRows.filter(s=>blockKey(s.areaId,s.blockId)===key);if(!sameBlockStock(want,actual))throw Error('Saved values differ from expected. Refresh and verify your stock.');}
  state.blockStock=nextRows;
  const events=eventsFn(before,stamp)||[];
  const statuses=await Promise.allSettled([appendEvents(events),appendBlockTransactions(events)]);
  const logged=statuses.every(v=>v.status==='fulfilled');if(!logged)for(const v of statuses)if(v.status==='rejected')console.warn('Transaction log issue',v.reason);
  render();renderBlockDetails();connection('Manager connected','live');toast(logged?'Block stock saved and logged.':'Stock saved, but a transaction log failed. Review Block Transactions and refresh.');
 }catch(err){connection('Save failed','error');throw err;}finally{busy=false;}
}
async function submitBlockStock(e){
 e.preventDefault();if(!selectedBlock||busy)return;setError('blockError','');
 const area=selectedId,block=selectedBlock,action=$('blockAction').value,qty=Number($('blockQty').value),pick=$('blockItem').value,item=pick==='__other__'&&action==='add'?$('blockCustom').value.trim():pick,reason=$('blockReason').value.trim();
 if(!Number.isSafeInteger(qty)||qty<1){setError('blockError','Enter a valid whole quantity.');return;}
 if(!item||item.length>90){setError('blockError','Choose an asset type.');return;}
 $('blockSave').disabled=true;
 try{
  await transactBlockStock([[area,block]],next=>{const m=next.get(blockKey(area,block)),key=norm(item),old=m.get(key);if(action==='remove'){if(!old||old.quantity<qty)throw Error('Not enough stock in this block.');m.set(key,{...old,quantity:old.quantity-qty});}else m.set(key,{areaId:area,blockId:block,item:old?.item||item,quantity:(old?.quantity||0)+qty});},(_prev,stamp)=>[{timestamp:stamp,action:action==='add'?'Add':'Remove',item,quantity:qty,from:action==='remove'?`${area}/${block}`:'',to:action==='add'?`${area}/${block}`:'',note:`Storage block${action==='remove'&&reason?' · Reason: '+reason:''}`,reason:action==='remove'?reason:''}]);
  $('blockQty').value=1;$('blockReason').value='';renderBlockDetails();
 }catch(err){setError('blockError',err.message);}finally{$('blockSave').disabled=false;}
}
async function clearBlock(){
 const area=selectedId,block=selectedBlock,items=blockStockAt(area,block);if(!isOwner()||busy||!items.length)return;
 const reason=$('blockReason').value.trim();
 if(!confirm(`Empty ${area} / ${block}?\n\n${items.map(s=>`• ${s.item}: ${fmt(s.quantity)}`).join('\n')}\n\nAll quantities in this block will be removed and logged.`))return;
 $('blockClear').disabled=true;
 try{await transactBlockStock([[area,block]],next=>next.get(blockKey(area,block)).clear(),(before,stamp)=>[...before.get(blockKey(area,block)).values()].map(s=>({timestamp:stamp,action:'Clear block',item:s.item,quantity:s.quantity,from:`${area}/${block}`,to:'',note:`Entire block cleared${reason?' · Reason: '+reason:''}`,reason})));$('blockReason').value='';}
 catch(err){toast(err.message);}finally{$('blockClear').disabled=false;}
}
async function transferBlock(){
 const area=selectedId,src=selectedBlock,dst=$('blockTo').value,items=blockStockAt(area,src);if(!dst||dst===src||!items.length||!validBlock(area,dst))return;
 if(!confirm(`Move everything from ${area}/${src} to ${area}/${dst}? Matching assets at the destination will be combined.`))return;
 try{
  await transactBlockStock([[area,src],[area,dst]],current=>{
   const a=current.get(blockKey(area,src)),b=current.get(blockKey(area,dst));for(const [key,s] of a){const prior=b.get(key);b.set(key,{areaId:area,blockId:dst,item:prior?.item||s.item,quantity:(prior?.quantity||0)+s.quantity});}a.clear();
  },(before,stamp)=>[...before.get(blockKey(area,src)).values()].map(s=>({timestamp:stamp,action:'Move',item:s.item,quantity:s.quantity,from:`${area}/${src}`,to:`${area}/${dst}`,note:'Block transfer'})));
  renderBlockDetails();
 }catch(err){toast(err.message);}
}
// Clear all stock in an area, including its nested blocks, with one verified batch of changes.

async function assignLooseToBlock(){
 const area=selectedId,block=$('assignLooseTo').value;
 if(!isOwner()||busy||!state.blocksReady||!validBlock(area,block))return;
 const current=looseStockAt(area);if(!current.length)return;
 if(!confirm(`Assign all unassigned stock in ${area} to block ${block}? Matching asset types will be combined.`))return;
 busy=true;$('assignLooseButton').disabled=true;connection('Assigning stock…');
 try{
  await ensureBlockTransactions();
  const [looseSheet,blockSheet]=await Promise.all([readApiTab(tabs.stock,'A:D'),readApiTab(tabs.blockStock,'A:E')]);
  verifyApiHeader(looseSheet,STOCK_HEADERS);verifyApiHeader(blockSheet,BLOCK_STOCK_HEADERS);
  const source=parseStock(tableFromApi(looseSheet)).filter(s=>norm(s.areaId)===norm(area));
  const target=parseBlockStock(tableFromApi(blockSheet)).filter(s=>blockKey(s.areaId,s.blockId)===blockKey(area,block));
  if(!equalStock(source,current)||!equalStock(target,blockStockAt(area,block)))throw Error('Stock changed. Refresh before assigning.');
  const rows=[],occupied=new Map(),blanks=[];
  for(let i=1;i<blockSheet.length;i++){const r=blockSheet[i],k=blockKey(r?.[0],r?.[1])+'|'+norm(r?.[2]);if(r?.[0]&&r?.[2]&&Number(r?.[3])>0)occupied.set(k,i+1);else blanks.push(i+1);}
  let nextRow=blockSheet.length+1;const stamp=now(),merged=mapStock(target);
  for(const item of source){const old=merged.get(norm(item.item));merged.set(norm(item.item),{areaId:area,blockId:block,item:old?.item||item.item,quantity:(old?.quantity||0)+item.quantity});}
  for(let i=1;i<looseSheet.length;i++)if(norm(looseSheet[i]?.[0])===norm(area)&&Number(looseSheet[i]?.[2])>0)rows.push({range:`'${tabs.stock}'!A${i+1}:D${i+1}`,values:[['','',0,stamp]]});
  for(const item of merged.values()){
   const key=blockKey(area,block)+'|'+norm(item.item),row=occupied.get(key)||blanks.shift()||nextRow++;
   rows.push({range:`'${tabs.blockStock}'!A${row}:E${row}`,values:[[area,block,item.item,item.quantity,stamp]]});
  }
  await apiWrites(rows);
  const [verifiedLoose,verifiedBlocks]=await Promise.all([readApiTab(tabs.stock,'A:D'),readApiTab(tabs.blockStock,'A:E')]);
  const nextLoose=parseStock(tableFromApi(verifiedLoose)),nextBlocks=parseBlockStock(tableFromApi(verifiedBlocks));
  if(nextLoose.some(s=>norm(s.areaId)===norm(area))||!equalStock(nextBlocks.filter(s=>blockKey(s.areaId,s.blockId)===blockKey(area,block)),[...merged.values()]))throw Error('Could not verify allocation. Refresh and check the stock.');
  state.stock=nextLoose;state.blockStock=nextBlocks;
  const events=source.map(item=>({timestamp:stamp,action:'Move',item:item.item,quantity:item.quantity,from:`${area} (unassigned)`,to:`${area}/${block}`,note:'Assigned existing stock to block'}));const logResults=await Promise.allSettled([appendEvents(events),appendBlockTransactions(events)]);if(logResults.some(r=>r.status==='rejected'))toast('Stock assigned; check block transaction logs.');
  render();connection('Manager connected','live');toast(`Stock assigned to ${area}/${block}.`);
 }catch(err){toast(err.message);connection('Save failed','error');}
 finally{busy=false;$('assignLooseButton').disabled=false;}
}

async function clearEntireArea(areaId,reason=''){
 if(busy||!isOwner())throw Error('Manager editing is required.');busy=true;connection('Clearing area…');
 try{
  const [a,b]=await Promise.all([readApiTab(tabs.stock,'A:D'),...(state.blocksReady?[readApiTab(tabs.blockStock,'A:E')]:[])]);
  verifyApiHeader(a,STOCK_HEADERS);if(b)verifyApiHeader(b,BLOCK_STOCK_HEADERS);
  const loose=parseStock(tableFromApi(a)).filter(s=>norm(s.areaId)===norm(areaId)),blocks=b?parseBlockStock(tableFromApi(b)).filter(s=>norm(s.areaId)===norm(areaId)):[];
  if(!equalStock(loose,looseStockAt(areaId)))throw Error('Unassigned stock changed. Refresh first.');
  for(const blockId of new Set(blocks.map(s=>s.blockId)))if(!equalStock(blocks.filter(s=>s.blockId===blockId),blockStockAt(areaId,blockId)))throw Error('Block stock changed. Refresh first.');
  if(blocks.length!==state.blockStock.filter(s=>norm(s.areaId)===norm(areaId)).length)throw Error('Block stock changed. Refresh first.');
  await ensureBlockTransactions();
  const writes=[];
  for(let i=1;i<a.length;i++)if(norm(a[i]?.[0])===norm(areaId)&&Number(a[i]?.[2])>0)writes.push({range:`'${tabs.stock}'!A${i+1}:D${i+1}`,values:[['','',0,now()]]});
  if(b)for(let i=1;i<b.length;i++)if(norm(b[i]?.[0])===norm(areaId)&&Number(b[i]?.[3])>0)writes.push({range:`'${tabs.blockStock}'!A${i+1}:E${i+1}`,values:[['','','',0,now()]]});
  if(!writes.length)return;
  await apiWrites(writes);
  const [checkedLoose,checkedBlocks]=await Promise.all([readApiTab(tabs.stock,'A:D'),...(state.blocksReady?[readApiTab(tabs.blockStock,'A:E')]:[])]);
  const latestLoose=parseStock(tableFromApi(checkedLoose)),latestBlocks=checkedBlocks?parseBlockStock(tableFromApi(checkedBlocks)):[];
  if(latestLoose.some(s=>norm(s.areaId)===norm(areaId))||latestBlocks.some(s=>norm(s.areaId)===norm(areaId)))throw Error('Area clear could not be verified; refresh and review inventory.');
  state.stock=latestLoose;state.blockStock=latestBlocks;
  const stamp=now(),events=[...loose.map(s=>({...s,location:areaId})),...blocks.map(s=>({...s,location:`${areaId}/${s.blockId}`}))].map(s=>({timestamp:stamp,action:'Clear area',item:s.item,quantity:s.quantity,from:s.location,to:'',note:`Whole area emptied${reason?' · Reason: '+reason:''}`,reason}));
  const results=await Promise.allSettled([appendEvents(events),appendBlockTransactions(events.filter(e=>e.from.includes('/')))]);if(results.some(r=>r.status==='rejected'))toast('Stock cleared; verify the transaction history.');
  render();connection('Manager connected','live');toast('Entire area emptied, including all storage blocks.');
 }catch(err){connection('Clear failed','error');throw err;}finally{busy=false;}
}

function signIn(){if(!cfg.OAUTH_CLIENT_ID||!cfg.OWNER_EMAIL){toast('Manager login is not configured.');return;}if(!gisReady||!window.google?.accounts?.oauth2){toast('Google sign-in is loading. Try again.');return;}
 if(!tokenClient)tokenClient=google.accounts.oauth2.initTokenClient({client_id:cfg.OAUTH_CLIENT_ID,scope:'openid email https://www.googleapis.com/auth/spreadsheets',callback:async r=>{if(!r.access_token){toast('Sign-in was cancelled.');return;}const candidate=r.access_token;try{const response=await fetch('https://www.googleapis.com/oauth2/v3/userinfo',{headers:{Authorization:'Bearer '+candidate}});if(!response.ok)throw new Error('Could not verify Google account.');const who=await response.json();if(!who.email_verified||norm(who.email)!==norm(cfg.OWNER_EMAIL))throw new Error('This account does not have manager access.');token=candidate;ownerEmail=who.email;mode='owner';render();toast('Manager editing unlocked.');}catch(err){token='';ownerEmail='';mode='viewer';render();toast(err.message);}}});tokenClient.requestAccessToken({prompt:'select_account'});}
function signOut(){bulkMode=false;bulkSelected.clear();blockSelection.clear();focusedBlock='';blockAnchor='';blockClipboard=null;clearStockDrag();clearBlockMove();token='';ownerEmail='';mode='viewer';designing=false;drawingTool='select';closeEditor();render();connection('Live · view only','live');toast('Editing locked.');}
function initialise(){
 $('warehouseTab').onclick=()=>showMainView('warehouse');$('transactionsTab').onclick=()=>showMainView('transactions');
 $('transactionType').onchange=()=>{transactionLimit=50;renderBlockTransactions();};$('transactionSearch').oninput=()=>{transactionLimit=50;renderBlockTransactions();};$('transactionsMore').onclick=()=>{transactionLimit+=50;renderBlockTransactions();};
 document.addEventListener('pointermove',moveBlockPointer,{passive:false});document.addEventListener('pointerup',finishBlockPointer);document.addEventListener('pointercancel',clearBlockMove);
 $('findForm').onsubmit=findAssets; $('findItem').addEventListener('change',()=>{findRequest=null;$('findResult').classList.add('hidden');});$('findQty').addEventListener('input',()=>{findRequest=null;$('findResult').classList.add('hidden');});
 window.addEventListener('copy',e=>{if(isOwner()&&blockSelection.size&&$('detailsOverlay').classList.contains('open')&&!$('blockOverlay').classList.contains('open')&&!e.target.closest('input,textarea,select,[contenteditable="true"]'))copyBlockSelection(e);});
 window.addEventListener('paste',e=>{if(!isOwner()||!focusedBlock||!$('detailsOverlay').classList.contains('open')||$('blockOverlay').classList.contains('open')||e.target.closest('input,textarea,select,[contenteditable="true"]'))return;const raw=e.clipboardData?.getData('text/plain')||'';if(raw){try{if(isValidClipboard(JSON.parse(raw))){e.preventDefault();pasteBlockSelection(raw);}}catch{}}});
 $('blockViewLanes').onclick=()=>switchBlockView('lanes');$('blockViewCompact').onclick=()=>switchBlockView('compact');$('blockCopySelection').onclick=()=>copyBlockSelection();$('blockPasteSelection').onclick=()=>pasteBlockSelection();$('blockEditSelection').onclick=()=>{if(blockSelection.size===1&&focusedBlock)openBlock(focusedBlock);};$('blockDeselect').onclick=()=>{blockSelection.clear();focusedBlock='';blockAnchor='';refreshBlockSelection();};$('blockQuantitiesForm').onsubmit=saveBlockQuantities;window.addEventListener('keydown',blockKeyboard);
 $('bulkStart').onclick=startBulk;$('bulkCancel').onclick=cancelBulk;$('bulkSelectAll').onclick=selectAllBulk;$('bulkRemove').onclick=removeBulk;
 $('areaFindAction').onclick=()=>{openWarehouseOverview();requestAnimationFrame(()=>$('findPanel').scrollIntoView({behavior:'smooth',block:'start'}));};
 $('enableBlocks').onclick=enableBlockStorage;$('saveBlockGrid').onclick=saveBlockGrid;$('blockFindGo').onclick=findBlock;$('blockFind').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();findBlock();}});$('areaOwnerButton').onclick=()=>isOwner()?signOut():signIn();$('openAreaMap').onclick=showAreaMap;$('closeMap').onclick=()=>closeOverlay('mapOverlay');$('returnArea').onclick=()=>closeOverlay('mapOverlay');$('warehouseOverview').onclick=openWarehouseOverview;$('mapOverlay').addEventListener('click',e=>{if(e.target===$('mapOverlay'))closeOverlay('mapOverlay');});$('closeBlock').onclick=()=>closeBlock();$('doneBlock').onclick=()=>closeBlock();$('blockForm').onsubmit=submitBlockStock;$('blockClear').onclick=clearBlock;$('blockMove').onclick=transferBlock;$('blockAction').onchange=renderBlockDetails;$('blockItem').onchange=renderBlockDetails;$('blockOverlay').addEventListener('click',e=>{if(e.target===$('blockOverlay'))closeBlock();});$('copyBlockLink').onclick=async()=>{const link=blockLink(selectedId,selectedBlock);try{await navigator.clipboard.writeText(link);toast('Block link copied.');}catch{prompt('Copy block link:',link);}};$('ownerButton').onclick=()=>isOwner()?signOut():signIn();$('refresh').onclick=()=>loadData();$('setupButton').onclick=setupWarehouse;$('designButton').onclick=()=>{designing=!designing;drawingTool='select';render();};$('addAreaButton').onclick=()=>openAreaEditor();$('clearFilter').onclick=()=>{filterItem='';render();};$('search').addEventListener('input',renderList);document.querySelectorAll('[data-draw-tool]').forEach(button=>button.onclick=()=>{drawingTool=button.dataset.drawTool;syncDrawingPalette();$('board').classList.toggle('drawing-tool',drawingTool!=='select');});for(const [id,which,delta] of [['mapZoomOut','board',-1],['mapZoomIn','board',1],['mapZoomFit','board',0],['focusZoomOut','focusMapBoard',-1],['focusZoomIn','focusMapBoard',1],['focusZoomFit','focusMapBoard',0]])$(id).onclick=()=>adjustMapZoom(which,delta);window.addEventListener('resize',()=>{applyMapSize('board');applyMapSize('focusMapBoard');if(selectedId&&blockView==='lanes'&&lastBlockGroupSize!==blockGroupSize())renderBlockGrid(selectedId);});
 $('closeDetails').onclick=()=>closeDetails();$('doneDetails').onclick=()=>closeDetails();$('closeEditor').onclick=closeEditor;
 $('detailsOverlay').addEventListener('click',e=>{if(e.target===$('detailsOverlay'))closeDetails();});$('editorOverlay').addEventListener('click',e=>{if(e.target===$('editorOverlay'))closeEditor();});
 $('copyLink').onclick=async()=>{const url=areaURL(selectedId);try{if(navigator.clipboard?.writeText)await navigator.clipboard.writeText(url);else {const el=make('textarea');el.value=url;document.body.append(el);el.select();document.execCommand('copy');el.remove();}$('linkStatus').textContent='Copied!';}catch{prompt('Copy this area link:',url);}};
 document.querySelectorAll('[data-action]').forEach(el=>el.onclick=()=>setStockAction(el.dataset.action));$('itemType').onchange=toggleCustom;$('stockForm').onsubmit=addOrRemove;$('clearArea').onclick=clearArea;$('clearWholeArea').onclick=clearArea;$('assignLooseButton').onclick=assignLooseToBlock;$('transferButton').onclick=async()=>{const to=$('transferTo').value,from=selectedId;if(!to){toast('Choose a destination first.');return;}await moveContents(from,to);};$('areaForm').onsubmit=submitArea;$('editType').onchange=syncEditorType;$('archiveArea').onclick=archiveArea;
 $('board').addEventListener('pointerdown',drawStart);$('board').addEventListener('pointermove',e=>{dragMove(e);drawMove(e);});$('board').addEventListener('pointerup',e=>{dragEnd(e);drawEnd(e);});$('board').addEventListener('pointercancel',()=>{drag=null;cancelDrawing();renderBoard();});document.addEventListener('pointermove',stockPointerMove,{passive:false});document.addEventListener('pointerup',stockPointerUp);document.addEventListener('pointercancel',clearStockDrag);
 window.addEventListener('popstate',()=>{if(!state.loaded)return;closeOverlay('mapOverlay');openFromURL();});
 window.addEventListener('keydown',e=>{if(e.key==='Escape'){if($('editorOverlay').classList.contains('open'))closeEditor();else if($('mapOverlay').classList.contains('open'))closeOverlay('mapOverlay');else if($('blockOverlay').classList.contains('open'))closeBlock();else if($('detailsOverlay').classList.contains('open'))closeDetails();}});
 document.addEventListener('visibilitychange',()=>{if(!document.hidden&&mode==='viewer')loadData(false);});setInterval(()=>{if(!document.hidden&&mode==='viewer')loadData(false);},90000);
 window.warehouseGisReady=()=>{gisReady=true;};render();loadData();
}
initialise();
})();
