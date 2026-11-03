/* Public whole-block warehouse retrieval planner.
   Every column is accessed from row N towards row 1. Previously occupied
   blocks can be cleared in order to reach deeper target assets. */
(function(root){
 'use strict';
 const norm=value=>String(value??'').trim().toLocaleLowerCase('en-GB');
 const columnName=n=>{let label='';for(let x=n;x>0;x=Math.floor((x-1)/26))label=String.fromCharCode(65+(x-1)%26)+label;return label;};
 const compareNames=(a,b)=>String(a).localeCompare(String(b),undefined,{numeric:true});
 const placeCompare=(a,b)=>(Number(a.y)||0)-(Number(b.y)||0)||(Number(a.x)||0)-(Number(b.x)||0)||compareNames(a.id,b.id);
 const whole=n=>Number.isSafeInteger(n)&&n>0;
 const MAX_STATES=14000;
 // Having met the required target quantity (including overpick), prefer fewer
 // emptied blocks, then fewer areas, then fewer unrelated items displaced.
 function better(a,b){return !b||a.visits<b.visits||(a.visits===b.visits&&(a.areaCount<b.areaCount||(a.areaCount===b.areaCount&&a.otherQty<b.otherQty)));}
 function trimStates(map,target){
  if(map.size<=MAX_STATES)return map;
  const all=[...map.values()].sort((a,b)=>a.qty-b.qty),picked=new Map();
  const keep=s=>{if(picked.size<MAX_STATES&&!picked.has(s.qty))picked.set(s.qty,s);};
  const below=all.filter(s=>s.qty<target),above=all.filter(s=>s.qty>=target);
  for(const s of above.slice(0,Math.floor(MAX_STATES*.18)))keep(s);
  for(const s of below.slice(-Math.floor(MAX_STATES*.40)))keep(s);
  const slots=MAX_STATES-picked.size;
  for(let i=0;i<slots;i++)keep(all[Math.floor(i*(all.length-1)/Math.max(1,slots-1))]);
  if(picked.size<MAX_STATES)for(const s of [...all].sort((a,b)=>Math.abs(a.qty-target)-Math.abs(b.qty-target)))keep(s);
  return picked;
 }
 function choosePrefixes(lanes,requested){
  const total=lanes.reduce((s,l)=>s+l.total,0);
  if(total<requested)return lanes.map(l=>({lane:l,count:l.blocks.length})).filter(s=>s.count);
  const maxBlock=Math.max(0,...lanes.flatMap(l=>l.blocks.map(b=>b.quantity))),ceiling=requested+maxBlock;
  let states=new Map([[0,{qty:0,visits:0,areaCount:0,areaMask:0n,otherQty:0,previous:null,lane:null,count:0}]]);
  for(const lane of lanes){
   const next=new Map(states),bit=1n<<BigInt(lane.areaIndex);
   const options=[];let qty=0,otherQty=0;
   for(let count=1;count<=lane.blocks.length;count++){
    const block=lane.blocks[count-1];qty+=block.quantity;otherQty+=block.otherQty;
    // Do not clear blocks for no target gain. Blockers are included in a later prefix.
    if(block.quantity===0||qty>ceiling)continue;
    options.push({count,qty,otherQty});
   }
   for(const state of states.values())for(const option of options){
    const sum=state.qty+option.qty;if(sum>ceiling)continue;
    const newArea=(state.areaMask&bit)===0n;
    const candidate={qty:sum,visits:state.visits+option.count,areaCount:state.areaCount+(newArea?1:0),areaMask:state.areaMask|bit,otherQty:state.otherQty+option.otherQty,previous:state,lane,count:option.count};
    if(better(candidate,next.get(sum)))next.set(sum,candidate);
   }
   states=trimStates(next,requested);
  }
  let best;
  for(const candidate of states.values())if(candidate.qty>=requested&&(!best||candidate.qty<best.qty||(candidate.qty===best.qty&&better(candidate,best))))best=candidate;
  if(!best){let remaining=requested;const selections=[];for(const lane of lanes){let count=0;while(count<lane.blocks.length&&remaining>0)remaining-=lane.blocks[count++].quantity;if(count)selections.push({lane,count});if(remaining<=0)break;}return selections;}
  const picks=[];for(let s=best;s&&s.lane;s=s.previous)if(s.count)picks.push({lane:s.lane,count:s.count});return picks.reverse();
 }
 function plan({areas=[],blockSettings=[],blockStock=[],looseStock=[],item='',quantity=0}){
  const requested=Number(quantity),asset=norm(item);
  if(!asset||!whole(requested))throw new Error('Choose an asset and enter a positive whole quantity.');
  const visibleAreas=areas.filter(a=>a.active!==false&&norm(a.type)!=='aisle').sort(placeCompare);
  const areaById=new Map(visibleAreas.map(a=>[norm(a.id),a])),settings=new Map(blockSettings.map(s=>[norm(s.areaId),s]));
  const stockAt=new Map();let totalLocated=0,unassigned=0;
  for(const s of looseStock)if(asset===norm(s.item)&&areaById.has(norm(s.areaId)))unassigned+=Number(s.quantity)||0;
  for(const s of blockStock){
   if(!areaById.has(norm(s.areaId)))continue;
   const qty=Number(s.quantity)||0;if(qty<=0)continue;
   if(asset===norm(s.item))totalLocated+=qty;
   const key=norm(s.areaId)+'|'+String(s.blockId||'').toUpperCase(),contents=stockAt.get(key)||[];
   contents.push(s);stockAt.set(key,contents);
  }
  const lanes=[];
  for(let index=0;index<visibleAreas.length;index++){
   const area=visibleAreas[index],grid=settings.get(norm(area.id));
   const rows=Math.min(30,Math.max(2,Number(grid?.rows)||8)),cols=Math.min(52,Math.max(1,Number(grid?.columns)||2));
   for(let col=1;col<=cols;col++){
    const column=columnName(col),blocks=[];
    for(let row=rows;row>=1;row--){
     const blockId=column+String(row).padStart(2,'0'),contents=stockAt.get(norm(area.id)+'|'+blockId)||[];
     if(!contents.length)continue;
     const matching=contents.filter(s=>norm(s.item)===asset),amount=matching.reduce((sum,s)=>sum+(Number(s.quantity)||0),0);
     const otherItems=contents.filter(s=>norm(s.item)!==asset).map(s=>({item:s.item,quantity:Number(s.quantity)||0}));
     blocks.push({areaId:area.id,areaName:area.name||area.id,blockId,column,row,quantity:amount,otherItems,otherQty:otherItems.reduce((sum,s)=>sum+s.quantity,0),mixed:otherItems.length>0,kind:amount>0?'collect':'clear'});
    }
    // Avoid routes which clear entire columns containing no requested asset.
    while(blocks.length&&!blocks[blocks.length-1].quantity)blocks.pop();
    if(blocks.length)lanes.push({areaId:area.id,areaName:area.name||area.id,areaIndex:index,column,blocks,total:blocks.reduce((sum,b)=>sum+b.quantity,0)});
   }
  }
  const accessible=lanes.reduce((sum,lane)=>sum+lane.total,0);
  const picks=choosePrefixes(lanes,requested),steps=[],groups=new Map();
  // Collect in the real front-to-back order. A preceding clearance step makes
  // the following block accessible without needing a new search.
  for(const {lane,count} of picks)for(const block of lane.blocks.slice(0,count)){
   const step={...block,take:block.quantity,number:steps.length+1};steps.push(step);
   const key=norm(block.areaId),group=groups.get(key)||{areaId:block.areaId,areaName:block.areaName,quantity:0,steps:[]};
   group.quantity+=step.quantity;group.steps.push(step);groups.set(key,group);
  }
  const planned=steps.reduce((sum,s)=>sum+s.quantity,0),shortfall=Math.max(0,requested-planned),otherRemoved=steps.reduce((sum,s)=>sum+s.otherQty,0);
  const clearanceSteps=steps.filter(s=>s.kind==='clear');
  return {item,requested,steps,areasUsed:[...groups.values()],areaCount:groups.size,blockCount:steps.length,planned,shortfall,excess:Math.max(0,planned-requested),otherRemoved,clearanceCount:clearanceSteps.length,accessible,blocked:0,unassigned,totalLocated};
 }
 root.WarehouseRetrieval={plan};
 if(typeof module!=='undefined'&&module.exports)module.exports={plan};
})(typeof window!=='undefined'?window:globalThis);
