/* Read-only retrieval planning. Each column is approached from its highest row (N). */
(function(root){
 'use strict';
 const norm=value=>String(value??'').trim().toLocaleLowerCase('en-GB');
 const columnName=n=>{let label='';for(let x=n;x>0;x=Math.floor((x-1)/26))label=String.fromCharCode(65+(x-1)%26)+label;return label;};
 function plan({areas=[],blockSettings=[],blockStock=[],looseStock=[],item='',quantity=0}){
  const requested=Number(quantity),asset=norm(item);
  if(!asset||!Number.isSafeInteger(requested)||requested<1)throw new Error('Choose an asset and enter a positive whole quantity.');
  const lanes=[],visibleAreas=areas.filter(a=>a.active!==false&&norm(a.type)!=='aisle');
  const settings=new Map(blockSettings.map(s=>[norm(s.areaId),s]));
  const at=new Map();let totalLocated=0,unassigned=0;
  for(const s of looseStock)if(asset===norm(s.item)&&visibleAreas.some(a=>norm(a.id)===norm(s.areaId)))unassigned+=Number(s.quantity)||0;
  for(const s of blockStock){
   if(!visibleAreas.some(a=>norm(a.id)===norm(s.areaId)))continue;
   if(asset===norm(s.item))totalLocated+=Number(s.quantity)||0;
   const key=norm(s.areaId)+'|'+String(s.blockId||'').toUpperCase(),list=at.get(key)||[];list.push(s);at.set(key,list);
  }
  for(const area of visibleAreas){
   const grid=settings.get(norm(area.id)),rows=Number(grid?.rows)||8,cols=Number(grid?.columns)||2;
   for(let col=1;col<=cols;col++){
    const column=columnName(col),blocks=[];
    for(let row=rows;row>=1;row--){
     const blockId=column+String(row).padStart(2,'0'),contents=at.get(norm(area.id)+'|'+blockId)||[];
     if(!contents.length)continue; // Empty blocks do not obstruct the route.
     const matching=contents.filter(s=>norm(s.item)===asset),amount=matching.reduce((sum,s)=>sum+(Number(s.quantity)||0),0);
     if(!amount)break; // Another asset physically blocks any deeper stock in this column.
     const mixed=contents.some(s=>norm(s.item)!==asset);
     blocks.push({areaId:area.id,areaName:area.name,blockId,column,row,quantity:amount,mixed});
     if(mixed)break; // Taking this asset alone cannot clear the mixed block.
    }
    if(blocks.length)lanes.push({area,blocks,total:blocks.reduce((sum,b)=>sum+b.quantity,0)});
   }
  }
  const accessible=lanes.reduce((sum,l)=>sum+l.total,0),blocked=Math.max(0,totalLocated-accessible);
  let remaining=requested;const steps=[],available=[...lanes];
  while(remaining>0&&available.length){
   const stopsFor=(lane,needed)=>{let collected=0,stops=0;for(const b of lane.blocks){collected+=b.quantity;stops++;if(collected>=needed)break;}return stops;};
   available.sort((a,b)=>{
    const fitsA=a.total>=remaining,fitsB=b.total>=remaining;
    if(fitsA!==fitsB)return fitsA?-1:1;
    if(fitsA){const d=stopsFor(a,remaining)-stopsFor(b,remaining);if(d)return d;}
    else if(a.total!==b.total)return b.total-a.total;
    return (Number(a.area.y)||0)-(Number(b.area.y)||0)||(Number(a.area.x)||0)-(Number(b.area.x)||0)||a.area.id.localeCompare(b.area.id,undefined,{numeric:true})||a.blocks[0].column.localeCompare(b.blocks[0].column,undefined,{numeric:true});
   });
   const lane=available.shift();
   for(const block of lane.blocks){
    if(remaining===0)break;
    const take=Math.min(remaining,block.quantity);
    steps.push({...block,take});remaining-=take;
   }
  }
  return {item,requested,steps,planned:requested-remaining,shortfall:remaining,accessible,blocked,unassigned,totalLocated};
 }
 root.WarehouseRetrieval={plan};
 if(typeof module!=='undefined'&&module.exports)module.exports={plan};
})(typeof window!=='undefined'?window:globalThis);
