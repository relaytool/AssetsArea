/* Public, read-only, multi-area picking planner. A column is entered at row N. */
(function(root){
 'use strict';
 const norm=value=>String(value??'').trim().toLocaleLowerCase('en-GB');
 const columnName=n=>{let label='';for(let x=n;x>0;x=Math.floor((x-1)/26))label=String.fromCharCode(65+(x-1)%26)+label;return label;};
 const compareNames=(a,b)=>String(a).localeCompare(String(b),undefined,{numeric:true});
 const placeCompare=(a,b)=>(Number(a.y)||0)-(Number(b.y)||0)||(Number(a.x)||0)-(Number(b.x)||0)||compareNames(a.id,b.id);
 function plan({areas=[],blockSettings=[],blockStock=[],looseStock=[],item='',quantity=0}){
  const requested=Number(quantity),asset=norm(item);
  if(!asset||!Number.isSafeInteger(requested)||requested<1)throw new Error('Choose an asset and enter a positive whole quantity.');
  const visibleAreas=areas.filter(a=>a.active!==false&&norm(a.type)!=='aisle');
  const areaById=new Map(visibleAreas.map(a=>[norm(a.id),a]));
  const settings=new Map(blockSettings.map(s=>[norm(s.areaId),s]));
  const stockAt=new Map();let totalLocated=0,unassigned=0;
  for(const s of looseStock)if(asset===norm(s.item)&&areaById.has(norm(s.areaId)))unassigned+=Number(s.quantity)||0;
  for(const s of blockStock){
   if(!areaById.has(norm(s.areaId)))continue;
   const qty=Number(s.quantity)||0;
   if(qty<=0)continue;
   if(asset===norm(s.item))totalLocated+=qty;
   const key=norm(s.areaId)+'|'+String(s.blockId||'').toUpperCase(),contents=stockAt.get(key)||[];
   contents.push(s);stockAt.set(key,contents);
  }
  const grouped=new Map();
  for(const area of visibleAreas){
   const grid=settings.get(norm(area.id)),rows=Math.min(30,Math.max(2,Number(grid?.rows)||8)),cols=Math.min(52,Math.max(1,Number(grid?.columns)||2));
   const lanes=[];
   for(let col=1;col<=cols;col++){
    const column=columnName(col),blocks=[];
    for(let row=rows;row>=1;row--){
     const blockId=column+String(row).padStart(2,'0'),contents=stockAt.get(norm(area.id)+'|'+blockId)||[];
     if(!contents.length)continue; // An empty position does not block access.
     const matching=contents.filter(s=>norm(s.item)===asset),amount=matching.reduce((sum,s)=>sum+(Number(s.quantity)||0),0);
     if(!amount)break; // A different asset type physically obstructs deeper rows.
     const mixed=contents.some(s=>norm(s.item)!==asset);
     blocks.push({areaId:area.id,areaName:area.name||area.id,blockId,column,row,quantity:amount,mixed});
     if(mixed)break; // The other asset in this block prevents going farther.
    }
    if(blocks.length)lanes.push({column,blocks,total:blocks.reduce((sum,b)=>sum+b.quantity,0)});
   }
   if(lanes.length)grouped.set(norm(area.id),{area,lanes,total:lanes.reduce((sum,l)=>sum+l.total,0)});
  }
  const accessible=[...grouped.values()].reduce((sum,a)=>sum+a.total,0);
  const blocked=Math.max(0,totalLocated-accessible);
  let remaining=requested;
  const steps=[],areasUsed=[];
  const stopsFor=(lane,need)=>{let qty=0,count=0;for(const block of lane.blocks){qty+=block.quantity;count++;if(qty>=need)break;}return count;};
  const groupStops=(group,need)=>{
   // Estimate the fewest front-to-back block visits inside this area.
   if(group.lanes.length===1)return stopsFor(group.lanes[0],need);
   let left=need,count=0;const lanes=group.lanes.map(l=>({lane:l,index:0}));
   while(left>0&&lanes.some(x=>x.index<x.lane.blocks.length)){
    lanes.sort((a,b)=>(b.lane.blocks[a.index]?.quantity||0)-(a.lane.blocks[b.index]?.quantity||0)||compareNames(a.lane.column,b.lane.column));
    const next=lanes[0];left-=next.lane.blocks[next.index++].quantity;count++;
   }
   return count;
  };
  // First select a *whole area*, then its accessible column fronts. Only leave
  // that area when it cannot supply the remaining requirement. This gives an
  // ordered route, minimises area changes and can span the entire warehouse.
  const available=[...grouped.values()];
  while(remaining>0&&available.length){
   available.sort((a,b)=>{
    const fitsA=a.total>=remaining,fitsB=b.total>=remaining;
    if(fitsA!==fitsB)return fitsA?-1:1;
    if(fitsA){const stops=groupStops(a,remaining)-groupStops(b,remaining);if(stops)return stops;}
    else if(a.total!==b.total)return b.total-a.total;
    return placeCompare(a.area,b.area);
   });
   const chosen=available.shift(),areaSteps=[];
   const lanes=[...chosen.lanes];
   while(remaining>0&&lanes.length){
    lanes.sort((a,b)=>{
     const fitsA=a.total>=remaining,fitsB=b.total>=remaining;
     if(fitsA!==fitsB)return fitsA?-1:1;
     if(fitsA){const stops=stopsFor(a,remaining)-stopsFor(b,remaining);if(stops)return stops;}
     else if(a.total!==b.total)return b.total-a.total;
     return compareNames(a.column,b.column);
    });
    const lane=lanes.shift();
    for(const block of lane.blocks){
     if(!remaining)break;
     const take=Math.min(remaining,block.quantity);
     const step={...block,take,number:steps.length+1};
     steps.push(step);areaSteps.push(step);remaining-=take;
    }
   }
   if(areaSteps.length)areasUsed.push({areaId:chosen.area.id,areaName:chosen.area.name||chosen.area.id,quantity:areaSteps.reduce((sum,s)=>sum+s.take,0),steps:areaSteps});
  }
  return {item,requested,steps,areasUsed,areaCount:areasUsed.length,blockCount:steps.length,planned:requested-remaining,shortfall:remaining,accessible,blocked,unassigned,totalLocated};
 }
 root.WarehouseRetrieval={plan};
 if(typeof module!=='undefined'&&module.exports)module.exports={plan};
})(typeof window!=='undefined'?window:globalThis);
