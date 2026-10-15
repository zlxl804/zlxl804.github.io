(function(root){
 'use strict';
 const labels=['기준파일','서울본사','대전지사','부산지사'];
 const colors=['FFFFFF','FFF2A6','FFD09B','BEDCFF'];
 const clean=v=>String(v??'').trim();
 function parse(wb,X){
  const name=wb.SheetNames.find(n=>n.replace(/\s/g,'')==='표지(국문)') || wb.SheetNames.find(n=>n==='통합항목');
  if(!name) throw Error('표지(국문) 또는 통합항목 시트를 찾을 수 없습니다.');
  const rows=X.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true});
  let header=-1,cols=[];
  for(let i=0;i<Math.min(rows.length,40);i++){
   const found=[]; rows[i].forEach((v,c)=>{if(clean(v)==='분류번호' && clean(rows[i][c+1])==='교정항목') found.push(c);});
   if(found.length){header=i;cols=found;break;}
  }
  if(header<0) throw Error('분류번호 / 교정항목 머리글을 찾을 수 없습니다.');
  const items=[],groups={};
  for(const c of cols){let last=null;
   for(let i=header+1;i<rows.length;i++){
    const code=clean(rows[i][c]),text=clean(rows[i][c+1]);
    const group=code.match(/^(\d{3})\s*[.．]\s*(.+)/);
    if(group){groups[group[1]]=group[2];last=null;continue;}
    if(/^\d{5}$/.test(code)){
     if(!text) throw Error(`${name} ${i+1}행: ${code} 항목명이 비어 있습니다.`);
     last={code,name:text}; items.push(last);
    } else if(code){last=null;}
    else if(last && text){last.name+=' '+text;}
   }
  }
  if(!items.length) throw Error('5자리 분류번호 항목이 없습니다.');
  return {items,groups,sheet:name};
 }
 function merge(inputs){
  const seen=new Map(),groups={},counts=[0,0,0,0],duplicates=[0,0,0,0];
  inputs.forEach((input,source)=>{
   Object.entries(input.groups).forEach(([k,v])=>{if(!groups[k])groups[k]=v;});
   input.items.forEach(item=>{
    if(seen.has(item.code)){duplicates[source]++;return;}
    seen.set(item.code,{...item,source});counts[source]++;
   });
  });
  return {items:[...seen.values()].sort((a,b)=>a.code.localeCompare(b.code)),groups,counts,duplicates};
 }
 function workbook(result,X){
  const rows=[['교정항목 분야별 통합'],['신규: 서울 노랑 / 대전 주황 / 부산 파랑'],['분류번호','교정항목','추가 출처']];
  const meta=[];let prev='';
  result.items.forEach(item=>{
   const group=item.code.slice(0,3);
   if(group!==prev){rows.push([group+'. '+(result.groups[group]||'분야명 미기재')]);meta.push({row:rows.length-1,group:true});prev=group;}
   rows.push([item.code,item.name,labels[item.source]]);meta.push({row:rows.length-1,source:item.source});
  });
  const s=X.utils.aoa_to_sheet(rows);
  s['!cols']=[{wch:18},{wch:70},{wch:16}];
  s['!rows']=rows.map(r=>({hpt:r[1]?Math.max(30,Math.ceil(String(r[1]).length/40)*20):28}));
  s['!merges']=[{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
  for(let r=0;r<rows.length;r++)for(let c=0;c<3;c++){
   const key=X.utils.encode_cell({r,c});if(!s[key])s[key]={t:'s',v:''};
   s[key].s={font:{name:'맑은 고딕',sz:11},alignment:{vertical:'center',wrapText:true},border:{bottom:{style:'hair',color:{rgb:'DCE3EE'}}}};
   if(r<3)s[key].s={...s[key].s,font:{name:'맑은 고딕',sz:r===0?16:11,bold:true,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'233D67'}}};
  }
  meta.forEach(m=>{for(let c=0;c<3;c++)s[X.utils.encode_cell({r:m.row,c})].s.fill={fgColor:{rgb:m.group?'E8EDF5':colors[m.source]}};if(m.group)s['!merges'].push({s:{r:m.row,c:0},e:{r:m.row,c:2}});});
  const wb=X.utils.book_new();X.utils.book_append_sheet(wb,s,'통합항목');return wb;
 }
 const api={parse,merge,workbook,labels,colors};if(typeof module!=='undefined')module.exports=api;else root.MergeCore=api;
})(typeof globalThis!=='undefined'?globalThis:this);
