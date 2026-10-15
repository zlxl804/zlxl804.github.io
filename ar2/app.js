'use strict';
const C=window.MergeCore,files=Array(4).fill(null),slots=['서울본사 기존 통합본','서울본사 최신','대전지사 최신','부산지사 최신'];let result=null;
const $=id=>document.getElementById(id);
function status(text,error=false){$('status').textContent=text;$('status').className=error?'error':'';}
slots.forEach((label,i)=>{
 const el=document.createElement('label');el.className='upload';el.style.setProperty('--color','#'+(i?C.colors[i]:'929DB6'));
 const strong=document.createElement('strong');strong.textContent=label;
 const small=document.createElement('small');small.textContent='클릭하여 파일 선택';
 const input=document.createElement('input');input.type='file';input.accept='.xlsx,.xls';input.setAttribute('aria-label',label);
 input.addEventListener('change',()=>{files[i]=input.files[0]||null;small.textContent=files[i]?.name||'클릭하여 파일 선택';result=null;$('result').hidden=true;$('merge').disabled=!files.every(Boolean);$('merge').textContent=files.every(Boolean)?'통합 결과 생성':'4개 파일을 선택해 주세요';status('파일은 서버에 전송되지 않고 이 브라우저에서만 처리됩니다.');});
 el.append(strong,small,input);$('uploads').append(el);
});
$('merge').onclick=async()=>{
 result=null;$('result').hidden=true;$('merge').disabled=true;document.querySelectorAll('input[type=file]').forEach(e=>e.disabled=true);
 try{
  if(!window.XLSX)throw Error('엑셀 도구를 불러오지 못했습니다. 페이지를 새로고침해 주세요.');
  const inputs=[];
  for(let i=0;i<4;i++){
   status(`${slots[i]} 읽는 중…`);await new Promise(r=>setTimeout(r,20));
   try{inputs.push(C.parse(XLSX.read(await files[i].arrayBuffer(),{type:'array'}),XLSX));}catch(e){throw Error(`${slots[i]}: ${e.message}`);}
  }
  result=C.merge(inputs);$('stats').replaceChildren();
  ['세부항목 전체','기준 유지','서울 신규','대전 신규','부산 신규'].forEach((label,i)=>{const el=document.createElement('div');el.className='stat';if(i>1)el.style.background='#'+C.colors[i-1];el.textContent=label;const b=document.createElement('b');b.textContent=(i===0?result.items.length:result.counts[i-1])+'개';el.append(b);$('stats').append(el);});
  const fieldCount=new Set(result.items.map(item=>item.code.slice(0,3))).size;
  const fieldStat=document.createElement('div');fieldStat.className='stat';fieldStat.style.gridColumn='1 / -1';fieldStat.textContent='분야 수 (분류번호 앞 3자리)';const fieldValue=document.createElement('b');fieldValue.textContent=fieldCount+'개';fieldStat.append(fieldValue);$('stats').prepend(fieldStat);
  $('detail').textContent=`중복 제외 ${result.duplicates.reduce((a,b)=>a+b,0)}건 · 색상은 최초 추가 출처 기준입니다.`;
  $('result').hidden=false;render();status('통합이 완료되었습니다. 결과를 확인하고 엑셀을 다운로드하세요.');
 }catch(e){status(e.message,true);}finally{$('merge').disabled=!files.every(Boolean);document.querySelectorAll('input[type=file]').forEach(e=>e.disabled=false);}
};
function render(){if(!result)return;const q=$('search').value.trim().toLowerCase(),source=$('source').value;const items=result.items.filter(i=>(source===''||String(i.source)===source)&&(!q||(i.code+' '+i.name).toLowerCase().includes(q)));$('items').replaceChildren();let prev='';for(const item of items){const g=item.code.slice(0,3);if(g!==prev){const tr=document.createElement('tr');tr.className='group';const td=document.createElement('td');td.colSpan=3;td.textContent=g+'. '+(result.groups[g]||'분야명 미기재');tr.append(td);$('items').append(tr);prev=g;}const tr=document.createElement('tr');tr.className='source-'+item.source;[item.code,item.name,C.labels[item.source]].forEach(v=>{const td=document.createElement('td');td.textContent=v;tr.append(td);});$('items').append(tr);}$('shown').textContent=`${items.length}개 항목 표시 / 전체 ${result.items.length}개`;}
$('search').oninput=render;$('source').onchange=render;
$('download').onclick=()=>{if(!result)return;try{XLSX.writeFile(C.workbook(result,XLSX),'교정항목_분야별_통합.xlsx');}catch(e){status('다운로드 실패: '+e.message,true);}};
