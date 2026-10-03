'use strict';
const $=id=>document.getElementById(id);let d=null;let tournamentClasses=[],activeClassIndex=0,excelPreparation=null;
const say=s=>{$('msg').textContent=s};const key=(a,b)=>[a,b].sort((x,y)=>x-y).join(':');
function stats(){let z=Object.fromEntries(d.players.map(p=>[p.no,{wins:0,losses:0,byes:0,hist:''}]));for(let round of Object.keys(d.results).map(Number).sort((a,b)=>a-b))for(let r of d.results[round]){if(r.winner)z[r.winner].wins++;if(r.bye&&r.winner)z[r.winner].byes++;if(r.p2&&!r.forfeit){if(r.winner===r.p1)z[r.p2].losses++;else if(r.winner===r.p2)z[r.p1].losses++;z[r.p1].hist+=r.winner===r.p1?'○':'×';z[r.p2].hist+=r.winner===r.p2?'○':'×'}}return z}
function rank(){
 const st=stats(), all=Object.keys(d.results).map(Number).sort((a,b)=>a-b).flatMap(r=>d.results[r]);
 const head=(a,b)=>all.find(r=>r.p2&&!r.forfeit&&key(r.p1,r.p2)===key(a,b))?.winner||0;
 let rows=d.players.map(p=>{const opp=all.filter(r=>r.p2&&!r.forfeit&&(r.p1===p.no||r.p2===p.no)).map(r=>r.p1===p.no?r.p2:r.p1),beaten=all.filter(r=>r.p2&&!r.forfeit&&r.winner===p.no).map(r=>r.p1===p.no?r.p2:r.p1),vals=beaten.map(n=>st[n].wins).sort((a,b)=>a-b);return {no:p.no,name:p.name,wins:st[p.no].wins,losses:st[p.no].losses,sc:opp.reduce((s,n)=>s+st[n].wins,0),sb:vals.reduce((s,n)=>s+n,0),md:vals.length<=2?0:vals.slice(1,-1).reduce((s,n)=>s+n,0),direct:0,directText:'－'}});
 let ordered=[];
 for(const wins of [...new Set(rows.map(x=>x.wins))].sort((a,b)=>b-a)){
  const g=rows.filter(x=>x.wins===wins);
  if(g.length===2){const winner=head(g[0].no,g[1].no);if(winner)for(const x of g){x.direct=x.no===winner?1:0;x.directText=x.no===winner?'勝':'負'}}
  g.sort((a,b)=>b.direct-a.direct||b.sc-a.sc||b.sb-a.sb||b.md-a.md||a.no-b.no);
  ordered.push(...g);
 }
 let currentRank=1;
 for(let i=0;i<ordered.length;i++){
  const x=ordered[i],prev=ordered[i-1],sameCount=rows.filter(z=>z.wins===x.wins).length;
  const same=prev&&prev.wins===x.wins&&(sameCount!==2||prev.direct===x.direct)&&prev.sc===x.sc&&prev.sb===x.sb&&prev.md===x.md;
  if(i>0&&!same)currentRank=i+1;
  x.rank=currentRank;
  x.note=`勝数=${x.wins} / 負数=${x.losses} / 直接対決=${x.directText} / SC=${x.sc} / SB=${x.sb} / MD=${x.md}`;
 }
 const counts={};for(const x of ordered)counts[x.rank]=(counts[x.rank]||0)+1;
 for(const x of ordered)if(counts[x.rank]>1)x.note+=' / 最終規定対象';
 return ordered;
}
function played(){return new Set(Object.values(d.results).flat().filter(r=>r.p2&&!r.forfeit).map(r=>key(r.p1,r.p2)))}
function makePairings(round){let active=d.players.filter(p=>!p.withdrawn).map(p=>p.no),st=stats(),history=played();if(round===1){active.sort((a,b)=>a-b);let ans=[];for(let i=0;i+1<active.length;i+=2)ans.push({p1:active[i],p2:active[i+1]});if(active.length%2)ans.push({p1:active.at(-1),p2:0});return ans}let bye=null;if(active.length%2){bye=[...active].sort((a,b)=>st[a].byes-st[b].byes||st[a].wins-st[b].wins||b-a)[0];active=active.filter(x=>x!==bye)}let maxGap=active.length?Math.max(...Object.values(st).map(x=>x.wins))-Math.min(...Object.values(st).map(x=>x.wins)):0;for(let gap=0;gap<=Math.max(maxGap,d.rounds);gap++){let ids=[...active].sort((a,b)=>st[b].wins-st[a].wins||a-b),failed=new Set(),out=[];function feasible(list){let left=new Set(list);while(left.size){let q=[left.values().next().value],count=0;left.delete(q[0]);while(q.length){let a=q.shift();count++;let has=false;for(let b of list){if(a===b||history.has(key(a,b))||Math.abs(st[a].wins-st[b].wins)>gap)continue;has=true;if(left.delete(b))q.push(b)}if(!has)return false}if(count%2)return false}return true}function search(list){if(!list.length)return true;let state=[...list].sort((a,b)=>a-b).join(',');if(failed.has(state))return false;let a=list[0],cand=list.slice(1).filter(b=>!history.has(key(a,b))&&Math.abs(st[a].wins-st[b].wins)<=gap).sort((b,c)=>Math.abs(st[a].wins-st[b].wins)-Math.abs(st[a].wins-st[c].wins)||(st[a].hist===st[b].hist?0:1)-(st[a].hist===st[c].hist?0:1)||Math.abs(a-b)-Math.abs(a-c));for(let b of cand){out.push({p1:a,p2:b});if(search(list.filter(x=>x!==a&&x!==b)))return true;out.pop()}failed.add(state);return false}if(feasible(ids)&&search(ids)){if(bye!==null)out.push({p1:bye,p2:0});return out}}throw Error('再戦なしで組合せを作成できませんでした')}
function el(tag,text){let e=document.createElement(tag);if(text!==undefined)e.textContent=text;return e}function button(text,fn){let b=el('button',text);b.onclick=fn;return b}let sortByRank=false;let sortColumn='NO';let sortAscending=true;let currentAux='';
function opponent(no,round){let p=d.pairings[round]?.find(x=>x.p1===no||x.p2===no);return p?(p.p1===no?p.p2:p.p1):null}
function resultFor(no,round){let p=d.results[round]?.find(x=>x.p1===no||x.p2===no);return p?(p.forfeit?(p.winner===no?'不戦勝':'棄権'):(p.winner===no?'〇':'×')):''}
function render(){if(!d)return;
 const tabs=$('classTabs');tabs.replaceChildren();
 if(!tournamentClasses.length)tournamentClasses=[d];
 tournamentClasses.forEach((c,i)=>{
   const b=el('button',c.className);b.className='tab'+(i===activeClassIndex?' selected':'');
   b.style.background=i===activeClassIndex?'white':'#dedede';
   b.style.borderBottom=i===activeClassIndex?'2px solid #356d9b':'1px solid #aaa';
   b.onclick=()=>{activeClassIndex=i;d=tournamentClasses[i];sortByRank=false;sortColumn='NO';sortAscending=true;currentAux='';$('aux').hidden=true;render();};
   tabs.append(b);
 });
$('workspace').hidden=false;$('eventtitle').textContent=d.name;$('viewclass').value=d.className;$('viewplayers').value=d.players.length;$('viewrounds').value=d.rounds;$('phase').textContent=d.completed?'大会終了':d.currentRound?'進行：第'+d.currentRound+'回戦':'進行：準備中';$('status').textContent=d.completed?'大会終了：最終結果を確認してください。':d.currentRound?'現在：第'+d.currentRound+'回戦　勝敗を入力して［入力完了］を押してください。　✓ 自動保存':'第1回戦を作成してください。';$('start').disabled=!!d.currentRound;$('complete').disabled=!d.currentRound||d.completed;$('undo').disabled=!d.currentRound||!!d.undoUsed||(d.completed?!d.results[d.currentRound]:d.currentRound<2||!d.results[d.currentRound-1]);
let st=stats(),ranks=d.results&&Object.keys(d.results).length?rank():[],byNo=Object.fromEntries(ranks.map(x=>[x.no,x]));let ordered=[...d.players].sort((a,b)=>{
 const val=p=>{const x=byNo[p.no]||{};if(sortColumn==='NO')return p.no;if(sortColumn==='名前')return p.name;if(sortColumn==='棄権 フラグ')return Number(!!p.withdrawn);if(sortColumn==='順位')return x.rank??999;if(sortColumn==='勝数')return x.wins??0;if(sortColumn==='負数')return x.losses??0;if(sortColumn==='直接対決')return x.direct??0;if(['SC','SB','MD'].includes(sortColumn))return x[sortColumn.toLowerCase()]??0;if(sortColumn==='備考')return x.note??'';const m=sortColumn.match(/^第(\d+)回戦 (相手番号|勝敗)$/);if(m){const r=Number(m[1]);return m[2]==='相手番号'?(opponent(p.no,r)||0):(resultFor(p.no,r)||'')}return p.no};
 const av=val(a),bv=val(b);const cmp=typeof av==='string'?String(av).localeCompare(String(bv),'ja'):av-bv;return (sortAscending?cmp:-cmp)||a.no-b.no;
});let t=el('table'),head=el('tr');for(let h of ['NO','名前','棄権 フラグ',...Array.from({length:d.currentRound||0},(_,i)=>['第'+(i+1)+'回戦 相手番号','第'+(i+1)+'回戦 勝敗']).flat(),'順位','勝数','負数','直接対決','SC','SB','MD','備考']){let th=el('th',(h==='順位表示'?'順位':h).replace(' 相手番号','\n相手\n番号').replace(' 勝敗','\n勝敗'));if(h==='順位表示')th.className='rank-display';if(h==='NO')th.className='no';if(h==='名前')th.className='name';if(h==='棄権 フラグ')th.className='withdraw-cell';if(h.endsWith(' 相手番号'))th.className='opponent-cell';if(h.endsWith(' 勝敗'))th.className='result-cell';if(['順位','勝数','負数','SC','SB','MD'].includes(h))th.className='stat-cell';if(h==='直接対決')th.className='direct-cell';if(h==='備考')th.className='note';th.title=(h==='順位表示'?'順位':h)+'で並べ替え';th.style.cursor='pointer';th.onclick=()=>{if(h==='順位表示')h='順位';if(sortColumn===h)sortAscending=!sortAscending;else{sortColumn=h;sortAscending=true}sortByRank=sortColumn==='順位';render()};if(sortColumn===h)th.setAttribute('aria-sort',sortAscending?'ascending':'descending');head.append(th)}t.append(head);
for(let p of ordered){let x=byNo[p.no],tr=el('tr');if(d.completed&&x?.rank<=3)tr.className='rank'+x.rank;let no=el('td',p.no);no.className='no';let name=el('td',p.name);name.className='name';if(p.withdrawn){no.classList.add('withdraw-gray');name.classList.add('withdraw-gray')}tr.append(no,name);let w=el('td');let check=el('input');check.type='checkbox';check.checked=!!p.withdrawn;const lockedWithdrawal=!!p.withdrawn&&!!d.currentRound&&((p.withdrawnRound>0&&p.withdrawnRound<d.currentRound)||(p.withdrawnRound>0&&!!d.results[p.withdrawnRound]));check.disabled=!d.currentRound||d.completed||lockedWithdrawal;if(lockedWithdrawal)check.title='確定済みの棄権です。訂正するには「元に戻す」を使用してください。';check.onchange=()=>{if(lockedWithdrawal){check.checked=true;say('確定済みの棄権は解除できません。元に戻して訂正してください。');return;}p.withdrawn=check.checked;p.withdrawnRound=check.checked?d.currentRound:0;if(check.checked&&d.currentRound&&!d.results[d.currentRound]){const opp=opponent(p.no,d.currentRound);if(opp)delete d.drafts[key(p.no,opp)];}save();render();say(check.checked?'棄権者の当該回戦の勝敗を消去しました。対戦相手は入力完了時に不戦勝となります。':'棄権を解除しました。必要に応じて勝敗を入力してください。')};w.className='withdraw-cell';w.append(check);tr.append(w);
for(let round=1;round<=d.currentRound;round++){let opp=opponent(p.no,round);let opponentCell=el('td',opp||'');opponentCell.className='opponent-cell';if(p.withdrawn&&round>=(p.withdrawnRound||d.currentRound||1))opponentCell.classList.add('withdraw-gray');tr.append(opponentCell);let cell=el('td');cell.className='result-cell';if(p.withdrawn&&round>=(p.withdrawnRound||d.currentRound||1))cell.classList.add('withdraw-gray');if(round===d.currentRound&&!d.completed&&!d.results[round]){if(p.withdrawn){cell.textContent='棄権'}else if(opp===0){cell.textContent='〇'}else if(opp){let sel=el('select');let k=key(p.no,opp);for(let [v,label] of [['',''],['W','〇'],['L','×']]){let opt=el('option',label);opt.value=v;sel.append(opt)}let chosen=d.drafts[k]||'';sel.value=chosen?(Number(chosen)===p.no?'W':'L'):'';if(d.players.find(q=>q.no===opp)?.withdrawn){sel.disabled=true;sel.value='W';}sel.setAttribute('aria-label',`${p.no}番 ${p.name} 第${round}回戦 勝敗`);sel.dataset.player=String(p.no);sel.dataset.opp=String(opp);sel.onchange=()=>{if(sel.value)d.drafts[k]=String(sel.value==='W'?p.no:opp);else delete d.drafts[k];save();render();const next=document.querySelector(`select[data-player=\"${p.no}\"]`);if(next)next.focus();updateInputAssist()};sel.onkeydown=e=>{if(e.key==='o'||e.key==='O'||e.key==='〇'||e.key==='○'){e.preventDefault();sel.value='W';sel.onchange()}else if(e.key==='x'||e.key==='X'||e.key==='×'){e.preventDefault();sel.value='L';sel.onchange()}else if(e.key==='Delete'||e.key==='Backspace'){e.preventDefault();sel.value='';sel.onchange()}};cell.append(sel)}}else cell.textContent=resultFor(p.no,round);tr.append(cell)}for(let [i,val] of [x?.rank??'',st[p.no].wins,st[p.no].losses,x?.directText??'',x?.sc??0,x?.sb??0,x?.md??0,x?.note??''].entries()){let cell=el('td',String(val));cell.className=i===7?'note':i===3?'direct-cell':'stat-cell';tr.append(cell)}t.append(tr)}$('matches').replaceChildren(t);$('sortRank').textContent=sortByRank?'NO順':'順位順';updateInputAssist();if(currentAux)showAux(currentAux)}
// 転記は別ウィンドウ。開くたびに選択中クラスの最新データを表示する。
function openTransferWindow(){
 if(!d)return;
 const win=window.open('','yukuhashi_transfer','width=680,height=850,resizable=yes,scrollbars=yes');
 if(!win){say('転記画面を開けません。ブラウザでポップアップを許可してください。');return;}
 const doc=win.document;
 doc.open();doc.write('<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>転記</title><style>body{font-family:"Yu Gothic",Meiryo,sans-serif;margin:10px;color:#111}h1{font-size:19px;text-align:center;margin:6px 0}h2{font-size:16px;margin:6px 0}table{border-collapse:collapse;width:auto;max-width:none;table-layout:auto}th,td{border:1px solid #555;padding:5px 6px;text-align:center;font-size:19px;line-height:1.3;white-space:nowrap}th{background:#eee}th:first-child,td:first-child{min-width:42px}th:nth-child(2),td:nth-child(2){min-width:165px;text-align:left;white-space:nowrap}th:last-child,td:last-child{min-width:64px}body.final td:nth-child(3){font-weight:bold;font-size:23px}@media print{body{margin:0}button{display:none}}</style></head><body></body></html>');doc.close();
 doc.body.classList.toggle('final',!!d.completed);const h=doc.createElement('h1');h.textContent=d.name;doc.body.append(h);
 const sub=doc.createElement('h2');sub.textContent=d.className+'　'+(d.completed?'最終順位':'第'+d.currentRound+'回戦 対戦表');doc.body.append(sub);
 const table=doc.createElement('table'),head=doc.createElement('tr');
 for(const label of (d.completed?['NO','名前','順位','勝数','負数']:['NO','名前','相手NO'])){const th=doc.createElement('th');th.textContent=label;head.append(th)}table.append(head);
 const ranked=d.completed?rank():[];
 for(const player of [...d.players].sort((a,b)=>a.no-b.no)){
  const tr=doc.createElement('tr');
  const x=d.completed?ranked.find(x=>x.no===player.no):null;const values=d.completed?[player.no,player.name,x?.rank??'',x?.wins??'',x?.losses??'']:[player.no,player.name,opponent(player.no,d.currentRound)||'不戦'];
  for(const value of values){const td=doc.createElement('td');td.textContent=String(value);tr.append(td)}tr.style.background=['#ffffff','#e8f2fc','#fff7dc','#ebf8eb','#f8eeee'][table.rows.length%5];table.append(tr);
 }
 doc.body.append(table);const close=doc.createElement('button');close.textContent='閉じて元に戻る';close.onclick=()=>win.close();doc.body.append(close);win.onkeydown=e=>{if(e.key==='Escape')win.close()};win.focus();
}
// V127-style independent round-history window: round selector and complete per-player grid.
function openHistoryWindow(){
 if(!d||!d.currentRound)return say('回戦履歴は第1回戦作成後に表示できます。');
 const source=d;
 const win=window.open('','yukuhashi_history','width=1220,height=760,resizable=yes,scrollbars=yes');
 if(!win)return say('回戦履歴を開けません。ブラウザのポップアップを許可してください。');
 const doc=win.document;
 doc.open();doc.write('<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>回戦履歴</title><style>body{font-family:"Yu Gothic UI",Meiryo,sans-serif;margin:10px;background:#f0f0f0;font-size:14px}header{display:flex;gap:12px;align-items:center;margin-bottom:12px}button,select{padding:7px;font:inherit}#grid{max-height:calc(100vh - 85px);overflow:auto;background:white;border:1px solid #aaa}table{border-collapse:collapse;white-space:nowrap;font-size:14px;width:max-content;table-layout:fixed}th,td{border:1px solid #bbb;padding:2px 3px;text-align:center;box-sizing:border-box}th{background:#eee;position:sticky;top:0;height:49px;z-index:1}th.name,td.name{width:150px;min-width:150px;text-align:left}th.note,td.note{width:540px;min-width:540px;text-align:left;white-space:normal;overflow-wrap:anywhere}th.number,td.number{width:52px;min-width:52px}th.opponent,td.opponent{width:48px;min-width:48px}th.result,td.result{width:42px;min-width:42px}th.direct,td.direct{width:76px;min-width:76px}.rank1{background:#83cceb}.rank2{background:#f7c7ac}.rank3{background:#f191ea}</style></head><body><header><label for="round">表示する回戦</label><select id="round"></select><button id="close">最新画面に戻る</button></header><div id="grid"></div></body></html>');doc.close();
 doc.title=source.className+' 回戦履歴';
 const selector=doc.getElementById('round');
 for(let r=1;r<=source.currentRound;r++){const opt=doc.createElement('option');opt.value=String(r);opt.textContent='第'+r+'回戦';selector.append(opt)}
 if(source.completed){const opt=doc.createElement('option');opt.value='final';opt.textContent='最終結果';selector.append(opt)}
 selector.selectedIndex=selector.options.length-1;
 const close=()=>win.close();doc.getElementById('close').onclick=close;
 win.onkeydown=e=>{if(e.key==='Escape')close()};
 const add=(tr,value,cls)=>{const cell=doc.createElement(tr.parentNode?.tagName==='THEAD'?'th':'td');cell.textContent=value??'';if(cls)cell.className=cls;tr.append(cell)};
 const draw=()=>{
  const final=selector.value==='final';const r=final?source.currentRound:Number(selector.value);
  const snapshot={...source,currentRound:r,pairings:Object.fromEntries(Object.entries(source.pairings).filter(([n])=>Number(n)<=r)),results:Object.fromEntries(Object.entries(source.results).filter(([n])=>Number(n)<=r)),drafts:{},completed:final};
  const previous=d;let ranks;try{d=snapshot;ranks=rank()}finally{d=previous}
  const byNo=Object.fromEntries(ranks.map(x=>[x.no,x]));const t=doc.createElement('table');const thead=doc.createElement('thead'),head=doc.createElement('tr');thead.append(head);t.append(thead);
  if(final){for(const label of ['順位','NO','名前','勝数','負数','直接対決','SC','SB','MD','備考'])add(head,label,label==='名前'?'name':label==='備考'?'note':label==='直接対決'?'direct':'number')}
  else{for(const label of ['NO','名前'])add(head,label,label==='名前'?'name':'number');for(let rr=1;rr<=r;rr++){add(head,'第'+rr+'回戦 相手番号','opponent');add(head,'勝敗','result')}for(const label of ['順位','勝数','負数','備考'])add(head,label,label==='備考'?'note':'number')}
  const body=doc.createElement('tbody');t.append(body);
  const players=final?[...source.players].sort((a,b)=>(byNo[a.no]?.rank??999)-(byNo[b.no]?.rank??999)||a.no-b.no):[...source.players].sort((a,b)=>a.no-b.no);
  for(const p of players){const x=byNo[p.no]||{},tr=doc.createElement('tr');if(final&&x.rank<=3)tr.className='rank'+x.rank;body.append(tr);
   if(final){for(const [v,c] of [[x.rank,'number'],[p.no,'number'],[p.name,'name'],[x.wins,'number'],[x.losses,'number'],[x.directText,'direct'],[x.sc,'number'],[x.sb,'number'],[x.md,'number'],[x.note,'note']])add(tr,v,c)}
   else{add(tr,p.no,'number');add(tr,p.name,'name');for(let rr=1;rr<=r;rr++){const pair=(source.pairings[rr]||[]).find(q=>q.p1===p.no||q.p2===p.no);const opp=pair?(pair.p1===p.no?pair.p2:pair.p1):null;add(tr,opp||'','opponent');const result=(source.results[rr]||[]).find(q=>pair&&((q.p1===pair.p1&&q.p2===pair.p2)||(q.p1===pair.p2&&q.p2===pair.p1)));add(tr,result?(result.winner===p.no?'〇':'×'):(opp===0?'〇':''),'result')}for(const [v,c] of [[x.rank,'number'],[x.wins,'number'],[x.losses,'number'],[x.note,'note']])add(tr,v,c)}
  }
  doc.getElementById('grid').replaceChildren(t);
 };
 selector.onchange=draw;draw();win.focus();
}
function showAux(kind){if(!d)return;currentAux=kind;let body=$('auxBody');body.replaceChildren();$('aux').hidden=false;let rs=rank();if(kind==='transfer'){$('auxTitle').textContent='転記';let t=el('table');let h=el('tr');for(let s of ['NO','名前',d.completed?'順位':'相手NO'])h.append(el('th',s));t.append(h);for(let p of [...d.players].sort((a,b)=>a.no-b.no)){let tr=el('tr');for(let s of [p.no,p.name,d.completed?rs.find(x=>x.no===p.no)?.rank??'':opponent(p.no,d.currentRound)||''])tr.append(el('td',String(s)));t.append(tr)}body.append(t)}else if(kind==='history'){$('auxTitle').textContent='回戦履歴';for(let [round,results] of Object.entries(d.results)){let h=el('h4','第'+round+'回戦');body.append(h);let t=el('table');for(let r of results){let tr=el('tr');tr.append(el('td',String(r.p1)),el('td',r.p2?String(r.p2):'不戦勝'),el('td','勝者 '+r.winner));t.append(tr)}body.append(t)}}else{$('auxTitle').textContent='ヘルプ';
  const sections=[
    ['1．大会の準備と基本操作',[
      'iPadのSafariで開きます。大会データは端末内に自動保存されます。定期的に大会バックアップを保存してください。',
      '「大会準備」で準備用Excelを選択して「第1回戦を作成」を押します。以前の大会は「保存大会を再開」で開けます。',
      '「大会進行」でクラスのタブを選択し、対局の勝敗（〇・×）を入力します。一方の結果を入力すると相手側にも反映されます。',
      '全対局の入力後に「入力完了」を押すと、次の回戦の対戦表を作成します。未入力や矛盾がある場合は入力を確認してください。',
      '「転記」で掲示用の相手番号、「回戦履歴」で過去の結果、「順位順」で現在の順位順表示、「元に戻す」で直前の入力完了の取り消しができます。',
      'ヘルプ内の「自動入力」は未入力の対局だけにランダムで結果を入れるデモ・テスト用です。実際の大会結果の入力には使用しないでください。'
    ]],
    ['2．対戦相手の決め方（現在のWeb版の実装）',[
      '第1回戦は準備用Excelの名簿をランダムに並べ替え、大会番号を振り直してから2人ずつ組み合わせます。奇数人数の場合は不戦勝が発生します。',
      '第2回戦以降は再戦を避け、勝数が同じ人同士の対戦を優先します。組合せが成立しない場合は勝数差を広げて探索します。',
      '奇数人数の不戦勝は、過去の不戦勝が少ない人を優先し、その後に勝数などで選びます。不戦勝は1勝で、実対局の再戦履歴には含めません。',
      '途中棄権者は以降の組合せから除外します。再戦なしで組合せが成立しない場合はエラーを表示します。',
      '※Windows版V127の組合せとの完全一致は未検証です。明日のデモではWeb版の現在の動作として説明してください。'
    ]],
    ['3．順位・タイブレークの見方',[
      '順位は「勝数 → 直接対決 → SC → SB → MD」の順で比較します。',
      '勝数：勝った回数です。不戦勝も1勝として加えます。',
      '直接対決：同じ勝数の人がちょうど2人で、互いに対戦している場合、その勝者を優先します。3人以上の同勝数グループには適用しません。',
      'SC（ソルコフ）：実際に対戦した相手全員の最終勝数を合計した値です。',
      'SB：自分が実際に勝った相手の最終勝数を合計した値です。',
      'MD：このWeb版では、勝った相手の最終勝数を小さい順に並べ、最小と最大を除いた残りを合計します。勝った相手が2人以下なら0です。',
      '不戦勝は勝数に加えますが、SC・SB・MDに対戦相手としての点数を加えません。',
      '表の「備考」に勝数・直接対決・SC・SB・MDの値を表示します。同点が解消しない場合は同順位となることがあります。'
    ]]
  ];
  for(const [heading,lines] of sections){const h=el('h4',heading);body.append(h);const ul=el('ul');for(const line of lines){const li=el('li',line);li.style.marginBottom='7px';ul.append(li)}body.append(ul)}
  body.append(el('p','このWeb版は検証中です。V127との完全一致は実機試験が必要です。'));
  const auto=el('button','自動入力（未入力対局のみ・テスト用）');auto.type='button';auto.id='ipadAutoFill';auto.className='primary';auto.onclick=()=>{if(confirm('未入力の対局にランダムで勝敗を入力します。実際の大会では使用しないでください。実行しますか？')){autoFillCurrentRound();showAux('help')}};body.append(auto);
 }}
// V127 F5: fill only missing pairs, preserve existing results and withdrawals.
function autoFillCurrentRound(){
 if(!d||!d.currentRound||d.completed||d.results[d.currentRound])return say('入力可能な回戦がありません');
 const pairs=d.pairings[d.currentRound]||[];let filled=0,preserved=0;
 for(const m of pairs){
  if(!m.p2)continue;
  const a=d.players.find(p=>p.no===m.p1),b=d.players.find(p=>p.no===m.p2);
  if(a?.withdrawn||b?.withdrawn){preserved++;continue;}
  const k=key(m.p1,m.p2);
  if(d.drafts[k]){preserved++;continue;}
  const random=new Uint32Array(1);crypto.getRandomValues(random);
  d.drafts[k]=String(random[0]%2===0?m.p1:m.p2);filled++;
 }
 save();render();say(`F5：未入力 ${filled}対局を自動入力。入力済み・棄権 ${preserved}対局は変更していません。`);
}
function checkCurrentRound(){
 if(!d||!d.currentRound||d.completed)return ['入力可能な回戦がありません'];
 let errors=[];
 for(const m of d.pairings[d.currentRound]||[]){
  if(!m.p2)continue;
  const a=d.players.find(p=>p.no===m.p1),b=d.players.find(p=>p.no===m.p2);
  if(a?.withdrawn&&b?.withdrawn){errors.push(`NO.${m.p1} と NO.${m.p2}：両者とも棄権しています。`);continue;}
  if(a?.withdrawn||b?.withdrawn)continue;
  const winner=Number(d.drafts[key(m.p1,m.p2)]||0);
  if(winner!==m.p1&&winner!==m.p2)errors.push(`NO.${m.p1} と NO.${m.p2}：勝敗が未入力です。`);
 }
 return errors;
}
function showCheck(){const errors=checkCurrentRound();alert(errors.length?'F6 入力チェック（'+errors.length+'件）\n\n'+errors.join('\n'):'全行をチェックしました。問題はありません。');say(errors.length?'F6：'+errors.length+'件の問題があります':'F6：問題はありません');}


$('undo').onclick=async()=>{
 if(!d||!d.currentRound||d.undoUsed)return say('直前の入力完了は既に取り消しました。再度入力完了するまで戻せません。');
 if(!yukFolder.handle)return say('大会フォルダを選択してから操作してください');
 const finalRound=!!d.completed,prev=finalRound?d.currentRound:d.currentRound-1;
 if(!d.results[prev])return say('取り消せる直前回戦がありません');
 const question=finalRound?`大会終了を取り消し、第${prev}回戦の入力画面に戻しますか？`:`第${prev}回戦の入力完了を取り消しますか？ 現在の第${d.currentRound}回戦の組合せと入力は破棄されます。`;
 if(!confirm(question))return;
 try{
  await yukFolder.flush();
  const stamp=historyStamp();
  await yukFolder.write('大会進行データ_元に戻す前_'+stamp+'.json',new Blob([JSON.stringify({classes:tournamentClasses,active:activeClassIndex},null,2)],{type:'application/json'}));
  const rs=d.results[prev];if(!finalRound)delete d.pairings[d.currentRound];delete d.results[prev];d.currentRound=prev;d.completed=false;d.undoUsed=true;d.drafts={};
  for(const r of rs)if(r.p2&&!r.bye)d.drafts[key(r.p1,r.p2)]=String(r.winner);
  save();await yukFolder.flush();render();
  const blob=await exportTournamentExcel(tournamentClasses,rankForClass);
  await yukFolder.write('対戦表_大会用.xlsx',blob);
  say(`第${prev}回戦の入力画面に戻しました。大会フォルダにバックアップを保存し、大会用Excelを更新しました。棄権フラグを確認してください。`);
 }catch(err){say('元に戻す処理でエラーが発生しました：'+err.message+'。大会フォルダのバックアップを確認してください。')}
};

document.addEventListener('keydown',e=>{if(e.key==='F5'||e.key==='F6'){if($('workspace').hidden)return;e.preventDefault();if(e.key==='F5')autoFillCurrentRound();else showCheck();}});
$('transfer').onclick=()=>openTransferWindow();$('historyBtn').onclick=()=>openHistoryWindow();$('help').onclick=()=>showAux('help');$('closeAux').onclick=()=>{currentAux='';$('aux').hidden=true};$('sortRank').onclick=()=>{if(sortColumn==='順位'&&sortAscending){sortColumn='NO';sortAscending=true;sortByRank=false}else{sortColumn='順位';sortAscending=true;sortByRank=true}render()};$('navprep').onclick=()=>{$('prep').hidden=false;$('workspace').hidden=true;$('navprep').className='active';$('navprogress').className=''};$('navprogress').onclick=()=>{if(!d)return say('先に大会を作成してください');$('prep').hidden=true;$('workspace').hidden=false;$('navprep').className='';$('navprogress').className='active'};
function save(){if(d){
 tournamentClasses[activeClassIndex]=d;
 if(window.yukFolder?.handle)window.yukFolder.queueState({classes:tournamentClasses,active:activeClassIndex});
}}function validate(x){if(!x||!Array.isArray(x.players)||!x.players.length||typeof x.results!=='object'||typeof x.pairings!=='object'||!Number.isInteger(x.rounds))throw Error('大会データ形式が不正です')}
$('create').onclick=()=>{try{let names=$('names').value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean),rounds=Number($('rounds').value);if(names.length<2||names.length>50||new Set(names).size!==names.length||rounds<3||rounds>10||rounds>names.length-1)throw Error('人数2～50名・名前重複なし・回戦3～10かつ人数−1以下にしてください');if(!confirm('現在の大会データを新規大会で置き換えますか？'))return;tournamentClasses=[];activeClassIndex=0;d={name:$('tname').value,className:$('cname').value,rounds,players:names.map((name,i)=>({no:i+1,name,withdrawn:false,withdrawnRound:0})),pairings:{},results:{},drafts:{},currentRound:0,completed:false};save();render();$('navprogress').click();say('新規大会を作成しました')}catch(e){say(e.message)}};
$('start').onclick=()=>{try{d.pairings[1]=makePairings(1);d.currentRound=1;save();render();say('1回戦を作成しました')}catch(e){say(e.message)}};
let lastHistoryStamp='',historySerial=0;function historyStamp(){const now=new Date(),pad=n=>String(n).padStart(2,'0');const base=now.getFullYear()+pad(now.getMonth()+1)+pad(now.getDate())+pad(now.getHours())+pad(now.getMinutes())+pad(now.getSeconds());historySerial=base===lastHistoryStamp?historySerial+1:0;lastHistoryStamp=base;return base+(historySerial?'_'+String(historySerial).padStart(2,'0'):'')}
$('complete').onclick=async()=>{let beforeCommit;try{let errors=checkCurrentRound();if(errors.length)throw Error(errors.join('\n'));let round=d.currentRound,rs=[];for(let p of d.pairings[round]){if(!p.p2){rs.push({p1:p.p1,p2:0,winner:p.p1,bye:true});continue}let a=d.players.find(x=>x.no===p.p1),b=d.players.find(x=>x.no===p.p2),winner=Number(d.drafts[key(p.p1,p.p2)]);if(a.withdrawn&&b.withdrawn)throw Error(`${p.p1}・${p.p2}が両方棄権しています。試作版では未対応です`);if(a.withdrawn)winner=p.p2;if(b.withdrawn)winner=p.p1;if(!winner)throw Error(`${p.p1}－${p.p2}の勝敗が未入力です`);rs.push({p1:p.p1,p2:p.p2,winner,bye:!!(a.withdrawn||b.withdrawn),forfeit:!!(a.withdrawn||b.withdrawn)})}if(!confirm(`${round}回戦を確定しますか？`))return;beforeCommit=JSON.stringify(d);d.results[round]=rs;d.undoUsed=false;d.drafts={};if(round>=d.rounds){d.completed=true}else{d.pairings[round+1]=makePairings(round+1);d.currentRound++}// Keep a pre-commit snapshot. A failed disk save must not advance the round.
 await yukFolder.queueState({classes:tournamentClasses,active:activeClassIndex});
 const blob=await exportTournamentExcel(tournamentClasses,rankForClass);
 await yukFolder.write('対戦表_大会用_'+historyStamp()+'.xlsx',blob);
 render();say(`${round}回戦を確定し、大会用Excelと日時付き履歴Excelを保存しました`)}catch(e){if(typeof beforeCommit!=='undefined'){d=JSON.parse(beforeCommit);tournamentClasses[activeClassIndex]=d;render();try{await yukFolder.queueState({classes:tournamentClasses,active:activeClassIndex})}catch(_){/* report original save error */}}say('回戦は確定していません：'+e.message)}};
function downloadBackup(){if(!d)return say('大会がありません');const blob=new Blob([JSON.stringify(d,null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=el('a');a.href=url;a.download='yukuhashi-'+new Date().toISOString().replace(/[:.]/g,'-')+'.json';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);say('JSONの保存を開始しました。ダウンロードフォルダーにファイルがあることを確認してください。');}


async function restoreFromFile(file){if(!file)return;try{let x=JSON.parse(await file.text());validate(x);if(!confirm('現在の大会データを置き換えますか？ 保存していないデータは失われます。'))return;localStorage.setItem('yukuhashi-pwa-before-restore',JSON.stringify(d));d=x;save();render();$('prep').hidden=true;$('workspace').hidden=false;say('復元しました')}catch(err){say('復元失敗：'+err.message)}}


// 起動時は必ず大会準備。旧ブラウザ保存データを自動読込しない。
// 大会進行は利用者が選んだ大会フォルダのファイルだけから復元する。
$('prep').hidden=false;$('workspace').hidden=true;
$('folderMock').onclick=async()=>{try{let data=await yukFolder.pick();if(data){data.classes.forEach(validate);if(!confirm('選択した大会フォルダの進行状況を開きますか？')){yukFolder.disconnect();return;}tournamentClasses=data.classes;activeClassIndex=Math.min(data.active||0,data.classes.length-1);d=tournamentClasses[activeClassIndex];render();$('navprogress').click();say('大会フォルダから再開しました')}else{d=null;tournamentClasses=[];activeClassIndex=0;$('prep').hidden=false;$('workspace').hidden=true;say('空の大会フォルダを選択しました。準備用Excelを選択してください。')}}catch(e){say('フォルダ選択：'+e.message)}};
if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});

// Windows版準備画面のUI検証: 未実装のExcel機能を成功したように表示しない。
$('excelMock').addEventListener('change', async e=>{
 const f=e.target.files?.[0];$('excelPath').value=f?.name||'';$('excelCreate').disabled=true;excelPreparation=null;
 if(!f)return;
 try{
  excelPreparation=await parsePreparationXlsx(f);
  excelPreparation.sourceBytes=new Uint8Array(await f.arrayBuffer());
  $('excelNotice').textContent='✓ チェックOK　大会名：'+excelPreparation.name+'　'+excelPreparation.classes.length+'クラス・参加者合計 '+excelPreparation.classes.reduce((n,c)=>n+c.players.length,0)+'名。第1回戦を作成できます。';
  $('excelCreate').disabled=false;
 }catch(err){$('excelNotice').textContent='準備用Excelに問題があります：'+err.message}
});
function confirmReplaceTournament(){
 return new Promise(resolve=>{
  const overlay=document.createElement('div');
  overlay.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.48);z-index:99999;display:flex;align-items:center;justify-content:center;padding:18px';
  const panel=document.createElement('div');
  panel.style.cssText='background:white;color:#202b3b;border-radius:12px;padding:22px;max-width:460px;width:100%;font-size:18px';
  const msg=document.createElement('p');msg.textContent='以前の大会をバックアップしますか？ 「いいえ」：保存せず新しい大会を作成。「はい」：バックアップしてから作成。';
  const actions=document.createElement('div');actions.style.cssText='display:flex;justify-content:flex-end;gap:16px;margin-top:22px';
  for(const [label,answer] of [['いいえ',false],['はい',true]]){
   const button=document.createElement('button');button.type='button';button.textContent=label;
   button.style.cssText='min-width:100px;padding:12px;font-size:18px';
   button.onclick=()=>{overlay.remove();resolve(answer)};actions.append(button);
  }
  panel.append(msg,actions);overlay.append(panel);document.body.append(overlay);
 });
}
$('excelCreate').onclick=async()=>{
 if(!excelPreparation)return;
 if(!window.yukFolder?.handle)return say('先に大会フォルダを選択してください。');
 if(await yukFolder.hasTournamentFiles()){
  const backupRequested=await confirmReplaceTournament();
  if(backupRequested){try{await yukFolder.flush();await yukFolder.backupTournamentFiles()}catch(err){return say('バックアップに失敗しました。新しい大会は作成していません：'+err.message)}}
 }
 tournamentClasses=excelPreparation.classes.map(c=>{
  // V127: 初回に名簿そのものをシャッフルし、その順番で大会NOを確定する。
  const shuffled=[...c.players];
  for(let i=shuffled.length-1;i>0;i--){let j=Math.floor(Math.random()*(i+1));[shuffled[i],shuffled[j]]=[shuffled[j],shuffled[i]]}
  const players=shuffled.map((p,i)=>({...p,no:i+1,withdrawn:false,withdrawnRound:0}));
  const pairs=[];for(let i=0;i<players.length;i+=2)pairs.push({p1:players[i].no,p2:i+1<players.length?players[i+1].no:0});
  return {name:excelPreparation.name,className:c.name,templateSheetName:c.sheetName,rounds:c.rounds,players,pairings:{1:pairs},results:{},drafts:{},currentRound:1,completed:false};
 });
 activeClassIndex=0;d=tournamentClasses[0];
 try{setExcelTemplate(excelPreparation.sourceBytes);await yukFolder.write('対戦表_準備用.xlsx',new Blob([excelPreparation.sourceBytes]));await exportTournamentExcel(tournamentClasses,rankForClass)}catch(err){return say('大会用Excelの作成に失敗しました：'+err.message)}
 save();await yukFolder.flush();render();$('navprogress').click();
 say('全'+tournamentClasses.length+'クラスの第1回戦を作成し、対戦表_大会用.xlsxを出力しました。');
};
$('exitPrep').onclick=()=>{$('navprogress').click()};

// 表内入力補助：対戦ペア単位の未入力チェック。既存の対戦相手への自動反映を使用。
function updateInputAssist(){
 const box=document.getElementById('inputAssist');if(!box||!d)return;
 if(!d.currentRound||d.completed){box.textContent=d.completed?'全回戦終了':'第1回戦の作成待ち';return}
 const pairs=d.pairings[d.currentRound]||[];let pending=[];let completed=0;
 for(const pair of pairs){if(!pair.p2)continue;const a=d.players.find(p=>p.no===pair.p1),b=d.players.find(p=>p.no===pair.p2);if(a?.withdrawn||b?.withdrawn||d.drafts[key(pair.p1,pair.p2)])completed++;else pending.push(pair.p1+'－'+pair.p2)}
 box.textContent=`入力済み ${completed}/${completed+pending.length}対局　未入力 ${pending.length}対局`+(pending.length?'（番号：'+pending.join('、')+'）':'　入力完了できます');
}

function rankForClass(cls){const old=d;try{d=cls;return rank()}finally{d=old}}
