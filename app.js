'use strict';
const $=id=>document.getElementById(id);let d=null;let tournamentClasses=[],activeClassIndex=0,excelPreparation=null;
const say=s=>{$('msg').textContent=s};const key=(a,b)=>[a,b].sort((x,y)=>x-y).join(':');
function stats(){let z=Object.fromEntries(d.players.map(p=>[p.no,{wins:0,losses:0,byes:0,hist:''}]));for(let round of Object.keys(d.results).map(Number).sort((a,b)=>a-b))for(let r of d.results[round]){if(r.winner)z[r.winner].wins++;if(r.bye)z[r.winner].byes++;if(r.p2){if(r.winner===r.p1)z[r.p2].losses++;else if(r.winner===r.p2)z[r.p1].losses++;z[r.p1].hist+=r.winner===r.p1?'○':'×';z[r.p2].hist+=r.winner===r.p2?'○':'×'}}return z}
function rank(){
 const st=stats(), all=Object.keys(d.results).map(Number).sort((a,b)=>a-b).flatMap(r=>d.results[r]);
 const head=(a,b)=>all.find(r=>r.p2&&key(r.p1,r.p2)===key(a,b))?.winner||0;
 let rows=d.players.map(p=>{const opp=all.filter(r=>r.p2&&(r.p1===p.no||r.p2===p.no)).map(r=>r.p1===p.no?r.p2:r.p1),beaten=all.filter(r=>r.p2&&r.winner===p.no).map(r=>r.p1===p.no?r.p2:r.p1),vals=beaten.map(n=>st[n].wins).sort((a,b)=>a-b);return {no:p.no,name:p.name,wins:st[p.no].wins,losses:st[p.no].losses,sc:opp.reduce((s,n)=>s+st[n].wins,0),sb:vals.reduce((s,n)=>s+n,0),md:vals.length<=2?0:vals.slice(1,-1).reduce((s,n)=>s+n,0),direct:0,directText:'－'}});
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
function played(){return new Set(Object.values(d.results).flat().filter(r=>r.p2).map(r=>key(r.p1,r.p2)))}
function makePairings(round){
 let active=d.players.filter(p=>!p.withdrawn).map(p=>p.no),st=stats(),history=played();
 const allActive=[...active];
 if(!d.pairingMeta)d.pairingMeta={};
 if(round===1){active.sort((a,b)=>a-b);let ans=[];for(let i=0;i+1<active.length;i+=2)ans.push({p1:active[i],p2:active[i+1],gap:0,special:false,specialReason:'',reason:'第1回戦'});if(active.length%2)ans.push({p1:active.at(-1),p2:0,gap:0,special:false,specialReason:'',reason:'不戦勝'});d.pairingMeta[round]={lookahead:false,maxGap:0,reason:'第1回戦は初期組合せ'};return ans}
 let bye=null;if(active.length%2){bye=[...active].sort((a,b)=>st[a].byes-st[b].byes||st[a].wins-st[b].wins||b-a)[0];active=active.filter(x=>x!==bye)}
 const remaining=Math.max(0,d.rounds-round),exactLookahead=active.length<=14,futureMemo=new Map();
 function futurePossible(baseHistory,players,roundsLeft){
  if(roundsLeft<=0||players.length<2)return true;
  const memoKey=[...baseHistory].sort().join(',')+'|'+roundsLeft+'|'+players.join(',');if(futureMemo.has(memoKey))return futureMemo.get(memoKey);
  // 小規模クラスは残り全回戦を再帰的に確認する。大人数は次回戦の成立を確認する。
  const depth=exactLookahead?roundsLeft:Math.min(1,roundsLeft);let budget=exactLookahead?400000:30000,exhausted=false;
  function rec(hist,left){
   if(left<=0)return true;if(--budget<0){exhausted=true;return false}
   let ids=[...players];
   const byeChoices=ids.length%2?ids:[null];
   for(const bbye of byeChoices){
    let work=bbye===null?ids:ids.filter(x=>x!==bbye),pairs=[];
    function match(list){
     if(!list.length){let nh=new Set(hist);for(const [a,b] of pairs)nh.add(key(a,b));return rec(nh,left-1)}
     let best=list[0],bestCand=null;
     for(const a of list){let c=list.filter(b=>b!==a&&!hist.has(key(a,b)));if(bestCand===null||c.length<bestCand.length){best=a;bestCand=c}}
     if(!bestCand.length)return false;
     bestCand.sort((a,b)=>a-b);
     for(const b of bestCand){pairs.push([best,b]);if(match(list.filter(x=>x!==best&&x!==b)))return true;pairs.pop()}
     return false;
    }
    if(match(work))return true;
   }
   return false;
  }
  const ok=rec(new Set(baseHistory),depth)&&!exhausted;futureMemo.set(memoKey,ok);return ok;
 }
 let maxGap=active.length?Math.max(...active.map(x=>st[x].wins))-Math.min(...active.map(x=>st[x].wins)):0;
 const pairKey=q=>key(q.p1,q.p2);
 for(let gap=0;gap<=Math.max(maxGap,d.rounds);gap++){
  let ids=[...active].sort((a,b)=>st[b].wins-st[a].wins||a-b),best=null,bestScore=null,baseline=null,baselineScore=null,visited=0;
  function scoreOf(out){let gs=out.map(q=>Math.abs(st[q.p1].wins-st[q.p2].wins)).sort((a,b)=>b-a);return [Math.max(0,...gs),...gs,gs.reduce((a,b)=>a+b,0)]}
  function better(a,b){if(!b)return true;for(let i=0;i<Math.max(a.length,b.length);i++){let x=a[i]||0,y=b[i]||0;if(x!==y)return x<y}return false}
  function search(list,out){
   if(++visited>60000&&active.length>14)return;
   if(!list.length){
    let sc=scoreOf(out);if(better(sc,baselineScore)){baseline=out.map(q=>({...q}));baselineScore=sc}
    let nh=new Set(history);for(const q of out)nh.add(key(q.p1,q.p2));
    if(remaining&&!futurePossible(nh,[...active],remaining))return;
    if(better(sc,bestScore)){best=out.map(q=>({...q}));bestScore=sc}return;
   }
   let a=list[0],cand=list.slice(1).filter(b=>!history.has(key(a,b))&&Math.abs(st[a].wins-st[b].wins)<=gap).sort((b,c)=>Math.abs(st[a].wins-st[b].wins)-Math.abs(st[a].wins-st[c].wins)||Math.abs(a-b)-Math.abs(a-c));
   for(const b of cand){out.push({p1:a,p2:b});search(list.filter(x=>x!==a&&x!==b),out);out.pop();if(best&&active.length>14)return}
  }
  search(ids,[]);
  if(best){
   const baselineSet=new Set((baseline||[]).map(pairKey)),chosenSet=new Set(best.map(pairKey));
   const lookaheadAdjusted=!!baseline&&([...baselineSet].some(k=>!chosenSet.has(k))||[...chosenSet].some(k=>!baselineSet.has(k)));
   let mg=Math.max(0,...best.map(q=>Math.abs(st[q.p1].wins-st[q.p2].wins)));
   for(const q of best){
    q.gap=Math.abs(st[q.p1].wins-st[q.p2].wins);
    q.lookahead=lookaheadAdjusted&&!baselineSet.has(pairKey(q));
    const parts=[];if(q.lookahead)parts.push('先読み');if(q.gap>=2)parts.push(q.gap+'勝差');
    q.special=parts.length>0;q.specialReason=parts.length?parts.join('・')+'調整':'';
    q.reason=q.specialReason||(q.gap===0?'同星優先':'近い勝数を優先');
   }
   if(bye!==null)best.push({p1:bye,p2:0,gap:0,special:false,specialReason:'',reason:'不戦勝'});
   d.pairingMeta[round]={lookahead:remaining>0,lookaheadAdjusted,maxGap:mg,reason:lookaheadAdjusted?'先読みで将来行き詰まる候補を除外':'星差最小の成立可能な組合せ'};return best;
  }
 }
 // 将来の完全先読み候補が見つからない場合でも大会を停止させない。
 // 現在回で可能な実対局数を最大化し、残りを不戦勝にする。
 // 再戦は絶対に作らない。棄権等で必要なら複数不戦勝を許容する。
 let fallbackBest=null,fallbackScore=null;
 function fbBetter(a,b){if(!b)return true;for(let i=0;i<Math.max(a.length,b.length);i++){const x=a[i]??0,y=b[i]??0;if(x!==y)return x<y}return false}
 function fbSearch(left,pairs,byes){
  if(!left.length){
   const gaps=pairs.map(q=>Math.abs(st[q.p1].wins-st[q.p2].wins)).sort((a,b)=>b-a);
   const byeCounts=byes.map(n=>st[n].byes).sort((a,b)=>b-a);
   // 実対局数最大 → 不戦勝偏り最小 → 最大星差 → 星差総量
   const sc=[-pairs.length,Math.max(0,...byeCounts),byeCounts.reduce((a,b)=>a+b,0),Math.max(0,...gaps),gaps.reduce((a,b)=>a+b,0)];
   if(fbBetter(sc,fallbackScore)){fallbackScore=sc;fallbackBest={pairs:pairs.map(q=>({...q})),byes:[...byes]}}
   return;
  }
  // 対戦可能相手が少ない選手から処理して詰みを避ける
  let a=left[0],min=Infinity;
  for(const x of left){const c=left.filter(y=>y!==x&&!history.has(key(x,y))).length;if(c<min){min=c;a=x}}
  const rest=left.filter(x=>x!==a);
  const cand=rest.filter(b=>!history.has(key(a,b))).sort((b,c)=>Math.abs(st[a].wins-st[b].wins)-Math.abs(st[a].wins-st[c].wins)||st[b].byes-st[c].byes||Math.abs(a-b)-Math.abs(a-c));
  for(const b of cand){pairs.push({p1:a,p2:b});fbSearch(rest.filter(x=>x!==b),pairs,byes);pairs.pop()}
  // a を不戦勝にする枝。複数不戦勝もここで自然に扱う。
  byes.push(a);fbSearch(rest,pairs,byes);byes.pop();
 }
 fbSearch([...allActive],[],[]);
 if(fallbackBest){
  const out=fallbackBest.pairs;
  for(const q of out){q.gap=Math.abs(st[q.p1].wins-st[q.p2].wins);const parts=[];if(q.gap>=2)parts.push(q.gap+'勝差');q.special=parts.length>0;q.specialReason=parts.length?parts.join('・')+'調整':'';q.reason=q.specialReason||(q.gap===0?'同星優先':'近い勝数を優先')}
  for(const n of fallbackBest.byes)out.push({p1:n,p2:0,gap:0,special:false,specialReason:'',reason:fallbackBest.byes.length>1?'複数不戦勝調整':'不戦勝'});
  const mg=Math.max(0,...fallbackBest.pairs.map(q=>q.gap));
  d.pairingMeta[round]={lookahead:true,lookaheadAdjusted:true,maxGap:mg,reason:fallbackBest.byes.length>1?'再戦を避け大会を継続するため複数不戦勝を適用':'再戦を避け大会を継続できる最大実対局数を採用'};
  return out;
 }
 // active が0人でも停止させない。
 d.pairingMeta[round]={lookahead:true,lookaheadAdjusted:false,maxGap:0,reason:'対局可能な参加者なし'};
 return [];
}
function el(tag,text){let e=document.createElement(tag);if(text!==undefined)e.textContent=text;return e}function button(text,fn){let b=el('button',text);b.onclick=fn;return b}let sortByRank=false;let sortColumn='NO';let sortAscending=true;let currentAux='';
function opponent(no,round){let p=d.pairings[round]?.find(x=>x.p1===no||x.p2===no);return p?(p.p1===no?p.p2:p.p1):null}
function resultFor(no,round){let p=d.results[round]?.find(x=>x.p1===no||x.p2===no);return p?(p.winner===no?'〇':'×'):''}
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
$('workspace').hidden=false;$('eventtitle').textContent=d.name;$('viewclass').value=d.className;$('viewplayers').value=d.players.length;$('viewrounds').value=d.rounds;let waiting=!!d.awaitingNextRound;$('phase').textContent=d.completed?'大会終了':waiting?'進行：次回戦準備':d.currentRound?'進行：第'+d.currentRound+'回戦':'進行：第1回戦準備';$('status').textContent=d.completed?'大会終了：最終結果を確認してください。':waiting?`第${d.currentRound}回戦を確定しました。次回戦から棄権する参加者を確認して［第${d.currentRound+1}回戦の組合せを作成］を押してください。　✓ 自動保存`:d.currentRound?`現在：第${d.currentRound}回戦　対戦相手を確認し、必要なら棄権を訂正して組合せを作り直してください。対局後に勝敗を入力して［入力完了］を押してください。　✓ 自動保存`:'当日の欠席者を棄権にチェックしてから［第1回戦の組合せを作成］を押してください。　✓ 自動保存';const pairingRound=waiting?d.currentRound+1:(d.currentRound||1);$('start').hidden=d.completed;$('start').disabled=d.completed;$('start').textContent=`第${pairingRound}回戦の組合せを作成`;$('complete').disabled=!d.currentRound||d.completed||waiting||!!d.results[d.currentRound];const undoRound=(d.completed||waiting)?d.currentRound:(d.currentRound>=2?d.currentRound-1:0);$('undo').disabled=!undoRound||!!d.undoUsed||!d.results[undoRound];
let st=stats(),ranks=d.results&&Object.keys(d.results).length?rank():[],byNo=Object.fromEntries(ranks.map(x=>[x.no,x]));let ordered=[...d.players].sort((a,b)=>{
 const val=p=>{const x=byNo[p.no]||{};if(sortColumn==='NO')return p.no;if(sortColumn==='名前')return p.name;if(sortColumn==='棄権')return Number(!!p.withdrawn);if(sortColumn==='順位')return x.rank??999;if(sortColumn==='勝数')return x.wins??0;if(sortColumn==='負数')return x.losses??0;if(sortColumn==='直接対決')return x.direct??0;if(['SC','SB','MD'].includes(sortColumn))return x[sortColumn.toLowerCase()]??0;if(sortColumn==='備考')return x.note??'';const m=sortColumn.match(/^第(\d+)回戦 (相手番号|勝敗)$/);if(m){const r=Number(m[1]);return m[2]==='相手番号'?(opponent(p.no,r)||0):(resultFor(p.no,r)||'')}return p.no};
 const av=val(a),bv=val(b);const cmp=typeof av==='string'?String(av).localeCompare(String(bv),'ja'):av-bv;return (sortAscending?cmp:-cmp)||a.no-b.no;
});let meta=d.pairingMeta?.[d.currentRound];if(meta&&d.currentRound>1){let info=el('div');info.className='pairing-info';let adj=meta.lookaheadAdjusted?'【先読みで組合せ調整あり】':'【先読み確認済み】';info.textContent=`${adj} 第${d.currentRound}回戦：最大勝数差 ${meta.maxGap}　${meta.reason}`;$('matches').replaceChildren(info)}else $('matches').replaceChildren();let t=el('table'),head=el('tr');for(let h of ['NO','名前','棄権',...Array.from({length:d.currentRound||0},(_,i)=>['第'+(i+1)+'回戦 相手番号','第'+(i+1)+'回戦 勝敗']).flat(),'順位','勝数','負数','直接対決','SC','SB','MD','備考']){let th=el('th',h.replace(' 相手番号','\n相手\n番号').replace(' 勝敗','\n勝敗'));if(h==='NO')th.className='no';if(h==='名前')th.className='name';if(h==='棄権')th.className='withdraw-cell';if(h.endsWith(' 相手番号'))th.className='opponent-cell';if(h.endsWith(' 勝敗'))th.className='result-cell';if(['順位','勝数','負数','SC','SB','MD'].includes(h))th.className='stat-cell';if(h==='直接対決')th.className='direct-cell';if(h==='備考')th.className='note';th.title=h+'で並べ替え';th.style.cursor='pointer';th.onclick=()=>{if(sortColumn===h)sortAscending=!sortAscending;else{sortColumn=h;sortAscending=true}sortByRank=sortColumn==='順位';render()};if(sortColumn===h)th.setAttribute('aria-sort',sortAscending?'ascending':'descending');head.append(th)}t.append(head);
for(let p of ordered){let x=byNo[p.no],tr=el('tr');if(d.completed&&x?.rank<=3)tr.className='rank'+x.rank;let no=el('td',p.no);no.className='no';let name=el('td',p.name);name.className='name';if(p.withdrawn){no.classList.add('withdraw-gray');name.classList.add('withdraw-gray')}tr.append(no,name);let w=el('td');let check=el('input');check.type='checkbox';check.checked=!!p.withdrawn;{const targetRound=d.currentRound?(waiting?d.currentRound+1:d.currentRound):1;check.disabled=d.completed||(p.withdrawn&&p.withdrawnRound<targetRound);check.onchange=()=>{if(p.withdrawn&&p.withdrawnRound<targetRound){check.checked=true;alert('過去の回戦で確定した棄権は解除できません。');return;}p.withdrawn=check.checked;p.withdrawnRound=check.checked?targetRound:0;if(check.checked&&!waiting&&d.currentRound){const opp=opponent(p.no,d.currentRound);if(opp)delete d.drafts[key(p.no,opp)];}showInputErrors([]);save();render()}};w.className='withdraw-cell';w.append(check);tr.append(w);
for(let round=1;round<=d.currentRound;round++){let opp=opponent(p.no,round);let pairNow=d.pairings[round]?.find(q=>q.p1===p.no||q.p2===p.no);let opponentCell=el('td',(opp||''));opponentCell.className='opponent-cell';if(pairNow?.special){opponentCell.classList.add('special-pairing');opponentCell.title=pairNow.specialReason}else if(pairNow?.reason)opponentCell.title=pairNow.reason;if(p.withdrawn&&round>=(p.withdrawnRound||d.currentRound||1))opponentCell.classList.add('withdraw-gray');tr.append(opponentCell);let cell=el('td');cell.className='result-cell';if(opp&&round===d.currentRound&&!d.completed&&!p.withdrawn&&!d.players.find(x=>x.no===opp)?.withdrawn&&!d.drafts[key(p.no,opp)])cell.classList.add('pending-result');if(p.withdrawn&&round>=(p.withdrawnRound||d.currentRound||1))cell.classList.add('withdraw-gray');if(round===d.currentRound&&!d.completed&&!d.results[round]){if(opp===0){cell.textContent='〇'}else if(opp){let sel=el('select');let k=key(p.no,opp);for(let [v,label] of [['',''],['W','〇'],['L','×']]){let opt=el('option',label);opt.value=v;sel.append(opt)}let chosen=d.drafts[k]||'';sel.value=chosen?(Number(chosen)===p.no?'W':'L'):'';sel.disabled=!!p.withdrawn||!!d.players.find(x=>x.no===opp)?.withdrawn;sel.setAttribute('aria-label',`${p.no}番 ${p.name} 第${round}回戦 勝敗`);sel.dataset.player=String(p.no);sel.dataset.opp=String(opp);sel.onchange=()=>{if(sel.value)d.drafts[k]=String(sel.value==='W'?p.no:opp);else delete d.drafts[k];save();showInputErrors([]);render();const next=document.querySelector(`select[data-player=\"${p.no}\"]`);if(next)next.focus();updateInputAssist()};sel.onkeydown=e=>{if(e.key==='o'||e.key==='O'||e.key==='〇'||e.key==='○'){e.preventDefault();sel.value='W';sel.onchange()}else if(e.key==='x'||e.key==='X'||e.key==='×'){e.preventDefault();sel.value='L';sel.onchange()}else if(e.key==='Delete'||e.key==='Backspace'){e.preventDefault();sel.value='';sel.onchange()}};cell.append(sel)}}else cell.textContent=resultFor(p.no,round);tr.append(cell)}for(let [i,val] of [x?.rank??'',st[p.no].wins,st[p.no].losses,x?.directText??'',x?.sc??0,x?.sb??0,x?.md??0,x?.note??''].entries()){let cell=el('td',String(val));cell.className=i===7?'note':i===3?'direct-cell':'stat-cell';tr.append(cell)}t.append(tr)}$('matches').append(t);$('sortRank').textContent=sortByRank?'NO順':'順位順';updateInputAssist();if(currentAux)showAux(currentAux)}
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
      'WindowsのChromeまたはEdgeで起動します。デモでは本番データと別の大会フォルダを用意してください。',
      '「大会準備」で「新しい大会を開始」を選び、先に大会フォルダを指定し、次に準備用Excelを選択してから「第1回戦の組合せを作成」を押します。',
      '「大会進行」でクラスのタブを選択し、対局の勝敗（〇・×）を入力します。一方の結果を入力すると相手側にも反映されます。',
      '全対局の入力後に「入力完了」を押すと、その回戦だけを確定します。次回戦からの棄権者を確認してから「第○回戦の組合せを作成」を押してください。',
      '「転記」で掲示用の相手番号、「回戦履歴」で過去の結果、「順位順」で現在の順位順表示、「入力完了取消」で直前の入力完了を取り消せます。',
      'ヘルプ内の「自動入力」は未入力の対局だけにランダムで結果を入れるデモ・テスト用です。実際の大会結果の入力には使用しないでください。'
    ]],
    ['2．対戦相手の決め方（現在のWeb版の実装）',[
      '前日に準備用Excelの名簿をランダムに並べ替えて大会NOだけを確定し、相手番号空欄の対戦表を印刷します。当日は欠席者を確認してから第1回戦の組合せを作成します。',
      '第2回戦以降は再戦を避け、残り回戦も成立可能な候補を先読み確認します。その中で最大勝数差を最小にし、さらに大きな勝数差の対局数が少ない組合せを選びます。',
      '対戦相手番号の「※」は、先読み調整または2勝差以上の理由付き組合せです。相手番号にマウスを合わせると理由を確認できます。',
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
  body.append(el('p','入力補助（PCでは F5・F6 キーでも操作できます）'));
  const actions=el('div');actions.className='help-actions';
  const fill=el('button','自動入力');fill.type='button';fill.disabled=!d.currentRound||d.completed;fill.onclick=()=>{if(!confirm('未入力の対局だけをランダムに〇×入力しますか？'))return;autoFillCurrentRound();showAux('help')};
  const check=el('button','入力チェック');check.type='button';check.disabled=!d.currentRound||d.completed;check.onclick=()=>showCheck();
  actions.append(fill,check);body.append(actions);
  const note=el('p','自動入力は未入力の対局のみが対象です。入力済みの勝敗や棄権の設定は変更しません。入力チェックは未入力や矛盾を確認します。大会終了後は、必要に応じて入力完了を取り消してから操作してください。');note.className='help-note';body.append(note);

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
function showInputErrors(errors){
 const box=$('inputError');if(!box)return;
 box.hidden=!errors.length;
 box.replaceChildren();
 if(!errors.length)return;
 const title=el('strong','入力エラー：'+errors.length+'件の問題があります');box.append(title);
 box.append(el('p','以下の対局を確認してください。すべて解消するまで次の回戦には進めません。'));
 const list=el('ul');for(const error of errors)list.append(el('li',error));box.append(list);
 box.scrollIntoView({behavior:'smooth',block:'center'});
}
function showCheck(){const errors=checkCurrentRound();showInputErrors(errors);alert(errors.length?'F6 入力チェック（'+errors.length+'件）\n\n'+errors.join('\n'):'全行をチェックしました。問題はありません。');say(errors.length?'F6：'+errors.length+'件の問題があります':'F6：問題はありません');}


$('undo').onclick=async()=>{
 if(!d||!d.currentRound||d.undoUsed)return say('取り消せる直前の入力完了がありません。');
 if(!yukFolder.handle)return say('大会フォルダを選択してから操作してください');
 const waiting=!!d.awaitingNextRound,finalRound=!!d.completed;
 const prev=(finalRound||waiting)?d.currentRound:(d.currentRound>=2?d.currentRound-1:0);
 if(!prev||!d.results[prev])return say('取り消せる直前回戦がありません');
 if(!confirm(`第${prev}回戦の入力完了を取り消しますか？`))return;
 try{
  await yukFolder.flush();
  const stamp=historyStamp();
  await yukFolder.write('大会進行データ_入力完了取消前_'+stamp+'.json',new Blob([JSON.stringify({classes:tournamentClasses,active:activeClassIndex},null,2)],{type:'application/json'}));
  const rs=d.results[prev];
  // 次回戦の組合せが作成済みなら、組合せ・入力途中の勝敗・組合せ理由をすべて破棄する。
  if(!finalRound&&!waiting&&d.currentRound>prev){delete d.pairings[d.currentRound];if(d.pairingMeta)delete d.pairingMeta[d.currentRound];}
  delete d.results[prev];d.currentRound=prev;d.completed=false;d.awaitingNextRound=false;d.undoUsed=true;d.drafts={};
  for(const r of rs)if(r.p2&&!r.bye)d.drafts[key(r.p1,r.p2)]=String(r.winner);
  save();await yukFolder.flush();render();
  const blob=await exportTournamentExcel(tournamentClasses,rankForClass);
  await yukFolder.write('対戦表_大会用.xlsx',blob);
  say(`第${prev}回戦の入力完了を取り消しました。勝敗を修正して、もう一度［入力完了］を押してください。`);
 }catch(err){say('入力完了取消でエラーが発生しました：'+err.message+'。大会フォルダのバックアップを確認してください。')}
};

document.addEventListener('keydown',e=>{if(e.key==='F5'||e.key==='F6'){if($('workspace').hidden)return;e.preventDefault();if(e.key==='F5')autoFillCurrentRound();else showCheck();}});
$('exportCurrentExcel').onclick=async()=>{
 const btn=$('exportCurrentExcel');btn.disabled=true;
 try{
  if(!d)throw Error('クラスが選択されていません');
  await downloadCurrentClassExcel(d,rankForClass);
  say(d.className+'クラスのExcelダウンロードを開始しました。');
 }catch(e){say('Excel出力失敗：'+e.message);alert('Excel出力失敗：'+e.message)}
 finally{btn.disabled=false}
};
$('transfer').onclick=()=>openTransferWindow();$('historyBtn').onclick=()=>openHistoryWindow();$('help').onclick=()=>showAux('help');$('closeAux').onclick=()=>{currentAux='';$('aux').hidden=true};$('sortRank').onclick=()=>{if(sortColumn==='順位'&&sortAscending){sortColumn='NO';sortAscending=true;sortByRank=false}else{sortColumn='順位';sortAscending=true;sortByRank=true}render()};$('navprep').onclick=()=>{$('prep').hidden=false;$('workspace').hidden=true;$('navprep').className='active';$('navprogress').className=''};$('navprogress').onclick=()=>{if(!d)return say('先に大会を作成してください');$('prep').hidden=true;$('workspace').hidden=false;$('navprep').className='';$('navprogress').className='active'};
function save(){if(d){
 tournamentClasses[activeClassIndex]=d;
 if(window.yukFolder?.handle)window.yukFolder.queueState({classes:tournamentClasses,active:activeClassIndex});
}}function validate(x){if(!x||!Array.isArray(x.players)||!x.players.length||typeof x.results!=='object'||typeof x.pairings!=='object'||!Number.isInteger(x.rounds))throw Error('大会データ形式が不正です')}
$('create').onclick=()=>{try{let names=$('names').value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean),rounds=Number($('rounds').value);if(names.length<2||names.length>50||new Set(names).size!==names.length||rounds<3||rounds>10||rounds>names.length-1)throw Error('人数2～50名・名前重複なし・回戦3～10かつ人数−1以下にしてください');if(!confirm('現在の大会データを新規大会で置き換えますか？'))return;tournamentClasses=[];activeClassIndex=0;d={name:$('tname').value,className:$('cname').value,rounds,players:names.map((name,i)=>({no:i+1,name,withdrawn:false,withdrawnRound:0})),pairings:{},pairingMeta:{},results:{},drafts:{},currentRound:0,awaitingNextRound:false,completed:false};save();render();$('navprogress').click();say('新規大会を作成しました')}catch(e){say(e.message)}};
$('start').onclick=async()=>{try{const round=d.currentRound?(d.awaitingNextRound?d.currentRound+1:d.currentRound):1;if(!round||d.completed)return;const label=`第${round}回戦`;/* 同じ回戦を再作成する場合は、その回戦の組合せと入力途中の勝敗を全て破棄して最初から組み直す。 */delete d.pairings[round];if(d.pairingMeta)delete d.pairingMeta[round];d.drafts={};d.pairings[round]=makePairings(round);d.currentRound=round;d.awaitingNextRound=false;save();await yukFolder.flush();try{await exportTournamentExcel(tournamentClasses,rankForClass)}catch(_){/* state remains saved; manual Excel output is still available */}render();say(`${label}の組合せを作成しました。全員の対戦相手を確認してから対局を開始してください。`)}catch(e){say(e.message)}};
let lastHistoryStamp='',historySerial=0;function historyStamp(){const now=new Date(),pad=n=>String(n).padStart(2,'0');const base=now.getFullYear()+pad(now.getMonth()+1)+pad(now.getDate())+pad(now.getHours())+pad(now.getMinutes())+pad(now.getSeconds());historySerial=base===lastHistoryStamp?historySerial+1:0;lastHistoryStamp=base;return base+(historySerial?'_'+String(historySerial).padStart(2,'0'):'')}
$('complete').onclick=async()=>{let beforeCommit;try{let errors=checkCurrentRound();if(errors.length){showInputErrors(errors);return;}showInputErrors([]);let round=d.currentRound,rs=[];for(let p of d.pairings[round]){if(!p.p2){rs.push({p1:p.p1,p2:0,winner:p.p1,bye:true});continue}let a=d.players.find(x=>x.no===p.p1),b=d.players.find(x=>x.no===p.p2),winner=Number(d.drafts[key(p.p1,p.p2)]);if(a.withdrawn&&b.withdrawn)throw Error(`${p.p1}・${p.p2}が両方棄権しています。試作版では未対応です`);if(a.withdrawn)winner=p.p2;if(b.withdrawn)winner=p.p1;if(!winner)throw Error(`${p.p1}－${p.p2}の勝敗が未入力です`);rs.push({p1:p.p1,p2:p.p2,winner,bye:false})}if(!confirm(`${round}回戦を確定しますか？`))return;beforeCommit=JSON.stringify(d);d.results[round]=rs;d.undoUsed=false;d.drafts={};if(round>=d.rounds){d.completed=true;d.awaitingNextRound=false}else{d.awaitingNextRound=true}// 次回戦の組合せは自動作成しない。棄権者確認後に［第○回戦を作成］で決定する。
 await yukFolder.queueState({classes:tournamentClasses,active:activeClassIndex});
 const blob=await exportTournamentExcel(tournamentClasses,rankForClass);
 await yukFolder.write('対戦表_大会用_'+historyStamp()+'.xlsx',blob);
 sortByRank=false;sortColumn='NO';sortAscending=true;render();say(round>=d.rounds?`${round}回戦を確定し、大会を終了しました`:`${round}回戦を確定しました。次回戦から棄権する参加者を確認してから第${round+1}回戦の組合せを作成してください。`)}catch(e){if(typeof beforeCommit!=='undefined'){d=JSON.parse(beforeCommit);tournamentClasses[activeClassIndex]=d;render();try{await yukFolder.queueState({classes:tournamentClasses,active:activeClassIndex})}catch(_){/* report original save error */}}say('回戦は確定していません：'+e.message)}};
function downloadBackup(){if(!d)return say('大会がありません');const blob=new Blob([JSON.stringify(d,null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=el('a');a.href=url;a.download='yukuhashi-'+new Date().toISOString().replace(/[:.]/g,'-')+'.json';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);say('JSONの保存を開始しました。ダウンロードフォルダーにファイルがあることを確認してください。');}


async function restoreFromFile(file){if(!file)return;try{let x=JSON.parse(await file.text());validate(x);if(!confirm('現在の大会データを置き換えますか？ 保存していないデータは失われます。'))return;localStorage.setItem('yukuhashi-pwa-before-restore',JSON.stringify(d));d=x;save();render();$('prep').hidden=true;$('workspace').hidden=false;say('復元しました')}catch(err){say('復元失敗：'+err.message)}}


// 起動時は必ず大会準備。旧ブラウザ保存データを自動読込しない。
// 大会進行は利用者が選んだ大会フォルダのファイルだけから復元する。
$('prep').hidden=false;$('workspace').hidden=true;
const showPrepMode=mode=>{ $('prepChoices').hidden=!!mode;$('newPrep').hidden=mode!=='new';$('resumePrep').hidden=mode!=='resume';};
$('newTournament').onclick=()=>{showPrepMode('new');$('newExcelRow').hidden=true;$('excelNotice').hidden=true;$('excelCreate').hidden=true;$('excelCreate').disabled=true;$('excelPath').value='';$('excelMock').value='';excelPreparation=null;yukFolder.disconnect();$('folderNotice').textContent='大会用データを保存するための新しいフォルダを作って選択してください。';};
$('resumeTournament').onclick=async()=>{showPrepMode('resume');$('resumeLast').disabled=true;try{const info=await yukFolder.getLastInfo();$('lastFolderNotice').textContent=info?'前回の大会フォルダ：'+info.name+'。再開する場合は下のボタンを押してください。':'前回の大会フォルダは記憶されていません。別のフォルダを選択してください。';$('resumeLast').disabled=!info}catch(e){$('lastFolderNotice').textContent='前回のフォルダを確認できません：'+e.message}};
document.querySelectorAll('.backPrep').forEach(b=>b.onclick=()=>{yukFolder.disconnect();showPrepMode(null)});
async function openSaved(data){if(!data)throw Error('このフォルダには保存済みの大会進行データがありません');data.classes.forEach(validate);tournamentClasses=data.classes;activeClassIndex=Math.min(data.active||0,data.classes.length-1);d=tournamentClasses[activeClassIndex];await yukFolder.remember();render();$('navprogress').click();say('大会フォルダから再開しました')}
$('resumeLast').onclick=async()=>{try{const data=await yukFolder.restoreLast();if(!data)throw Error('保存済みの大会進行データがありません');if(!confirm('前回の大会「'+data.classes[0].name+'」を再開しますか？'))return;await openSaved(data)}catch(e){$('lastFolderNotice').textContent='再開できません：'+e.message+'。必要なら別の大会フォルダを選択してください。'}};
$('resumeOther').onclick=async()=>{try{const data=await yukFolder.pick();if(!data)throw Error('保存済みの大会進行データがありません');if(!confirm('「'+data.classes[0].name+'」を再開しますか？'))return;await openSaved(data)}catch(e){$('lastFolderNotice').textContent='再開できません：'+e.message}};
$('folderMock').onclick=async()=>{try{await yukFolder.pick();await yukFolder.remember();$('folderNotice').textContent='大会フォルダ：'+yukFolder.handle.name+'（選択済み）';$('newExcelRow').hidden=false;$('excelNotice').hidden=false;say('準備用Excelを選択してください')}catch(e){if(e.name!=='AbortError')say('フォルダ選択：'+e.message)}};
if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});

// Windows版準備画面のUI検証: 未実装のExcel機能を成功したように表示しない。
$('excelMock').addEventListener('change', async e=>{
 const f=e.target.files?.[0];$('excelPath').value=f?.name||'';$('excelCreate').disabled=true;excelPreparation=null;
 if(!f)return;
 try{
  excelPreparation=await parsePreparationXlsx(f);
  excelPreparation.sourceBytes=new Uint8Array(await f.arrayBuffer());
  $('excelNotice').textContent='✓ チェックOK　大会名：'+excelPreparation.name+'　'+excelPreparation.classes.length+'クラス・参加者合計 '+excelPreparation.classes.reduce((n,c)=>n+c.players.length,0)+'名。大会進行表を作成できます。';
  $('excelCreate').disabled=false;$('excelCreate').hidden=false;
 }catch(err){$('excelNotice').textContent='準備用Excelに問題があります：'+err.message}
});
$('excelCreate').onclick=async()=>{
 if(!excelPreparation)return;
 if(!window.yukFolder?.handle)return say('先に大会フォルダを選択してください。');
 tournamentClasses=excelPreparation.classes.map(c=>{
  // V127: 初回に名簿そのものをシャッフルし、その順番で大会NOを確定する。
  const shuffled=[...c.players];
  for(let i=shuffled.length-1;i>0;i--){const limit=Math.floor(4294967296/(i+1))*(i+1);let n;do{const b=new Uint32Array(1);crypto.getRandomValues(b);n=b[0]}while(n>=limit);const j=n%(i+1);[shuffled[i],shuffled[j]]=[shuffled[j],shuffled[i]]}if(shuffled.length>2&&shuffled.every((v,i)=>v===c.players[i]))[shuffled[0],shuffled[1]]=[shuffled[1],shuffled[0]];
  const players=shuffled.map((p,i)=>({...p,no:i+1,withdrawn:false,withdrawnRound:0}));
  return {name:excelPreparation.name,className:c.name,rounds:c.rounds,players,pairings:{},pairingMeta:{},results:{},drafts:{},currentRound:0,awaitingNextRound:false,completed:false};
 });
 activeClassIndex=0;d=tournamentClasses[0];
 try{setExcelTemplate(excelPreparation.sourceBytes);await yukFolder.write('対戦表_準備用.xlsx',new Blob([excelPreparation.sourceBytes]));await exportTournamentExcel(tournamentClasses,rankForClass)}catch(err){return say('大会用Excelの作成に失敗しました：'+err.message)}
 save();await yukFolder.flush();render();$('navprogress').click();
 say('全'+tournamentClasses.length+'クラスの大会進行表を作成しました。当日は欠席者を確認してから第1回戦の組合せを作成してください。');
};
$('exitPrep').onclick=()=>{$('navprogress').click()};

// 表内入力補助：対戦ペア単位の未入力チェック。既存の対戦相手への自動反映を使用。
function updateInputAssist(){
 const box=document.getElementById('inputAssist');if(!box||!d)return;
 if(!d.currentRound||d.completed||d.awaitingNextRound){box.textContent=d.completed?'全回戦終了':d.awaitingNextRound?`第${d.currentRound+1}回戦の作成待ち（棄権者を確認してください）`:'第1回戦の作成待ち（当日欠席者を確認してください）';return}
 const pairs=d.pairings[d.currentRound]||[];let pending=[];let completed=0;
 for(const pair of pairs){if(!pair.p2)continue;const a=d.players.find(p=>p.no===pair.p1),b=d.players.find(p=>p.no===pair.p2);if(a?.withdrawn||b?.withdrawn||d.drafts[key(pair.p1,pair.p2)])completed++;else pending.push(pair.p1+'－'+pair.p2)}
 box.textContent=`入力済み ${completed}/${completed+pending.length}対局　未入力 ${pending.length}対局`+(pending.length?'（番号：'+pending.join('、')+'）':'　入力完了できます');
}

function rankForClass(cls){const old=d;try{d=cls;return rank()}finally{d=old}}


// V17: vertically resize the match table by dragging the divider below it.
(function(){
  function setupMatchResizer(){
    const box=document.getElementById('matches');
    const bar=document.getElementById('matches-resizer');
    if(!box||!bar||bar.dataset.ready==='1') return;
    bar.dataset.ready='1';
    let startY=0,startH=0;
    const move=(e)=>{
      const h=Math.max(180,Math.min(window.innerHeight-180,startH+(e.clientY-startY)));
      box.style.height=h+'px';
      box.style.maxHeight='none';
    };
    const up=()=>{
      document.removeEventListener('pointermove',move);
      document.removeEventListener('pointerup',up);
      document.body.classList.remove('resizing-matches');
    };
    bar.addEventListener('pointerdown',(e)=>{
      startY=e.clientY; startH=box.getBoundingClientRect().height;
      document.body.classList.add('resizing-matches');
      document.addEventListener('pointermove',move);
      document.addEventListener('pointerup',up);
      e.preventDefault();
    });
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',setupMatchResizer);
  else setupMatchResizer();
})();
