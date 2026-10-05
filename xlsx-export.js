// Windows V127 template-preserving XLSX writer. All operations local/offline.
'use strict';
const XLSX_STORE='yukuhashi-excel-template-v1';
let excelTemplateBytes=null;
function toB64(bytes){let parts=[];for(let i=0;i<bytes.length;i+=8192)parts.push(String.fromCharCode(...bytes.subarray(i,i+8192)));return btoa(parts.join(''))}
function fromB64(s){let raw=atob(s),b=new Uint8Array(raw.length);for(let i=0;i<raw.length;i++)b[i]=raw.charCodeAt(i);return b}
function setExcelTemplate(bytes){excelTemplateBytes=new Uint8Array(bytes)}
function getExcelTemplate(){if(!excelTemplateBytes)throw Error('大会フォルダの準備用Excelが読み込まれていません。大会フォルダを選び直してください');return excelTemplateBytes}
async function unzipEntries(bytes){let v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),end=-1;for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(v.getUint32(i,true)===0x06054b50){end=i;break}if(end<0)throw Error('ExcelファイルのZIP形式が不正です');let count=v.getUint16(end+10,true),at=v.getUint32(end+16,true),out=new Map(),dec=new TextDecoder();for(let i=0;i<count;i++){if(v.getUint32(at,true)!==0x02014b50)throw Error('Excelのファイル一覧が不正です');let method=v.getUint16(at+10,true),size=v.getUint32(at+20,true),nl=v.getUint16(at+28,true),extra=v.getUint16(at+30,true),comment=v.getUint16(at+32,true),offset=v.getUint32(at+42,true),name=dec.decode(bytes.subarray(at+46,at+46+nl));if(v.getUint32(offset,true)!==0x04034b50)throw Error('Excelファイルの内容が不正です');let start=offset+30+v.getUint16(offset+26,true)+v.getUint16(offset+28,true),raw=bytes.slice(start,start+size),data;if(method===0)data=raw;else if(method===8){if(!globalThis.DecompressionStream)throw Error('このブラウザはExcel展開に対応していません');data=new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())}else throw Error('未対応のExcel圧縮方式です');out.set(name,data);at+=46+nl+extra+comment}return out}
let crcTable=null;function crc32(bytes){if(!crcTable){crcTable=Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0})}let c=0xffffffff;for(let b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0}
function zipStored(entries){let enc=new TextEncoder(),parts=[],central=[],offset=0;for(let [name,data] of entries){let n=enc.encode(name),crc=crc32(data),local=new Uint8Array(30+n.length),lv=new DataView(local.buffer);lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);lv.setUint16(8,0,true);lv.setUint32(14,crc,true);lv.setUint32(18,data.length,true);lv.setUint32(22,data.length,true);lv.setUint16(26,n.length,true);local.set(n,30);parts.push(local,data);let c=new Uint8Array(46+n.length),cv=new DataView(c.buffer);cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint32(16,crc,true);cv.setUint32(20,data.length,true);cv.setUint32(24,data.length,true);cv.setUint16(28,n.length,true);cv.setUint32(42,offset,true);c.set(n,46);central.push(c);offset+=local.length+data.length}let centralSize=central.reduce((s,b)=>s+b.length,0),end=new Uint8Array(22),ev=new DataView(end.buffer);ev.setUint32(0,0x06054b50,true);ev.setUint16(8,entries.size,true);ev.setUint16(10,entries.size,true);ev.setUint32(12,centralSize,true);ev.setUint32(16,offset,true);return new Blob([...parts,...central,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
function xmlParse(bytes){let doc=new DOMParser().parseFromString(new TextDecoder().decode(bytes),'application/xml');if(doc.querySelector('parsererror'))throw Error('Excel XMLの解析に失敗しました');return doc}
function colName(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function putCell(doc,row,num,value){let ns=doc.documentElement.namespaceURI,ref=colName(num)+row.getAttribute('r'),c=[...row.children].find(x=>x.localName==='c'&&x.getAttribute('r')===ref);if(value===null||value===undefined||value===''){if(c){c.removeAttribute('t');for(let x of [...c.children])if(x.localName==='v'||x.localName==='is'||x.localName==='f')x.remove()}return}if(!c){c=doc.createElementNS(ns,'c');c.setAttribute('r',ref);let after=[...row.children].find(x=>x.localName==='c'&&cellColumn(x.getAttribute('r'))>num);row.insertBefore(c,after||null)}for(let x of [...c.children])if(x.localName==='v'||x.localName==='is'||x.localName==='f')x.remove();if(typeof value==='number'){c.removeAttribute('t');let v=doc.createElementNS(ns,'v');v.textContent=String(value);c.append(v)}else{c.setAttribute('t','inlineStr');let is=doc.createElementNS(ns,'is'),t=doc.createElementNS(ns,'t');t.textContent=String(value);is.append(t);c.append(is)}}
function cellColumn(ref){let n=0;for(let c of (ref||'').match(/^[A-Z]+/)?.[0]||'')n=n*26+c.charCodeAt(0)-64;return n}
function getRow(doc,num){let sheet=doc.getElementsByTagName('sheetData')[0],row=[...sheet.children].find(x=>x.localName==='row'&&Number(x.getAttribute('r'))===num);if(!row){row=doc.createElementNS(doc.documentElement.namespaceURI,'row');row.setAttribute('r',String(num));let after=[...sheet.children].find(x=>x.localName==='row'&&Number(x.getAttribute('r'))>num);sheet.insertBefore(row,after||null)}return row}
function sheetPaths(entries){let wb=xmlParse(entries.get('xl/workbook.xml')),rels=xmlParse(entries.get('xl/_rels/workbook.xml.rels')),targets=new Map();for(let r of rels.getElementsByTagName('Relationship'))targets.set(r.getAttribute('Id'),r.getAttribute('Target'));let paths=new Map();for(let s of wb.getElementsByTagName('sheet')){let id=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),t=targets.get(id);if(!t)continue;let path=t.startsWith('/')?t.slice(1):'xl/'+t.replace(/^\.\//,'');path=path.replace(/^xl\/xl\//,'xl/');paths.set(s.getAttribute('name'),path)}return paths}
// Normalize summary columns to the actual round count; the source template may
// have room for only five summary fields, whereas the web app writes eight.
// Approximate Excel's Home > Format > AutoFit Column Width at export time.
// XLSX bestFit alone does not reliably resize on open, so measure actual output.
function autoFitNoteWidth(notes){
 // Measure the *actual cell text*; rank callback notes may differ from exports.
 // Fullwidth characters occupy approximately two Excel character units.
 const units=text=>[...String(text??'')].reduce((n,ch)=>n+(/[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7ff\uf900-\ufaff\uff01-\uff60]/.test(ch)?2:1),0);
 return Math.min(255,Math.max(12,Math.ceil(Math.max(units('備考'),...notes.map(units))+3)));
}
function noteTextsFromSheet(doc,col){
 const result=[];
 for(const row of doc.getElementsByTagName('row')){
  if(Number(row.getAttribute('r'))<7)continue;
  const ref=colName(col)+row.getAttribute('r');
  const c=[...row.children].find(x=>x.localName==='c'&&x.getAttribute('r')===ref);
  if(c){const t=c.getElementsByTagName('t')[0]||c.getElementsByTagName('v')[0];if(t)result.push(t.textContent)}
 }
 return result;
}
function alignedStyle(styles,base,alignment){
 const ns=styles.documentElement.namespaceURI,cellXfs=styles.getElementsByTagName('cellXfs')[0];
 const xfs=[...cellXfs.children].filter(x=>x.localName==='xf');
 const xf=xfs[Number(base)]?.cloneNode(true)||xfs[0].cloneNode(true);
 let a=[...xf.children].find(x=>x.localName==='alignment');
 if(!a){a=styles.createElementNS(ns,'alignment');xf.append(a)}
 for(const [k,v] of Object.entries(alignment))a.setAttribute(k,v);
 xf.setAttribute('applyAlignment','1');cellXfs.append(xf);cellXfs.setAttribute('count',String(xfs.length+1));
 return String(xfs.length);
}

function specialPairStyle(styles,base){
 const ns=styles.documentElement.namespaceURI;
 const cellXfs=styles.getElementsByTagName('cellXfs')[0];
 const fonts=styles.getElementsByTagName('fonts')[0];
 const borders=styles.getElementsByTagName('borders')[0];
 const xfs=[...cellXfs.children].filter(x=>x.localName==='xf');
 const src=xfs[Number(base)]||xfs[0];
 // Preserve the original cell font (Yu Gothic, size etc.) and add only bold.
 const srcFont=fonts.children[Number(src.getAttribute('fontId')||0)]||fonts.children[0];
 const font=srcFont.cloneNode(true);
 if(![...font.children].some(x=>x.localName==='b'))font.insertBefore(styles.createElementNS(ns,'b'),font.firstChild);
 fonts.append(font);fonts.setAttribute('count',String(fonts.children.length));
 const fontId=fonts.children.length-1;
 // Four-sided double border.  This is deliberately independent from fills so
 // prize-ranking background colours can coexist with the special-pair marker.
 const border=styles.createElementNS(ns,'border');
 for(const side of ['left','right','top','bottom']){
  const e=styles.createElementNS(ns,side);e.setAttribute('style','double');
  const color=styles.createElementNS(ns,'color');color.setAttribute('auto','1');e.append(color);border.append(e);
 }
 border.append(styles.createElementNS(ns,'diagonal'));
 borders.append(border);borders.setAttribute('count',String(borders.children.length));
 const borderId=borders.children.length-1;
 const xf=src.cloneNode(true);
 xf.setAttribute('fontId',String(fontId));xf.setAttribute('borderId',String(borderId));
 xf.setAttribute('applyFont','1');xf.setAttribute('applyBorder','1');
 cellXfs.append(xf);cellXfs.setAttribute('count',String(cellXfs.children.length));
 return String(cellXfs.children.length-1);
}
function fixSummaryLayout(doc,stat,styles,noteWidth){
 const ns=doc.documentElement.namespaceURI,last=stat+7;
 const titles=['順位','勝数','負数','直接対決','SC','SB','MD','備考'];
 const existing=(row,n)=>[...row.children].find(c=>c.localName==='c'&&c.getAttribute('r')===colName(n)+row.getAttribute('r'));
 const ensure=(row,n)=>{let c=existing(row,n);if(!c){c=doc.createElementNS(ns,'c');c.setAttribute('r',colName(n)+row.getAttribute('r'));let after=[...row.children].find(x=>x.localName==='c'&&cellColumn(x.getAttribute('r'))>n);row.insertBefore(c,after||null)}return c};
 // Read actual reference styles BEFORE moving anything. Original workbook uses
 // thin internal lines, medium outside borders, and medium every fifth row.
 const h=getRow(doc,5),sub=getRow(doc,6),refStat=stat,refEnd=stat+5;
 const hMid=existing(h,refStat+1)?.getAttribute('s')||existing(h,refStat)?.getAttribute('s');
 const hFirst=existing(h,refStat)?.getAttribute('s')||hMid;
 const hEnd=existing(h,refEnd)?.getAttribute('s')||hMid;
 const subMid=existing(sub,refStat+1)?.getAttribute('s')||existing(sub,refStat)?.getAttribute('s');
 const subEnd=existing(sub,refEnd)?.getAttribute('s')||subMid;
 const noteHeaderStyle=alignedStyle(styles,hEnd||hMid||'0',{horizontal:'center',vertical:'center',textRotation:'0',wrapText:'0'});
 const noteBodyStyles=new Map();
 const noteBodyStyle=base=>{if(!noteBodyStyles.has(base))noteBodyStyles.set(base,alignedStyle(styles,base,{horizontal:'left',vertical:'center',textRotation:'0'}));return noteBodyStyles.get(base)};
 const samples=[];
 for(let r=7;r<=56;r++){
  let row=getRow(doc,r),mid=existing(row,refStat)?.getAttribute('s'),end=existing(row,refEnd)?.getAttribute('s');
  samples.push({row,mid:mid||'3',end:end||mid||'3'});
 }
 for(let i=0;i<8;i++){
  let c=ensure(h,stat+i);c.setAttribute('s',i===7?noteHeaderStyle:i===0?hFirst:hMid);putCell(doc,h,stat+i,titles[i]);
  let h2=ensure(sub,stat+i);h2.setAttribute('s',i===7?subEnd:subMid);putCell(doc,sub,stat+i,null);
 }
 for(const {row,mid,end} of samples)for(let i=0;i<8;i++)ensure(row,stat+i).setAttribute('s',i===7?noteBodyStyle(end):mid);
 const merges=doc.getElementsByTagName('mergeCells')[0];
 if(merges){
  for(let m of [...merges.children]){
   let ref=m.getAttribute('ref')||'',match=ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
   if(match&&Number(match[2])<=6&&Number(match[4])>=5&&cellColumn(match[3])>=stat)m.remove();
  }
  for(let i=0;i<8;i++){let m=doc.createElementNS(ns,'mergeCell');m.setAttribute('ref',colName(stat+i)+'5:'+colName(stat+i)+'6');merges.append(m)}
  merges.setAttribute('count',String(merges.children.length));
 }
 const cols=doc.getElementsByTagName('cols')[0];
 if(cols){
  for(let c of [...cols.children]){
   if(c.localName!=='col')continue;
   let a=Number(c.getAttribute('min')),b=Number(c.getAttribute('max'));
   if(b<stat)continue;
   if(a<stat)c.setAttribute('max',String(stat-1));else c.remove();
  }
  [5.4,5.4,5.4,6.5,5.4,5.4,5.4,noteWidth].forEach((w,i)=>{let c=doc.createElementNS(ns,'col');c.setAttribute('min',String(stat+i));c.setAttribute('max',String(stat+i));c.setAttribute('width',String(w));c.setAttribute('customWidth','1');if(i===7)c.setAttribute('bestFit','1');cols.append(c)});
 }
 // Never leave old summary values, formatting or merged ranges to the right.
 for(let row of doc.getElementsByTagName('row'))for(let c of [...row.children])if(c.localName==='c'&&cellColumn(c.getAttribute('r'))>last)c.remove();
 if(merges){for(let m of [...merges.children]){let ref=m.getAttribute('ref')||'',a=ref.match(/^([A-Z]+)\d+/);if(a&&cellColumn(a[1])>last)m.remove()}merges.setAttribute('count',String(merges.children.length))}
 let dim=doc.getElementsByTagName('dimension')[0];if(dim)dim.setAttribute('ref','A1:'+colName(last)+'56');
}
function sheetUpdate(entries,paths,cls,withRanking,rankCallback){
 let path=paths.get(cls.className);if(!path||!entries.has(path))throw Error(cls.className+'のExcelシートが見つかりません');
 let doc=xmlParse(entries.get(path)),stat=5+cls.rounds*2,rs=withRanking&&Object.values(cls.results).some(x=>x.length)?rankCallback(cls):null,byNo=new Map((rs||[]).map(x=>[x.no,x])),specialCells=[];
 for(let row=7;row<=56;row++){
  let p=cls.players[row-7],r=getRow(doc,row);for(let col=2;col<=stat+7;col++)putCell(doc,r,col,null);if(!p)continue;
  putCell(doc,r,2,p.no);putCell(doc,r,3,p.name);if(p.withdrawn)putCell(doc,r,4,1);
  for(let round=1;round<=cls.rounds;round++){
   let pair=(cls.pairings[round]||[]).find(x=>x.p1===p.no||x.p2===p.no),opp=pair?(pair.p1===p.no?pair.p2:pair.p1):null,oppCol=5+(round-1)*2;
   if(opp){putCell(doc,r,oppCol,opp);if(pair?.special)specialCells.push({row:row,col:oppCol})}
   let match=(cls.results[round]||[]).find(x=>x.p1===p.no||x.p2===p.no);
   if(match)putCell(doc,r,6+(round-1)*2,match.bye||match.winner===p.no?'〇':'×');else if(round===cls.currentRound&&cls.drafts&&opp){let k=[p.no,opp].sort((a,b)=>a-b).join(':'),winner=Number(cls.drafts[k]);if(winner)putCell(doc,r,6+(round-1)*2,winner===p.no?'〇':'×')}
  }
  let q=byNo.get(p.no);if(q){for(let [i,v] of [q.rank,q.wins,q.losses,q.directText,q.sc,q.sb,q.md,q.note].entries())putCell(doc,r,stat+i,v)}
 }
 // 現行表の下を2行空け、B列から特殊組合せ理由を1対局1行で記録する。
 const lastTableRow=6+cls.players.length,reasonStart=lastTableRow+3;
 // 参加者最終行より下にテンプレート由来の不要な罫線・書式を残さない。
 for(let rr=lastTableRow+1;rr<=56;rr++){let row=getRow(doc,rr);for(const c of [...row.children]){if(c.localName!=='c')continue;let m=(c.getAttribute('r')||'').match(/^([A-Z]+)(\d+)$/);if(!m)continue;let n=0;for(const ch of m[1])n=n*26+ch.charCodeAt(0)-64;if(n>=2&&n<=stat+7)c.remove()}}
 let reasonRow=reasonStart,reasonCells=[];
 for(let round=1;round<=cls.rounds;round++)for(const pair of (cls.pairings?.[round]||[]))if(pair.p2&&pair.special){let rr=reasonRow++,r=getRow(doc,rr);putCell(doc,r,2,`${round}回戦-NO_${pair.p1}とNO_${pair.p2}-${pair.specialReason}`);reasonCells.push({row:rr,col:2})}
 const styles=xmlParse(entries.get('xl/styles.xml'));
 // 表の下に追加した特殊組合せメッセージは、元セルの中央揃えを継承せず必ず左寄せにする。
 const reasonStyleCache=new Map();for(const x of reasonCells){let row=getRow(doc,x.row),ref=colName(x.col)+x.row,c=[...row.children].find(v=>v.localName==='c'&&v.getAttribute('r')===ref);if(!c)continue;let base=c.getAttribute('s')||'0';if(!reasonStyleCache.has(base))reasonStyleCache.set(base,alignedStyle(styles,base,{horizontal:'left',vertical:'center',textRotation:'0'}));c.setAttribute('s',reasonStyleCache.get(base))}
 // 特殊組合せの相手番号セルだけ、文字を太字＋四辺二重罫線にする。塗りつぶし色は変更しない。
 const styleCache=new Map();for(const x of specialCells){let row=getRow(doc,x.row),ref=colName(x.col)+x.row,c=[...row.children].find(v=>v.localName==='c'&&v.getAttribute('r')===ref);if(!c)continue;let base=c.getAttribute('s')||'0';if(!styleCache.has(base))styleCache.set(base,specialPairStyle(styles,base));c.setAttribute('s',styleCache.get(base))}
 const noteWidth=autoFitNoteWidth(noteTextsFromSheet(doc,stat+7));fixSummaryLayout(doc,stat,styles,noteWidth);
 // fixSummaryLayout may adjust styles elsewhere, but special opponent cells retain their dedicated style.
 let dim=doc.getElementsByTagName('dimension')[0];if(dim&&reasonRow>reasonStart)dim.setAttribute('ref','A1:'+colName(stat+7)+Math.max(56,reasonRow-1));
 entries.set('xl/styles.xml',new TextEncoder().encode(new XMLSerializer().serializeToString(styles)));entries.set(path,new TextEncoder().encode(new XMLSerializer().serializeToString(doc)))
}
// Keep only the selected class worksheet in a manually downloaded workbook.
function keepSelectedWorksheet(entries,selectedName){
 const wb=xmlParse(entries.get('xl/workbook.xml'));
 const rel=xmlParse(entries.get('xl/_rels/workbook.xml.rels'));
 const sheetList=wb.getElementsByTagName('sheets')[0];
 const all=[...sheetList.children].filter(x=>x.localName==='sheet');
 const selected=all.find(x=>x.getAttribute('name')===selectedName);
 if(!selected)throw Error('選択中のクラスのシートが見つかりません：'+selectedName);
 const relId=x=>x.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')||x.getAttribute('r:id');
 const keepId=relId(selected);
 const removedIds=new Set();
 for(const sh of all)if(sh!==selected){removedIds.add(relId(sh));sh.remove()}
 const relationships=[...rel.documentElement.children].filter(x=>x.localName==='Relationship');
 for(const item of relationships){
  if(!removedIds.has(item.getAttribute('Id')))continue;
  const target=item.getAttribute('Target')||'';
  const path=(target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'')).replace(/^xl\/xl\//,'xl/');
  entries.delete(path);
  const sub=path.replace(/^(.+\/)([^/]+)$/,'$1_rels/$2.rels');
  entries.delete(sub);
  item.remove();
 }
 // Workbook's active tab must be the only retained sheet.
 for(const view of wb.getElementsByTagName('workbookView'))view.setAttribute('activeTab','0');
 const defined=wb.getElementsByTagName('definedNames')[0];
 if(defined)for(const item of [...defined.children])if(item.hasAttribute('localSheetId'))item.remove();
 const enc=new TextEncoder(),serialize=x=>enc.encode(new XMLSerializer().serializeToString(x));
 entries.set('xl/workbook.xml',serialize(wb));
 entries.set('xl/_rels/workbook.xml.rels',serialize(rel));
}
// Generate a fresh one-sheet workbook. Never inherit phantom template borders.
function cleanClassWorkbook(cls,rankCallback){
 const enc=new TextEncoder(),entries=new Map(),add=(path,xml)=>entries.set(path,enc.encode(xml));
 const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
 const col=n=>colName(n),last=5+cls.rounds*2+7,stat=last-7;
 const rows=[],row=(r,cells,extra='')=>rows.push(`<row r="${r}"${extra}>${cells.join('')}</row>`);
 const cell=(r,c,v,style=1)=>{
  if(v===null||v===undefined||v==='')return `<c r="${col(c)}${r}" s="${style}"/>`;
  if(typeof v==='number')return `<c r="${col(c)}${r}" s="${style}"><v>${v}</v></c>`;
  return `<c r="${col(c)}${r}" s="${style}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
 };
 row(1,[]);
 row(2,[cell(2,2,'クラス名',2),cell(2,4,'対戦数',2)]);
 row(3,[cell(3,2,cls.className,1),cell(3,4,cls.rounds,1)]);
 row(4,[]);
 const h=[cell(5,2,'NO',2),cell(5,3,'名前',2),cell(5,4,'棄権有無',2)];
 for(let n=1;n<=cls.rounds;n++){h.push(cell(5,5+(n-1)*2,n+'回戦',2),cell(5,6+(n-1)*2,'',2))}
 ['順位','勝数','負数','直接対決','SC','SB','MD','備考'].forEach((v,i)=>h.push(cell(5,stat+i,v,i===7?2:5)));
 row(5,h,' ht="26" customHeight="1"');
 const sub=[cell(6,2,'',2),cell(6,3,'',2),cell(6,4,'',2)];
 for(let n=1;n<=cls.rounds;n++)sub.push(cell(6,5+(n-1)*2,'相手番号',5),cell(6,6+(n-1)*2,'勝敗',5));
 for(let i=0;i<8;i++)sub.push(cell(6,stat+i,'',2));
 row(6,sub,' ht="88" customHeight="1"');
 const ranked=Object.values(cls.results||{}).some(v=>v.length)?rankCallback(cls):[];
 const byNo=new Map((ranked||[]).map(v=>[v.no,v]));
 for(let i=0;i<cls.players.length;i++){
  const p=cls.players[i],r=i+7,cells=[cell(r,2,p.no),cell(r,3,p.name,3),cell(r,4,p.withdrawn?1:'')];
  for(let n=1;n<=cls.rounds;n++){
   const pair=(cls.pairings?.[n]||[]).find(v=>v.p1===p.no||v.p2===p.no);
   const opponent=pair?(pair.p1===p.no?pair.p2:pair.p1):null;
   const match=(cls.results?.[n]||[]).find(v=>v.p1===p.no||v.p2===p.no);
   cells.push(cell(r,5+(n-1)*2,opponent||''));
   const draftWinner=n===cls.currentRound&&opponent?Number(cls.drafts?.[[p.no,opponent].sort((a,b)=>a-b).join(':')]):0;
   cells.push(cell(r,6+(n-1)*2,match?(match.bye||match.winner===p.no?'〇':'×'):draftWinner?(draftWinner===p.no?'〇':'×'):''));
  }
  const q=byNo.get(p.no),values=q?[q.rank,q.wins,q.losses,q.directText,q.sc,q.sb,q.md,q.note]:Array(8).fill('');
  values.forEach((v,j)=>cells.push(cell(r,stat+j,v,j===7?3:1)));
  row(r,cells);
 }
 const widths=[{n:1,w:1.75},{n:2,w:5.25},{n:3,w:15},{n:4,w:4.875}];
 for(let i=5;i<stat;i++)widths.push({n:i,w:i%2===1?5.625:6.875});
 [4.625,4.625,4.625,4.625,4.625,4.625,4.625,autoFitNoteWidth((ranked||[]).map(x=>x.note))].forEach((w,i)=>widths.push({n:stat+i,w}));
 const cols=widths.map(({n,w})=>`<col min="${n}" max="${n}" width="${w}" customWidth="1"/>`).join('');
 const merges=['B2:C2','D2:G2','B3:C3','D3:G3','B5:B6','C5:C6','D5:D6'];
 for(let n=1;n<=cls.rounds;n++){let a=5+(n-1)*2;merges.push(`${col(a)}5:${col(a+1)}5`)}
 for(let i=0;i<8;i++)merges.push(`${col(stat+i)}5:${col(stat+i)}6`);
 const mergeXml=merges.map(ref=>`<mergeCell ref="${ref}"/>`).join('');
 add('xl/worksheets/sheet1.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${col(last)}${Math.max(7,cls.players.length+6)}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>${cols}</cols><sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${mergeXml}</mergeCells><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>`);
 add('xl/styles.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Yu Gothic"/></font><font><b/><sz val="11"/><name val="Yu Gothic"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center" textRotation="255" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
 add('xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${esc(cls.className)}" sheetId="1" r:id="rId1"/></sheets></workbook>`);
 add('xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
 add('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
 add('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
 return zipStored(entries);
}
// Export the currently selected class by editing the supplied ORIGINAL workbook.
// No cell styles, widths, row heights or merged ranges are recreated.
async function downloadCurrentClassExcel(cls,rankCallback,historyLabel=""){
 if(typeof EXACT_EXCEL_FORMAT_B64!=='string')throw Error('Excel原本が読み込まれていません');
 const entries=await unzipEntries(fromB64(EXACT_EXCEL_FORMAT_B64));
 const paths=sheetPaths(entries);
 if(!paths.has(cls.className))throw Error('原本に「'+cls.className+'」シートがありません');
 sheetUpdate(entries,paths,cls,true,rankCallback);
 const workbook=xmlParse(entries.get('xl/workbook.xml'));
 const sheets=workbook.getElementsByTagName('sheets')[0];
 const keep=[...sheets.children].find(x=>x.getAttribute('name')===cls.className);
 if(!keep)throw Error('選択クラスのシートが見つかりません');
 const nsRel='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const keepRel=keep.getAttributeNS(nsRel,'id')||keep.getAttribute('r:id');
 for(const sh of [...sheets.children])if(sh!==keep)sh.remove();
 for(const view of workbook.getElementsByTagName('workbookView'))view.setAttribute('activeTab','0');
 // Defined names referring to other sheets would otherwise trigger Excel repair warnings.
 for(const defs of [...workbook.getElementsByTagName('definedNames')])defs.remove();
 const rels=xmlParse(entries.get('xl/_rels/workbook.xml.rels'));
 const removed=[];
 for(const rel of [...rels.documentElement.children]){
  if(rel.localName!=='Relationship'||!rel.getAttribute('Type')?.endsWith('/worksheet'))continue;
  if(rel.getAttribute('Id')===keepRel)continue;
  let target=rel.getAttribute('Target');
  let path=(target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'' )).replace(/^xl\/xl\//,'xl/');
  removed.push('/'+path);entries.delete(path);rel.remove();
 }
 const ct=xmlParse(entries.get('[Content_Types].xml'));
 for(const node of [...ct.documentElement.children])if(removed.includes(node.getAttribute('PartName')))node.remove();
 const enc=new TextEncoder(),ser=doc=>enc.encode(new XMLSerializer().serializeToString(doc));
 entries.set('xl/workbook.xml',ser(workbook));
 entries.set('xl/_rels/workbook.xml.rels',ser(rels));
 entries.set('[Content_Types].xml',ser(ct));
 const blob=zipStored(entries);
 const now=new Date(),pad=n=>String(n).padStart(2,'0');const stamp=now.getFullYear()+pad(now.getMonth()+1)+pad(now.getDate())+pad(now.getHours())+pad(now.getMinutes())+pad(now.getSeconds());
 const safeLabel=historyLabel?'_'+String(historyLabel).replace(/[\\/:*?"<>|]/g,'_'):'';const fileName='対戦表_'+String(cls.className).replace(/[\\/:*?"<>|]/g,'_')+safeLabel+'_'+stamp+'.xlsx';
 if(!window.yukFolder?.handle)throw Error('大会フォルダが選択されていません');
 await window.yukFolder.write(fileName,blob);
 return {blob,fileName};
}
async function buildTournamentExcel(classes,rankCallback){let entries=await unzipEntries(getExcelTemplate()),paths=sheetPaths(entries);for(let cls of classes)sheetUpdate(entries,paths,cls,true,rankCallback);return zipStored(entries)}
async function exportTournamentExcel(classes,rankCallback){let blob=await buildTournamentExcel(classes,rankCallback);if(window.yukFolder?.handle){await window.yukFolder.write('対戦表_大会用.xlsx',blob);return blob}let url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='対戦表_大会用.xlsx';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);return blob}
