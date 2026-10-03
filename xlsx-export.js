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
function fixSummaryLayout(doc,stat,styles){
 const ns=doc.documentElement.namespaceURI,head=getRow(doc,5),last=stat+7;
 const titles=['順位','勝数','負数','直接対決','SC','SB','MD','備考'];
 const row6=getRow(doc,6),data=getRow(doc,7);
 // Capture template styles before changing the cells. The template's final
 // summary column has a right border; intermediate columns must not inherit it.
 const sample=(r,n)=>[...r.children].find(c=>c.localName==='c'&&c.getAttribute('r')===colName(n)+r.getAttribute('r'));
 const styleAt=(row,n)=>sample(row,n)?.getAttribute('s');
 const midHead=styleAt(head,stat+1)||styleAt(head,stat),endHead=styleAt(head,stat+5)||midHead;
 const midSub=styleAt(row6,stat+1)||styleAt(row6,stat),endSub=styleAt(row6,stat+5)||midSub;
 const midBody=styleAt(data,stat)||'3',endBody=styleAt(data,stat+5)||midBody;
 // Create dedicated, uniformly bordered styles rather than reusing the
 // template's last-column styles (which have internal right-edge borders).
 const borderList=styles.getElementsByTagName('borders')[0];
 const b=styles.createElementNS(styles.documentElement.namespaceURI,'border');
 for(const side of ['left','right','top','bottom']){
  const el=styles.createElementNS(styles.documentElement.namespaceURI,side);
  el.setAttribute('style','thin');
  const color=styles.createElementNS(styles.documentElement.namespaceURI,'color');color.setAttribute('indexed','64');el.append(color);b.append(el);
 }
 b.append(styles.createElementNS(styles.documentElement.namespaceURI,'diagonal'));
 const borderId=borderList.children.length;borderList.append(b);borderList.setAttribute('count',String(borderList.children.length));
 const xfs=styles.getElementsByTagName('cellXfs')[0],borderStyle=new Map();
 function withBorder(id){id=String(id||'0');if(borderStyle.has(id))return borderStyle.get(id);
  const original=xfs.children[Number(id)]||xfs.children[0];const xf=original.cloneNode(true);
  xf.setAttribute('borderId',String(borderId));xf.setAttribute('applyBorder','1');
  const next=String(xfs.children.length);xfs.append(xf);xfs.setAttribute('count',String(xfs.children.length));borderStyle.set(id,next);return next;
 }
 const midHeadBorder=withBorder(midHead),endHeadBorder=withBorder(endHead),midSubBorder=withBorder(midSub),endSubBorder=withBorder(endSub),midBodyBorder=withBorder(midBody),endBodyBorder=withBorder(endBody);
 for(let i=0;i<8;i++){
  const n=stat+i, ref=colName(n);
  let c=sample(head,n);if(!c){c=doc.createElementNS(ns,'c');c.setAttribute('r',ref+'5');head.append(c)}
  c.setAttribute('s',i===7?endHeadBorder:midHeadBorder);
  putCell(doc,head,n,titles[i]);
  let h=sample(row6,n);if(!h){h=doc.createElementNS(ns,'c');h.setAttribute('r',ref+'6');row6.append(h)}h.setAttribute('s',i===7?endSubBorder:midSubBorder);
  for(let r=7;r<=56;r++){
   const row=getRow(doc,r);let cell=sample(row,n);
   if(!cell){cell=doc.createElementNS(ns,'c');cell.setAttribute('r',ref+r);row.append(cell)}
   // Reuse the template's existing formatted cells, not its unformatted spillover.
   cell.setAttribute('s',i===7?endBodyBorder:midBodyBorder);
  }
 }
 const cols=doc.getElementsByTagName('cols')[0]||doc.createElementNS(ns,'cols');
 if(!cols.parentNode){let sheet=doc.getElementsByTagName('sheetData')[0];doc.documentElement.insertBefore(cols,sheet)}
 // Existing template often defines one combined range over the summary cells.
 // Remove overlapping definitions so individual widths reliably take effect.
 for(const c of [...cols.children])if(c.localName==='col'&&Number(c.getAttribute('max'))>=stat&&Number(c.getAttribute('min'))<=last){
  let a=Number(c.getAttribute('min')),b=Number(c.getAttribute('max'));
  if(a<stat){let left=c.cloneNode(true);left.setAttribute('max',String(stat-1));cols.insertBefore(left,c)}
  if(b>last){let right=c.cloneNode(true);right.setAttribute('min',String(last+1));cols.insertBefore(right,c)}c.remove();
 }
 [7,7,7,13,9,9,9,64].forEach((w,i)=>{let c=doc.createElementNS(ns,'col');c.setAttribute('min',String(stat+i));c.setAttribute('max',String(stat+i));c.setAttribute('width',String(w));c.setAttribute('customWidth','1');cols.append(c)});
 const merges=doc.getElementsByTagName('mergeCells')[0];if(merges){
  for(const m of [...merges.children]){const range=m.getAttribute('ref')||'',match=range.match(/^([A-Z]+)5:([A-Z]+)6$/);if(match&&cellColumn(match[1])<=last&&cellColumn(match[2])>=stat)m.remove()}
  for(let i=0;i<8;i++){let m=doc.createElementNS(ns,'mergeCell');m.setAttribute('ref',colName(stat+i)+'5:'+colName(stat+i)+'6');merges.append(m)}
  merges.setAttribute('count',String(merges.children.length));
 }
 // The supplied template retains formatted phantom cells after the final
 // summary column. Remove them; otherwise Excel displays stray vertical rules.
 for(const row of doc.getElementsByTagName('row')){
  for(const cell of [...row.children])
   if(cell.localName==='c'&&cellColumn(cell.getAttribute('r'))>last)cell.remove();
 }
 for(const col of [...cols.children]){
  if(col.localName!=='col')continue;
  const min=Number(col.getAttribute('min')),max=Number(col.getAttribute('max'));
  if(min>last)col.remove();
  else if(max>last)col.setAttribute('max',String(last));
 }
 if(merges)for(const m of [...merges.children]){
  const range=m.getAttribute('ref')||'',colsInRange=range.match(/([A-Z]+)[0-9]+:([A-Z]+)[0-9]+/);
  if(colsInRange&&cellColumn(colsInRange[1])>last)m.remove();
 }
 if(merges)merges.setAttribute('count',String(merges.children.length));
 const selection=doc.getElementsByTagName('selection')[0];
 if(selection){selection.setAttribute('activeCell','B7');selection.setAttribute('sqref','B7')}
 const dim=doc.getElementsByTagName('dimension')[0];if(dim){const old=dim.getAttribute('ref')||'A1';const end=old.match(/([A-Z]+)([0-9]+)$/);if(end&&cellColumn(end[1])<last)dim.setAttribute('ref','A1:'+colName(last)+end[2])}
}
function sheetUpdate(entries,paths,cls,withRanking,rankCallback){let path=paths.get(cls.className);if(!path||!entries.has(path))throw Error(cls.className+'のExcelシートが見つかりません');let doc=xmlParse(entries.get(path)),stat=5+cls.rounds*2,rs=withRanking&&Object.values(cls.results).some(x=>x.length)?rankCallback(cls):null,byNo=new Map((rs||[]).map(x=>[x.no,x]));for(let row=7;row<=56;row++){let p=cls.players[row-7],r=getRow(doc,row);for(let col=2;col<=stat+7;col++)putCell(doc,r,col,null);if(!p)continue;putCell(doc,r,2,p.no);putCell(doc,r,3,p.name);if(p.withdrawn)putCell(doc,r,4,1);for(let round=1;round<=cls.rounds;round++){let pair=(cls.pairings[round]||[]).find(x=>x.p1===p.no||x.p2===p.no);if(pair){let opp=pair.p1===p.no?pair.p2:pair.p1;if(opp)putCell(doc,r,5+(round-1)*2,opp)}let match=(cls.results[round]||[]).find(x=>x.p1===p.no||x.p2===p.no);if(match)putCell(doc,r,6+(round-1)*2,match.bye||match.winner===p.no?'〇':'×')}let q=byNo.get(p.no);if(q){for(let [i,v] of [q.rank,q.wins,q.losses,q.directText,q.sc,q.sb,q.md,q.note].entries())putCell(doc,r,stat+i,v)}}const styles=xmlParse(entries.get('xl/styles.xml'));fixSummaryLayout(doc,stat,styles);entries.set('xl/styles.xml',new TextEncoder().encode(new XMLSerializer().serializeToString(styles)));entries.set(path,new TextEncoder().encode(new XMLSerializer().serializeToString(doc)))}
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
 row(2,[cell(2,2,cls.name||'',2)]);
 row(3,[cell(3,2,cls.className+' クラス',2)]);
 row(4,[]);
 const h=[cell(5,2,'NO',2),cell(5,3,'名前',2),cell(5,4,'棄権',2)];
 for(let n=1;n<=cls.rounds;n++){h.push(cell(5,5+(n-1)*2,n+'回戦',2),cell(5,6+(n-1)*2,'',2))}
 ['順位','勝数','負数','直接対決','SC','SB','MD','備考'].forEach((v,i)=>h.push(cell(5,stat+i,v,2)));
 row(5,h,' ht="42" customHeight="1"');
 const sub=[cell(6,2,'',2),cell(6,3,'',2),cell(6,4,'',2)];
 for(let n=1;n<=cls.rounds;n++)sub.push(cell(6,5+(n-1)*2,'相手番号',2),cell(6,6+(n-1)*2,'勝敗',2));
 for(let i=0;i<8;i++)sub.push(cell(6,stat+i,'',2));
 row(6,sub,' ht="31" customHeight="1"');
 const ranked=Object.values(cls.results||{}).some(v=>v.length)?rankCallback(cls):[];
 const byNo=new Map((ranked||[]).map(v=>[v.no,v]));
 for(let i=0;i<cls.players.length;i++){
  const p=cls.players[i],r=i+7,cells=[cell(r,2,p.no),cell(r,3,p.name,3),cell(r,4,p.withdrawn?1:'')];
  for(let n=1;n<=cls.rounds;n++){
   const pair=(cls.pairings?.[n]||[]).find(v=>v.p1===p.no||v.p2===p.no);
   const opponent=pair?(pair.p1===p.no?pair.p2:pair.p1):null;
   const match=(cls.results?.[n]||[]).find(v=>v.p1===p.no||v.p2===p.no);
   cells.push(cell(r,5+(n-1)*2,opponent||''));
   cells.push(cell(r,6+(n-1)*2,match?(match.bye||match.winner===p.no?'〇':'×'):''));
  }
  const q=byNo.get(p.no),values=q?[q.rank,q.wins,q.losses,q.directText,q.sc,q.sb,q.md,q.note]:Array(8).fill('');
  values.forEach((v,j)=>cells.push(cell(r,stat+j,v,j===7?3:1)));
  row(r,cells);
 }
 const widths=[{n:1,w:3},{n:2,w:7},{n:3,w:21},{n:4,w:7}];
 for(let i=5;i<stat;i++)widths.push({n:i,w:i%2===1?12:8});
 [8,8,8,14,9,9,9,70].forEach((w,i)=>widths.push({n:stat+i,w}));
 const cols=widths.map(({n,w})=>`<col min="${n}" max="${n}" width="${w}" customWidth="1"/>`).join('');
 const merges=['B2:'+col(Math.min(last,8))+'2','B3:'+col(Math.min(last,8))+'3','B5:B6','C5:C6','D5:D6'];
 for(let n=1;n<=cls.rounds;n++){let a=5+(n-1)*2;merges.push(`${col(a)}5:${col(a+1)}5`)}
 for(let i=0;i<8;i++)merges.push(`${col(stat+i)}5:${col(stat+i)}6`);
 const mergeXml=merges.map(ref=>`<mergeCell ref="${ref}"/>`).join('');
 add('xl/worksheets/sheet1.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${col(last)}${Math.max(7,cls.players.length+6)}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>${cols}</cols><sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${mergeXml}</mergeCells><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>`);
 add('xl/styles.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Yu Gothic"/></font><font><b/><sz val="11"/><name val="Yu Gothic"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
 add('xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${esc(cls.className)}" sheetId="1" r:id="rId1"/></sheets></workbook>`);
 add('xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
 add('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
 add('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
 return zipStored(entries);
}
async function downloadCurrentClassExcel(cls,rankCallback){
 const blob=cleanClassWorkbook(cls,rankCallback),url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download='対戦表_'+String(cls.className).replace(/[\\/:*?"<>|]/g,'_')+'.xlsx';
 document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
 return blob;
}
async function buildTournamentExcel(classes,rankCallback){let entries=await unzipEntries(getExcelTemplate()),paths=sheetPaths(entries);for(let cls of classes)sheetUpdate(entries,paths,cls,true,rankCallback);return zipStored(entries)}
async function exportTournamentExcel(classes,rankCallback){let blob=await buildTournamentExcel(classes,rankCallback);if(window.yukFolder?.handle){await window.yukFolder.write('対戦表_大会用.xlsx',blob);return blob}let url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='対戦表_大会用.xlsx';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);return blob}
