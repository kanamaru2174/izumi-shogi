'use strict';
// Local-only XLSX preparation reader; no network requests.
async function parsePreparationXlsx(file){
 const buf=await file.arrayBuffer(),view=new DataView(buf),bytes=new Uint8Array(buf),decoder=new TextDecoder();
 let end=-1;for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054b50){end=i;break}
 if(end<0)throw Error('xlsx形式のZIP構造を読み取れません');
 const files=new Map();let at=view.getUint32(end+16,true),count=view.getUint16(end+10,true);
 for(let n=0;n<count;n++){
  if(view.getUint32(at,true)!==0x02014b50)throw Error('xlsxのファイル一覧が壊れています');
  const method=view.getUint16(at+10,true),size=view.getUint32(at+20,true),nameLen=view.getUint16(at+28,true),extra=view.getUint16(at+30,true),comment=view.getUint16(at+32,true),offset=view.getUint32(at+42,true);
  const name=decoder.decode(bytes.slice(at+46,at+46+nameLen));files.set(name,{method,size,offset});at+=46+nameLen+extra+comment;
 }
 async function read(name){let f=files.get(name);if(!f)throw Error('Excel内に'+name+'がありません');let off=f.offset;if(view.getUint32(off,true)!==0x04034b50)throw Error('xlsxの内容が壊れています');let begin=off+30+view.getUint16(off+26,true)+view.getUint16(off+28,true),data=bytes.slice(begin,begin+f.size);
  if(f.method===8){if(typeof DecompressionStream==='undefined')throw Error('このブラウザはExcel展開に対応していません');return decoder.decode(await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())}
  if(f.method!==0)throw Error('未対応のExcel圧縮方式です');return decoder.decode(data);
 }
 const xml=s=>{let x=new DOMParser().parseFromString(s,'application/xml');if(x.querySelector('parsererror'))throw Error('ExcelのXMLを解析できません');return x};
 const shared=[];if(files.has('xl/sharedStrings.xml')){let doc=xml(await read('xl/sharedStrings.xml'));for(let si of doc.getElementsByTagName('si'))shared.push([...si.getElementsByTagName('t')].map(t=>t.textContent).join(''))}
 const workbook=xml(await read('xl/workbook.xml')),rels=xml(await read('xl/_rels/workbook.xml.rels'));
 const relation={};for(let r of rels.getElementsByTagName('Relationship'))relation[r.getAttribute('Id')]=r.getAttribute('Target');
 const sheets=[];
 for(let sheet of workbook.getElementsByTagName('sheet')){
  let id=sheet.getAttribute('r:id')||sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');let target=relation[id];if(!target)throw Error('シートの参照先がありません');let path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');path=path.replace(/^xl\/xl\//,'xl/');
  let doc=xml(await read(path)),cells={};for(let c of doc.getElementsByTagName('c')){let ref=c.getAttribute('r'),v=c.getElementsByTagName('v')[0]?.textContent||'';if(c.getAttribute('t')==='s')v=shared[Number(v)]||'';else if(c.getAttribute('t')==='inlineStr')v=[...c.getElementsByTagName('t')].map(t=>t.textContent).join('');cells[ref]=v}
  sheets.push({name:sheet.getAttribute('name'),cells});
 }
 if(sheets.length<2)throw Error('大会名シートとクラス別シートが必要です');
 const first=sheets.find(s=>s.name==='大会名');if(!first)throw Error('「大会名」シートがありません');
 const name=String(first.cells.C2||'').trim();if(!name)throw Error('大会名シート C2 の大会名が空欄です');
 const classes=[];for(let sh of sheets.filter(s=>s!==first)){
  const className=String(sh.cells.B3||'').trim(),rounds=Number(sh.cells.D3);if(!className)throw Error(sh.name+'：B3のクラス名が空欄です');if(!Number.isInteger(rounds)||rounds<3||rounds>10)throw Error(sh.name+'：D3の回戦数は3～10にしてください');
  let players=[];for(let row=7;row<=200;row++){let no=sh.cells['B'+row],pname=String(sh.cells['C'+row]||'').trim();if(no===''||no===undefined){if(pname)throw Error(sh.name+'：'+row+'行の番号がありません');continue}if(Number(no)!==players.length+1||!pname)throw Error(sh.name+'：'+row+'行の番号・氏名を確認してください');players.push({no:Number(no),name:pname})}
  if(players.length<2||players.length>50||rounds>players.length-1)throw Error(sh.name+'：人数は2～50名、回戦数は人数−1以下にしてください');
  if(new Set(players.map(p=>p.name)).size!==players.length)throw Error(sh.name+'：参加者名が重複しています');classes.push({name:className,sheetName:sh.name,rounds,players});
 }
 if(!classes.length||classes.length>10)throw Error('クラス数は1～10にしてください');if(new Set(classes.map(c=>c.name)).size!==classes.length)throw Error('クラス名が重複しています');
 return {name,classes};
}
