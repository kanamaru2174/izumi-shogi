"use strict";
(()=>{
 const bar=document.createElement('div');bar.className='ipad-tools';bar.innerHTML='<button id="ipadBackup" type="button">大会バックアップ</button><button id="ipadPrepDownload" type="button">準備用Excel保存</button><label class="filebtn">バックアップ読込<input id="ipadImport" type="file" accept=".json,application/json" hidden></label><button id="ipadExcel" type="button">Excel出力</button>';
 document.querySelector('#workspace').insertBefore(bar,document.querySelector('#workspace .status'));
 const title=document.getElementById('eventtitle');title.textContent='大会結果';
 const help=document.createElement('div');help.id='ipadTips';help.innerHTML='<strong>❔ 使い方のヒント</strong><p>対局結果を入力すると順位・勝敗数・各種指標を自動計算します。回戦履歴で結果を確認できます。大会バックアップでデータを保存してください。</p>';document.querySelector('#workspace .toolbar:last-of-type').after(help);
 const label=document.createElement('span');label.className='ipad-trophy';label.textContent='🏆';title.prepend(label);
 const info=document.createElement('p');info.className='ipad-info';info.textContent='大会はこのiPad内に自動保存されます。Safariの履歴・Webサイトデータを消す前にバックアップしてください。';document.querySelector('#prep .prepactions').after(info);
 const download=(blob,name)=>{const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),60000)};
 document.getElementById('ipadBackup').onclick=async()=>{try{await yukFolder.flush();const f=await yukFolder.read('大会進行データ.json');if(!f)throw Error('保存済み大会がありません');const data=JSON.parse(await f.text());const x=await yukFolder.read('対戦表_準備用.xlsx');const packageData={format:'yukuhashi-ipad-backup-v2',data,excel:x?toB64(new Uint8Array(await x.arrayBuffer())):null};download(new Blob([JSON.stringify(packageData)],{type:'application/json'}),'行橋支部方式_iPad_大会バックアップ.json')}catch(e){alert(e.message)}};
 document.getElementById('ipadImport').onchange=async e=>{try{const f=e.target.files[0];if(!f)return;const parsed=JSON.parse(await f.text());const data=parsed.format==='yukuhashi-ipad-backup-v2'?parsed.data:parsed;if(!Array.isArray(data.classes)||!data.classes.length||!data.classes.every(c=>Array.isArray(c.players)&&c.pairings&&c.results))throw Error('大会データ形式が不正です');if(!confirm('保存中の大会を置き換えます。続行しますか？'))return;if(parsed.format==='yukuhashi-ipad-backup-v2'&&parsed.excel){const bytes=fromB64(parsed.excel);if(bytes.length<4||bytes[0]!==80||bytes[1]!==75)throw Error('Excelデータが不正です');await yukFolder.write('対戦表_準備用.xlsx',new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));setExcelTemplate(bytes)}else{await yukFolder.remove('対戦表_準備用.xlsx');setExcelTemplate(new Uint8Array())}await yukFolder.queueState(data);alert('バックアップを保存しました。「保存大会を再開」を押してください。')}catch(err){alert(err.message)}finally{e.target.value=''}};
 document.getElementById('ipadPrepDownload').onclick=async()=>{const f=await yukFolder.read('対戦表_準備用.xlsx');if(f)download(f,'対戦表_準備用.xlsx');else alert('準備用Excelが保存されていません')};
 document.getElementById('ipadExcel').onclick=async()=>{
  try{
    const blob=await exportTournamentExcel(tournamentClasses,rankForClass);
    if(!blob)throw Error('Excelファイルを生成できませんでした');
    // exportTournamentExcel stores the workbook in the virtual folder;
    // iPad requires a separate, explicit user-initiated download.
    const old=document.getElementById('ipadExcelReady');if(old)old.remove();
    const url=URL.createObjectURL(blob);
    const link=document.createElement('a');link.id='ipadExcelReady';link.href=url;
    link.download='対戦表_大会用.xlsx';link.textContent='Excelファイルを保存（ここをタップ）';
    link.style.cssText='display:inline-block;padding:14px;margin:8px;background:#e0f2e5;border:2px solid #237346;font-size:20px;color:#123';
    document.getElementById('ipadExcel').after(link);
    alert('Excelの準備ができました。緑色の「Excelファイルを保存」をタップしてください。');
  }catch(e){alert('Excel出力：'+e.message)}
};
 const original=document.getElementById('folderMock');original.textContent='保存大会を再開';
 const notice=document.getElementById('folderNotice');notice.textContent='iPad内に自動保存します。新しい大会は準備用Excelを選択してください。';
 // Safari: keep selects native and accessible; enlarge tap targets in CSS.
})();
