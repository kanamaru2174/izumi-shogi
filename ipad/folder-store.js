// User-approved local tournament folder. Supported in desktop Chromium secure contexts.
'use strict';
window.yukFolder=(()=>{
 let handle=null, pending=Promise.resolve();
 const notice=()=>document.getElementById('folderNotice');
 const show=t=>{if(notice())notice().textContent=t};
 function db(){return new Promise((resolve,reject)=>{const req=indexedDB.open('yukuhashi-folder-v1',1);req.onupgradeneeded=()=>req.result.createObjectStore('handles');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
 async function storeHandle(h){const d=await db();await new Promise((resolve,reject)=>{const tx=d.transaction('handles','readwrite');tx.objectStore('handles').put(h,'last');tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)});d.close()}
 async function lastHandle(){const d=await db();const h=await new Promise((resolve,reject)=>{const tx=d.transaction('handles');const req=tx.objectStore('handles').get('last');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)});d.close();return h}
 async function read(name){if(!handle)return null;try{return await (await handle.getFileHandle(name)).getFile()}catch(e){if(e.name==='NotFoundError')return null;throw e}}
 async function write(name,blob){if(!handle)throw Error('大会フォルダが未選択です');if(await handle.queryPermission({mode:'readwrite'})!=='granted')throw Error('大会フォルダの書き込み許可がありません。フォルダを選択し直してください');const f=await handle.getFileHandle(name,{create:true}),w=await f.createWritable();await w.write(blob);await w.close();show('大会フォルダ：'+handle.name+'（保存済み）')}
 async function load(){let file=await read('大会進行データ.json');if(!file)return null;const data=JSON.parse(await file.text());if(!Array.isArray(data.classes)||!data.classes.length)throw Error('大会データが不正です');const excel=await read('対戦表_準備用.xlsx');if(excel)setExcelTemplate(new Uint8Array(await excel.arrayBuffer()));return data}
 async function pick(){if(!window.showDirectoryPicker)throw Error('このブラウザはフォルダへの直接読み書きに対応していません。Windows版Chrome／EdgeのlocalhostまたはHTTPSで起動してください');const h=await showDirectoryPicker({mode:'readwrite',id:'yukuhashi-tournament'});handle=h;await storeHandle(h);show('大会フォルダ：'+h.name);return await load()}
 async function autoRestore(callback){try{if(!window.showDirectoryPicker)return;const h=await lastHandle();if(!h)return;handle=h;show('前回の大会フォルダ：'+h.name+'（アクセス権確認中）');if(await h.queryPermission({mode:'readwrite'})==='granted'){await callback(await load());show('大会フォルダ：'+h.name+'（復元済み）')}else{handle=null;show('前回の大会フォルダ：'+h.name+'。再開するには「大会フォルダを選択／再開」を押してください')}}catch(e){handle=null;show('フォルダの自動復元失敗：'+e.message)}}
 async function backupTournamentFiles(){
  if(!handle)throw Error('大会フォルダが未選択です');
  const stamp=new Date().toISOString().replace(/[-:TZ.]/g,'').slice(0,14);
  const names=['大会進行データ.json','対戦表_大会用.xlsx','対戦表_準備用.xlsx'];
  const files=[];
  for(const name of names){const f=await read(name);if(f)files.push({name,blob:await f.arrayBuffer()})}
  if(!files.length)throw Error('バックアップ対象がありません');
  for(const f of files){const dot=f.name.lastIndexOf('.');const backupName=f.name.slice(0,dot)+'_backup_'+stamp+f.name.slice(dot);await write(backupName,new Blob([f.blob]))}
 }
 function queueState(data){const json=JSON.stringify(data,null,2);pending=pending.catch(()=>{}).then(()=>write('大会進行データ.json',new Blob([json],{type:'application/json'}))).catch(e=>{show('フォルダ保存エラー：'+e.message);throw e});return pending}
 return {get handle(){return handle},pick,hasTournamentFiles:async()=>!!(await read('大会進行データ.json')||await read('対戦表_大会用.xlsx')),autoRestore,write,backupTournamentFiles,queueState,flush:()=>pending,disconnect:()=>{handle=null}};
})();
