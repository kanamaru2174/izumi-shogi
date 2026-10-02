"use strict";
// iPad local storage adapter: same yukFolder interface as desktop, without directory permissions.
window.yukFolder=(()=>{
 const DB='yukuhashi-ipad-tournament-v1';let pending=Promise.resolve();
 function db(){return new Promise((resolve,reject)=>{const r=indexedDB.open(DB,1);r.onupgradeneeded=()=>r.result.createObjectStore('files');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)})}
 async function get(name){const d=await db();try{return await new Promise((resolve,reject)=>{const r=d.transaction('files').objectStore('files').get(name);r.onsuccess=()=>resolve(r.result||null);r.onerror=()=>reject(r.error)})}finally{d.close()}}
 async function put(name,blob){const d=await db();try{await new Promise((resolve,reject)=>{const t=d.transaction('files','readwrite');t.objectStore('files').put(blob,name);t.oncomplete=resolve;t.onerror=()=>reject(t.error);t.onabort=()=>reject(t.error)})}finally{d.close()}}
 async function remove(name){const d=await db();try{await new Promise((resolve,reject)=>{const t=d.transaction('files','readwrite');t.objectStore('files').delete(name);t.oncomplete=resolve;t.onerror=()=>reject(t.error)})}finally{d.close()}}
 async function read(name){const blob=await get(name);return blob?new File([blob],name,{type:blob.type}):null}
 async function write(name,blob){await put(name,blob)}
 async function load(){const file=await read('大会進行データ.json');if(!file)return null;const data=JSON.parse(await file.text());if(!Array.isArray(data.classes)||!data.classes.length)throw Error('保存大会データが不正です');const excel=await read('対戦表_準備用.xlsx');if(excel)setExcelTemplate(new Uint8Array(await excel.arrayBuffer()));return data}
 function queueState(data){pending=pending.catch(()=>{}).then(()=>put('大会進行データ.json',new Blob([JSON.stringify(data)],{type:'application/json'})));return pending}
 return {handle:{name:'iPad端末内'},pick:load,read,write,remove,load,queueState,flush:()=>pending,hasTournamentFiles:async()=>!!(await get('大会進行データ.json')),disconnect:()=>{}};
})();
