// Only the registered BI frame can use this fixed loopback reserve transport.
(function(){
 'use strict';
 const endpoint='http://127.0.0.1:53627';
 const routes={health:['GET','/reserve-health'],start:['POST','/reserve-jobs/start'],status:['POST','/reserve-jobs/status'],result:['POST','/reserve-jobs/result']};
 const pending=new Set();
 window.addEventListener('message',async function(e){
  const m=e.data||{};
  if(m.type!=='bi-reserve-local-request'||e.source!==viewportClient||e.origin!==viewportOrigin||!isAppDescendant(e.source)||!/^[a-zA-Z0-9_-]{16,100}$/.test(String(m.id||''))||!Object.hasOwn(routes,m.action)||pending.has(m.id))return;
  pending.add(m.id);
  const send=x=>e.source.postMessage(Object.assign({type:'bi-reserve-local-result',id:m.id},x),e.origin);
  try{
   const [method,path]=routes[m.action],controller=new AbortController(),timer=setTimeout(()=>controller.abort(),20000);
   let response;
   try{response=await fetch(endpoint+path,{method,mode:'cors',credentials:'omit',cache:'no-store',targetAddressSpace:'loopback',signal:controller.signal,...(method==='POST'?{headers:{'Content-Type':'application/json'},body:JSON.stringify(m.payload||{})}:{})});}finally{clearTimeout(timer);}
   const data=await response.json();if(!response.ok)throw Error(data.error||'Загрузчик не выполнил запрос');send({result:data});
  }catch(error){send({error:error.name==='AbortError'?'Загрузчик не ответил вовремя. Проверяем состояние перед повтором.':'Нет связи с загрузчиком BI. Проверьте, что он запущен; если браузер запрашивает доступ к приложению на этом компьютере, разрешите его.'});}
  finally{pending.delete(m.id);}
 });
})();
