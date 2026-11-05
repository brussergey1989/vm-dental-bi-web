// Private bytes arrive only from the embedding BI / Mini App window. No file upload or external viewer.
const allowed=origin=>origin===location.origin||origin==='https://vm-dental-bi.ru'||/^https:\/\/[a-z0-9-]+\.googleusercontent\.com$/.test(origin)||origin==='https://script.google.com';
const content=document.getElementById('content'),status=document.getElementById('status'),controls=document.getElementById('controls');
const nonce=crypto.randomUUID();let accepted=false,objectUrl='',printPdf=null,canPrint=false,printBusy=false;
function message(text){status.textContent=text;}
function text(value){const pre=document.createElement('pre');pre.textContent=value;content.replaceChildren(pre);}
function zipGuard(bytes){
 if(bytes[0]!==80||bytes[1]!==75)return;
 let total=0,count=0;
 for(let i=Math.max(0,bytes.length-65557);i<bytes.length-21;i++)if(bytes[i]===80&&bytes[i+1]===75&&bytes[i+2]===5&&bytes[i+3]===6){
  const data=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),entries=data.getUint16(i+10,true),offset=data.getUint32(i+16,true);if(entries>5000)throw Error('Слишком сложный архив документа. Скачайте оригинал.');
  let at=offset;for(let j=0;j<entries;j++){if(at+46>bytes.length||data.getUint32(at,true)!==0x02014b50)throw Error('Архив документа повреждён.');const size=data.getUint32(at+24,true);total+=size;count++;if(total>24*1024*1024||size>12*1024*1024)throw Error('Документ слишком большой для просмотра. Скачайте оригинал.');at+=46+data.getUint16(at+28,true)+data.getUint16(at+30,true)+data.getUint16(at+32,true);}return;
 }
 throw Error('Архив документа повреждён или защищён паролем.');
}
async function loadPdf(bytes){
 const pdfjs=await import('./vendor/pdf.mjs');pdfjs.GlobalWorkerOptions.workerSrc=new URL('./vendor/pdf.worker.mjs',import.meta.url).href;
 return pdfjs.getDocument({data:bytes,isEvalSupported:false,useSystemFonts:true,disableFontFace:true,stopAtErrors:true}).promise;
}
async function showPdf(bytes){
 const pdf=await loadPdf(bytes);let number=1,rendering=false;
 printPdf=pdf;const back=document.createElement('button'),next=document.createElement('button'),label=document.createElement('span');back.textContent='←';back.setAttribute('aria-label','Предыдущая страница');next.textContent='→';next.setAttribute('aria-label','Следующая страница');controls.replaceChildren(back,label,next);
 async function page(){if(rendering)return;rendering=true;back.disabled=next.disabled=true;try{message('Открываем страницу…');const p=await pdf.getPage(number),original=p.getViewport({scale:1}),scale=Math.min(1.7,Math.sqrt(2500000/(original.width*original.height))),viewport=p.getViewport({scale}),canvas=document.createElement('canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);await p.render({canvasContext:canvas.getContext('2d'),viewport}).promise;content.replaceChildren(canvas);label.textContent=number+' / '+pdf.numPages;message('');p.cleanup();}finally{rendering=false;back.disabled=number<=1;next.disabled=number>=pdf.numPages;}}
 back.onclick=()=>{if(number>1&&!rendering){number--;page().catch(e=>message(e.message));}};next.onclick=()=>{if(number<pdf.numPages&&!rendering){number++;page().catch(e=>message(e.message));}};await page();
}
async function showSheet(bytes){
 await import('./vendor/xlsx.full.min.js');const book=window.XLSX.read(bytes,{type:'array',cellFormula:false,cellHTML:false,cellStyles:false,sheetRows:501,bookVBA:false});
 const select=document.createElement('select');select.setAttribute('aria-label','Лист документа');book.SheetNames.forEach(name=>{const option=document.createElement('option');option.textContent=option.value=name;select.append(option);});controls.replaceChildren(select);
 function render(){const rows=window.XLSX.utils.sheet_to_json(book.Sheets[select.value],{header:1,raw:false,defval:''}),table=document.createElement('table');rows.slice(0,500).forEach(row=>{const tr=document.createElement('tr');row.slice(0,60).forEach(value=>{const td=document.createElement('td');td.textContent=String(value);tr.append(td);});table.append(tr);});content.replaceChildren(table);message('Просмотр значений: до 500 строк и 60 столбцов выбранного листа. Формулы не выполняются.');}select.onchange=render;render();
}
async function showTextDocument(bytes,ext){
 const {unzipSync,strFromU8}=await import('./vendor/fflate.mjs'),path=ext==='docx'?'word/document.xml':'content.xml',files=unzipSync(bytes,{filter:file=>file.name===path});if(!files[path])throw Error('Не найден текст документа.');const xml=new DOMParser().parseFromString(strFromU8(files[path]),'application/xml');if(xml.querySelector('parsererror'))throw Error('Не удалось прочитать документ.');
 const paragraphs=[...xml.getElementsByTagNameNS('*','p')],value=paragraphs.map(p=>p.textContent).join('\n\n');text(value||'В документе нет доступного текста.');message('Текст документа без исходного оформления. Оригинал можно скачать отдельно.');
}
window.addEventListener('message',async event=>{
 if(event.source===parent&&allowed(event.origin)&&event.data?.type==='request-document-print'&&event.data.nonce===nonce){await printDocument();return;}
 if(event.source===parent&&allowed(event.origin)&&event.data?.type==='request-document-print-batch'&&event.data.nonce===nonce){
  await printBatch(event.data,event.origin);
  return;
 }
 if(accepted||event.source!==parent||!allowed(event.origin)||event.data?.type!=='request-document-bytes'||event.data.nonce!==nonce)return;
 accepted=true;
 try{const file=event.data.file,bytes=new Uint8Array(file.bytes),ext=String(file.name).split('.').pop().toLowerCase();const limit=ext==='pdf'&&file.printable===true?10:5;if(!bytes.length||bytes.length>limit*1024*1024)throw Error('Размер файла — до '+limit+' МБ.');document.getElementById('name').textContent=file.name;zipGuard(bytes);
  if(ext==='pdf'){await showPdf(bytes);canPrint=file.printable===true;if(canPrint){const print=document.createElement('button');print.textContent='Печать';print.onclick=printDocument;controls.append(print);if(file.autoPrint)await printDocument();}}else if(/^(jpe?g)$/.test(ext)){const image=document.createElement('img');image.alt=file.name;objectUrl=URL.createObjectURL(new Blob([bytes],{type:'image/jpeg'}));image.src=objectUrl;content.replaceChildren(image);message('');}
  else if(['xlsx','xls','ods','csv'].includes(ext))await showSheet(bytes);else if(['docx','odt'].includes(ext))await showTextDocument(bytes,ext);else throw Error('Формат не поддерживается для просмотра. Скачайте оригинал.');
 }catch(error){message('Не удалось открыть документ. '+(error.message||'Скачайте оригинал.'));}
});
window.addEventListener('unload',()=>{if(objectUrl)URL.revokeObjectURL(objectUrl);});
parent.postMessage({type:'request-document-ready',nonce},'*');

async function printDocument(){
 if(!canPrint||!printPdf||printBusy)return;printBusy=true;let pages;
 try{
  message('Готовим все страницы к печати…');pages=document.getElementById('print-content');if(pages)pages.remove();pages=document.createElement('div');pages.id='print-content';document.body.append(pages);
  if(printPdf.numPages>100)throw Error('Для печати документа более 100 страниц скачайте PDF.');
  for(let i=1;i<=printPdf.numPages;i++){const p=await printPdf.getPage(i),viewport=p.getViewport({scale:1.5}),canvas=document.createElement('canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);await p.render({canvasContext:canvas.getContext('2d'),viewport}).promise;pages.append(canvas);p.cleanup();}
  message('');window.print();
 }catch(error){message('Откройте скачанный PDF и выберите печать в просмотрщике. '+(error.message||''));}
 finally{printBusy=false;}
}
async function printBatch(data,origin){
 if(!canPrint||typeof data.jobId!=='string')return;
 function respond(stage,message){
  parent.postMessage({type:'request-document-print-batch-status',nonce,jobId:data.jobId,stage,message},origin);
 }
 if(printBusy){respond('failed','Дождитесь завершения текущей печати.');return;}
 printBusy=true;
 let pages;
 try{
  const files=data.files;
  if(!Array.isArray(files)||!files.length||files.length>40)throw Error('Выберите от 1 до 40 PDF.');
  let bytesTotal=0;
  for(const file of files){
   if(file?.mime!=='application/pdf'||!(file.bytes instanceof Uint8Array)||!file.bytes.length||file.bytes.length>10*1024*1024)throw Error('Один из PDF недоступен для печати.');
   bytesTotal+=file.bytes.length;
  }
  if(bytesTotal>45*1024*1024)throw Error('Размер задания превышает 45 МБ.');
  respond('started');
  document.getElementById('print-content')?.remove();
  pages=document.createElement('div');
  pages.id='print-content';
  pages.className='print-batch';
  document.body.append(pages);
  let count=0;
  for(const [index,file] of files.entries()){
   message('Готовим документ '+(index+1)+' из '+files.length+'…');
   const pdf=await loadPdf(file.bytes);
   try{
    count+=pdf.numPages;
    if(count>100)throw Error('Для одного задания выберите до 100 страниц.');
    for(let number=1;number<=pdf.numPages;number++){
     const page=await pdf.getPage(number);
     const size=page.getViewport({scale:1});
     const scale=Math.min(1.5,Math.sqrt(2000000/(size.width*size.height)));
     const viewport=page.getViewport({scale});
     const canvas=document.createElement('canvas');
     canvas.width=Math.ceil(viewport.width);
     canvas.height=Math.ceil(viewport.height);
     await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise;
     const sheet=document.createElement('section');
     sheet.className='print-sheet';
     sheet.append(canvas);
     pages.append(sheet);
     page.cleanup();
    }
   }finally{await pdf.destroy();}
  }
  message('');
  window.print();
  respond('finished');
 }catch(error){
  const detail=error.message||'Не удалось подготовить документы.';
  message(detail);
  respond('failed',detail);
 }finally{
  pages?.remove();
  printBusy=false;
 }
}
window.addEventListener('afterprint',()=>document.getElementById('print-content')?.remove());
