import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.js?url";

function localOcrPath(name="") {
  const base=String(import.meta.env.BASE_URL||"/").replace(/\/$/,"");
  return `${base}/ocr${name?`/${name}`:""}`;
}

async function extractPdfText(file) {
  const pdfjsLib=await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc=pdfWorkerUrl;
  const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
  let fullText="";
  for(let i=1;i<=pdf.numPages;i++){
    const page=await pdf.getPage(i),content=await page.getTextContent(),lineGroups=[];
    for(const item of content.items){
      const y=item.transform[5],x=item.transform[4];
      let group=lineGroups.find(candidate=>Math.abs(candidate.y-y)<3);
      if(!group){group={y,words:[]};lineGroups.push(group);}
      group.words.push({x,str:item.str});
    }
    lineGroups.sort((a,b)=>b.y-a.y);
    for(const group of lineGroups){group.words.sort((a,b)=>a.x-b.x);fullText+=group.words.map(word=>word.str).join(" ")+"\n";}
  }
  if(fullText.trim().length>40)return fullText;

  const {createWorker}=await import("tesseract.js");
  let worker;
  try{
    worker=await createWorker("eng",1,{workerPath:localOcrPath("worker.min.js"),corePath:localOcrPath("core"),langPath:localOcrPath()});
  }catch(error){throw new Error(`This scanned PDF could not start local text recognition: ${error.message||String(error)}`);}
  let ocrText="";
  try{
    for(let i=1;i<=pdf.numPages;i++){
      const page=await pdf.getPage(i),viewport=page.getViewport({scale:2}),canvas=document.createElement("canvas");
      canvas.width=viewport.width;canvas.height=viewport.height;
      await page.render({canvasContext:canvas.getContext("2d"),viewport}).promise;
      const {data:{text}}=await worker.recognize(canvas);ocrText+=text+"\n";
    }
  }finally{await worker.terminate();}
  return ocrText;
}

function htmlToText(html){
  const doc=new DOMParser().parseFromString(String(html||""),"text/html");
  doc.querySelectorAll("script,style,iframe,object,svg").forEach(element=>element.remove());
  for(const row of [...doc.querySelectorAll("tr")]){
    const cells=[...row.querySelectorAll(":scope > th,:scope > td")].map(element=>element.textContent.trim().replace(/\s+/g," "));
    row.replaceWith(doc.createTextNode("\n"+cells.join("\t")+"\n"));
  }
  for(const br of [...doc.querySelectorAll("br")])br.replaceWith(doc.createTextNode("\n"));
  for(const element of [...doc.querySelectorAll("p,div,li,h1,h2,h3")])element.append(doc.createTextNode("\n"));
  return (doc.body?.textContent||"").replace(/\n[ \t]+/g,"\n");
}

function emailToText(raw){
  const source=String(raw||"").replace(/\r\n/g,"\n"),headerEnd=source.indexOf("\n\n");
  if(headerEnd<0)return source;
  const headers=source.slice(0,headerEnd),body=source.slice(headerEnd+2);
  const boundary=headers.match(/boundary\s*=\s*["']?([^"';\s]+)/i)?.[1];
  const parts=boundary?body.split(`--${boundary}`):[source];
  const decode=(partHeaders,partBody)=>{
    const encoding=partHeaders.match(/content-transfer-encoding\s*:\s*([^\s;]+)/i)?.[1]?.toLowerCase();
    if(encoding==="base64"){
      try{return new TextDecoder().decode(Uint8Array.from(atob(partBody.replace(/\s/g,"")),character=>character.charCodeAt(0)));}catch{return "";}
    }
    if(encoding==="quoted-printable"){
      const normalized=partBody.replace(/=\n/g,"");
      try{const bytes=[];for(let i=0;i<normalized.length;i++){
        if(normalized[i]==="="&&/^[0-9a-f]{2}$/i.test(normalized.slice(i+1,i+3))){bytes.push(parseInt(normalized.slice(i+1,i+3),16));i+=2;}
        else bytes.push(...new TextEncoder().encode(normalized[i]));
      }return new TextDecoder().decode(new Uint8Array(bytes));}catch{return normalized;}
    }
    return partBody;
  };
  for(const kind of ["text/plain","text/html"]){
    for(const part of parts){
      const index=part.indexOf("\n\n");if(index<0)continue;
      const partHeaders=part.slice(0,index),partBody=part.slice(index+2);
      if(!new RegExp(`content-type\\s*:\\s*${kind.replace("/","\\/")}`,"i").test(partHeaders))continue;
      if(/content-disposition\s*:\s*attachment/i.test(partHeaders))continue;
      const value=decode(partHeaders,partBody.replace(/\n--\s*$/, ""));
      return kind==="text/html"?htmlToText(value):headers.split("\n").filter(line=>/^(?:from|subject|date):/i.test(line)).join("\n")+"\n"+value;
    }
  }
  return headers.split("\n").filter(line=>/^(?:from|subject|date):/i.test(line)).join("\n")+"\n"+body;
}

export async function fileToText(file){
  const name=file.name.toLowerCase();
  if(name.endsWith(".eml"))return emailToText(await file.text());
  if(name.endsWith(".html")||name.endsWith(".htm"))return htmlToText(await file.text());
  if(name.endsWith(".xlsx")||name.endsWith(".xls")){
    const XLSX=await import("xlsx"),workbook=XLSX.read(await file.arrayBuffer(),{type:"array"});
    return XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]);
  }
  if(name.endsWith(".pdf"))return extractPdfText(file);
  return file.text();
}

export {emailToText,htmlToText,extractPdfText};
