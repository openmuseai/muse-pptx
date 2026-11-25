/** Standard browser/MCP Apps services for the original desktop preload API.
 * No native paths, IPC, external AI provider or unprompted clipboard reads. */
import {Buffer} from 'buffer'
import {installBrowserFont,browserFontFaces,browserFontData} from './font-metrics'
import {buildPrintDocumentHtml,normalizePrintRatio} from '../../vendor/genoffice/apps/slides/src/shared/print-html'
import {presentation} from './presentation'
export function pickBrowserFile(accept:string):Promise<File|null>{
  return new Promise(resolve=>{const input=document.createElement('input');input.type='file';input.accept=accept;input.hidden=true;let done=false
    const finish=(file:File|null)=>{if(done)return;done=true;input.remove();resolve(file)}
    input.addEventListener('change',()=>finish(input.files?.[0]??null),{once:true});input.addEventListener('cancel',()=>finish(null),{once:true});document.body.append(input);input.click()
  })
}
interface BrowserPort {download(name:string,mime:string,bytes:Uint8Array):Promise<boolean>;printDocument?(html:string,widthMils:number,heightMils:number):Promise<{ok:boolean}|null>;emit(name:string,value:any):void;language():string;theme():string}
export function createBrowserServices(p:BrowserPort){
  let aiPrefs={fontSize:14,spellcheck:false}
  const name=(value:string,fallback:string)=>value?.split(/[\\/]/).pop()?.replace(/[\x00-\x1f]/g,'').slice(0,200)||fallback
  return {
    privateFontFaces:async()=>browserFontFaces(),privateFontData:async(id:string)=>browserFontData(id),
    fontCatalog:async()=>[{family:'Carlito',script:'latin',license:'OFL-1.1',installed:true,bytes:0},{family:'Arimo',script:'latin',license:'OFL-1.1',installed:true,bytes:0}],
    fontDownload:async(family:string)=>['Carlito','Arimo'].includes(family)?{ok:true}:{ok:false,error:'Font is not in the bundled catalog'},
    fontInstallLocal:async()=>{const file=await pickBrowserFile('.ttf,.otf');if(!file)return {families:[]};const family=await installBrowserFont(file);p.emit('onFontsChanged',null);return {families:[family]}},
    getAiPanelPrefs:async()=>aiPrefs,setAiPanelPrefs:async(value:any)=>{aiPrefs={fontSize:Math.min(32,Math.max(10,Number(value.fontSize??aiPrefs.fontSize))),spellcheck:!!(value.spellcheck??aiPrefs.spellcheck)};p.emit('onAiPanelPrefsChanged',aiPrefs);return aiPrefs},
    nativeClipboard:async(op:'copy'|'cut'|'paste')=>{if(op==='paste'){const text=await navigator.clipboard.readText();document.execCommand('insertText',false,text);return}if(!document.execCommand(op))throw Error('Clipboard permission denied')},
    microphoneAccess:async()=>typeof navigator.mediaDevices?.getUserMedia==='function',
    setShowFullScreen:async(on:boolean)=>{if(on){if(document.documentElement.requestFullscreen&&!document.fullscreenElement)try{await document.documentElement.requestFullscreen()}catch{/* The original show still fills its embedded viewport. */}}else if(document.fullscreenElement)await document.exitFullscreen()},
    presenterStart:async()=>({audience:false}),presenterEnd:async()=>{},presenterSync:(value:any)=>p.emit('onShowSync',value),presenterInk:(value:any)=>p.emit('onShowInk',value),presenterSwap:async()=>({audience:false}),
    audienceReady:async()=>({audience:false}),audienceNav:(value:any)=>p.emit('onAudienceNav',value),
    pickExportDir:async()=>'download',pickExportPdfPath:async(defaultName:string)=>name(defaultName,'presentation.pdf'),
    exportImages:async(o:any)=>{const paths=[];for(const [i,png] of o.pngsBase64.entries()){const path=name(o.baseName,'presentation')+'-'+String(i+1).padStart(2,'0')+'.png';if(!await p.download(path,'image/png',pngBytes(png)))return {ok:false,error:'Download canceled'};paths.push(path)}return {ok:true,paths}},
    savePicture:async(o:any)=>{const path=name(o.defaultName,'picture').replace(/\.png$/i,'')+'.png';const ok=await p.download(path,'image/png',pngBytes(o.pngBase64));return {ok,path:ok?path:undefined}},
    exportPdf:async(o:any)=>{const path=name(o.filePath,'presentation.pdf');const bytes=await rasterPdf(o);const ok=await p.download(path,'application/pdf',bytes);return {ok,path:ok?path:undefined}},
    printSlides:async(o:any)=>{const srcs=o.pngsBase64.map((png:string)=>'data:image/png;base64,'+Buffer.from(pngBytes(png)).toString('base64')),ratio=normalizePrintRatio(o.widthPx/o.heightPx)
      const options={...o,srcs,ratio,layout:o.layout??'full'},html=buildPrintDocumentHtml(options)
      const full=options.layout==='full',landscape=!full&&o.orientation==='landscape',width=full?ratio*7.5:landscape?11.69:8.27,height=full?7.5:landscape?8.27:11.69
      try{const native=await p.printDocument?.(html,Math.round(width*1000),Math.round(height*1000));if(native)return native}catch{return {ok:false,error:presentation.integration('printingUnavailable')}}
      const frame=document.createElement('iframe');frame.hidden=true;document.body.append(frame)
      return new Promise<{ok:boolean,error?:string}>(resolve=>{let done=false;const finish=(ok:boolean)=>{if(done)return;done=true;frame.remove();resolve(ok?{ok:true}:{ok:false,error:presentation.integration('printingUnavailable')})}
        frame.onload=async()=>{try{const w=frame.contentWindow;if(!w){finish(false);return}await Promise.all([...w.document.images].map(i=>i.decode()));let opened=false;w.addEventListener('beforeprint',()=>{opened=true},{once:true});w.addEventListener('afterprint',()=>finish(true),{once:true});w.print();setTimeout(()=>finish(opened),1500)}catch{finish(false)}};setTimeout(()=>finish(false),30000);frame.srcdoc=html
      })
    },
    reportCloseSaveResult:async()=>{},headlessExportDone:async()=>{},
  }
}
function pngBytes(base64:string){const bytes=new Uint8Array(Buffer.from(base64,'base64'));if(bytes.length>64*1024*1024||bytes[0]!==137||bytes[1]!==80||bytes[2]!==78||bytes[3]!==71)throw Error('Invalid PNG export');return bytes}

/** Portable browser fallback: real rendered pages in a standards-compliant PDF.
 * The desktop vector exporter uses Electron printToPDF and cannot run in WebView.
 * Links remain annotations; pages are rasterized locally, never uploaded. */
async function rasterPdf(o:any):Promise<Uint8Array>{
  if(!Array.isArray(o.pages)||!o.pages.length||!Number.isFinite(o.widthPx)||!Number.isFinite(o.heightPx)||o.widthPx<=0||o.heightPx<=0)throw Error('Invalid PDF pages')
  const encoder=new TextEncoder(),chunks:Uint8Array[]=[],offsets=[0];let size=0
  const put=(s:string|Uint8Array)=>{const bytes=typeof s==='string'?encoder.encode(s):s;chunks.push(bytes);size+=bytes.length}
  const objects:Array<string|{header:string;bytes:Uint8Array}>=['<< /Type /Catalog /Pages 2 0 R >>','']
  const kids:number[]=[]
  for(const [pageIndex,page] of o.pages.entries()){const image=new Image();let url:string
    if(page.png)url='data:image/png;base64,'+page.png
    else {if(typeof page.svg!=='string')throw Error('Invalid PDF page');const svg=page.svg.replace(/(<svg\b[^>]*>)/,'$1<style>'+String(o.fontCss??'').replaceAll('</style','')+'</style>');url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}))}
    let jpeg:Uint8Array,w:number,h:number
    try{image.src=url;await image.decode();const scale=Math.min(2,4096/Math.max(o.widthPx,o.heightPx));w=Math.max(1,Math.round(o.widthPx*scale));h=Math.max(1,Math.round(o.heightPx*scale));const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;const ctx=canvas.getContext('2d');if(!ctx)throw Error('PDF canvas unavailable');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(image,0,0,w,h);jpeg=new Uint8Array(Buffer.from(canvas.toDataURL('image/jpeg',.95).split(',')[1],'base64'));canvas.width=canvas.height=0}finally{if(url.startsWith('blob:'))URL.revokeObjectURL(url)}
    const pageId=objects.length+1,imageId=pageId+1,contentId=pageId+2,pw=o.widthPx*.75,ph=o.heightPx*.75;kids.push(pageId)
    const content=`q ${pw} 0 0 ${ph} 0 0 cm /Im0 Do Q`
    const annotations=[]
    for(const link of o.links?.[pageIndex]??[]){
      if(![link.x,link.y,link.w,link.h].every(Number.isFinite)||link.w<=0||link.h<=0)continue
      const x=Math.max(0,link.x)*pw,y=Math.max(0,link.y)*ph,r=Math.min(pw,(link.x+link.w)*pw),top=Math.min(ph,(link.y+link.h)*ph)
      if(r<=x||top<=y)continue
      let action:string
      const internal=/^#(?:pg|page=)?(\d+)$/.exec(link.href)
      if(internal){const target=Number(internal[1])-1;if(target<0||target>=o.pages.length)continue;action=`/Dest [PAGE_${target} 0 R /Fit]`}
      else{let url:URL;try{url=new URL(link.href)}catch{continue}if(!['https:','http:','mailto:','tel:'].includes(url.protocol)||url.username||url.password)continue;action='/A << /S /URI /URI ('+url.href.replaceAll('\\','\\\\').replaceAll('(','\\(').replaceAll(')','\\)')+') >>'}
      annotations.push(`<< /Type /Annot /Subtype /Link /Rect [${x} ${ph-top} ${r} ${ph-y}] /Border [0 0 0] ${action} >>`)
    }
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /XObject << /Im0 ${imageId} 0 R >> >> /Contents ${contentId} 0 R /Annots [${annotations.join(' ')}] >>`,{header:`<< /Type /XObject /Subtype /Image /Width ${w!} /Height ${h!} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpeg!.length} >>`,bytes:jpeg!},{header:`<< /Length ${encoder.encode(content).length} >>`,bytes:encoder.encode(content)})
  }
  for(let i=0;i<objects.length;i++)if(typeof objects[i]==='string')objects[i]=(objects[i] as string).replace(/PAGE_(\d+) 0 R/g,(_,n)=>kids[Number(n)]+' 0 R')
  objects[1]=`<< /Type /Pages /Count ${kids.length} /Kids [${kids.map(id=>id+' 0 R').join(' ')}] >>`
  put('%PDF-1.4\n');for(const [i,obj] of objects.entries()){offsets.push(size);put(`${i+1} 0 obj\n`);if(typeof obj==='string')put(obj);else{put(obj.header+'\nstream\n');put(obj.bytes);put('\nendstream')}put('\nendobj\n')}
  const xref=size;put(`xref\n0 ${objects.length+1}\n0000000000 65535 f \n`);for(const off of offsets.slice(1))put(String(off).padStart(10,'0')+' 00000 n \n');put(`trailer\n<< /Size ${objects.length+1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`)
  const result=new Uint8Array(size);let pos=0;for(const bytes of chunks){result.set(bytes,pos);pos+=bytes.length}return result
}
