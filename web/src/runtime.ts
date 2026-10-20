/** Plugin-owned browser adapter for the unmodified GenOffice editor preload API.
 * The Host transfers standard MCP Apps tool input and download-file messages. All PPTX parsing/rendering/editing
 * stays in the upstream libraries bundled with this plugin.
 */
import { Buffer } from 'buffer'
import { App } from '@modelcontextprotocol/ext-apps'
import { openPptx, savePptx, createBlankPptx, getSlideNotes, type OpenedPptx } from '@genoffice/pptx-engine'
import { buildAllRenderSlides, rebuildSlide } from '../../vendor/genoffice/apps/slides/src/main/session-state'
import { runTxn, type Op } from '@genoffice/pptx-ops'
import {fontsReady} from './font-metrics'
import { AI_PROVIDERS } from '@genoffice/ai-provider/browser'

const mcpApp = window.parent !== window ? new App({name:'GenOffice PPTX',version:'0.4.0'},{}) : null
if (mcpApp) {
  mcpApp.ontoolinput = ({arguments:args}) => {
    if(typeof args?.content === 'string') void load(new File([Buffer.from(args.content,'base64')], typeof args.name==='string'?args.name:'Presentation.pptx', {type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})).catch(showError)
  }
  void mcpApp.connect().catch(showError)
}
let opened: OpenedPptx | null = null
let fileName = 'Presentation.pptx'
let fitWidthPx = 1280
let dirty = false
let pendingResolve: ((value: any) => void) | null = null
let firstOpen: any = null
const listeners = new Map<string, Set<(...args: any[]) => void>>()
const emit = (name: string, value: any) => listeners.get(name)?.forEach(fn => fn(value))
const undoStack: Uint8Array[] = []
const redoStack: Uint8Array[] = []
// The same Desktop session builder (including TIFF, themed SVG and JPEG
// orientation handling) produces the tree consumed by the original React UI.
const render = (index: number) => opened ? rebuildSlide({opened,fitWidthPx} as any,index) : null
const allSlides = () => buildAllRenderSlides(opened!,fitWidthPx)
const result = () => ({path:fileName,slides:allSlides(),size:opened!.deck.size})
const history = () => emit('onHistoryChanged',{canUndo:undoStack.length>0,canRedo:redoStack.length>0})
async function load(file: File) {
  if (file.size > 64*1024*1024) throw new Error('PPTX exceeds 64 MiB')
  await fontsReady()
  opened = await openPptx(new Uint8Array(await file.arrayBuffer()))
  fileName = file.name; dirty=false; undoStack.length=0;redoStack.length=0
  firstOpen=result()
  if(pendingResolve) {pendingResolve(firstOpen);pendingResolve=null}
  emit('onOpened',firstOpen);history()
  document.documentElement.dataset.pptxReady=String(opened.deck.slides.length)
}
window.addEventListener('unhandledrejection',event=>showError(event.reason))
window.addEventListener('message',event=>{
  if (event.data instanceof File) void load(event.data).catch(showError)
})
function showError(error: unknown) {
  console.error(error)
  const banner=document.createElement('div');banner.textContent=String(error)
  banner.style.cssText='position:fixed;bottom:20px;left:20px;background:#fff2f2;color:#9b1c1c;padding:12px;z-index:99999'
  document.body.append(banner)
}
async function apply(ops: Op[], index: number, all=false) {
  if(!opened) throw new Error('No presentation open')
  const before=await savePptx(opened)
  const outcome=runTxn(opened,{ops})
  if(!outcome.applied) throw new Error(outcome.failures?.map(f=>f.error).join('; ')||'Edit failed')
  undoStack.push(before);if(undoStack.length>30)undoStack.shift();redoStack.length=0;dirty=true;history()
  return all?allSlides():render(index)
}
const target=(op:any)=>({target:{slide:op.slideIndex,el:op.sourceId},...(op.groupId?{group:op.groupId}:{})})
async function save() {
  if(!opened)return {ok:false}
  const bytes=await savePptx(opened)
  const file=new File([bytes as BlobPart],fileName,{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})
  if(mcpApp) {
    const response=await mcpApp.downloadFile({contents:[{type:'resource',resource:{uri:'file:///'+encodeURIComponent(fileName),mimeType:file.type,blob:Buffer.from(bytes).toString('base64')}}]})
    if(response.isError) return {ok:false}
  } else {
    const a=document.createElement('a');a.href=URL.createObjectURL(file);a.download=fileName;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)
  }
  dirty=false
  return {ok:true,path:fileName,slides:allSlides()}
}
const settings={provider:'custom',providers:Object.fromEntries(AI_PROVIDERS.map(p=>[p.id,{apiKey:'',baseUrl:'',model:p.defaultModel}]))}
const api: Record<string, (...args:any[])=>any>={
  getLanguage:async()=>'zh',getTheme:async()=>'light',getAutoSaveDefault:async()=>false,
  getAiPanelPrefs:async()=>({fontSize:14,spellcheck:false}),setAiPanelPrefs:async(p:any)=>p,
  consumePendingOpen:async(width:number)=>{fitWidthPx=width;return opened?result():new Promise(resolve=>pendingResolve=resolve)},
  consumeHeadlessExport:async()=>null,
  openPptx:async()=>new Promise(resolve=>{const input=document.createElement('input');input.type='file';input.accept='.pptx';input.onchange=async()=>{if(input.files?.[0]){await load(input.files[0]);resolve(result())}else resolve(null)};input.click()}),
  newBlank:async(width:number)=>{fitWidthPx=width;await load(new File([await createBlankPptx() as BlobPart],'Presentation.pptx'));return result()},
  getLayouts:async()=>({layouts:[],size:opened?.deck.size??{cx:12192000,cy:6858000}}),
  privateFontFaces:async()=>[],privateFontData:async()=>null,fontMissing:async()=>[],fontCatalog:async()=>[],
  clipboardProbe:async()=>false,hasSlideClipboard:async()=>false,
  getAiSettings:async()=>settings,getRecentFiles:async()=>[],getTransition:async()=>'none',getAnimations:async()=>[],
  getComments:async()=>[],getSections:async()=>[],getNotes:async(index:number)=>opened?getSlideNotes(opened.archive,opened.deck.slides[index].path):'',
  setNotes:async(op:any)=>{await apply([{op:'setNotes',target:{slide:op.slideIndex},text:op.text}],op.slideIndex);return true},
  isDirty:async()=>dirty,save,saveAs:save,
  undo:async()=>{if(!opened||!undoStack.length)return null;redoStack.push(await savePptx(opened));opened=await openPptx(undoStack.pop()!);dirty=true;history();return allSlides()},
  redo:async()=>{if(!opened||!redoStack.length)return null;undoStack.push(await savePptx(opened));opened=await openPptx(redoStack.pop()!);dirty=true;history();return allSlides()},
  editText:async(op:any)=>apply([{op:'setText',...target(op),paragraphs:op.paragraphs}],op.slideIndex),
  setElementFont:async(op:any)=>apply(op.sourceIds.map((id:string)=>({op:'setFont',...target({...op,sourceId:id}),font:op})),op.slideIndex),
  setElementParagraphFormat:async(op:any)=>apply(op.sourceIds.map((id:string)=>({op:'setParagraphFormat',...target({...op,sourceId:id}),format:op})),op.slideIndex),
  deleteElement:async(op:any)=>apply([{op:'deleteElement',...target(op)}],op.slideIndex),
  deleteElements:async(op:any)=>apply(op.sourceIds.map((id:string)=>({op:'deleteElement',...target({...op,sourceId:id})})),op.slideIndex),
  editFill:async(op:any)=>{
    const g=op.fill?.gradient
    const path=g?.path??(g?.radial?'circle':undefined)
    const fill=g?{stops:g.stops?.length?g.stops:[{pos:0,color:g.from},{pos:1,color:g.to}],
      ...(path?{path,...(g.center?{fillTo:{l:g.center.x,t:g.center.y,r:1-g.center.x,b:1-g.center.y}}:{})}:{angle:Math.round((g.angleDeg??0)*60000)})}:op.fill??'none'
    return apply([{op:'setFill',...target(op),fill}],op.slideIndex)
  },
  editTransform:async(op:any)=>{
    const factor=opened!.deck.size.cx/op.fitWidthPx
    const box={x:Math.round(op.xPx*factor),y:Math.round(op.yPx*factor),cx:Math.round(op.wPx*factor),cy:Math.round(op.hPx*factor)}
    return apply([{op:'setTransform',...target(op),box,rotDeg:op.rotationDeg,resizeTableGrid:!op.groupId}],op.slideIndex)
  },
  getSlideSize:async()=>opened?.deck.size,
  setSlideSize:async(op:any)=>apply([{op:'setSlideSize',...op}],0,true),
  reportSelection:async()=>{},setAutoSavePref:async()=>{},setShowFullScreen:async()=>{},
}
// Optional event subscriptions preserve the upstream preload contract entirely
// inside the plugin. Unsupported native/cloud actions fail visibly.
;(window as any).slidesApi = new Proxy(api,{get(target,key:string){
  if(key in target)return target[key]
  if(key.startsWith('on'))return (fn:any)=>{if(!listeners.has(key))listeners.set(key,new Set());listeners.get(key)!.add(fn);return()=>listeners.get(key)?.delete(fn)}
  return async()=>{throw new Error(`此离线插件暂不支持此操作：${key}`)}
}})
// A browser test harness can select files through the standard file input.
const input=document.createElement('input');input.type='file';input.accept='.pptx';input.id='pptx-file-input';input.style.display='none';input.onchange=()=>{if(input.files?.[0])void load(input.files[0]).catch(showError)}
document.addEventListener('DOMContentLoaded',()=>document.body.append(input))
