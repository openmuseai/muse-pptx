/** Plugin-owned browser adapter for the unmodified GenOffice editor preload API.
 * The Host transfers standard MCP Apps tool input and download-file messages. All PPTX parsing/rendering/editing
 * stays in the upstream libraries bundled with this plugin.
 */
import { Buffer } from 'buffer'
import { App } from '@modelcontextprotocol/ext-apps'
import { openPptx, savePptx, createBlankPptx, getSlideNotes,reparseDeck,materializeSlide,commitSaved, type OpenedPptx } from '@genoffice/pptx-engine'
import { buildAllRenderSlides, rebuildSlide,makeMediaResolver } from '../../vendor/genoffice/apps/slides/src/main/session-state'
import { runTxn, type Op, type TxnRequest } from '@genoffice/pptx-ops'
import {fontsReady,fontMetrics} from './font-metrics'
import {presentation,pluginConfiguration} from './presentation'
import packageInfo from '../../package.json'
import {elementInsertOp,pickImageFile,imageSize} from './insert-bridge'
import {inkInsertOp} from './ink-bridge'
import {pictureFrame} from '../../vendor/genoffice/apps/slides/src/main/picture-frame'
import {imageDpiFromBytes} from '@genoffice/pptx-render'
import {tiffToPng} from '../../vendor/genoffice/apps/slides/src/main/tiff-decode'
import {createToolbarBridge,type ToolbarPort} from './toolbar-bridge'
import {pickBrowserFile,createBrowserServices} from './browser-services'
import {t} from '../../vendor/genoffice/apps/slides/src/renderer/i18n/locale'
import {confirmBrowserDialog,saveNameDialog} from './browser-dialogs'
import {z} from 'zod'

const mcpApp = window.parent !== window ? new App({name:'OpenMuse Slides',version:packageInfo.version},{tools:{listChanged:false}}) : null
if (mcpApp) {
  mcpApp.ontoolinput = ({arguments:args}) => {
    if(typeof args?.content === 'string') void load(new File([Buffer.from(args.content,'base64')], typeof args.name==='string'?args.name:presentation.limits.defaultFileName, {type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})).catch(showError)
  }
  mcpApp.onlisttools = async()=>({tools:agentTools})
  mcpApp.oncalltool = async params=>{
    const task=agentQueue.then(()=>callAgentTool(params.name,params.arguments??{}))
    agentQueue=task.catch(()=>{})
    try {return {content:[{type:'text',text:JSON.stringify(await task)}]}}
    catch(error){return {isError:true,content:[{type:'text',text:String(error)}]}}
  }
  mcpApp.onhostcontextchanged=()=>acceptHostContext(mcpApp.getHostContext())
}

function acceptHostContext(context:unknown){
  try {presentation.apply(context);if((context as any)?.platform==='mobile')document.documentElement.dataset.platform='mobile';else delete document.documentElement.dataset.platform}catch{showError(new PluginPresentationError('languageRejected'))}
}
const hostReady=mcpApp ? mcpApp.connect().then(()=>{acceptHostContext(mcpApp.getHostContext());publishContext()}).catch(showError) : Promise.resolve()
if(!mcpApp)presentation.apply({locale:navigator.language})
function requireFeature(feature:keyof typeof pluginConfiguration.features){if(!presentation.features[feature])throw new PluginPresentationError('featureDisabled')}
class PluginPresentationError extends Error {
  constructor(readonly code:string,readonly arguments_:Record<string,string|number>={}){super(presentation.integration(code,arguments_))}
}
presentation.subscribe(()=>{emit('onLanguageChanged',presentation.upstream);emit('onThemeChanged',presentation.theme)})

let opened: OpenedPptx | null = null
let fileName = presentation.limits.defaultFileName
let fitWidthPx = 1280
let dirty = false
let revision = 0
let agentQueue: Promise<unknown> = Promise.resolve()
let editorContext: {slideIndex:number,selectedIds:string[],beforeRun?:()=>Promise<void>,goTo?:(index:number)=>Promise<void>} = {slideIndex:0,selectedIds:[]}
export function setAgentEditorContext(value: typeof editorContext) {editorContext=value;publishContext()}
function publishContext() {
  document.documentElement.dataset.pptxRevision=String(revision)
  if(mcpApp?.getHostCapabilities()?.updateModelContext) void mcpApp.updateModelContext({structuredContent:{name:fileName,revision,dirty,slides:opened?.deck.slides.length??0,slideIndex:editorContext.slideIndex,selectedIds:editorContext.selectedIds}}).catch(()=>{})
}
export async function askAgent(text:string) {
  requireFeature('agent')
  if(!mcpApp?.getHostCapabilities()?.message){const error=new PluginPresentationError('agentUnavailable');showError(error);throw error}
  await editorContext.beforeRun?.()
  await mcpApp.sendMessage({role:'user',content:[{type:'text',text}]})
}
let pendingResolve: ((value: any) => void) | null = null
let firstOpen: any = null
const listeners = new Map<string, Set<(...args: any[]) => void>>()
const emit = (name: string, value: any) => listeners.get(name)?.forEach(fn => fn(value))
const undoStack: Uint8Array[] = []
const redoStack: Uint8Array[] = []
let historyBatch:{depth:number;snapshotTaken:boolean}|null=null
let mutationTail:Promise<unknown>=Promise.resolve()
let preview:{before:Uint8Array;redo:Uint8Array[];dirty:boolean;undoLength:number}|null=null
// The same Desktop session builder (including TIFF, themed SVG and JPEG
// orientation handling) produces the tree consumed by the original React UI.
const render = (index: number) => opened ? rebuildSlide({opened,fitWidthPx} as any,index) : null
const allSlides = () => buildAllRenderSlides(opened!,fitWidthPx)
const result = () => ({path:fileName,slides:allSlides(),size:opened!.deck.size})
const history = () => emit('onHistoryChanged',{canUndo:undoStack.length>0,canRedo:redoStack.length>0})
presentation.subscribe(()=>{
  const limit=presentation.limits.historyLimit
  if(undoStack.length>limit||redoStack.length>limit){
    undoStack.splice(0,Math.max(0,undoStack.length-limit));redoStack.splice(0,Math.max(0,redoStack.length-limit));history()
  }
})
async function load(file: File) {
  if (file.size > presentation.limits.maxFileBytes) throw new PluginPresentationError('fileTooLarge')
  await fontsReady()
  opened = await openPptx(new Uint8Array(await file.arrayBuffer()))
  fileName = file.name; dirty=false; undoStack.length=0;redoStack.length=0;historyBatch=null;preview=null
  api.resetSession?.()
  revision++;publishContext()
  firstOpen=result()
  if(pendingResolve) {pendingResolve(firstOpen);pendingResolve=null}
  emit('onOpened',firstOpen);history()
  document.documentElement.dataset.pptxReady=String(opened.deck.slides.length)
}
window.addEventListener('unhandledrejection',event=>showError(event.reason))
window.addEventListener('message',event=>{
  if (event.source===(mcpApp?window.parent:window) && event.data instanceof File) void load(event.data).catch(showError)
})
let errorBanner:HTMLDivElement|null=null
let errorMessage:{code:string,args:Record<string,string|number>}|null=null
let errorText:HTMLSpanElement|null=null
function clearError(){errorBanner?.remove();errorBanner=null;errorText=null;errorMessage=null}
function renderError(){if(errorText&&errorMessage)errorText.textContent=presentation.integration(errorMessage.code,errorMessage.args);const close=errorBanner?.querySelector('button');if(close)close.setAttribute('aria-label',presentation.integration('dismissError'))}
presentation.subscribe(renderError)
function showError(error: unknown) {
  errorMessage={code:error instanceof PluginPresentationError?error.code:'operationFailed',args:error instanceof PluginPresentationError?error.arguments_:{}}
  if(!errorBanner){errorBanner=document.createElement('div');errorBanner.setAttribute('role','alert');errorBanner.style.cssText='position:fixed;bottom:20px;inset-inline-start:20px;background:#fff2f2;color:#9b1c1c;padding:12px;z-index:99999';errorText=document.createElement('span');const close=document.createElement('button');close.type='button';close.textContent='×';close.style.cssText='margin-inline-start:10px;border:0;background:transparent;color:inherit;cursor:pointer';close.onclick=clearError;errorBanner.append(errorText,close);document.body.append(errorBanner)}
  renderError()
}

function serializeMutation<T>(action:()=>Promise<T>):Promise<T>{const next=mutationTail.then(action);mutationTail=next.catch(()=>{});return next}
const transact:ToolbarPort['transact']=(ops,options={})=>serializeMutation(async()=>{
  requireFeature('editing');if(!opened)throw new PluginPresentationError('noPresentation')
  if(!ops.length)throw Error('No edit operations')
  if(options.expectedRevision!==undefined&&options.expectedRevision!==revision)throw Error('Document revision conflict')
  const source=opened,before=await savePptx(source),beforeRevision=revision
  if(opened!==source||revision!==beforeRevision)throw Error('Document changed while preparing edit')
  const previousRedo=[...redoStack],previousDirty=dirty
  try{
    if(options.replaceLast){if(!undoStack.length)throw Error('Paste options expired');opened=await openPptx(options.replaceLast)}
    options.prepare?.()
    const active=opened!
    const outcome=runTxn(active,{ops,...(options.parts?{parts:options.parts}:{})})
    if(!outcome.applied)throw Error(outcome.failures?.map(f=>f.error).join('; ')||'Edit failed')
    const continuingPreview=!!preview
    if(options.preview&&!preview)preview={before,redo:previousRedo,dirty:previousDirty,undoLength:undoStack.length}
    if(!options.replaceLast&&!continuingPreview&&!historyBatch?.snapshotTaken){undoStack.push(before);if(historyBatch)historyBatch.snapshotTaken=true}
    if(!options.preview)preview=null
    if(options.reparse)opened=reparseDeck(active)
    if(options.materialize!==undefined)materializeSlide(active,options.materialize)
    while(undoStack.length>presentation.limits.historyLimit)undoStack.shift()
    redoStack.length=0;dirty=true;revision++;clearError();history();publishContext();return outcome
  }catch(error){opened=await openPptx(before);emit('onDeckChanged',{slides:allSlides()});throw error}
})
async function apply(ops: Op[], index: number, all=false, expectedRevision?:number,creation=false) {
  const outcome=await transact(ops,{expectedRevision})
  if(creation){const sourceId=outcome.records?.flatMap(record=>record.created??[])[0];if(!sourceId)throw Error('Insertion produced no element');return {slide:render(index),sourceId}}
  return all?allSlides():render(index)
}
const target=(op:any)=>({target:{slide:op.slideIndex,el:op.sourceId},...(op.groupId?{group:op.groupId}:{})})
async function save(exportName=fileName) {
  requireFeature('export')
  if(!opened)return {ok:false}
  await mutationTail
  const savedRevision=revision
  const bytes=await savePptx(opened)
  const file=new File([bytes as BlobPart],exportName,{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})
  if(mcpApp) {
    const response=await mcpApp.downloadFile({contents:[{type:'resource',resource:{uri:'file:///'+encodeURIComponent(exportName),mimeType:file.type,blob:Buffer.from(bytes).toString('base64')}}]})
    if(response.isError) return {ok:false}
  } else {
    const a=document.createElement('a');a.href=URL.createObjectURL(file);a.download=exportName;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)
  }
  if(savedRevision===revision){commitSaved(opened);dirty=false}
  fileName=exportName
  emit('onDeckChanged',{slides:allSlides()});publishContext()
  return {ok:true,path:fileName,slides:allSlides()}
}
const settings={provider:'custom',providers:{custom:{apiKey:'',baseUrl:'',model:''}}}
const api: Record<string, (...args:any[])=>any>={
  getLanguage:async()=>{let timer:ReturnType<typeof setTimeout>;await Promise.race([hostReady,new Promise<void>(resolve=>{timer=setTimeout(resolve,presentation.limits.hostInitializationTimeoutMs)})]);clearTimeout(timer!);return presentation.upstream},getTheme:async()=>presentation.theme,getAutoSaveDefault:async()=>false,
  getAiPanelPrefs:async()=>({fontSize:14,spellcheck:false}),setAiPanelPrefs:async(p:any)=>p,
  consumePendingOpen:async(width:number)=>{fitWidthPx=width;return opened?result():new Promise(resolve=>pendingResolve=resolve)},
  consumeHeadlessExport:async()=>null,
  openPptx:async()=>{const file=await pickBrowserFile('.pptx');if(!file)return null;await load(file);return result()},
  newBlank:async(width:number)=>{requireFeature('editing');fitWidthPx=width;await load(new File([await createBlankPptx() as BlobPart],presentation.limits.defaultFileName));return result()},
  beginHistoryBatch:async()=>{requireFeature('editing');if(!opened)return false;if(historyBatch)historyBatch.depth++;else historyBatch={depth:1,snapshotTaken:false};return true},
  endHistoryBatch:async()=>{if(!historyBatch)return false;if(--historyBatch.depth===0)historyBatch=null;return true},
  addElement:async(op:any)=>{requireFeature('editing');if(!opened)throw new PluginPresentationError('noPresentation');return apply([elementInsertOp(opened.deck.size,op) as Op],op.slideIndex,false,undefined,true)},
  addInk:async(op:any)=>{requireFeature('editing');if(!opened)throw new PluginPresentationError('noPresentation');const operation=inkInsertOp(opened.deck.size,op);fitWidthPx=op.fitWidthPx;return apply([operation],op.slideIndex,false,undefined,true)},
  pickPictureFile:async()=>{requireFeature('editing');const file=await pickImageFile();if(!file)return null;if(file.size>16*1024*1024)throw new PluginPresentationError('fileTooLarge');return {base64:Buffer.from(await file.arrayBuffer()).toString('base64'),ext:file.name.split('.').pop()!.toLowerCase()}},
  insertImage:async(index:number)=>{
    requireFeature('editing');if(!opened)throw new PluginPresentationError('noPresentation')
    const source=opened,expected=revision,file=await pickImageFile();if(!file)return null
    if(file.size>16*1024*1024)throw new PluginPresentationError('fileTooLarge')
    const ext=file.name.split('.').pop()!.toLowerCase()
    if(!['png','jpg','jpeg','gif','bmp','webp','tif','tiff'].includes(ext))return {error:'unsupported',ext}
    const bytes=new Uint8Array(await file.arrayBuffer())
    let natural:{width:number;height:number};try{if(ext==='tif'||ext==='tiff'){const decoded=tiffToPng(bytes);if(!decoded)return {error:'unsupported',ext};natural={width:decoded.width,height:decoded.height}}else natural=await imageSize(file)}catch{return {error:'unsupported',ext}}
    if(opened!==source||revision!==expected)throw new PluginPresentationError('operationFailed')
    const size=source.deck.size
    return apply([{op:'addPicture',target:{slide:index},bytes,ext,offset:pictureFrame({width:size.cx,height:size.cy},natural,imageDpiFromBytes(bytes),96)}],index,false,expected,true)
  },
  flipElements:async(op:any)=>apply([{op:'flipElements',target:{slide:op.slideIndex},els:op.sourceIds,axis:op.axis}],op.slideIndex),
  getLayouts:async()=>({layouts:[],size:opened?.deck.size??{cx:12192000,cy:6858000}}),
  privateFontFaces:async()=>[],privateFontData:async()=>null,fontMissing:async()=>[],fontCatalog:async()=>[],
  clipboardProbe:async()=>false,hasSlideClipboard:async()=>false,
  getAiSettings:async()=>settings,getRecentFiles:async()=>[],getTransition:async()=>'none',getAnimations:async()=>[],
  getComments:async()=>[],getSections:async()=>[],getNotes:async(index:number)=>opened?getSlideNotes(opened.archive,opened.deck.slides[index].path):'',
  setNotes:async(op:any)=>{await apply([{op:'setNotes',target:{slide:op.slideIndex},text:op.text}],op.slideIndex);return true},
  isDirty:async()=>dirty,save,saveAs:async(defaultName:string)=>{const stem=(defaultName||fileName).replace(/\.pptx$/i,'');const name=await saveNameDialog(stem+'-copy.pptx');return name?save(name):{ok:false}},
  undo:async()=>serializeMutation(async()=>{requireFeature('editing');if(!opened||!undoStack.length)return null;historyBatch=null;preview=null;redoStack.push(await savePptx(opened));opened=await openPptx(undoStack.pop()!);dirty=true;revision++;history();publishContext();return allSlides()}),
  redo:async()=>serializeMutation(async()=>{requireFeature('editing');if(!opened||!redoStack.length)return null;historyBatch=null;preview=null;undoStack.push(await savePptx(opened));while(undoStack.length>presentation.limits.historyLimit)undoStack.shift();opened=await openPptx(redoStack.pop()!);dirty=true;revision++;history();publishContext();return allSlides()}),
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
    await transact([{op:'setTransform',...target(op),...(op.groupId?{absBox:box}:{box}),rotDeg:op.rotationDeg,resizeTableGrid:!op.groupId}],{preview:op.preview});return render(op.slideIndex)
  },
  reportSelection:async()=>{},setAutoSavePref:async()=>{},setShowFullScreen:async()=>{},
}
Object.assign(api,createToolbarBridge({
  document:()=>{if(!opened)throw new PluginPresentationError('noPresentation');return opened},revision:()=>revision,
  width:(value?:number)=>{if(value!==undefined){if(!Number.isFinite(value)||value<=0)throw Error('Invalid viewport');fitWidthPx=value}return fitWidthPx},
  render,renderAll:allSlides,transact,
  cancelPreview:index=>serializeMutation(async()=>{if(!preview)return render(index);const before=preview;preview=null;opened=await openPptx(before.before);undoStack.splice(before.undoLength);redoStack.splice(0,redoStack.length,...before.redo);dirty=before.dirty;revision++;history();publishContext();return render(index)}),
  pickFile:pickBrowserFile,imageFile:async()=>{const file=await pickImageFile();if(!file)return null;if(file.size>16*1024*1024)throw new PluginPresentationError('fileTooLarge');return {bytes:new Uint8Array(await file.arrayBuffer()),ext:file.name.split('.').pop()!.toLowerCase()}},
  confirm:confirmBrowserDialog,translate:key=>key.startsWith('integration.')?presentation.integration(key.slice(12)):t(key as any),metrics:fontMetrics,media:makeMediaResolver,
}),createBrowserServices({download:async(name,mime,bytes)=>{
  requireFeature('export')
  if(mcpApp){const r=await mcpApp.downloadFile({contents:[{type:'resource',resource:{uri:'file:///'+encodeURIComponent(name),mimeType:mime,blob:Buffer.from(bytes).toString('base64')}}]});return !r.isError}
  const url=URL.createObjectURL(new Blob([bytes as BlobPart],{type:mime}));const link=document.createElement('a');link.href=url;link.download=name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);return true
},printDocument:async(html,widthMils,heightMils)=>{
  requireFeature('export')
  if(!mcpApp?.getHostCapabilities()?.experimental?.['io.openmuse/browser-print'])return null
  const result=await mcpApp.request({method:'io.openmuse/browser-print',params:{html,title:fileName,widthMils,heightMils}},z.object({dialogOpened:z.boolean()}),{timeout:35000})
  return {ok:result.dialogOpened}
},emit,language:()=>presentation.upstream,theme:()=>presentation.theme}))
// Optional event subscriptions preserve the upstream preload contract entirely
// inside the plugin. Unsupported native/cloud actions fail visibly.
;(window as any).slidesApi = new Proxy(api,{get(target,key:string){
  if(key in target)return target[key]
  if(key.startsWith('on'))return (fn:any)=>{if(!listeners.has(key))listeners.set(key,new Set());listeners.get(key)!.add(fn);return()=>listeners.get(key)?.delete(fn)}
  return async()=>{throw new PluginPresentationError('unsupportedOperation',{operation:key})}
}})
// A browser test harness can select files through the standard file input.
const input=document.createElement('input');input.type='file';input.accept='.pptx';input.id='pptx-file-input';input.style.display='none';input.onchange=()=>{if(input.files?.[0])void load(input.files[0]).catch(showError)}
document.addEventListener('DOMContentLoaded',()=>document.body.append(input))

const integer={type:'integer',minimum:0}
const identity={slideIndex:integer,elementId:{type:'string',minLength:1}}
const expected={expectedRevision:integer}
function tool(name:string,description:string,properties:any,required:string[],readOnlyHint=false){
  return {name,description,inputSchema:{type:'object' as const,properties,required,additionalProperties:false},annotations:{readOnlyHint}}
}
const agentTools=[
  tool('pptx_get_state','OpenMuse live slides v1: read current page and document revision.',{},[],true),
  tool('pptx_read_page','Read actual live text and speaker notes for a zero-based slide; no file snapshot.',{slideIndex:integer},['slideIndex'],true),
  tool('pptx_go_to_slide','Navigate the original editor without changing the document.',{slideIndex:integer},['slideIndex']),
  tool('pptx_read_deck','Read live document revision and editable element IDs/text. Optionally read one zero-based slide. Never guess IDs. Read again after a revision conflict.',{slideIndex:integer},[],true),
  tool('pptx_edit_text','Replace text of one element, preserving original formatting. Zero-based slideIndex and exact elementId from pptx_read_deck. Revision must match; uses the same undoable transaction as manual editing.',{...identity,...expected,text:{type:'string',maxLength:50000}},['slideIndex','elementId','expectedRevision','text']),
  tool('pptx_set_font','Change text color/size/style of an element. Colors are six hexadecimal digits without #; sizes in points.',{...identity,...expected,font:{type:'object',properties:{color:{type:'string',pattern:'^[0-9A-Fa-f]{6}$'},fontSizePt:{type:'number',minimum:1,maximum:400},bold:{type:'boolean'},italic:{type:'boolean'}},additionalProperties:false}},['slideIndex','elementId','expectedRevision','font']),
  tool('pptx_set_notes','Set speaker notes of one zero-based slide.',{slideIndex:integer,...expected,text:{type:'string',maxLength:50000}},['slideIndex','expectedRevision','text']),
  tool('pptx_undo','Undo the last edit in this live editor.',expected,['expectedRevision']),
  tool('pptx_save','Save the current document through standard MCP Apps download-file to its Host-authorized destination. Call after successful edits. Reports failure if Host did not save.',expected,['expectedRevision']),
]
function checkedSlide(value:unknown) {
  if(!opened||!Number.isInteger(value)||Number(value)<0||Number(value)>=opened.deck.slides.length)throw new Error('Invalid slideIndex')
  return Number(value)
}
function elementSummary(el:any):any {
  const paragraphs=el.text?.paragraphs??el.textBody?.paragraphs??el.paragraphs
  return {elementId:el.id,type:el.type,...(paragraphs?{text:paragraphs.map((p:any)=>p.runs?.map((r:any)=>r.text??'').join('')??'').join('\n')}:{ }),...(el.children?{children:el.children.map(elementSummary)}:{})}
}
async function callAgentTool(name:string,args:Record<string,any>) {
  if(!opened)throw new Error('No presentation open')
  if(!agentTools.some(tool=>tool.name===name))throw new Error('Unknown tool')
  if(name==='pptx_get_state')return {name:fileName,revision,dirty,pageCount:opened.deck.slides.length,slideIndex:editorContext.slideIndex}
  if(name==='pptx_read_page') {
    const index=checkedSlide(args.slideIndex)
    const elements=opened.deck.slides[index].elements.map(elementSummary)
    const texts=(items:any[]):string[]=>items.flatMap(e=>[...(e.text?[e.text]:[]),...texts(e.children??[])])
    return {name:fileName,revision,index,text:texts(elements).join('\n'),notes:getSlideNotes(opened.archive,opened.deck.slides[index].path)}
  }
  if(name==='pptx_go_to_slide') {
    const index=checkedSlide(args.slideIndex)
    if(!editorContext.goTo)throw Error('Navigation unavailable')
    await editorContext.goTo(index)
    return {slideIndex:editorContext.slideIndex,revision}
  }
  if(name==='pptx_read_deck') {
    const indices=args.slideIndex===undefined?opened.deck.slides.map((_,i)=>i):[checkedSlide(args.slideIndex)]
    return {name:fileName,revision,dirty,slideIndex:editorContext.slideIndex,selectedIds:editorContext.selectedIds,slides:indices.map(index=>({index,elements:opened!.deck.slides[index].elements.map(elementSummary)}))}
  }
  // Flush pending original UI drafts before checking the agent's revision.
  // A draft mutation therefore rejects a stale agent edit rather than losing it.
  await editorContext.beforeRun?.()
  if(!Number.isInteger(args.expectedRevision)||args.expectedRevision!==revision)throw new Error(`Revision conflict: expected ${args.expectedRevision}, actual ${revision}. Read the document again.`)
  if(name==='pptx_save') {
    const outcome=await save()
    if(!outcome.ok)throw new Error('Host did not save the document')
    return {revision,saved:true,name:fileName}
  }
  if(name==='pptx_undo') {
    if(!undoStack.length)throw new Error('No edit to undo')
    await api.undo();emit('onDeckChanged',{slides:allSlides()});return {revision,dirty}
  }
  const index=checkedSlide(args.slideIndex)
  let op:any
  if(name==='pptx_set_notes') {
    if(typeof args.text!=='string'||args.text.length>50000)throw new Error('Invalid notes')
    op={op:'setNotes',target:{slide:index},text:args.text}
  } else {
    if(typeof args.elementId!=='string'||!args.elementId)throw new Error('Invalid elementId')
    const target={slide:index,el:args.elementId}
    if(name==='pptx_edit_text') {
      if(typeof args.text!=='string'||args.text.length>50000)throw new Error('Invalid text')
      op={op:'setText',target,paragraphs:args.text.split('\n').map(text=>({runs:[{text}]}))}
    }else if(name==='pptx_set_font') {
      const font=args.font
      if(!font||typeof font!=='object'||Array.isArray(font)||Object.keys(font).some(key=>!['color','fontSizePt','bold','italic'].includes(key)))throw new Error('Invalid font patch')
      if(font.color!==undefined&&!/^[0-9a-f]{6}$/i.test(font.color))throw new Error('Invalid color')
      if(font.fontSizePt!==undefined&&(!Number.isFinite(font.fontSizePt)||font.fontSizePt<1||font.fontSizePt>400))throw new Error('Invalid font size')
      if(['bold','italic'].some(key=>font[key]!==undefined&&typeof font[key]!=='boolean'))throw new Error('Invalid font style')
      op={op:'setFont',target,font}
    }else throw new Error('Unknown tool')
  }
  await apply([op],index,false,args.expectedRevision)
  emit('onDeckChanged',{slides:allSlides()})
  return {revision,dirty,slideIndex:index}
}
