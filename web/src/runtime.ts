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

const mcpApp = window.parent !== window ? new App({name:'GenOffice PPTX',version:'0.5.0'},{tools:{listChanged:false}}) : null
if (mcpApp) {
  mcpApp.ontoolinput = ({arguments:args}) => {
    if(typeof args?.content === 'string') void load(new File([Buffer.from(args.content,'base64')], typeof args.name==='string'?args.name:'Presentation.pptx', {type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})).catch(showError)
  }
  mcpApp.onlisttools = async()=>({tools:agentTools})
  mcpApp.oncalltool = async params=>{
    const task=agentQueue.then(()=>callAgentTool(params.name,params.arguments??{}))
    agentQueue=task.catch(()=>{})
    try {return {content:[{type:'text',text:JSON.stringify(await task)}]}}
    catch(error){return {isError:true,content:[{type:'text',text:String(error)}]}}
  }
  void mcpApp.connect().then(()=>publishContext()).catch(showError)
}
let opened: OpenedPptx | null = null
let fileName = 'Presentation.pptx'
let fitWidthPx = 1280
let dirty = false
let revision = 0
let agentQueue: Promise<unknown> = Promise.resolve()
let editorContext: {slideIndex:number,selectedIds:string[],beforeRun?:()=>Promise<void>} = {slideIndex:0,selectedIds:[]}
export function setAgentEditorContext(value: typeof editorContext) {editorContext=value;publishContext()}
function publishContext() {
  document.documentElement.dataset.pptxRevision=String(revision)
  if(mcpApp?.getHostCapabilities()?.updateModelContext) void mcpApp.updateModelContext({structuredContent:{name:fileName,revision,dirty,slides:opened?.deck.slides.length??0,slideIndex:editorContext.slideIndex,selectedIds:editorContext.selectedIds}}).catch(()=>{})
}
export async function askAgent(text:string) {
  if(!mcpApp?.getHostCapabilities()?.message)throw new Error('当前 Host 尚未连接 Agent 插件，请在 Host 的 Agent 面板中操作。')
  await editorContext.beforeRun?.()
  await mcpApp.sendMessage({role:'user',content:[{type:'text',text}]})
}
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
  revision++;publishContext()
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
async function apply(ops: Op[], index: number, all=false, expectedRevision?:number) {
  if(!opened) throw new Error('No presentation open')
  const before=await savePptx(opened)
  if(expectedRevision!==undefined&&expectedRevision!==revision)throw new Error(`Revision conflict: expected ${expectedRevision}, actual ${revision}. Read the document again.`)
  const outcome=runTxn(opened,{ops})
  if(!outcome.applied) throw new Error(outcome.failures?.map(f=>f.error).join('; ')||'Edit failed')
  undoStack.push(before);if(undoStack.length>30)undoStack.shift();redoStack.length=0;dirty=true;history()
  revision++;publishContext()
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
  emit('onDeckChanged',{slides:allSlides()});publishContext()
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
  undo:async()=>{if(!opened||!undoStack.length)return null;redoStack.push(await savePptx(opened));opened=await openPptx(undoStack.pop()!);dirty=true;revision++;history();publishContext();return allSlides()},
  redo:async()=>{if(!opened||!redoStack.length)return null;undoStack.push(await savePptx(opened));opened=await openPptx(redoStack.pop()!);dirty=true;revision++;history();publishContext();return allSlides()},
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

const integer={type:'integer',minimum:0}
const identity={slideIndex:integer,elementId:{type:'string',minLength:1}}
const expected={expectedRevision:integer}
function tool(name:string,description:string,properties:any,required:string[],readOnlyHint=false){
  return {name,description,inputSchema:{type:'object' as const,properties,required,additionalProperties:false},annotations:{readOnlyHint}}
}
const agentTools=[
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
