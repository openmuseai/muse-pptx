/** Browser translation of the upstream SlidesApi. Document rules stay in
 * GenOffice's canonical ops, not in the Host or a second OOXML implementation.
 * Contract reference: vendor/genoffice/apps/slides/src/main/slides-main.ts. */
import {Buffer} from 'buffer'
import {
  listSlideLayouts,builtinLayoutInfos,shouldOfferBuiltinLayouts,ensureBuiltinLayout,BUILTIN_LAYOUT_PREFIX,
  copySlide,copyElementData,getSections,normalizeSections,getSlideComments,getSlideTransition,
  getSlideAnimations,elementSpid,getSlideAdvanceTime,narrationSpids,getChartElementData,
  getElementLink,getSlideLinks,getRunLinks,readHeaderFooter,parseMasterPart,listMasterParts,
  savePptx,parseTheme,type OpenedPptx,type Slide,
} from '@genoffice/pptx-engine'
import {type Op,type TxnResult,mapScriptOps,runTxn} from '@genoffice/pptx-ops'
import {buildRenderSlide} from '@genoffice/pptx-render'
import {planSlideDuplicates,planSlideMoves} from '../../vendor/genoffice/apps/slides/src/shared/slide-selection'
import {newPasteCascade,pageKey,pasteShiftPx,recordPaste} from '../../vendor/genoffice/apps/slides/src/main/paste-cascade'
import {pictureFrame} from '../../vendor/genoffice/apps/slides/src/main/picture-frame'
import {imageDpiFromBytes} from '@genoffice/pptx-render'
import {videoFrame,audioFrame,videoSize} from '../../vendor/genoffice/apps/slides/src/main/video-size'

export interface ToolbarPort {
  document():OpenedPptx
  revision():number
  width(value?:number):number
  render(index:number):any
  renderAll():any[]
  transact(ops:Op[],options?:{prepare?:()=>void;reparse?:boolean;materialize?:number;parts?:Map<string,Slide>;preview?:boolean;expectedRevision?:number;replaceLast?:Uint8Array}):Promise<TxnResult>
  cancelPreview(index:number):Promise<any>
  pickFile(accept:string):Promise<File|null>
  imageFile():Promise<{bytes:Uint8Array;ext:string}|null>
  confirm(message:string):Promise<boolean>
  translate(key:string):string
  metrics:any
  media(opened:OpenedPptx):(path:string)=>string|undefined
}
export function createToolbarBridge(p:ToolbarPort){
  let slidesClipboard:{bundles:any[];pngs?:string[]}|null=null
  let elementsClipboard:{items:any[];token:string;cascade:ReturnType<typeof newPasteCascade>;image?:string}|null=null
  let clipboardKind:'slide'|'internal'|null=null
  let lastPaste:{revision:number;before:Uint8Array;afterIndex:number}|null=null
  let master:{partPath:string;slide:Slide}|null=null
  const doc=()=>p.document()
  const slide=(i:number)=>{if(!Number.isInteger(i)||!doc().deck.slides[i])throw Error('Invalid slide index');return doc().deck.slides[i]!}
  const target=(o:any)=>({target:{slide:o.slideIndex,el:o.sourceId},...(o.groupId?{group:o.groupId}:{})})
  const emu=(width?:number)=>{const w=width??p.width();if(!Number.isFinite(w)||w<=0)throw Error('Invalid viewport');return (n:number)=>{if(!Number.isFinite(n))throw Error('Invalid geometry');return Math.round(n*doc().deck.size.cx/w)}}
  const box=(o:any)=>{const e=emu(o.fitWidthPx);return {x:e(o.xPx),y:e(o.yPx),cx:e(o.wPx),cy:e(o.hPx)}}
  const run=async(ops:Op[],o?:Parameters<ToolbarPort['transact']>[1])=>{if(!ops.length)throw Error('No operations');return p.transact(ops,o)}
  const edit=async(ops:Op[],i:number,options?:Parameters<ToolbarPort['transact']>[1])=>{await run(ops,options);return p.render(i)}
  const whole=async(ops:Op[],width?:number,options?:Parameters<ToolbarPort['transact']>[1])=>{await run(ops,options);if(width)p.width(width);return p.renderAll()}
  const created=async(ops:Op[],o:any)=>{const r=await run(ops);p.width(o.fitWidthPx);return {slide:p.render(o.slideIndex),sourceId:r.records?.flatMap(r=>r.created??[])[0]}}
  const layout=(value:string|undefined)=>value?.startsWith(BUILTIN_LAYOUT_PREFIX)?ensureBuiltinLayout(doc().archive,doc().deck.size,value.slice(BUILTIN_LAYOUT_PREFIX.length))??undefined:value
  const indexes=(index:number)=>index===-1?doc().deck.slides.map((_,i)=>i):(slide(index),[index])
  const section=async(op:Op)=>{await run([op]);return getSections(doc())}
  const stroke=(s:any)=>s?{color:s.color,widthEmu:Math.round(s.widthPt*12700),...(s.dash?{dash:s.dash}:{}),...(s.cap?{cap:s.cap}:{}),...(s.join?{join:s.join}:{}),...(s.compound?{compound:s.compound}:{}),...(s.gradient?{gradient:{stops:s.gradient.stops,angle:Math.round(s.gradient.angleDeg*60000)}}:{})}:null
  const fill=(v:any)=>{if(typeof v==='string')return v;const g=v?.gradient;if(!g)return v??'none';const path=g.path??(g.radial?'circle':undefined);return {stops:g.stops?.length?g.stops:[{pos:0,color:g.from},{pos:1,color:g.to}],...(path?{path,...(g.center?{fillTo:{l:g.center.x,t:g.center.y,r:1-g.center.x,b:1-g.center.y}}:{})}:{angle:Math.round((g.angleDeg??0)*60000)})}}
  const masterRender=()=>master?buildRenderSlide(master.slide,doc().deck.size,{fitWidthPx:p.width(),media:p.media(doc()),metrics:p.metrics}):null
  const masterEdit=async(op:Op)=>{if(!master)return null;await run([{...op,target:{part:master.partPath,el:(op.target as any)?.el}}],{parts:new Map([[master.partPath,master.slide]])});return masterRender()}
  const api:Record<string,(...args:any[])=>any>={
    // Home: slide structure, layout, sections and clipboard.
    getLayouts:async()=>{const layouts=listSlideLayouts(doc().archive);if(shouldOfferBuiltinLayouts(layouts))layouts.push(...builtinLayoutInfos(doc().deck.size,new Set(layouts.map(l=>l.name))));return {layouts,size:{...doc().deck.size}}},
    addBlankSlide:async(o:any)=>({slides:await whole([{op:'addBlankSlide',target:{slide:o.sourceIndex}},...(o.before?[{op:'moveSlide',target:{slide:o.sourceIndex},to:o.sourceIndex+1}]:[])],o.fitWidthPx),index:o.before?o.sourceIndex:o.sourceIndex+1}),
    addSlide:async(o:any)=>({slides:await whole([{op:'duplicateSlide',target:{slide:o.sourceIndex},clearText:o.clearText}],o.fitWidthPx),index:o.sourceIndex+1}),
    addSlideWithLayout:async(o:any)=>{const op:Op={op:'addSlideWithLayout',target:{slide:o.sourceIndex},layoutPath:o.layoutPath};return {slides:await whole([op],o.fitWidthPx,{prepare:()=>{op.layoutPath=layout(o.layoutPath);if(!op.layoutPath)throw Error('Layout not found')}}),index:o.sourceIndex+1}},
    setSlideLayout:async(o:any)=>{const op:Op={op:'setSlideLayout',target:{slide:o.slideIndex},...(o.layoutPath?{layoutPath:o.layoutPath}:{})};return edit([op],o.slideIndex,{prepare:()=>{if(o.layoutPath){op.layoutPath=layout(o.layoutPath);if(!op.layoutPath)throw Error('Layout not found')}}})},
    deleteSlide:async(i:number)=>api.deleteSlides({slideIndexes:[i]}),
    deleteSlides:async(o:any)=>{const ids=[...new Set<number>(o.slideIndexes)].sort((a,b)=>b-a);if(!ids.length||ids.length>=doc().deck.slides.length)return null;return whole(ids.map(i=>({op:'deleteSlide',target:{slide:i}})))},
    duplicateSlides:async(o:any)=>{const ops=planSlideDuplicates(o.slideIndexes).map(s=>'duplicate'in s?{op:'duplicateSlide',target:{slide:s.duplicate}}:{op:'moveSlide',target:{slide:s.from},to:s.to});if(!ops.length)return null;return {slides:await whole(ops,o.fitWidthPx),index:Math.max(...o.slideIndexes)+1}},
    moveSlide:async(o:any)=>({slides:await whole([{op:'moveSlide',target:{slide:o.fromIndex},to:o.toIndex}]),sections:getSections(doc())}),
    moveSlides:async(o:any)=>{const ops=planSlideMoves(doc().deck.slides.length,o.slideIndexes,o.insertAt).map(s=>({op:'moveSlide',target:{slide:s.from},to:s.to}));if(!ops.length)return null;return {slides:await whole(ops),sections:getSections(doc())}},
    getSections:async()=>getSections(doc()),setSections:async(sections:any[])=>section({op:'setSections',sections}),
    addSection:async(o:any)=>section({op:'addSection',atSlideIndex:o.atSlideIndex,name:o.name}),
    renameSection:async(o:any)=>section({op:'renameSection',id:o.id,name:o.name}),removeSection:async(o:any)=>section({op:'removeSection',id:o.id}),
    moveSection:async(o:any)=>{await section({op:'moveSection',id:o.id,dir:o.dir});return {slides:p.renderAll(),sections:getSections(doc())}},
    removeSectionSlides:async(o:any)=>{const {lead,sections}=normalizeSections(getSections(doc()),doc().deck.slides.length);const ids=o.id==null?lead:sections.find(s=>s.id===o.id)?.slideIndices;if(!ids?.length||ids.length>=doc().deck.slides.length)return null;await run([...ids].reverse().map(i=>({op:'deleteSlide',target:{slide:i}})).concat(o.id==null?[]:[{op:'removeSection',id:o.id}] as any));return {slides:p.renderAll(),sections:getSections(doc())}},
    findReplace:async(o:any)=>{const r=await run([{op:'findReplace',find:o.find,replace:o.replace,matchCase:o.matchCase,firstOnly:o.firstOnly,slideIndex:o.slideIndex,elementId:o.elementId}]);return {count:(r.records?.[0]?.after as any)?.count??0,slides:p.renderAll()}},
    copySlides:async(o:any)=>{const bundles=o.slideIndexes.map((i:number)=>{slide(i);const bundle=copySlide(doc(),i);if(!bundle)throw Error('Cannot copy slide');return bundle});if(!bundles.length)return false;slidesClipboard={bundles,...(o.pngs?.length===bundles.length?{pngs:o.pngs}:{})};clipboardKind='slide';return true},
    hasSlideClipboard:async()=>clipboardKind==='slide'&&!!slidesClipboard,
    clipboardProbe:async()=>!!clipboardKind,
    clipboardExternal:async()=>clipboardKind?{kind:clipboardKind}:{kind:'none'},
    pasteSlide:async(o:any)=>{if(!slidesClipboard)return null;const {bundles,pngs}=slidesClipboard;if(o.mode==='picture'&&!pngs)return null;const before=await savePptx(doc()),expectedRevision=p.revision();const r=await run(bundles.map((bundle,j)=>({op:'pasteSlide',afterIndex:o.mode==='picture'?o.afterIndex:o.afterIndex+j,mode:o.mode,...(o.mode==='picture'?{png:pngs![j]}:{bundle})})),{expectedRevision});lastPaste={before,afterIndex:o.afterIndex,revision:p.revision()};p.width(o.fitWidthPx);return {slides:p.renderAll(),index:(r.records?.[0]?.after as any)?.index,count:o.mode==='picture'?0:bundles.length,sourceIds:r.records?.flatMap(r=>r.created??[])}},
    repasteSlide:async(o:any)=>{if(!lastPaste||!slidesClipboard||p.revision()!==lastPaste.revision)return null;const {bundles,pngs}=slidesClipboard;if(o.mode==='picture'&&!pngs)return null;const rec=lastPaste;const r=await run(bundles.map((bundle,j)=>({op:'pasteSlide',afterIndex:o.mode==='picture'?rec.afterIndex:rec.afterIndex+j,mode:o.mode,...(o.mode==='picture'?{png:pngs![j]}:{bundle})})),{expectedRevision:rec.revision,replaceLast:rec.before});rec.revision=p.revision();p.width(o.fitWidthPx);return {slides:p.renderAll(),index:(r.records?.[0]?.after as any)?.index,count:o.mode==='picture'?0:bundles.length,sourceIds:r.records?.flatMap(r=>r.created??[])}},
    copyElements:async(o:any)=>{const s=slide(o.slideIndex);const items=o.sourceIds.map((id:string)=>s.elements.find(e=>e.id===id)).filter(Boolean).map((e:any)=>copyElementData(doc(),s,e));if(items.length){elementsClipboard={items,token:o.clipboardToken??crypto.randomUUID(),cascade:newPasteCascade(o.cut?null:pageKey(0,o.slideIndex))};clipboardKind='internal'}return items.length},
    copyElementsImage:async(token:string,png:string)=>{if(elementsClipboard?.token!==token)return false;elementsClipboard.image=png;return true},
    pasteElements:async(o:any)=>{if(!elementsClipboard)return null;const key=pageKey(0,o.slideIndex),shift=emu(o.fitWidthPx)(pasteShiftPx(elementsClipboard.cascade,key));const r=await run([{op:'pasteElements',target:{slide:o.slideIndex},items:elementsClipboard.items,dx:shift,dy:shift}]);recordPaste(elementsClipboard.cascade,key);p.width(o.fitWidthPx);return {slide:p.render(o.slideIndex),sourceIds:r.records?.flatMap(r=>r.created??[])}},
    duplicateElements:async(o:any)=>{const s=slide(o.slideIndex),items=o.sourceIds.map((id:string)=>s.elements.find(e=>e.id===id)).filter(Boolean).map((e:any)=>copyElementData(doc(),s,e));const e=emu(o.fitWidthPx);const r=await run([{op:'pasteElements',target:{slide:o.slideIndex},items,dx:e(o.dxPx),dy:e(o.dyPx)}]);return {slide:p.render(o.slideIndex),sourceIds:r.records?.flatMap(r=>r.created??[])}},
    editStroke:async(o:any)=>edit([{op:'setStroke',...target(o),stroke:stroke(o.stroke)}],o.slideIndex),
    reorderElement:async(o:any)=>edit([{op:'reorderElement',...target(o),dir:o.dir}],o.slideIndex),
    editTransformMulti:async(o:any)=>edit(o.items.map((item:any)=>({op:'setTransform',...target({...o,...item}),box:box({...o,...item}),rotDeg:item.rotationDeg,resizeTableGrid:!item.groupId})),o.slideIndex),
    batchEditTransform:async(o:any)=>edit(o.items.map((item:any)=>({op:'setTransform',...target({...o,...item}),box:box({...o,...item}),rotDeg:item.rotationDeg})),o.slideIndex),
    cancelTransformPreview:async(o:any)=>p.cancelPreview(o.slideIndex),
    groupElements:async(o:any)=>{const r=await run([{op:'groupElements',target:{slide:o.slideIndex},els:o.sourceIds}]);return {slide:p.render(o.slideIndex),groupId:r.records?.[0]?.created?.[0]}},
    ungroupElement:async(o:any)=>edit([{op:'ungroupElement',...target(o)}],o.slideIndex),
    getRenderSlides:async()=>p.renderAll(),
    // Insert: native editable table/chart/SmartArt and image bytes.
    addTable:async(o:any)=>created([{op:'addTable',target:{slide:o.slideIndex},rows:o.rows,cols:o.cols,offset:{...box(o),...(o.rowHeightEmu>0?{cy:Math.round(o.rowHeightEmu)*o.rows}:{})},...(o.rowHeightEmu>0?{rowHeightsEmu:Array.from({length:o.rows},()=>Math.round(o.rowHeightEmu))}:{})}],o),
    addChart:async(o:any)=>created([{op:'addChart',target:{slide:o.slideIndex},kind:o.kind==='barH'?'bar':o.kind,...(o.kind==='barH'?{barDir:'bar'}:{}),title:o.title,categories:o.categories,series:o.series,offset:box(o)}],o),
    addSmartArt:async(o:any)=>created([{op:'addSmartArt',target:{slide:o.slideIndex},layout:o.layout,items:o.items,offset:box(o)}],o),
    addImageBytes:async(o:any)=>{const bytes=bytesFrom(o.base64,16*1024*1024);const offset=o.naturalPx?pictureFrame({width:doc().deck.size.cx,height:doc().deck.size.cy},o.naturalPx,imageDpiFromBytes(bytes),96,{...(o.centerPx?{center:{x:emu(o.fitWidthPx)(o.centerPx.x),y:emu(o.fitWidthPx)(o.centerPx.y)}}:{}),minEmu:emu(o.fitWidthPx)(24)}):box(o);return created([{op:'addPicture',target:{slide:o.slideIndex},bytes,ext:o.ext,offset,...(o.name?{name:o.name}:{})}],o)},
    replacePictureBytes:async(o:any)=>edit([{op:'replacePicture',...target(o),bytes:bytesFrom(o.base64,16*1024*1024),ext:o.ext}],o.slideIndex),
    setLink:async(o:any)=>{
      if(o.target?.kind==='url'){let url:URL;try{url=new URL(o.target.url)}catch{throw Error('Invalid link URL')}if(!['https:','http:','mailto:','tel:'].includes(url.protocol)||url.username||url.password)throw Error('Unsafe link URL')}
      // Canonical hyperlink surgery materializes the whole slide and assigns
      // fresh parser IDs. Its element structure is unchanged: retain the UI's
      // identity so the selected object can immediately be edited again.
      const remember=(elements:any[]):any[]=>elements.map(e=>({id:e.id,children:e.children?remember(e.children):undefined}))
      const ids=remember(slide(o.slideIndex).elements)
      await run([{op:'setLink',...target(o),link:o.target}])
      const restore=(elements:any[],saved:any[])=>elements.forEach((e,i)=>{e.id=saved[i].id;if(e.children&&saved[i].children)restore(e.children,saved[i].children)})
      restore(slide(o.slideIndex).elements,ids)
      return p.render(o.slideIndex)
    },
    getLink:async(i:number,id:string)=>getElementLink(doc(),i,id),
    getSlideLinks:async(i:number)=>getSlideLinks(doc(),i).map(({elementId,...r})=>({sourceId:elementId,...r})),
    getRunLinks:async(i:number)=>getRunLinks(doc(),i).map(({elementId,...r})=>({sourceId:elementId,...r})),
    applyHeaderFooter:async(o:any)=>whole([{op:'applyHeaderFooter',settings:{footer:o.footer??null,slideNum:!!o.slideNum,date:o.date??null,...(o.dateAuto?{dateAuto:true}:{})}}],o.fitWidthPx),
    getHeaderFooter:async(i:number)=>readHeaderFooter(slide(i)),
    addComment:async(o:any)=>{await run([{op:'addComment',target:{slide:o.slideIndex},author:'OpenMuse',text:o.text}]);return api.getComments(o.slideIndex)},
    deleteComment:async(o:any)=>{await run([{op:'deleteComment',target:{slide:o.slideIndex},authorId:o.authorId,idx:o.idx}]);return api.getComments(o.slideIndex)},
    getComments:async(i:number)=>getSlideComments(doc().archive,slide(i).path),
    insertMedia:async(i:number,kind:'audio'|'video',fitWidthPx:number)=>{const revision=p.revision(),file=await p.pickFile(kind==='audio'?'audio/*':'video/*');if(!file)return null;if(file.size>64*1024*1024)throw Error('Media exceeds 64 MiB');if(p.revision()!==revision)throw Error('Document changed while selecting media');return api.addMediaBytes({slideIndex:i,fitWidthPx,kind,base64:Buffer.from(await file.arrayBuffer()).toString('base64'),ext:file.name.split('.').pop()?.toLowerCase(),name:file.name})},
    addMediaBytes:async(o:any)=>{if('path'in o)throw Error('Native paths are not browser file grants');const bytes=bytesFrom(o.base64,64*1024*1024),size={width:doc().deck.size.cx,height:doc().deck.size.cy};const center=o.centerPx?{x:emu(o.fitWidthPx)(o.centerPx.x),y:emu(o.fitWidthPx)(o.centerPx.y)}:undefined;return created([{op:'addMedia',target:{slide:o.slideIndex},kind:o.kind,bytes,ext:o.ext,name:o.name,offset:o.kind==='audio'?audioFrame(size,center):videoFrame(size,o.natural??videoSize(bytes,o.ext),center)}],o)},
    getMediaData:async(i:number,id:string)=>{const el:any=slide(i).elements.find(e=>e.id===id),media=el?.media;if(!media)return null;if(media.external)return {kind:media.kind,dataUrl:media.target};const bytes=doc().archive.readBytes(media.target);if(!bytes)return null;const ext=media.target.split('.').pop()?.toLowerCase(),mime:Record<string,string>={mp3:'audio/mpeg',wav:'audio/wav',m4a:'audio/mp4',mp4:'video/mp4',webm:media.kind==='audio'?'audio/webm':'video/webm',ogg:'audio/ogg',mov:'video/quicktime'};return {kind:media.kind,dataUrl:`data:${mime[ext]??(media.kind==='audio'?'audio/mpeg':'video/mp4')};base64,${Buffer.from(bytes).toString('base64')}`}},
    pickModel3d:async()=>{const file=await p.pickFile('.glb');if(!file)return null;if(file.size>64*1024*1024)throw Error('Model exceeds 64 MiB');return {name:file.name,bytes:new Uint8Array(await file.arrayBuffer())}},
    addModel3d:async(o:any)=>{const s=doc().deck.size,cy=Math.round(s.cy*.5),cx=Math.round(cy*o.aspect);return created([{op:'addModel3d',target:{slide:o.slideIndex},bytes:o.bytes,ext:'glb',name:o.name,...(o.posterBase64?{poster:{bytes:bytesFrom(o.posterBase64,16*1024*1024),ext:'png'}}:{}),offset:{x:Math.round((s.cx-cx)/2),y:Math.round((s.cy-cy)/2),cx,cy}}],o)},
    getModel3dData:async(i:number,id:string)=>{const e:any=slide(i).elements.find(e=>e.id===id);return e?.model3d?.target?doc().archive.readBytes(e.model3d.target)??null:null},
    setModel3dView:async(o:any)=>edit([{op:'setModel3dView',...target(o),...o.rot,...(o.posterBase64?{poster:o.posterBase64,posterExt:'png'}:{})}],o.slideIndex),
    // Contextual format tabs retain the original interactions.
    editPictureOpacity:async(o:any)=>edit([{op:'setPictureOpacity',...target(o),opacity:o.opacity}],o.slideIndex),
    editPictureSrcRect:async(o:any)=>edit([{op:'setPictureSrcRect',...target(o),srcRect:o.srcRect,...(o.boxPx?{box:box({...o,xPx:o.boxPx.x,yPx:o.boxPx.y,wPx:o.boxPx.w,hPx:o.boxPx.h})}:{})}],o.slideIndex),
    changeShape:async(o:any)=>edit([{op:'setShapeGeometry',...target(o),prst:o.prst}],o.slideIndex),
    setShapeAdjust:async(o:any)=>edit([{op:'setShapeAdjust',...target(o),adjust:o.adjust}],o.slideIndex,{preview:o.preview}),
    setShapeGeometry:async(o:any)=>{const e=emu(o.fitWidthPx);return edit([{op:'setShapeCustomGeometry',...target(o),path:{w:e(o.pathPx.w),h:e(o.pathPx.h),cmds:o.pathPx.cmds.map((c:any)=>({op:c.op,pts:c.pts.map(e)}))}}],o.slideIndex,{preview:o.preview})},
    setTextAnchor:async(o:any)=>edit([{op:'setTextAnchor',...target(o),anchor:o.anchor}],o.slideIndex),
    setTextBodyProps:async(o:any)=>edit([{op:'setTextBodyProps',...target(o),props:o.props}],o.slideIndex),
    setEffects:async(o:any)=>edit([{op:'setEffects',...target(o),effects:o.effects}],o.slideIndex),
    editConnectorEndpoints:async(o:any)=>{const e=emu(o.fitWidthPx);return edit([{op:'setConnectorEndpoints',...target(o),p1:{x:e(o.x1Px),y:e(o.y1Px)},p2:{x:e(o.x2Px),y:e(o.y2Px)},start:o.start,end:o.end}],o.slideIndex)},
    editTableCell:async(o:any)=>edit([{op:'setTableCell',...target(o),row:o.row,col:o.col,paragraphs:o.paragraphs}],o.slideIndex),
    tableStructure:async(o:any)=>{const r=await run([{op:'tableStructure',...target(o),kind:o.kind,index:o.index,...(o.before?{before:true}:{})}]);return {slide:p.render(o.slideIndex),sourceId:(r.records?.[0]?.after as any)?.elementId}},
    tableMerge:async(o:any)=>{const r=await run([{op:'tableMerge',...target(o),kind:o.kind,row:o.row,col:o.col}]);return {slide:p.render(o.slideIndex),sourceId:(r.records?.[0]?.after as any)?.elementId}},
    setTableColWidth:async(o:any)=>edit([{op:'setTableColWidth',...target(o),col:o.col,wEmu:emu(o.fitWidthPx)(o.wPx)}],o.slideIndex),
    setTableRowHeight:async(o:any)=>edit([{op:'setTableRowHeight',...target(o),row:o.row,hEmu:emu(o.fitWidthPx)(o.hPx)}],o.slideIndex),
    setTableCellAnchor:async(o:any)=>edit([{op:'setTableCellAnchor',...target(o),row:o.row,col:o.col,anchor:o.anchor}],o.slideIndex),
    editTableStyle:async(o:any)=>{const idx=slide(o.slideIndex).elements.findIndex(e=>e.id===o.sourceId),{slideIndex,sourceId,...style}=o;await run([{op:'setTableStyle',target:{slide:slideIndex,el:sourceId},...style}],{materialize:slideIndex});return {slide:p.render(slideIndex),sourceId:slide(slideIndex).elements[idx]?.id}},
    getChartData:async(i:number,id:string)=>getChartElementData(slide(i),id),
    getChartColorSchemes:async()=>chartSchemes(doc()),
    editChart:async(o:any)=>{const s=slide(o.slideIndex),idx=s.elements.findIndex(e=>e.id===o.sourceId),el=s.elements[idx];if(el?.type==='chart'&&el.descr!=='aislides-chart'&&!await p.confirm(p.translate('integration.chartSimplifyBody')))return null;const {slideIndex,sourceId,...patch}=o;if(patch.kind==='barH'){patch.kind='bar';patch.barDir='bar'}if(patch.colorScheme)patch.colorScheme=chartSchemes(doc()).find(s=>s.key===patch.colorScheme)?.colors;await run([{op:'setChart',target:{slide:slideIndex,el:sourceId},patch}],{materialize:slideIndex});return {slide:p.render(slideIndex),sourceId:slide(slideIndex).elements[idx]?.id}},
    // Design, transitions, animations, show, review and view read real document data.
    getSlideSize:async()=>({...doc().deck.size}),
    setSlideSize:async(o:any)=>doc().deck.size.cx===o.cx&&doc().deck.size.cy===o.cy?p.renderAll():whole([{op:'setSlideSize',cx:o.cx,cy:o.cy}]),
    applyTheme:async(o:any)=>whole([{op:'applyTheme',name:o.name,colors:o.colors,...(o.majorFont?{majorFont:o.majorFont}:{}),...(o.minorFont?{minorFont:o.minorFont}:{})}],o.fitWidthPx,{reparse:true}),
    editBackground:async(o:any)=>{const ids=indexes(o.slideIndex);let common:any;if(o.kind==='image'){const revision=p.revision();let source;if(o.pick!==false){const f=await p.imageFile();if(!f)return null;if(p.revision()!==revision)throw Error('Document changed while selecting image');source=f}else{const bg=slide(o.sourceSlideIndex??o.slideIndex).background;if(bg?.type!=='image')return null;source={mediaPath:bg.mediaRef}}common={kind:'image',source,tile:o.mode==='tile'}}else common=o.kind==='solid'?{kind:'solid',color:o.color}:o.kind==='gradient'?{kind:'gradient',from:o.from,to:o.to,angleDeg:o.angleDeg,...(o.radial?{radial:true}:{})}:o.kind==='reset'?{kind:'reset'}:{kind:'graphics',hidden:o.hidden};return whole(ids.map(i=>({op:'setBackground',target:{slide:i},...common})),o.fitWidthPx)},
    editImageFill:async(o:any)=>{const revision=p.revision(),source=o.source?{bytes:bytesFrom(o.source.base64,16*1024*1024),ext:o.source.ext}:await p.imageFile();if(!source)return null;if(p.revision()!==revision)throw Error('Document changed while selecting image');return edit(o.targets.map((t:any)=>({op:'setImageFill',...target({...o,...t}),source,tile:o.mode==='tile'})),o.slideIndex)},
    getTransition:async(i:number)=>getSlideTransition(slide(i)),
    setTransition:async(o:any)=>{await run(indexes(o.slideIndex).map(i=>({op:'setTransition',target:{slide:i},kind:o.kind})));return true},
    getAnimations:async(i:number)=>{const s=slide(i);const bySpid=new Map(s.elements.map(e=>[elementSpid(e),e]));return getSlideAnimations(s).flatMap(a=>{const e=bySpid.get(a.spid);if(!e)return [];const {spid,...rest}=a;return [{...rest,sourceId:e.id,targetName:e.name||e.type}]})},
    setAnimations:async(o:any)=>{await run([{op:'setAnimations',target:{slide:o.slideIndex},items:o.items}]);return true},
    getShapeKeys:async(i:number)=>slide(i).elements.map(e=>({sourceId:e.id,spid:elementSpid(e),name:e.name??''})),
    getShowTiming:async(i:number)=>{const s=slide(i),narration=new Set(narrationSpids(s));return {advanceMs:getSlideAdvanceTime(s),narrationIds:s.elements.filter(e=>narration.has(elementSpid(e)??-1)).map(e=>e.id)}},
    setAdvanceTimes:async(o:any)=>{await run(o.times.map((t:any)=>({op:'setAdvanceTime',target:{slide:t.slideIndex},ms:t.ms})));return true},
    setSlideHidden:async(o:any)=>edit([{op:'setHidden',target:{slide:o.slideIndex},hidden:o.hidden}],o.slideIndex),
    setSlidesHidden:async(o:any)=>whole(o.slideIndexes.map((i:number)=>({op:'setHidden',target:{slide:i},hidden:o.hidden}))),
    saveRecording:async(o:any)=>whole(o.slides.flatMap((s:any)=>[{op:'addNarration',target:{slide:s.slideIndex},bytes:bytesFrom(s.wavBase64,64*1024*1024),ext:'wav',durationMs:s.durationMs},{op:'setAdvanceTime',target:{slide:s.slideIndex},ms:s.advanceMs}])),
    clearNarration:async(ids:number[])=>{const targets=ids.filter(i=>narrationSpids(slide(i)).length);return targets.length?whole(targets.map(i=>({op:'clearNarration',target:{slide:i}}))):p.renderAll()},
    masterEnter:async(width:number)=>{p.width(width);const items=listMasterParts(doc().archive).flatMap(part=>{const s=parseMasterPart(doc().archive,part.partPath);if(!s)return [];if(!master)master={partPath:part.partPath,slide:s};return [{...part,slide:buildRenderSlide(s,doc().deck.size,{fitWidthPx:width,media:p.media(doc()),metrics:p.metrics})}]});return items.length?{items}:null},
    masterOpen:async(partPath:string)=>{if(!listMasterParts(doc().archive).some(p=>p.partPath===partPath))throw Error('Unknown master part');const s=parseMasterPart(doc().archive,partPath);if(!s)return null;master={partPath,slide:s};return masterRender()},
    masterClose:async()=>{master=null;return p.renderAll()},
    masterEditText:async(o:any)=>masterEdit({op:'setText',target:{el:o.sourceId},paragraphs:o.paragraphs}),
    masterEditTransform:async(o:any)=>masterEdit({op:'setTransform',target:{el:o.sourceId},box:box(o),rotDeg:o.rotationDeg}),
    masterEditFill:async(o:any)=>masterEdit({op:'setFill',target:{el:o.sourceId},fill:fill(o.fill)}),
    masterEditStroke:async(o:any)=>masterEdit({op:'setStroke',target:{el:o.sourceId},stroke:stroke(o.stroke)}),
    masterDeleteElement:async(o:any)=>masterEdit({op:'deleteElement',target:{el:o.sourceId}}),
    applyTxn:async(o:any)=>{const r=o.dryRun?runTxn(doc(),{ops:o.ops,dryRun:true}):await run(o.ops);return {...r,records:r.records?.map(rec=>({op:rec.op.op,...(rec.op.target?.el?{target:rec.op.target.el}:{}),...(rec.created?{created:rec.created}:{})})),...(!o.dryRun&&r.applied?{slides:p.renderAll()}:{})}},
    applyEditScript:async(o:any)=>{const ops=mapScriptOps(doc(),o);if(!ops.length)return null;return {slide:await edit(ops,o.slideIndex)}},
  }
  api.resetSession=()=>{master=null;lastPaste=null}
  return api
}

function bytesFrom(base64:string,limit:number){
  if(typeof base64!=='string'||base64.length>Math.ceil(limit/3)*4||!/^[A-Za-z0-9+/]*={0,2}$/.test(base64))throw Error('Invalid encoded file')
  const bytes=new Uint8Array(Buffer.from(base64,'base64'));if(!bytes.length||bytes.length>limit)throw Error('Invalid file size');return bytes
}
function chartSchemes(opened:OpenedPptx){
  const chain=opened.deck.slides[0]?opened.archive.resolveSlideChain(opened.deck.slides[0].path):null
  const xml=chain?.themePath?opened.archive.readText(chain.themePath):null
  const colors=xml?parseTheme(xml).colors:{}
  let accents=['accent1','accent2','accent3','accent4','accent5','accent6'].map(k=>colors[k]).filter((c):c is string=>!!c)
  if(accents.length<3)accents=['#4472C4','#ED7D31','#A5A5A5','#FFC000','#5B9BD5','#70AD47']
  const mix=(hex:string,to:number,ratio:number)=>'#'+hex.replace('#','').match(/../g)!.map(ch=>Math.round(parseInt(ch,16)+(to-parseInt(ch,16))*ratio).toString(16).padStart(2,'0')).join('')
  return [{key:'default',label:'Theme',colors:[]},{key:'colorful',label:'Colorful',colors:accents},{key:'colorful2',label:'Colorful 2',colors:[...accents.slice(3),...accents.slice(0,3)]},...accents.map((c,i)=>({key:`mono-accent${i+1}`,label:`Accent ${i+1}`,colors:[mix(c,0,.25),c,mix(c,255,.25),mix(c,255,.45),mix(c,255,.65)]}))]
}
