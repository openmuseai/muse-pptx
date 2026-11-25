/** Translate the original preload contract; canonical ops own PPTX mutation. */
export function elementInsertOp(size:{cx:number;cy:number},input:any){
  if(!Number.isInteger(input.slideIndex)||input.slideIndex<0||!Number.isFinite(input.fitWidthPx)||input.fitWidthPx<=0)throw Error('Invalid insertion viewport')
  if(![input.xPx,input.yPx,input.wPx,input.hPx].every(Number.isFinite)||input.wPx<0||input.hPx<0)throw Error('Invalid insertion bounds')
  const emu=(px:number)=>Math.round(px*size.cx/input.fitWidthPx)
  const paragraphs=input.paragraphs?.length?input.paragraphs:input.text?String(input.text).split('\n').map(text=>({runs:[{text}]})):undefined
  return {op:'addElement',target:{slide:input.slideIndex},kind:input.kind,
    offset:{x:emu(input.xPx),y:emu(input.yPx),cx:emu(input.wPx),cy:emu(input.hPx)},
    ...(paragraphs?{paragraphs}:{}),...(input.fillColor?{fill:input.fillColor}:{}),
    ...(input.stroke?{stroke:{color:input.stroke.color,widthEmu:Math.round(input.stroke.widthPt*12700)}}:{}),
    ...(input.bodyPr?{bodyPr:input.bodyPr}:{}),}
}

export const imageAccept='image/png,image/jpeg,image/gif,image/bmp,image/webp,image/tiff'
/** A user-selected browser File never exposes a native path to the plugin. */
export function pickImageFile():Promise<File|null>{
  return new Promise(resolve=>{
    const input=document.createElement('input');input.type='file';input.accept=imageAccept;input.hidden=true
    const finish=(file:File|null)=>{input.remove();resolve(file)}
    input.addEventListener('change',()=>finish(input.files?.[0]??null),{once:true})
    input.addEventListener('cancel',()=>finish(null),{once:true})
    document.body.append(input);input.click()
  })
}
export async function imageSize(file:File):Promise<{width:number;height:number}>{
  const url=URL.createObjectURL(file)
  try{const image=new Image();image.src=url;await image.decode();return {width:image.naturalWidth,height:image.naturalHeight}}
  finally{URL.revokeObjectURL(url)}
}
