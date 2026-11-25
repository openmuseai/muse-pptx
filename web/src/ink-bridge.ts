import {Buffer} from 'buffer'
import {decodeInkPayload,INK_NAME_PREFIX} from '../../vendor/genoffice/apps/slides/src/renderer/ink'

/** Keep the upstream ink representation: a transparent PNG plus vector
 * metadata, so save/reopen, eraser and ordinary picture transactions agree. */
export function inkInsertOp(size:{cx:number;cy:number},input:any){
  if(!Number.isInteger(input.slideIndex)||input.slideIndex<0||!Number.isFinite(input.fitWidthPx)||input.fitWidthPx<=0)throw Error('Invalid ink viewport')
  if(![input.xPx,input.yPx,input.wPx,input.hPx].every(Number.isFinite)||input.wPx<=0||input.hPx<=0)throw Error('Invalid ink bounds')
  if(typeof input.payload!=='string'||input.payload.length>1024*1024||!decodeInkPayload(input.payload))throw Error('Invalid ink payload')
  if(typeof input.base64!=='string'||input.base64.length>24*1024*1024||!/^[A-Za-z0-9+/]+={0,2}$/.test(input.base64))throw Error('Invalid ink image')
  const bytes=Buffer.from(input.base64,'base64')
  if(bytes.toString('base64')!==input.base64||bytes.length>16*1024*1024||!bytes.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10])))throw Error('Invalid ink PNG')
  const emu=(px:number)=>Math.round(px*size.cx/input.fitWidthPx)
  return {op:'addPicture' as const,target:{slide:input.slideIndex},bytes:new Uint8Array(bytes),ext:'png',offset:{x:emu(input.xPx),y:emu(input.yPx),cx:Math.max(1,emu(input.wPx)),cy:Math.max(1,emu(input.hPx))},name:`${INK_NAME_PREFIX} ${crypto.randomUUID()}`,descr:input.payload}
}
