/** Records only the current presentation canvas, never the phone screen or UI.
 * Canvas geometry is intersected with the slide bounds to exclude its bleed. */
export async function capturePresentation(width:number,height:number):Promise<MediaStream>{
  if(typeof MediaRecorder==='undefined')throw Error('MediaRecorder unavailable')
  const canvas=document.createElement('canvas')
  const scale=Math.min(1,1920/Math.max(width,height))
  canvas.width=Math.max(2,Math.round(width*scale));canvas.height=Math.max(2,Math.round(height*scale))
  const ctx=canvas.getContext('2d');if(!ctx||!canvas.captureStream)throw Error('Canvas recording unavailable')
  let timer:ReturnType<typeof setTimeout>|undefined,closed=false
  const draw=()=>{
    if(closed)return
    // The editor's slide box contains only this document's Konva layers.
    const box=document.querySelector<HTMLElement>('.slideshow .ss-frame .konvajs-content')??document.querySelector<HTMLElement>('.stage-rel')
    if(!box?.isConnected){stop();return}
    const bounds=box.getBoundingClientRect()
    if(bounds.width&&bounds.height){
      ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height)
      const layers:Array<HTMLCanvasElement|HTMLVideoElement>=[...box.querySelectorAll('canvas')]
      // Playing media is a DOM overlay above the show canvas, not a Konva node.
      if(box.closest('.slideshow'))layers.push(...(box.closest('.ss-frame')?.querySelectorAll('video')??[]))
      for(const layer of layers){
        if(layer instanceof HTMLVideoElement&&layer.readyState<2)continue
        const r=layer.getBoundingClientRect();if(!r.width||!r.height)continue
        const x=Math.max(bounds.left,r.left),y=Math.max(bounds.top,r.top)
        const w=Math.min(bounds.right,r.right)-x,h=Math.min(bounds.bottom,r.bottom)-y
        const sourceW=layer instanceof HTMLVideoElement?layer.videoWidth:layer.width,sourceH=layer instanceof HTMLVideoElement?layer.videoHeight:layer.height
        if(!sourceW||!sourceH)continue
        if(w>0&&h>0)ctx.drawImage(layer,(x-r.left)*sourceW/r.width,(y-r.top)*sourceH/r.height,w*sourceW/r.width,h*sourceH/r.height,(x-bounds.left)*canvas.width/bounds.width,(y-bounds.top)*canvas.height/bounds.height,w*canvas.width/bounds.width,h*canvas.height/bounds.height)
      }
      if(document.querySelector('.slideshow .ss-black')){ctx.fillStyle='#000';ctx.fillRect(0,0,canvas.width,canvas.height)}
      else if(document.querySelector('.slideshow .ss-white')){ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height)}
    }
    timer=setTimeout(draw,1000/15)
  }
  const stop=()=>{if(closed)return;closed=true;if(timer)clearTimeout(timer);window.removeEventListener('pagehide',stop);document.removeEventListener('visibilitychange',hidden);stream.getTracks().forEach(t=>t.stop());canvas.width=canvas.height=0}
  const hidden=()=>{if(document.hidden){stream.getVideoTracks().forEach(t=>t.dispatchEvent(new Event('ended')));stop()}}
  const stream=canvas.captureStream(15)
  draw()
  // MediaStreamTrack.stop does not dispatch ended; release the timer explicitly.
  for(const track of stream.getTracks()){const original=track.stop.bind(track);track.stop=()=>{original();stop()}}
  window.addEventListener('pagehide',stop,{once:true});document.addEventListener('visibilitychange',hidden)
  return stream
}
