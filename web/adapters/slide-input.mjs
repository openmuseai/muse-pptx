/** Rebind only the Stage input surface; retain upstream drawing/model logic. */
export function adaptSlideInput(source) {
  const start=source.indexOf('<Stage\n'),end=source.indexOf('onContextMenu=',start)
  if(start<0||end<start)throw Error('Upstream Stage input contract changed')
  const original=source.slice(start,end)
  const down=original.indexOf('onMouseDown='),move=original.indexOf('onMouseMove='),up=original.indexOf('onMouseUp=')
  if(down<0||move<down||up<move)throw Error('Upstream Stage input contract changed')
  const normalize="event.evt.preventDefault(); const e = {...event, evt: {...event.evt, button: 0, shiftKey: false, ctrlKey: false, metaKey: false, altKey: false}};"
  const touchDown=original.slice(down,move).replace('onMouseDown={(e) => {','onTouchStart={(event) => { '+normalize)
  const touchMove=original.slice(move,up).replace('onMouseMove={(e) => {','onTouchMove={(event) => { '+normalize)
  const touchUp=original.slice(up).replace('onMouseUp={() => {','onTouchEnd={(event) => { event.evt.preventDefault();')
  if(!touchDown.startsWith('onTouchStart=')||!touchMove.startsWith('onTouchMove=')||!touchUp.startsWith('onTouchEnd='))throw Error('Upstream Stage handler signature changed')
  const handlers=original+touchDown+touchMove+touchUp+'onTouchCancel={() => { if (!drawMode) return; drawRef.current = null; setDrawPreview(null); onDrawCancel?.() }}\n      '
  return source.slice(0,start)+handlers+source.slice(end)
}
