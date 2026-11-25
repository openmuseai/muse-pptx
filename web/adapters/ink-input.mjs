/** Small, checked product overlays; the upstream checkout stays untouched. */
export function adaptInkApp(source){
 const commit='(stroke: InkStroke) => arrangeActions.commitInk(ctxRef.current, stroke),'
 if(!source.includes(commit))throw Error('Upstream ink commit contract changed')
 source=source.replace(commit,`async (stroke: InkStroke) => {
      try { await arrangeActions.commitInk(ctxRef.current, stroke) }
      catch (error) { setInkTool('select'); throw error }
    },`)
 const state="const [inkTool, setInkTool] = useState<InkTool>('select')"
 if(!source.includes(state))throw Error('Upstream ink state contract changed')
 return source.replace(state,state+`
  useEffect(() => {
    const exit = (event: KeyboardEvent) => { if (event.key === 'Escape') setInkTool('select') }
    window.addEventListener('keydown', exit)
    return () => window.removeEventListener('keydown', exit)
  }, [])`)
}
export function adaptInkRibbon(source){
 const tab='setTab(tb)'
 if(source.split(tab).length!==2)throw Error('Upstream ink navigation contract changed')
 return source.replace(tab,"if (tb !== 'draw') onInkTool('select'); "+tab)
}
