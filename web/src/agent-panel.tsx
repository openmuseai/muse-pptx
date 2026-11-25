import {useEffect,useRef,useState} from 'react'
import {setAgentEditorContext,askAgent} from './runtime'
import {presentation} from './presentation'

/** Retain the original editor's selection/flush contract without its AI chat. */
export function AiPanel(props: any) {
  const submitted=useRef<number|null>(null)
  const [error,setError]=useState('')
  useEffect(() => {
    setAgentEditorContext({slideIndex:props.current,selectedIds:props.selectedIds,beforeRun:props.onBeforeRun,goTo:async(index:number)=>{
      props.onNavigate(index)
      await new Promise<void>(resolve=>requestAnimationFrame(()=>requestAnimationFrame(()=>resolve())))
    }})
  },[props.current,JSON.stringify(props.selectedIds),props.onBeforeRun,props.onNavigate])
  useEffect(()=>{
    const preset=props.preset
    if(!preset?.text||submitted.current===preset.nonce)return
    submitted.current=preset.nonce;setError('')
    // The existing review/translate confirmation precedes this handoff; the
    // Host Agent retains its provider/data consent and live-tool revision gates.
    void askAgent(preset.text).catch(()=>setError(presentation.integration('agentUnavailable')))
  },[props.preset?.nonce])
  return error?<span role="alert">{error}</span>:null
}
