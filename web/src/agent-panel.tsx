import {useEffect} from 'react'
import {setAgentEditorContext} from './runtime'

/** Retain the original editor's selection/flush contract without its AI chat. */
export function AiPanel(props: any) {
  useEffect(() => {
    setAgentEditorContext({slideIndex:props.current,selectedIds:props.selectedIds,beforeRun:props.onBeforeRun})
  },[props.current,JSON.stringify(props.selectedIds),props.onBeforeRun])
  return null
}
