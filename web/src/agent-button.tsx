import {useState} from 'react'
import {askAgent} from './runtime'

export function AgentButton() {
  const [pending,setPending]=useState(false)
  const [error,setError]=useState('')
  return <div className="openmuse-agent-entry">
    <button className="rb-big" disabled={pending} onClick={async()=>{
      setPending(true);setError('')
      try {await askAgent('请查看当前演示文稿和选中的幻灯片，先询问我希望如何修改。通过文档插件提供的工具读取、修改和保存。')}
      catch(e){setError(String(e))}finally{setPending(false)}
    }}><span className="rb-big-icon" aria-hidden="true">✦</span><span>{pending?'连接 Agent…':'Agent'}</span></button>
    {error&&<span role="alert">{error}</span>}
  </div>
}
