/** View-local asynchronous dialogs work without granting sandbox allow-modals.
 * Text is inserted as textContent; existing modal tokens/style remain shared. */
import {t} from './locale-adapter'
import {presentation} from './presentation'
let tail:Promise<unknown>=Promise.resolve()
function dialog(message:string,initial?:string):Promise<string|null>{
  const task=tail.then(()=>new Promise<string|null>(resolve=>{
    const previous=document.activeElement as HTMLElement|null
    const backdrop=document.createElement('div');backdrop.className='modal-backdrop';backdrop.style.zIndex='100001'
    const panel=document.createElement('div');panel.className='modal';panel.style.cssText='max-width:calc(100vw - 32px);box-sizing:border-box';panel.setAttribute('role','dialog');panel.setAttribute('aria-modal','true');panel.tabIndex=-1
    const title=document.createElement('h2');title.textContent=initial===undefined?presentation.integration('confirmAction'):t('ribbonFileSaveAs');title.id='dialog-'+crypto.randomUUID();panel.setAttribute('aria-labelledby',title.id);panel.append(title)
    const note=document.createElement('p');note.className='modal-note';note.textContent=message;panel.append(note)
    let input:HTMLInputElement|undefined
    if(initial!==undefined){input=document.createElement('input');input.type='text';input.value=initial;input.maxLength=200;input.setAttribute('aria-label',presentation.integration('fileName'));input.style.cssText='width:100%;box-sizing:border-box';panel.append(input)}
    const actions=document.createElement('div');actions.className='modal-actions'
    const cancel=document.createElement('button');cancel.textContent=t('ribbonCancel');const ok=document.createElement('button');ok.className='primary';ok.textContent=t('ribbonOk')
    const valid=()=>!input||!!input.value.trim()&&input.value.length<=200&&!/[\\/\x00-\x1f]/.test(input.value)&&/\.pptx$/i.test(input.value)
    const finish=(value:string|null)=>{backdrop.remove();if(previous?.isConnected)previous.focus({preventScroll:true});resolve(value)}
    cancel.onclick=()=>finish(null);ok.onclick=()=>{if(valid())finish(input?.value.trim()??'yes')};input?.addEventListener('input',()=>{ok.disabled=!valid()});ok.disabled=!valid()
    actions.append(cancel,ok);panel.append(actions);backdrop.append(panel);document.body.append(backdrop)
    backdrop.addEventListener('click',e=>{if(e.target===backdrop)finish(null)})
    panel.addEventListener('keydown',e=>{if(e.key==='Escape'){e.preventDefault();e.stopPropagation();finish(null)}else if(e.key==='Enter'&&valid()){e.preventDefault();e.stopPropagation();finish(input?.value.trim()??'yes')}else if(e.key==='Tab'){const focusable=[...(input?[input]:[]),cancel,ok].filter(e=>!e.disabled);const first=focusable[0],last=focusable.at(-1);if(e.shiftKey&&document.activeElement===first){e.preventDefault();last?.focus()}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first?.focus()}}})
    if(input){input.focus();input.select()}else cancel.focus()
  }));tail=task.catch(()=>{});return task
}
export const confirmBrowserDialog=async(message:string)=>(await dialog(message))!==null
export const saveNameDialog=(initial:string)=>dialog(presentation.integration('saveCopyName'),initial)
