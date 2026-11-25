// Complete installed OpenMuse WebView QA. No model mutation API is used.
import {chromium} from 'playwright-core'
import {execFileSync} from 'node:child_process'
import fs from 'node:fs/promises'
import path from 'node:path'
const args=process.argv.slice(2),arg=(key,fallback)=>{const i=args.indexOf(key);return i<0?fallback:args[i+1]}
const serial=arg('--serial'),scenario=arg('--scenario','inspect'),out=arg('--out'),port=arg('--port','19222')
if(!serial)throw Error('--serial is required; target only the authorized QA device')
if(!['inspect','core','record','link'].includes(scenario))throw Error('Unknown scenario')
if(scenario!=='inspect'&&!out)throw Error('A JSONL --out evidence path is required for edits')
const focus=execFileSync('adb',['-s',serial,'shell','dumpsys','window'],{encoding:'utf8'}).split('\n').find(s=>s.includes('mCurrentFocus='))??''
if(!focus.includes('com.openmuseai.mobile.acceptance/'))throw Error('Resume/cancel the native system flow first; OpenMuse QA must be foreground')
const browser=await chromium.connectOverCDP('http://127.0.0.1:'+port,{timeout:10000})
const page=browser.contexts().flatMap(c=>c.pages()).find(p=>p.url().includes('.plugins.openmuse.io/host/index.html'))
if(!page)throw Error('Complete OpenMuse plugin WebView not found')
page.setDefaultTimeout(8000)
const frame=page.frameLocator('#view'),errors=[]
page.on('pageerror',e=>errors.push(e.message))
try{
 const own=await frame.locator('html').evaluate(async()=>(await window.slidesApi.getRenderSlides()).some(s=>s.nodes.some(n=>(n.text?.lines??[]).map(l=>l.runs.map(r=>r.text).join('')).join('\n').includes('OpenMuse Toolbar'))))
 if(!own)throw Error('Open the self-authored OpenMuse Toolbar QA fixture first; user documents are not edited')
 let result
 if(scenario==='inspect')result={host:page.url(),status:await frame.locator('.status-left').innerText(),tabs:await frame.locator('.ribbon-tab').allTextContents()}
 else{
  const dir=import.meta.dirname,helpers=await fs.readFile(path.join(dir,'helpers.js'),'utf8')
  const files=scenario==='link'?['link-open','link']:[scenario]
  const code=helpers+'\n'+(await Promise.all(files.map(f=>fs.readFile(path.join(dir,'scenarios',f+'.js'),'utf8')))).join('\n')
  // Scenario code is versioned with this runner, never received from the plugin.
  result=await new Function('page','frame','errors','fs','return (async()=>{'+code+'})()')(page,frame,errors,fs)
 }
 const receipt={schema:'openmuse.mobile-toolbar-e2e/1',time:new Date().toISOString(),scenario,host:page.url(),result,errors}
 if(out)await fs.appendFile(out,JSON.stringify(receipt)+'\n')
 console.log(JSON.stringify(receipt))
}catch(error){if(out)await fs.appendFile(out,JSON.stringify({time:new Date().toISOString(),scenario,automationError:String(error),errors})+'\n');console.error(error.stack);process.exitCode=1}
finally{process.exit(process.exitCode??0)}
