import test from 'node:test'
import assert from 'node:assert/strict'
import {readFile} from 'node:fs/promises'
import {chromium} from 'playwright-core'

test('mobile section thumbnails stay in one scrollable row; desktop stays vertical',async()=>{
 const css=(await readFile(new URL('../vendor/genoffice/apps/slides/src/renderer/styles.css',import.meta.url),'utf8'))+'\n'+await readFile(new URL('../web/src/mobile.css',import.meta.url),'utf8')
 const browser=await chromium.launch({executablePath:process.env.CHROME_PATH??'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:true})
 try{
  const page=await browser.newPage()
  const thumb=(i:number)=>`<div class="thumb"><canvas width="100" height="56"></canvas><span class="thumb-num">${i}</span></div>`
  for(const width of [390,691,1351]){
   await page.setViewportSize({width,height:700})
   await page.setContent(`<style>${css}</style><div class="workspace"><div class="slide-list"><div class="section-group"><div class="section-header">Section 1</div>${thumb(1)}${thumb(2)}</div><div class="section-group"><div class="section-header">Section 2</div>${thumb(3)}${thumb(4)}</div></div><div class="stage-col"></div></div>`)
   const bounds=await page.locator('.thumb').evaluateAll(es=>es.map(e=>({x:e.getBoundingClientRect().x,y:e.getBoundingClientRect().y})))
   if(width<=700){assert.ok(bounds.every(b=>Math.abs(b.y-bounds[0].y)<1));assert.ok(bounds.every((b,i)=>i===0||b.x>bounds[i-1].x));if(width===390)assert.ok(await page.locator('.slide-list').evaluate(e=>e.scrollWidth>e.clientWidth))}
   else{assert.ok(bounds[1].y>bounds[0].y);assert.equal(bounds[1].x,bounds[0].x)}
  }
 }finally{await browser.close()}
})
