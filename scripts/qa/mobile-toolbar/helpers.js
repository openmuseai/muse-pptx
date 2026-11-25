const escapeRegex=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
const tab=async(label)=>{const b=frame.locator('.ribbon-tab').filter({hasText:new RegExp('^'+escapeRegex(label)+'$')});if(!(await b.getAttribute('class')).includes(' active')||!await frame.locator('.ribbon-body').isVisible())await b.click()};
const btn=(label)=>frame.locator('.ribbon-body button').filter({hasText:new RegExp('^'+escapeRegex(label)+'$')});
const delay=async(ms=180)=>frame.locator('html').evaluate((_,ms)=>new Promise(r=>setTimeout(r,ms)),ms);
const mutate=async(action)=>{const r=await frame.locator('html').getAttribute('data-pptx-revision');await action();await frame.locator('html').evaluate(async(_,prev)=>{for(let i=0;i<100;i++){if(document.documentElement.dataset.pptxRevision!==prev)return;await new Promise(r=>setTimeout(r,30))}throw Error('Revision unchanged')},r);return await frame.locator('html').getAttribute('data-pptx-revision')};
const nodes=async()=>frame.locator('html').evaluate(async()=>(await window.slidesApi.getRenderSlides())[0].nodes.map(n=>({id:n.sourceId,t:n.type,box:n.box,prst:n.prst})));
