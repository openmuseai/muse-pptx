import test from 'node:test'
import assert from 'node:assert/strict'
import {build} from 'esbuild'
import {mkdtemp,readFile,writeFile,readdir,rm} from 'node:fs/promises'
import path from 'node:path'
import {tmpdir} from 'node:os'
import {pathToFileURL} from 'node:url'
import {elementInsertOp} from '../web/src/insert-bridge'
import {inkInsertOp} from '../web/src/ink-bridge'

test('insertion rejects invalid viewport/bounds before a document mutation',()=>{
  for(const input of [{fitWidthPx:0},{fitWidthPx:NaN},{wPx:-1},{slideIndex:-1}])assert.throws(()=>elementInsertOp({cx:12192000,cy:6858000},{slideIndex:0,fitWidthPx:1280,xPx:40,yPx:50,wPx:200,hPx:100,...input}))
})
test('ink rejects invalid PNG, empty vector metadata and non-finite geometry',()=>{
  const valid={slideIndex:0,fitWidthPx:1280,xPx:40,yPx:50,wPx:200,hPx:100,payload:JSON.stringify({v:1,tool:'pen',color:'C00000',width:2,points:[[0,0],[100,100]]}),base64:'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aOQAAAABJRU5ErkJggg=='}
  assert.equal(inkInsertOp({cx:12192000,cy:6858000},valid).offset.cx,1905000)
  for(const patch of [{payload:'{}'},{base64:'SGVsbG8='},{wPx:0},{xPx:NaN},{fitWidthPx:0}])assert.throws(()=>inkInsertOp({cx:12192000,cy:6858000},{...valid,...patch}))
})
test('browser insertion bridge uses canonical ops: shape, textbox and image survive real OOXML save/reopen',async()=>{
  const root=path.resolve(import.meta.dirname,'..'),packages=new Map()
  for(const name of await readdir(path.join(root,'vendor/genoffice/packages'))){try{const base=path.join(root,'vendor/genoffice/packages',name),p=JSON.parse(await readFile(path.join(base,'package.json'),'utf8'));packages.set(p.name,{base,exports:p.exports})}catch{}}
  const dir=await mkdtemp(path.join(tmpdir(),'muse-pptx-edit-'))
  try{
    const result=await build({stdin:{contents:`
      import {openPptx,savePptx,createBlankPptx} from '@genoffice/pptx-engine';
      import {runTxn} from '@genoffice/pptx-ops';
      import {elementInsertOp} from ${JSON.stringify(path.join(root,'web/src/insert-bridge.ts'))};
      import {inkInsertOp} from ${JSON.stringify(path.join(root,'web/src/ink-bridge.ts'))};
      export async function exercise(){
        const deck=await openPptx(await createBlankPptx()),before=deck.deck.slides[0].elements.length;
        const shape=elementInsertOp(deck.deck.size,{slideIndex:0,fitWidthPx:1280,kind:'rect',xPx:80,yPx:100,wPx:240,hPx:120,fillColor:'#C43E1C',stroke:{color:'#000000',widthPt:1}});
        const text=elementInsertOp(deck.deck.size,{slideIndex:0,fitWidthPx:1280,kind:'textbox',xPx:140,yPx:280,wPx:360,hPx:100,text:'Muse Mobile insertion'});
        const picture={op:'addPicture',target:{slide:0},bytes:new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aOQAAAABJRU5ErkJggg==','base64')),ext:'png',offset:{x:2000000,y:2000000,cx:1000000,cy:1000000}};
        const ink=inkInsertOp(deck.deck.size,{slideIndex:0,fitWidthPx:1280,xPx:40,yPx:50,wPx:200,hPx:100,base64:Buffer.from(picture.bytes).toString('base64'),payload:JSON.stringify({v:1,tool:'pen',color:'C00000',width:2,points:[[0,0],[100,100]]})});
        const outcome=runTxn(deck,{ops:[shape,text,picture,ink]});
        if(!outcome.applied)throw Error(JSON.stringify(outcome.failures));
        const saved=await savePptx(deck),reopened=await openPptx(saved);
        return {before,after:reopened.deck.slides[0].elements.length,created:outcome.records.flatMap(r=>r.created||[]).length,xml:reopened.archive.readText(reopened.deck.slides[0].path)};
      }
    `,resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'esm',write:false,logLevel:'silent',loader:{'.md':'text'},banner:{js:"import {createRequire} from 'node:module'; const require=createRequire(import.meta.url);"},plugins:[{name:'same-canonical-packages-as-web',setup(b){b.onResolve({filter:/^@genoffice\//},args=>{const parts=args.path.split('/'),p=packages.get(parts.slice(0,2).join('/'));if(!p)return;const sub=parts.length===2?'.':'./'+parts.slice(2).join('/');let target=typeof p.exports==='string'?p.exports:p.exports[sub];if(!target)for(const [key,value] of Object.entries(p.exports))if(key.endsWith('*')&&sub.startsWith(key.slice(0,-1)))target=String(value).replace('*',sub.slice(key.length-1));if(!target)throw Error('Missing canonical export '+args.path);return {path:path.resolve(p.base,target)}})}}]})
    const file=path.join(dir,'canonical.mjs');await writeFile(file,result.outputFiles[0].contents)
    const exercised=await (await import(pathToFileURL(file).href)).exercise()
    assert.equal(exercised.after,exercised.before+4);assert.equal(exercised.created,4)
    assert.match(exercised.xml,/Muse Mobile insertion/);assert.match(exercised.xml,/<p:pic/);assert.match(exercised.xml,/prst="rect"/)
    assert.match(exercised.xml,/aislides-ink/);assert.match(exercised.xml,/points/)
  }finally{await rm(dir,{recursive:true,force:true})}
})
