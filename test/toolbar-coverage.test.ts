import test from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import ts from 'typescript'

test('every preload method used by the bundled editor has a browser implementation',async()=>{
  const root=path.resolve(import.meta.dirname,'..'),implemented=new Set<string>()
  for(const file of ['runtime.ts','toolbar-bridge.ts','browser-services.ts']){
    const name=path.join(root,'web/src',file),source=ts.createSourceFile(name,await fs.readFile(name,'utf8'),ts.ScriptTarget.Latest,true)
    const members=(object:ts.ObjectLiteralExpression)=>{for(const member of object.properties)if(ts.isPropertyAssignment(member)||ts.isShorthandPropertyAssignment(member)||ts.isMethodDeclaration(member))implemented.add(member.name.getText(source).replace(/^['"]|['"]$/g,''))}
    function visit(node:ts.Node){
      if(ts.isVariableDeclaration(node)&&node.name.getText(source)==='api'&&node.initializer&&ts.isObjectLiteralExpression(node.initializer))members(node.initializer)
      if(ts.isFunctionDeclaration(node)&&node.name?.text==='createBrowserServices')for(const statement of node.body?.statements??[])if(ts.isReturnStatement(statement)&&statement.expression&&ts.isObjectLiteralExpression(statement.expression))members(statement.expression)
      ts.forEachChild(node,visit)
    }visit(source)
  }
  const meta=JSON.parse(await fs.readFile(path.join(root,'dist/web-metafile.json'),'utf8')),missing=new Map<string,string[]>(),used=new Set<string>()
  for(const file of Object.keys(meta.inputs).filter(f=>f.includes('/renderer/')&&/\.tsx?$/.test(f))){
    const source=await fs.readFile(path.resolve(root,file),'utf8')
    for(const match of source.matchAll(/slidesApi\.(\w+)/g)){
      const name=match[1];if(name.startsWith('on'))continue;used.add(name)
      if(!implemented.has(name)){const files=missing.get(name)??[];files.push(file);missing.set(name,files)}
    }
  }
  assert(used.size>100,'The production editor import graph must be present, not an empty/stub entry')
  assert.deepEqual([...missing],[],'A toolbar/dialog must never fall through to unsupportedOperation')
})
