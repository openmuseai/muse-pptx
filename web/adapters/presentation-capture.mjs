/** Keep the original record/stop/insert workflow, change only the capture source. */
export function adaptPresentationCapture(source,importPath){
 const original='await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false })'
 if(source.split(original).length!==2)throw Error('Upstream screen recording contract changed')
 return `import {capturePresentation} from ${JSON.stringify(importPath)};\n`+source.replace(original,'await capturePresentation(ctx.slide.widthPx, ctx.slide.heightPx)')
}
