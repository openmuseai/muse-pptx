/** The browser sandbox intentionally disallows blocking native JS modals. */
export function adaptBrowserDialogsRibbon(source,importPath){
 const original="const confirmAiRewrite = () => {"
 if(source.split(original).length!==2||source.split("if (confirmAiRewrite())").length!==3)throw Error('Upstream AI confirmation contract changed')
 source=source.replace(original,'const confirmAiRewrite = async () => {').replace("if (!window.confirm(t('ribbonAiRewriteConfirm'))) return false","if (!await confirmBrowserDialog(t('ribbonAiRewriteConfirm'))) return false")
 source=source.replace("onClick={() => {\n                  if (confirmAiRewrite())", "onClick={async () => {\n                  if (await confirmAiRewrite())")
 source=source.replace("onClick={() => {\n                          setTranslateOpen(false)\n                          if (confirmAiRewrite())", "onClick={async () => {\n                          setTranslateOpen(false)\n                          if (await confirmAiRewrite())")
 if(source.includes('if (confirmAiRewrite())'))throw Error('Unadapted blocking confirmation')
 return `import {confirmBrowserDialog} from ${JSON.stringify(importPath)};\n`+source
}
