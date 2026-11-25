/** Fixed upstream snapshot omits its local CRC helper import for media posters.
 * Add the existing canonical helper at bundle time without editing vendor. */
export function adaptMediaCrc(source){
 if(source.includes("import { crc32 } from './crc32'"))return source
 if(!source.includes('crc32(Buffer.concat('))throw Error('Upstream media poster contract changed')
 return "import { crc32 } from './crc32';\n"+source
}
