import fs from 'node:fs/promises';
import path from 'node:path';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const base=path.dirname(new URL(import.meta.url).pathname.slice(1));
const src=path.join(base,'integrated','应收账款管理完整工具包（融合修订版）.xlsx');
const out=path.join(base,'qa','integrated_xlsx'); await fs.mkdir(out,{recursive:true});
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(src));
const err=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:500},summary:'integrated workbook error scan'});
await fs.writeFile(path.join(out,'errors.ndjson'),err.ndjson??String(err));
const names=wb.worksheets.items.map(x=>x.name);
for(const n of names){const png=await wb.render({sheetName:n,autoCrop:'all',scale:.5,format:'png'});await fs.writeFile(path.join(out,`${n}.png`),new Uint8Array(await png.arrayBuffer()))}
console.log(JSON.stringify({sheets:names.length,names}));
