import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const file="C:/Users/HZJ/Desktop/Git/work/finance_risk_20260908/deliverables/02-风险框架与行动清单-v1.1.xlsx";
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(file));
const sheets=await wb.inspect({kind:"sheet",include:"id,name",maxChars:3000});
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},maxChars:4000});
console.log(JSON.stringify({sheets:sheets.ndjson??String(sheets),errors:errors.ndjson??String(errors)}));
