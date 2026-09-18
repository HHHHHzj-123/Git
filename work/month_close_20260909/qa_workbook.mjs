import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const file="C:/Users/HZJ/Desktop/Git/work/month_close_20260909/deliverables/M02-001-年度工作与月结执行清单.xlsx";
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(file));
const sheets=await wb.inspect({kind:"sheet",include:"id,name",maxChars:5000});
const formulas=await wb.inspect({kind:"table",range:"04科目与报表勾稽!A4:N8",include:"values,formulas",tableMaxRows:8,tableMaxCols:14,maxChars:10000});
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},maxChars:4000});
console.log(JSON.stringify({sheets:sheets.ndjson??String(sheets),formulas:formulas.ndjson??String(formulas),errors:errors.ndjson??String(errors)}));
