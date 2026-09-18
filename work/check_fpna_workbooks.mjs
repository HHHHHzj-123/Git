import {FileBlob,SpreadsheetFile} from 'file:///C:/Users/HZJ/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
for(const p of ['C:/Users/HZJ/Desktop/Git/work/fpna-output/财务与经营分析模型.xlsx','C:/Users/HZJ/Desktop/Git/work/fpna-output/常见行业毛利率参考数据库.xlsx']){
 const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(p));
 console.log(p);
 console.log((await wb.inspect({kind:'sheet',include:'id,name',maxChars:8000})).ndjson);
 if(p.includes('模型')) console.log((await wb.inspect({kind:'table',range:'07-收入PVM!A6:I13',include:'values,formulas',tableMaxRows:10,tableMaxCols:10})).ndjson);
 else console.log((await wb.inspect({kind:'table',range:'01-官方行业数据!A6:J12',include:'values,formulas',tableMaxRows:10,tableMaxCols:10})).ndjson);
 console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:100},summary:'post export errors'})).ndjson);
}
