import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { SpreadsheetFile,FileBlob } from '@oai/artifact-tool';
const root=fileURLToPath(new URL('.',import.meta.url)).replaceAll('\\','/').replace(/\/$/,'');
const fmm=process.argv.includes('--fmm');
const out=fmm?'outputs/bem_vba_stage3/Elastic_BEM_VBA_Full.xlsm':'outputs/bem_vba_stage2/Elastic_BEM_VBA_Solver.xlsm';
const qa=fmm?'qa_stage3':'qa_stage2';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(`${root}/${out}`));
wb.recalculate();
const checks=[];
for (const [sheetName,range] of [['操作','A16:B23'],['操作','A34:B43'],...(fmm?[['操作','A49:B61']]:[]),['内点結果','A5:Q6']]) {
 const inspected=await wb.inspect({kind:'table',range:`${sheetName}!${range}`,include:'values,formulas',tableMaxRows:10,tableMaxCols:17,maxChars:2500});
 checks.push(inspected.ndjson);console.log(inspected.ndjson);
}
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},maxChars:1000});checks.push(errors.ndjson);console.log(errors.ndjson);
for(const [sheetName,range,name] of [
 ['操作','A1:D30','操作'],['操作','A33:D47','解析設定'],...(fmm?[['操作','A48:D62','FMM設定']]:[]),['境界結果','A1:M12','境界結果'],['内点結果','A1:Q8','内点結果'],['使い方','A1:B23','使い方']]){
 const blob=await wb.render({sheetName,range,scale:1.4,format:'png'});
 await fs.writeFile(`${root}/${qa}/final_${name}.png`,new Uint8Array(await blob.arrayBuffer()));
}
await fs.writeFile(`${root}/${qa}/visual_inspection.ndjson`,checks.join('\n'));
console.log('Final solver workbook rendered and inspected.');
