// Define strict cell type for inference
type CellVal = string | number | boolean | Date;

/**
 * Unified Office Script: Normalize, populate, and flag Main sheet.
 * Steps:
 * 1. Normalize cols A (Material) & B (Season).
 * 2. Normalize & format cols C-D (dates).
 * 3. Copy E-F from GFE with "INACTIVE" logic.
 * 4. Flag changes in G (first available) and H (discontinue).
 */
function main(workbook: ExcelScript.Workbook): void {
  const SHEET = "Main";
  const GFE_SHEET = "GFE";
  const START_ROW = 2;

  const ws = workbook.getWorksheet(SHEET);
  const wsGfe = workbook.getWorksheet(GFE_SHEET);
  if (!ws || !wsGfe) return;

  // --- 1. Normalize A-B ---
  const used = ws.getUsedRange(); if (!used) return;
  const rows = used.getRowCount();
  const abRange = ws.getRangeByIndexes(START_ROW, 0, rows - START_ROW, 2);
  const abVals = abRange.getValues() as CellVal[][];
  for (let i = 0; i < abVals.length; i++) {
    // A: Material
    let m = abVals[i][0];
    if (typeof m === 'string') {
      const t = m.replace(/^'/,'').trim();
      const n = Number(t);
      abVals[i][0] = !isNaN(n) ? n : t;
    }
    // B: Season
    let s = abVals[i][1];
    if (typeof s === 'string') abVals[i][1] = s.trim().toUpperCase();
  }
  abRange.setValues(abVals);

  // --- 2. Normalize & format C-D ---
  function excelDate(num: number): Date {
    const d = num > 60 ? num - 1 : num;
    const e = new Date(Date.UTC(1899,11,31));
    return new Date(e.getTime()+d*86400000);
  }
  const cdRange = ws.getRangeByIndexes(START_ROW,2,rows-START_ROW,2);
  const cdVals = cdRange.getValues() as CellVal[][];
  const cdOut: string[][] = [];
  for (const row of cdVals) {
    const outRow: string[] = ["",""];
    for (let j=0;j<2;j++) {
      const v = row[j];
      let dt: Date|null = null;
      if (v instanceof Date) dt=v;
      else if (typeof v==='number') dt=excelDate(v);
      else if (typeof v==='string' && v.trim()) {
        const x=new Date(v); if(!isNaN(x.getTime())) dt=x;
      }
      if (dt) outRow[j]=`${String(dt.getMonth()+1).padStart(2,'0')}/${String(dt.getDate()).padStart(2,'0')}/${dt.getFullYear()}`;
    }
    cdOut.push(outRow);
  }
  const fmtCD = Array(cdOut.length).fill(["@","@"]); cdRange.setNumberFormatLocal(fmtCD);
  cdRange.setValues(cdOut);

  // --- 3. Populate E-F from GFE ---
  // build lookup
  const gfeUsed = wsGfe.getUsedRange(); if (!gfeUsed) return;
  const gfe = gfeUsed.getValues() as CellVal[][];
  const lookup: Record<string,{fa:CellVal,disco:CellVal,status:string}> = {};
  for (let i=1;i<gfe.length;i++) {
    const key = String(gfe[i][0]).trim().toLowerCase()+'|'+String(gfe[i][1]).trim().toLowerCase();
    lookup[key]={fa:gfe[i][2],disco:gfe[i][3],status:String(gfe[i][4]).trim().toUpperCase()};
  }
  const efRange = ws.getRangeByIndexes(START_ROW,4,rows-START_ROW,2);
  const efOut:string[][]=[];
  for (let i=START_ROW;i<rows;i++){
    const m=String(abVals[i-START_ROW][0]).trim().toLowerCase();
    const s=String(abVals[i-START_ROW][1]).trim().toLowerCase();
    const e=lookup[m+'|'+s];
    let fa="",di="";
    if(e){
      if(e.status!="ACTIVE") fa=di="INACTIVE";
      else{
        const d=parseDate(e.fa),dd=parseDate(e.disco);
        if(d) fa=formatDate(d); if(dd) di=formatDate(dd);
      }
    }
    efOut.push([fa,di]);
  }
  efRange.setNumberFormatLocal(Array(efOut.length).fill(["@","@"]));
  efRange.setValues(efOut);

  // --- 4. Flag G-H ---
  function toMs(v:CellVal):number|""{if(v instanceof Date)return v.getTime();if(typeof v==='number')return v;const x=String(v).trim();if(!x)return"";const z=new Date(x);return isNaN(z.getTime())?"":z.getTime();}
  const ghRange=ws.getRangeByIndexes(START_ROW,6,rows-START_ROW,2);
  const ghVals=ghRange.getValues() as CellVal[][];
  const ghOut:string[][]=[];
  for (let i=START_ROW;i<rows;i++){
    const c=toMs(cdOut[i-START_ROW][0]||"");
    const e=toMs(efOut[i-START_ROW][0]||"");
    const d=toMs(cdOut[i-START_ROW][1]||"");
    const f=toMs(efOut[i-START_ROW][1]||"");
    ghOut.push([c!==e?"Y":"",d!==f?"Y":""]);
  }
  ghRange.setValues(ghOut);

  // parseDate reused from above
  function parseDate(raw: CellVal): Date|null { if(raw instanceof Date)return raw; if(typeof raw==='number')return excelDate(raw); if(typeof raw==='string'&&raw.trim()){const d=new Date(raw);return isNaN(d.getTime())?null:d;}return null; }
  function excelDate(n:number):Date{const o=n>60?n-1:n;const e=new Date(Date.UTC(1899,11,31));return new Date(e.getTime()+o*86400000);}  
}
