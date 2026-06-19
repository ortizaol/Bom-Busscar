// Test script — same logic as bom_generator.html
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import { createRequire } from 'module';

// ---- PDF.js (Node) ----
const require = createRequire(import.meta.url);
let pdfjs;
try {
  pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
} catch(e) {
  pdfjs = require('pdfjs-dist');
}

async function extractPdfText(filePath) {
  const data = new Uint8Array(fs.readFileSync(filePath));
  const doc = await pdfjs.getDocument({ data, useWorkerFetch: false, isEvalSupported: false, useSystemFonts: true }).promise;
  let txt = '';
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    txt += content.items.map(x => x.str).join(' ') + '\n';
  }
  return txt;
}

// ---- Helpers (same as HTML) ----
function norm(s) { return (s||'').toString().normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().trim(); }
function parseDimNum(s) {
  if (!s) return 0;
  s = s.toString().trim().replace(/\s/g,'');
  const parts = s.split('.');
  if (parts.length <= 1) return parseFloat(s.replace(',','.')) || 0;
  const last = parts[parts.length-1];
  if (last.length <= 2) return parseFloat(parts.slice(0,-1).join('') + '.' + last) || 0;
  return parseFloat(parts.join('')) || 0;
}
function parseFactura(txt) {
  const marcaM = txt.match(/MARCA\s+CHASIS\s*:\s*(.+?)(?=\n|SERIE\s+DE|NUMERO\s+DE\s+MOTOR|COLOR\s*:)/is);
  const serieM = txt.match(/SERIE\s+DE\s+CHASIS\s*:\s*([A-Z0-9]{6,20})/i);
  const opM    = txt.match(/\bOP\s*:\s*(\d{4,6})/i) || txt.match(/\bOP\s+(\d{4,6})\s/i);
  const trmM   = txt.match(/TRM\s*\$\s*([\d.,]+)/i);
  const tipoM  = txt.match(/\d+\s+(CARROCER[ÍI]A\s+PARA\s+BUS\s+[^\n]+)/i) ||
                 txt.match(/CARROCER[ÍI]A\s+PARA\s+BUS\s+([^\n]+)/i);
  let valorExworks = '', valorGastosOrigen = '';
  const summaryIdx = txt.search(/VALOR\s*\/\s*VALUE\s+EXW/i);
  if (summaryIdx >= 0) {
    const chunk = txt.slice(summaryIdx, summaryIdx + 600);
    const allUsd = [...chunk.matchAll(/US\$([\d.,]+)/gi)];
    console.log('  USD values in EXW block:', allUsd.map(m=>m[1]));
    if (allUsd.length >= 2) valorExworks      = allUsd[1][1];
    if (allUsd.length >= 3) valorGastosOrigen = allUsd[2][1];
  }
  if (!valorExworks) {
    const allUsd = [...txt.matchAll(/US\$([\d.,]+)/gi)];
    if (allUsd.length >= 2) valorExworks = allUsd[1][1];
  }
  return { marcaChasis: marcaM?.[1]?.trim()||'', tipoCarroceria: tipoM?.[1]?.trim().replace(/\s+/g,' ')||'',
    numOP: opM?.[1]||'', vin: serieM?.[1]||'', trm: trmM?.[1]||'', valorExworks, valorGastosOrigen };
}

function parseDeclaracion(txt) {
  const trmM  = txt.match(/58\s*\.?\s*Tasa\s+de\s+cambio[^0-9]*([\d.]+)/i) || txt.match(/Tasa\s+de\s+cambio[^0-9]*([\d.]+)/i);
  const cantM = txt.match(/77\s*\.?\s*Cantidad\s+dcms?\.?\s*([\d.]+)/i)     || txt.match(/Cantidad\s+dcms?\.?\s*([\d.]+)/i);
  const fobM  = txt.match(/78\s*\.?\s*Valor\s+FOB\s+USD\s*([\d.,]+)/i);
  const fletM = txt.match(/79\s*\.?\s*Valor\s+fletes\s+USD\s*([\d.,]+)/i);
  const segM  = txt.match(/80\s*\.?\s*Valor\s+seguro\s+USD\s*([\d.,]+)/i);
  const fob = parseDimNum(fobM?.[1]); const flete = parseDimNum(fletM?.[1]); const seg = parseDimNum(segM?.[1]);
  return { valorAduana: fob>0 ? (fob+flete+seg).toFixed(2) : '',
    tasaCambio: parseDimNum(trmM?.[1]).toFixed(2),
    cantidadDcms: String(Math.round(parseDimNum(cantM?.[1]||'1'))) };
}

// ---- Excel helpers ----
function sheetHeaders(sheet, rowNum=1) {
  const h=[]; sheet.getRow(rowNum).eachCell({includeEmpty:true},(c,col)=>{ h[col-1]=c.text||''; }); return h;
}
function findCol(headers, ...pats) {
  for (let i=0;i<headers.length;i++) {
    const h=norm(headers[i]); if(!h) continue;
    for (const p of pats) { if(h.includes(norm(p))) return i+1; }
  } return -1;
}
function cellVal(sheet,r,col) {
  if(col<1) return '';
  const v=sheet.getRow(r).getCell(col).value;
  if(v&&typeof v==='object'&&v.result!==undefined) return v.result;
  if(v&&typeof v==='object'&&v.formula) return '';
  return v??'';
}
function cellTxt(sheet,r,col) { return (cellVal(sheet,r,col)??'').toString().trim(); }

const UPLOADS = '/root/.claude/uploads/c44649cc-d03f-535e-a4b4-5a3f2bed3772';

async function main() {
  console.log('\n========== FACTURA PDF ==========');
  const factTxt = await extractPdfText(`${UPLOADS}/8caf3e96-2._FACTURA_DEFINITIVA_EP13928.pdf`);
  const fd = parseFactura(factTxt);
  console.log(fd);

  console.log('\n========== DIM PDF ==========');
  const dimTxt = await extractPdfText(`${UPLOADS}/8c93ea8a-6._DIM_4820250008423620.pdf`);
  const dd = parseDeclaracion(dimTxt);
  console.log(dd);
  console.log('  FOB+Flete+Seguro (valorAduana):', dd.valorAduana);

  console.log('\n========== EXCEL CONSUMOS ==========');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(`${UPLOADS}/f48a3635-OPS_COMEX_EXPO_ABRIL_2026__prueba.xlsx`);
  console.log('Hojas:', wb.worksheets.map(w=>w.name));

  const finalSheet = wb.worksheets.find(w=>norm(w.name)===norm('FINAL'));
  if (finalSheet) {
    const h = sheetHeaders(finalSheet);
    console.log('FINAL headers:', h);
    const cItem=findCol(h,'ítem','item'), cDesc=findCol(h,'descripción','descripcion','descripci');
    const cUM=findCol(h,'u.m','um'), cCant=findCol(h,'suma de cantidad','cantidad');
    const cCosto=findCol(h,'suma de costo total','costo total','costo');
    const cProv=findCol(h,'proveedor','razón social','razon social'), cOrig=findCol(h,'origen','pais');
    console.log(`  cols: item=${cItem} desc=${cDesc} um=${cUM} cant=${cCant} costo=${cCosto} prov=${cProv} orig=${cOrig}`);
    console.log(`  Row 2 sample: item=${cellTxt(finalSheet,2,cItem)} desc=${cellTxt(finalSheet,2,cDesc)} cant=${cellVal(finalSheet,2,cCant)} costo=${cellVal(finalSheet,2,cCosto)}`);
    console.log(`  Total rows in FINAL: ${finalSheet.rowCount}`);
  }

  // DATOS cost/qty
  const datosSheet = wb.worksheets.find(w=>norm(w.name).includes('dato'));
  if (datosSheet) {
    const dH=sheetHeaders(datosSheet);
    const dItem=findCol(dH,'ítem','item'), dCosto=findCol(dH,'costo mp (sal)','costo mp','costo total','costo');
    const dCant=findCol(dH,'salidas (inv.)','salidas','suma de cantidad','cantidad');
    const dIncl=findCol(dH,'incluir'), dOP=findCol(dH,'referencia1','o.p. referencia','valida','op numero','op');
    console.log(`\nDATOS cols: item=${dItem} costo=${dCosto} cant=${dCant} incl=${dIncl} op=${dOP}`);
    const OP='EP13928';
    let included=0, excluded=0, wrongOP=0;
    for(let r=2;r<=datosSheet.rowCount;r++){
      const incl=cellTxt(datosSheet,r,dIncl).toUpperCase();
      const rowOP=cellTxt(datosSheet,r,dOP);
      if(incl==='NO'){excluded++;continue;}
      if(rowOP!==OP){wrongOP++;continue;}
      included++;
    }
    console.log(`  For OP=${OP}: included=${included}, excluded(NO)=${excluded}, wrong OP=${wrongOP}`);
    console.log(`  DATO row 2: incl=${cellTxt(datosSheet,2,dIncl)} op=${cellTxt(datosSheet,2,dOP)} item=${cellTxt(datosSheet,2,dItem)}`);
  }

  const cfSheet = wb.worksheets.find(w=>norm(w.name).includes('consumo final'));
  if (cfSheet) {
    const cfH=sheetHeaders(cfSheet);
    const cfItem=findCol(cfH,'ítem','item'), cfProv=findCol(cfH,'razón social','razon social','razon social proveedor','proveedor');
    const cfGast=findCol(cfH,'% gastos destino','gastos destino','% gastos','cif');
    console.log(`\nCONSUMO FINAL cols: item=${cfItem} prov=${cfProv} %gastos=${cfGast}`);
    console.log(`  Row 2: item=${cellTxt(cfSheet,2,cfItem)} prov=${cellTxt(cfSheet,2,cfProv)} pct=${cellVal(cfSheet,2,cfGast)}`);
  }
}

main().catch(e=>{ console.error(e); process.exit(1); });
