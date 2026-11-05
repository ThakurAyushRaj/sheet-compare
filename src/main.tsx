import React, {useState} from 'react';
import {createRoot} from 'react-dom/client';
import * as XLSX from 'xlsx';
import {ArrowRight, ArrowDownToLine, Check, ChevronDown, FileSpreadsheet, Files, Github, Layers3, LockKeyhole, Plus, ShieldCheck, Sparkles, Upload, X, ArrowLeftRight} from 'lucide-react';
import './style.css';

type Book = {name:string; book:XLSX.WorkBook};
type Diff = {sheet:string;cell:string;rightCell:string;unit:string;field:string;old:string;value:string;type:'Changed'};
export function compare(a:XLSX.WorkBook,b:XLSX.WorkBook,trim=true,ignoreCase=false):{diffs:Diff[];total:number; sheets:string[]} {
 const singleSheet=a.SheetNames.length===1&&b.SheetNames.length===1; const sheets=singleSheet?[a.SheetNames[0]]:Array.from(new Set([...a.SheetNames,...b.SheetNames])); const diffs:Diff[]=[];let total=0;
 const fields:Record<string,string>={unitno:'Unit No.',unitnumber:'Unit No.',status:'Status',unitstatus:'Status',unitdescription:'Unit Description',plotsize:'Plot Size',width:'Width',length:'Length',dimensionunit:'Dimension Unit',dimensionunits:'Dimension Unit'};
 const normalizeHeader=(v:unknown)=>String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');
 const headers=(s:XLSX.WorkSheet)=>{let best={row:0,columns:new Map<string,number>()};for(let row=0;row<20;row++){const columns=new Map<string,number>();for(const key of Object.keys(s)){if(key.startsWith('!'))continue;const address=XLSX.utils.decode_cell(key);if(address.r!==row)continue;const field=fields[normalizeHeader(s[key]?.v)];if(field)columns.set(field,address.c);}if(columns.size>best.columns.size)best={row,columns};}return best;};
 let matchedFields=0,matchedUnits=0;
 const normalize=(v:string)=>{if(trim)v=v.trim();return ignoreCase?v.toLowerCase():v;};
 for(const sheet of sheets){
  const left=a.Sheets[sheet],right=b.Sheets[singleSheet?b.SheetNames[0]:sheet];
  if(!left||!right)continue;
  const lh=headers(left),rh=headers(right);
  const leftUnitColumn=lh.columns.get('Unit No.'),rightUnitColumn=rh.columns.get('Unit No.');
  if(leftUnitColumn===undefined||rightUnitColumn===undefined)throw new Error('Both files must have a Unit No. column so units can be matched correctly.');
  const indexUnits=(s:XLSX.WorkSheet,headerRow:number,column:number,label:string)=>{
   const units=new Map<string,{row:number;unit:string}>();
   const end=XLSX.utils.decode_range(s['!ref']||'A1').e.r;
   for(let row=headerRow+1;row<=end;row++){
    const unit=String(s[XLSX.utils.encode_cell({r:row,c:column})]?.v??'').trim();
    const key=normalize(unit);
    if(!key)continue;
    const existing=units.get(key);
    if(existing)throw new Error('Duplicate Unit No. "'+unit+'" in '+label+' sheet "'+sheet+'" (rows '+(existing.row+1)+' and '+(row+1)+'). Each unit number must be unique.');
    units.set(key,{row,unit});
   }
   return units;
  };
  const leftUnits=indexUnits(left,lh.row,leftUnitColumn,'original'),rightUnits=indexUnits(right,rh.row,rightUnitColumn,'updated');
  const sharedFields=Array.from(lh.columns).filter(([field])=>field!=='Unit No.'&&rh.columns.has(field));
  matchedFields+=sharedFields.length;
  for(const [key,original] of leftUnits){
   const updated=rightUnits.get(key);if(!updated)continue;matchedUnits++;
   for(const [field,column] of sharedFields){
    const cell=XLSX.utils.encode_cell({r:original.row,c:column});
    const rightCell=XLSX.utils.encode_cell({r:updated.row,c:rh.columns.get(field)!});
    const old=String(left[cell]?.v??''),value=String(right[rightCell]?.v??'');total++;
    if(normalize(old)!==''&&normalize(value)!==''&&normalize(old)!==normalize(value))diffs.push({sheet,cell,rightCell,unit:original.unit,field,old,value,type:'Changed'});
   }
  }
 }
 if(!matchedFields)throw new Error('No matching comparison columns found. Include Status, Unit Description, Plot Size, Width, Length, or Dimension Unit alongside Unit No.');
 if(!matchedUnits)throw new Error('No matching unit numbers were found. Check the Unit No. values in both files.');
 return {diffs,total,sheets};
}
function App(){
 const [books,setBooks]=useState<(Book|null)[]>([null,null]); const [error,setError]=useState('');const [busy,setBusy]=useState(false);const [result,setResult]=useState<ReturnType<typeof compare>|null>(null);const [trim,setTrim]=useState(true);const [ignoreCase,setIgnoreCase]=useState(false);const [sheet,setSheet]=useState('All sheets');const [drag,setDrag]=useState(-1);
 async function upload(file:File|undefined,index:number){if(!file)return;setError('');if(!/\.(xlsx|xls|csv)$/i.test(file.name)){setError('Please choose an .xlsx, .xls, or .csv file.');return;}if(file.size>20*1024*1024){setError('Please choose a file smaller than 20 MB.');return;}try{const book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});if(!book.SheetNames.length)throw Error();setBooks(prev=>prev.map((b,i)=>i===index?{name:file.name,book}:b));setResult(null);}catch{setError('This file could not be read. Please check that it is a valid Excel workbook.');}}
 function run(){if(!books[0]||!books[1])return;setError('');if(!(books[0].book.SheetNames.length===1&&books[1].book.SheetNames.length===1)&&!books[0].book.SheetNames.some(name=>books[1]!.book.SheetNames.includes(name))){setError('The files have no matching sheet names. Give the sheets you want to compare the same name, save both files, and select them again.');setResult(null);return;}setBusy(true);setTimeout(()=>{try{setResult(compare(books[0]!.book,books[1]!.book,trim,ignoreCase));setSheet('All sheets');}catch(error){setResult(null);setError(error instanceof Error?error.message:'Unable to compare these workbooks.');}finally{setBusy(false);}},60);}
 function demo(){const first=[['Unit No.','Unit Description','Plot Size','Status'],['Notebook','Stationery',12,120],['Desk lamp','Lighting',45,32],['Water bottle','Accessories',18,64],['Canvas tote','Accessories',24,80],['Pen set','Stationery',9,150]];const second=[['Unit No.','Unit Description','Plot Size','Status'],['Notebook','Stationery',14,120],['Desk lamp','Lighting',45,28],['Water bottle','Accessories',18,64],['Canvas tote','Accessories',24,95],['Pen set','Stationery',9,150],['Desk organizer','Stationery',22,40]];const make=(data:(string|number)[][])=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet(data),'Inventory');return b;};const a=make(first),b=make(second);setBooks([{name:'inventory_original.xlsx',book:a},{name:'inventory_updated.xlsx',book:b}]);setResult(compare(a,b,trim,ignoreCase));setSheet('All sheets');setError('');}
 const rows=result?.diffs.filter(d=>sheet==='All sheets'||d.sheet===sheet)||[];
 function download(){if(!result)return;const csv=XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(result.diffs.map(d=>({Sheet:d.sheet,'Unit No.':d.unit,'Original cell':d.cell,'Updated cell':d.rightCell,Field:d.field,'Original value':d.old,'Updated value':d.value}))));const url=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'}));const a=document.createElement('a');a.href=url;a.download='sheetcompare-results.csv';a.click();URL.revokeObjectURL(url);}
 return <div className="app"><header><a className="brand" href="./"><span className="brand-icon"><Files size={21}/></span>SheetCompare<span className="beta">BETA</span></a><div className="header-right"><span><LockKeyhole size={14}/>100% private. Always.</span><a href="#how-it-works">How it works <ArrowRight size={14}/></a></div></header>
 <main><div className="intro"><div className="eyebrow"><span/> LESS CHECKING. MORE CLARITY.</div><h1>Two spreadsheets.<br/>Every difference, <span>spotted.</span></h1><p>Compare Excel files in seconds. See exactly what changed,<br className="desktop"/> without the endless scrolling.</p></div>
 <section className="workspace"><div className="section-top"><div><span className="step">01</span><h2>Add your spreadsheets</h2></div><span className="formats">.xlsx, .xls, or .csv <span>·</span> Up to 20 MB each</span></div><div className="uploads">{[0,1].map(i=><React.Fragment key={i}>{i===1&&<div className="between"><ArrowLeftRight size={18}/></div>}<div className={'upload '+(drag===i?'dragging ':'')+(books[i]?'loaded':'')} onDragOver={e=>{e.preventDefault();setDrag(i);}} onDragLeave={()=>setDrag(-1)} onDrop={e=>{e.preventDefault();setDrag(-1);void upload(e.dataTransfer.files[0],i);}}><div className="upload-heading"><span className={'file-index index-'+i}>{i===0?'A':'B'}</span><strong>{i===0?'Original file':'Updated file'}</strong><span>{i===0?'Your starting point':'Your latest version'}</span></div><label className="drop-zone"><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{void upload(e.target.files?.[0],i);e.target.value='';}}/><span className="upload-icon">{books[i]?<FileSpreadsheet size={25}/>:<Upload size={23}/>}</span><strong>{books[i]?books[i]!.name:<>Drop your file here, or <em>browse</em></>}</strong><small>{books[i]?`${books[i]!.book.SheetNames.length} sheet${books[i]!.book.SheetNames.length===1?'':'s'} · Click to replace`:'Excel workbooks & CSV files'}</small></label>{books[i]&&<button className="remove" aria-label={'Remove '+(i===0?'original':'updated')+' file'} onClick={()=>{setBooks(prev=>prev.map((b,j)=>i===j?null:b));setResult(null);}}><X size={15}/></button>}</div></React.Fragment>)}</div><p className="file-notice">Matching by Unit No. Comparing Status, Unit Description, Plot Size, Width, Length, and Dimension Unit, regardless of row order. Save and reselect files after editing.</p><div className="options"><div><label><input type="checkbox" checked={trim} onChange={e=>{setTrim(e.target.checked);setResult(null);}}/>Ignore extra spaces</label><label><input type="checkbox" checked={ignoreCase} onChange={e=>{setIgnoreCase(e.target.checked);setResult(null);}}/>Ignore letter case</label></div><span><ShieldCheck size={14}/> Files stay in your browser</span></div>{error&&<div role="alert" className="error">{error}</div>}<div className="compare-bar"><span><Layers3 size={16}/>Only your selected unit details.</span><button className="primary" disabled={!books[0]||!books[1]||busy} onClick={run}>{busy?'Comparing…':'Compare files'}<ArrowRight size={17}/></button></div></section>
 {!result?<><div className="sample">Just looking around? <button onClick={demo}>Try a sample comparison <ArrowRight size={14}/></button></div><section className="features" id="how-it-works"><article><span><FileSpreadsheet size={21}/></span><h3>Nothing slips through</h3><p>Find cells with different values<br/> in both files, across every sheet.</p></article><article><span><Sparkles size={21}/></span><h3>Differences, made obvious</h3><p>Clear, color-coded results let you see<br/> what matters at a glance.</p></article><article><span><ShieldCheck size={21}/></span><h3>Your files are yours</h3><p>Everything happens in your browser.<br/> No uploads, accounts, or worries.</p></article></section></>:<section className="results"><div className="result-title"><div><span className="step">02</span><h2>Comparison results</h2><span className="complete"><Check size={12}/> Complete</span></div><button className="export" onClick={download}><ArrowDownToLine size={15}/> Export CSV</button></div><div className="unit-summary"><strong>Units with differences ({new Set(rows.map(d=>d.unit)).size})</strong><div>{Array.from(new Set(rows.map(d=>d.unit))).map(unit=><span key={unit}>Unit {unit}</span>)}</div></div><div className="stats">{[['Different cells',result.diffs.length,'changed']].map(([label,count,cls])=><div className={String(cls)} key={String(label)}><span>{label}</span><strong>{count}</strong></div>)}</div><div className="table-toolbar"><span>{result.total.toLocaleString()} cells checked · {result.sheets.length} sheets</span><div><select aria-label="Filter by sheet" value={sheet} onChange={e=>setSheet(e.target.value)}>{['All sheets',...result.sheets].map(s=><option key={s}>{s}</option>)}</select></div></div><div className="table-scroll"><table><thead><tr><th>Sheet</th><th>Unit No.</th><th>Original cell</th><th>Updated cell</th><th>Field</th><th>Original value</th><th>Updated value</th></tr></thead><tbody>{rows.slice(0,1000).map(d=><tr key={d.sheet+d.cell}><td>{d.sheet}</td><td><strong>Unit {d.unit}</strong></td><td><code>{d.cell}</code></td><td><code>{d.rightCell}</code></td><td>{d.field}</td><td className="old-value">{d.old||<span className="empty">Empty</span>}</td><td className="new-value">{d.value||<span className="empty">Empty</span>}</td></tr>)}</tbody></table>{!rows.length&&<div className="no-results"><Check size={26}/><strong>{result.diffs.length?'No changes match these filters.':'No differing values found in cells filled in both files.'}</strong></div>}</div>{rows.length>1000&&<p className="limit">Showing the first 1,000 differences. Export CSV for the complete results.</p>}<p className="result-note">Only cells with values in both files are shown when their values differ. Single-sheet files are compared directly, regardless of sheet name. Multi-sheet files match by sheet name. Columns match by field name; rows match by Unit No., regardless of row order. Units found in only one file are excluded. Formatting and formulas with the same saved value are not compared.</p></section>}
 <footer><span className="footer-brand"><Files size={15}/> SheetCompare</span><span>A little less manual work. A lot more peace of mind.</span><span><LockKeyhole size={12}/> Private by design</span></footer></main></div>
}
createRoot(document.getElementById('root')!).render(<React.StrictMode><App/></React.StrictMode>);






