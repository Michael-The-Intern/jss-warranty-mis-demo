import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';

// ══════════════════════════════════════════════════════
// SHARED PRIMITIVES / WARRANTY MIS CONSTANTS (Increment 0A)
// Declarations only — no reads/writes of wmis.* keys in this increment.
// Naming rule: Recall Database uses `jss_*_v2` legacy keys; Warranty MIS
// will use `wmis.*.v1` keys when domain-specific blocks are implemented.
// ══════════════════════════════════════════════════════
const DOMAIN = Object.freeze({
  RECALL_INTELLIGENCE: 'RECALL_INTELLIGENCE',
  GM_WARRANTY:         'GM_WARRANTY',
  STELLANTIS_RECOVERY: 'STELLANTIS_RECOVERY',
  PARTS_CROSSREF:      'PARTS_CROSSREF',
  OPERATIONS:          'OPERATIONS',
});
const OEM = Object.freeze({
  GM: 'GM', STELLANTIS: 'STELLANTIS', FORD: 'FORD', OTHER: 'OTHER',
});
const FIELD_ORIGIN = Object.freeze({
  IMPORTED: 'IMPORTED', DERIVED: 'DERIVED', USER: 'USER', SEED: 'SEED',
});
const PART_MAP_STATUS = Object.freeze({
  MAPPED: 'MAPPED', UNMAPPED: 'UNMAPPED', ASSUMED: 'ASSUMED', CONFLICT: 'CONFLICT',
});
const STORAGE_KEYS = Object.freeze({
  nhtsa:      Object.freeze({ recalls: 'wmis.nhtsa.recalls.v1' }),
  gm:         Object.freeze({ bills: 'wmis.gm.bills.v1', claims: 'wmis.gm.claims.v1' }),
  stellantis: Object.freeze({
    cases:          'wmis.stellantis.cases.v1',
    swrs:           'wmis.stellantis.swrs.v1',
    debits:         'wmis.stellantis.debits.v1',
    claimAlloc:     'wmis.stellantis.claimAllocations.v1',
    costRecovery:   'wmis.stellantis.costRecoveryLines.v1',
  }),
  parts:      Object.freeze({ crossRef: 'wmis.parts.crossref.v1' }),
  import_:    Object.freeze({ batches: 'wmis.import.batches.v1' }),
  audit:      Object.freeze({ wmis: 'wmis.audit.v1' }),
});
// ══════════════════════════════════════════════════════
// SHARED PRIMITIVES (existing) — hooks, atoms, seed data, legacy helpers.
// ══════════════════════════════════════════════════════


const C = {
  // Canvas — BEIGE PRESERVED
  bg:'#F5F4F0', panel:'#FFFFFF', panelAlt:'#FBFAF7',
  // Borders — warm beige family retained
  border:'#E8E4DA', borderSoft:'#F0EDE5', borderStrong:'#D9D3C5',
  // Text hierarchy (4 levels)
  ink:'#1B2A5E', inkSub:'#3B4A73', inkMute:'#8A8579', inkFaint:'#B5AE9E',
  navy:'#1B2A5E',coral:'#E8551F',coralSoft:'#FF7A3D',
  navBg:'#1B2A5E',navBgDeep:'#0F1A3D',navText:'#C7CFE4',
  green:'#3F8B4E', amber:'#D97706', rose:'#C8102E', info:'#2C4A9E',
  tileBlue:'#E8ECF5',tilePeach:'#FCE8DE',tileGreen:'#E6EFE6',tilePurple:'#ECE8F5',
  // Semantic tint surfaces for badges
  tintNavy:'#E8ECF5', tintAmber:'#FBEFD9', tintGreen:'#E6EFE6', tintRose:'#F7E1E4', tintNeutral:'#F0EDE5',
  focusRing:'rgba(232,85,31,0.25)',
};
// Typography scale — enterprise hierarchy
const T = {
  pageTitle:{fontSize:'21px',fontWeight:600,letterSpacing:'-0.015em',lineHeight:1.2},
  section:  {fontSize:'15px',fontWeight:600,letterSpacing:'-0.005em'},
  cardHead: {fontSize:'13px',fontWeight:600},
  body:     {fontSize:'13px',fontWeight:400},
  metadata: {fontSize:'12px',fontWeight:400,color:'#8A8579'},
  caption:  {fontSize:'11px',fontWeight:600,letterSpacing:'0.06em',textTransform:'uppercase',color:'#8A8579'},
  num:      {fontVariantNumeric:'tabular-nums',fontFeatureSettings:'"tnum"'},
};
// Spacing / radius / elevation — rounded enterprise (buttons/inputs 8px, cards 10px)
const S = {
  r:{sm:'4px',md:'8px',lg:'10px',xl:'14px',pill:'999px'},
  elev1:'0 1px 2px rgba(27,42,94,0.04), 0 1px 3px rgba(27,42,94,0.06)',
  elev2:'0 4px 8px -2px rgba(27,42,94,0.08), 0 2px 4px -2px rgba(27,42,94,0.04)',
  elev3:'0 12px 24px -8px rgba(27,42,94,0.18), 0 4px 8px -4px rgba(27,42,94,0.08)',
};

// ── VIN ENGINE ──
const VIN_VALS = {0:0,1:1,2:2,3:3,4:4,5:5,6:6,7:7,8:8,9:9,A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,J:1,K:2,L:3,M:4,N:5,P:7,R:9,S:2,T:3,U:4,V:5,W:6,X:7,Y:8,Z:9};
const VIN_WEIGHTS = [8,7,6,5,4,3,2,10,0,9,8,7,6,5,4,3,2];
const VALID_CHARS = '0123456789ABCDEFGHJKLMNPRSTUVWXYZ';

function validateVIN(vin) {
  if (!vin || vin.length !== 17) return { valid:false, error:'Must be 17 characters ('+( vin?vin.length:0)+' given)' };
  const v = vin.toUpperCase();
  for (let c of v) { if (!VALID_CHARS.includes(c)) return { valid:false, error:'Invalid char "'+c+'" — I, O, Q not allowed' }; }
  let sum = 0;
  for (let i=0;i<17;i++) sum += (VIN_VALS[v[i]]||0)*VIN_WEIGHTS[i];
  const rem = sum%11;
  const check = rem===10?'X':String(rem);
  return { valid:true, checkOk:v[8]===check, warning:v[8]!==check?'Check digit mismatch (pos 9: got '+v[8]+', expected '+check+')':null };
}

const WMI_TABLE = [
  ['5YJ','TESLA','Tesla'],['1G1','GM','Chevrolet'],['1G2','GM','Pontiac'],['1G6','GM','Cadillac'],['1GC','GM','GMC'],['2G1','GM','Chevrolet'],
  ['1FA','FORD','Ford'],['1FB','FORD','Ford'],['1FC','FORD','Ford'],['1FD','FORD','Ford'],['1FM','FORD','Ford'],['1FT','FORD','Ford'],['3FA','FORD','Ford MX'],
  ['1C3','STELLANTIS','Chrysler'],['1C4','STELLANTIS','Jeep'],['1C6','STELLANTIS','Ram'],['2C3','STELLANTIS','Chrysler CA'],['3C4','STELLANTIS','Chrysler MX'],
  ['1HG','HONDA','Honda'],['2HG','HONDA','Honda CA'],['JHM','HONDA','Honda JP'],
  ['WBA','BMW','BMW'],['WBS','BMW','BMW M'],['WBY','BMW','BMW'],
  ['WDB','MERCEDES','Mercedes'],['WDD','MERCEDES','Mercedes'],
  ['WVW','VW','VW DE'],['1VW','VW','VW'],['3VW','VW','VW MX'],
  ['1N4','NISSAN','Nissan'],['JN1','NISSAN','Nissan JP'],
  ['4T1','TOYOTA','Toyota'],['5TD','TOYOTA','Toyota'],['2T1','TOYOTA','Toyota CA'],['JT2','TOYOTA','Toyota JP'],
  ['KM8','HYUNDAI/KIA','Hyundai'],['KNA','HYUNDAI/KIA','Kia'],
  ['YV1','VOLVO','Volvo'],['YV4','VOLVO','Volvo'],
  ['WAU','AUDI','Audi'],['WA1','AUDI','Audi'],
];
const MY_MAP = {A:2010,B:2011,C:2012,D:2013,E:2014,F:2015,G:2016,H:2017,J:2018,K:2019,L:2020,M:2021,N:2022,P:2023,R:2024,S:2025,T:2026,V:2027,'1':2001,'2':2002,'3':2003,'4':2004,'5':2005,'6':2006,'7':2007,'8':2008,'9':2009};

function decodeVIN(vin) {
  if (!vin||vin.length<10) return null;
  const v = vin.toUpperCase();
  const hit = WMI_TABLE.find(w=>v.startsWith(w[0])) || WMI_TABLE.find(w=>v.startsWith(w[0].slice(0,2)));
  return { oem:hit?hit[1]:'UNKNOWN', make:hit?hit[2]:'Unknown', year:MY_MAP[v[9]]||null };
}

// ── OEM PIE CHART ──
function OEMPieChart({vins}){
  const counts={};
  vins.forEach(v=>{ if(v.oem) counts[v.oem]=(counts[v.oem]||0)+1; });
  const entries=Object.entries(counts).sort((a,b)=>b[1]-a[1]);
  const total=entries.reduce((s,[,v])=>s+v,0);
  if(total===0) return <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'140px',color:'#8A8579',fontSize:'12px'}}>No VIN data yet</div>;
  const COLORS=['#1B2A5E','#E8551F','#3F8B4E','#1C69D4','#D97706','#C8102E','#0057A8','#910A2D','#6B7A7E','#002C5F'];
  let cum=0;
  const cx=90,cy=90,ro=72,ri=40;
  const slices=entries.map(([oem,cnt],i)=>{
    const pct=cnt/total;
    const sa=cum*2*Math.PI-Math.PI/2;
    const ea=(cum+pct)*2*Math.PI-Math.PI/2;
    cum+=pct;
    const x1=cx+ro*Math.cos(sa),y1=cy+ro*Math.sin(sa);
    const x2=cx+ro*Math.cos(ea),y2=cy+ro*Math.sin(ea);
    const ix1=cx+ri*Math.cos(sa),iy1=cy+ri*Math.sin(sa);
    const ix2=cx+ri*Math.cos(ea),iy2=cy+ri*Math.sin(ea);
    const lg=pct>0.5?1:0;
    const d=`M${ix1},${iy1} L${x1},${y1} A${ro},${ro} 0 ${lg},1 ${x2},${y2} L${ix2},${iy2} A${ri},${ri} 0 ${lg},0 ${ix1},${iy1} Z`;
    return {oem,cnt,pct,d,color:COLORS[i%COLORS.length]};
  });
  return(
    <div style={{display:'flex',alignItems:'center',gap:'16px'}}>
      <svg width="180" height="180" viewBox="0 0 180 180" style={{flexShrink:0}}>
        {slices.map((s,i)=><path key={i} d={s.d} fill={s.color} stroke="white" strokeWidth="1.5"/>)}
        <text x="90" y="86" textAnchor="middle" fontSize="18" fontWeight="700" fill="#1B2A5E">{total}</text>
        <text x="90" y="100" textAnchor="middle" fontSize="9" fill="#8A8579">TOTAL VINs</text>
      </svg>
      <div style={{display:'flex',flexDirection:'column',gap:'5px',flex:1}}>
        {slices.map((s,i)=>(
          <div key={i} style={{display:'flex',alignItems:'center',gap:'7px',fontSize:'11px'}}>
            <span style={{width:'8px',height:'8px',borderRadius:'2px',background:s.color,flexShrink:0}}/>
            <span style={{color:'#1B2A5E',fontWeight:500,flex:1}}>{s.oem}</span>
            <span style={{color:'#8A8579'}}>{s.cnt}</span>
            <span style={{color:'#8A8579',minWidth:'32px',textAlign:'right'}}>{Math.round(s.pct*100)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── SHEETJS LOADER (Vite runtime-validation compatibility) ──
function useXLSX() { return true; }
// ── COLUMN AUTO-DETECTION ──
const VIN_RE = /^[A-HJ-NPR-Z0-9]{17}$/i;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$|^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/;
const NUM_RE = /^\d+(\.\d+)?$/;

function scoreColumn(col, values) {
  const sample = values.filter(Boolean).slice(0,20);
  if(!sample.length) return {};
  const scores = {};
  const vinMatches = sample.filter(v=>VIN_RE.test(String(v).trim())).length;
  const dateMatches = sample.filter(v=>DATE_RE.test(String(v).trim())).length;
  const numMatches = sample.filter(v=>NUM_RE.test(String(v).trim())).length;
  const colL = String(col).toLowerCase();
  // VIN score
  scores.vin = (vinMatches/sample.length)*70 + (colL.includes('vin')?30:colL.includes('vehicle_id')?20:0);
  // Date score
  scores.date = (dateMatches/sample.length)*60 + (colL.includes('date')?30:colL.includes('repair')?20:colL.includes('time')?10:0);
  // OEM score
  const oemKw = ['oem','make','brand','manufacturer','customer'];
  scores.oem = oemKw.some(k=>colL.includes(k)) ? 80 : (sample.filter(v=>OEMS.includes(String(v).toUpperCase())).length/sample.length)*70;
  // Campaign score
  const campKw = ['campaign','recall','program','id','code'];
  scores.campaign = campKw.some(k=>colL.includes(k)) ? 70 : 0;
  // Part/JSS PN score
  const pnKw = ['jss','part','pn','number','component'];
  scores.partPN = pnKw.some(k=>colL.includes(k)) ? 70 : 0;
  // Price score
  const priceKw = ['price','cost','rate','agreed','pc_price','pcprice'];
  scores.price = (priceKw.some(k=>colL.includes(k)) ? 60 : 0) + (numMatches/sample.length)*30;
  // Vehicle score
  const vehKw = ['vehicle','model','car','truck','program'];
  scores.vehicle = vehKw.some(k=>colL.includes(k)) ? 70 : 0;
  // Description score
  const descKw = ['desc','description','type','component','part_name','name'];
  scores.description = descKw.some(k=>colL.includes(k)) ? 70 : 0;
  return scores;
}

function autoDetectColumns(headers, rows) {
  const mapping = {};
  const taken = new Set();
  const fields = ['vin','date','oem','campaign','partPN','price','vehicle','description'];
  // Score each header for each field
  const scoreboard = headers.map(h=>{
    const vals = rows.map(r=>r[h]);
    return {header:h, scores:scoreColumn(h,vals)};
  });
  fields.forEach(field=>{
    const best = scoreboard
      .filter(s=>!taken.has(s.header))
      .sort((a,b)=>(b.scores[field]||0)-(a.scores[field]||0))[0];
    if(best && (best.scores[field]||0)>25){
      mapping[field]=best.header;
      taken.add(best.header);
    }
  });
  return mapping;
}

function detectFileType(headers, rows, mapping) {
  // VIN file: has a VIN column with real VINs
  if(mapping.vin){
    const sample=rows.slice(0,10).map(r=>r[mapping.vin]);
    const vinCount=sample.filter(v=>VIN_RE.test(String(v||'').trim())).length;
    if(vinCount>0) return 'VIN_LIST';
  }
  // Parts file: has JSS PN / price / OEM columns
  const pnKw=['jss','part','pn'];
  const hasPNCol=headers.some(h=>pnKw.some(k=>String(h).toLowerCase().includes(k)));
  const hasPriceCol=headers.some(h=>['price','cost','rate'].some(k=>String(h).toLowerCase().includes(k)));
  if(hasPNCol||hasPriceCol) return 'PARTS_REF';
  return 'UNKNOWN';
}

function normalizeDate(v) {
  if(!v) return '';
  const s = String(v).trim();
  if(DATE_RE.test(s)){
    // Convert M/D/YYYY or D-M-YYYY to YYYY-MM-DD
    const parts = s.split(/[\/\-]/);
    if(parts.length===3 && parts[0].length<=2){
      const [m,d,y] = parts;
      const yr = y.length===2?'20'+y:y;
      return yr+'-'+m.padStart(2,'0')+'-'+d.padStart(2,'0');
    }
    return s;
  }
  // Excel serial date
  if(NUM_RE.test(s)){
    const serial=parseInt(s);
    if(serial>40000&&serial<50000){
      const d=new Date((serial-25569)*86400*1000);
      return d.toISOString().slice(0,10);
    }
  }
  return '';
}

// ── SMART IMPORTER COMPONENT ──
function SmartImporter({camps,parts,vinCounts,onImportVINs,onImportParts,onClose,xlsxReady,toast$}){
  const [stage,setStage]=useState('upload'); // upload | sheet | mapping | preview | done
  const [sheets,setSheets]=useState([]);
  const [selSheet,setSelSheet]=useState('');
  const [rawRows,setRawRows]=useState([]);
  const [headers,setHeaders]=useState([]);
  const [mapping,setMapping]=useState({});
  const [fileType,setFileType]=useState('UNKNOWN');
  const [preview,setPreview]=useState([]);
  const [result,setResult]=useState(null);
  const [defCampId,setDefCampId]=useState('');
  const [wb,setWb]=useState(null);
  const [drag,setDrag]=useState(false);
  const [fileName,setFileName]=useState('');

  const parseSheet=(workbook,sheetName)=>{
    const ws=workbook.Sheets[sheetName];
    const json=XLSX.utils.sheet_to_json(ws,{defval:'',raw:false});
    if(!json.length){toast$('Sheet is empty','warn');return;}
    const hdrs=Object.keys(json[0]);
    setHeaders(hdrs);
    setRawRows(json);
    const mp=autoDetectColumns(hdrs,json);
    setMapping(mp);
    const ft=detectFileType(hdrs,json,mp);
    setFileType(ft);
    setStage('mapping');
  };

  const processFile=file=>{
    setFileName(file.name);
    const isCsv=file.name.toLowerCase().endsWith('.csv')||file.type==='text/csv';
    const reader=new FileReader();
    if(isCsv){
      reader.onload=ev=>{
        const text=ev.target.result;
        const lines=text.split('\n').filter(l=>l.trim());
        const hdrs=lines[0].split(',').map(s=>s.replace(/"/g,'').trim());
        const rows=lines.slice(1).map(line=>{
          const vals=line.split(',').map(s=>s.replace(/"/g,'').trim());
          const obj={};
          hdrs.forEach((h,i)=>obj[h]=vals[i]||'');
          return obj;
        }).filter(r=>Object.values(r).some(v=>v));
        setHeaders(hdrs);setRawRows(rows);
        const mp=autoDetectColumns(hdrs,rows);
        setMapping(mp);
        const ft=detectFileType(hdrs,rows,mp);
        setFileType(ft);
        setStage('mapping');
      };
      reader.readAsText(file);
    } else {
      if(!xlsxReady||!XLSX){toast$('Excel parser loading, try again in a moment','warn');return;}
      reader.onload=ev=>{
        const data=new Uint8Array(ev.target.result);
        const workbook=XLSX.read(data,{type:'array'});
        setWb(workbook);
        const sheetNames=workbook.SheetNames;
        if(sheetNames.length===1){
          parseSheet(workbook,sheetNames[0]);
        } else {
          setSheets(sheetNames);
          setSelSheet(sheetNames.find(s=>s!=='Sheet1'&&s!=='sheet1')||sheetNames[0]);
          setStage('sheet');
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const onDrop=e=>{
    e.preventDefault();setDrag(false);
    const f=e.dataTransfer.files[0];
    if(f) processFile(f);
  };
  const onFileInput=e=>{
    const f=e.target.files[0];
    if(f) processFile(f);
    e.target.value='';
  };

  const buildPreview=()=>{
    if(fileType==='VIN_LIST'){
      const rows=rawRows.slice(0,200);
      const prev=rows.map((r,i)=>{
        const vin=String(r[mapping.vin]||'').trim().toUpperCase();
        const date=normalizeDate(r[mapping.date]);
        const res=VIN_RE.test(vin)?validateVIN(vin):{valid:false,error:'Not a VIN'};
        const dec=VIN_RE.test(vin)?decodeVIN(vin):null;
        const isDupe=(vinCounts[vin]||0)>0;
        return {_idx:i,vin,date,oem:dec?.oem||r[mapping.oem]||'UNKNOWN',year:dec?.year,make:dec?.make,valid:res.valid,error:res.error,isDupe};
      });
      setPreview(prev);
    } else {
      const rows=rawRows.slice(0,200);
      const prev=rows.map((r,i)=>({
        _idx:i,
        oem:r[mapping.oem]||'',
        vehicle:r[mapping.vehicle]||'',
        jssPN:r[mapping.partPN]||'',
        description:r[mapping.description]||'',
        pcPrice:parseFloat(r[mapping.price])||0,
        valid:!!(r[mapping.partPN]||r[mapping.oem]),
      }));
      setPreview(prev.filter(p=>p.valid));
    }
    setStage('preview');
  };

  const doImport=()=>{
    if(fileType==='VIN_LIST'){
      const valid=preview.filter(p=>p.valid);
      const dupes=valid.filter(p=>p.isDupe).length;
      const newVins=valid.map((p,i)=>({
        id:Date.now()+i,
        vin:p.vin,
        campaignId:+defCampId||0,
        oem:p.oem,
        repairDate:p.date,
        type:'REPAIRED',
        addedDate:new Date().toISOString().slice(0,10),
      }));
      onImportVINs(newVins);
      setResult({type:'VIN_LIST',imported:valid.length,dupes,errors:preview.filter(p=>!p.valid).length});
    } else {
      const newParts=preview.filter(p=>p.valid).map((p,i)=>({
        id:Date.now()+i,
        oem:String(p.oem).toUpperCase()||'UNKNOWN',
        vehicle:p.vehicle,
        jssPN:p.jssPN,
        custPN:'',
        description:p.description,
        plant:'',
        volume:0,
        stdCost:0,
        pcPrice:p.pcPrice,
        sop:'',
        status:'ACTIVE',
      }));
      onImportParts(newParts);
      setResult({type:'PARTS_REF',imported:newParts.length,errors:0});
    }
    setStage('done');
  };

  const FIELD_LABELS={vin:'VIN Column',date:'Repair Date',oem:'OEM/Make',campaign:'Campaign',partPN:'JSS / Part PN',price:'PC Price / Rate',vehicle:'Vehicle / Model',description:'Description'};
  const relevantFields=fileType==='VIN_LIST'?['vin','date','oem']:['partPN','oem','vehicle','price','description'];

  return(
    <div style={{position:'fixed',inset:0,background:'rgba(15,26,61,0.65)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:9990,backdropFilter:'blur(3px)'}}>
      <Card style={{width:'720px',maxWidth:'95vw',maxHeight:'88vh',display:'flex',flexDirection:'column',overflow:'hidden'}}>
        {/* Header */}
        <div style={{padding:'18px 22px',borderBottom:`1px solid ${C.border}`,display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0}}>
          <div>
            <div style={{fontSize:'15px',fontWeight:700,color:C.ink}}>🔍 Smart File Importer</div>
            <div style={{fontSize:'11px',color:C.inkMute,marginTop:'2px'}}>Auto-detects VIN lists, parts references, or campaign data from Excel &amp; CSV</div>
          </div>
          <div style={{display:'flex',gap:'7px',alignItems:'center'}}>
            {stage!=='upload'&&stage!=='done'&&<Btn variant="ghost" size="sm" onClick={()=>setStage('upload')}>← Back</Btn>}
            <button onClick={onClose} style={{background:'none',border:'none',fontSize:'18px',cursor:'pointer',color:C.inkMute}}>✕</button>
          </div>
        </div>

        {/* Progress */}
        <div style={{padding:'10px 22px',background:C.bg,borderBottom:`1px solid ${C.borderSoft}`,display:'flex',gap:'6px',alignItems:'center',flexShrink:0}}>
          {['Upload','Detect','Map Columns','Preview & Import'].map((s,i)=>{
            const stageIdx={upload:0,sheet:1,mapping:1,preview:2,done:3}[stage];
            const active=i<=stageIdx;
            return <React.Fragment key={s}>
              {i>0&&<div style={{flex:1,height:'2px',background:active?C.navy:'#ddd',borderRadius:'2px'}}/>}
              <div style={{display:'flex',alignItems:'center',gap:'5px',flexShrink:0}}>
                <div style={{width:'20px',height:'20px',borderRadius:'50%',background:active?C.navy:'#ddd',color:active?'#fff':'#aaa',fontSize:'10px',fontWeight:700,display:'flex',alignItems:'center',justifyContent:'center'}}>{i+1}</div>
                <span style={{fontSize:'10.5px',fontWeight:active?600:400,color:active?C.ink:C.inkMute}}>{s}</span>
              </div>
            </React.Fragment>;
          })}
        </div>

        <div style={{flex:1,overflow:'auto',padding:'20px 22px'}}>
          {/* STAGE: UPLOAD */}
          {stage==='upload'&&(
            <div>
              <div
                onDragOver={e=>{e.preventDefault();setDrag(true);}}
                onDragLeave={()=>setDrag(false)}
                onDrop={onDrop}
                style={{border:`2px dashed ${drag?C.coral:C.border}`,borderRadius:'16px',padding:'48px 24px',textAlign:'center',background:drag?C.tilePeach:C.bg,transition:'all 0.2s',cursor:'pointer'}}
                onClick={()=>document.getElementById('sfi-input').click()}
              >
                <div style={{fontSize:'36px',marginBottom:'10px'}}>📊</div>
                <div style={{fontSize:'15px',fontWeight:600,marginBottom:'6px'}}>Drop your Excel or CSV file here</div>
                <div style={{fontSize:'12px',color:C.inkMute,marginBottom:'16px'}}>Supports .xlsx, .xls, .csv — any format, any column order</div>
                <Btn variant="accent">Choose File</Btn>
                <input id="sfi-input" type="file" accept=".xlsx,.xls,.csv,.txt" onChange={onFileInput} style={{display:'none'}}/>
              </div>
              <div style={{marginTop:'16px',padding:'14px',background:C.tileBlue,borderRadius:'12px',fontSize:'11.5px',color:C.inkMute,lineHeight:1.6}}>
                <div style={{fontWeight:600,color:C.ink,marginBottom:'4px'}}>🤖 What gets auto-detected:</div>
                <div>• <b>VIN list</b> — any file with a column of 17-character VINs (repair records)</div>
                <div>• <b>Parts reference</b> — files with JSS PN, OEM, vehicle, and price columns (like your Campaign Summary Excel)</div>
                <div>• Column names don't matter — the scanner reads the actual data values to figure it out</div>
              </div>
            </div>
          )}

          {/* STAGE: SHEET PICKER */}
          {stage==='sheet'&&(
            <div>
              <div style={{fontSize:'13px',fontWeight:600,marginBottom:'12px'}}>Select a sheet from <b>{fileName}</b></div>
              <div style={{display:'flex',flexDirection:'column',gap:'7px'}}>
                {sheets.map(s=>(
                  <button key={s} onClick={()=>setSelSheet(s)} style={{padding:'12px 16px',borderRadius:'10px',border:`2px solid ${selSheet===s?C.navy:C.border}`,background:selSheet===s?C.tileBlue:C.panel,cursor:'pointer',textAlign:'left',fontSize:'13px',fontWeight:selSheet===s?600:400}}>
                    📋 {s} {s==='Sheet1'||s==='sheet1'?<span style={{fontSize:'10px',color:C.inkMute}}>(usually empty)</span>:''}
                  </button>
                ))}
              </div>
              <Btn variant="accent" onClick={()=>parseSheet(wb,selSheet)} style={{marginTop:'14px',width:'100%'}}>Scan Selected Sheet →</Btn>
            </div>
          )}

          {/* STAGE: COLUMN MAPPING */}
          {stage==='mapping'&&(
            <div>
              <div style={{display:'flex',alignItems:'center',gap:'10px',marginBottom:'14px'}}>
                <div style={{padding:'6px 14px',borderRadius:'999px',fontSize:'11px',fontWeight:700,background:fileType==='VIN_LIST'?C.tileGreen:C.tileBlue,color:fileType==='VIN_LIST'?C.green:C.navy}}>
                  {fileType==='VIN_LIST'?'✓ VIN Repair List detected':'✓ Parts Reference detected'}
                </div>
                <button onClick={()=>setFileType(fileType==='VIN_LIST'?'PARTS_REF':'VIN_LIST')} style={{fontSize:'11px',color:C.inkMute,background:'none',border:'none',cursor:'pointer',textDecoration:'underline'}}>wrong? switch</button>
              </div>
              <div style={{fontSize:'11.5px',color:C.inkMute,marginBottom:'12px'}}><b>{rawRows.length}</b> rows found in <b>{fileName}</b>. Verify the column assignments below:</div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'9px',marginBottom:'16px'}}>
                {relevantFields.map(field=>(
                  <div key={field}>
                    <div style={{fontSize:'10.5px',fontWeight:600,color:C.inkMute,textTransform:'uppercase',letterSpacing:'0.07em',marginBottom:'4px'}}>{FIELD_LABELS[field]}</div>
                    <Sel value={mapping[field]||''} onChange={e=>setMapping(m=>({...m,[field]:e.target.value||undefined}))} style={{width:'100%'}}>
                    <option value="">(not mapped)</option>
                    {headers.map(h=><option key={h} value={h}>{h}</option>)}
                    </Sel>
                    {mapping[field]&&<div style={{fontSize:'10px',color:C.inkMute,marginTop:'2px'}}>
                    Sample: {rawRows.slice(0,3).map(r=>r[mapping[field]]).filter(Boolean).join(' · ')}
                    </div>}
                  </div>
                ))}
              </div>
              {fileType==='VIN_LIST'&&(
                <div style={{marginBottom:'14px'}}>
                  <div style={{fontSize:'10.5px',fontWeight:600,color:C.inkMute,textTransform:'uppercase',letterSpacing:'0.07em',marginBottom:'4px'}}>Assign to Campaign</div>
                  <Sel value={defCampId} onChange={e=>setDefCampId(e.target.value)} style={{width:'100%'}}>
                    <option value="">No campaign (assign later)</option>
                    {camps.map(c=><option key={c.id} value={c.id}>{c.name} — {c.oem}</option>)}
                  </Sel>
                </div>
              )}
              <Btn variant="accent" onClick={buildPreview} style={{width:'100%'}}>Preview {rawRows.length} Rows →</Btn>
            </div>
          )}

          {/* STAGE: PREVIEW */}
          {stage==='preview'&&(
            <div>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'10px'}}>
                <div style={{fontSize:'13px',fontWeight:600}}>
                  {fileType==='VIN_LIST'
                    ? preview.filter(p=>p.valid).length+' valid VINs ready to import'
                    : preview.length+' parts rows ready to import'}
                  {fileType==='VIN_LIST'&&preview.filter(p=>p.isDupe).length>0&&
                    <span style={{color:C.rose,marginLeft:'8px'}}>· {preview.filter(p=>p.isDupe).length} duplicates</span>}
                  {preview.filter(p=>!p.valid).length>0&&
                    <span style={{color:C.amber,marginLeft:'8px'}}>· {preview.filter(p=>!p.valid).length} invalid (will skip)</span>}
                </div>
              </div>
              <div style={{maxHeight:'320px',overflow:'auto',border:`1px solid ${C.border}`,borderRadius:'10px'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                  <thead style={{position:'sticky',top:0,background:C.bg}}>
                    <tr>
                    {fileType==='VIN_LIST'
                    ?['VIN','Year/Make','OEM','Repair Date','Status'].map(h=><th key={h} style={{padding:'7px 12px',textAlign:'left',fontSize:'9px',fontWeight:700,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`}}>{h}</th>)
                    :['JSS PN','OEM','Vehicle','Description','PC Price','Status'].map(h=><th key={h} style={{padding:'7px 12px',textAlign:'left',fontSize:'9px',fontWeight:700,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`}}>{h}</th>)
                    }
                    </tr>
                  </thead>
                  <tbody>
                    {preview.map((p,i)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${C.borderSoft}`,background:!p.valid?'rgba(200,16,46,0.04)':p.isDupe?'rgba(217,119,6,0.04)':''}}>
                    {fileType==='VIN_LIST'?<>
                    <td style={{padding:'6px 12px',fontFamily:'monospace',fontSize:'10.5px',fontWeight:500}}>{p.vin}</td>
                    <td style={{padding:'6px 12px',color:C.inkMute}}>{p.year?p.year+' '+p.make:'—'}</td>
                    <td style={{padding:'6px 12px'}}>{p.oem&&p.oem!=='UNKNOWN'?<Pill color={oemColor(p.oem)}>{p.oem}</Pill>:<span style={{color:C.inkMute}}>—</span>}</td>
                    <td style={{padding:'6px 12px',color:C.inkMute}}>{p.date||'—'}</td>
                    <td style={{padding:'6px 12px'}}>
                    {!p.valid?<Pill color={C.rose}>✗ {p.error}</Pill>:p.isDupe?<Pill color={C.amber}>⚠ Duplicate</Pill>:<Pill color={C.green}>✓ Valid</Pill>}
                    </td>
                    </>:<>
                    <td style={{padding:'6px 12px',fontFamily:'monospace',fontSize:'10.5px',fontWeight:600,color:C.navy}}>{p.jssPN||'—'}</td>
                    <td style={{padding:'6px 12px'}}>{p.oem?<Pill color={oemColor(p.oem.toUpperCase())}>{p.oem}</Pill>:'—'}</td>
                    <td style={{padding:'6px 12px'}}>{p.vehicle||'—'}</td>
                    <td style={{padding:'6px 12px'}}>{p.description||'—'}</td>
                    <td style={{padding:'6px 12px',fontWeight:600,color:p.pcPrice?C.navy:C.inkMute}}>{p.pcPrice?fmt$d(p.pcPrice):'—'}</td>
                    <td style={{padding:'6px 12px'}}><Pill color={C.green}>✓</Pill></td>
                    </>}
                    </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{marginTop:'14px',display:'flex',gap:'9px'}}>
                <Btn variant="accent" onClick={doImport} style={{flex:1}}>
                  ✓ Import {fileType==='VIN_LIST'?preview.filter(p=>p.valid).length+' VINs':preview.length+' Parts'}
                </Btn>
                <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
              </div>
            </div>
          )}

          {/* STAGE: DONE */}
          {stage==='done'&&result&&(
            <div style={{textAlign:'center',padding:'32px 0'}}>
              <div style={{fontSize:'48px',marginBottom:'12px'}}>✅</div>
              <div style={{fontSize:'16px',fontWeight:700,marginBottom:'6px'}}>Import Complete!</div>
              <div style={{fontSize:'13px',color:C.inkMute,marginBottom:'20px'}}>
                {result.type==='VIN_LIST'
                  ?result.imported+' VINs imported · '+result.dupes+' duplicates flagged · '+result.errors+' skipped'
                  :result.imported+' parts added to Parts Reference'}
              </div>
              <Btn variant="accent" onClick={onClose}>Close</Btn>
            </div>
          )}
        </div>
      </Card>
    </div>
  );
}

// ── CSV EXPORT ──
function downloadCSV(rows, filename) {
  if (!rows||!rows.length) return;
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(','), ...rows.map(r=>headers.map(h=>JSON.stringify(r[h]??'')).join(','))].join('\n');
  const blob = new Blob([csv],{type:'text/csv'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href=url; a.download=filename+'_'+new Date().toISOString().slice(0,10)+'.csv'; a.click();
  URL.revokeObjectURL(url);
}

// ── LOCAL STORAGE ──
function useLS(key, def) {
  const [val, setVal] = useState(() => {
    try { const s=localStorage.getItem(key); return s?JSON.parse(s):def; } catch { return def; }
  });
  useEffect(()=>{ try { localStorage.setItem(key,JSON.stringify(val)); } catch {} },[val]);
  return [val, setVal];
}

// ── SEED DATA ──
const SEED_CAMPS = [
  {id:1,name:'STLA LX LD Retractor Recall',oem:'STELLANTIS',vehicle:'LX LD',jssPN:'2433862A6M',agreedRate:193.08,maxVINs:5000,takeRate:72,sop:'2026-04-01',status:'IN PROGRESS',statusLog:[{status:'OPEN',ts:'2026-04-01T00:00:00'},{status:'IN PROGRESS',ts:'2026-04-15T00:00:00'}]},
  {id:2,name:'Ford CD338/334 PAB Recall',oem:'FORD',vehicle:'FUSION',jssPN:'2541116-SAD',agreedRate:169.42,maxVINs:8200,takeRate:68,sop:'2026-02-10',status:'IN PROGRESS',statusLog:[{status:'OPEN',ts:'2026-02-10T00:00:00'},{status:'IN PROGRESS',ts:'2026-02-20T00:00:00'}]},
  {id:3,name:'BMW E7X PAB Recall',oem:'BMW',vehicle:'X7 X3',jssPN:'3104281001',agreedRate:174.73,maxVINs:3100,takeRate:55,sop:'2026-03-15',status:'IN PROGRESS',statusLog:[{status:'OPEN',ts:'2026-03-15T00:00:00'},{status:'IN PROGRESS',ts:'2026-03-22T00:00:00'}]},
  {id:4,name:'STLA Buckle Inflator Recall',oem:'STELLANTIS',vehicle:'Jeep Cherokee',jssPN:'JSS-A2',agreedRate:148.50,maxVINs:2800,takeRate:65,sop:'2025-11-01',status:'CLOSED',statusLog:[{status:'OPEN',ts:'2025-11-01T00:00:00'},{status:'IN PROGRESS',ts:'2025-11-15T00:00:00'},{status:'CLOSED',ts:'2026-01-30T00:00:00'}]},
  {id:5,name:'Ford F-150 MGG Recall',oem:'FORD',vehicle:'F-150',jssPN:'JSS-B1',agreedRate:212.00,maxVINs:4500,takeRate:71,sop:'2025-09-01',status:'CLOSED',statusLog:[{status:'OPEN',ts:'2025-09-01T00:00:00'},{status:'IN PROGRESS',ts:'2025-09-20T00:00:00'},{status:'CLOSED',ts:'2025-12-15T00:00:00'}]},
];
const SEED_PARTS = [
  {id:1,oem:'STELLANTIS',vehicle:'LX LD',jssPN:'2433862A6M',custPN:'SC-1001',description:'Retractor',plant:'TRN',volume:12000,stdCost:112.40,pcPrice:193.08,sop:'2026-04-01',status:'ACTIVE'},
  {id:2,oem:'STELLANTIS',vehicle:'Jeep Cherokee',jssPN:'JSS-A2',custPN:'SC-1002',description:'Buckle Inflator',plant:'TRN',volume:8500,stdCost:82.30,pcPrice:148.50,sop:'2025-11-01',status:'ACTIVE'},
  {id:3,oem:'FORD',vehicle:'Fusion',jssPN:'2541116-SAD',custPN:'FC-2001',description:'PAB Module',plant:'MVA',volume:3259,stdCost:95.10,pcPrice:169.42,sop:'2026-02-10',status:'ACTIVE'},
  {id:4,oem:'FORD',vehicle:'F-150',jssPN:'JSS-B1',custPN:'FC-2002',description:'Micro Gas Generator',plant:'MVA',volume:4392,stdCost:121.80,pcPrice:212.00,sop:'2025-09-01',status:'ACTIVE'},
  {id:5,oem:'BMW',vehicle:'X7 X3',jssPN:'3104281001',custPN:'BC-3001',description:'PAB E7X LHD',plant:'TRN',volume:113200,stdCost:106.55,pcPrice:174.73,sop:'2026-03-15',status:'ACTIVE'},
  {id:6,oem:'VW',vehicle:'Beetle',jssPN:'JSS-D1',custPN:'VC-4001',description:'DAB Module',plant:'TRN',volume:45000,stdCost:55.10,pcPrice:99.21,sop:'2026-05-18',status:'ACTIVE'},
  {id:7,oem:'STELLANTIS',vehicle:'Ram 1500',jssPN:'JSS-A3',custPN:'SC-1003',description:'Retractor Assembly',plant:'ACU',volume:6200,stdCost:88.75,pcPrice:155.20,sop:'2024-06-01',status:'INACTIVE'},
  {id:8,oem:'GM',vehicle:'Silverado',jssPN:'JSS-G1',custPN:'GC-5001',description:'Buckle Assembly',plant:'TRN',volume:9800,stdCost:44.20,pcPrice:78.90,sop:'2024-03-01',status:'INACTIVE'},
];
// Sample repaired VINs covering multiple campaigns and quarters
const SEED_VINS = [
  // Campaign 5 (Ford F-150 - CLOSED, Q3/Q4 2025)
  {id:101,vin:'1FTFW1ET5NFA12345',campaignId:5,oem:'FORD',repairDate:'2025-09-15',type:'REPAIRED',addedDate:'2025-09-15'},
  {id:102,vin:'1FTFW1ET5NFA12346',campaignId:5,oem:'FORD',repairDate:'2025-09-18',type:'REPAIRED',addedDate:'2025-09-18'},
  {id:103,vin:'1FTFW1ET5NFA12347',campaignId:5,oem:'FORD',repairDate:'2025-09-22',type:'REPAIRED',addedDate:'2025-09-22'},
  {id:104,vin:'1FTFW1ET5NFA12348',campaignId:5,oem:'FORD',repairDate:'2025-10-05',type:'REPAIRED',addedDate:'2025-10-05'},
  {id:105,vin:'1FTFW1ET5NFA12349',campaignId:5,oem:'FORD',repairDate:'2025-10-12',type:'REPAIRED',addedDate:'2025-10-12'},
  {id:106,vin:'1FTFW1ET5NFA12350',campaignId:5,oem:'FORD',repairDate:'2025-11-03',type:'REPAIRED',addedDate:'2025-11-03'},
  {id:107,vin:'1FTFW1ET5NFA12351',campaignId:5,oem:'FORD',repairDate:'2025-11-14',type:'REPAIRED',addedDate:'2025-11-14'},
  {id:108,vin:'1FTFW1ET5NFA12352',campaignId:5,oem:'FORD',repairDate:'2025-12-02',type:'REPAIRED',addedDate:'2025-12-02'},
  // Campaign 4 (STLA Buckle - CLOSED, Q4 2025 / Q1 2026)
  {id:109,vin:'1C4RJFBG5NC123456',campaignId:4,oem:'STELLANTIS',repairDate:'2025-11-10',type:'REPAIRED',addedDate:'2025-11-10'},
  {id:110,vin:'1C4RJFBG5NC123457',campaignId:4,oem:'STELLANTIS',repairDate:'2025-11-20',type:'REPAIRED',addedDate:'2025-11-20'},
  {id:111,vin:'1C4RJFBG5NC123458',campaignId:4,oem:'STELLANTIS',repairDate:'2025-12-08',type:'REPAIRED',addedDate:'2025-12-08'},
  {id:112,vin:'1C4RJFBG5NC123459',campaignId:4,oem:'STELLANTIS',repairDate:'2026-01-07',type:'REPAIRED',addedDate:'2026-01-07'},
  {id:113,vin:'1C4RJFBG5NC123460',campaignId:4,oem:'STELLANTIS',repairDate:'2026-01-15',type:'REPAIRED',addedDate:'2026-01-15'},
  // Campaign 2 (Ford Fusion - IN PROGRESS, Q1/Q2 2026)
  {id:114,vin:'3FA6P0H77GR123456',campaignId:2,oem:'FORD',repairDate:'2026-02-14',type:'REPAIRED',addedDate:'2026-02-14'},
  {id:115,vin:'3FA6P0H77GR123457',campaignId:2,oem:'FORD',repairDate:'2026-02-21',type:'REPAIRED',addedDate:'2026-02-21'},
  {id:116,vin:'3FA6P0H77GR123458',campaignId:2,oem:'FORD',repairDate:'2026-03-05',type:'REPAIRED',addedDate:'2026-03-05'},
  {id:117,vin:'3FA6P0H77GR123459',campaignId:2,oem:'FORD',repairDate:'2026-03-18',type:'REPAIRED',addedDate:'2026-03-18'},
  {id:118,vin:'3FA6P0H77GR123460',campaignId:2,oem:'FORD',repairDate:'2026-04-02',type:'REPAIRED',addedDate:'2026-04-02'},
  {id:119,vin:'3FA6P0H77GR123461',campaignId:2,oem:'FORD',repairDate:'2026-05-10',type:'REPAIRED',addedDate:'2026-05-10'},
  {id:120,vin:'3FA6P0H77GR123462',campaignId:2,oem:'FORD',repairDate:'2026-06-15',type:'REPAIRED',addedDate:'2026-06-15'},
  // Campaign 3 (BMW E7X - IN PROGRESS, Q2 2026)
  {id:121,vin:'WBA7R4C59LA123456',campaignId:3,oem:'BMW',repairDate:'2026-04-08',type:'REPAIRED',addedDate:'2026-04-08'},
  {id:122,vin:'WBA7R4C59LA123457',campaignId:3,oem:'BMW',repairDate:'2026-04-22',type:'REPAIRED',addedDate:'2026-04-22'},
  {id:123,vin:'WBA7R4C59LA123458',campaignId:3,oem:'BMW',repairDate:'2026-05-06',type:'REPAIRED',addedDate:'2026-05-06'},
  {id:124,vin:'WBA7R4C59LA123459',campaignId:3,oem:'BMW',repairDate:'2026-06-01',type:'REPAIRED',addedDate:'2026-06-01'},
  // Campaign 1 (STLA LX LD - IN PROGRESS, Q2 2026)
  {id:125,vin:'1C6RR7LT5NS123456',campaignId:1,oem:'STELLANTIS',repairDate:'2026-05-12',type:'REPAIRED',addedDate:'2026-05-12'},
  {id:126,vin:'1C6RR7LT5NS123457',campaignId:1,oem:'STELLANTIS',repairDate:'2026-05-20',type:'REPAIRED',addedDate:'2026-05-20'},
  {id:127,vin:'1C6RR7LT5NS123458',campaignId:1,oem:'STELLANTIS',repairDate:'2026-06-04',type:'REPAIRED',addedDate:'2026-06-04'},
  {id:128,vin:'1C6RR7LT5NS123459',campaignId:1,oem:'STELLANTIS',repairDate:'2026-07-01',type:'REPAIRED',addedDate:'2026-07-01'},
  // Recall-only VINs (not yet repaired)
  {id:201,vin:'1C4RJFBG5NC200001',campaignId:1,oem:'STELLANTIS',repairDate:'',type:'RECALL_ONLY',addedDate:'2026-04-05'},
  {id:202,vin:'1C4RJFBG5NC200002',campaignId:1,oem:'STELLANTIS',repairDate:'',type:'RECALL_ONLY',addedDate:'2026-04-05'},
  {id:203,vin:'3FA6P0H77GR200001',campaignId:2,oem:'FORD',repairDate:'',type:'RECALL_ONLY',addedDate:'2026-02-12'},
  {id:204,vin:'3FA6P0H77GR200002',campaignId:2,oem:'FORD',repairDate:'',type:'RECALL_ONLY',addedDate:'2026-02-12'},
  {id:205,vin:'WBA7R4C59LA200001',campaignId:3,oem:'BMW',repairDate:'',type:'RECALL_ONLY',addedDate:'2026-03-16'},
];

const OEMS = ['STELLANTIS','HONDA','GM','FORD','TOYOTA','NISSAN','SUBARU','HYUNDAI/KIA','VW','BMW','TESLA','MERCEDES','MAZDA','AUDI','RIVIAN','VOLVO','SLATE','UNKNOWN'];
const DEF_TR = {STELLANTIS:72,HONDA:65,GM:65,FORD:68,TOYOTA:70,NISSAN:62,SUBARU:60,'HYUNDAI/KIA':63,VW:60,BMW:55,TESLA:75,MERCEDES:58,MAZDA:60,AUDI:57,RIVIAN:50,VOLVO:55,SLATE:50,UNKNOWN:50};
const oemColor = o => ({STELLANTIS:'#1B2A5E',HONDA:'#C8102E',GM:'#0057A8',FORD:'#1B5FA8',TOYOTA:'#E8551F',NISSAN:'#C3002F',SUBARU:'#004B87','HYUNDAI/KIA':'#002C5F',VW:'#0046AD',BMW:'#1C69D4',TESLA:'#CC0000',MERCEDES:'#5B7F99',MAZDA:'#910A2D',AUDI:'#BB0A30',RIVIAN:'#6B7A7E',VOLVO:'#1C355E',SLATE:'#2A2A2A',UNKNOWN:'#8A8579'}[o]||'#8A8579');

const fmt$ = n=>'$'+(n||0).toLocaleString(undefined,{minimumFractionDigits:0,maximumFractionDigits:0});
const fmt$d = n=>'$'+(n||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtN = n=>(n||0).toLocaleString();
const getY = d=>d?new Date(d).getFullYear():null;
const getQ = d=>d?Math.floor(new Date(d).getMonth()/3)+1:null;
const now=new Date(); const CY=now.getFullYear(); const CQ=Math.floor(now.getMonth()/3)+1;

// ── ATOMS ──
const Card=({children,variant='default',style={},...p})=>{
  const v={
    default:{background:C.panel,border:`1px solid ${C.border}`,boxShadow:S.elev1},
    inset:  {background:C.panelAlt,border:`1px solid ${C.borderSoft}`,boxShadow:'none'},
    flat:   {background:C.panel,border:`1px solid ${C.border}`,boxShadow:'none'},
  }[variant]||{background:C.panel,border:`1px solid ${C.border}`,boxShadow:S.elev1};
  return <div {...p} style={{borderRadius:S.r.lg,...v,...style}}>{children}</div>;
};
const Btn=({children,variant='primary',size='md',style={},onFocus,onBlur,...p})=>{
  const vs={
    primary:{background:C.ink,color:'#fff',border:`1px solid ${C.ink}`},
    accent: {background:C.coral,color:'#fff',border:`1px solid ${C.coral}`},
    ghost:  {background:C.panel,color:C.ink,border:`1px solid ${C.border}`},
    danger: {background:C.panel,color:C.rose,border:`1px solid ${C.border}`},
    success:{background:C.green,color:'#fff',border:`1px solid ${C.green}`},
  };
  const ss={
    md:{height:'34px',padding:'0 14px',fontSize:'13px',borderRadius:S.r.md},
    sm:{height:'28px',padding:'0 10px',fontSize:'12px',borderRadius:S.r.md},
    xs:{height:'24px',padding:'0 8px', fontSize:'11px',borderRadius:S.r.sm},
  };
  return <button {...p}
    onFocus={e=>{e.currentTarget.style.boxShadow=`0 0 0 3px ${C.focusRing}`;onFocus?.(e);}}
    onBlur ={e=>{e.currentTarget.style.boxShadow='none';onBlur?.(e);}}
    style={{...(vs[variant]||vs.primary),...(ss[size]||ss.md),fontWeight:500,cursor:'pointer',
      display:'inline-flex',alignItems:'center',justifyContent:'center',gap:'6px',lineHeight:1,
      transition:'filter .12s, box-shadow .12s',...style}}>{children}</button>;
};
const _fieldBase={height:'34px',padding:'0 12px',borderRadius:S.r.md,background:C.panelAlt,border:`1px solid ${C.border}`,fontSize:'13px',outline:'none',color:C.ink,transition:'border-color .12s, box-shadow .12s',boxSizing:'border-box'};
const _fon =e=>{e.currentTarget.style.borderColor=C.coral;e.currentTarget.style.boxShadow=`0 0 0 3px ${C.focusRing}`;};
const _foff=e=>{e.currentTarget.style.borderColor=C.border;e.currentTarget.style.boxShadow='none';};
const Inp=({style={},onFocus,onBlur,...p})=>(
  <input {...p}
    onFocus={e=>{_fon(e);onFocus?.(e);}}
    onBlur ={e=>{_foff(e);onBlur?.(e);}}
    style={{..._fieldBase,...style}}/>);
const Sel=({children,style={},onFocus,onBlur,...p})=>(
  <select {...p}
    onFocus={e=>{_fon(e);onFocus?.(e);}}
    onBlur ={e=>{_foff(e);onBlur?.(e);}}
    style={{..._fieldBase,cursor:'pointer',paddingRight:'28px',appearance:'none',
      backgroundImage:`url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238A8579' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='6 9 12 15 18 9'/></svg>")`,
      backgroundRepeat:'no-repeat',backgroundPosition:'right 10px center',...style}}>{children}</select>);
const Pill=({children,color,soft=true,style={}})=>(
  <span style={{
    display:'inline-flex',alignItems:'center',gap:'4px',
    padding:'0 8px',height:'20px',borderRadius:S.r.sm,
    fontSize:'11px',fontWeight:500,letterSpacing:'0.01em',lineHeight:1,
    background:soft?(color+'1F'):color, color:soft?color:'#fff',
    ...style
  }}>{children}</span>
);
const statusColors={OPEN:C.info,'IN PROGRESS':C.amber,CLOSED:C.green};
const SBadge=({status,onClick})=>{
  const c=statusColors[status]||C.inkMute;
  return <button onClick={onClick} style={{
    display:'inline-flex',alignItems:'center',gap:'5px',height:'22px',padding:'0 8px',
    borderRadius:S.r.sm,fontSize:'11px',fontWeight:500,border:'none',
    cursor:onClick?'pointer':'default',background:c+'1F',color:c,lineHeight:1
  }}>
    <span style={{width:'6px',height:'6px',borderRadius:'50%',background:c}}/>{status}
  </button>;
};

function Toast({msg,type,onDone}){
  useEffect(()=>{const t=setTimeout(onDone,3000);return()=>clearTimeout(t);},[]);
  const c=type==='error'?C.rose:type==='warn'?C.amber:C.green;
  const icon=type==='error'?'✕':type==='warn'?'!':'✓';
  return <div style={{position:'fixed',bottom:'24px',right:'24px',minHeight:'40px',
    display:'flex',alignItems:'center',gap:'10px',
    background:C.panel,color:C.ink,padding:'8px 14px 8px 10px',
    borderRadius:S.r.lg,border:`1px solid ${C.border}`,borderLeft:`3px solid ${c}`,
    fontSize:'13px',fontWeight:500,zIndex:9999,boxShadow:S.elev3}}>
    <span style={{width:'20px',height:'20px',borderRadius:'50%',background:c+'1F',
      color:c,display:'flex',alignItems:'center',justifyContent:'center',
      fontSize:'12px',fontWeight:700,flexShrink:0}}>{icon}</span>
    {msg}
  </div>;
}
function Modal({msg,onOk,onCancel}){
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(15,23,42,0.45)',
      backdropFilter:'blur(2px)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:9998}}>
      <Card style={{padding:'22px 24px 18px',maxWidth:'420px',width:'90%',boxShadow:S.elev3,borderRadius:S.r.xl}}>
        <div style={{...T.section,marginBottom:'6px',color:C.ink}}>Confirm deletion</div>
        <div style={{fontSize:'13px',color:C.inkMute,marginBottom:'18px',lineHeight:1.5}}>{msg}</div>
        <div style={{display:'flex',gap:'8px',justifyContent:'flex-end'}}>
          <Btn variant="ghost" onClick={onCancel}>Cancel</Btn>
          <Btn onClick={onOk} style={{background:C.rose,borderColor:C.rose,color:'#fff'}}>Delete</Btn>
        </div>
      </Card>
    </div>
  );
}

// ── KPI atom (Phase B5) — compact enterprise tile: accent bar · caption · big tabular metric · sub
const KPI=({label,value,sub,accent=C.navy,subTone='mute'})=>(
  <Card style={{padding:'14px 16px',display:'flex',flexDirection:'column',gap:'6px',minHeight:'92px'}}>
    <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
      <span style={{width:'3px',height:'18px',borderRadius:'2px',background:accent,flexShrink:0}}/>
      <div style={{fontSize:'10.5px',fontWeight:600,letterSpacing:'0.08em',textTransform:'uppercase',color:C.inkMute}}>{label}</div>
    </div>
    <div style={{fontSize:'26px',fontWeight:700,letterSpacing:'-0.02em',lineHeight:1.05,color:C.ink,fontVariantNumeric:'tabular-nums',fontFeatureSettings:'"tnum"'}}>{value}</div>
    {sub&&<div style={{fontSize:'11px',color:subTone==='pos'?C.green:subTone==='neg'?C.rose:C.inkMute,lineHeight:1.3}}>{sub}</div>}
  </Card>
);

// ── Table styling injector (Phase B3) — Fluent/TanStack-inspired polish: header, hover, tabular-nums
// Scoped to .jss-app so SmartImporter modal tables also benefit. No structural JSX changes.
function TableStyles(){
  useEffect(()=>{
    if(document.getElementById('jss-tbl-css'))return;
    const s=document.createElement('style');s.id='jss-tbl-css';
    s.textContent=`
      .jss-app table{border-collapse:separate;border-spacing:0;}
      .jss-app thead th{background:#FBFAF7 !important;color:#6B6558 !important;border-bottom:1px solid #D9D3C5 !important;position:relative;}
      .jss-app tbody tr{transition:background-color .12s ease;}
      .jss-app tbody tr:hover{background-color:rgba(27,42,94,0.045) !important;}
      .jss-app tbody td{font-variant-numeric:tabular-nums;font-feature-settings:"tnum";}
      .jss-app tbody tr:last-child td{border-bottom:none !important;}
    `;
    document.head.appendChild(s);
  },[]);
  return null;
}

// ── JSS LOGO — resilient, retina-optimized, with text fallback ──
function JSSLogo(){
  const SOURCES=[
    (typeof window!=='undefined'&&window.__JSS_LOGO_URL__)||'',
    'jss-logo-white on blue@3x-50.png',
    'jss-logo-white on blue@3x-50.jpg',
    './jss-logo-white on blue@3x-50.png',
    'assets/jss-logo-white on blue@3x-50.png',
  ].filter(Boolean);
  const [idx,setIdx]=useState(0);
  const [failed,setFailed]=useState(false);
  const onErr=()=>{ if(idx<SOURCES.length-1) setIdx(idx+1); else setFailed(true); };
  const tile={padding:'11px',background:'#fff',borderRadius:'10px',marginBottom:'10px',
    display:'flex',alignItems:'center',justifyContent:'center',height:'58px',boxSizing:'border-box'};
  if(failed||SOURCES.length===0){
    return(
      <div style={tile} aria-label="Joyson Safety Systems">
        <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
          <div style={{width:'30px',height:'30px',borderRadius:'50%',background:C.navy,
            display:'flex',alignItems:'center',justifyContent:'center',
            color:'#fff',fontSize:'13px',fontWeight:800,letterSpacing:'-0.03em'}}>J</div>
          <div style={{lineHeight:1}}>
            <div style={{fontSize:'15px',fontWeight:800,color:C.navy,letterSpacing:'-0.01em'}}>JSS</div>
            <div style={{fontSize:'7.5px',fontWeight:600,color:C.inkMute,letterSpacing:'0.14em',marginTop:'2px'}}>SAFETY SYSTEMS</div>
          </div>
        </div>
      </div>
    );
  }
  return(
    <div style={tile}>
      <img src={SOURCES[idx]} alt="Joyson Safety Systems"
        onError={onErr}
        style={{maxHeight:'36px',maxWidth:'100%',objectFit:'contain'}}/>
    </div>
  );
}

const NAV=[
  {id:'dashboard',label:'Dashboard',ico:'M3 13h8V3H3zM13 21h8V11h-8zM3 21h8v-6H3zM13 9h8V3h-8z'},
  {id:'campaigns',label:'Campaigns',ico:'M19 4h-1V2h-2v2H8V2H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V9h14z'},
  {id:'repairs',label:'VIN Repairs',ico:'M9 5H7c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h10c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2h-2M9 5c0 1.1.9 2 2 2h2c1.1 0 2-.9 2-2M9 5c0-1.1.9-2 2-2h2c1.1 0 2 .9 2 2'},
  {id:'recall',label:'Recall Tracker',ico:'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z'},
  {id:'rampup',label:'Ramp-Up',ico:'M3 3v18h18M7 14l4-4 4 4 5-5'},
  {id:'spend',label:'Spend Analysis',ico:'M12 1v22M17 5H9.5a3.5 3.5 0 100 7h5a3.5 3.5 0 110 7H6'},
  {id:'parts',label:'Parts Reference',ico:'M20 7l-8-4-8 4v10l8 4 8-4zM12 3v18M4 7l8 4 8-4'},
  {id:'audit',label:'Audit Log',ico:'M9 5H7c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h10c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2h-2M9 5c0 1.1.9 2 2 2h2c1.1 0 2-.9 2-2'},
];
const NIcon=({d})=>(
  <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    {d.split('M').filter(Boolean).map((s,i)=><path key={i} d={'M'+s}/>)}
  </svg>
);
const PH=({title,sub})=>(
  <div style={{marginBottom:'2px',paddingBottom:'12px',borderBottom:`1px solid ${C.borderSoft}`,width:'100%'}}>
    <h1 style={{...T.pageTitle,margin:0,color:C.ink}}>{title}</h1>
    {sub&&<p style={{fontSize:'13px',color:C.inkMute,margin:'4px 0 0',lineHeight:1.45}}>{sub}</p>}
  </div>
);

// Field wrapper — persistent top-aligned label above data-entry controls.
// Presentation only; children control retains all logic, handlers, values.
const Field=({label,required,hint,children,style={}})=>(
  <div style={{display:'flex',flexDirection:'column',gap:'4px',minWidth:0,...style}}>
    <label style={{fontSize:'10.5px',fontWeight:600,letterSpacing:'0.06em',textTransform:'uppercase',color:C.inkSub,lineHeight:1.2}}>
      {label}{required&&<span style={{color:C.coral,marginLeft:'3px'}}>*</span>}
    </label>
    {children}
    {hint&&<div style={{fontSize:'10.5px',color:C.inkMute,lineHeight:1.3}}>{hint}</div>}
  </div>
);

// Section header for create/edit form cards — replaces micro-caps caption.
// Signals a composition surface (data entry) vs. a query surface (filters).
const FormHeader=({title,editing,hint})=>(
  <div style={{display:'flex',alignItems:'baseline',gap:'10px',marginBottom:'12px',paddingBottom:'10px',borderBottom:`1px solid ${C.borderSoft}`}}>
    <div style={{width:'3px',alignSelf:'stretch',background:editing?C.amber:C.coral,borderRadius:'2px',minHeight:'18px'}}/>
    <div style={{...T.cardHead,color:C.ink}}>{editing?'Edit '+title:'New '+title}</div>
    {hint&&<div style={{fontSize:'11.5px',color:C.inkMute}}>{hint}</div>}
    <div style={{flex:1}}/>
    <div style={{fontSize:'10px',fontWeight:600,letterSpacing:'0.1em',textTransform:'uppercase',color:C.inkMute}}>{editing?'Editing':'Data entry'}</div>
  </div>
);

// ══════════════════════════════════════════════════════
// PROTECTED LEGACY RECALL DATABASE — DO NOT MODIFY BEHAVIOR
// ══════════════════════════════════════════════════════
function RecallDatabaseApp({areaSwitch,tab,setTab}){
  const xlsxReady=useXLSX();
  const [oem,setOem]=useState('ALL');
  const [from,setFrom]=useState('');
  const [to,setTo]=useState('');
  const [toast,setToast]=useState(null);
  const [modal,setModal]=useState(null);

  const [camps,setCamps]=useLS('jss_camps_v2',SEED_CAMPS);
  const [vins,setVins]=useLS('jss_vins_v2',SEED_VINS);
  const [parts,setParts]=useLS('jss_parts_v2',SEED_PARTS);
  const [auditLog,setAuditLog]=useLS('jss_audit_v2',[]);
  const [ramps,setRamps]=useLS('jss_ramps_v2',{});

  const log=(action,detail)=>setAuditLog(prev=>[{id:Date.now(),ts:new Date().toISOString(),action,detail},...prev].slice(0,500));
  const toast$=(msg,type='success')=>setToast({msg,type});
  const confirm$=(msg,onOk)=>setModal({msg,onOk});

  const matchOEM=o=>oem==='ALL'||o===oem;
  const inRange=d=>{ if(!d)return true; if(from&&d<from)return false; if(to&&d>to)return false; return true; };

  const filtCamps=camps.filter(c=>matchOEM(c.oem)&&inRange(c.sop));
  const campIds=new Set(filtCamps.map(c=>c.id));
  // Dashboard/Spend use campaign-cascaded VIN set (aggregate by campaign — correct)
  const filtVins=vins.filter(v=>campIds.has(v.campaignId)&&inRange(v.repairDate||v.addedDate));
  // VIN Repairs / Recall Tracker use VIN-direct filtering (OEM on VIN, date on VIN activity — no campaign SOP cascade)
  const filtVinsDirect=vins.filter(v=>matchOEM(v.oem)&&inRange(v.repairDate||v.addedDate));
  // Parts Reference honors OEM filter directly (no date — parts have no activity date)
  const filtParts=parts.filter(p=>matchOEM(p.oem));
  const vinCounts=useMemo(()=>{ const m={}; vins.forEach(v=>{m[v.vin]=(m[v.vin]||0)+1;}); return m; },[vins]);

  const shared={camps,setCamps,vins,setVins,parts,setParts,filtCamps,filtVins,filtVinsDirect,filtParts,vinCounts,ramps,setRamps,log,toast$,confirm$,xlsxReady};

  const resetAll=()=>confirm$('Reset ALL data to defaults? This cannot be undone.',()=>{
    setCamps(SEED_CAMPS);setVins(SEED_VINS);setParts(SEED_PARTS);setAuditLog([]);setRamps({});
    toast$('Reset to defaults','warn');setModal(null);
  });

  return(
    <div className="jss-app" style={{display:'flex',height:'100vh',background:C.bg,fontFamily:'"Inter",-apple-system,sans-serif',color:C.ink,overflow:'hidden'}}>
      <TableStyles/>
      {/* Sidebar */}
      <aside style={{width:'215px',flexShrink:0,background:C.navBg,display:'flex',flexDirection:'column',padding:'14px 10px 10px',overflowY:'auto'}}>
        {areaSwitch}
        <JSSLogo/>
        <div style={{padding:'4px 8px 6px',marginBottom:'4px'}}>
          <span style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.14em',color:'rgba(255,255,255,0.42)',textTransform:'uppercase'}}>Recall Intelligence</span>
        </div>
        <nav style={{flex:1,display:'flex',flexDirection:'column',gap:'1px'}}>
          {NAV.map(n=>{
            const a=tab===n.id;
            return(
              <button key={n.id} onClick={()=>setTab(n.id)} style={{position:'relative',display:'flex',alignItems:'center',gap:'10px',height:'32px',padding:'0 10px 0 12px',borderRadius:'6px',border:'none',cursor:'pointer',textAlign:'left',fontSize:'12.5px',fontWeight:a?600:400,background:a?'rgba(255,255,255,0.08)':'transparent',color:a?'#fff':'rgba(255,255,255,0.62)',transition:'background 120ms, color 120ms',boxShadow:a?`inset 2px 0 0 ${C.coral}`:'none'}} onMouseEnter={e=>{if(!a)e.currentTarget.style.background='rgba(255,255,255,0.05)';}} onMouseLeave={e=>{if(!a)e.currentTarget.style.background='transparent';}}>
                <NIcon d={n.ico}/>{n.label}
              </button>
            );
          })}
        </nav>
        <div style={{marginTop:'8px',paddingTop:'8px',borderTop:'1px solid rgba(255,255,255,0.08)'}}>
          <button onClick={resetAll} style={{width:'100%',padding:'6px 10px',borderRadius:'6px',border:'none',background:'transparent',color:'rgba(255,255,255,0.42)',fontSize:'11px',cursor:'pointer',textAlign:'left',transition:'background 120ms'}} onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.05)'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>⟳ Reset Defaults</button>
        </div>
      </aside>

      {/* Main */}
      <div style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
        <header style={{padding:'12px 24px',display:'flex',alignItems:'center',gap:'10px',background:C.navBg,borderBottom:'1px solid rgba(0,0,0,0.35)',boxShadow:'inset 0 -1px 0 rgba(255,255,255,0.06)',flexShrink:0}}>
          <div style={{flex:1}}>
            <div style={{fontSize:'15px',fontWeight:600,color:'#fff',letterSpacing:'-0.005em'}}>JSS Recall Database</div>
            <div style={{fontSize:'11px',color:'rgba(255,255,255,0.55)',marginTop:'1px'}}>Joyson Safety Systems — Recall Operations</div>
          </div>
          <span style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.14em',color:'rgba(255,255,255,0.45)',textTransform:'uppercase',marginRight:'2px'}}>Filters</span>
          <Sel value={oem} onChange={e=>setOem(e.target.value)} style={{minWidth:'132px',height:'30px',padding:'0 10px',background:'rgba(255,255,255,0.10)',border:'1px solid rgba(255,255,255,0.15)',borderRadius:'6px',color:'#fff',fontSize:'12px'}}>
            <option value="ALL" style={{color:'#1B2A5E'}}>All OEMs</option>
            {OEMS.map(o=><option key={o} value={o} style={{color:'#1B2A5E'}}>{o}</option>)}
          </Sel>
          <Inp type="date" value={from} onChange={e=>setFrom(e.target.value)} style={{height:'30px',padding:'0 10px',background:'rgba(255,255,255,0.10)',border:'1px solid rgba(255,255,255,0.15)',borderRadius:'6px',color:'#fff',fontSize:'12px',colorScheme:'dark'}}/>
          <span style={{color:'rgba(255,255,255,0.45)',fontSize:'12px'}}>→</span>
          <Inp type="date" value={to} onChange={e=>setTo(e.target.value)} style={{height:'30px',padding:'0 10px',background:'rgba(255,255,255,0.10)',border:'1px solid rgba(255,255,255,0.15)',borderRadius:'6px',color:'#fff',fontSize:'12px',colorScheme:'dark'}}/>
          {(from||to)&&<button onClick={()=>{setFrom('');setTo('');}} style={{height:'30px',padding:'0 10px',fontSize:'11px',fontWeight:500,color:'rgba(255,255,255,0.75)',background:'rgba(255,255,255,0.08)',border:'1px solid rgba(255,255,255,0.15)',borderRadius:'6px',cursor:'pointer'}}>✕ Clear</button>}
        </header>
        <div style={{flex:1,overflow:'auto',padding:'18px 26px 32px'}}>
          {tab==='dashboard'&&<DashTab {...shared}/>}
          {tab==='campaigns'&&<CampsTab {...shared}/>}
          {tab==='repairs'&&<RepairsTab {...shared}/>}
          {tab==='recall'&&<RecallTab {...shared}/>}
          {tab==='rampup'&&<RampTab {...shared}/>}
          {tab==='spend'&&<SpendTab {...shared}/>}
          {tab==='parts'&&<PartsTab {...shared}/>}
          {tab==='audit'&&<AuditTab auditLog={auditLog} setAuditLog={setAuditLog} toast$={toast$}/>}
        </div>
      </div>

      {toast&&<Toast msg={toast.msg} type={toast.type} onDone={()=>setToast(null)}/>}
      {modal&&<Modal msg={modal.msg} onOk={modal.onOk} onCancel={()=>setModal(null)}/>}
    </div>
  );
}

// ════
// DASHBOARD
// ════
function DashTab({camps,vins,filtCamps,filtVins}){
  const repairs=filtVins.filter(v=>v.type==='REPAIRED');
  const uniqR=new Set(repairs.map(v=>v.vin)).size;
  const potential=filtCamps.reduce((s,c)=>s+Math.round((c.maxVINs||0)*(c.takeRate||0)/100),0);
  const ytd=filtCamps.reduce((s,c)=>{
    const u=new Set(repairs.filter(v=>v.campaignId===c.id&&getY(v.repairDate)===CY).map(v=>v.vin)).size;
    return s+u*(c.agreedRate||0);
  },0);
  const openCount=filtCamps.filter(c=>c.status!=='CLOSED').length;
  const allDupes=Object.values(useMemo(()=>{const m={};vins.forEach(v=>{m[v.vin]=(m[v.vin]||0)+1;});return m;},[vins])).filter(n=>n>1).length;
  const sc={OPEN:0,'IN PROGRESS':0,CLOSED:0};
  filtCamps.forEach(c=>{sc[c.status]=(sc[c.status]||0)+1;});
  const totalC=filtCamps.length||1;
  const tiles=[
    {l:'Active Campaigns',v:openCount,s:'currently open',accent:C.navy},
    {l:'Potential VINs',v:fmtN(potential),s:'max × take rate',accent:'#6B4EA8'},
    {l:'Unique Repaired',v:fmtN(uniqR),s:'confirmed VINs',accent:C.green,subTone:'pos'},
    {l:'YTD Spend',v:fmt$(ytd),s:CY+' year to date',accent:C.coral},
  ];
  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <PH title="Dashboard" sub="Live overview of your recall portfolio"/>
      {allDupes>0&&<div style={{padding:'9px 14px',background:C.tilePeach,borderRadius:'10px',border:`1px solid ${C.coral}30`,fontSize:'12.5px',color:C.rose,fontWeight:500}}>⚠ {allDupes} duplicate VIN(s) detected globally — check VIN Repairs tab</div>}
      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'12px'}}>
        {tiles.map(t=><KPI key={t.l} label={t.l} value={t.v} sub={t.s} accent={t.accent} subTone={t.subTone}/>)}
      </div>
      <Card style={{padding:'18px'}}>
        <div style={{fontSize:'13.5px',fontWeight:600,marginBottom:'12px'}}>Campaign Status Distribution</div>
        <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'12px'}}>
          {[{k:'OPEN',l:'Open',col:C.navy},{k:'IN PROGRESS',l:'In Progress',col:C.amber},{k:'CLOSED',l:'Closed',col:C.green}].map(s=>{
            const pct=Math.round((sc[s.k]||0)/totalC*100);
            return(
              <div key={s.k}>
                <div style={{display:'flex',alignItems:'center',gap:'5px',fontSize:'11px',color:C.inkMute,marginBottom:'4px'}}>
                  <span style={{width:'6px',height:'6px',borderRadius:'50%',background:s.col}}/>{s.l}
                </div>
                <div style={{fontSize:'22px',fontWeight:700}}>{pct}%</div>
                <div style={{height:'5px',borderRadius:'4px',background:'#eee',marginTop:'7px'}}>
                  <div style={{width:pct+'%',height:'100%',borderRadius:'4px',background:s.col}}/>
                </div>
                <div style={{fontSize:'11px',color:C.inkMute,marginTop:'3px'}}>{sc[s.k]||0} campaigns</div>
              </div>
            );
          })}
        </div>
      </Card>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px'}}>
        <Card style={{padding:'18px'}}>
          <div style={{fontSize:'13.5px',fontWeight:600,marginBottom:'12px'}}>VINs by OEM</div>
          <OEMPieChart vins={filtVins}/>
        </Card>
        <Card style={{padding:'18px'}}>
          <div style={{fontSize:'13.5px',fontWeight:600,marginBottom:'10px'}}>Quick Stats</div>
          {[
            {l:'Total VINs logged',v:fmtN(vins.length)},
            {l:'Unique VINs',v:fmtN(new Set(vins.map(v=>v.vin)).size)},
            {l:'Duplicate flags',v:fmtN(allDupes),warn:allDupes>0},
            {l:'Total campaigns',v:fmtN(camps.length)},
            {l:'Total potential',v:fmtN(camps.reduce((s,c)=>s+Math.round((c.maxVINs||0)*(c.takeRate||0)/100),0))},
          ].map(r=>(
            <div key={r.l} style={{display:'flex',justifyContent:'space-between',padding:'7px 0',borderBottom:`1px solid ${C.borderSoft}`,fontSize:'12px'}}>
              <span style={{color:C.inkMute}}>{r.l}</span>
              <span style={{fontWeight:600,color:r.warn?C.rose:C.ink}}>{r.v}</span>
            </div>
          ))}
        </Card>
      </div>
      <Card style={{padding:'4px 0'}}>
        <div style={{padding:'14px 18px 10px',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
          <div style={{fontSize:'13.5px',fontWeight:600}}>Campaign Breakdown</div>
          <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(filtCamps.map(c=>({Campaign:c.name,OEM:c.oem,MaxVINs:c.maxVINs,TakeRate:c.takeRate,AgreedRate:c.agreedRate,Status:c.status})),'dashboard')}>↓ Export CSV</Btn>
        </div>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
          <thead>
            <tr>{['Campaign','OEM','Max VINs','Take %','Potential','Repaired','Rate','Total Paid','Status'].map(h=>(
              <th key={h} style={{textAlign:'left',padding:'8px 18px',fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`}}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {filtCamps.map(c=>{
              const pot=Math.round((c.maxVINs||0)*(c.takeRate||0)/100);
              const u=new Set(repairs.filter(v=>v.campaignId===c.id).map(v=>v.vin)).size;
              return(
                <tr key={c.id} style={{borderBottom:`1px solid ${C.borderSoft}`}}>
                  <td style={{padding:'10px 18px',fontWeight:500}}>{c.name}</td>
                  <td style={{padding:'10px 18px'}}><Pill color={oemColor(c.oem)}>{c.oem}</Pill></td>
                  <td style={{padding:'10px 18px',textAlign:'right'}}>{fmtN(c.maxVINs)}</td>
                  <td style={{padding:'10px 18px',color:C.amber,fontWeight:600}}>{c.takeRate}%</td>
                  <td style={{padding:'10px 18px',textAlign:'right',fontWeight:600}}>{fmtN(pot)}</td>
                  <td style={{padding:'10px 18px',textAlign:'right',color:C.green,fontWeight:600}}>{fmtN(u)}</td>
                  <td style={{padding:'10px 18px',color:C.inkMute}}>{fmt$d(c.agreedRate)}</td>
                  <td style={{padding:'10px 18px',fontWeight:600}}>{fmt$(u*(c.agreedRate||0))}</td>
                  <td style={{padding:'10px 18px'}}><SBadge status={c.status}/></td>
                </tr>
              );
            })}
            {filtCamps.length===0&&<tr><td colSpan="9" style={{textAlign:'center',padding:'28px',color:C.inkMute}}>No campaigns match current filters</td></tr>}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ════
// CAMPAIGNS
// ════
function CampsTab({camps,setCamps,parts,filtCamps,log,toast$,confirm$}){
  const blank={name:'',oem:'STELLANTIS',vehicle:'',jssPN:'',agreedRate:'',maxVINs:'',takeRate:'',sop:'',status:'OPEN'};
  const [form,setForm]=useState(blank);
  const [editId,setEditId]=useState(null);

  const onPN=pn=>{
    const m=parts.find(p=>p.jssPN===pn);
    setForm(f=>({...f,jssPN:pn,agreedRate:m?m.pcPrice:f.agreedRate}));
  };
  const save=()=>{
    if(!form.name){toast$('Campaign name required','error');return;}
    const e={...form,agreedRate:+form.agreedRate||0,maxVINs:+form.maxVINs||0,takeRate:+form.takeRate||DEF_TR[form.oem]||65};
    if(editId){
      setCamps(camps.map(c=>c.id===editId?{...e,id:editId,statusLog:c.statusLog||[]}:c));
      log('CAMPAIGN_EDIT','Edited: '+form.name); toast$('Campaign updated'); setEditId(null);
    } else {
      setCamps([...camps,{...e,id:Date.now(),statusLog:[{status:'OPEN',ts:new Date().toISOString()}]}]);
      log('CAMPAIGN_ADD','Added: '+form.name); toast$('Campaign added');
    }
    setForm(blank);
  };
  const del=id=>{
    const c=camps.find(x=>x.id===id);
    confirm$('Delete "'+c?.name+'"?',()=>{
      setCamps(camps.filter(x=>x.id!==id));
      log('CAMPAIGN_DELETE','Deleted: '+c?.name); toast$('Deleted','warn');
    });
  };
  const cycle=c=>{
    const nx={OPEN:'IN PROGRESS','IN PROGRESS':'CLOSED',CLOSED:'OPEN'};
    const ns=nx[c.status];
    setCamps(camps.map(x=>x.id===c.id?{...x,status:ns,statusLog:[...(x.statusLog||[]),{status:ns,ts:new Date().toISOString()}]}:x));
    log('STATUS_CHANGE',c.name+': '+c.status+' → '+ns); toast$('Status → '+ns);
  };

  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="Campaigns" sub="Agreed rate auto-fills from Parts Reference when JSS PN is selected"/>
        <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(filtCamps.map(c=>({Name:c.name,OEM:c.oem,Vehicle:c.vehicle,JSSPN:c.jssPN,MaxVINs:c.maxVINs,TakeRate:c.takeRate,AgreedRate:c.agreedRate,Status:c.status,SOP:c.sop})),'campaigns')}>↓ Export CSV</Btn>
      </div>
      <Card style={{padding:'16px 18px'}}>
        <FormHeader title="Campaign" editing={!!editId} hint="Agreed rate auto-fills when a JSS PN is selected"/>
        <div style={{display:'grid',gridTemplateColumns:'2fr 1.2fr 1fr 1fr 1fr 1fr 1fr 1fr auto',gap:'10px',alignItems:'end'}}>
          <Field label="Campaign Name" required><Inp placeholder="e.g. STLA LX LD Retractor Recall" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/></Field>
          <Field label="OEM" required><Sel value={form.oem} onChange={e=>setForm({...form,oem:e.target.value,takeRate:DEF_TR[e.target.value]||65})}>{OEMS.map(o=><option key={o}>{o}</option>)}</Sel></Field>
          <Field label="Vehicle"><Inp placeholder="e.g. LX LD" value={form.vehicle} onChange={e=>setForm({...form,vehicle:e.target.value})}/></Field>
          <Field label="JSS PN"><Sel value={form.jssPN} onChange={e=>onPN(e.target.value)}>
            <option value="">Select…</option>
            {parts.map(p=><option key={p.id} value={p.jssPN}>{p.jssPN}</option>)}
          </Sel></Field>
          <Field label="Max VINs"><Inp type="number" placeholder="0" value={form.maxVINs} onChange={e=>setForm({...form,maxVINs:e.target.value})}/></Field>
          <Field label="Take %"><Inp type="number" placeholder="0" value={form.takeRate} onChange={e=>setForm({...form,takeRate:e.target.value})}/></Field>
          <Field label="Rate $"><Inp type="number" placeholder="0.00" value={form.agreedRate} onChange={e=>setForm({...form,agreedRate:e.target.value})}/></Field>
          <Field label="SOP Date"><Inp type="date" value={form.sop} onChange={e=>setForm({...form,sop:e.target.value})}/></Field>
          <div style={{display:'flex',gap:'5px'}}>
            <Btn variant="accent" onClick={save}>{editId?'Update':'Add'}</Btn>
            {editId&&<Btn variant="ghost" size="sm" onClick={()=>{setForm(blank);setEditId(null);}}>✕</Btn>}
          </div>
        </div>
        {form.jssPN&&parts.find(p=>p.jssPN===form.jssPN)&&(
          <div style={{marginTop:'7px',fontSize:'11px',color:C.green}}>✓ Rate auto-filled from Parts Reference: {fmt$d(parts.find(p=>p.jssPN===form.jssPN)?.pcPrice)}</div>
        )}
      </Card>
      <Card>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
          <thead>
            <tr>{['Campaign','OEM','Vehicle','JSS PN','Max VINs','Take %','Potential','Rate','SOP','Status',''].map(h=>(
              <th key={h} style={{textAlign:'left',padding:'9px 16px',fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`,whiteSpace:'nowrap'}}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {filtCamps.map(c=>{
              const pot=Math.round((c.maxVINs||0)*(c.takeRate||0)/100);
              return(
                <tr key={c.id} style={{borderBottom:`1px solid ${C.borderSoft}`}}>
                  <td style={{padding:'10px 16px',fontWeight:500}}>{c.name}</td>
                  <td style={{padding:'10px 16px'}}><Pill color={oemColor(c.oem)}>{c.oem}</Pill></td>
                  <td style={{padding:'10px 16px',color:C.inkMute}}>{c.vehicle}</td>
                  <td style={{padding:'10px 16px',fontFamily:'monospace',fontSize:'11px',color:C.inkMute}}>{c.jssPN}</td>
                  <td style={{padding:'10px 16px',textAlign:'right'}}>{fmtN(c.maxVINs)}</td>
                  <td style={{padding:'10px 16px',color:C.amber,fontWeight:600}}>{c.takeRate}%</td>
                  <td style={{padding:'10px 16px',textAlign:'right',fontWeight:600}}>{fmtN(pot)}</td>
                  <td style={{padding:'10px 16px',color:C.inkMute}}>{fmt$d(c.agreedRate)}</td>
                  <td style={{padding:'10px 16px',color:C.inkMute,fontSize:'11px'}}>{c.sop}</td>
                  <td style={{padding:'10px 16px'}}><SBadge status={c.status} onClick={()=>cycle(c)}/></td>
                  <td style={{padding:'10px 16px',display:'flex',gap:'6px'}}>
                    <Btn variant="ghost" size="xs" onClick={()=>{setForm({...c});setEditId(c.id);}}>Edit</Btn>
                    <Btn variant="danger" size="xs" onClick={()=>del(c.id)}>Del</Btn>
                  </td>
                </tr>
              );
            })}
            {filtCamps.length===0&&<tr><td colSpan="11" style={{textAlign:'center',padding:'28px',color:C.inkMute}}>No campaigns — add one above</td></tr>}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ════
// VIN REPAIRS
// ════
function RepairsTab({camps,vins,setVins,parts,setParts,filtVinsDirect,vinCounts,log,toast$,confirm$,xlsxReady}){
  const filtVins=filtVinsDirect;
  const [form,setForm]=useState({vin:'',campaignId:'',oem:'',repairDate:''});
  const [vinfo,setVinfo]=useState(null);
  const [search,setSearch]=useState('');
  const [showImporter,setShowImporter]=useState(false);

  const repairs=filtVins.filter(v=>v.type==='REPAIRED');
  const visible=repairs.filter(v=>!search||v.vin.toLowerCase().includes(search.toLowerCase())||(v.oem||'').toLowerCase().includes(search.toLowerCase()));

  const onVin=raw=>{
    const v=raw.toUpperCase().trim();
    setForm(f=>({...f,vin:v}));
    if(v.length===17){
      const res=validateVIN(v);
      const dec=decodeVIN(v);
      const camp=camps.find(c=>c.id===+form.campaignId);
      setVinfo({...res,...dec,isDupe:(vinCounts[v]||0)>0,oemMismatch:dec&&dec.oem!=='UNKNOWN'&&camp&&dec.oem!==camp.oem,campOEM:camp?.oem});
      if(dec&&dec.oem!=='UNKNOWN') setForm(f=>({...f,vin:v,oem:dec.oem}));
    } else { setVinfo(null); }
  };
  const onCamp=id=>{
    const c=camps.find(x=>x.id===+id);
    setForm(f=>({...f,campaignId:id,oem:c?c.oem:f.oem}));
  };
  const vinBorder=()=>{
    if(!vinfo||!form.vin)return C.border;
    if(!vinfo.valid)return C.rose;
    if(vinfo.warning||vinfo.oemMismatch||vinfo.isDupe)return C.amber;
    return C.green;
  };
  const add=()=>{
    if(!form.vin){toast$('VIN required','error');return;}
    if(!form.campaignId){toast$('Select a campaign','error');return;}
    const res=validateVIN(form.vin);
    if(!res.valid){toast$('Invalid VIN: '+res.error,'error');return;}
    const isDupe=(vinCounts[form.vin]||0)>0;
    setVins(p=>[...p,{id:Date.now(),vin:form.vin,campaignId:+form.campaignId,oem:form.oem,repairDate:form.repairDate,type:'REPAIRED',addedDate:new Date().toISOString().slice(0,10)}]);
    log('VIN_ADD','VIN '+form.vin+(isDupe?' [DUPLICATE]':''));
    toast$(isDupe?'⚠ VIN added — duplicate flagged':'VIN added',isDupe?'warn':'success');
    setForm(f=>({...f,vin:'',repairDate:''}));
    setVinfo(null);
  };
  const del=(id,vin)=>confirm$('Remove VIN '+vin+'?',()=>{
    setVins(p=>p.filter(v=>v.id!==id));
    log('VIN_DELETE','VIN '+vin+' removed'); toast$('Removed','warn');
  });
  const handleImportVINs=newVins=>{
    const dupes=newVins.filter(v=>(vinCounts[v.vin]||0)>0).length;
    setVins(p=>[...p,...newVins]);
    log('SMART_IMPORT','Imported '+newVins.length+' VINs, '+dupes+' dupes');
    toast$('Imported '+newVins.length+' VINs'+(dupes>0?', '+dupes+' duplicates flagged':''),dupes>0?'warn':'success');
  };
  const handleImportParts=newParts=>{
    setParts(p=>[...p,...newParts]);
    log('SMART_IMPORT','Imported '+newParts.length+' parts to Parts Reference');
    toast$('Imported '+newParts.length+' parts to Parts Reference');
  };

  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="VIN Repairs" sub="Confirmed repairs — VINs validated on entry, duplicates flagged globally"/>
        <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(visible.map(v=>({VIN:v.vin,Campaign:camps.find(c=>c.id===v.campaignId)?.name||'',OEM:v.oem,RepairDate:v.repairDate,Duplicate:vinCounts[v.vin]>1?'YES':'NO',Year:decodeVIN(v.vin)?.year||'',Make:decodeVIN(v.vin)?.make||''})),'vin_repairs')}>↓ Export CSV</Btn>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px'}}>
        <Card style={{padding:'16px 18px'}}>
          <FormHeader title="VIN Repair" hint="Enter a single confirmed repair"/>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'10px',marginBottom:'8px'}}>
            <Field label="VIN (17 characters)" required style={{gridColumn:'1/-1'}}>
              <Inp placeholder="e.g. 1FTFW1ET5NFA12345" value={form.vin} onChange={e=>onVin(e.target.value)} style={{border:`2px solid ${vinBorder()}`}}/>
            </Field>
            <Field label="Campaign" required hint="Assigns this VIN to a recall campaign">
              <Sel value={form.campaignId} onChange={e=>onCamp(e.target.value)}>
                <option value="">Select campaign…</option>
                {camps.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
              </Sel>
            </Field>
            <Field label="Repair Date">
              <Inp type="date" value={form.repairDate} onChange={e=>setForm({...form,repairDate:e.target.value})}/>
            </Field>
          </div>
          {vinfo&&(
            <div style={{padding:'7px 10px',borderRadius:'9px',background:vinfo.valid?(vinfo.warning||vinfo.isDupe||vinfo.oemMismatch?C.tilePeach:C.tileGreen):'#fde8e8',fontSize:'11px',marginBottom:'7px',lineHeight:1.5}}>
              {!vinfo.valid&&<span style={{color:C.rose}}>✗ {vinfo.error}</span>}
              {vinfo.valid&&<>
                <span style={{color:C.green}}>✓ Valid</span>
                {vinfo.year&&<span style={{color:C.inkMute}}> · {vinfo.year} {vinfo.make}</span>}
                {vinfo.warning&&<span style={{color:C.amber}}> · ⚠ {vinfo.warning}</span>}
                {vinfo.oemMismatch&&<span style={{color:C.rose}}> · ⚠ OEM mismatch: VIN={vinfo.oem}, Campaign={vinfo.campOEM}</span>}
                {vinfo.isDupe&&<span style={{color:C.rose}}> · ⚠ DUPLICATE — VIN already exists</span>}
              </>}
            </div>
          )}
          <Btn variant="accent" onClick={add} style={{width:'100%'}}>+ Add VIN</Btn>
        </Card>
        <Card style={{padding:'24px 22px',display:'flex',flexDirection:'column',gap:'12px',alignItems:'center',justifyContent:'center',background:`linear-gradient(135deg,${C.tileBlue},${C.bg})`,border:`1.5px dashed ${C.border}`,minHeight:'160px'}}>
          <div style={{fontSize:'32px'}}>🔍</div>
          <div style={{fontSize:'13.5px',fontWeight:600,color:C.ink}}>Smart File Import</div>
          <div style={{fontSize:'11.5px',color:C.inkMute,textAlign:'center',maxWidth:'260px',lineHeight:1.5}}>Drop any Excel or CSV — columns auto-detected, VINs validated, duplicates flagged before committing</div>
          <Btn variant="accent" onClick={()=>setShowImporter(true)}>📂 Open Smart Importer</Btn>
          <div style={{fontSize:'10.5px',color:C.inkMute}}>Supports .xlsx · .xls · .csv · any column order</div>
        </Card>
      </div>
      <Card>
        <div style={{padding:'11px 16px',borderBottom:`1px solid ${C.border}`,display:'flex',alignItems:'center',gap:'9px'}}>
          <Inp placeholder="🔍 Search VINs or OEM…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1}}/>
          <div style={{fontSize:'11.5px',color:C.inkMute,whiteSpace:'nowrap'}}>{visible.length} records · <b style={{color:C.rose}}>{visible.filter(v=>vinCounts[v.vin]>1).length} dupes</b></div>
        </div>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
          <thead><tr>{['VIN','Campaign','OEM','Yr / Make','Repair Date','Flag',''].map(h=>(
            <th key={h} style={{textAlign:'left',padding:'8px 16px',fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`}}>{h}</th>
          ))}</tr></thead>
          <tbody>
            {visible.map(v=>{
              const d=vinCounts[v.vin]>1;
              const dec=decodeVIN(v.vin);
              return(
                <tr key={v.id} style={{borderBottom:`1px solid ${C.borderSoft}`,background:d?'rgba(200,16,46,0.04)':''}}>
                  <td style={{padding:'9px 16px',fontFamily:'monospace',fontSize:'11.5px',fontWeight:500}}>{v.vin}</td>
                  <td style={{padding:'9px 16px',color:C.inkMute,fontSize:'11.5px'}}>{camps.find(c=>c.id===v.campaignId)?.name||'—'}</td>
                  <td style={{padding:'9px 16px'}}><Pill color={oemColor(v.oem)}>{v.oem||'?'}</Pill></td>
                  <td style={{padding:'9px 16px',color:C.inkMute,fontSize:'11px'}}>{dec?(dec.year||'?')+' '+dec.make:'—'}</td>
                  <td style={{padding:'9px 16px',color:C.inkMute,fontSize:'11.5px'}}>{v.repairDate||'—'}</td>
                  <td style={{padding:'9px 16px'}}>{d?<Pill color={C.rose} soft={false}>⚠ DUPE</Pill>:<Pill color={C.green}>✓ OK</Pill>}</td>
                  <td style={{padding:'9px 16px',textAlign:'right'}}><button onClick={()=>del(v.id,v.vin)} style={{background:'none',border:'none',color:C.rose,fontSize:'12px',cursor:'pointer'}}>Del</button></td>
                </tr>
              );
            })}
            {visible.length===0&&<tr><td colSpan="7" style={{textAlign:'center',padding:'28px',color:C.inkMute}}>No repair VINs yet — add or import above</td></tr>}
          </tbody>
        </table>
      </Card>
      {showImporter&&<SmartImporter camps={camps} parts={parts} vinCounts={vinCounts} onImportVINs={handleImportVINs} onImportParts={handleImportParts} onClose={()=>setShowImporter(false)} xlsxReady={xlsxReady} toast$={toast$}/>}
    </div>
  );
}

// ════
// RECALL TRACKER
// ════
function RecallTab({camps,vins,setVins,filtVinsDirect,log,toast$,confirm$}){
  const filtVins=filtVinsDirect;
  const [form,setForm]=useState({vin:'',campaignId:'',oem:''});
  const [search,setSearch]=useState('');
  const rcl=filtVins.filter(v=>v.type==='RECALL_ONLY');
  const vis=rcl.filter(v=>!search||v.vin.toLowerCase().includes(search.toLowerCase()));
  const allR=new Set(vins.filter(v=>v.type==='REPAIRED').map(v=>v.vin));
  const add=()=>{
    if(!form.vin||!form.campaignId){toast$('VIN and Campaign required','error');return;}
    const res=validateVIN(form.vin.toUpperCase());
    if(!res.valid){toast$('Invalid VIN: '+res.error,'error');return;}
    setVins(p=>[...p,{id:Date.now(),vin:form.vin.toUpperCase(),campaignId:+form.campaignId,oem:form.oem,repairDate:'',type:'RECALL_ONLY',addedDate:new Date().toISOString().slice(0,10)}]);
    log('RECALL_ADD','Recall VIN '+form.vin);
    toast$('Recall VIN added');
    setForm(f=>({...f,vin:''}));
  };
  const mark=v=>{
    setVins(p=>p.map(x=>x.id===v.id?{...x,type:'REPAIRED',repairDate:new Date().toISOString().slice(0,10)}:x));
    log('VIN_REPAIRED','VIN '+v.vin+' marked repaired'); toast$('Marked as repaired ✓');
  };
  const del=(id,vin)=>confirm$('Remove recall VIN '+vin+'?',()=>{
    setVins(p=>p.filter(v=>v.id!==id));
    log('RECALL_DELETE','Recall VIN '+vin); toast$('Removed','warn');
  });
  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="Recall Tracker" sub="Recall population — VINs not yet confirmed repaired"/>
        <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(vis.map(v=>({VIN:v.vin,Campaign:camps.find(c=>c.id===v.campaignId)?.name||'',OEM:v.oem,Repaired:allR.has(v.vin)?'YES':'NO'})),'recall_tracker')}>↓ Export CSV</Btn>
      </div>
      <Card style={{padding:'16px 18px'}}>
        <FormHeader title="Recall Population VIN" hint="Adds a VIN pending repair — this is a data-entry form, not a filter"/>
        <div style={{display:'grid',gridTemplateColumns:'2fr 2fr auto',gap:'10px',alignItems:'end'}}>
          <Field label="VIN" required>
            <Inp placeholder="17 characters" value={form.vin} onChange={e=>setForm({...form,vin:e.target.value.toUpperCase()})}/>
          </Field>
          <Field label="Campaign" required hint="Assigns this VIN to a recall campaign">
            <Sel value={form.campaignId} onChange={e=>{const c=camps.find(x=>x.id===+e.target.value);setForm(f=>({...f,campaignId:e.target.value,oem:c?c.oem:''}));}}>
              <option value="">Select campaign…</option>
              {camps.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
            </Sel>
          </Field>
          <Btn variant="accent" onClick={add}>+ Add VIN</Btn>
        </div>
      </Card>
      <Card>
        <div style={{padding:'11px 16px',borderBottom:`1px solid ${C.border}`,display:'flex',gap:'9px',alignItems:'center'}}>
          <Inp placeholder="🔍 Search VINs…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1}}/>
          <div style={{fontSize:'11.5px',color:C.inkMute}}>{vis.length} in recall · <b style={{color:C.green}}>{vis.filter(v=>allR.has(v.vin)).length} repaired</b></div>
        </div>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
          <thead><tr>{['VIN','Campaign','OEM','Status',''].map(h=>(
            <th key={h} style={{textAlign:'left',padding:'8px 16px',fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`}}>{h}</th>
          ))}</tr></thead>
          <tbody>
            {vis.map(v=>{const r=allR.has(v.vin);return(
              <tr key={v.id} style={{borderBottom:`1px solid ${C.borderSoft}`}}>
                <td style={{padding:'9px 16px',fontFamily:'monospace',fontSize:'11.5px',fontWeight:500}}>{v.vin}</td>
                <td style={{padding:'9px 16px',color:C.inkMute,fontSize:'11.5px'}}>{camps.find(c=>c.id===v.campaignId)?.name||'—'}</td>
                <td style={{padding:'9px 16px'}}><Pill color={oemColor(v.oem)}>{v.oem||'?'}</Pill></td>
                <td style={{padding:'9px 16px'}}>{r?<Pill color={C.green}>✓ Repaired</Pill>:<Pill color={C.amber}>Pending</Pill>}</td>
                <td style={{padding:'9px 16px',textAlign:'right',display:'flex',gap:'7px',justifyContent:'flex-end'}}>
                  {!r&&<Btn size="xs" variant="success" onClick={()=>mark(v)}>✓ Mark Repaired</Btn>}
                  <button onClick={()=>del(v.id,v.vin)} style={{background:'none',border:'none',color:C.rose,fontSize:'12px',cursor:'pointer'}}>Del</button>
                </td>
              </tr>
            );})}
            {vis.length===0&&<tr><td colSpan="5" style={{textAlign:'center',padding:'28px',color:C.inkMute}}>No recall VINs yet</td></tr>}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ════
// RAMP-UP
// ════
const DEF_RAMP=[5,10,15,20,20,10,8,5,3,2,1,1];
function RampTab({filtCamps,ramps,setRamps,toast$}){
  const [selId,setSelId]=useState(filtCamps[0]?.id||'');
  const camp=filtCamps.find(c=>c.id===+selId);
  const key=selId?String(selId):null;
  const pcts=key&&ramps[key]?ramps[key]:[...DEF_RAMP];
  const upd=(i,val)=>{
    const n=[...pcts]; n[i]=Math.max(0,Math.min(100,+val||0));
    setRamps(r=>({...r,[key]:n}));
  };
  const even=()=>{ setRamps(r=>({...r,[key]:Array(12).fill(Math.round(100/12))})); toast$('Distributed evenly'); };
  const reset=()=>{ setRamps(r=>({...r,[key]:[...DEF_RAMP]})); toast$('Reset to default curve'); };
  const total=pcts.reduce((s,v)=>s+v,0);
  const pot=camp?Math.round((camp.maxVINs||0)*(camp.takeRate||0)/100):0;
  const projVINs=camp?pcts.reduce((s,p)=>s+Math.round(pot*p/100),0):0;
  const projRev =camp?pcts.reduce((s,p)=>s+Math.round(pot*p/100*(camp.agreedRate||0)),0):0;
  const avgRev  =camp?Math.round(projRev/12):0;
  const peakIdx =camp?pcts.indexOf(Math.max(...pcts)):-1;
  const maxPct  =camp?Math.max(...pcts,1):1;
  // Compact SVG area chart (C2): 96px tall, gridlines, filled curve, peak dot
  const chartW=560, chartH=72, padX=2, padY=4;
  const stepX =(chartW-padX*2)/11;
  const yFor  =v=>chartH-padY-((v/maxPct)*(chartH-padY*2));
  const pts   =pcts.map((p,i)=>[padX+i*stepX, yFor(p)]);
  const linePath=pts.map((pt,i)=>(i?'L':'M')+pt[0].toFixed(1)+','+pt[1].toFixed(1)).join(' ');
  const areaPath=linePath+` L${(padX+11*stepX).toFixed(1)},${chartH-padY} L${padX},${chartH-padY} Z`;

  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      {/* C6 — Header + toolbar alignment */}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="Ramp-Up Schedule" sub="Editable weekly repair volume percentage per campaign"/>
        <div style={{display:'flex',gap:'7px',alignItems:'center'}}>
          {/* C5 — Validation chip beside actions */}
          {camp&&<Pill color={total===100?C.green:C.amber}>{total===100?'✓ 100%':'⚠ '+total+'%'}</Pill>}
          <Btn variant="ghost" size="sm" onClick={even}>Even</Btn>
          <Btn variant="ghost" size="sm" onClick={reset}>Reset</Btn>
          <Btn variant="ghost" size="sm" onClick={()=>{
            if(!camp)return;
            downloadCSV(pcts.map((p,i)=>({Week:'Wk '+(i+1),Pct:p,ProjVINs:Math.round(pot*p/100),ProjRevenue:Math.round(pot*p/100*(camp.agreedRate||0))})),'rampup_'+camp.name.replace(/\s/g,'_'));
          }}>↓ Export</Btn>
        </div>
      </div>
      <Card style={{padding:'14px 18px 16px'}}>
        {/* Selector row */}
        <div style={{display:'flex',alignItems:'center',gap:'10px',marginBottom:'12px'}}>
          <Sel value={selId} onChange={e=>setSelId(e.target.value)} style={{minWidth:'260px',flex:'0 0 auto'}}>
            <option value="">Select Campaign</option>
            {filtCamps.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
          </Sel>
          {/* C1 — Context strip */}
          {camp&&(
            <div style={{display:'flex',gap:'22px',alignItems:'center',paddingLeft:'6px',borderLeft:`1px solid ${C.borderSoft}`,marginLeft:'2px',paddingTop:'2px',paddingBottom:'2px'}}>
              <div>
                <div style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',textTransform:'uppercase',color:C.inkMute}}>OEM</div>
                <div style={{marginTop:'2px'}}><Pill color={oemColor(camp.oem)}>{camp.oem}</Pill></div>
              </div>
              <div>
                <div style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',textTransform:'uppercase',color:C.inkMute}}>Potential VINs</div>
                <div style={{fontSize:'15px',fontWeight:700,color:C.ink,...T.num,lineHeight:1.2,marginTop:'1px'}}>{fmtN(pot)}</div>
              </div>
              <div>
                <div style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',textTransform:'uppercase',color:C.inkMute}}>Agreed Rate</div>
                <div style={{fontSize:'15px',fontWeight:700,color:C.ink,...T.num,lineHeight:1.2,marginTop:'1px'}}>{fmt$d(camp.agreedRate)}</div>
              </div>
              <div>
                <div style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',textTransform:'uppercase',color:C.inkMute}}>Peak Week</div>
                <div style={{fontSize:'15px',fontWeight:700,color:C.ink,...T.num,lineHeight:1.2,marginTop:'1px'}}>{peakIdx>=0?'Wk '+(peakIdx+1)+' · '+pcts[peakIdx]+'%':'—'}</div>
              </div>
            </div>
          )}
        </div>
        {camp?(
          <>
            {/* C2 — Compact SVG area chart with gridlines and peak marker */}
            <div style={{marginBottom:'10px'}}>
              <svg viewBox={`0 0 ${chartW} ${chartH}`} preserveAspectRatio="none" style={{width:'100%',height:'72px',display:'block'}}>
                {[0.25,0.5,0.75].map(f=>(
                  <line key={f} x1={padX} x2={chartW-padX} y1={chartH-padY-f*(chartH-padY*2)} y2={chartH-padY-f*(chartH-padY*2)} stroke={C.borderSoft} strokeWidth="1" strokeDasharray="2 3"/>
                ))}
                <line x1={padX} x2={chartW-padX} y1={chartH-padY} y2={chartH-padY} stroke={C.border} strokeWidth="1"/>
                <path d={areaPath} fill={C.navy} fillOpacity="0.10"/>
                <path d={linePath} fill="none" stroke={C.navy} strokeWidth="1.6" strokeLinejoin="round" strokeLinecap="round"/>
                {pts.map((pt,i)=>(
                  <circle key={i} cx={pt[0]} cy={pt[1]} r={i===peakIdx?3.2:2} fill={i===peakIdx?C.coral:C.navy}/>
                ))}
              </svg>
            </div>
            {/* C3 — Weekly grid with peak highlight */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(12,1fr)',gap:'5px'}}>
              {pcts.map((p,i)=>{
                const isPeak=i===peakIdx;
                return(
                  <div key={i} style={{textAlign:'center',paddingTop:'4px',borderTop:isPeak?`2px solid ${C.coral}`:`2px solid transparent`}}>
                    <div style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.06em',color:isPeak?C.coral:C.inkMute,textTransform:'uppercase',marginBottom:'3px'}}>Wk {i+1}</div>
                    <Inp type="number" value={p} onChange={e=>upd(i,e.target.value)} style={{textAlign:'center',padding:'5px 2px',fontSize:'11.5px',width:'100%',boxSizing:'border-box',height:'28px'}}/>
                    <div style={{fontSize:'10px',color:C.inkMute,marginTop:'3px',...T.num}}>{fmtN(Math.round(pot*p/100))}</div>
                    <div style={{fontSize:'10px',color:C.green,fontWeight:500,...T.num}}>{fmt$(Math.round(pot*p/100*(camp.agreedRate||0)))}</div>
                  </div>
                );
              })}
            </div>
            {/* C4 — 3-metric summary footer */}
            <div style={{marginTop:'12px',padding:'10px 14px',background:C.tileBlue,borderRadius:S.r.md,display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'12px'}}>
              {[
                {l:'Projected VINs',v:fmtN(projVINs),c:C.ink},
                {l:'Projected Revenue',v:fmt$(projRev),c:C.green},
                {l:'Avg Weekly Revenue',v:fmt$(avgRev),c:C.ink},
              ].map(m=>(
                <div key={m.l}>
                  <div style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',textTransform:'uppercase',color:C.inkMute}}>{m.l}</div>
                  <div style={{fontSize:'17px',fontWeight:700,color:m.c,...T.num,lineHeight:1.2,marginTop:'2px'}}>{m.v}</div>
                </div>
              ))}
            </div>
          </>
        ):(
          /* C7 — Empty state */
          <div style={{padding:'32px 16px',textAlign:'center'}}>
            <div style={{fontSize:'13px',fontWeight:600,color:C.ink,marginBottom:'4px'}}>No campaign selected</div>
            <div style={{fontSize:'12px',color:C.inkMute}}>Choose a campaign above to view and edit its 12-week ramp-up schedule.</div>
          </div>
        )}
      </Card>
    </div>
  );
}

// ════
// SPEND
// ════
function SpendTab({filtCamps,filtVins}){
  const repairs=filtVins.filter(v=>v.type==='REPAIRED');
  const byQ={};
  repairs.forEach(v=>{
    if(!v.repairDate)return;
    const c=filtCamps.find(x=>x.id===v.campaignId); if(!c)return;
    const y=getY(v.repairDate),q=getQ(v.repairDate),k=y+' Q'+q;
    if(!byQ[k])byQ[k]={l:k,spend:0,vins:new Set(),y,q};
    if(!byQ[k].vins.has(v.vin)){byQ[k].vins.add(v.vin);byQ[k].spend+=c.agreedRate||0;}
  });
  const qs=Object.values(byQ).sort((a,b)=>a.y!==b.y?a.y-b.y:a.q-b.q);
  const maxS=Math.max(...qs.map(q=>q.spend),1);
  const byO={};
  filtCamps.forEach(c=>{
    const u=new Set(repairs.filter(v=>v.campaignId===c.id).map(v=>v.vin)).size;
    byO[c.oem]=(byO[c.oem]||0)+u*(c.agreedRate||0);
  });
  const ytd=qs.filter(q=>q.y===CY).reduce((s,q)=>s+q.spend,0);
  const total=qs.reduce((s,q)=>s+q.spend,0);
  const tiles=[
    {l:'Q'+CQ+' '+CY,v:fmt$(qs.find(q=>q.y===CY&&q.q===CQ)?.spend||0),s:'current quarter',accent:C.coral},
    {l:CY+' YTD',v:fmt$(ytd),s:CY+' year to date',accent:C.navy},
    {l:'All-Time',v:fmt$(total),s:'cumulative spend',accent:C.green,subTone:'pos'},
  ];
  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="Spend Analysis" sub="Quarterly OEM reimbursements from confirmed repairs"/>
        <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(qs.map(q=>({Quarter:q.l,Spend:q.spend,UniqueVINs:q.vins.size})),'spend_analysis')}>↓ Export CSV</Btn>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'12px'}}>
        {tiles.map(t=><KPI key={t.l} label={t.l} value={t.v} sub={t.s} accent={t.accent} subTone={t.subTone}/>)}
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1.6fr 1fr',gap:'12px'}}>
        <Card style={{padding:'18px'}}>
          <div style={{fontSize:'13.5px',fontWeight:600,marginBottom:'12px'}}>Spend by Quarter</div>
          {qs.length===0&&<div style={{textAlign:'center',color:C.inkMute,padding:'22px'}}>No spend data — add repaired VINs with dates</div>}
          {qs.map(q=>(
            <div key={q.l} style={{marginBottom:'11px'}}>
              <div style={{display:'flex',justifyContent:'space-between',fontSize:'12px',marginBottom:'4px'}}>
                <span style={{fontWeight:q.y===CY&&q.q===CQ?600:400}}>{q.l}{q.y===CY&&q.q===CQ?' ★':''}</span>
                <span style={{fontWeight:600}}>{fmt$(q.spend)}</span>
              </div>
              <div style={{height:'6px',borderRadius:'4px',background:'#eee'}}>
                <div style={{height:'100%',borderRadius:'4px',width:((q.spend/maxS)*100)+'%',background:q.y===CY&&q.q===CQ?C.navy:C.inkMute}}/>
              </div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginTop:'2px'}}>{q.vins.size} unique VINs</div>
            </div>
          ))}
        </Card>
        <Card style={{padding:'18px'}}>
          <div style={{fontSize:'13.5px',fontWeight:600,marginBottom:'10px'}}>Paid by OEM</div>
          {Object.entries(byO).sort((a,b)=>b[1]-a[1]).map(([o,s])=>(
            <div key={o} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'8px 0',borderBottom:`1px solid ${C.borderSoft}`}}>
              <Pill color={oemColor(o)}>{o}</Pill>
              <span style={{fontWeight:600,fontSize:'12.5px'}}>{fmt$(s)}</span>
            </div>
          ))}
          {Object.keys(byO).length===0&&<div style={{textAlign:'center',color:C.inkMute,padding:'14px'}}>No data yet</div>}
        </Card>
      </div>
    </div>
  );
}

// ════
// PARTS
// ════
function PartsTab({parts,setParts,filtParts,log,toast$,confirm$}){
  const blank={oem:'STELLANTIS',vehicle:'',jssPN:'',custPN:'',description:'',plant:'',volume:'',stdCost:'',pcPrice:'',sop:'',status:'ACTIVE'};
  const [form,setForm]=useState(blank);
  const [editId,setEditId]=useState(null);
  const [search,setSearch]=useState('');
  const [sf,setSf]=useState('ALL');
  const save=()=>{
    if(!form.jssPN){toast$('JSS PN required','error');return;}
    const e={...form,volume:+form.volume||0,stdCost:+form.stdCost||0,pcPrice:+form.pcPrice||0};
    if(editId){
      setParts(parts.map(p=>p.id===editId?{...e,id:editId}:p));
      log('PART_EDIT','Edited: '+form.jssPN); toast$('Part updated'); setEditId(null);
    } else {
      setParts([...parts,{...e,id:Date.now()}]);
      log('PART_ADD','Added: '+form.jssPN); toast$('Part added');
    }
    setForm(blank);
  };
  const del=(id,pn)=>confirm$('Delete part '+pn+'?',()=>{
    setParts(parts.filter(p=>p.id!==id));
    log('PART_DELETE','Deleted: '+pn); toast$('Deleted','warn');
  });
  const vis=filtParts.filter(p=>{
    const s=(p.oem+p.vehicle+p.jssPN+p.description+p.plant).toLowerCase();
    return(!search||s.includes(search.toLowerCase()))&&(sf==='ALL'||p.status===sf);
  });
  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="Parts Reference" sub="PC Price auto-fills Agreed Rate when JSS PN selected in Campaigns"/>
        <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(vis.map(p=>({OEM:p.oem,Vehicle:p.vehicle,JSSPN:p.jssPN,CustPN:p.custPN,Description:p.description,Plant:p.plant,Volume:p.volume,StdCost:p.stdCost,PCPrice:p.pcPrice,SOP:p.sop,Status:p.status})),'parts_reference')}>↓ Export CSV</Btn>
      </div>
      <Card style={{padding:'16px 18px'}}>
        <FormHeader title="Part" editing={!!editId} hint="PC Price becomes the Agreed Rate when JSS PN is chosen in a Campaign"/>
        <div style={{display:'grid',gridTemplateColumns:'1.2fr 1fr 1fr 1fr 1fr 1fr',gap:'10px',marginBottom:'10px'}}>
          <Field label="OEM" required><Sel value={form.oem} onChange={e=>setForm({...form,oem:e.target.value})}>{OEMS.map(o=><option key={o}>{o}</option>)}</Sel></Field>
          <Field label="Vehicle"><Inp placeholder="e.g. F-150" value={form.vehicle} onChange={e=>setForm({...form,vehicle:e.target.value})}/></Field>
          <Field label="JSS PN" required><Inp placeholder="e.g. JSS-B1" value={form.jssPN} onChange={e=>setForm({...form,jssPN:e.target.value})}/></Field>
          <Field label="Customer PN"><Inp placeholder="e.g. FC-2002" value={form.custPN} onChange={e=>setForm({...form,custPN:e.target.value})}/></Field>
          <Field label="Description"><Inp placeholder="e.g. PAB Module" value={form.description} onChange={e=>setForm({...form,description:e.target.value})}/></Field>
          <Field label="Plant"><Inp placeholder="TRN / MVA / ACU" value={form.plant} onChange={e=>setForm({...form,plant:e.target.value})}/></Field>
        </div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr 1fr 1fr auto',gap:'10px',alignItems:'end'}}>
          <Field label="Volume"><Inp type="number" placeholder="0" value={form.volume} onChange={e=>setForm({...form,volume:e.target.value})}/></Field>
          <Field label="Std Cost ($)"><Inp type="number" placeholder="0.00" value={form.stdCost} onChange={e=>setForm({...form,stdCost:e.target.value})}/></Field>
          <Field label="PC Price / Agreed Rate ($)"><Inp type="number" placeholder="0.00" value={form.pcPrice} onChange={e=>setForm({...form,pcPrice:e.target.value})}/></Field>
          <Field label="SOP Date"><Inp type="date" value={form.sop} onChange={e=>setForm({...form,sop:e.target.value})}/></Field>
          <Field label="Status"><Sel value={form.status} onChange={e=>setForm({...form,status:e.target.value})}>
            <option value="ACTIVE">ACTIVE</option>
            <option value="INACTIVE">INACTIVE</option>
          </Sel></Field>
          <div style={{display:'flex',gap:'5px'}}>
            <Btn variant="accent" onClick={save}>{editId?'Save':'Add'}</Btn>
            {editId&&<Btn variant="ghost" size="sm" onClick={()=>{setForm(blank);setEditId(null);}}>✕</Btn>}
          </div>
        </div>
      </Card>
      <Card style={{padding:'11px 16px',display:'flex',gap:'9px',alignItems:'center'}}>
        <Inp placeholder="🔍 Search parts…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1}}/>
        <Sel value={sf} onChange={e=>setSf(e.target.value)} style={{minWidth:'130px'}}>
          <option value="ALL">All Status</option>
          <option value="ACTIVE">Active Only</option>
          <option value="INACTIVE">Inactive Only</option>
        </Sel>
        <span style={{fontSize:'11.5px',color:C.inkMute}}>{vis.length} parts</span>
      </Card>
      <Card>
        <div style={{overflowX:'auto'}}>
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11.5px'}}>
            <thead><tr>{['OEM','Vehicle','JSS PN','Cust PN','Desc','Plant','Volume','Std Cost','PC Price','SOP','Status',''].map(h=>(
              <th key={h} style={{textAlign:'left',padding:'8px 14px',fontSize:'9px',fontWeight:600,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`,whiteSpace:'nowrap'}}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {vis.map(p=>(
                <tr key={p.id} style={{borderBottom:`1px solid ${C.borderSoft}`,opacity:p.status==='INACTIVE'?0.55:1}}>
                  <td style={{padding:'8px 14px'}}><Pill color={oemColor(p.oem)}>{p.oem}</Pill></td>
                  <td style={{padding:'8px 14px',fontWeight:500}}>{p.vehicle}</td>
                  <td style={{padding:'8px 14px',fontFamily:'monospace',fontSize:'11px',fontWeight:600,color:C.navy}}>{p.jssPN}</td>
                  <td style={{padding:'8px 14px',fontFamily:'monospace',fontSize:'11px',color:'#bbb'}}>{p.custPN}</td>
                  <td style={{padding:'8px 14px'}}>{p.description}</td>
                  <td style={{padding:'8px 14px'}}><Pill color={C.inkMute}>{p.plant}</Pill></td>
                  <td style={{padding:'8px 14px',textAlign:'right'}}>{fmtN(p.volume)}</td>
                  <td style={{padding:'8px 14px',textAlign:'right',color:C.inkMute}}>{fmt$d(p.stdCost)}</td>
                  <td style={{padding:'8px 14px',textAlign:'right',fontWeight:700,color:C.navy}}>{fmt$d(p.pcPrice)}</td>
                  <td style={{padding:'8px 14px',color:C.inkMute,fontSize:'10.5px'}}>{p.sop}</td>
                  <td style={{padding:'8px 14px'}}><Pill color={p.status==='ACTIVE'?C.green:'#999'}>{p.status}</Pill></td>
                  <td style={{padding:'8px 14px',display:'flex',gap:'5px'}}>
                    <Btn variant="ghost" size="xs" onClick={()=>{setForm({...p});setEditId(p.id);}}>Edit</Btn>
                    <Btn variant="danger" size="xs" onClick={()=>del(p.id,p.jssPN)}>Del</Btn>
                  </td>
                </tr>
              ))}
              {vis.length===0&&<tr><td colSpan="12" style={{textAlign:'center',padding:'28px',color:C.inkMute}}>No parts match</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ════
// AUDIT
// ════
function AuditTab({auditLog,setAuditLog,toast$}){
  const [search,setSearch]=useState('');
  const vis=auditLog.filter(e=>!search||(e.action+e.detail).toLowerCase().includes(search.toLowerCase()));
  const ac=a=>({CAMPAIGN_ADD:C.green,CAMPAIGN_EDIT:C.navy,CAMPAIGN_DELETE:C.rose,VIN_ADD:C.green,VIN_DELETE:C.rose,BULK_IMPORT:C.navy,STATUS_CHANGE:C.amber,RECALL_ADD:C.green,RECALL_DELETE:C.rose,VIN_REPAIRED:C.green,PART_ADD:C.green,PART_EDIT:C.navy,PART_DELETE:C.rose}[a]||C.inkMute);
  return(
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end'}}>
        <PH title="Audit Log" sub="Every add, edit, delete, and status change — timestamped automatically"/>
        <div style={{display:'flex',gap:'7px'}}>
          <Btn variant="ghost" size="sm" onClick={()=>downloadCSV(vis.map(e=>({Timestamp:e.ts,Action:e.action,Detail:e.detail})),'audit_log')}>↓ Export CSV</Btn>
          <Btn variant="danger" size="sm" onClick={()=>{setAuditLog([]);toast$('Log cleared','warn');}}>Clear</Btn>
        </div>
      </div>
      <Card style={{padding:'11px 16px',display:'flex',gap:'9px',alignItems:'center'}}>
        <Inp placeholder="🔍 Search log…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1}}/>
        <span style={{fontSize:'11.5px',color:C.inkMute}}>{vis.length} entries</span>
      </Card>
      <Card>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
          <thead><tr>{['Timestamp','Action','Detail'].map(h=>(
            <th key={h} style={{textAlign:'left',padding:'8px 16px',fontSize:'9.5px',fontWeight:600,letterSpacing:'0.08em',color:C.inkMute,textTransform:'uppercase',borderBottom:`1px solid ${C.border}`}}>{h}</th>
          ))}</tr></thead>
          <tbody>
            {vis.map(e=>(
              <tr key={e.id} style={{borderBottom:`1px solid ${C.borderSoft}`}}>
                <td style={{padding:'8px 16px',color:C.inkMute,fontSize:'11px',fontFamily:'monospace',whiteSpace:'nowrap'}}>{new Date(e.ts).toLocaleString()}</td>
                <td style={{padding:'8px 16px'}}><Pill color={ac(e.action)}>{e.action.replace(/_/g,' ')}</Pill></td>
                <td style={{padding:'8px 16px',color:C.ink}}>{e.detail}</td>
              </tr>
            ))}
            {vis.length===0&&<tr><td colSpan="3" style={{textAlign:'center',padding:'28px',color:C.inkMute}}>No log entries — all actions will appear here</td></tr>}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ══════════════════════════════════════════════════════
// WARRANTY MIS SHELL (Increment 1) — six placeholder terminals.
// No wmis.* localStorage keys are read or written in this increment.
// Ford exists in the OEM selector but has no terminal.
// ══════════════════════════════════════════════════════
const WMIS_NAV = [
  { id: 'wmis-overview',            label: 'Executive Overview' },
  { id: 'wmis-recall-intelligence', label: 'Recall Intelligence' },
  { id: 'wmis-gm',                  label: 'GM Workspace' },
  { id: 'wmis-stellantis',          label: 'Stellantis Workspace' },
  { id: 'wmis-parts',               label: 'Parts & Cross-Reference' },
  { id: 'wmis-operations',          label: 'Operations' },
];
const WMIS_TERMINAL_META = {
  'wmis-overview': {
    title: 'Executive Overview',
    purpose: 'Cross-domain management summary for recall exposure, GM warranty billing, Stellantis cost recovery, part mappings, and data-quality exceptions.',
    future: ['Portfolio KPI roll-up','Exposure vs. recovery snapshot','Data-quality exception queue','OEM-level status tiles'],
  },
  'wmis-recall-intelligence': {
    title: 'Recall Intelligence',
    purpose: 'NHTSA recall registry, population and completion analysis, and future JSS exposure forecasting.',
    future: ['NHTSA campaign registry','Population & completion analytics','Outstanding population estimation','JSS exposure forecasting'],
  },
  'wmis-gm': {
    title: 'GM Workspace',
    purpose: 'GM warranty bills, transaction-level claims, Bill Status, LBE Status, allocation values, and claim drill-downs.',
    future: ['GM bill registry','Transaction-level claim viewer','Bill Status / LBE Status tracking','Claim-to-bill allocation drill-down'],
  },
  'wmis-stellantis': {
    title: 'Stellantis Workspace',
    purpose: 'Campaign cases, SWRS groups, debit records, claim allocations, and detailed cost-recovery lines.',
    future: ['Case registry (e.g. Case #4551, #4750)','SWRS group aggregation','Debit records ledger','Claim allocation & cost-recovery lines'],
  },
  'wmis-parts': {
    title: 'Parts & Cross-Reference',
    purpose: 'Future mapping between OEM part numbers and JSS part numbers, including mapping status, side, description, and optional assumptions.',
    future: ['OEM PN ↔ JSS PN mapping','Mapping status (MAPPED / UNMAPPED / ASSUMED / CONFLICT)','Side, description, assumption notes','Bulk import & review workflow'],
  },
  'wmis-operations': {
    title: 'Operations',
    purpose: 'Future Import Center, Data Quality, Reports & Exports, and Warranty MIS Audit Log.',
    future: ['Import Center (batch registry)','Data Quality exceptions','Reports & Exports','Warranty MIS Audit Log'],
  },
};
const WMIS_OEM_OPTIONS = [
  { value: 'ALL',        label: 'All OEMs' },
  { value: 'GM',         label: 'General Motors' },
  { value: 'STELLANTIS', label: 'Stellantis' },
  { value: 'FORD',       label: 'Ford' },
  { value: 'OTHER',      label: 'Other' },
];

function AreaSwitch({area, setArea}){
  const opt = (val, label) => {
    const active = area === val;
    return (
      <button key={val} onClick={()=>setArea(val)} style={{
        flex:1, height:'26px', padding:'0 8px', border:'none', cursor:'pointer',
        borderRadius:'5px', fontSize:'11px', fontWeight:active?600:500,
        background: active ? C.coral : 'transparent',
        color: active ? '#fff' : 'rgba(255,255,255,0.70)',
        transition:'background 120ms, color 120ms', letterSpacing:'0.01em',
      }}>{label}</button>
    );
  };
  return (
    <div style={{marginBottom:'10px'}}>
      <div style={{fontSize:'9px', fontWeight:600, letterSpacing:'0.14em',
        color:'rgba(255,255,255,0.45)', textTransform:'uppercase',
        padding:'0 4px 5px'}}>Application Area</div>
      <div style={{display:'flex', gap:'3px', padding:'3px',
        background:'rgba(0,0,0,0.28)', borderRadius:'7px',
        border:'1px solid rgba(255,255,255,0.08)'}}>
        {opt('recall','Recall Database')}
        {opt('warranty','Warranty MIS')}
      </div>
    </div>
  );
}

function WMIS_Placeholder({id}){
  const meta = WMIS_TERMINAL_META[id];
  if(!meta) return <div style={{padding:'24px',color:C.inkMute}}>Unknown terminal: {id}</div>;
  return (
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <PH title={meta.title} sub={meta.purpose}/>
      <Card style={{padding:'28px 24px'}}>
        <div style={{display:'flex',alignItems:'center',gap:'10px',marginBottom:'14px'}}>
          <Pill color={C.amber}>Planned — Not yet implemented</Pill>
          <span style={{fontSize:'11.5px',color:C.inkMute}}>Shell placeholder · no data written</span>
        </div>
        <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.08em',
          textTransform:'uppercase',color:C.inkMute,marginBottom:'8px'}}>Future contents</div>
        <ul style={{margin:0,paddingLeft:'18px',fontSize:'13px',color:C.ink,lineHeight:1.7}}>
          {meta.future.map((f,i)=><li key={i}>{f}</li>)}
        </ul>
        <div style={{marginTop:'18px',padding:'12px 14px',background:C.panelAlt,
          border:`1px dashed ${C.border}`,borderRadius:S.r.md,fontSize:'11.5px',color:C.inkMute,lineHeight:1.55}}>
          This terminal is a shell only. No KPIs, claim totals, imports, or persisted
          data are active. It will be built out in a future increment.
        </div>
      </Card>
    </div>
  );
}


// ══════════════════════════════════════════════════════
// RECALL INTELLIGENCE — Increment 2A (read-only)
// Source: uploaded NHTSA HPH Recalls workbook (July 6 2026)
// No wmis.* localStorage reads or writes in this increment.
// ══════════════════════════════════════════════════════
const NHTSA_RECALLS_SEED = [
  {
    "_id": "nhtsa:20V736000",
    "nhtsaId": "20V736000",
    "reportReceivedDate": "2020-11-25",
    "sourceManufacturer": "General Motors, LLC",
    "normalizedOem": "General Motors",
    "recallLink": "Go to Recall",
    "subject": "Roof Rail Air Bag Inflator Endcap may Detach",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "N202309680",
    "mfrCampaignNumbers": [
      "N202309680"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 9279,
    "completionRatePct": null,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "General Motors LLC (GM) is recalling certain 2015-2016 GMC Sierra 1500 and Chevrolet Silverado 1500 and 2015 GMC Sierra 2500 and 3500, and Chevrolet Silverado 2500 and 3500 trucks. The roof-rail air bag (RRAB) inflator end cap may detach from the inflator.",
    "consequenceSummary": "If the end cap separates from the inflator, the compressed gas will escape and the end cap can be propelled into the vehicle, increasing the risk of injury or a crash.",
    "correctiveAction": "GM will notify owners, and dealers will replace the RRAB modules on the left and/or right side, free of charge. Interim owner notification letters were mailed on December 22, 2020. Owners may contact GMC customer service at 1-888-988-7267 or Chevrolet customer service at 1-800-222-1020. Note: GM filed a new recall, number 21V-504 that supersedes this recall. All vehicles previously recalled under 20V-736 will now be repaired under the new recall number 21V-504. GM's numbers for this recall are N202309680, N202309681, and N202309682.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": [
      {
        "type": "supersededBy",
        "otherNhtsaId": "21V504000",
        "reviewed": false
      }
    ]
  },
  {
    "_id": "nhtsa:21V504000",
    "nhtsaId": "21V504000",
    "reportReceivedDate": "2021-07-01",
    "sourceManufacturer": "General Motors, LLC",
    "normalizedOem": "General Motors",
    "recallLink": "Go to Recall",
    "subject": "Roof Rail Air Bag Inflator May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "N202324251",
    "mfrCampaignNumbers": [
      "N202324251"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 410019,
    "completionRatePct": 59.68,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "General Motors LLC (GM) is recalling certain 2015-2016 GMC Sierra 1500, 2500, 3500, and Chevrolet Silverado 1500, 2500, and 3500 trucks. The roof-rail air bag (RRAB) inflator end cap may detach from the inflator, or the inflator sidewall may rupture.",
    "consequenceSummary": "A separated inflator end cap or inflator rupture can allow compressed gas to escape, resulting in the end cap or other components being propelled into the vehicle, increasing the risk of injury or crash.",
    "correctiveAction": "Dealers will replace the left and right side RRAB modules, free of charge. Interim owner notification letters informing owners of the safety risk were mailed on August 18, 2021. Owners will receive a second notice when the remedy is available. Phase I owner notification letters were mailed on May 5, 2022. Phase II owner notification letters are expected to be mailed on May 31, 2022. Owners may contact GMC customer service at 1-888-988-7267 or Chevrolet customer service at 1-800-222-1020. This recall supersedes NHTSA recall number 20V-736. GM's number for this recall is N202324251.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": [
      {
        "type": "supersedes",
        "otherNhtsaId": "20V736000",
        "reviewed": false
      }
    ]
  },
  {
    "_id": "nhtsa:21E074000",
    "nhtsaId": "21E074000",
    "reportReceivedDate": "2021-08-12",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflator May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "Y56",
    "mfrCampaignNumbers": [
      "Y56"
    ],
    "recallType": "Equipment",
    "rawListedPotentialPopulation": 5373,
    "completionRatePct": 0.07,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) is recalling certain Mopar right and left-side Crew Cab, Quad Cab, Mega Cab, and Standard Cab side air bag inflatable curtains (SABIC) sold as replacement parts. Please see the part 573 report for a list of affected part numbers. The SABIC inflator end cap may detach from the inflator, or the inflator sidewall may rupture.",
    "consequenceSummary": "A separated inflator end cap or inflator rupture can allow compressed gas to escape, resulting in the end cap or other components being propelled towards a person, or into a vehicle, increasing the risk of injury or crash.",
    "correctiveAction": "Dealers will inspect and replace as necessary, the SABIC assemblies, free of charge. Interim owner notification letters were mailed on October 4, 2021. Owner notification letters were mailed on May 26, 2022. Owners may contact FCA US, LLC customer service at 1-800-853-1403. FCA US, LLC's number for this recall is Y56.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:21V632000",
    "nhtsaId": "21V632000",
    "reportReceivedDate": "2021-08-12",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflator May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "Y55",
    "mfrCampaignNumbers": [
      "Y55"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 212373,
    "completionRatePct": 47.41,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) is recalling certain 2015-2020 Ram 1500 Classic, 2015-2016 Ram 3500, and Ram 2500 vehicles, and 2016 Ram 3500 Cab Chassis with a gross vehicle weight rating less than 10,000 lbs. The side air bag inflatable curtain (SABIC) inflator end cap may detach from the inflator, or the inflator sidewall may rupture.",
    "consequenceSummary": "A separated inflator end cap or inflator rupture can allow compressed gas to escape, resulting in the end cap or other components being propelled into the vehicle, increasing the risk of injury or crash.",
    "correctiveAction": "Dealers will inspect and replace as necessary, the SABIC assemblies, free of charge. Interim owner notification letters were mailed on October 4, 2021. Owner notification letters were mailed on May 26, 2022. Owners may contact FCA US, LLC customer service at 1-800-853-1403. FCA US, LLC's number for this recall is Y55.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:24V198000",
    "nhtsaId": "24V198000",
    "reportReceivedDate": "2024-03-14",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflators May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "19B",
    "mfrCampaignNumbers": [
      "19B"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 284982,
    "completionRatePct": 28.29,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) (Stellantis) is recalling certain 2018-2021 Dodge Charger and Chrysler 300 vehicles. The right and left side curtain air bag inflators may rupture due to a manufacturing defect.",
    "consequenceSummary": "An inflator rupture may result in sharp metal fragments striking occupants, resulting in injury or death.",
    "correctiveAction": "Dealers will replace both side curtain air bags, free of charge. Owner notification letters were mailed beginning February 20, 2025. Owners may contact FCA US, LLC customer service at 1-800-853-1403. FCA US, LLC's number for this recall is 19B.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:24E024000",
    "nhtsaId": "24E024000",
    "reportReceivedDate": "2024-03-14",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflators May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "20B",
    "mfrCampaignNumbers": [
      "20B"
    ],
    "recallType": "Equipment",
    "rawListedPotentialPopulation": 982,
    "completionRatePct": 0.0,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) is recalling certain MOPAR right and left side curtain air bags, with part numbers 68222742AF, 68222743AF, 05108050AF, and 05108049AF, sold as replacement parts for certain 2011-2019 Dodge Charger and 2011-2020 Chrysler 300 vehicles. The replacement right and left side curtain air bag inflators may rupture due to a manufacturing defect.",
    "consequenceSummary": "An inflator rupture may result in sharp metal fragments striking occupants, resulting in injury or death. In addition, an inflator rupture when not installed in a vehicle can result in injury or death of people nearby.",
    "correctiveAction": "Dealers will replace or repurchase the side curtain air bags, free of charge. Owner notification letters were mailed April 30, 2025. Owners may contact FCA customer service at 1-800-853-1403. FCA's number for this recall is 20B.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:25V010000",
    "nhtsaId": "25V010000",
    "reportReceivedDate": "2025-01-16",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflators May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "05C, 09C, 10C",
    "mfrCampaignNumbers": [
      "05C",
      "09C",
      "10C"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 6066,
    "completionRatePct": 30.65,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) is recalling certain 2016-2019 Ram 3500, Ram 1500, 2016-2020 Ram 2500, and 2016 Ram 3500 Cab Chassis vehicles. The right and left side curtain air bag inflators may rupture due to a manufacturing defect.",
    "consequenceSummary": "An inflator rupture may result in sharp metal fragments striking occupants, resulting in injury.",
    "correctiveAction": "Dealers will replace one or both side curtain air bags, as necessary, free of charge. Owner notification letters were mailed February 13, 2025. Owners may contact FCA customer service at 1-800-853-1403. FCA's numbers for this recall are 05C, 09C, and 10C.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:25V432000",
    "nhtsaId": "25V432000",
    "reportReceivedDate": "2025-06-26",
    "sourceManufacturer": "General Motors, LLC",
    "normalizedOem": "General Motors",
    "recallLink": "Go to Recall",
    "subject": "Roof Rail Air Bag Inflator Endcap May Detach",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "N252513060",
    "mfrCampaignNumbers": [
      "N252513060"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 1658,
    "completionRatePct": 39.53,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "General Motors (GM) is recalling certain 2018 GMC Sierra 1500, 2019 Sierra 2500, 3500, Chevrolet Silverado 3500, and Chevrolet Silverado 2500 and 2018 Chevrolet Silverado 1500 Crew Cab vehicles. The left or right-side roof-rail air bag (RRAB) inflator end cap may detach from the inflator or the inflator sidewall may rupture.",
    "consequenceSummary": "A detached end cap or inflator rupture can allow compressed gas to escape and project the end cap or fragments of the inflator into the vehicle, increases the risk of injury or crash.",
    "correctiveAction": "Dealers will replace the left and right side RRAB modules, free of charge. Owner notification letters were mailed July 23, 2025. Owners may contact GMC customer service at 1-888-988-7267 or Chevrolet customer service at 1-800-222-1020. GM's number for this recall is N252513060.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:25V592000",
    "nhtsaId": "25V592000",
    "reportReceivedDate": "2025-09-11",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflators May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "83C, 91C, 92C",
    "mfrCampaignNumbers": [
      "83C",
      "91C",
      "92C"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 1761,
    "completionRatePct": 21.31,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) is recalling certain 2018 Ram 1500, Ram 2500, and Ram 3500 vehicles. The right and left side curtain air bag inflators may rupture due to a manufacturing defect.",
    "consequenceSummary": "An inflator rupture may result in sharp metal fragments striking occupants, resulting in injury.",
    "correctiveAction": "Dealers will replace one or both side curtain air bags, as necessary, free of charge. Owner notification letters were mailed September 25, 2025. Owners may contact FCA customer service at 1-800-853-1403. FCA's numbers for this recall are 83C, 91C, and 92C. Vehicle Identification Numbers (VINs) involved in this recall will be searchable on NHTSA.gov beginning September 18, 2025.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:25V824000",
    "nhtsaId": "25V824000",
    "reportReceivedDate": "2025-12-01",
    "sourceManufacturer": "Chrysler (FCA US, LLC)",
    "normalizedOem": "Stellantis",
    "recallLink": "Go to Recall",
    "subject": "Side Curtain Air Bag Inflators May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "C2C, C7C, C8C",
    "mfrCampaignNumbers": [
      "C2C",
      "C7C",
      "C8C"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 1589,
    "completionRatePct": 5.37,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "Chrysler (FCA US, LLC) is recalling certain 2019 Ram 1500 Classic, Ram 2500, and Ram 3500 vehicles. The right and left side curtain air bag inflators may rupture due to a manufacturing defect.",
    "consequenceSummary": "An inflator rupture may result in sharp metal fragments striking occupants, resulting in injury.",
    "correctiveAction": "Dealers will replace one or both side curtain air bags, as necessary, free of charge. Owner notification letters were mailed December 18, 2025. Owners may contact FCA customer service at 1-800-853-1403. FCA's numbers for this recall are C2C, C7C, and C8C. Vehicle Identification Numbers (VINs) involved in this recall will be searchable on NHTSA.gov beginning December 4, 2025.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:26V166000",
    "nhtsaId": "26V166000",
    "reportReceivedDate": "2026-03-19",
    "sourceManufacturer": "General Motors, LLC",
    "normalizedOem": "General Motors",
    "recallLink": "Go to Recall",
    "subject": "Roof Rail Air Bag Inflator Endcap May Detach",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "N262549710",
    "mfrCampaignNumbers": [
      "N262549710"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 2819,
    "completionRatePct": null,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "General Motors, LLC (GM) is recalling certain 2018 Chevrolet Silverado 1500, GMC Sierra 1500, 2019 Chevrolet Silverado 2500, Silverado 3500, GMC Sierra 2500, and Sierra 3500 Crew Cab vehicles. The left or right-side roof rail air bag (RRAB) inflator end cap may detach from the inflator or the inflator sidewall may rupture.",
    "consequenceSummary": "A detached end cap or inflator rupture can allow compressed gas to escape and project the end cap or fragments of the inflator into the vehicle, increasing the risk of injury.",
    "correctiveAction": "Dealers will replace the left and right-side RRAB modules, free of charge. Owner notification letters were mailed April 16, 2026. Owners may contact GMC customer service at 1-888-988-7267 or Chevrolet customer service at 1-800-222-1020. GM's number for this recall is N262549710. Vehicle Identification Numbers (VINs) involved in this recall became searchable on NHTSA.gov on March 19, 2026.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": [
      {
        "type": "expandedBy",
        "otherNhtsaId": "26V325000",
        "reviewed": false
      }
    ]
  },
  {
    "_id": "nhtsa:26V329000",
    "nhtsaId": "26V329000",
    "reportReceivedDate": "2026-05-21",
    "sourceManufacturer": "General Motors, LLC",
    "normalizedOem": "General Motors",
    "recallLink": "Go to Recall",
    "subject": "Air Bag Inflator May Rupture",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "N262555870",
    "mfrCampaignNumbers": [
      "N262555870"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 4125,
    "completionRatePct": null,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "General Motors, LLC (GM) is recalling certain 2015 Cadillac Escalade, Escalade ESV, XTS, 2015-2016 Cadillac ATS, CTS, and SRX vehicles. The front-driver air bag inflator may rupture during deployment.",
    "consequenceSummary": "An inflator rupture may result in sharp metal fragments striking the driver or other occupants, resulting in serious injury or death.",
    "correctiveAction": "Dealers will replace the front-driver air bag module, free of charge. Interim letters notifying owners of the safety risk were mailed June 19, 2026. Additional letters will be mailed once the remedy is available. Owners may contact Cadillac customer service at 1-800-333-4223. The manufacturer's number for this recall is N262555870. Vehicle Identification Numbers (VINs) involved in this recall became searchable on NHTSA.gov on May 21, 2026.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": []
  },
  {
    "_id": "nhtsa:26V325000",
    "nhtsaId": "26V325000",
    "reportReceivedDate": "2026-05-21",
    "sourceManufacturer": "General Motors, LLC",
    "normalizedOem": "General Motors",
    "recallLink": "Go to Recall",
    "subject": "Roof Rail Air Bag Inflator Endcap May Detach",
    "component": "AIR BAGS",
    "mfrCampaignNumberRaw": "N262557310",
    "mfrCampaignNumbers": [
      "N262557310"
    ],
    "recallType": "Vehicle",
    "rawListedPotentialPopulation": 2785,
    "completionRatePct": null,
    "sourceStatusRaw": null,
    "normalizedStatus": "Unknown",
    "statusConfidence": "Low",
    "recallDescription": "General Motors, LLC (GM) is recalling certain 2018 Chevrolet Silverado 1500, GMC Sierra 1500, 2019 Chevrolet Silverado 2500, Silverado 3500, GMC Sierra 2500, and Sierra 3500 vehicles. The left-side or right-side roof rail air bag (RRAB) inflator end cap may detach from the inflator or the inflator sidewall may rupture.",
    "consequenceSummary": "A detached end cap or inflator rupture can allow compressed gas to escape and project the end cap or fragments of the inflator into the vehicle, increasing the risk of injury.",
    "correctiveAction": "Dealers will replace both RRAB modules, free of charge. Owner notification letters are expected to be mailed July 6, 2026. Owners may contact GMC customer service at 1-888-988-7267 or Chevrolet customer service at 1-800-222-1020. GM's number for this recall is N262557310. Vehicle Identification Numbers (VINs) involved in this recall became searchable on NHTSA.gov on May 21, 2026. This recall expands previous NHTSA recall number 26V166.",
    "parkOutsideAdvisory": "No",
    "doNotDriveAdvisory": "No",
    "relationships": [
      {
        "type": "expands",
        "otherNhtsaId": "26V166000",
        "reviewed": false
      }
    ]
  }
];


const RI_STATUS_OPTIONS = [
  {value:'ALL',      label:'All statuses'},
  {value:'Active',   label:'Active'},
  {value:'Completed',label:'Completed'},
  {value:'Unknown',  label:'Unknown'},
];
const RI_COMPLETION_OPTIONS = [
  {value:'ALL',        label:'All completion'},
  {value:'REPORTED',   label:'Reported'},
  {value:'NOT_REPORTED',label:'Not Reported'},
  {value:'ZERO',       label:'0%'},
  {value:'GT_ZERO',    label:'Greater than 0%'},
];
const RI_OEM_OPTIONS = [
  {value:'ALL',label:'All OEMs'},
  {value:'General Motors',label:'General Motors'},
  {value:'Stellantis',label:'Stellantis'},
  {value:'Other',label:'Other'},
];
const RI_TYPE_OPTIONS = [
  {value:'ALL',label:'All types'},
  {value:'Vehicle',label:'Vehicle'},
  {value:'Equipment',label:'Equipment'},
];

function riFmtInt(n){
  if(n===null||n===undefined||isNaN(n)) return '—';
  return Math.round(n).toLocaleString('en-US');
}
function riFmtPct(n){
  if(n===null||n===undefined) return 'Not Reported';
  return (Math.round(n*100)/100).toFixed(2)+'%';
}
function riEstCompleted(rec){
  if(rec.completionRatePct===null||rec.completionRatePct===undefined) return null;
  if(rec.rawListedPotentialPopulation===null||rec.rawListedPotentialPopulation===undefined) return null;
  return rec.rawListedPotentialPopulation * rec.completionRatePct / 100;
}
function riEstOutstanding(rec){
  const c = riEstCompleted(rec);
  if(c===null) return null;
  return rec.rawListedPotentialPopulation - c;
}
function riRelIndicator(rec){
  if(!rec.relationships||rec.relationships.length===0) return null;
  const types = rec.relationships.map(r=>r.type);
  const short = {supersedes:'Supersedes',supersededBy:'Superseded by',expands:'Expands',expandedBy:'Expanded by'};
  return types.map(t=>short[t]||t).join(' · ');
}
function getNhtsaRecallUrl(nhtsaId) {
  return `https://www.nhtsa.gov/recalls?nhtsaId=${encodeURIComponent(nhtsaId)}`;
}

function RI_KPI({label,value,sub,accent}){
  return (
    <Card style={{padding:'14px 16px',minWidth:0,borderTop:`3px solid ${accent||C.navy}`}}>
      <div style={{fontSize:'10.5px',fontWeight:600,letterSpacing:'0.10em',
        textTransform:'uppercase',color:C.inkMute,marginBottom:'6px'}}>{label}</div>
      <div style={{fontSize:'22px',fontWeight:600,color:C.ink,fontVariantNumeric:'tabular-nums',lineHeight:1.1}}>{value}</div>
      {sub && <div style={{fontSize:'11px',color:C.inkMute,marginTop:'4px',lineHeight:1.4}}>{sub}</div>}
    </Card>
  );
}

function RecallDetail({rec,onClose}){
  if(!rec) return null;
  const est = riEstCompleted(rec);
  const out = riEstOutstanding(rec);
  const Section = ({title,children}) => (
    <div style={{marginBottom:'16px'}}>
      <div style={{fontSize:'10.5px',fontWeight:700,letterSpacing:'0.10em',
        textTransform:'uppercase',color:C.coral,marginBottom:'8px',
        paddingBottom:'4px',borderBottom:`1px solid ${C.borderSoft}`}}>{title}</div>
      <div style={{fontSize:'12.5px',color:C.ink,lineHeight:1.6}}>{children}</div>
    </div>
  );
  const KV = ({k,v}) => (
    <div style={{display:'grid',gridTemplateColumns:'180px 1fr',gap:'8px',marginBottom:'4px'}}>
      <span style={{color:C.inkMute,fontSize:'11.5px'}}>{k}</span>
      <span style={{color:C.ink,fontSize:'12.5px'}}>{v===null||v===undefined||v===''?'—':v}</span>
    </div>
  );
  return (
    <div style={{position:'fixed',inset:0,background:'rgba(15,23,42,0.45)',
      backdropFilter:'blur(2px)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:9998}}
      onClick={onClose}>
      <Card onClick={e=>e.stopPropagation()}
        style={{padding:'22px 26px',maxWidth:'860px',width:'92%',maxHeight:'86vh',
          overflowY:'auto',boxShadow:S.elev3,borderRadius:S.r.xl}}>
        <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:'14px',gap:'12px'}}>
          <div>
            <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.10em',
              textTransform:'uppercase',color:C.inkMute}}>NHTSA Recall</div>
            <div style={{fontSize:'20px',fontWeight:700,color:C.ink,marginTop:'2px'}}>{rec.nhtsaId}</div>
            <div style={{fontSize:'12.5px',color:C.inkSub,marginTop:'2px'}}>{rec.subject}</div>
          </div>
          <Btn variant="ghost" size="sm" onClick={onClose}>Close</Btn>
        </div>

        <Section title="Identity">
          <KV k="NHTSA ID" v={rec.nhtsaId}/>
          <KV k="Source Manufacturer" v={rec.sourceManufacturer}/>
          <KV k="Normalized OEM" v={rec.normalizedOem}/>
          <KV k="Manufacturer Campaign(s)" v={rec.mfrCampaignNumbers.length?rec.mfrCampaignNumbers.join(', '):rec.mfrCampaignNumberRaw}/>
          <KV k="Recall Type" v={rec.recallType}/>
          <KV k="Report Received" v={rec.reportReceivedDate}/>
          <KV k="Source Link" v={rec.nhtsaId
            ? <a href={getNhtsaRecallUrl(rec.nhtsaId)} target="_blank" rel="noopener noreferrer" style={{color:C.coral,textDecoration:'underline'}}>View on NHTSA.gov</a>
            : '—'}/>
        </Section>

        <Section title="Population">
          <KV k="Raw Listed Potential Population" v={riFmtInt(rec.rawListedPotentialPopulation)}/>
          <KV k="Completion Rate" v={riFmtPct(rec.completionRatePct)}/>
          <KV k="Estimated Completed" v={est===null?'—':riFmtInt(est)}/>
          <KV k="Estimated Outstanding" v={out===null?'—':riFmtInt(out)}/>
        </Section>

        <Section title="Issue">
          <KV k="Subject" v={rec.subject}/>
          <KV k="Component" v={rec.component}/>
          <div style={{marginTop:'8px'}}>
            <div style={{color:C.inkMute,fontSize:'11.5px',marginBottom:'2px'}}>Recall Description</div>
            <div style={{whiteSpace:'pre-wrap'}}>{rec.recallDescription||'—'}</div>
          </div>
          <div style={{marginTop:'8px'}}>
            <div style={{color:C.inkMute,fontSize:'11.5px',marginBottom:'2px'}}>Consequence Summary</div>
            <div style={{whiteSpace:'pre-wrap'}}>{rec.consequenceSummary||'—'}</div>
          </div>
          <div style={{marginTop:'8px'}}>
            <div style={{color:C.inkMute,fontSize:'11.5px',marginBottom:'2px'}}>Corrective Action</div>
            <div style={{whiteSpace:'pre-wrap'}}>{rec.correctiveAction||'—'}</div>
          </div>
        </Section>

        <Section title="Safety">
          <KV k="Park Outside Advisory" v={rec.parkOutsideAdvisory||'—'}/>
          <KV k="Do Not Drive Advisory"  v={rec.doNotDriveAdvisory||'—'}/>
        </Section>

        <Section title="Status">
          <KV k="Source Status Raw"  v={rec.sourceStatusRaw}/>
          <KV k="Normalized Status"  v={rec.normalizedStatus}/>
          <KV k="Status Confidence"  v={rec.statusConfidence}/>
        </Section>

        <Section title="Relationships">
          {rec.relationships && rec.relationships.length>0 ? (
            <div>
              {rec.relationships.map((r,i)=>(
                <div key={i} style={{marginBottom:'4px'}}>
                  <b>{r.type}</b> → {r.otherNhtsaId} <span style={{color:C.inkMute,fontSize:'11px'}}>(reviewed: {String(r.reviewed)})</span>
                </div>
              ))}
              <div style={{marginTop:'8px',fontSize:'11.5px',color:C.amber}}>
                Caution: relationships have NOT been used to deduplicate population.
              </div>
            </div>
          ) : <div style={{color:C.inkMute}}>No source-identified relationships.</div>}
        </Section>
      </Card>
    </div>
  );
}

function RecallIntelligence(){
  const [tab,setTab] = useState('registry'); // registry | popcomp | exposure
  const [search,setSearch] = useState('');
  const [oem,setOem] = useState('ALL');
  const [rtype,setRtype] = useState('ALL');
  const [compFilt,setCompFilt] = useState('ALL');
  const [statusFilt,setStatusFilt] = useState('ALL');
  const [dateFrom,setDateFrom] = useState('');
  const [dateTo,setDateTo] = useState('');
  const [selected,setSelected] = useState(null);

  const records = NHTSA_RECALLS_SEED;
  const filtered = records.filter(r=>{
    if(oem!=='ALL' && r.normalizedOem!==oem) return false;
    if(rtype!=='ALL' && r.recallType!==rtype) return false;
    if(statusFilt!=='ALL' && r.normalizedStatus!==statusFilt) return false;
    if(compFilt==='REPORTED'    && r.completionRatePct===null) return false;
    if(compFilt==='NOT_REPORTED'&& r.completionRatePct!==null) return false;
    if(compFilt==='ZERO'        && !(r.completionRatePct===0)) return false;
    if(compFilt==='GT_ZERO'     && !(r.completionRatePct!==null && r.completionRatePct>0)) return false;
    if(dateFrom && (!r.reportReceivedDate || r.reportReceivedDate < dateFrom)) return false;
    if(dateTo   && (!r.reportReceivedDate || r.reportReceivedDate > dateTo)) return false;
    if(search){
      const q = search.toLowerCase();
      const hay = [r.nhtsaId,r.sourceManufacturer,r.mfrCampaignNumberRaw,
        r.subject,r.component,r.recallDescription].filter(Boolean).join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });

  const clearFilters = () => {
    setSearch(''); setOem('ALL'); setRtype('ALL');
    setCompFilt('ALL'); setStatusFilt('ALL');
    setDateFrom(''); setDateTo('');
  };

  const rawTotal = filtered.reduce((s,r)=>s+(r.rawListedPotentialPopulation||0),0);
  const reported = filtered.filter(r=>r.completionRatePct!==null);
  const estCompletedTotal = reported.reduce((s,r)=>s+(riEstCompleted(r)||0),0);
  const estOutstandingTotal = reported.reduce((s,r)=>s+(riEstOutstanding(r)||0),0);
  const notReportedCount = filtered.filter(r=>r.completionRatePct===null).length;

  const tabBtn = (id,label) => (
    <button key={id} onClick={()=>setTab(id)} style={{
      height:'32px',padding:'0 14px',border:'none',cursor:'pointer',
      fontSize:'12.5px',fontWeight:tab===id?600:500,
      color:tab===id?C.ink:C.inkMute,
      background:tab===id?C.panel:'transparent',
      borderBottom:tab===id?`2px solid ${C.coral}`:`2px solid transparent`}}>{label}</button>
  );

  return (
    <div style={{display:'flex',flexDirection:'column',gap:'14px'}}>
      <PH title="Recall Intelligence"
          sub={`NHTSA recall registry — ${records.length} records loaded from source workbook (read-only prototype)`}/>

      <div style={{display:'flex',gap:'2px',borderBottom:`1px solid ${C.border}`}}>
        {tabBtn('registry','Recall Registry')}
        {tabBtn('popcomp','Population & Completion')}
        {tabBtn('exposure','Exposure Forecasting')}
      </div>

      {(tab==='registry' || tab==='popcomp') && (
        <Card style={{padding:'12px 14px'}}>
          <div style={{display:'flex',flexWrap:'wrap',gap:'8px',alignItems:'flex-end'}}>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>Search</div>
              <Inp placeholder="NHTSA ID, subject, campaign…" value={search}
                onChange={e=>setSearch(e.target.value)} style={{width:'240px'}}/>
            </div>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>OEM</div>
              <Sel value={oem} onChange={e=>setOem(e.target.value)}>
                {RI_OEM_OPTIONS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
              </Sel>
            </div>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>Recall Type</div>
              <Sel value={rtype} onChange={e=>setRtype(e.target.value)}>
                {RI_TYPE_OPTIONS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
              </Sel>
            </div>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>Completion</div>
              <Sel value={compFilt} onChange={e=>setCompFilt(e.target.value)}>
                {RI_COMPLETION_OPTIONS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
              </Sel>
            </div>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>Normalized Status</div>
              <Sel value={statusFilt} onChange={e=>setStatusFilt(e.target.value)}>
                {RI_STATUS_OPTIONS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
              </Sel>
            </div>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>Report Received Date From</div>
              <Inp type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)}/>
            </div>
            <div>
              <div style={{fontSize:'10.5px',color:C.inkMute,marginBottom:'2px'}}>To</div>
              <Inp type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)}/>
            </div>
            <Btn variant="ghost" size="sm" onClick={clearFilters}>Clear Filters</Btn>
            <div style={{marginLeft:'auto',fontSize:'12px',color:C.inkMute}}>
              <b style={{color:C.ink}}>{filtered.length}</b> of {records.length} recalls
            </div>
          </div>
        </Card>
      )}

      {tab==='registry' && (
        <Card style={{padding:'0',overflow:'hidden'}}>
          <div style={{overflowX:'auto'}}>
            <table className="jss-table" style={{width:'100%',minWidth:'1180px',borderCollapse:'collapse',fontSize:'12px'}}>
              <thead>
                <tr style={{background:C.panelAlt,borderBottom:`1px solid ${C.border}`}}>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>NHTSA ID</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>OEM</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Mfr Campaign</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Report Received</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Subject</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Type</th>
                  <th style={{textAlign:'right',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Potentially Affected</th>
                  <th style={{textAlign:'right',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Completion</th>
                  <th style={{textAlign:'right',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Est. Outstanding</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Status</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Relationship</th>
                  <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:C.inkSub}}>Source</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map(r=>{
                  const out = riEstOutstanding(r);
                  const rel = riRelIndicator(r);
                  return (
                    <tr key={r._id} onClick={()=>setSelected(r)}
                      style={{borderBottom:`1px solid ${C.borderSoft}`,cursor:'pointer'}}
                      onMouseEnter={e=>e.currentTarget.style.background=C.panelAlt}
                      onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                      <td style={{padding:'9px 12px',fontWeight:600,color:C.coral}}>{r.nhtsaId}</td>
                      <td style={{padding:'9px 12px',color:C.ink}}>{r.normalizedOem}</td>
                      <td style={{padding:'9px 12px',color:C.ink}}>
                        {r.mfrCampaignNumbers.length>1
                          ? r.mfrCampaignNumbers.map((c,i)=>(
                              <span key={i} style={{display:'inline-block',padding:'1px 6px',marginRight:'3px',
                                background:C.tintNavy,border:`1px solid ${C.border}`,borderRadius:'3px',fontSize:'11px'}}>{c}</span>))
                          : (r.mfrCampaignNumberRaw||'—')}
                      </td>
                      <td style={{padding:'9px 12px',color:C.inkSub,fontVariantNumeric:'tabular-nums'}}>{r.reportReceivedDate||'—'}</td>
                      <td style={{padding:'9px 12px',color:C.ink,maxWidth:'320px'}}>{r.subject}</td>
                      <td style={{padding:'9px 12px',color:C.inkSub}}>{r.recallType||'—'}</td>
                      <td style={{padding:'9px 12px',textAlign:'right',fontVariantNumeric:'tabular-nums',color:C.ink}}>
                        {riFmtInt(r.rawListedPotentialPopulation)}</td>
                      <td style={{padding:'9px 12px',textAlign:'right',fontVariantNumeric:'tabular-nums',
                        color:r.completionRatePct===null?C.inkFaint:C.ink}}>
                        {riFmtPct(r.completionRatePct)}</td>
                      <td style={{padding:'9px 12px',textAlign:'right',fontVariantNumeric:'tabular-nums',
                        color:out===null?C.inkFaint:C.ink}}>
                        {out===null?'—':riFmtInt(out)}</td>
                      <td style={{padding:'9px 12px'}}>
                        <Pill color={C.inkMute}>{r.normalizedStatus}</Pill>
                      </td>
                      <td style={{padding:'9px 12px',fontSize:'11px',color:C.amber}}>{rel||''}</td>
                      <td style={{padding:'9px 12px'}}>
                        {r.nhtsaId
                          ? <a href={getNhtsaRecallUrl(r.nhtsaId)} target="_blank" rel="noopener noreferrer"
                              onClick={e=>e.stopPropagation()}
                              style={{color:C.coral,textDecoration:'underline',fontSize:'11.5px'}}>NHTSA →</a>
                          : '—'}
                      </td>
                    </tr>
                  );
                })}
                {filtered.length===0 && (
                  <tr><td colSpan={12} style={{padding:'24px',textAlign:'center',color:C.inkMute}}>
                    No recalls match current filters.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {tab==='popcomp' && (
        <>
          <div style={{display:'grid',gridTemplateColumns:'repeat(5, minmax(0, 1fr))',gap:'10px'}}>
            <RI_KPI label="Recall Records" value={filtered.length} accent={C.navy}/>
            <RI_KPI label="Raw Listed Potential Population" value={riFmtInt(rawTotal)}
              sub="May include overlap from superseded/expanded recalls" accent={C.coral}/>
            <RI_KPI label="Estimated Completed" value={riFmtInt(estCompletedTotal)}
              sub="Excludes recalls without reported completion" accent={C.green}/>
            <RI_KPI label="Estimated Outstanding" value={riFmtInt(estOutstandingTotal)}
              sub="Excludes recalls without reported completion" accent={C.amber}/>
            <RI_KPI label="Completion Rate Not Reported" value={notReportedCount}
              sub={`of ${filtered.length} filtered records`} accent={C.rose}/>
          </div>

          <Card style={{padding:'12px 14px',background:C.tintAmber,border:`1px solid ${C.amber}`}}>
            <div style={{fontSize:'11.5px',color:C.ink,lineHeight:1.5}}>
              <b>Caution:</b> Raw Listed Potential Population may include overlap from superseded or expanded recalls.
              This total is not deduplicated. Relationship-Adjusted Population is not yet calculated.
            </div>
          </Card>

          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px'}}>
            <Card style={{padding:'14px 16px'}}>
              <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.08em',
                textTransform:'uppercase',color:C.inkMute,marginBottom:'10px'}}>Population by Normalized OEM</div>
              {(() => {
                const groups = {};
                filtered.forEach(r=>{
                  if(!groups[r.normalizedOem]) groups[r.normalizedOem]={count:0,pop:0};
                  groups[r.normalizedOem].count++;
                  groups[r.normalizedOem].pop += r.rawListedPotentialPopulation||0;
                });
                const rows = Object.entries(groups).sort((a,b)=>b[1].pop-a[1].pop);
                const max = Math.max(1, ...rows.map(([,v])=>v.pop));
                return rows.map(([k,v])=>(
                  <div key={k} style={{marginBottom:'8px'}}>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:'12px',marginBottom:'3px'}}>
                      <span style={{color:C.ink}}>{k} <span style={{color:C.inkMute}}>({v.count})</span></span>
                      <span style={{fontVariantNumeric:'tabular-nums',color:C.inkSub}}>{riFmtInt(v.pop)}</span>
                    </div>
                    <div style={{height:'6px',background:C.tintNeutral,borderRadius:'3px',overflow:'hidden'}}>
                      <div style={{height:'100%',width:`${(v.pop/max)*100}%`,background:C.coral}}/>
                    </div>
                  </div>
                ));
              })()}
            </Card>

            <Card style={{padding:'14px 16px'}}>
              <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.08em',
                textTransform:'uppercase',color:C.inkMute,marginBottom:'10px'}}>Vehicle vs Equipment Population</div>
              {(() => {
                const groups = {Vehicle:{count:0,pop:0},Equipment:{count:0,pop:0}};
                filtered.forEach(r=>{
                  const k = r.recallType==='Equipment'?'Equipment':'Vehicle';
                  groups[k].count++; groups[k].pop += r.rawListedPotentialPopulation||0;
                });
                const max = Math.max(1, groups.Vehicle.pop, groups.Equipment.pop);
                return Object.entries(groups).map(([k,v])=>(
                  <div key={k} style={{marginBottom:'8px'}}>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:'12px',marginBottom:'3px'}}>
                      <span style={{color:C.ink}}>{k} <span style={{color:C.inkMute}}>({v.count})</span></span>
                      <span style={{fontVariantNumeric:'tabular-nums',color:C.inkSub}}>{riFmtInt(v.pop)}</span>
                    </div>
                    <div style={{height:'6px',background:C.tintNeutral,borderRadius:'3px',overflow:'hidden'}}>
                      <div style={{height:'100%',width:`${(v.pop/max)*100}%`,background:C.navy}}/>
                    </div>
                  </div>
                ));
              })()}
            </Card>
          </div>

          <Card style={{padding:'14px 16px'}}>
            <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.08em',
              textTransform:'uppercase',color:C.inkMute,marginBottom:'10px'}}>Completion Reporting Coverage</div>
            <div style={{fontSize:'12.5px',color:C.ink,lineHeight:1.7}}>
              Reported: <b>{filtered.length - notReportedCount}</b> of <b>{filtered.length}</b> filtered records
              {' · '}Not Reported: <b>{notReportedCount}</b>
              {' · '}Literal 0%: <b>{filtered.filter(r=>r.completionRatePct===0).length}</b>
            </div>
          </Card>

          <Card style={{padding:'14px 16px'}}>
            <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.08em',
              textTransform:'uppercase',color:C.inkMute,marginBottom:'10px'}}>Largest Recalls by Potentially Affected</div>
            <table style={{width:'100%',fontSize:'12px',borderCollapse:'collapse'}}>
              <thead><tr style={{color:C.inkSub}}>
                <th style={{textAlign:'left',padding:'4px 6px'}}>NHTSA ID</th>
                <th style={{textAlign:'left',padding:'4px 6px'}}>Subject</th>
                <th style={{textAlign:'right',padding:'4px 6px'}}>Potentially Affected</th>
                <th style={{textAlign:'right',padding:'4px 6px'}}>Completion</th>
              </tr></thead>
              <tbody>
                {[...filtered].sort((a,b)=>(b.rawListedPotentialPopulation||0)-(a.rawListedPotentialPopulation||0))
                  .slice(0,5).map(r=>(
                    <tr key={r._id} style={{borderTop:`1px solid ${C.borderSoft}`}}>
                      <td style={{padding:'6px',color:C.coral,fontWeight:600}}>{r.nhtsaId}</td>
                      <td style={{padding:'6px',color:C.ink}}>{r.subject}</td>
                      <td style={{padding:'6px',textAlign:'right',fontVariantNumeric:'tabular-nums'}}>{riFmtInt(r.rawListedPotentialPopulation)}</td>
                      <td style={{padding:'6px',textAlign:'right',fontVariantNumeric:'tabular-nums',
                        color:r.completionRatePct===null?C.inkFaint:C.ink}}>{riFmtPct(r.completionRatePct)}</td>
                    </tr>
                ))}
              </tbody>
            </table>
          </Card>
        </>
      )}

      {tab==='exposure' && (
        <Card style={{padding:'28px 24px'}}>
          <div style={{display:'flex',alignItems:'center',gap:'10px',marginBottom:'14px'}}>
            <Pill color={C.amber}>Planned — Not yet implemented</Pill>
            <span style={{fontSize:'11.5px',color:C.inkMute}}>Editable JSS exposure assumptions deferred to a later increment</span>
          </div>
          <div style={{fontSize:'11px',fontWeight:600,letterSpacing:'0.08em',
            textTransform:'uppercase',color:C.inkMute,marginBottom:'8px'}}>Future contents</div>
          <ul style={{margin:0,paddingLeft:'18px',fontSize:'13px',color:C.ink,lineHeight:1.7}}>
            <li>JSS exposure assumptions (per recall / per campaign)</li>
            <li>Projected JSS exposure calculations</li>
            <li>Source-versus-assumption tagging on every value</li>
            <li>Optional price-per-module assumptions</li>
            <li>Actual-versus-projected comparison view</li>
          </ul>
          <div style={{marginTop:'18px',padding:'12px 14px',background:C.panelAlt,
            border:`1px dashed ${C.border}`,borderRadius:S.r.md,fontSize:'11.5px',color:C.inkMute,lineHeight:1.55}}>
            This view is a planned empty state. No editable inputs and no calculations are active.
          </div>
        </Card>
      )}

      {selected && <RecallDetail rec={selected} onClose={()=>setSelected(null)}/>}
    </div>
  );
}


// ════════════════════════════════════════════════════════════════════
// GM WORKSPACE — Increment 3A (read-only)
// ════════════════════════════════════════════════════════════════════
// SYNTHETIC BETA FIXTURE DATA ONLY.
// All Bill and Claim records in this block are fictional and were created
// solely for UI, calculation, filtering, and workflow validation.
// Do not replace these fixtures with production data in frontend source.
// Production and confidential records must be retrieved only through an
// approved authenticated backend.
// This block is self-contained. It does not read/write localStorage.
// STORAGE_KEYS.gm remains declared but inactive.
// ════════════════════════════════════════════════════════════════════

const GM_FIELD_KEYS = ["txnNum", "txnId", "billNum", "stmtDate", "billApprovalDate", "billStatus", "lbeStatus", "lbe", "recoveryGroup", "vin", "modelYear", "make", "model", "platform", "bodyStyle", "plant", "buildCountry", "buildRegion", "buildDate", "vehStatusDesc", "repairDate", "procDate", "jobCard", "odoMiles", "odoKm", "mis", "claimExpenseCat", "customerVerbatim", "causalVerbatim", "causalPartVerbatim", "correctionVerbatim", "generalVerbatim", "complaintCd", "complaintDesc", "bomId", "bomDesc", "actionDesc", "glc", "glcDesc", "baseLaborHrs", "otherLaborHrs", "supLaborHrs", "diagLaborHrs", "causalSvcPN", "causalSvcPNDesc", "causalPartQty", "causalProdPN", "causalProdPNDesc", "componentPN", "nonCausalSvcPN", "nonCausalSvcPNDesc", "nonCausalPartQty", "dealerName", "dealerCode", "dealerCity", "dealerState", "dealerCountry", "dealerRegion", "repairSubRegion", "dealerCurrency", "laborCostUSD", "baseLaborCostUSD", "otherLaborCostUSD", "supLaborCostUSD", "totalPartCostUSD", "towingCostUSD", "altTransCostUSD", "miscCostUSD", "deductibleCostUSD", "taxCostUSD", "totalCostUSD", "billingExchRate", "billCurrencyType", "owtClaimTotal", "adjustment", "adjustmentReason", "owtAdjClaimTotal", "shareableAmount", "supAlloc50", "techFactorPct", "supplierDisputeComment", "mfgSupplierName", "mfgSupplierDUNS", "remitToDuns", "techFactorApprovalDate", "projectId", "projectOwner", "reqOrg", "partReturnStatus", "partReceivedDate", "facility", "inspDisposition", "inspDispositionDate", "inspectedBy", "inspComment", "invStatus", "storageLoc", "storedDate", "pickReqNbr"];
const GM_FIELD_LABELS = {"txnNum": "Transaction Number", "txnId": "Transaction Id", "billNum": "Billing Number", "stmtDate": "Statement Date", "billApprovalDate": "Bill Approval Date", "billStatus": "Bill Status", "lbeStatus": "LBE Status", "lbe": "Legal Business Entity", "recoveryGroup": "Recovery Group", "vin": "VIN", "modelYear": "Model Year", "make": "Make", "model": "Model", "platform": "Platform", "bodyStyle": "Body Style", "plant": "Plant", "buildCountry": "Build Country", "buildRegion": "Build Region", "buildDate": "Build Date", "vehStatusDesc": "Vehicle Status Cd Desc", "repairDate": "Repair Date", "procDate": "Proc Date", "jobCard": "Job Card", "odoMiles": "Vehicle Odometer - Miles", "odoKm": "Vehicle Odometer - Kilometer", "mis": "Months in Service", "claimExpenseCat": "Claim Expense Category", "customerVerbatim": "Customer Verbatim", "causalVerbatim": "Causal Verbatim", "causalPartVerbatim": "Causal Part Verbatim", "correctionVerbatim": "Correction Verbatim", "generalVerbatim": "General Verbatim", "complaintCd": "Complaint CD", "complaintDesc": "Complaint Cd Desc", "bomId": "BOM ID", "bomDesc": "BOM Desc", "actionDesc": "Action Description", "glc": "Global Labor Code", "glcDesc": "Global Labor Code Desc", "baseLaborHrs": "Base Labor Hours", "otherLaborHrs": "Other Labor Hours", "supLaborHrs": "Supplemental Labor Hours", "diagLaborHrs": "Diagnostic Labor Hours", "causalSvcPN": "Causal Service PN", "causalSvcPNDesc": "Causal Service PN Description", "causalPartQty": "Causal Part Qty", "causalProdPN": "Causal Production PN", "causalProdPNDesc": "Causal Production PN Description", "componentPN": "Component Part Number", "nonCausalSvcPN": "Non-Causal Service PN", "nonCausalSvcPNDesc": "Non-Causal Service PN Description", "nonCausalPartQty": "Non-Causal Part Qty", "dealerName": "Repair Dealer Name", "dealerCode": "Repairing Dealer Code", "dealerCity": "Repair Dealer City", "dealerState": "Repair Dealer State", "dealerCountry": "Repair Dealer Country", "dealerRegion": "Repair Dealer Region", "repairSubRegion": "Repair Sub Region", "dealerCurrency": "Dealer Submitted Currency", "laborCostUSD": "Labor Cost (USD)", "baseLaborCostUSD": "Base Labor Cost (USD)", "otherLaborCostUSD": "Other Labor Cost (USD)", "supLaborCostUSD": "Supplemental Labor Cost (USD)", "totalPartCostUSD": "Total Part Cost (USD)", "towingCostUSD": "Towing Cost (USD)", "altTransCostUSD": "Alt Transportation Cost (USD)", "miscCostUSD": "Misc. Cost (USD)", "deductibleCostUSD": "Deductible Cost (USD)", "taxCostUSD": "Tax Cost (USD)", "totalCostUSD": "Total Cost (USD)", "billingExchRate": "Billing Exchange Rate (USD to Contract Currency)", "billCurrencyType": "Bill Currency Type", "owtClaimTotal": "OWT Claim Total", "adjustment": "Adjustment", "adjustmentReason": "Adjustment Reason", "owtAdjClaimTotal": "OWT Adjusted Claim Total", "shareableAmount": "Shareable Amount", "supAlloc50": "50% Supplier Allocation Amount", "techFactorPct": "Technical Factor %", "supplierDisputeComment": "Supplier Dispute Comment", "mfgSupplierName": "Mfg Supplier Name", "mfgSupplierDUNS": "Mfg Supplier DUNS", "remitToDuns": "Remit to Duns", "techFactorApprovalDate": "Technical Factor Approval Date", "projectId": "Project Id", "projectOwner": "Project Owner", "reqOrg": "Requesting Organization", "partReturnStatus": "Part Return Status", "partReceivedDate": "Part Received Date From Dealer", "facility": "Facility", "inspDisposition": "Claim Inspection Disposition", "inspDispositionDate": "Inspection Disposition Date", "inspectedBy": "Inspected By", "inspComment": "Inspection Comment", "invStatus": "Inventory Status", "storageLoc": "Storage Location", "storedDate": "Stored Date", "pickReqNbr": "Pick Request Nbr"};
const GM_SECTIONS = [["ib", "Identity & Billing"], ["veh", "Vehicle"], ["cr", "Claim & Repair"], ["pt", "Parts"], ["dr", "Dealer & Region"], ["fin", "Labor & Financial"], ["sa", "Supplier & Allocation"], ["si", "Inspection & Return"]];
const GM_SECTION_KEYS = {"ib": ["txnNum", "txnId", "billNum", "stmtDate", "billApprovalDate", "billStatus", "lbeStatus", "lbe", "recoveryGroup"], "veh": ["vin", "modelYear", "make", "model", "platform", "bodyStyle", "plant", "buildCountry", "buildRegion", "buildDate", "vehStatusDesc", "repairDate", "procDate", "jobCard", "odoMiles", "odoKm", "mis"], "cr": ["claimExpenseCat", "customerVerbatim", "causalVerbatim", "causalPartVerbatim", "correctionVerbatim", "generalVerbatim", "complaintCd", "complaintDesc", "bomId", "bomDesc", "actionDesc", "glc", "glcDesc", "baseLaborHrs", "otherLaborHrs", "supLaborHrs", "diagLaborHrs"], "pt": ["causalSvcPN", "causalSvcPNDesc", "causalPartQty", "causalProdPN", "causalProdPNDesc", "componentPN", "nonCausalSvcPN", "nonCausalSvcPNDesc", "nonCausalPartQty"], "dr": ["dealerName", "dealerCode", "dealerCity", "dealerState", "dealerCountry", "dealerRegion", "repairSubRegion"], "fin": ["dealerCurrency", "laborCostUSD", "baseLaborCostUSD", "otherLaborCostUSD", "supLaborCostUSD", "totalPartCostUSD", "towingCostUSD", "altTransCostUSD", "miscCostUSD", "deductibleCostUSD", "taxCostUSD", "totalCostUSD", "billingExchRate", "billCurrencyType"], "sa": ["owtClaimTotal", "adjustment", "adjustmentReason", "owtAdjClaimTotal", "shareableAmount", "supAlloc50", "techFactorPct"], "si": ["supplierDisputeComment", "mfgSupplierName", "mfgSupplierDUNS", "remitToDuns", "techFactorApprovalDate", "projectId", "projectOwner", "reqOrg", "partReturnStatus", "partReceivedDate", "facility", "inspDisposition", "inspDispositionDate", "inspectedBy", "inspComment", "invStatus", "storageLoc", "storedDate", "pickReqNbr"]};
const GM_NARRATIVE_KEYS = ["customerVerbatim", "causalVerbatim", "causalPartVerbatim", "correctionVerbatim", "generalVerbatim", "supplierDisputeComment", "inspComment", "adjustmentReason"];

const GM_BILLS_SEED = [{"billNum": "BETA-BILL-001", "stmtDate": "2021-06-01", "billStatus": "Approved", "lbeStatus": "Manual Submit Required", "billApprovalDate": "2021-06-11", "claimCount": 3}, {"billNum": "BETA-BILL-002", "stmtDate": "2021-07-04", "billStatus": "Pending Approval", "lbeStatus": "Pending Approval", "billApprovalDate": null, "claimCount": 3}, {"billNum": "BETA-BILL-003", "stmtDate": "2021-08-09", "billStatus": "Approved", "lbeStatus": "Pending Approval", "billApprovalDate": "2021-08-19", "claimCount": 3}];

// Positional claim rows aligned to GM_FIELD_KEYS.
const GM_CLAIMS_SEED_ROWS = [["BETA-TXN-001", "BETA-TXN-001", "BETA-BILL-001", "2021-06-01", "2021-06-11", "Approved", "Manual Submit Required", "Beta Test Dealer", "Synthetic Recovery Group", "BETA001TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-001", 10500, 16800, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-001", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-002", "BETA-TXN-002", "BETA-BILL-001", "2021-06-01", "2021-06-11", "Approved", "Manual Submit Required", "Beta Test Dealer", "Synthetic Recovery Group", "BETA002TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-002", 11000, 17600, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-002", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-003", "BETA-TXN-003", "BETA-BILL-001", "2021-06-01", "2021-06-11", "Approved", "Manual Submit Required", "Beta Test Dealer", "Synthetic Recovery Group", "BETA003TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-003", 11500, 18400, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-003", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, "USD", 0.0, 0.0, null, 0.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-004", "BETA-TXN-004", "BETA-BILL-002", "2021-07-04", null, "Pending Approval", "Pending Approval", "Beta Test Dealer", "Synthetic Recovery Group", "BETA004TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-004", 12000, 19200, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-004", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-005", "BETA-TXN-005", "BETA-BILL-002", "2021-07-04", null, "Pending Approval", "Pending Approval", "Beta Test Dealer", "Synthetic Recovery Group", "BETA005TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-005", 12500, 20000, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-005", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-006", "BETA-TXN-006", "BETA-BILL-002", "2021-07-04", null, "Pending Approval", "Pending Approval", "Beta Test Dealer", "Synthetic Recovery Group", "BETA006TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-006", 13000, 20800, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-006", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 25.0, "Synthetic goodwill adjustment for demonstration", 180.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-007", "BETA-TXN-007", "BETA-BILL-003", "2021-08-09", "2021-08-19", "Approved", "Pending Approval", "Beta Test Dealer", "Synthetic Recovery Group", "BETA007TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-007", 13500, 21600, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-007", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-008", "BETA-TXN-008", "BETA-BILL-003", "2021-08-09", "2021-08-19", "Approved", "Pending Approval", "Beta Test Dealer", "Synthetic Recovery Group", "BETA008TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-008", 14000, 22400, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-008", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 20.0, 10.0, 50.0, 15.0, 0.0, 0.0, 0.0, 5.0, 200.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null], ["BETA-TXN-009", "BETA-TXN-009", "BETA-BILL-003", "2021-08-09", "2021-08-19", "Approved", "Pending Approval", "Beta Test Dealer", "Synthetic Recovery Group", "BETA009TESTIQOXXX", 2025, "SyntheticMake", "SyntheticModel", "SyntheticPlatform", "Sedan", "Synthetic Assembly Plant", "Testland", "Test Region", "2021-05-01", "In Service", "2021-06-06", "2021-06-07", "BETA-JOB-009", null, null, "N", "Warranty", "Fictional beta-test customer description for UI validation purposes.", "Synthetic causal description used solely for demonstration.", "Synthetic causal part description.", "Synthetic correction narrative describing a fictional repair action.", "General synthetic beta-test remark.", "SYN-01", "Synthetic complaint description", "SYN-BOM-01", "Synthetic BOM description", "Synthetic corrective action", "SYN-GLC", "Synthetic GLC description", 1.5, 0.0, 0.0, 0.5, "SYN-PN-0001", "Synthetic part description", 1, "SYN-PPN-0001", "Synthetic production part description", "SYN-COMP-0001", null, null, 0, "Beta Test Dealer", "SYN-009", "Testville", "TS", "Testland", "Synthetic Region", "Synthetic Sub-Region", "USD", 100.0, 75.0, 0.0, 0.0, 50.0, 0.0, 0.0, 0.0, 0.0, 5.0, 155.0, 1.0, "USD", 155.0, 0.0, null, 155.0, 0.0, 0.0, 0.0, null, "Synthetic Supplier Alpha", "000000000", "000000000", null, "SYN-PROJ-01", "Demonstration Repair Center", "Test Business Unit", "Not Required", null, "Demonstration Repair Center", "Not Inspected", null, null, null, "N/A", null, null, null]];
function gmClaimObj(row) {
  const o = {};
  for (let i=0;i<GM_FIELD_KEYS.length;i++) o[GM_FIELD_KEYS[i]] = row[i];
  return o;
}
const GM_CLAIMS = GM_CLAIMS_SEED_ROWS.map(gmClaimObj);

function gmFmtMoney(n) {
  if (n===null || n===undefined || n==='' || Number.isNaN(Number(n))) return '—';
  const v = Number(n);
  return v.toLocaleString('en-US',{style:'currency',currency:'USD',maximumFractionDigits:2});
}
function gmFmtNum(n) {
  if (n===null || n===undefined || n==='' || Number.isNaN(Number(n))) return '—';
  return Number(n).toLocaleString('en-US');
}
function gmFmtPct(n) {
  if (n===null || n===undefined || n==='') return '—';
  const v = Number(n);
  if (Number.isNaN(v)) return String(n);
  return (v<=1 ? (v*100) : v).toFixed(2)+'%';
}
function gmVal(v) {
  if (v===null || v===undefined || v==='') return '—';
  return String(v);
}
function gmSum(arr, k) {
  let s = 0;
  for (const c of arr) { const v = Number(c[k]); if (!Number.isNaN(v)) s += v; }
  return s;
}
function gmUnique(arr, k) {
  const s = new Set();
  for (const c of arr) if (c[k]!==null && c[k]!==undefined && c[k]!=='') s.add(c[k]);
  return s.size;
}
function gmGroupCount(arr, k) {
  const m = new Map();
  for (const c of arr) {
    const v = (c[k]===null||c[k]===undefined||c[k]==='') ? '—' : String(c[k]);
    m.set(v,(m.get(v)||0)+1);
  }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]);
}

function GM_KPI({label, value, hint}) {
  return (
    <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,padding:'12px 14px'}}>
      <div style={{fontSize:11,textTransform:'uppercase',letterSpacing:.4,color:'#7C7669'}}>{label}</div>
      <div style={{fontSize:22,fontWeight:600,color:'#1B2A44',marginTop:4}}>{value}</div>
      {hint && <div style={{fontSize:11,color:'#7C7669',marginTop:2}}>{hint}</div>}
    </div>
  );
}

function GMClaimDetail({claim, onBack}) {
  if (!claim) return null;
  return (
    <div style={{background:'#FBF9F4',border:'1px solid #E7E4DC',borderRadius:8,padding:16}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
        <div>
          <div style={{fontSize:12,color:'#7C7669'}}>Transaction Id</div>
          <div style={{fontSize:20,fontWeight:600,color:'#1B2A44'}}>{gmVal(claim.txnId)}</div>
          <div style={{fontSize:12,color:'#7C7669',marginTop:4}}>
            Bill {gmVal(claim.billNum)} · VIN {gmVal(claim.vin)} · {gmVal(claim.make)} {gmVal(claim.model)} {gmVal(claim.modelYear)}
          </div>
        </div>
        <Btn onClick={onBack}>← Back to Claims</Btn>
      </div>
      {GM_SECTIONS.map(([sid,sLabel]) => (
        <div key={sid} style={{marginBottom:14,background:'#fff',border:'1px solid #E7E4DC',borderRadius:6,padding:12}}>
          <div style={{fontSize:12,fontWeight:600,textTransform:'uppercase',letterSpacing:.5,color:'#C0574E',marginBottom:8}}>{sLabel}</div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(240px,1fr))',gap:'8px 16px'}}>
            {GM_SECTION_KEYS[sid].map(k => {
              const isNarr = GM_NARRATIVE_KEYS.indexOf(k)>=0;
              if (isNarr) return null;
              return (
                <div key={k}>
                  <div style={{fontSize:10,textTransform:'uppercase',color:'#7C7669',letterSpacing:.3}}>{GM_FIELD_LABELS[k]}</div>
                  <div style={{fontSize:13,color:'#1B2A44',wordBreak:'break-word'}}>{gmVal(claim[k])}</div>
                </div>
              );
            })}
          </div>
          {GM_SECTION_KEYS[sid].filter(k=>GM_NARRATIVE_KEYS.indexOf(k)>=0).map(k => (
            <div key={k} style={{marginTop:10,borderTop:'1px dashed #E7E4DC',paddingTop:8}}>
              <div style={{fontSize:10,textTransform:'uppercase',color:'#7C7669',letterSpacing:.3}}>{GM_FIELD_LABELS[k]}</div>
              <div style={{fontSize:13,color:'#1B2A44',whiteSpace:'pre-wrap'}}>{claim[k] ? String(claim[k]) : 'Not provided'}</div>
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

function GMWorkspace() {
  const [tab,setTab] = React.useState('dashboard');
  const [selectedBill,setSelectedBill] = React.useState('');
  const [selectedClaimId,setSelectedClaimId] = React.useState(null);
  const [search,setSearch] = React.useState('');
  const [fBill,setFBill] = React.useState('');
  const [fBillStatus,setFBillStatus] = React.useState('');
  const [fLbeStatus,setFLbeStatus] = React.useState('');
  const [fMake,setFMake] = React.useState('');
  const [fPlatform,setFPlatform] = React.useState('');
  const [fCurrency,setFCurrency] = React.useState('');
  const [page,setPage] = React.useState(1);
  const PAGE_SIZE = 50;

  const claims = GM_CLAIMS;

  const filteredClaims = React.useMemo(()=>{
    const s = search.trim().toLowerCase();
    return claims.filter(c => {
      if (fBill && String(c.billNum)!==fBill) return false;
      if (selectedBill && String(c.billNum)!==selectedBill) return false;
      if (fBillStatus && c.billStatus!==fBillStatus) return false;
      if (fLbeStatus && c.lbeStatus!==fLbeStatus) return false;
      if (fMake && c.make!==fMake) return false;
      if (fPlatform && c.platform!==fPlatform) return false;
      if (fCurrency && c.dealerCurrency!==fCurrency) return false;
      if (s) {
        const hay = [c.txnId,c.txnNum,c.vin,c.billNum,c.make,c.model,c.causalSvcPN,c.causalSvcPNDesc,c.bomDesc,c.dealerName,c.customerVerbatim,c.causalVerbatim,c.correctionVerbatim].map(x=>x==null?'':String(x).toLowerCase()).join(' | ');
        if (hay.indexOf(s)<0) return false;
      }
      return true;
    });
  },[claims,search,fBill,fBillStatus,fLbeStatus,fMake,fPlatform,fCurrency,selectedBill]);

  React.useEffect(()=>{ setPage(1); },[search,fBill,fBillStatus,fLbeStatus,fMake,fPlatform,fCurrency,selectedBill,tab]);

  const uniqueVals = (k) => {
    const s = new Set();
    for (const c of claims) if (c[k]!==null && c[k]!==undefined && c[k]!=='') s.add(String(c[k]));
    return [...s].sort();
  };
  const makes = uniqueVals('make');
  const platforms = uniqueVals('platform');
  const currencies = uniqueVals('dealerCurrency');
  const billStatuses = uniqueVals('billStatus');
  const lbeStatuses = uniqueVals('lbeStatus');
  const billNumbers = GM_BILLS_SEED.map(b=>b.billNum);

  // Dashboard KPIs
  const totalClaims = claims.length;
  const totalBills = GM_BILLS_SEED.length;
  const totalVINs = gmUnique(claims,'vin');
  const owtAdj = gmSum(claims,'owtAdjClaimTotal');
  const shareable = gmSum(claims,'shareableAmount');
  const alloc50 = gmSum(claims,'supAlloc50');
  const approvedBills = GM_BILLS_SEED.filter(b=>b.billStatus==='Approved').length;
  const pendingBills = GM_BILLS_SEED.filter(b=>b.billStatus==='Pending Approval').length;

  const claimsByMake = gmGroupCount(claims,'make').slice(0,10);
  const claimsByPlatform = gmGroupCount(claims,'platform').slice(0,10);
  const claimsByBom = gmGroupCount(claims,'bomDesc').slice(0,10);
  const claimsByGLC = gmGroupCount(claims,'glcDesc').slice(0,10);
  const billStatusDist = gmGroupCount(claims,'billStatus');
  const lbeStatusDist = gmGroupCount(claims,'lbeStatus');

  const selectedClaim = selectedClaimId ? claims.find(c=>String(c.txnId)===String(selectedClaimId)) : null;

  const clearFilters = () => { setSearch(''); setFBill(''); setFBillStatus(''); setFLbeStatus(''); setFMake(''); setFPlatform(''); setFCurrency(''); };

  const totalPages = Math.max(1, Math.ceil(filteredClaims.length/PAGE_SIZE));
  const pageClaims = filteredClaims.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  const tabBtn = (id,label) => (
    <button key={id} onClick={()=>{setTab(id); setSelectedClaimId(null);}}
      style={{padding:'8px 14px',border:'1px solid '+(tab===id?'#C0574E':'#E7E4DC'),background:tab===id?'#C0574E':'#fff',color:tab===id?'#fff':'#1B2A44',borderRadius:6,fontSize:13,cursor:'pointer',marginRight:6,fontWeight:tab===id?600:400}}>
      {label}
    </button>
  );

  return (
    <div style={{padding:16}}>
      <div style={{marginBottom:12}}>
        <div style={{fontSize:20,fontWeight:600,color:'#1B2A44'}}>GM Workspace</div>
        <div style={{fontSize:12,color:'#7C7669'}}>Read-only beta prototype · {totalBills} synthetic bills · {totalClaims} synthetic claims · fixture data only</div>
      </div>
      <div style={{marginBottom:14}}>
        {tabBtn('dashboard','GM Dashboard')}
        {tabBtn('bills','Warranty Bills')}
        {tabBtn('claims','Claims')}
        {tabBtn('approval','Approval & Allocation')}
      </div>

      {selectedClaim ? (
        <GMClaimDetail claim={selectedClaim} onBack={()=>setSelectedClaimId(null)} />
      ) : tab==='dashboard' ? (
        <div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(180px,1fr))',gap:10,marginBottom:16}}>
            <GM_KPI label="Bills" value={gmFmtNum(totalBills)} hint={approvedBills+' Approved · '+pendingBills+' Pending'} />
            <GM_KPI label="Claims" value={gmFmtNum(totalClaims)} />
            <GM_KPI label="Unique VINs" value={gmFmtNum(totalVINs)} hint={(totalClaims-totalVINs)+' repeat VINs'} />
            <GM_KPI label="OWT Adjusted Claim Total" value={gmFmtMoney(owtAdj)} />
            <GM_KPI label="Shareable Amount" value={gmFmtMoney(shareable)} />
            <GM_KPI label="50% Supplier Allocation" value={gmFmtMoney(alloc50)} />
          </div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
            <GMBreakdown title="Claims by Make" data={claimsByMake} />
            <GMBreakdown title="Claims by Platform" data={claimsByPlatform} />
            <GMBreakdown title="Claims by BOM Description" data={claimsByBom} />
            <GMBreakdown title="Claims by Global Labor Code Desc" data={claimsByGLC} />
            <GMBreakdown title="Bill Status Distribution" data={billStatusDist} />
            <GMBreakdown title="LBE Status Distribution" data={lbeStatusDist} />
          </div>
        </div>
      ) : tab==='bills' ? (
        <div>
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:13,background:'#fff',border:'1px solid #E7E4DC'}}>
            <thead style={{background:'#F5F2EA'}}>
              <tr>
                {['Billing Number','Statement Date','Claim Count','Bill Status','LBE Status','OWT Adjusted Claim Total','Shareable Amount','50% Supplier Allocation','Bill Approval Date',''].map(h=>(
                  <th key={h} style={{textAlign:'left',padding:'8px 10px',borderBottom:'1px solid #E7E4DC',color:'#1B2A44',fontWeight:600}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {GM_BILLS_SEED.map(b => {
                const bc = claims.filter(c=>String(c.billNum)===b.billNum);
                return (
                  <tr key={b.billNum}>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4',fontWeight:600}}>{b.billNum}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.stmtDate)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtNum(b.claimCount)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.billStatus)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.lbeStatus)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'owtAdjClaimTotal'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'shareableAmount'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'supAlloc50'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.billApprovalDate)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>
                      <Btn onClick={()=>{setSelectedBill(b.billNum); setTab('claims');}}>View Claims</Btn>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : tab==='claims' ? (
        <div>
          <div style={{display:'flex',flexWrap:'wrap',gap:8,marginBottom:10,alignItems:'center'}}>
            <Inp placeholder="Search Transaction Id, VIN, PN, dealer, verbatim…" value={search} onChange={e=>setSearch(e.target.value)} style={{minWidth:280}} />
            <Sel value={selectedBill||fBill} onChange={e=>{setFBill(e.target.value); setSelectedBill('');}}>
              <option value="">All Bills</option>
              {billNumbers.map(b=><option key={b} value={b}>{b}</option>)}
            </Sel>
            <Sel value={fBillStatus} onChange={e=>setFBillStatus(e.target.value)}>
              <option value="">All Bill Statuses</option>
              {billStatuses.map(b=><option key={b} value={b}>{b}</option>)}
            </Sel>
            <Sel value={fLbeStatus} onChange={e=>setFLbeStatus(e.target.value)}>
              <option value="">All LBE Statuses</option>
              {lbeStatuses.map(b=><option key={b} value={b}>{b}</option>)}
            </Sel>
            <Sel value={fMake} onChange={e=>setFMake(e.target.value)}>
              <option value="">All Makes</option>
              {makes.map(b=><option key={b} value={b}>{b}</option>)}
            </Sel>
            <Sel value={fPlatform} onChange={e=>setFPlatform(e.target.value)}>
              <option value="">All Platforms</option>
              {platforms.map(b=><option key={b} value={b}>{b}</option>)}
            </Sel>
            <Sel value={fCurrency} onChange={e=>setFCurrency(e.target.value)}>
              <option value="">All Currencies</option>
              {currencies.map(b=><option key={b} value={b}>{b}</option>)}
            </Sel>
            <Btn onClick={()=>{clearFilters(); setSelectedBill('');}}>Clear Filters</Btn>
            <div style={{fontSize:12,color:'#7C7669',marginLeft:'auto'}}>{filteredClaims.length} of {claims.length} claims</div>
          </div>
          <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:6,overflow:'auto',maxHeight:'60vh'}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
              <thead style={{background:'#F5F2EA',position:'sticky',top:0}}>
                <tr>
                  {['Transaction Id','Bill','Repair Date','VIN','MY','Make','Model','Platform','BOM Desc','Causal Svc PN','Dealer','Curr.','Total (USD)','OWT Adj (USD)','Bill Status','LBE Status',''].map(h=>(
                    <th key={h} style={{textAlign:'left',padding:'6px 8px',borderBottom:'1px solid #E7E4DC',fontWeight:600,color:'#1B2A44',whiteSpace:'nowrap'}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pageClaims.map(c => (
                  <tr key={c.txnId}>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',fontFamily:'monospace'}}>{gmVal(c.txnId)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.billNum)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',whiteSpace:'nowrap'}}>{gmVal(c.repairDate)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',fontFamily:'monospace'}}>{gmVal(c.vin)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.modelYear)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.make)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.model)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.platform)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',maxWidth:200,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{gmVal(c.bomDesc)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',fontFamily:'monospace'}}>{gmVal(c.causalSvcPN)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',maxWidth:160,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{gmVal(c.dealerName)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.dealerCurrency)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',textAlign:'right'}}>{gmFmtMoney(c.totalCostUSD)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA',textAlign:'right'}}>{gmFmtMoney(c.owtAdjClaimTotal)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.billStatus)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>{gmVal(c.lbeStatus)}</td>
                    <td style={{padding:'6px 8px',borderBottom:'1px solid #F5F2EA'}}>
                      <Btn onClick={()=>setSelectedClaimId(c.txnId)}>Open</Btn>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{display:'flex',alignItems:'center',gap:8,marginTop:8,fontSize:12,color:'#7C7669'}}>
            <Btn onClick={()=>setPage(p=>Math.max(1,p-1))}>‹ Prev</Btn>
            <span>Page {page} of {totalPages}</span>
            <Btn onClick={()=>setPage(p=>Math.min(totalPages,p+1))}>Next ›</Btn>
          </div>
        </div>
      ) : tab==='approval' ? (
        <div>
          <div style={{fontSize:12,color:'#7C7669',marginBottom:10}}>Read-only summary of source-provided approval and allocation fields. No approval, rejection, allocation, or submit actions are available in this prototype.</div>
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:13,background:'#fff',border:'1px solid #E7E4DC'}}>
            <thead style={{background:'#F5F2EA'}}>
              <tr>
                {['Billing Number','Bill Status','LBE Status','Bill Approval Date','Tech Factor Approval Date (any)','Technical Factor %','OWT Claim Total','Adjustment','OWT Adj. Claim Total','Shareable','50% Supplier Allocation','Adjustment Reason (distinct)'].map(h=>(
                  <th key={h} style={{textAlign:'left',padding:'8px 10px',borderBottom:'1px solid #E7E4DC',fontWeight:600,color:'#1B2A44'}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {GM_BILLS_SEED.map(b => {
                const bc = claims.filter(c=>String(c.billNum)===b.billNum);
                const tfSet = new Set(bc.map(c=>c.techFactorPct).filter(v=>v!==null&&v!==undefined&&v!==''));
                const tfApprovals = new Set(bc.map(c=>c.techFactorApprovalDate).filter(v=>v!==null&&v!==undefined&&v!==''));
                const adjReasons = new Set(bc.map(c=>c.adjustmentReason).filter(v=>v!==null&&v!==undefined&&v!==''));
                return (
                  <tr key={b.billNum}>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4',fontWeight:600}}>{b.billNum}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.billStatus)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.lbeStatus)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmVal(b.billApprovalDate)}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{tfApprovals.size ? [...tfApprovals].join(', ') : '—'}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{tfSet.size ? [...tfSet].map(v=>gmFmtPct(v)).join(', ') : '—'}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'owtClaimTotal'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'adjustment'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'owtAdjClaimTotal'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'shareableAmount'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4'}}>{gmFmtMoney(gmSum(bc,'supAlloc50'))}</td>
                    <td style={{padding:'8px 10px',borderBottom:'1px solid #F0EDE4',fontSize:11,color:'#7C7669'}}>{adjReasons.size ? [...adjReasons].join(' | ') : '—'}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : null}
    </div>
  );
}

function GMBreakdown({title, data}) {
  const max = data.reduce((m,[,c])=>Math.max(m,c),0) || 1;
  return (
    <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:6,padding:12}}>
      <div style={{fontSize:12,fontWeight:600,color:'#1B2A44',marginBottom:8,textTransform:'uppercase',letterSpacing:.4}}>{title}</div>
      {data.length===0 ? <div style={{fontSize:12,color:'#7C7669'}}>—</div> : data.map(([k,v]) => (
        <div key={k} style={{display:'grid',gridTemplateColumns:'160px 1fr 40px',gap:8,alignItems:'center',fontSize:12,padding:'2px 0'}}>
          <div style={{color:'#1B2A44',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={k}>{k}</div>
          <div style={{background:'#F5F2EA',height:8,borderRadius:4,overflow:'hidden'}}><div style={{background:'#C0574E',height:'100%',width:(v/max*100)+'%'}} /></div>
          <div style={{textAlign:'right',color:'#7C7669'}}>{v}</div>
        </div>
      ))}
    </div>
  );
}


const STLA_CAMPAIGNS = [
  {
    "campaignNumber": "4551",
    "customer": "Stellantis",
    "plant": "Demo Assembly Plant",
    "modelYear": 2018,
    "claimClassification": "Synthetic inflator recovery demonstration",
    "billingCycle": "2026-06",
    "currency": "USD",
    "status": "Recovery Review",
    "dataClassification": "SYNTHETIC"
  },
  {
    "campaignNumber": "4750",
    "customer": "Stellantis",
    "plant": "Demo Assembly Plant",
    "modelYear": 2019,
    "claimClassification": "Synthetic inflator recovery demonstration",
    "billingCycle": "2026-06",
    "currency": "USD",
    "status": "Allocation Review",
    "dataClassification": "SYNTHETIC"
  }
];

const STLA_SWRS_GROUPS = [
  {
    "id": "SWRS-DEMO-4551-A",
    "campaignNumber": "4551",
    "label": "SWRS Group A",
    "reviewStatus": "Reconciled"
  },
  {
    "id": "SWRS-DEMO-4551-B",
    "campaignNumber": "4551",
    "label": "SWRS Group B",
    "reviewStatus": "Review Required"
  },
  {
    "id": "SWRS-DEMO-4551-C",
    "campaignNumber": "4551",
    "label": "SWRS Group C",
    "reviewStatus": "Pending Allocation"
  },
  {
    "id": "SWRS-DEMO-4750-A",
    "campaignNumber": "4750",
    "label": "SWRS Group A",
    "reviewStatus": "Reconciled"
  },
  {
    "id": "SWRS-DEMO-4750-B",
    "campaignNumber": "4750",
    "label": "SWRS Group B",
    "reviewStatus": "Review Required"
  },
  {
    "id": "SWRS-DEMO-4750-C",
    "campaignNumber": "4750",
    "label": "SWRS Group C",
    "reviewStatus": "Pending Allocation"
  }
];

const STLA_DEBITS = [
  {
    "id": "DEMO-DB-4551-C-A",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "market": "C",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 561.75,
    "reviewStatus": "Reconciled"
  },
  {
    "id": "DEMO-DB-4551-U-A",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "market": "U",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 306.3,
    "reviewStatus": "Review Required"
  },
  {
    "id": "DEMO-DB-4551-C-B",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "market": "C",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 897.0,
    "reviewStatus": "Pending Allocation"
  },
  {
    "id": "DEMO-DB-4551-U-B",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "market": "U",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 505.2,
    "reviewStatus": "Reconciled"
  },
  {
    "id": "DEMO-DB-4551-C-C",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "market": "C",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 1232.25,
    "reviewStatus": "Review Required"
  },
  {
    "id": "DEMO-DB-4551-U-C",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "market": "U",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 704.1,
    "reviewStatus": "Pending Allocation"
  },
  {
    "id": "DEMO-DB-4750-C-A",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "market": "C",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 576.75,
    "reviewStatus": "Reconciled"
  },
  {
    "id": "DEMO-DB-4750-U-A",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "market": "U",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 321.3,
    "reviewStatus": "Review Required"
  },
  {
    "id": "DEMO-DB-4750-C-B",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "market": "C",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 912.0,
    "reviewStatus": "Pending Allocation"
  },
  {
    "id": "DEMO-DB-4750-U-B",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "market": "U",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 520.2,
    "reviewStatus": "Reconciled"
  },
  {
    "id": "DEMO-DB-4750-C-C",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "market": "C",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 1247.25,
    "reviewStatus": "Review Required"
  },
  {
    "id": "DEMO-DB-4750-U-C",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "market": "U",
    "billingCycle": "2026-06",
    "currency": "USD",
    "billableAmount": 719.1,
    "reviewStatus": "Pending Allocation"
  }
];

const STLA_CLAIMS = [
  {
    "id": "CLM-4551-001",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "displayVin": "DEMO-VIN-4551-001",
    "claimNumber": "STLA-DEMO-4551-001",
    "modelYear": 2018,
    "familyCode": "DEMO-DS",
    "failedPart": "DEMO-PART-A",
    "otherParts": "DEMO-PART-B",
    "quantity": 1,
    "causalLop": "DEMO-LOP-01",
    "market": "C",
    "repairDate": "2025-10-01",
    "adjustedPartExpense": 150.0,
    "adjustedLopExpense": 80.0
  },
  {
    "id": "CLM-4551-002",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "displayVin": "DEMO-VIN-4551-002",
    "claimNumber": "STLA-DEMO-4551-002",
    "modelYear": 2018,
    "familyCode": "DEMO-DJ",
    "failedPart": "DEMO-PART-B",
    "otherParts": "DEMO-PART-C",
    "quantity": 2,
    "causalLop": "DEMO-LOP-02",
    "market": "U",
    "repairDate": "2025-11-02",
    "adjustedPartExpense": 187.25,
    "adjustedLopExpense": 102.1
  },
  {
    "id": "CLM-4551-003",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "displayVin": "DEMO-VIN-4551-003",
    "claimNumber": "STLA-DEMO-4551-003",
    "modelYear": 2018,
    "familyCode": "DEMO-D2",
    "failedPart": "DEMO-PART-C",
    "otherParts": "DEMO-PART-D",
    "quantity": 3,
    "causalLop": "DEMO-LOP-03",
    "market": "C",
    "repairDate": "2025-12-03",
    "adjustedPartExpense": 224.5,
    "adjustedLopExpense": 124.2
  },
  {
    "id": "CLM-4551-004",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "displayVin": "DEMO-VIN-4551-004",
    "claimNumber": "STLA-DEMO-4551-004",
    "modelYear": 2018,
    "familyCode": "DEMO-DJ",
    "failedPart": "DEMO-PART-D",
    "otherParts": "DEMO-PART-A",
    "quantity": 2,
    "causalLop": "DEMO-LOP-02",
    "market": "U",
    "repairDate": "2026-01-04",
    "adjustedPartExpense": 261.75,
    "adjustedLopExpense": 146.3
  },
  {
    "id": "CLM-4551-005",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "displayVin": "DEMO-VIN-4551-005",
    "claimNumber": "STLA-DEMO-4551-005",
    "modelYear": 2018,
    "familyCode": "DEMO-D2",
    "failedPart": "DEMO-PART-A",
    "otherParts": "DEMO-PART-B",
    "quantity": 3,
    "causalLop": "DEMO-LOP-03",
    "market": "C",
    "repairDate": "2026-02-05",
    "adjustedPartExpense": 299.0,
    "adjustedLopExpense": 168.4
  },
  {
    "id": "CLM-4551-006",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "displayVin": "DEMO-VIN-4551-006",
    "claimNumber": "STLA-DEMO-4551-006",
    "modelYear": 2018,
    "familyCode": "DEMO-DS",
    "failedPart": "DEMO-PART-B",
    "otherParts": "DEMO-PART-C",
    "quantity": 4,
    "causalLop": "DEMO-LOP-01",
    "market": "U",
    "repairDate": "2026-03-06",
    "adjustedPartExpense": 336.25,
    "adjustedLopExpense": 190.5
  },
  {
    "id": "CLM-4551-007",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "displayVin": "DEMO-VIN-4551-007",
    "claimNumber": "STLA-DEMO-4551-007",
    "modelYear": 2018,
    "familyCode": "DEMO-D2",
    "failedPart": "DEMO-PART-C",
    "otherParts": "DEMO-PART-D",
    "quantity": 3,
    "causalLop": "DEMO-LOP-03",
    "market": "C",
    "repairDate": "2026-04-07",
    "adjustedPartExpense": 373.5,
    "adjustedLopExpense": 212.6
  },
  {
    "id": "CLM-4551-008",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "displayVin": "DEMO-VIN-4551-008",
    "claimNumber": "STLA-DEMO-4551-008",
    "modelYear": 2018,
    "familyCode": "DEMO-DS",
    "failedPart": "DEMO-PART-D",
    "otherParts": "DEMO-PART-A",
    "quantity": 4,
    "causalLop": "DEMO-LOP-01",
    "market": "U",
    "repairDate": "2026-05-08",
    "adjustedPartExpense": 410.75,
    "adjustedLopExpense": 234.7
  },
  {
    "id": "CLM-4551-009",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "displayVin": "DEMO-VIN-4551-009",
    "claimNumber": "STLA-DEMO-4551-009",
    "modelYear": 2018,
    "familyCode": "DEMO-DJ",
    "failedPart": "DEMO-PART-A",
    "otherParts": "DEMO-PART-B",
    "quantity": 5,
    "causalLop": "DEMO-LOP-02",
    "market": "C",
    "repairDate": "2026-06-09",
    "adjustedPartExpense": 448.0,
    "adjustedLopExpense": 256.8
  },
  {
    "id": "CLM-4750-001",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "displayVin": "DEMO-VIN-4750-001",
    "claimNumber": "STLA-DEMO-4750-001",
    "modelYear": 2019,
    "familyCode": "DEMO-DS",
    "failedPart": "DEMO-PART-A",
    "otherParts": "DEMO-PART-B",
    "quantity": 1,
    "causalLop": "DEMO-LOP-01",
    "market": "C",
    "repairDate": "2025-10-01",
    "adjustedPartExpense": 150.0,
    "adjustedLopExpense": 80.0
  },
  {
    "id": "CLM-4750-002",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "displayVin": "DEMO-VIN-4750-002",
    "claimNumber": "STLA-DEMO-4750-002",
    "modelYear": 2019,
    "familyCode": "DEMO-DJ",
    "failedPart": "DEMO-PART-B",
    "otherParts": "DEMO-PART-C",
    "quantity": 2,
    "causalLop": "DEMO-LOP-02",
    "market": "U",
    "repairDate": "2025-11-02",
    "adjustedPartExpense": 187.25,
    "adjustedLopExpense": 102.1
  },
  {
    "id": "CLM-4750-003",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "displayVin": "DEMO-VIN-4750-003",
    "claimNumber": "STLA-DEMO-4750-003",
    "modelYear": 2019,
    "familyCode": "DEMO-D2",
    "failedPart": "DEMO-PART-C",
    "otherParts": "DEMO-PART-D",
    "quantity": 3,
    "causalLop": "DEMO-LOP-03",
    "market": "C",
    "repairDate": "2025-12-03",
    "adjustedPartExpense": 224.5,
    "adjustedLopExpense": 124.2
  },
  {
    "id": "CLM-4750-004",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "displayVin": "DEMO-VIN-4750-004",
    "claimNumber": "STLA-DEMO-4750-004",
    "modelYear": 2019,
    "familyCode": "DEMO-DJ",
    "failedPart": "DEMO-PART-D",
    "otherParts": "DEMO-PART-A",
    "quantity": 2,
    "causalLop": "DEMO-LOP-02",
    "market": "U",
    "repairDate": "2026-01-04",
    "adjustedPartExpense": 261.75,
    "adjustedLopExpense": 146.3
  },
  {
    "id": "CLM-4750-005",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "displayVin": "DEMO-VIN-4750-005",
    "claimNumber": "STLA-DEMO-4750-005",
    "modelYear": 2019,
    "familyCode": "DEMO-D2",
    "failedPart": "DEMO-PART-A",
    "otherParts": "DEMO-PART-B",
    "quantity": 3,
    "causalLop": "DEMO-LOP-03",
    "market": "C",
    "repairDate": "2026-02-05",
    "adjustedPartExpense": 299.0,
    "adjustedLopExpense": 168.4
  },
  {
    "id": "CLM-4750-006",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "displayVin": "DEMO-VIN-4750-006",
    "claimNumber": "STLA-DEMO-4750-006",
    "modelYear": 2019,
    "familyCode": "DEMO-DS",
    "failedPart": "DEMO-PART-B",
    "otherParts": "DEMO-PART-C",
    "quantity": 4,
    "causalLop": "DEMO-LOP-01",
    "market": "U",
    "repairDate": "2026-03-06",
    "adjustedPartExpense": 336.25,
    "adjustedLopExpense": 190.5
  },
  {
    "id": "CLM-4750-007",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "displayVin": "DEMO-VIN-4750-007",
    "claimNumber": "STLA-DEMO-4750-007",
    "modelYear": 2019,
    "familyCode": "DEMO-D2",
    "failedPart": "DEMO-PART-C",
    "otherParts": "DEMO-PART-D",
    "quantity": 3,
    "causalLop": "DEMO-LOP-03",
    "market": "C",
    "repairDate": "2026-04-07",
    "adjustedPartExpense": 373.5,
    "adjustedLopExpense": 212.6
  },
  {
    "id": "CLM-4750-008",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "displayVin": "DEMO-VIN-4750-008",
    "claimNumber": "STLA-DEMO-4750-008",
    "modelYear": 2019,
    "familyCode": "DEMO-DS",
    "failedPart": "DEMO-PART-D",
    "otherParts": "DEMO-PART-A",
    "quantity": 4,
    "causalLop": "DEMO-LOP-01",
    "market": "U",
    "repairDate": "2026-05-08",
    "adjustedPartExpense": 410.75,
    "adjustedLopExpense": 234.7
  },
  {
    "id": "CLM-4750-009",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "displayVin": "DEMO-VIN-4750-009",
    "claimNumber": "STLA-DEMO-4750-009",
    "modelYear": 2019,
    "familyCode": "DEMO-DJ",
    "failedPart": "DEMO-PART-A",
    "otherParts": "DEMO-PART-B",
    "quantity": 5,
    "causalLop": "DEMO-LOP-02",
    "market": "C",
    "repairDate": "2026-06-09",
    "adjustedPartExpense": 448.0,
    "adjustedLopExpense": 256.8
  }
];

const STLA_RECOVERY_LINES = [
  {
    "id": "REC-DEMO-4551-C-A",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "market": "C",
    "debitId": "DEMO-DB-4551-C-A",
    "lop": "DEMO-LOP-01",
    "numberOfConditions": 1,
    "adjustedLaborExpense": 196.61,
    "adjustedPartExpense": 365.14,
    "supplierResponsibilityPct": 50
  },
  {
    "id": "REC-DEMO-4551-U-A",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-A",
    "market": "U",
    "debitId": "DEMO-DB-4551-U-A",
    "lop": "DEMO-LOP-02",
    "numberOfConditions": 2,
    "adjustedLaborExpense": 107.2,
    "adjustedPartExpense": 199.1,
    "supplierResponsibilityPct": 75
  },
  {
    "id": "REC-DEMO-4551-C-B",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "market": "C",
    "debitId": "DEMO-DB-4551-C-B",
    "lop": "DEMO-LOP-03",
    "numberOfConditions": 3,
    "adjustedLaborExpense": 313.95,
    "adjustedPartExpense": 583.05,
    "supplierResponsibilityPct": 100
  },
  {
    "id": "REC-DEMO-4551-U-B",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-B",
    "market": "U",
    "debitId": "DEMO-DB-4551-U-B",
    "lop": "DEMO-LOP-01",
    "numberOfConditions": 4,
    "adjustedLaborExpense": 176.82,
    "adjustedPartExpense": 328.38,
    "supplierResponsibilityPct": 50
  },
  {
    "id": "REC-DEMO-4551-C-C",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "market": "C",
    "debitId": "DEMO-DB-4551-C-C",
    "lop": "DEMO-LOP-02",
    "numberOfConditions": 1,
    "adjustedLaborExpense": 431.29,
    "adjustedPartExpense": 800.96,
    "supplierResponsibilityPct": 75
  },
  {
    "id": "REC-DEMO-4551-U-C",
    "campaignNumber": "4551",
    "swrsId": "SWRS-DEMO-4551-C",
    "market": "U",
    "debitId": "DEMO-DB-4551-U-C",
    "lop": "DEMO-LOP-03",
    "numberOfConditions": 2,
    "adjustedLaborExpense": 246.44,
    "adjustedPartExpense": 457.67,
    "supplierResponsibilityPct": 100
  },
  {
    "id": "REC-DEMO-4750-C-A",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "market": "C",
    "debitId": "DEMO-DB-4750-C-A",
    "lop": "DEMO-LOP-01",
    "numberOfConditions": 3,
    "adjustedLaborExpense": 201.86,
    "adjustedPartExpense": 374.89,
    "supplierResponsibilityPct": 50
  },
  {
    "id": "REC-DEMO-4750-U-A",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-A",
    "market": "U",
    "debitId": "DEMO-DB-4750-U-A",
    "lop": "DEMO-LOP-02",
    "numberOfConditions": 4,
    "adjustedLaborExpense": 112.45,
    "adjustedPartExpense": 208.85,
    "supplierResponsibilityPct": 75
  },
  {
    "id": "REC-DEMO-4750-C-B",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "market": "C",
    "debitId": "DEMO-DB-4750-C-B",
    "lop": "DEMO-LOP-03",
    "numberOfConditions": 1,
    "adjustedLaborExpense": 319.2,
    "adjustedPartExpense": 592.8,
    "supplierResponsibilityPct": 100
  },
  {
    "id": "REC-DEMO-4750-U-B",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-B",
    "market": "U",
    "debitId": "DEMO-DB-4750-U-B",
    "lop": "DEMO-LOP-01",
    "numberOfConditions": 2,
    "adjustedLaborExpense": 182.07,
    "adjustedPartExpense": 338.13,
    "supplierResponsibilityPct": 50
  },
  {
    "id": "REC-DEMO-4750-C-C",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "market": "C",
    "debitId": "DEMO-DB-4750-C-C",
    "lop": "DEMO-LOP-02",
    "numberOfConditions": 3,
    "adjustedLaborExpense": 436.54,
    "adjustedPartExpense": 810.71,
    "supplierResponsibilityPct": 75
  },
  {
    "id": "REC-DEMO-4750-U-C",
    "campaignNumber": "4750",
    "swrsId": "SWRS-DEMO-4750-C",
    "market": "U",
    "debitId": "DEMO-DB-4750-U-C",
    "lop": "DEMO-LOP-03",
    "numberOfConditions": 4,
    "adjustedLaborExpense": 251.69,
    "adjustedPartExpense": 467.42,
    "supplierResponsibilityPct": 100
  }
];

// ═════════════════════════════════════════════════════════════════
// STELLANTIS WORKSPACE — Increment (Phase B5 demo)
// Synthetic fixture data only. Read-only, non-persistent, no API calls.
// Terminology: "Campaign" (not "Case") per Stellantis SWRS convention.
// ═════════════════════════════════════════════════════════════════

function stlaCurrency(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) {
    return '\u2014';
  }
  return Number(value).toLocaleString('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function stlaClaimTotal(claim) {
  return Number(claim.adjustedPartExpense || 0)
    + Number(claim.adjustedLopExpense || 0);
}

function stlaRecoveryTotal(line) {
  return Number(line.adjustedLaborExpense || 0)
    + Number(line.adjustedPartExpense || 0);
}

function stlaSupplierResponsibility(line) {
  return stlaRecoveryTotal(line)
    * Number(line.supplierResponsibilityPct || 0)
    / 100;
}

const STLA_STATUS_COLORS = {
  'Reconciled':'#3F8F5B','Review Required':'#C08A2E','Pending Allocation':'#7C7669',
  'Recovery Review':'#1B6FB8','Allocation Review':'#C08A2E',
};
const StlaPill = ({label}) => {
  const c = STLA_STATUS_COLORS[label] || '#7C7669';
  return <span style={{display:'inline-flex',alignItems:'center',gap:5,height:20,padding:'0 8px',borderRadius:6,fontSize:11,fontWeight:500,background:c+'1F',color:c,lineHeight:1}}>
    <span style={{width:6,height:6,borderRadius:'50%',background:c}}/>{label}
  </span>;
};

function StlaKPI({label,value,hint}) {
  return (
    <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,padding:'12px 14px',display:'flex',flexDirection:'column',gap:4}}>
      <div style={{fontSize:11,color:'#7C7669',fontWeight:500,textTransform:'uppercase',letterSpacing:'0.02em'}}>{label}</div>
      <div style={{fontSize:20,fontWeight:600,color:'#1B2A44',fontVariantNumeric:'tabular-nums'}}>{value}</div>
      {hint && <div style={{fontSize:11,color:'#7C7669'}}>{hint}</div>}
    </div>
  );
}

const STLA_SYNTHETIC_NOTICE = 'Synthetic demonstration data only \u2014 no original VIN, claim, dealer, supplier, or financial records are included.';

function StlaNotice() {
  return (
    <div style={{background:'#FBF4E9',border:'1px solid #E9D9B6',borderRadius:8,padding:'8px 12px',fontSize:12,color:'#7C5A22',marginBottom:12}}>
      {STLA_SYNTHETIC_NOTICE}
    </div>
  );
}

function StlaEmpty({message}) {
  return <div style={{padding:'24px 12px',textAlign:'center',color:'#7C7669',fontSize:13}}>{message}</div>;
}

const STLA_TH = {textAlign:'left',fontSize:11,fontWeight:600,color:'#7C7669',textTransform:'uppercase',letterSpacing:'0.02em',padding:'8px 10px',borderBottom:'1px solid #E7E4DC',whiteSpace:'nowrap'};
const STLA_TD = {fontSize:12,color:'#1B2A44',padding:'8px 10px',borderBottom:'1px solid #F1EFE8',whiteSpace:'nowrap'};

function StellantisWorkspace(){
  const [tab, setTab] = React.useState('dashboard');
  const [campaignFilter, setCampaignFilter] = React.useState('ALL');
  const [search, setSearch] = React.useState('');
  const [marketFilter, setMarketFilter] = React.useState('ALL');
  const [detailCampaign, setDetailCampaign] = React.useState(null);

  const clearFilters = () => {
    setCampaignFilter('ALL');
    setSearch('');
    setMarketFilter('ALL');
  };

  const q = search.trim().toLowerCase();

  const campaigns = React.useMemo(() => STLA_CAMPAIGNS.filter(c =>
    (campaignFilter==='ALL' || c.campaignNumber===campaignFilter)
  ), [campaignFilter]);

  const swrsGroups = React.useMemo(() => STLA_SWRS_GROUPS.filter(g =>
    (campaignFilter==='ALL' || g.campaignNumber===campaignFilter) &&
    (!q || g.campaignNumber.toLowerCase().includes(q) || g.id.toLowerCase().includes(q) || g.reviewStatus.toLowerCase().includes(q))
  ), [campaignFilter, q]);

  const debits = React.useMemo(() => STLA_DEBITS.filter(d =>
    (campaignFilter==='ALL' || d.campaignNumber===campaignFilter) &&
    (marketFilter==='ALL' || d.market===marketFilter) &&
    (!q || d.campaignNumber.toLowerCase().includes(q) || d.id.toLowerCase().includes(q) || d.swrsId.toLowerCase().includes(q) || d.market.toLowerCase().includes(q) || d.reviewStatus.toLowerCase().includes(q))
  ), [campaignFilter, marketFilter, q]);

  const claims = React.useMemo(() => STLA_CLAIMS.filter(c =>
    (campaignFilter==='ALL' || c.campaignNumber===campaignFilter) &&
    (marketFilter==='ALL' || c.market===marketFilter) &&
    (!q || [c.campaignNumber,c.swrsId,c.displayVin,c.claimNumber,c.familyCode,c.failedPart,c.causalLop].some(v=>String(v).toLowerCase().includes(q)))
  ), [campaignFilter, marketFilter, q]);

  const recoveryLines = React.useMemo(() => STLA_RECOVERY_LINES.filter(r =>
    (campaignFilter==='ALL' || r.campaignNumber===campaignFilter) &&
    (marketFilter==='ALL' || r.market===marketFilter) &&
    (!q || [r.campaignNumber,r.swrsId,r.debitId,r.lop].some(v=>String(v).toLowerCase().includes(q)))
  ), [campaignFilter, marketFilter, q]);

  // Dashboard derived values (campaign-filtered, both markets shown for comparison)
  const dashDebits = React.useMemo(() => STLA_DEBITS.filter(d => campaignFilter==='ALL' || d.campaignNumber===campaignFilter), [campaignFilter]);
  const dashRecovery = React.useMemo(() => STLA_RECOVERY_LINES.filter(r => campaignFilter==='ALL' || r.campaignNumber===campaignFilter), [campaignFilter]);

  const kpiCampaignCount = campaigns.length;
  const kpiSwrsCount = swrsGroups.length;
  const kpiClaimCount = claims.length;
  const kpiUniqueVins = new Set(claims.map(c=>c.displayVin)).size;
  const kpiDebitCount = dashDebits.length;
  const kpiDebitTotal = dashDebits.reduce((s,d)=>s+Number(d.billableAmount||0),0);
  const kpiClaimTotal = claims.reduce((s,c)=>s+stlaClaimTotal(c),0);
  const kpiSupplierResp = dashRecovery.reduce((s,r)=>s+stlaSupplierResponsibility(r),0);

  const marketCTotal = dashDebits.filter(d=>d.market==='C').reduce((s,d)=>s+Number(d.billableAmount||0),0);
  const marketUTotal = dashDebits.filter(d=>d.market==='U').reduce((s,d)=>s+Number(d.billableAmount||0),0);

  const swrsStatusDist = React.useMemo(() => {
    const dist = {};
    STLA_SWRS_GROUPS.filter(g => campaignFilter==='ALL' || g.campaignNumber===campaignFilter).forEach(g => {
      dist[g.reviewStatus] = (dist[g.reviewStatus]||0)+1;
    });
    return dist;
  }, [campaignFilter]);

  const recoveryByCampaign = React.useMemo(() => {
    const map = {};
    ['4551','4750'].forEach(cn => {
      if (campaignFilter!=='ALL' && campaignFilter!==cn) return;
      const lines = STLA_RECOVERY_LINES.filter(r=>r.campaignNumber===cn);
      map[cn] = lines.reduce((s,r)=>s+stlaSupplierResponsibility(r),0);
    });
    return map;
  }, [campaignFilter]);

  const swrsMap = React.useMemo(() => {
    const m = {};
    STLA_SWRS_GROUPS.forEach(g => { m[g.id] = g; });
    return m;
  }, []);

  const TabBtn = ({id,label}) => (
    <button onClick={()=>setTab(id)} style={{
      padding:'8px 14px',fontSize:12,fontWeight:600,borderRadius:6,border:'1px solid '+(tab===id?'#C0574E':'#E7E4DC'),
      background:tab===id?'#C0574E':'#fff',color:tab===id?'#fff':'#1B2A44',cursor:'pointer'
    }}>{label}</button>
  );

  return (
    <div style={{padding:20,display:'flex',flexDirection:'column',gap:16}}>
      <div>
        <div style={{fontSize:18,fontWeight:700,color:'#1B2A44'}}>Stellantis Workspace</div>
        <div style={{fontSize:12,color:'#7C7669'}}>Read-only beta prototype for Campaign recovery analysis using synthetic demonstration fixtures.</div>
      </div>

      <StlaNotice/>

      <div style={{display:'flex',flexWrap:'wrap',gap:10,alignItems:'center'}}>
        <select value={campaignFilter} onChange={e=>setCampaignFilter(e.target.value)} style={{padding:'6px 10px',fontSize:12,borderRadius:6,border:'1px solid #E7E4DC'}}>
          <option value="ALL">All Campaigns</option>
          <option value="4551">Campaign 4551</option>
          <option value="4750">Campaign 4750</option>
        </select>
        <select value={marketFilter} onChange={e=>setMarketFilter(e.target.value)} style={{padding:'6px 10px',fontSize:12,borderRadius:6,border:'1px solid #E7E4DC'}}>
          <option value="ALL">All Markets</option>
          <option value="C">Market C</option>
          <option value="U">Market U</option>
        </select>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search..." style={{padding:'6px 10px',fontSize:12,borderRadius:6,border:'1px solid #E7E4DC',minWidth:200}}/>
        <button onClick={clearFilters} style={{padding:'6px 12px',fontSize:12,borderRadius:6,border:'1px solid #E7E4DC',background:'#fff',cursor:'pointer'}}>Clear Filters</button>
      </div>

      <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
        <TabBtn id="dashboard" label="Campaign Dashboard"/>
        <TabBtn id="registry" label="Campaign Registry"/>
        <TabBtn id="swrs" label="SWRS Groups"/>
        <TabBtn id="debits" label="Debit Registry"/>
        <TabBtn id="claims" label="Claim Explorer"/>
        <TabBtn id="recovery" label="Cost Recovery"/>
      </div>

      {tab==='dashboard' && (
        <div style={{display:'flex',flexDirection:'column',gap:16}}>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:10}}>
            <StlaKPI label="Campaigns" value={kpiCampaignCount}/>
            <StlaKPI label="SWRS Groups" value={kpiSwrsCount}/>
            <StlaKPI label="Claims" value={kpiClaimCount}/>
            <StlaKPI label="Unique Demo VINs" value={kpiUniqueVins}/>
            <StlaKPI label="Debit Records" value={kpiDebitCount}/>
            <StlaKPI label="Total Debit Amount" value={stlaCurrency(kpiDebitTotal)}/>
            <StlaKPI label="Adjusted Claim Total" value={stlaCurrency(kpiClaimTotal)}/>
            <StlaKPI label="Supplier Responsibility" value={stlaCurrency(kpiSupplierResp)}/>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(260px,1fr))',gap:16}}>
            <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,padding:14}}>
              <div style={{fontSize:12,fontWeight:600,color:'#1B2A44',marginBottom:8,textTransform:'uppercase'}}>Recovery by Campaign</div>
              {Object.keys(recoveryByCampaign).length===0 ? <StlaEmpty message="No synthetic Campaigns match the current filters."/> :
                Object.entries(recoveryByCampaign).map(([cn,val]) => (
                  <div key={cn} style={{display:'flex',justifyContent:'space-between',fontSize:12,padding:'4px 0'}}>
                    <span>Campaign {cn}</span><span style={{fontVariantNumeric:'tabular-nums'}}>{stlaCurrency(val)}</span>
                  </div>
                ))}
            </div>
            <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,padding:14}}>
              <div style={{fontSize:12,fontWeight:600,color:'#1B2A44',marginBottom:8,textTransform:'uppercase'}}>Market C vs Market U Debit Totals</div>
              <div style={{display:'flex',justifyContent:'space-between',fontSize:12,padding:'4px 0'}}><span>Market C</span><span style={{fontVariantNumeric:'tabular-nums'}}>{stlaCurrency(marketCTotal)}</span></div>
              <div style={{display:'flex',justifyContent:'space-between',fontSize:12,padding:'4px 0'}}><span>Market U</span><span style={{fontVariantNumeric:'tabular-nums'}}>{stlaCurrency(marketUTotal)}</span></div>
            </div>
            <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,padding:14}}>
              <div style={{fontSize:12,fontWeight:600,color:'#1B2A44',marginBottom:8,textTransform:'uppercase'}}>SWRS Review Status Distribution</div>
              {Object.keys(swrsStatusDist).length===0 ? <StlaEmpty message="No synthetic SWRS groups match the current filters."/> :
                Object.entries(swrsStatusDist).map(([s,c]) => (
                  <div key={s} style={{display:'flex',justifyContent:'space-between',fontSize:12,padding:'4px 0'}}><StlaPill label={s}/><span>{c}</span></div>
                ))}
            </div>
          </div>
          <div style={{fontSize:11,color:'#7C7669'}}>All calculations use synthetic demonstration fixtures.</div>
        </div>
      )}

      {tab==='registry' && (
        <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,overflowX:'auto'}}>
          {campaigns.length===0 ? <StlaEmpty message="No synthetic Campaigns match the current filters."/> : (
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>
              <th style={STLA_TH}>Campaign Number</th><th style={STLA_TH}>Customer</th><th style={STLA_TH}>Plant</th>
              <th style={STLA_TH}>Model Year</th><th style={STLA_TH}>Claim Classification</th><th style={STLA_TH}>Billing Cycle</th>
              <th style={STLA_TH}>SWRS Groups</th><th style={STLA_TH}>Debit Records</th><th style={STLA_TH}>Synthetic Claims</th>
              <th style={STLA_TH}>Debit Total</th><th style={STLA_TH}>Status</th>
            </tr></thead>
            <tbody>
              {campaigns.map(c => {
                const cSwrs = STLA_SWRS_GROUPS.filter(g=>g.campaignNumber===c.campaignNumber);
                const cDebits = STLA_DEBITS.filter(d=>d.campaignNumber===c.campaignNumber);
                const cClaims = STLA_CLAIMS.filter(cl=>cl.campaignNumber===c.campaignNumber);
                const debitTotal = cDebits.reduce((s,d)=>s+Number(d.billableAmount||0),0);
                return (
                  <tr key={c.campaignNumber} onClick={()=>setDetailCampaign(c.campaignNumber)} style={{cursor:'pointer'}}>
                    <td style={STLA_TD}>Campaign {c.campaignNumber}</td>
                    <td style={STLA_TD}>{c.customer}</td>
                    <td style={STLA_TD}>{c.plant}</td>
                    <td style={STLA_TD}>{c.modelYear}</td>
                    <td style={STLA_TD}>{c.claimClassification}</td>
                    <td style={STLA_TD}>{c.billingCycle}</td>
                    <td style={STLA_TD}>{cSwrs.length}</td>
                    <td style={STLA_TD}>{cDebits.length}</td>
                    <td style={STLA_TD}>{cClaims.length}</td>
                    <td style={STLA_TD}>{stlaCurrency(debitTotal)}</td>
                    <td style={STLA_TD}><StlaPill label={c.status}/></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          )}
          {detailCampaign && (() => {
            const c = STLA_CAMPAIGNS.find(x=>x.campaignNumber===detailCampaign);
            const cSwrs = STLA_SWRS_GROUPS.filter(g=>g.campaignNumber===detailCampaign);
            const cDebits = STLA_DEBITS.filter(d=>d.campaignNumber===detailCampaign);
            const cClaims = STLA_CLAIMS.filter(cl=>cl.campaignNumber===detailCampaign);
            const debitTotal = cDebits.reduce((s,d)=>s+Number(d.billableAmount||0),0);
            const claimTotal = cClaims.reduce((s,cl)=>s+stlaClaimTotal(cl),0);
            return (
              <div style={{padding:14,borderTop:'1px solid #E7E4DC',fontSize:12}}>
                <div style={{fontWeight:600,marginBottom:6}}>Campaign {c.campaignNumber} Detail</div>
                <div>SWRS Groups: {cSwrs.length}</div>
                <div>Debit Records: {cDebits.length}</div>
                <div>Claims: {cClaims.length}</div>
                <div>Debit Total: {stlaCurrency(debitTotal)}</div>
                <div>Adjusted Claim Total: {stlaCurrency(claimTotal)}</div>
                <div>Status: <StlaPill label={c.status}/></div>
                <div style={{marginTop:6,color:'#7C7669'}}>{STLA_SYNTHETIC_NOTICE}</div>
                <button onClick={()=>setDetailCampaign(null)} style={{marginTop:8,padding:'4px 10px',fontSize:11,borderRadius:6,border:'1px solid #E7E4DC',background:'#fff',cursor:'pointer'}}>Close</button>
              </div>
            );
          })()}
        </div>
      )}

      {tab==='swrs' && (
        <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,overflowX:'auto'}}>
          {swrsGroups.length===0 ? <StlaEmpty message="No synthetic SWRS groups match the current filters."/> : (
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>
              <th style={STLA_TH}>Campaign</th><th style={STLA_TH}>SWRS Identifier</th><th style={STLA_TH}>Synthetic Claim Count</th>
              <th style={STLA_TH}>Market C Debit Amount</th><th style={STLA_TH}>Market U Debit Amount</th><th style={STLA_TH}>Total Debit Amount</th>
              <th style={STLA_TH}>Related Debit Count</th><th style={STLA_TH}>Review Status</th>
            </tr></thead>
            <tbody>
              {swrsGroups.map(g => {
                const gDebits = STLA_DEBITS.filter(d=>d.swrsId===g.id);
                const cAmt = gDebits.filter(d=>d.market==='C').reduce((s,d)=>s+Number(d.billableAmount||0),0);
                const uAmt = gDebits.filter(d=>d.market==='U').reduce((s,d)=>s+Number(d.billableAmount||0),0);
                const gClaims = STLA_CLAIMS.filter(cl=>cl.swrsId===g.id);
                return (
                  <tr key={g.id}>
                    <td style={STLA_TD}>Campaign {g.campaignNumber}</td>
                    <td style={STLA_TD}>{g.id}</td>
                    <td style={STLA_TD}>{gClaims.length}</td>
                    <td style={STLA_TD}>{stlaCurrency(cAmt)}</td>
                    <td style={STLA_TD}>{stlaCurrency(uAmt)}</td>
                    <td style={STLA_TD}>{stlaCurrency(cAmt+uAmt)}</td>
                    <td style={STLA_TD}>{gDebits.length}</td>
                    <td style={STLA_TD}><StlaPill label={g.reviewStatus}/></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          )}
        </div>
      )}

      {tab==='debits' && (
        <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,overflowX:'auto'}}>
          {debits.length===0 ? <StlaEmpty message="No synthetic debit records match the current filters."/> : (
          <>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>
              <th style={STLA_TH}>Campaign</th><th style={STLA_TH}>Debit Identifier</th><th style={STLA_TH}>SWRS Identifier</th>
              <th style={STLA_TH}>Market</th><th style={STLA_TH}>Billing Cycle</th><th style={STLA_TH}>Currency</th>
              <th style={STLA_TH}>Billable Amount</th><th style={STLA_TH}>Review Status</th>
            </tr></thead>
            <tbody>
              {debits.map(d => (
                <tr key={d.id}>
                  <td style={STLA_TD}>Campaign {d.campaignNumber}</td>
                  <td style={STLA_TD}>{d.id}</td>
                  <td style={STLA_TD}>{d.swrsId}</td>
                  <td style={STLA_TD}>{d.market}</td>
                  <td style={STLA_TD}>{d.billingCycle}</td>
                  <td style={STLA_TD}>{d.currency}</td>
                  <td style={STLA_TD}>{stlaCurrency(d.billableAmount)}</td>
                  <td style={STLA_TD}><StlaPill label={d.reviewStatus}/></td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{padding:'8px 10px',fontSize:11,color:'#7C7669',borderTop:'1px solid #E7E4DC'}}>
            Visible: {debits.length} debit records \u2014 {stlaCurrency(debits.reduce((s,d)=>s+Number(d.billableAmount||0),0))}
          </div>
          </>
          )}
        </div>
      )}

      {tab==='claims' && (
        <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,overflowX:'auto'}}>
          {claims.length===0 ? <StlaEmpty message="No synthetic claims match the current filters."/> : (
          <>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>
              <th style={STLA_TH}>Campaign</th><th style={STLA_TH}>SWRS</th><th style={STLA_TH}>Demo VIN</th>
              <th style={STLA_TH}>Synthetic Claim Number</th><th style={STLA_TH}>Model Year</th><th style={STLA_TH}>Family Code</th>
              <th style={STLA_TH}>Failed Part</th><th style={STLA_TH}>Other Parts</th><th style={STLA_TH}>Quantity</th>
              <th style={STLA_TH}>Causal LOP</th><th style={STLA_TH}>Market</th><th style={STLA_TH}>Repair Date</th>
              <th style={STLA_TH}>Adjusted Part Expense</th><th style={STLA_TH}>Adjusted LOP Expense</th><th style={STLA_TH}>Adjusted Total</th>
            </tr></thead>
            <tbody>
              {claims.map(c => (
                <tr key={c.id}>
                  <td style={STLA_TD}>Campaign {c.campaignNumber}</td>
                  <td style={STLA_TD}>{c.swrsId}</td>
                  <td style={STLA_TD}>{c.displayVin}</td>
                  <td style={STLA_TD}>{c.claimNumber}</td>
                  <td style={STLA_TD}>{c.modelYear}</td>
                  <td style={STLA_TD}>{c.familyCode}</td>
                  <td style={STLA_TD}>{c.failedPart}</td>
                  <td style={STLA_TD}>{c.otherParts}</td>
                  <td style={STLA_TD}>{c.quantity}</td>
                  <td style={STLA_TD}>{c.causalLop}</td>
                  <td style={STLA_TD}>{c.market}</td>
                  <td style={STLA_TD}>{c.repairDate}</td>
                  <td style={STLA_TD}>{stlaCurrency(c.adjustedPartExpense)}</td>
                  <td style={STLA_TD}>{stlaCurrency(c.adjustedLopExpense)}</td>
                  <td style={STLA_TD}>{stlaCurrency(stlaClaimTotal(c))}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{padding:'8px 10px',fontSize:11,color:'#7C7669',borderTop:'1px solid #E7E4DC'}}>
            Visible: {claims.length} claims \u2014 {stlaCurrency(claims.reduce((s,c)=>s+stlaClaimTotal(c),0))} \u2014 {new Set(claims.map(c=>c.displayVin)).size} distinct Demo VINs
          </div>
          </>
          )}
        </div>
      )}

      {tab==='recovery' && (
        <div style={{background:'#fff',border:'1px solid #E7E4DC',borderRadius:8,overflowX:'auto'}}>
          {recoveryLines.length===0 ? <StlaEmpty message="No synthetic recovery lines match the current filters."/> : (
          <>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>
              <th style={STLA_TH}>Campaign</th><th style={STLA_TH}>SWRS</th><th style={STLA_TH}>Market</th><th style={STLA_TH}>Debit Identifier</th>
              <th style={STLA_TH}>LOP</th><th style={STLA_TH}>Number of Conditions</th><th style={STLA_TH}>Adjusted Labor Expense</th>
              <th style={STLA_TH}>Adjusted Part Expense</th><th style={STLA_TH}>Adjusted Total Expense</th>
              <th style={STLA_TH}>Supplier Responsibility %</th><th style={STLA_TH}>Supplier Responsibility Expense</th>
            </tr></thead>
            <tbody>
              {recoveryLines.map(r => (
                <tr key={r.id}>
                  <td style={STLA_TD}>Campaign {r.campaignNumber}</td>
                  <td style={STLA_TD}>{r.swrsId}</td>
                  <td style={STLA_TD}>{r.market}</td>
                  <td style={STLA_TD}>{r.debitId}</td>
                  <td style={STLA_TD}>{r.lop}</td>
                  <td style={STLA_TD}>{r.numberOfConditions}</td>
                  <td style={STLA_TD}>{stlaCurrency(r.adjustedLaborExpense)}</td>
                  <td style={STLA_TD}>{stlaCurrency(r.adjustedPartExpense)}</td>
                  <td style={STLA_TD}>{stlaCurrency(stlaRecoveryTotal(r))}</td>
                  <td style={STLA_TD}>{r.supplierResponsibilityPct}%</td>
                  <td style={STLA_TD}>{stlaCurrency(stlaSupplierResponsibility(r))}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{padding:'8px 10px',fontSize:11,color:'#7C7669',borderTop:'1px solid #E7E4DC'}}>
            Visible conditions: {recoveryLines.reduce((s,r)=>s+Number(r.numberOfConditions||0),0)} \u2014 Adjusted Total: {stlaCurrency(recoveryLines.reduce((s,r)=>s+stlaRecoveryTotal(r),0))} \u2014 Supplier Responsibility: {stlaCurrency(recoveryLines.reduce((s,r)=>s+stlaSupplierResponsibility(r),0))}
          </div>
          </>
          )}
        </div>
      )}
    </div>
  );
}



function WarrantyMISApp({areaSwitch,wmisTab,setWmisTab}){
  const [wmisOem,   setWmisOem]   = useState('ALL');
  const [wmisSearch,setWmisSearch]= useState('');
  return (
    <div className="jss-app" style={{display:'flex',height:'100vh',background:C.bg,
      fontFamily:'"Inter",-apple-system,sans-serif',color:C.ink,overflow:'hidden'}}>
      <TableStyles/>
      <aside style={{width:'215px',flexShrink:0,background:C.navBg,display:'flex',
        flexDirection:'column',padding:'14px 10px 10px',overflowY:'auto'}}>
        {areaSwitch}
        <JSSLogo/>
        <div style={{padding:'4px 8px 6px',marginBottom:'4px'}}>
          <span style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.14em',
            color:'rgba(255,255,255,0.42)',textTransform:'uppercase'}}>Warranty MIS</span>
        </div>
        <nav style={{flex:1,display:'flex',flexDirection:'column',gap:'1px'}}>
          {WMIS_NAV.map(n=>{
            const a = wmisTab === n.id;
            return (
              <button key={n.id} onClick={()=>setWmisTab(n.id)} style={{
                position:'relative',display:'flex',alignItems:'center',gap:'10px',
                height:'32px',padding:'0 10px 0 12px',borderRadius:'6px',border:'none',
                cursor:'pointer',textAlign:'left',fontSize:'12.5px',
                fontWeight:a?600:400,
                background:a?'rgba(255,255,255,0.08)':'transparent',
                color:a?'#fff':'rgba(255,255,255,0.62)',
                transition:'background 120ms, color 120ms',
                boxShadow:a?`inset 2px 0 0 ${C.coral}`:'none'}}
                onMouseEnter={e=>{if(!a)e.currentTarget.style.background='rgba(255,255,255,0.05)';}}
                onMouseLeave={e=>{if(!a)e.currentTarget.style.background='transparent';}}>
                {n.label}
              </button>
            );
          })}
        </nav>
      </aside>
      <div style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
        <header style={{padding:'12px 24px',display:'flex',alignItems:'center',gap:'10px',
          background:C.navBg,borderBottom:'1px solid rgba(0,0,0,0.35)',
          boxShadow:'inset 0 -1px 0 rgba(255,255,255,0.06)',flexShrink:0}}>
          <div style={{flex:1}}>
            <div style={{fontSize:'15px',fontWeight:600,color:'#fff',letterSpacing:'-0.005em'}}>JSS Warranty MIS</div>
            <div style={{fontSize:'11px',color:'rgba(255,255,255,0.55)',marginTop:'1px'}}>Warranty Management Information System</div>
          </div>
          <span style={{fontSize:'9.5px',fontWeight:600,letterSpacing:'0.14em',
            color:'rgba(255,255,255,0.45)',textTransform:'uppercase',marginRight:'2px'}}>Filters</span>
          <Sel value={wmisOem} onChange={e=>setWmisOem(e.target.value)}
            style={{minWidth:'150px',height:'30px',padding:'0 10px',
              background:'rgba(255,255,255,0.10)',border:'1px solid rgba(255,255,255,0.15)',
              borderRadius:'6px',color:'#fff',fontSize:'12px'}}>
            {WMIS_OEM_OPTIONS.map(o=>(
              <option key={o.value} value={o.value} style={{color:'#1B2A5E'}}>{o.label}</option>
            ))}
          </Sel>
          <Inp placeholder="Search…" value={wmisSearch} onChange={e=>setWmisSearch(e.target.value)}
            style={{height:'30px',padding:'0 10px',width:'200px',
              background:'rgba(255,255,255,0.10)',border:'1px solid rgba(255,255,255,0.15)',
              borderRadius:'6px',color:'#fff',fontSize:'12px'}}/>
        </header>
        <div style={{flex:1,overflow:'auto',padding:'18px 26px 32px'}}>
          {wmisTab==='wmis-overview'            && <WMIS_Placeholder id="wmis-overview"/>}
          {wmisTab==='wmis-recall-intelligence' && <RecallIntelligence/>}
          {wmisTab==='wmis-gm'                  && <GMWorkspace/>}
          {wmisTab==='wmis-stellantis'          && <StellantisWorkspace/>}
          {wmisTab==='wmis-parts'               && <WMIS_Placeholder id="wmis-parts"/>}
          {wmisTab==='wmis-operations'          && <WMIS_Placeholder id="wmis-operations"/>}
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════
// FUTURE DOMAIN-SPECIFIC BLOCKS
// (NHTSA / GM / Stellantis / Parts CrossRef / Operations)
// will be added here in later increments.
// ══════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════
// TOP-LEVEL APP ROUTER — application-area switch (UI-only)
// ══════════════════════════════════════════════════════
function App(){
  const [area, setArea] = useState('recall'); // 'recall' | 'warranty' — not persisted
  // Lifted terminal state — persists across application-area switches (session only, not localStorage)
  const [recallTab, setRecallTab] = useState('dashboard');
  const [wmisTab,   setWmisTab]   = useState('wmis-overview');
  const areaSwitch = <AreaSwitch area={area} setArea={setArea}/>;
  return area === 'warranty'
    ? <WarrantyMISApp areaSwitch={areaSwitch} wmisTab={wmisTab} setWmisTab={setWmisTab}/>
    : <RecallDatabaseApp areaSwitch={areaSwitch} tab={recallTab} setTab={setRecallTab}/>;
}

export default App;
