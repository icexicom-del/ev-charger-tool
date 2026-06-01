/**
 * Vercel Serverless Function — Word Document Export
 * POST /api/export-word  body: quote JSON
 * Returns: .docx binary
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  Header, Footer, HeadingLevel, TabStopType, SimpleField, PageBreak
} = require('docx');

module.exports = async (req, res) => {
  if (req.method !== 'POST') { res.status(405).end('Method not allowed'); return; }

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(204).end(); return; }

  try {
    const q = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    const buf = await buildDoc(q);
    const fname = `Quotation_${q.id}_${(q.customer||'Customer').replace(/\s+/g,'_')}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('Content-Length', buf.length);
    res.status(200).end(buf);
  } catch(e) {
    console.error('export-word error:', e);
    res.status(500).json({ error: e.message });
  }
};

// ── Helpers ───────────────────────────────────────────────────────
const INR = v => '\u20B9' + Math.round(v).toLocaleString('en-IN');
const nb  = { style: BorderStyle.NIL, size: 0, color: 'FFFFFF' };
const tb  = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E0' };
const hb  = { style: BorderStyle.SINGLE, size: 4, color: '1A3A5C' };
const allB = b => ({ top:b, bottom:b, left:b, right:b });
const nbAll = allB(nb);
const cm  = { top:80, bottom:80, left:120, right:120 };
const PW  = 9638; // A4 2cm margins content width

function hdrCell(txt, w, centre) {
  return new TableCell({ width:{size:w,type:WidthType.DXA}, borders:allB(hb),
    shading:{fill:'1A3A5C',type:ShadingType.CLEAR}, margins:cm, verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({ alignment: centre?AlignmentType.CENTER:AlignmentType.LEFT,
      children:[new TextRun({text:txt,bold:true,color:'FFFFFF',size:18,font:'Calibri'})] })] });
}
function bdCell(txt, w, opts={}) {
  return new TableCell({ width:{size:w,type:WidthType.DXA}, borders:allB(tb),
    shading: opts.shade?{fill:opts.shade,type:ShadingType.CLEAR}:undefined,
    margins:cm, verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({
      alignment: opts.right?AlignmentType.RIGHT:opts.centre?AlignmentType.CENTER:AlignmentType.LEFT,
      children:[new TextRun({text:String(txt),bold:!!opts.bold,size:18,font:'Calibri',color:opts.color||'1A202C'})] })] });
}
function totCell(txt, w, dark) {
  const fill = dark ? '1A3A5C' : 'EBF4FF';
  const col  = dark ? 'FFFFFF' : '1A3A5C';
  return new TableCell({ width:{size:w,type:WidthType.DXA},
    borders: allB(dark?{style:BorderStyle.SINGLE,size:8,color:'1A3A5C'}:tb),
    shading:{fill,type:ShadingType.CLEAR}, margins:cm,
    children:[new Paragraph({alignment:AlignmentType.RIGHT,
      children:[new TextRun({text:String(txt),bold:true,size:dark?20:18,font:'Calibri',color:col})] })] });
}
function p(txt,opts={}) {
  return new Paragraph({ alignment:opts.centre?AlignmentType.CENTER:opts.right?AlignmentType.RIGHT:AlignmentType.LEFT,
    spacing:{before:opts.before||80,after:opts.after||80},
    children:[new TextRun({text:txt,bold:!!opts.bold,size:opts.size||20,font:'Calibri',
      color:opts.color||'1A202C',italics:!!opts.italic,underline:opts.ul?{}:undefined})] });
}
function divP(color='2E6DB4') {
  return new Paragraph({spacing:{before:60,after:60},border:{bottom:{style:BorderStyle.SINGLE,size:8,color}}});
}
const ep = (h=80) => new Paragraph({spacing:{before:h,after:h},children:[new TextRun('')]});

async function buildDoc(q) {
  const today = new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'long',year:'numeric'});
  const appr  = (q.history||[]).filter(h=>h.status==='Approved').pop();

  // col widths for main table
  const cw = [
    Math.round(PW*0.05), Math.round(PW*0.41), Math.round(PW*0.09),
    Math.round(PW*0.08), Math.round(PW*0.16), Math.round(PW*0.08), Math.round(PW*0.13)
  ];

  const makeRows = (items, startI) => items.map((it,i) => {
    const idx = startI + i;
    const sh  = idx%2===0 ? 'F7FAFC' : 'FFFFFF';
    const qty = it.qty ? (String(it.qty)+(it.unit==='Metre'?'m':'')) : '1';
    const rate = it.rate ? INR(it.rate)+(it.unit==='Metre'?'/m':'') : '—';
    const amt  = it.amt || ((it.qty||1)*(it.rate||0));
    return new TableRow({ children:[
      bdCell(idx+1, cw[0], {centre:true, shade:sh}),
      bdCell(it.name, cw[1], {shade:sh}),
      bdCell(it.unit||'Unit', cw[2], {centre:true, shade:sh}),
      bdCell(qty, cw[3], {centre:true, shade:sh}),
      bdCell(rate, cw[4], {right:true, shade:sh}),
      bdCell('18%', cw[5], {centre:true, shade:sh}),
      bdCell(INR(amt), cw[6], {right:true, shade:sh}),
    ]});
  });

  const aRows = makeRows(q.bomSnapshot||[], 0);
  const bRows = makeRows(q.cableSnapshot||[], aRows.length);
  const cRows = makeRows(q.optionalsSnapshot||[], aRows.length+bRows.length);

  const totRows = [
    new TableRow({children:[
      new TableCell({columnSpan:6,width:{size:cw[0]+cw[1]+cw[2]+cw[3]+cw[4]+cw[5],type:WidthType.DXA},
        borders:allB(tb),shading:{fill:'EBF4FF',type:ShadingType.CLEAR},margins:cm,
        children:[new Paragraph({alignment:AlignmentType.RIGHT,
          children:[new TextRun({text:'Sub Total (excl. GST)',bold:true,size:18,font:'Calibri',color:'1A3A5C'})]})]}),
      totCell(INR(q.sub||0), cw[6], false)
    ]}),
    new TableRow({children:[
      new TableCell({columnSpan:6,width:{size:cw[0]+cw[1]+cw[2]+cw[3]+cw[4]+cw[5],type:WidthType.DXA},
        borders:allB(tb),shading:{fill:'F7FAFC',type:ShadingType.CLEAR},margins:cm,
        children:[new Paragraph({alignment:AlignmentType.RIGHT,
          children:[new TextRun({text:'GST @ 18%',size:18,font:'Calibri',color:'4A5568'})]})]}),
      bdCell(INR(q.gst||0), cw[6], {right:true, shade:'F7FAFC', color:'4A5568'})
    ]}),
    new TableRow({children:[
      new TableCell({columnSpan:6,width:{size:cw[0]+cw[1]+cw[2]+cw[3]+cw[4]+cw[5],type:WidthType.DXA},
        borders:allB({style:BorderStyle.SINGLE,size:8,color:'1A3A5C'}),
        shading:{fill:'1A3A5C',type:ShadingType.CLEAR},margins:cm,
        children:[new Paragraph({alignment:AlignmentType.RIGHT,
          children:[new TextRun({text:'GRAND TOTAL (incl. GST)',bold:true,size:20,font:'Calibri',color:'FFFFFF'})]})]}),
      totCell(INR(q.total||0), cw[6], true)
    ]}),
  ];

  function tcSec(title, body) {
    return [
      new Paragraph({spacing:{before:140,after:60},children:[new TextRun({text:title,bold:true,size:20,font:'Calibri',color:'1A3A5C',underline:{}})]}),
      new Paragraph({spacing:{before:40,after:80},children:[new TextRun({text:body,size:18,font:'Calibri',color:'2D3748'})]}),
    ];
  }

  const doc = new Document({
    styles:{default:{document:{run:{font:'Calibri',size:20}}}},
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
      headers:{default: new Header({children:[
        new Table({width:{size:PW,type:WidthType.DXA},columnWidths:[Math.round(PW*0.55),Math.round(PW*0.45)],
          borders:{top:nb,bottom:{style:BorderStyle.SINGLE,size:6,color:'2E6DB4'},left:nb,right:nb,insideH:nb,insideV:nb},
          rows:[new TableRow({children:[
            new TableCell({width:{size:Math.round(PW*0.55),type:WidthType.DXA},borders:nbAll,margins:{top:60,bottom:60,left:0,right:120},children:[
              new Paragraph({children:[new TextRun({text:'EV CHARGER INSTALLATION SERVICES',bold:true,size:28,font:'Calibri',color:'1A3A5C'})]}),
              new Paragraph({children:[new TextRun({text:'Professional EV Charging Solutions',size:18,font:'Calibri',color:'4A5568',italics:true})]}),
            ]}),
            new TableCell({width:{size:Math.round(PW*0.45),type:WidthType.DXA},borders:nbAll,margins:{top:60,bottom:60,left:120,right:0},verticalAlign:VerticalAlign.CENTER,children:[
              new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:'QUOTATION',bold:true,size:36,font:'Calibri',color:'2E6DB4'})]}),
            ]}),
          ]})]})
      ]})},
      footers:{default: new Footer({children:[
        new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:6,color:'2E6DB4'}},spacing:{before:80,after:40},
          tabStops:[{type:TabStopType.RIGHT,position:PW}],
          children:[
            new TextRun({text:'EV Charger Installation Services  |  evse@company.in',size:16,font:'Calibri',color:'4A5568'}),
            new TextRun({text:'\tPage ',size:16,font:'Calibri',color:'4A5568'}),
            new SimpleField({instruction:'PAGE',cachedValue:'1'}),
          ]})
      ]})},
      children:[
        // Ref / Date
        new Table({width:{size:PW,type:WidthType.DXA},columnWidths:[Math.round(PW*0.5),Math.round(PW*0.5)],
          borders:{top:nb,bottom:nb,left:nb,right:nb,insideH:nb,insideV:nb},
          rows:[new TableRow({children:[
            new TableCell({width:{size:Math.round(PW*0.5),type:WidthType.DXA},borders:nbAll,margins:{top:0,bottom:0,left:0,right:0},children:[new Paragraph({children:[new TextRun({text:`Ref: ${q.id}`,size:18,font:'Calibri',color:'4A5568'})]})]}),
            new TableCell({width:{size:Math.round(PW*0.5),type:WidthType.DXA},borders:nbAll,margins:{top:0,bottom:0,left:0,right:0},children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:`Date: ${q.quote_date||q.date||today}`,size:18,font:'Calibri',color:'4A5568'})]})]}),
          ]})]
        }),
        // Addressee
        ep(120),
        p('To,',{size:18}),
        p(q.customer,{bold:true,size:22}),
        ...(q.addr&&q.addr!=='—'?[p(q.addr,{size:18})]:[]),
        ...(q.phone?[p(q.phone,{size:18})]:[]),
        ...(q.email?[p(q.email,{size:18,color:'2E6DB4'})]:[]),
        ep(80),
        // Subject
        new Paragraph({spacing:{before:80,after:120},children:[
          new TextRun({text:'Subject: ',bold:true,size:20,font:'Calibri'}),
          new TextRun({text:`Quotation for EV Charger Installation — ${q.model}${q.site_desc||q.siteDesc?' at '+(q.site_desc||q.siteDesc):''}`,size:20,font:'Calibri',underline:{}}),
        ]}),
        divP(),ep(80),
        // Intro
        p('This is with reference to your requirement of EV charger installation. We thank you for providing us an opportunity to associate with your esteemed organization. Please find below our commercial offer for your acceptance.',{size:19}),
        p('We hope this offer is in line with your expectations. In case of any query, please feel free to contact us. We assure you the best of our products and services.',{size:19,after:160}),
        // Commercials label
        new Paragraph({spacing:{before:200,after:80},children:[new TextRun({text:'Commercials',bold:true,size:24,font:'Calibri',color:'1A3A5C',underline:{}})]}),
        // Main table
        new Table({width:{size:PW,type:WidthType.DXA},columnWidths:cw,rows:[
          new TableRow({tableHeader:true,children:[
            hdrCell('S.No',cw[0],true), hdrCell('ITEMS',cw[1],false), hdrCell('Unit',cw[2],true),
            hdrCell('Qty',cw[3],true), hdrCell('Price / Unit',cw[4],true), hdrCell('Tax',cw[5],true), hdrCell('Amount',cw[6],true)
          ]}),
          ...aRows, ...bRows, ...cRows, ...totRows
        ]}),
        // Approval
        ep(120),
        new Table({width:{size:PW,type:WidthType.DXA},columnWidths:[PW],rows:[new TableRow({children:[
          new TableCell({width:{size:PW,type:WidthType.DXA},borders:allB({style:BorderStyle.SINGLE,size:4,color:'9AE6B4'}),
            shading:{fill:'F0FFF4',type:ShadingType.CLEAR},margins:{top:100,bottom:100,left:150,right:150},children:[
              new Paragraph({children:[new TextRun({text:'\u2714 APPROVED QUOTATION',bold:true,size:20,font:'Calibri',color:'276749'})]}),
              new Paragraph({children:[new TextRun({text:`Approved by: ${appr?.by||'—'}   |   Date: ${appr?.date||'—'}${appr?.note?`   |   "${appr.note}"`:''}`,size:18,font:'Calibri',color:'2D6A4F'})]})
            ]}),
        ]})]})  ,
        // T&C
        ep(200), divP(),
        new Paragraph({spacing:{before:160,after:100},children:[new TextRun({text:'Terms and Conditions',bold:true,size:24,font:'Calibri',color:'1A3A5C',underline:{}})]}),
        new Paragraph({spacing:{before:60,after:60},indent:{left:360},children:[new TextRun({text:'Warranty: ',bold:true,size:19,font:'Calibri'}),new TextRun({text:'12 months on installation workmanship from date of commissioning.',size:19,font:'Calibri',color:'2D3748'})]}),
        new Paragraph({spacing:{before:60,after:60},indent:{left:360},children:[new TextRun({text:'Lead Time: ',bold:true,size:19,font:'Calibri'}),new TextRun({text:'Work commences within 5–7 working days of advance payment and site readiness.',size:19,font:'Calibri',color:'2D3748'})]}),
        new Paragraph({spacing:{before:60,after:60},indent:{left:360},children:[new TextRun({text:'Scope of I&C – Inclusion: ',bold:true,size:19,font:'Calibri'}),new TextRun({text:'Cable termination, commissioning, CMS onboarding (if applicable).',size:19,font:'Calibri',color:'2D3748'})]}),
        new Paragraph({spacing:{before:60,after:60},indent:{left:360},children:[new TextRun({text:'Scope of I&C – Exclusions: ',bold:true,size:19,font:'Calibri'}),new TextRun({text:'Civil work, DISCOM charges, earthing (unless quoted), canopy, charger stand, cable beyond quoted run.',size:19,font:'Calibri',color:'2D3748'})]}),
        ...tcSec('Terms of Payment','100% advance on order confirmation before commencement of work.'),
        ...tcSec('Delivery','Work scheduled within 5–7 working days from receipt of PO and advance payment. Subject to force majeure.'),
        ...tcSec('Validity','This offer is valid for 30 days from the date of issue.'),
        ...tcSec('General','This proposal is based on mutual discussions. We reserve rights to revise if specifications or site conditions change.'),
        // Signatures
        ep(200), divP(),ep(80),
        new Table({width:{size:PW,type:WidthType.DXA},columnWidths:[Math.round(PW*0.45),Math.round(PW*0.1),Math.round(PW*0.45)],
          borders:{top:nb,bottom:nb,left:nb,right:nb,insideH:nb,insideV:nb},
          rows:[new TableRow({children:[
            new TableCell({width:{size:Math.round(PW*0.45),type:WidthType.DXA},borders:nbAll,children:[
              p('Customer acceptance',{size:18,color:'4A5568'}),
              ep(200),
              new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'CBD5E0'}},spacing:{before:0,after:0},children:[new TextRun({text:'Signature & date',size:17,font:'Calibri',color:'A0AEC0',italics:true})]}),
            ]}),
            new TableCell({width:{size:Math.round(PW*0.1),type:WidthType.DXA},borders:nbAll,children:[ep()]}),
            new TableCell({width:{size:Math.round(PW*0.45),type:WidthType.DXA},borders:nbAll,children:[
              new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:80,after:40},children:[new TextRun({text:'For EV Charger Installation Services',bold:true,size:18,font:'Calibri',color:'1A3A5C'})]}),
              new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:40,after:80},children:[new TextRun({text:q.prepared_by||q.by||'',size:20,font:'Calibri',bold:true})]}),
              ep(200),
              new Paragraph({alignment:AlignmentType.RIGHT,border:{top:{style:BorderStyle.SINGLE,size:4,color:'CBD5E0'}},spacing:{before:0,after:0},children:[new TextRun({text:'Authorised Signatory',size:17,font:'Calibri',color:'A0AEC0',italics:true})]}),
            ]}),
          ]})]
        }),
        // Footer note
        ep(160),
        new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'2E6DB4'}},alignment:AlignmentType.CENTER,spacing:{before:80,after:0},children:[new TextRun({text:`Quote ${q.id}  |  Generated ${today}  |  Computer-generated document`,size:16,font:'Calibri',color:'718096'})]})
      ]
    }]
  });
  return Packer.toBuffer(doc);
}
