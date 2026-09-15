const SECTION_DEFS = [
  ['finalidade', 'FINALIDADE'],
  ['abrangencia', 'ABRANGÊNCIA'],
  ['executantes', 'EXECUTANTES'],
  ['descricao', 'DESCRIÇÃO'],
  ['orientacoes', 'ORIENTAÇÕES AOS PACIENTES CLIENTES E COLABORADORES'],
  ['riscos', 'RISCOS RELACIONADOS E AÇÕES PREVENTIVAS'],
  ['observacoes', 'OBSERVAÇÕES'],
  ['historico', 'HISTÓRICO DAS REVISÕES']
];

const aliases = {
  finalidade:['FINALIDADE','OBJETIVO'],
  abrangencia:['ABRANGENCIA','ABRANGÊNCIA','APLICACAO','APLICAÇÃO'],
  executantes:['EXECUTANTES','RESPONSAVEIS','RESPONSÁVEIS'],
  descricao:['DESCRICAO','DESCRIÇÃO','PROCEDIMENTO','DESENVOLVIMENTO'],
  orientacoes:['ORIENTACOES AOS PACIENTES CLIENTES E COLABORADORES','ORIENTAÇÕES AOS PACIENTES CLIENTES E COLABORADORES','ORIENTACOES AOS PACIENTES CLIENTES','ORIENTAÇÕES AOS PACIENTES CLIENTES'],
  riscos:['RISCOS RELACIONADOS E ACOES PREVENTIVAS','RISCOS RELACIONADOS E AÇÕES PREVENTIVAS','RISCOS E ACOES PREVENTIVAS','RISCOS E AÇÕES PREVENTIVAS'],
  observacoes:['OBSERVACOES','OBSERVAÇÕES'],
  historico:['HISTORICO DAS REVISOES','HISTÓRICO DAS REVISÕES']
};

const $ = (id) => document.getElementById(id);
const fileInput=$('fileInput'), dropzone=$('dropzone'), reviewPanel=$('reviewPanel'), metadataPanel=$('metadataPanel'), exportPanel=$('exportPanel'), sectionsEl=$('sections');
let currentFileName='POP';

function normalize(value){return value.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/^\s*\d+(?:\.\d+)*\s*[-.)]?\s*/,'').replace(/[^A-Z0-9 ]/gi,' ').replace(/\s+/g,' ').trim().toUpperCase()}
function detectSection(text){const n=normalize(text);return SECTION_DEFS.find(([key])=>aliases[key].some(a=>n===normalize(a)))?.[0]||null}
function escapeHtml(v){return v.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]))}
function toast(message){const el=$('toast');el.textContent=message;el.classList.add('show');setTimeout(()=>el.classList.remove('show'),2800)}

async function parseDocx(file){
  if(!window.JSZip) throw new Error('Não foi possível carregar o leitor de DOCX. Verifique a internet e tente novamente.');
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const xmlFile=zip.file('word/document.xml');
  if(!xmlFile) throw new Error('Este arquivo não parece ser um DOCX válido.');
  const xml=new DOMParser().parseFromString(await xmlFile.async('string'),'application/xml');
  const body=xml.getElementsByTagNameNS('*','body')[0];
  const items=[];
  for(const node of body.children){
    if(node.localName==='p'){
      const runParts=[...node.getElementsByTagNameNS('*','t')].map(t=>t.textContent.trim()).filter(Boolean);
      const runTable=inferRunTable(runParts);
      if(runTable){items.push(runTable);continue}
      let text='';
      const walker=xml.createTreeWalker(node,NodeFilter.SHOW_ELEMENT);
      while(walker.nextNode()){
        const part=walker.currentNode;
        if(part.localName==='t') text+=part.textContent;
        else if(part.localName==='tab') text+='\t';
        else if(part.localName==='br'||part.localName==='cr') text+='\n';
      }
      text=text.replace(/[ ]+\n/g,'\n').replace(/\n[ ]+/g,'\n').trim();
      if(text) items.push(inferStructuredItem(text));
    } else if(node.localName==='tbl'){
      const rows=[...node.children].filter(n=>n.localName==='tr').map(tr=>[...tr.children].filter(n=>n.localName==='tc').map(tc=>[...tc.getElementsByTagNameNS('*','t')].map(t=>t.textContent).join(' ').trim()));
      if(rows.length) items.push({type:'table',rows});
    }
  }
  const result=Object.fromEntries(SECTION_DEFS.map(([k])=>[k,[]]));
  const unassigned=[]; let active=null;
  for(const item of items){
    if(item.type==='p'){
      const found=detectSection(item.text);
      if(found){active=found;continue}
    }
    if(active) result[active].push(item); else unassigned.push(item);
  }
  return {result,unassigned};
}

function inferRunTable(parts){
  const codeIndexes=parts.map((part,index)=>/^\d{2,6}$/.test(part)?index:-1).filter(index=>index>=0);
  if(codeIndexes.length<3) return null;
  const first=codeIndexes[0], rows=[], header=parts.slice(0,first).join(' ').trim();
  if(header) rows.push([header,header]);
  for(let i=0;i<codeIndexes.length;i++){
    const codeIndex=codeIndexes[i], end=i+1<codeIndexes.length?codeIndexes[i+1]:parts.length;
    const description=parts.slice(codeIndex+1,end).join(' ').trim();
    if(description) rows.push([parts[codeIndex],description]);
  }
  return rows.length>=3?{type:'table',rows}:null;
}

function inferStructuredItem(text){
  const lines=text.split(/\n+/).map(x=>x.trim()).filter(Boolean);
  if(text.includes('\t')){
    const rows=lines.map(line=>line.split(/\t+/).map(x=>x.trim()).filter(Boolean)).filter(row=>row.length>1);
    if(rows.length) return {type:'table',rows};
  }
  // Alguns POPs antigos simulam tabelas com tabulações ou caixas de texto.
  // Após a extração, elas podem aparecer como: TÍTULO977MEDICAMENTO1040MEDICAMENTO.
  const matches=[...text.matchAll(/\b(\d{3,6})\s*(?=[A-ZÀ-Ü])/g)];
  if(matches.length>=3){
    const first=matches[0].index, header=text.slice(0,first).trim();
    const rows=header?[[header,header]]:[];
    matches.forEach((match,index)=>{
      const start=match.index+match[0].length;
      const end=index+1<matches.length?matches[index+1].index:text.length;
      rows.push([match[1],text.slice(start,end).trim()]);
    });
    if(rows.filter(r=>r[1]).length>=3) return {type:'table',rows};
  }
  return {type:'p',text};
}

function itemsToHtml(items){return items.map(item=>{
  if(item.type==='p') return `<p>${escapeHtml(item.text)}</p>`;
  return `<table><tbody>${item.rows.map((row,ri)=>`<tr>${row.map(cell=>`<${ri?'td':'th'}>${escapeHtml(cell)}</${ri?'td':'th'}>`).join('')}</tr>`).join('')}</tbody></table>`;
}).join('')}

function renderReview(data){
  sectionsEl.innerHTML='';
  SECTION_DEFS.forEach(([key,title],i)=>{
    let html=itemsToHtml(data.result[key]);
    if(key==='historico' && !data.result[key].some(item=>item.type==='table')){
      html='<table><tbody><tr><th>REVISÃO</th><th>NOME DO REVISOR</th><th>DESCRIÇÃO SUMÁRIA</th><th>DATA</th></tr><tr><td>Versão 00</td><td></td><td></td><td></td></tr></tbody></table>';
    }
    const card=document.createElement('article');card.className='section-card';
    card.innerHTML=`<div class="section-title"><span class="section-number">${i+1}</span><strong>${title}</strong></div><div class="editor" id="editor-${key}" contenteditable="true" spellcheck="true">${html}</div>`;
    sectionsEl.appendChild(card);
  });
  const missing=SECTION_DEFS.filter(([key])=>!data.result[key].length).map(([,title])=>title);
  const messages=[];
  if(missing.length) messages.push(`Revise as seções sem conteúdo identificado: ${missing.join(', ')}.`);
  if(data.unassigned.length) messages.push(`${data.unassigned.length} elemento(s) antes da primeira seção não foram incluídos; normalmente são capa, cabeçalho ou sumário.`);
  $('warnings').hidden=!messages.length;$('warnings').innerHTML=messages.map(m=>`<div>${escapeHtml(m)}</div>`).join('');
  reviewPanel.hidden=false;metadataPanel.hidden=false;exportPanel.hidden=false;reviewPanel.scrollIntoView({behavior:'smooth',block:'start'});
}

async function handleFile(file){
  if(!file||!file.name.toLowerCase().endsWith('.docx')){toast('Selecione um arquivo no formato DOCX.');return}
  currentFileName=file.name.replace(/\.docx$/i,'');$('metaTitle').value=currentFileName.replace(/^POP\s*[-–]?\s*\d+\s*[-–]?\s*/i,'').trim();$('fileStatus').hidden=false;$('fileStatus').textContent=`Lendo ${file.name}…`;
  try{const data=await parseDocx(file);$('fileStatus').textContent=`Arquivo carregado: ${file.name}`;renderReview(data)}
  catch(err){$('fileStatus').textContent='Não foi possível ler o arquivo.';toast(err.message);console.error(err)}
}

function htmlToDocxChildren(editor){
  const D=window.docx, children=[];
  for(const node of editor.children){
    if(node.tagName==='TABLE'){
      const rows=[...node.rows].map((row,ri)=>new D.TableRow({tableHeader:ri===0,children:[...row.cells].map(cell=>new D.TableCell({shading:ri===0?{fill:'D9D9D9'}:undefined,verticalAlign:D.VerticalAlign.CENTER,children:[new D.Paragraph({alignment:D.AlignmentType.CENTER,children:[new D.TextRun({text:cell.innerText.trim(),font:'Arial',size:18,bold:ri===0})]})]}))}));
      children.push(new D.Table({alignment:D.AlignmentType.CENTER,width:{size:100,type:D.WidthType.PERCENTAGE},borders:{top:{style:D.BorderStyle.SINGLE,size:1,color:'808080'},bottom:{style:D.BorderStyle.SINGLE,size:1,color:'808080'},left:{style:D.BorderStyle.SINGLE,size:1,color:'808080'},right:{style:D.BorderStyle.SINGLE,size:1,color:'808080'},insideHorizontal:{style:D.BorderStyle.SINGLE,size:1,color:'A0A0A0'},insideVertical:{style:D.BorderStyle.SINGLE,size:1,color:'A0A0A0'}},rows}));
      children.push(new D.Paragraph({spacing:{after:80}}));
    }else{
      const text=node.innerText.trim();if(text) children.push(new D.Paragraph({spacing:{after:120,line:276},children:[new D.TextRun({text,font:'Arial',size:24})]}));
    }
  }
  if(!children.length){children.push(new D.Paragraph({children:[new D.TextRun({text:'NA (Não se aplica).',font:'Arial',size:24})]}))}
  return children;
}

function value(id,fallback=''){return $(id).value.trim()||fallback}
function brDate(raw){if(!raw)return 'DATA';const [y,m,d]=raw.split('-');return `${d}/${m}/${y}`}
function cellText(D,text,{bold=false,size=16,align=D.AlignmentType.LEFT}={}){
  return new D.Paragraph({alignment:align,spacing:{before:0,after:0},children:[new D.TextRun({text,font:'Arial',size,bold,color:'000000'})]});
}
async function imageBytes(path){return new Uint8Array(await (await fetch(path)).arrayBuffer())}

async function buildHeader(D){
  const logo=await imageBytes('assets/imip-logo.png');
  const border={style:D.BorderStyle.SINGLE,size:1,color:'808080'};
  const title=value('metaTitle','TÍTULO DO DOCUMENTO').toUpperCase();
  const process=value('metaProcess','NOME DO PROCESSO').toUpperCase();
  const code=value('metaCode','IMIP.POP.XXX.XXX').toUpperCase();
  const version=value('metaVersion','00');
  return new D.Header({children:[
    new D.Table({width:{size:100,type:D.WidthType.PERCENTAGE},columnWidths:[1200,7000,2100],borders:{top:border,bottom:border,left:border,right:border,insideHorizontal:border,insideVertical:border},rows:[
      new D.TableRow({children:[
        new D.TableCell({rowSpan:2,verticalAlign:D.VerticalAlign.CENTER,width:{size:1200,type:D.WidthType.DXA},children:[new D.Paragraph({alignment:D.AlignmentType.CENTER,children:[new D.ImageRun({data:logo,transformation:{width:38,height:27},type:'png'})]})]}),
        new D.TableCell({verticalAlign:D.VerticalAlign.CENTER,width:{size:7000,type:D.WidthType.DXA},children:[cellText(D,title,{bold:true,size:16,align:D.AlignmentType.CENTER})]}),
        new D.TableCell({verticalAlign:D.VerticalAlign.CENTER,width:{size:2100,type:D.WidthType.DXA},children:[cellText(D,`Versão: ${version}`,{size:14,align:D.AlignmentType.CENTER})]})
      ]}),
      new D.TableRow({children:[
        new D.TableCell({verticalAlign:D.VerticalAlign.CENTER,width:{size:7000,type:D.WidthType.DXA},children:[cellText(D,`Processo: ${process}`,{bold:true,size:13,align:D.AlignmentType.CENTER})]}),
        new D.TableCell({verticalAlign:D.VerticalAlign.CENTER,width:{size:2100,type:D.WidthType.DXA},children:[cellText(D,`Código: ${code}`,{bold:true,size:13,align:D.AlignmentType.CENTER})]})
      ]})
    ]})
  ]});
}

async function buildFooter(D){
  const nature=await imageBytes('assets/footer-natureza.png');
  const border={style:D.BorderStyle.SINGLE,size:1,color:'808080'};
  const footerTable=new D.Table({width:{size:100,type:D.WidthType.PERCENTAGE},columnWidths:[5150,5150],borders:{top:border,bottom:border,left:border,right:border,insideHorizontal:border,insideVertical:border},rows:[
    new D.TableRow({children:[
      new D.TableCell({children:[cellText(D,`Elaboração: ${value('metaAuthor','NOME')}`,{size:13})]}),
      new D.TableCell({children:[cellText(D,`Data da Elaboração: ${brDate($('metaAuthorDate').value)}`,{size:13})]})
    ]}),
    new D.TableRow({children:[
      new D.TableCell({children:[cellText(D,`Aprovação: ${value('metaApproval','NOME')}`,{size:13})]}),
      new D.TableCell({children:[cellText(D,`Data da Aprovação: ${brDate($('metaApprovalDate').value)}`,{size:13})]})
    ]}),
    new D.TableRow({children:[
      new D.TableCell({children:[cellText(D,`Vigência: ${value('metaValidity','2 anos')}`,{size:13})]}),
      new D.TableCell({children:[new D.Paragraph({alignment:D.AlignmentType.RIGHT,spacing:{before:0,after:0},children:[new D.TextRun({text:'Página ',font:'Arial',size:13}),new D.TextRun({children:[D.PageNumber.CURRENT],font:'Arial',size:13}),new D.TextRun({text:' de ',font:'Arial',size:13}),new D.TextRun({children:[D.PageNumber.TOTAL_PAGES],font:'Arial',size:13})]})]})
    ]})
  ]});
  return new D.Footer({children:[footerTable,new D.Paragraph({alignment:D.AlignmentType.CENTER,spacing:{before:40,after:0},children:[new D.ImageRun({data:nature,transformation:{width:575,height:49},type:'png'})]})]});
}

async function generateDocx(){
  if(!window.docx){toast('Não foi possível carregar o gerador de Word. Verifique a internet.');return}
  const button=$('generateBtn');button.disabled=true;button.textContent='Gerando…';
  try{
    const D=window.docx, children=[];
    children.push(new D.Paragraph({spacing:{after:220},children:[new D.TextRun({text:'SUMÁRIO',font:'Arial',size:24,bold:true,color:'000000'})]}));
    children.push(new D.TableOfContents('Sumário',{hyperlink:true,headingStyleRange:'1-1',pageNumbersEntryLevelsRange:'1-1'}));
    children.push(new D.Paragraph({pageBreakBefore:true,children:[]}));
    SECTION_DEFS.forEach(([key,title],i)=>{
      children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_1,keepNext:true,spacing:{before:200,after:100},children:[new D.TextRun({text:`${i+1}. ${title}`,font:'Arial',size:24,bold:true,color:'000000'})]}));
      children.push(...htmlToDocxChildren($(`editor-${key}`)));
    });
    const [header,footer]=await Promise.all([buildHeader(D),buildFooter(D)]);
    const doc=new D.Document({
      title:value('metaTitle','POP padronizado'),creator:'POP Zero',features:{updateFields:true},
      styles:{default:{document:{run:{font:'Arial',size:24,color:'000000'},paragraph:{spacing:{line:276}}}},paragraphStyles:[{id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{font:'Arial',size:24,bold:true,color:'000000'},paragraph:{spacing:{before:200,after:100},outlineLevel:0}}]},
      sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1800,right:1134,bottom:650,left:1276,header:320,footer:300}}},headers:{default:header},footers:{default:footer},children}]
    });
    const blob=await D.Packer.toBlob(doc), url=URL.createObjectURL(blob), a=document.createElement('a');
    a.href=url;a.download=`${currentFileName}_Norma_Zero.docx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);toast('POP completo gerado com sucesso.');
  }catch(err){console.error(err);toast('Ocorreu um erro ao gerar o Word.')}finally{button.disabled=false;button.textContent='Baixar POP completo'}
}

fileInput.addEventListener('change',e=>handleFile(e.target.files[0]));
['dragenter','dragover'].forEach(name=>dropzone.addEventListener(name,e=>{e.preventDefault();dropzone.classList.add('drag')}));
['dragleave','drop'].forEach(name=>dropzone.addEventListener(name,e=>{e.preventDefault();dropzone.classList.remove('drag')}));
dropzone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));
$('clearBtn').addEventListener('click',()=>{fileInput.value='';reviewPanel.hidden=true;metadataPanel.hidden=true;exportPanel.hidden=true;$('fileStatus').hidden=true;scrollTo({top:0,behavior:'smooth'})});
$('generateBtn').addEventListener('click',generateDocx);
