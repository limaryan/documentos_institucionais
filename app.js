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
const fileInput=$('fileInput'), dropzone=$('dropzone'), reviewPanel=$('reviewPanel'), exportPanel=$('exportPanel'), sectionsEl=$('sections');
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
      const text=[...node.getElementsByTagNameNS('*','t')].map(t=>t.textContent).join('').trim();
      if(text) items.push({type:'p',text});
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

function itemsToHtml(items){return items.map(item=>{
  if(item.type==='p') return `<p>${escapeHtml(item.text)}</p>`;
  return `<table><tbody>${item.rows.map((row,ri)=>`<tr>${row.map(cell=>`<${ri?'td':'th'}>${escapeHtml(cell)}</${ri?'td':'th'}>`).join('')}</tr>`).join('')}</tbody></table>`;
}).join('')}

function renderReview(data){
  sectionsEl.innerHTML='';
  SECTION_DEFS.forEach(([key,title],i)=>{
    const card=document.createElement('article');card.className='section-card';
    card.innerHTML=`<div class="section-title"><span class="section-number">${i+1}</span><strong>${title}</strong></div><div class="editor" id="editor-${key}" contenteditable="true" spellcheck="true">${itemsToHtml(data.result[key])}</div>`;
    sectionsEl.appendChild(card);
  });
  const missing=SECTION_DEFS.filter(([key])=>!data.result[key].length).map(([,title])=>title);
  const messages=[];
  if(missing.length) messages.push(`Revise as seções sem conteúdo identificado: ${missing.join(', ')}.`);
  if(data.unassigned.length) messages.push(`${data.unassigned.length} elemento(s) antes da primeira seção não foram incluídos; normalmente são capa, cabeçalho ou sumário.`);
  $('warnings').hidden=!messages.length;$('warnings').innerHTML=messages.map(m=>`<div>${escapeHtml(m)}</div>`).join('');
  reviewPanel.hidden=false;exportPanel.hidden=false;reviewPanel.scrollIntoView({behavior:'smooth',block:'start'});
}

async function handleFile(file){
  if(!file||!file.name.toLowerCase().endsWith('.docx')){toast('Selecione um arquivo no formato DOCX.');return}
  currentFileName=file.name.replace(/\.docx$/i,'');$('fileStatus').hidden=false;$('fileStatus').textContent=`Lendo ${file.name}…`;
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

async function generateDocx(){
  if(!window.docx){toast('Não foi possível carregar o gerador de Word. Verifique a internet.');return}
  const button=$('generateBtn');button.disabled=true;button.textContent='Gerando…';
  try{
    const D=window.docx, children=[];
    SECTION_DEFS.forEach(([key,title],i)=>{
      children.push(new D.Paragraph({keepNext:true,spacing:{before:200,after:100},children:[new D.TextRun({text:`${i+1}. ${title}`,font:'Arial',size:24,bold:true})]}));
      children.push(...htmlToDocxChildren($(`editor-${key}`)));
    });
    const doc=new D.Document({styles:{default:{document:{run:{font:'Arial',size:24},paragraph:{spacing:{line:276}}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,right:1134,bottom:1134,left:1134}}},children}]});
    const blob=await D.Packer.toBlob(doc), url=URL.createObjectURL(blob), a=document.createElement('a');
    a.href=url;a.download=`${currentFileName}_Conteudo_Norma_Zero.docx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);toast('Word gerado com sucesso.');
  }catch(err){console.error(err);toast('Ocorreu um erro ao gerar o Word.')}finally{button.disabled=false;button.textContent='Baixar Word padronizado'}
}

fileInput.addEventListener('change',e=>handleFile(e.target.files[0]));
['dragenter','dragover'].forEach(name=>dropzone.addEventListener(name,e=>{e.preventDefault();dropzone.classList.add('drag')}));
['dragleave','drop'].forEach(name=>dropzone.addEventListener(name,e=>{e.preventDefault();dropzone.classList.remove('drag')}));
dropzone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));
$('clearBtn').addEventListener('click',()=>{fileInput.value='';reviewPanel.hidden=true;exportPanel.hidden=true;$('fileStatus').hidden=true;scrollTo({top:0,behavior:'smooth'})});
$('generateBtn').addEventListener('click',generateDocx);
