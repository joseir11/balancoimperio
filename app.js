const SUPABASE_URL="https://znivpoqcvxkgscqoszfj.supabase.co";
const SUPABASE_KEY="sb_publishable_-77CwES71Z8Js-zUoADACQ_WAPZDOk6";
const TOKEN_KEY="balanco_imperio_team_access";
let accessToken="",team=null,materials=[],statuses={},own=[],selected=null,pendingCode="",ownStatusFilter="TODOS",ownListCollapsed=false;
const $=id=>document.getElementById(id);
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&","<":"<",">":">",'"':""","'":"&#39;"}[c]));
function msg(t){$("toast").textContent=t;$("toast").classList.remove("hidden");setTimeout(()=>$("toast").classList.add("hidden"),3200)}
function badge(s){return '<span class="status-dot '+s+'" title="'+(s==="OK"?"OK":s==="DIVERGENTE"?"Divergente":"Aguardando")+'"></span>'}
let adminSession=null;
function showStart(section){$("start").classList.remove("hidden");$("app").classList.add("hidden");["createPanel","codePanel","adminLoginPanel","adminPanel"].forEach(id=>$(id).classList.add("hidden"));$(section==="create"?"createPanel":section==="code"?"codePanel":section==="adminLogin"?"adminLoginPanel":"adminPanel").classList.remove("hidden")}
async function adminApi(action,payload={}){const response=await fetch(SUPABASE_URL+"/functions/v1/manage-team",{method:"POST",headers:{"Content-Type":"application/json","apikey":SUPABASE_KEY,"Authorization":"Bearer "+adminSession.access_token},body:JSON.stringify({action,...payload})});const result=await response.json();if(!response.ok)throw new Error(result.error||"Falha ao administrar equipes.");return result}
let preparedShareFile=null,sharePreparation=false,shareRefreshTimer=null;
function downloadExcelFile(file){
 const url=URL.createObjectURL(file);const link=document.createElement("a");link.href=url;link.download=file.name;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function buildExcelFile(){
 if(!window.ExcelJS)throw new Error("Não foi possível carregar o gerador de Excel. Verifique sua conexão e tente novamente.");
 const data=await adminApi("export");const allTeams=data.teams||[],materials=data.materials||[],counts=data.counts||[];
 const teams=allTeams.filter(t=>counts.some(c=>c.team_id===t.id));
 if(!teams.length)throw new Error("Ainda não existem contagens registradas pelas equipes.");
 const byKey=new Map(counts.map(c=>[c.material_id+"|"+c.team_id,c.quantity]));
 const workbook=new ExcelJS.Workbook();
 workbook.creator="BALANÇO IMPÉRIO";workbook.created=new Date();workbook.modified=new Date();
 const sheet=workbook.addWorksheet("CONTAGEM",{views:[{state:"frozen",ySplit:1,showGridLines:false}],properties:{defaultRowHeight:21}});
 const headers=["CÓDIGO","DESCRIÇÃO","UNIDADE",...teams.map(t=>t.name),"CONTAGEM CONFERIDA"];
 sheet.addRow(headers);
 const header=sheet.getRow(1);header.height=30;
 header.eachCell(cell=>{cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF111C2D"}};cell.font={name:"Arial",size:10,bold:true,color:{argb:"FFFFFFFF"}};cell.alignment={vertical:"middle",horizontal:"center",wrapText:true};cell.border={bottom:{style:"medium",color:{argb:"FFF28B22"}}}});
 materials.forEach(m=>{
  const values=teams.map(t=>byKey.get(m.id+"|"+t.id));
  const complete=values.every(v=>v!==undefined&&v!==null&&v!=="");
  const equal=complete&&values.every(v=>Number(v)===Number(values[0]));
  const row=sheet.addRow([m.code??"",m.description??"",m.unit??"",...values.map(v=>v===undefined||v===null?"":Number(v)),equal?Number(values[0]):null]);
  row.height=22;
  row.eachCell({includeEmpty:true},(cell,col)=>{
   cell.font={name:"Arial",size:10,color:{argb:"FF26364A"}};
   cell.alignment={vertical:"middle",horizontal:col<=2?"left":col===3?"center":"right"};
   if(row.number%2===0)cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF4F7FB"}};
   cell.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
   if(col>=4&&cell.value!==null&&cell.value!=="")cell.numFmt="#,##0.###";
  });
  const finalCell=row.getCell(headers.length);
  finalCell.font={name:"Arial",size:10,bold:true,color:{argb:equal?"FF168653":complete?"FFCE4848":"FF9A741B"}};
  finalCell.fill={type:"pattern",pattern:"solid",fgColor:{argb:equal?"FFE7F6ED":complete?"FFFFF0EF":"FFFFF5DD"}};
  finalCell.alignment={vertical:"middle",horizontal:"right"};
 });
 sheet.columns=[{width:18},{width:52},{width:13},...teams.map(()=>({width:18})),{width:24}];
 sheet.autoFilter={from:{row:1,column:1},to:{row:Math.max(1,sheet.rowCount),column:headers.length}};
 sheet.pageSetup={orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0};
 sheet.headerFooter.oddHeader="&CBALANÇO IMPÉRIO";sheet.headerFooter.oddFooter="&P de &N";
 const buffer=await workbook.xlsx.writeBuffer();
 const filename="BALANCO_IMPERIO_CONTAGEM_"+new Date().toISOString().slice(0,10)+".xlsx";
 const mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
 return new File([new Blob([buffer],{type:mime})],filename,{type:mime});
}
async function prepareShareFile(showMessage=false){
 if(sharePreparation||!adminSession)return;
 sharePreparation=true;const button=$("shareCounts");button.disabled=true;button.title="Preparando planilha para compartilhar";
 try{
  const file=await buildExcelFile();preparedShareFile=file;
  button.title="Compartilhar Excel";button.setAttribute("aria-label","Compartilhar Excel");
  if(showMessage)msg("Planilha atualizada e pronta para compartilhar.");
 }catch(e){
  button.title="Não foi possível preparar a planilha";
  if(showMessage)msg(e.message||"Não foi possível preparar a planilha.");
 }finally{
  sharePreparation=false;button.disabled=false;
 }
}
async function shareExcelFile(){
 const file=preparedShareFile;
 if(!file){msg("A planilha ainda está sendo preparada. Aguarde alguns segundos e tente novamente.");return;}
 if(!navigator.share||!navigator.canShare||!navigator.canShare({files:[file]})){
  msg("Este navegador não oferece compartilhamento direto de arquivos. Use o ícone de download e anexe o Excel no aplicativo desejado.");return;
 }
 try{
  await navigator.share({files:[file],title:"Balanço Império - Contagem",text:"Planilha de contagem das equipes."});
 }catch(e){
  if(e.name==="AbortError")return;
  msg("O navegador não permitiu abrir o compartilhamento. Tente novamente ou use o botão de download.");
 }
}
async function exportCounts(){
 const button=$("exportCounts");button.disabled=true;
 try{
  const file=await buildExcelFile();downloadExcelFile(file);
  preparedShareFile=file;msg("Arquivo Excel baixado e atualizado.");
 }catch(e){msg(e.message||"Não foi possível gerar o arquivo Excel.")}finally{button.disabled=false}
}
$("exportCounts").addEventListener("click",exportCounts);
$("shareCounts").addEventListener("click",shareExcelFile);
