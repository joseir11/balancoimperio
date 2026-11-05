const SUPABASE_URL="https://znivpoqcvxkgscqoszfj.supabase.co";
const SUPABASE_KEY="sb_publishable_-77CwES71Z8Js-zUoADACQ_WAPZDOk6";
const TOKEN_KEY="balanco_imperio_team_access";
let accessToken="",team=null,materials=[],statuses={},own=[],selected=null,pendingCode="",ownStatusFilter="TODOS",ownListCollapsed=false;
const $=id=>document.getElementById(id);
function esc(s){return String(s??"").replace(/&/g,"&").replace(/</g,"<").replace(/>/g,">").replace(/"/g,""").replace(/'/g,"&#39;")}
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
 if(!window.ExcelJS)throw new Error("Nao foi possivel carregar o gerador de Excel.");
 const data=await adminApi("export");const allTeams=data.teams||[],mats=data.materials||[],counts=data.counts||[];
 const teams=allTeams.filter(t=>counts.some(c=>c.team_id===t.id));
 if(!teams.length)throw new Error("Ainda nao existem contagens registradas pelas equipes.");
 const byKey=new Map(counts.map(c=>[c.material_id+"|"+c.team_id,c.quantity]));
 const workbook=new ExcelJS.Workbook();
 workbook.creator="BALANCO IMPERIO";workbook.created=new Date();workbook.modified=new Date();
 const sheet=workbook.addWorksheet("CONTAGEM",{views:[{state:"frozen",ySplit:1,showGridLines:false}],properties:{defaultRowHeight:21}});
 const headers=["CODIGO","DESCRICAO","UNIDADE",...teams.map(t=>t.name),"CONTAGEM CONFERIDA"];
 sheet.addRow(headers);
 const header=sheet.getRow(1);header.height=30;
 header.eachCell(cell=>{cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF111C2D"}};cell.font={name:"Arial",size:10,bold:true,color:{argb:"FFFFFFFF"}};cell.alignment={vertical:"middle",horizontal:"center",wrapText:true};cell.border={bottom:{style:"medium",color:{argb:"FFF28B22"}}}});
 mats.forEach(m=>{
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
 const buffer=await workbook.xlsx.writeBuffer();
 const filename="BALANCO_IMPERIO_CONTAGEM_"+new Date().toISOString().slice(0,10)+".xlsx";
 const mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
 return new File([new Blob([buffer],{type:mime})],filename,{type:mime});
}
async function prepareShareFile(showMessage=false){
 if(sharePreparation||!adminSession)return;
 sharePreparation=true;const button=$("shareCounts");if(button)button.disabled=true;
 try{
  const file=await buildExcelFile();preparedShareFile=file;
  if(showMessage)msg("Planilha pronta.");
 }catch(e){
  if(showMessage)msg(e.message||"Nao foi possivel preparar a planilha.");
 }finally{
  sharePreparation=false;if(button)button.disabled=false;
 }
}
async function shareExcelFile(){
 const file=preparedShareFile;
 if(!file){msg("A planilha ainda esta sendo preparada.");return;}
 if(!navigator.share||!navigator.canShare||!navigator.canShare({files:[file]})){
  msg("Use o icone de download e anexe o Excel.");return;
 }
 try{
  await navigator.share({files:[file],title:"Balanco Imperio",text:"Contagem das equipes."});
 }catch(e){
  if(e.name==="AbortError")return;
  msg("Nao foi possivel compartilhar. Use o download.");
 }
}
async function exportCounts(){
 const button=$("exportCounts");if(button)button.disabled=true;
 try{
  const file=await buildExcelFile();downloadExcelFile(file);
  preparedShareFile=file;msg("Excel baixado.");
 }catch(e){msg(e.message||"Nao foi possivel gerar o Excel.")}finally{if(button)button.disabled=false}
}
document.addEventListener("DOMContentLoaded",()=>{
 const ex=$("exportCounts");const sh=$("shareCounts");
 if(ex)ex.addEventListener("click",exportCounts);
 if(sh)sh.addEventListener("click",shareExcelFile);
});
