const SUPABASE_URL="https://znivpoqcvxkgscqoszfj.supabase.co";
const SUPABASE_KEY="sb_publishable_-77CwES71Z8Js-zUoADACQ_WAPZDOk6";
const TOKEN_KEY="balanco_imperio_team_access";
const ADMIN_SESSION_KEY="balanco_imperio_admin_session";
let accessToken="",team=null,materials=[],statuses={},own=[],selected=null,pendingCode="",ownStatusFilter="TODOS",ownListCollapsed=false;
const $=id=>document.getElementById(id);
function esc(s){const m={38:"amp",60:"lt",62:"gt",34:"quot",39:"#39"};return String(s??"").replace(/[&<>"']/g,c=>"&"+(m[c.charCodeAt(0)]||"")+";")}
function msg(t){$("toast").textContent=t;$("toast").classList.remove("hidden");setTimeout(()=>$("toast").classList.add("hidden"),3200)}
function badge(s){return '<span class="status-dot '+s+'" title="'+(s==="OK"?"OK":s==="DIVERGENTE"?"Divergente":"Aguardando")+'"></span>'}
let adminSession=null;
async function ensureAdminSession(){
 if(!adminSession)throw new Error("Sessao administrativa nao inicializada.");
 const expiresAt=Number(adminSession.expires_at||0);
 if(expiresAt&&expiresAt>Date.now()/1000+60)return;
 if(!adminSession.refresh_token)throw new Error("Sessao administrativa expirada. Entre novamente.");
 const response=await fetch(SUPABASE_URL+"/auth/v1/token?grant_type=refresh_token",{method:"POST",headers:{"Content-Type":"application/json","apikey":SUPABASE_KEY},body:JSON.stringify({refresh_token:adminSession.refresh_token})});
 const data=await response.json();
 if(!response.ok)throw new Error(data.error_description||data.msg||"Sessao administrativa expirada.");
 adminSession=data;localStorage.setItem(ADMIN_SESSION_KEY,JSON.stringify(data));
}
function showStart(section){$("start").classList.remove("hidden");$("app").classList.add("hidden");["createPanel","codePanel","adminLoginPanel","adminPanel"].forEach(id=>$(id).classList.add("hidden"));$(section==="create"?"createPanel":section==="code"?"codePanel":section==="adminLogin"?"adminLoginPanel":"adminPanel").classList.remove("hidden")}
async function adminApi(action,payload={}){await ensureAdminSession();const response=await fetch(SUPABASE_URL+"/functions/v1/manage-team",{method:"POST",headers:{"Content-Type":"application/json","apikey":SUPABASE_KEY,"Authorization":"Bearer "+adminSession.access_token},body:JSON.stringify({action,...payload})});const result=await response.json();if(!response.ok)throw new Error(result.error||"Falha ao administrar equipes.");return result}
let preparedShareFile=null,sharePreparation=false;
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
 const buffer=await workbook.xlsx.writeBuffer();
 const filename="BALANCO_IMPERIO_CONTAGEM_"+new Date().toISOString().slice(0,10)+".xlsx";
 const mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
 return new File([buffer],filename,{type:mime,lastModified:Date.now()});
}
async function prepareShareFile(showMessage=false){
 if(sharePreparation||!adminSession)return;
 sharePreparation=true;const button=$("shareCounts");if(button)button.disabled=true;
 try{const file=await buildExcelFile();preparedShareFile=file;if(showMessage)msg("Planilha pronta.");}
 catch(e){if(showMessage)msg(e.message||"Erro ao preparar planilha.");}
 finally{sharePreparation=false;if(button)button.disabled=false;}
}
async function shareExcelFile(){
 const file=preparedShareFile;
 if(!file){msg("Planilha ainda sendo preparada. Aguarde e tente novamente.");return;}
 if(!window.isSecureContext){msg("O compartilhamento exige abrir o aplicativo por HTTPS.");return;}
 if(typeof navigator.share!=="function"){msg("O Chrome nao disponibilizou o compartilhamento neste contexto. Abra o app diretamente no Chrome.");return;}
 try{
  await navigator.share({
   files:[file],
   title:"BALANCO IMPERIO",
   text:"Contagem das equipes."
  });
 }catch(e){
  if(e&&e.name==="AbortError")return;
  const reason=(e&&e.name?e.name+": ":"")+(e&&e.message?e.message:"falha desconhecida");
  msg("Falha no compartilhamento: "+reason);
 }
}
async function exportCounts(){
 const button=$("exportCounts");if(button)button.disabled=true;
 try{const file=await buildExcelFile();downloadExcelFile(file);preparedShareFile=file;msg("Excel baixado.");}
 catch(e){msg(e.message||"Nao foi possivel gerar o Excel.");}
 finally{if(button)button.disabled=false;}
}
function bindExportButtons(){
 const ex=$("exportCounts");const sh=$("shareCounts");
 if(ex)ex.addEventListener("click",exportCounts);
 if(sh)sh.addEventListener("click",shareExcelFile);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",bindExportButtons);
else bindExportButtons();


async function loadAdminTeams(){
 const data=await adminApi("list");
 $("adminTeamsList").innerHTML=(data.teams||[]).map(t=>'<div class="item team-card"><div class="iteminfo"><div class="desc">'+esc(t.name)+'</div><div class="unit">'+(t.active?'Ativa':'Inativa')+'</div></div><div class="item-tools"><button class="icon-btn edit" type="button" data-config="'+t.id+'" title="Editar" aria-label="Editar"><svg viewBox="0 0 24 24"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4Z"/></svg></button><button class="icon-btn '+(t.active?'toggle':'off')+'" type="button" data-status="'+t.id+'" data-active="'+(!t.active)+'" title="'+(t.active?'Desativar':'Ativar')+'" aria-label="Status"><svg viewBox="0 0 24 24"><path d="M12 22c5.5 0 10-4.5 10-10S17.5 2 12 2 2 6.5 2 12s4.5 10 10 10z"/><path d="m9 12 2 2 4-4"/></svg></button><button class="icon-btn delete" type="button" data-delete="'+t.id+'" title="Excluir" aria-label="Excluir"><svg viewBox="0 0 24 24"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="m19 6-1 14H6L5 6"/><path d="M10 11v5M14 11v5"/></svg></button></div></div>').join("")||'<div class="empty">Nenhuma equipe.</div>';
 $("adminTeamsList").querySelectorAll("[data-config]").forEach(b=>b.addEventListener("click",async()=>{const t=(data.teams||[]).find(x=>x.id===b.dataset.config);const name=prompt("Nome da equipe:",t.name);if(name===null)return;const password=prompt("Nova senha (min. 6). Vazio = manter:");if(password===null)return;try{await adminApi("update",{team_id:t.id,name,password});msg("Equipe atualizada.");await loadAdminTeams()}catch(e){msg(e.message)}}));
 $("adminTeamsList").querySelectorAll("[data-status]").forEach(b=>b.addEventListener("click",async()=>{try{await adminApi("status",{team_id:b.dataset.status,active:b.dataset.active==="true"});await loadAdminTeams();msg("Situacao atualizada.")}catch(e){msg(e.message)}}));
 $("adminTeamsList").querySelectorAll("[data-delete]").forEach(b=>b.addEventListener("click",async()=>{const t=(data.teams||[]).find(x=>x.id===b.dataset.delete);if(!confirm('Excluir equipe "'+t.name+'"?'))return;try{await adminApi("delete",{team_id:t.id});await loadAdminTeams();msg("Equipe excluida.")}catch(e){msg(e.message)}}));
}
async function api(action,payload={}){if(!accessToken)throw new Error("Sessão da equipe não inicializada. Volte e entre novamente.");const response=await fetch(SUPABASE_URL+"/functions/v1/team-api",{method:"POST",headers:{"Content-Type":"application/json","apikey":SUPABASE_KEY,"Authorization":"Bearer "+String(accessToken).trim()},body:JSON.stringify({action,...payload})});const result=await response.json();if(!response.ok)throw new Error(result.error||"Falha na comunicacao.");return result}
async function loadData(){const data=await api("data");team=data.team;materials=data.materials||[];statuses=Object.fromEntries((data.statuses||[]).map(x=>[x.material_id,x]));own=data.counts||[];renderOwn();$("teamTitle").textContent=team.name;$("app").classList.remove("hidden");$("start").classList.add("hidden");$("connection").className="status-dot OK";$("connection").title="Online";$("connection").textContent=""}
async function openWithToken(token){accessToken=token;try{await loadData();localStorage.setItem(TOKEN_KEY,token)}catch(e){localStorage.removeItem(TOKEN_KEY);accessToken="";throw e}}
$("showCode").addEventListener("click",()=>showStart("code"));
$("showAdmin").addEventListener("click",()=>showStart("adminLogin"));
$("backHomeAdmin").addEventListener("click",()=>showStart("create"));
$("backCreate").addEventListener("click",()=>showStart("create"));
$("adminLoginForm").addEventListener("submit",async e=>{e.preventDefault();const btn=$("adminLoginBtn");btn.disabled=true;try{const res=await fetch(SUPABASE_URL+"/auth/v1/token?grant_type=password",{method:"POST",headers:{"Content-Type":"application/json","apikey":SUPABASE_KEY},body:JSON.stringify({email:$("adminEmail").value.trim(),password:$("adminPassword").value})});const data=await res.json();if(!res.ok)throw new Error(data.error_description||data.msg||"Falha no login.");adminSession=data;localStorage.setItem(ADMIN_SESSION_KEY,JSON.stringify(data));showStart("adminPanel");await loadAdminTeams();prepareShareFile()}catch(err){msg(err.message)}finally{btn.disabled=false}});
$("adminLogout").addEventListener("click",()=>{adminSession=null;localStorage.removeItem(ADMIN_SESSION_KEY);showStart("create")});
$("adminCreateForm").addEventListener("submit",async e=>{e.preventDefault();const btn=$("adminCreateBtn");btn.disabled=true;try{await adminApi("create",{name:$("adminTeamName").value.trim(),password:$("adminTeamPassword").value});$("adminTeamName").value="";$("adminTeamPassword").value="";msg("Equipe cadastrada.");await loadAdminTeams()}catch(err){msg(err.message)}finally{btn.disabled=false}});
$("codeForm").addEventListener("submit",async e=>{
 e.preventDefault();
 const btn=$("enterBtn"),name=$("accessCode").value.trim().replace(/\s+/g," "),password=$("teamPassword").value;
 if(!name||!password){msg("Informe o nome da equipe e a senha.");return;}
 btn.disabled=true;
 try{
  const response=await fetch(SUPABASE_URL+"/functions/v1/team-login",{
   method:"POST",
   headers:{"Content-Type":"application/json","apikey":SUPABASE_KEY},
   body:JSON.stringify({name,password})
  });
  const result=await response.json().catch(()=>({}));
  if(!response.ok)throw new Error(result.error||"Não foi possível acessar. Confira o nome da equipe e a senha.");
  const teamToken=String(result.access_code||"").trim();
  if(!teamToken)throw new Error("O servidor não retornou o código de acesso da equipe.");
  await openWithToken(teamToken);
  $("teamPassword").value="";
  $("accessCode").value="";
  msg("Acesso realizado com sucesso.");
 }catch(err){
  msg(err.message||"Falha ao acessar a equipe.");
 }finally{
  btn.disabled=false;
 }
});
function renderOwn(){
 const items=own.map(c=>({...c,m:materials.find(m=>m.id===c.material_id),status:statuses[c.material_id]?.status||"AGUARDANDO"})).filter(c=>c.m);
 const filtered=ownStatusFilter==="TODOS"?items:items.filter(c=>c.status===ownStatusFilter);
 $("ownTotal").textContent=items.length;
 $("ownList").innerHTML=filtered.length?filtered.map(c=>'<div class="item"><div class="iteminfo"><div class="code">'+esc(c.m.code)+'</div><div class="desc">'+esc(c.m.description)+'</div><div class="unit">'+esc(c.m.unit)+'</div><div class="observation">'+(c.observation?'OBS: '+esc(c.observation):'')+'</div></div><div class="item-tools"><b class="item-value">'+esc(c.quantity)+'</b><span class="status-dot '+esc(c.status)+'"></span><button class="icon-btn edit" type="button" data-edit="'+c.material_id+'" title="Editar" aria-label="Editar"><svg viewBox="0 0 24 24"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4Z"/></svg></button><button class="icon-btn delete" type="button" data-del="'+c.id+'" title="Excluir" aria-label="Excluir"><svg viewBox="0 0 24 24"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="m19 6-1 14H6L5 6"/><path d="M10 11v5M14 11v5"/></svg></button></div></div>').join(""):(items.length?'<div class="empty">Nenhum neste status.</div>':'<div class="empty">Nenhuma contagem.</div>');
 $("ownListContent").classList.toggle("own-list-collapsed",ownListCollapsed);
 $("toggleOwnList").setAttribute("aria-expanded",String(!ownListCollapsed));
 $("ownList").querySelectorAll("[data-edit]").forEach(b=>b.addEventListener("click",()=>choose(b.dataset.edit)));
 $("ownList").querySelectorAll("[data-del]").forEach(b=>b.addEventListener("click",()=>deleteCount(b.dataset.del)));
}
$("toggleOwnList").addEventListener("click",()=>{ownListCollapsed=!ownListCollapsed;renderOwn()});
$("ownStatusFilter").addEventListener("change",()=>{ownStatusFilter=$("ownStatusFilter").value;renderOwn()});
$("search").addEventListener("input",()=>{const q=$("search").value.toLowerCase().trim();const found=q?materials.filter(m=>String(m.code).toLowerCase().includes(q)||String(m.description).toLowerCase().includes(q)).slice(0,30):[];$("options").innerHTML=found.map(m=>'<button class="option" type="button" data-id="'+m.id+'"><b>'+esc(m.code)+'</b> — '+esc(m.description)+' <small>'+esc(m.unit)+'</small></button>').join("")||(q?'<div class="empty">Nenhum material.</div>':"");$("options").classList.toggle("hidden",!q);$("options").querySelectorAll("[data-id]").forEach(b=>b.addEventListener("click",()=>choose(b.dataset.id)))});
function choose(id){selected=materials.find(m=>m.id===id);if(!selected)return;const c=own.find(x=>x.material_id===id);$("options").classList.add("hidden");$("search").value=selected.code+" — "+selected.description;$("chosenCode").textContent=selected.code;$("chosenDesc").textContent=selected.description;$("chosenUnit").textContent=selected.unit;$("qty").value=c?.quantity??"";$("observation").value=c?.observation??"";$("chosen").classList.remove("hidden");$("qty").focus()}
$("save").addEventListener("click",async()=>{if(!selected||$("qty").value===""||!Number.isFinite(+$("qty").value)||+$("qty").value<0)return msg("Informe uma quantidade valida.");const btn=$("save");btn.disabled=true;try{await api("save",{material_id:selected.id,quantity:+$("qty").value,observation:$("observation").value});selected=null;$("search").value="";$("qty").value="";$("observation").value="";$("options").innerHTML="";$("options").classList.add("hidden");$("chosen").classList.add("hidden");await loadData();$("search").focus();msg("Contagem salva.")}catch(e){msg(e.message)}finally{btn.disabled=false;btn.innerHTML='<svg viewBox="0 0 24 24"><path d="M5 3h12l4 4v14H3V3h2z"/><path d="M7 3v6h10V3M7 21v-7h10v7"/></svg>'}});
async function deleteCount(id){if(!confirm("Excluir esta contagem?"))return;try{await api("delete",{count_id:id});await loadData();$("chosen").classList.add("hidden");msg("Contagem excluida.")}catch(e){msg(e.message)}}
$("switchTeam").addEventListener("click",()=>{if(confirm("Sair desta equipe neste aparelho?")){localStorage.removeItem(TOKEN_KEY);accessToken="";team=null;showStart("create")}});
$("copyTeamCode")?.addEventListener("click",()=>{const t=$("teamCodeText")?.textContent;if(t)navigator.clipboard?.writeText(t).then(()=>msg("Codigo copiado."))});
$("toggleObs")?.addEventListener("click",()=>$("obsWrap").classList.toggle("hidden"));
(async()=>{
 showStart("create");
 const savedTeamToken=localStorage.getItem(TOKEN_KEY);
 if(savedTeamToken){
  try{await openWithToken(savedTeamToken);return;}
  catch(e){localStorage.removeItem(TOKEN_KEY);accessToken="";showStart("create");}
 }
 const savedAdmin=localStorage.getItem(ADMIN_SESSION_KEY);
 if(savedAdmin){
  try{
   adminSession=JSON.parse(savedAdmin);
   await ensureAdminSession();
   showStart("adminPanel");
   await loadAdminTeams();
   prepareShareFile();
   return;
  }catch(e){adminSession=null;localStorage.removeItem(ADMIN_SESSION_KEY);showStart("create");}
 }
})();
if("serviceWorker" in navigator){window.addEventListener("load",()=>navigator.serviceWorker.register("./sw.js").catch(()=>{}))}
