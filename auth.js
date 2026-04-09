
// ═══ CONFIG + SHARED AUTH (from Gestão de Reuniões) ═══
const CLIENT_ID='b3a66901-ec96-424b-9bcb-370b39a2a210';
const TENANT_ID='09ae14e9-8398-4ca0-a5cd-994620a0ab3a';
const SP='https://fjuventude.sharepoint.com/sites/PlataformasInternasFJ';
const REDIRECT=window.location.href.split('?')[0].split('#')[0];
const SCOPES=['https://fjuventude.sharepoint.com/AllSites.Write','https://fjuventude.sharepoint.com/AllSites.Read'];
const GRAPH_SCOPES=['User.Read','GroupMember.Read.All'];

// GROUP IDs
const GROUPS={
  CA:'0c05918f-4508-45fd-82f4-d5cc3925cd4f',
  Gestao:'c85b392f-df39-413d-897a-6b0af64a7a8e',
  Impacto:'cd0afde0-060e-4838-9e34-1ed8e2416c15',
  Financeiro:'dd09b5dd-79e8-4ecf-a152-ae79098ab711',
  OpSul:'192fb910-a387-4838-bad5-db6a9e831022',
  OpNorte:'bbc476e9-8f75-46d1-9fa7-cf9cc5b23535'
};

// Reunião type → required group
const REUNIAO_GROUPS={
  'Reunião com CAdm':[GROUPS.CA],
  'Reunião Gestão':[GROUPS.CA,GROUPS.Gestao],
  'Reunião O.Norte':[GROUPS.CA,GROUPS.Gestao,GROUPS.OpNorte],
  'Reunião O.Sul':[GROUPS.CA,GROUPS.Gestao,GROUPS.OpSul],
  'Reunião Financeiro':[GROUPS.CA,GROUPS.Gestao,GROUPS.Financeiro],
  'Reunião Impacto':[GROUPS.CA,GROUPS.Gestao,GROUPS.Impacto]
};


// Meeting level colors
const REUNIAO_CORES={
  'Reunião com CAdm':{bg:'#003087',text:'#fff',light:'#E6F1FB',badge:'b-ca',ppt:{header:'003087',accent:'1A5BC9',light:'E6F1FB'}},
  'Reunião Gestão':{bg:'#0F6E56',text:'#fff',light:'#E1F5EE',badge:'b-opnorte',ppt:{header:'0F6E56',accent:'1D9E75',light:'E1F5EE'}},
  'Reunião O.Norte':{bg:'#B8860B',text:'#fff',light:'#FAEEDA',badge:'b-ambos',ppt:{header:'B8860B',accent:'D4A017',light:'FAEEDA'}},
  'Reunião O.Sul':{bg:'#B8860B',text:'#fff',light:'#FAEEDA',badge:'b-ambos',ppt:{header:'B8860B',accent:'D4A017',light:'FAEEDA'}},
  'Reunião Financeiro':{bg:'#534AB7',text:'#fff',light:'#EEEDFE',badge:'b-fin',ppt:{header:'534AB7',accent:'7B74D0',light:'EEEDFE'}},
  'Reunião Impacto':{bg:'#D85A30',text:'#fff',light:'#FDE8DF',badge:'b-imp',ppt:{header:'D85A30',accent:'E8784F',light:'FDE8DF'}}
};
function getCores(tipo){return REUNIAO_CORES[tipo]||REUNIAO_CORES['Reunião com CAdm'];}


// ── MODULE SWITCHER ──
let _currentModule='rn';
function switchModule(mod){
  _currentModule=mod;
  ['rn','pl','gf'].forEach(m=>{
    document.getElementById('tab-'+m)?.classList.toggle('active',m===mod);
    const sb=document.getElementById('sidebar-'+m);
    const mc=document.getElementById('module-'+m);
    if(sb)sb.style.display=m===mod?'block':'none';
    if(mc)mc.style.display=m===mod?'block':'none';
  });
  if(mod==='pl'){
    if(!_plLoaded){_plLoaded=true;plInit();}
    plShowPage('dash');
  }
  if(mod==='gf'&&!_gfLoaded){_gfLoaded=true;gfInit();}
  if(mod==='rn')showPage('assuntos');
}
let _plLoaded=false,_gfLoaded=false;

// ── SHARED PROJECTS ──
let sharedProjetos=[];
async function loadSharedProjetos(){
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Projetos')/items?$select=*&$orderby=Title asc&$top=500`);
    sharedProjetos=d.d.results.map(r=>({
      id:r.Id,nome:r.Title||'',codigo:r.Codigo||'',dep:r.Departamento||'',
      estado:r.Estado||'Ativo',dataInicio:r.DataInicio?r.DataInicio.split('T')[0]:'',
      dataFim:r.DataFim?r.DataFim.split('T')[0]:'',
      anoOrcamental:r.AnoOrcamental||new Date().getFullYear(),
      isBancoGlobal:r.IsBancoGlobal||false
    }));
  }catch(e){console.warn('loadSharedProjetos',e);}
}

// ── FINANCE ROLE ──
function getFinanceRole(groups){
  if(groups.includes(GROUPS.CA))return'pca';
  if(groups.includes(GROUPS.Gestao))return'pe';
  if(groups.includes(GROUPS.Financeiro))return'df';
  if([GROUPS.OpNorte,GROUPS.OpSul,GROUPS.Impacto].some(g=>groups.includes(g)))return'area_director';
  return'worker';
}
let _financeRole='worker';

const msalCfg={auth:{clientId:CLIENT_ID,authority:`https://login.microsoftonline.com/${TENANT_ID}`,redirectUri:REDIRECT}};
const msalI=new msal.PublicClientApplication(msalCfg);

let user=null,spToken=null,graphToken=null;
let userGroups=[];
let items=[],extItems=[],rnTarefas=[];
let editId=null,editExtId=null,editTarId=null;
let T={conf:false,atras:false};
let carrItems=[],carrIdx=0,carrDec={},fsOpen=false,currReuniao='';

async function init(){
  await msalI.initialize();
  const r=await msalI.handleRedirectPromise();
  if(r){user=r.account;spToken=r.accessToken;await loadUserGroups();showApp();}
  else{const a=msalI.getAllAccounts();if(a.length){user=a[0];await getSPToken();await loadUserGroups();showApp();}}
}
async function doLogin(){
  document.getElementById('login-msg').textContent='A autenticar...';
  try{await msalI.loginRedirect({scopes:[...SCOPES,...GRAPH_SCOPES]});}
  catch(e){document.getElementById('login-msg').textContent='Erro: '+e.message;}
}
function doLogout(){msalI.logoutRedirect();}
async function getSPToken(){
  try{const r=await msalI.acquireTokenSilent({scopes:SCOPES,account:user});spToken=r.accessToken;}
  catch(e){await msalI.acquireTokenRedirect({scopes:SCOPES});}
}
async function getGraphToken(){
  try{const r=await msalI.acquireTokenSilent({scopes:GRAPH_SCOPES,account:user});graphToken=r.accessToken;}
  catch(e){console.warn('Graph token error',e);}
}
async function loadUserGroups(){
  try{
    await getGraphToken();
    // Use transitiveMemberOf to catch all group memberships including nested
    let url='https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=100';
    let allGroups=[];
    while(url){
      const r=await fetch(url,{headers:{Authorization:`Bearer ${graphToken}`}});
      const d=await r.json();
      if(d.value)allGroups=[...allGroups,...d.value];
      url=d['@odata.nextLink']||null;
    }
    userGroups=allGroups.map(g=>g.id);
    console.log('User groups loaded:',allGroups.map(g=>g.displayName));
  }catch(e){
    console.warn('Could not load groups, trying fallback',e);
    // Fallback: try memberOf
    try{
      await getGraphToken();
      const r=await fetch('https://graph.microsoft.com/v1.0/me/memberOf?$select=id,displayName&$top=100',{headers:{Authorization:`Bearer ${graphToken}`}});
      const d=await r.json();
      userGroups=(d.value||[]).map(g=>g.id);
    }catch(e2){console.warn('Both group methods failed',e2);}
  }
}

// ═══ SISTEMA DE UTILIZADORES ═══
let orgUsers=[];
let userRoles={}; // {email: {role:'worker'|'area_director'|'df'|'pe'|'pca', dept:'', projects:[]}}

// Carregar utilizadores da organização via Graph API
async function loadOrgUsers(){
  try{
    await getGraphToken();
    const r=await fetch('https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,department,jobTitle&$top=100',
      {headers:{Authorization:`Bearer ${graphToken}`}});
    const d=await r.json();
    orgUsers=(d.value||[]).map(u=>({
      id:u.id,
      name:u.displayName||'',
      email:u.mail||u.userPrincipalName||'',
      dept:u.department||'',
      title:u.jobTitle||''
    }));
    console.log('Org users loaded:',orgUsers.length);
    // Carregar roles guardados no localStorage (ou futuramente de uma lista SP)
    const savedRoles=localStorage.getItem('fj_user_roles');
    if(savedRoles)userRoles=JSON.parse(savedRoles);
  }catch(e){console.warn('loadOrgUsers error',e);}
}

function saveUserRoles(){
  localStorage.setItem('fj_user_roles',JSON.stringify(userRoles));
}

function getUserByEmail(email){
  if(!email)return null;
  const u=orgUsers.find(u=>u.email?.toLowerCase()===email.toLowerCase());
  if(u){
    const role=userRoles[email]||{role:'worker',dept:u.dept,projects:[]};
    return{...u,...role};
  }
  return null;
}

function getUserName(email){
  const u=getUserByEmail(email);
  return u?.name||email||'—';
}

// Determinar role do utilizador atual baseado nos grupos
function getCurrentUserRole(){
  if(userGroups.includes(GROUPS.CA))return'pca';
  if(userGroups.includes(GROUPS.Gestao))return'pe';
  if(userGroups.includes(GROUPS.Financeiro))return'df';
  if([GROUPS.OpNorte,GROUPS.OpSul,GROUPS.Impacto].some(g=>userGroups.includes(g)))return'area_director';
  return'worker';
}

// Verificar se utilizador atual é admin (CA ou Gestão)
function isAdmin(){
  return userGroups.includes(GROUPS.CA)||userGroups.includes(GROUPS.Gestao);
}

function hasGroupAccess(groupId){return userGroups.includes(groupId)||userGroups.includes(GROUPS.CA);}
function canSeeReuniao(tipoReuniao){
  const req=REUNIAO_GROUPS[tipoReuniao];
  if(!req)return true;
  return req.some(gid=>userGroups.includes(gid));
}

function toggleDropdown(){
  document.getElementById('user-dropdown').classList.toggle('open');
}
document.addEventListener('click',e=>{
  if(!e.target.closest('.topbar-right'))document.getElementById('user-dropdown').classList.remove('open');
});


function showApp(){
  document.getElementById('login-screen').style.display='none';
  document.getElementById('app-screen').style.display='block';
  const name=user.name||user.username;
  const initials=name.split(' ').map(n=>n[0]).slice(0,2).join('').toUpperCase();
  sv('user-name',name);sv('user-email',user.username);
  document.getElementById('user-initials').textContent=initials;
  document.getElementById('dd-name').textContent=name;
  document.getElementById('dd-email').textContent=user.username;
  const groupNames=Object.entries(GROUPS).filter(([k,v])=>userGroups.includes(v)).map(([k])=>k).join(', ')||'Sem grupos';
  document.getElementById('dd-groups').textContent=groupNames;
  _financeRole=getFinanceRole(userGroups);
  
  // Mostrar menu de administração se for admin (CA ou Gestão)
  if(isAdmin()){
    const adminLabel=document.getElementById('admin-section-label');
    if(adminLabel)adminLabel.style.display='block';
    const navAdmin=document.getElementById('nav-admin');
    if(navAdmin)navAdmin.classList.remove('hidden');
    // Carregar utilizadores da organização
    loadOrgUsers();
  }
  
  loadSharedProjetos();
  loadItems();loadExtItems();loadTarefas();loadHistReunioes();loadSavedSessions();
}


async function getDigest(){
  await getSPToken();
  const r=await fetch(`${SP}/_api/contextinfo`,{method:'POST',headers:{Accept:'application/json;odata=verbose',Authorization:`Bearer ${spToken}`,'Content-Length':'0'}});
  const d=await r.json();return d.d.GetContextWebInformation.FormDigestValue;
}
async function spRead(url){
  await getSPToken();
  const r=await fetch(url,{headers:{Accept:'application/json;odata=verbose',Authorization:`Bearer ${spToken}`}});
  if(!r.ok)throw new Error('Erro '+r.status);return r.json();
}
async function spWrite(url,body,method,etag){
  await getSPToken();
  const digest=await getDigest();
  const h={'Content-Type':'application/json;odata=verbose','Accept':'application/json;odata=verbose','Authorization':`Bearer ${spToken}`,'X-RequestDigest':digest};
  if(method==='MERGE'||method==='DELETE'){h['X-HTTP-Method']=method;h['IF-MATCH']=etag||'*';}
  const r=await fetch(url,{method:'POST',headers:h,body:body?JSON.stringify(body):undefined});
  if(!r.ok){const t=await r.text();throw new Error(`HTTP ${r.status}: ${t.slice(0,200)}`);}
  return r;
}
async function uploadAttachment(list,itemId,file){
  await getSPToken();const digest=await getDigest();
  const url=`${SP}/_api/web/lists/getbytitle('${list}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
  const buf=await file.arrayBuffer();
  const r=await fetch(url,{method:'POST',headers:{Authorization:`Bearer ${spToken}`,'X-RequestDigest':digest,'content-length':buf.byteLength},body:buf});
  if(!r.ok)throw new Error('Upload falhou');
}

// PEOPLE PICKER
async function searchPeople(q,suggestId,inputId){
  if(q.length<2){document.getElementById(suggestId).classList.remove('open');return;}
  try{
    await getGraphToken();
    const r=await fetch(`https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${encodeURIComponent(q)}') or startswith(mail,'${encodeURIComponent(q)}')&$select=displayName,mail,id&$top=8`,{headers:{Authorization:`Bearer ${graphToken}`}});
    const d=await r.json();
    const s=document.getElementById(suggestId);
    if(!d.value?.length){s.classList.remove('open');return;}
    s.innerHTML=d.value.map(p=>`<div class="people-item" onclick="selectPerson('${inputId}','${suggestId}','${p.displayName.replace(/'/g,"\\'")}')">
      <div class="people-avatar">${p.displayName.split(' ').map(n=>n[0]).slice(0,2).join('').toUpperCase()}</div>
      <div><div style="font-weight:500">${p.displayName}</div><div style="font-size:11px;color:var(--text3)">${p.mail||''}</div></div>
    </div>`).join('');
    s.classList.add('open');
  }catch(e){console.warn(e);}
}
function selectPerson(inputId,suggestId,name){
  document.getElementById(inputId).value=name;
  document.getElementById(suggestId).classList.remove('open');
}


// Inicializar aplicação
init();
