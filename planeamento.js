// ═══ MÓDULO: PLANEAMENTO ═══
// ═══ PLANEAMENTO MODULE ═══

function plShowPage(p){
  document.querySelectorAll('#module-pl .page').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('#sidebar-pl .nav-item').forEach(e=>e.classList.remove('active'));
  document.getElementById('pl-page-'+p)?.classList.add('active');
  document.getElementById('pl-nav-'+p)?.classList.add('active');
  if(p==='gantt') renderGantt();
}
/* ===== STATE ===== */

let plProjetos=[],plAcoes=[],plTarefas=[],plKpdefs=[],plKpregs=[],plPlanoTrim=[];
let plSemView='pessoal',plViewWeek,plViewYear;
let plEditTrimId=null,plEditKdefId=null,plEditKregId=null,plEditTarId=null;
let plHierOpen={};

/* ===== LOAD ===== */
async function plLoadAll(){
  showNotif('A carregar dados...','info');
  try{
    
    await Promise.all([loadProjetos(),loadAcoes(),loadTarefas(),loadKPIDefs(),loadKPIRegs(),loadPlanoTrim()]);
    renderDash();renderHier();renderTrim();renderSem();renderKdefs();renderKregs();populateDrops();
    showNotif(`Dados carregados! ${plTarefas.length} plTarefas · ${plProjetos.length} plProjetos · ${plAcoes.length} ações`,'success');
  }catch(e){showNotif('Erro ao carregar: '+e.message,'error');console.error(e);}
}
async function loadProjetos(){
  // Factorial_Projetos = plProjetos top-level (vindos do Factorial)
  const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Projetos')/items?$select=*&$orderby=Title asc&$top=1000`);
  plProjetos=d.d.results.map(r=>({id:r.Id,nome:r.Title||'',dep:r.Departamento||'',estado:r.Estado||'active',dataInicio:r.DataInicio?r.DataInicio.split('T')[0]:'',dataFim:r.DataFim?r.DataFim.split('T')[0]:'',notas:r.Notas||'',tipoFJ:r.Tipo_FJ||(r.Title?.includes('(INOV)')?'INOV':'Recorrente')}));
}
async function loadAcoes(){
  // Acoes = lista manual criada na plataforma, filho de Factorial_Projetos
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Acoes')/items?$select=*&$orderby=Title asc&$top=1000`);
    plAcoes=d.d.results.map(r=>({id:r.Id,nome:r.Title||'',estado:r.Estado||'Ativo',dataInicio:r.DataInicio?r.DataInicio.split('T')[0]:'',dataFim:r.DataFim?r.DataFim.split('T')[0]:'',projetoId:r.ProjetoId||null,dep:r.Departamento||'',tipoFJ:r.TipoAcao||'Recorrente',responsavel:'',notas:r.Notas||''}));
  }catch(e){console.warn('Acoes list error:',e.message);plAcoes=[];}
}
async function loadTarefas(){
  let all=[];
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items?$select=*&$orderby=Id desc&$top=2000`);
    const ft=d.d.results.map(r=>({
      id:'f_'+r.Id,spId:r.Id,
      titulo:r.Title||'',
      estado:r.Estado||r.Status||'todo',
      dataLimite:r.DataLimite?r.DataLimite.split('T')[0]:(r.due_on||''),
      semana:r.Semana||null,
      anoFJ:r.AnoFJ||null,
      notas:r.Notas||'',
      prioridade:r.Prioridade||'Média',
      projetoId:r.ProjetoId||null,
      acaoId:r.AcaoId||null,
      responsavel:r.Responsavel||'',
      dep:r.Departamento||'',
      lista:'FJ_Tarefas',
      attachments:[]
    }));
    all=[...all,...ft];
    showNotif(`✓ ${ft.length} plTarefas carregadas`,'success');
  }catch(e){
    showNotif('Erro a carregar plTarefas: '+e.message,'error');
    console.error('ERRO Factorial_Tarefas:',e.message);
  }
  all=all.map(t=>{
    if((!t.semana||!t.anoFJ)&&t.dataLimite){
      try{const d=new Date(t.dataLimite+'T12:00:00');t.semana=isoWeek(d);t.anoFJ=d.getFullYear();}catch(e){}
    }
    return t;
  });
  plTarefas=all;
}
async function loadKPIDefs(){
  const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_KPI_Definicoes')/items?$select=*&$orderby=Title asc&$top=500`);
  plKpdefs=d.d.results.map(r=>({id:r.Id,nome:r.Title||'',projetoId:r.ProjetoId||null,acaoId:r.AcaoId||null,tipo:r.TipoKPI||'Manual',dir:r.Direcao||'meta',uni:r.Unidade||'',meta:r.Meta||0,freq:r.Frequencia||'Mensal',resp:r.Responsavel||'',ativo:r.Ativo!==false,desc:r.Descricao||''}));
}
async function loadKPIRegs(){
  const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_KPI_Registos')/items?$select=*,InseridoPor/Title&$expand=InseridoPor&$orderby=Created desc&$top=500`);
  plKpregs=d.d.results.map(r=>({id:r.Id,periodo:r.Title||'',kpiId:r.KPIId||null,valor:r.Valor||0,trim:r.Trimestre||'',sem:r.Semana||null,ano:r.AnoFJ||null,dataReg:r.DataRegisto?r.DataRegisto.split('T')[0]:'',inseridoPor:r.InseridoPor?.Title||'',notas:r.Notas||''}));
}
async function loadPlanoTrim(){
  const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_PlanoTrimestral')/items?$select=*&$orderby=AnoFJ desc,Trimestre asc&$top=500`);
  plPlanoTrim=d.d.results.map(r=>({id:r.Id,titulo:r.Title||'',projetoId:r.ProjetoId||null,acaoId:r.AcaoId||null,trim:r.Trimestre||'',ano:r.AnoFJ||null,dep:r.Departamento||'',estado:r.Estado||'Previsto',aprovacao:r.EstadoAprov||'Rascunho',tarPrev:r.TarefasPrevistas||0,notas:r.Notas||''}));
}

/* ===== DROPDOWNS ===== */
function populateDrops(){
  // Projetos = Factorial_Projetos
  const po=plProjetos.map(p=>`<option value="${p.id}">${p.nome}</option>`).join('');
  ['ft-proj','fk-proj','fa-proj'].forEach(id=>{const e=el(id);if(e)e.innerHTML='<option value="">Selecionar...</option>'+po;});
  const ffkd=el('ff-kd-proj');if(ffkd)ffkd.innerHTML='<option value="">Todos os plProjetos</option>'+po;
  const ffap=el('ff-ac-proj');if(ffap)ffap.innerHTML='<option value="">Todos os plProjetos</option>'+po;
  // Ações = lista Acoes (manual)
  const ao=plAcoes.map(a=>`<option value="${a.id}">${a.nome}</option>`).join('');
  const ftA=el('ftar-acao');if(ftA)ftA.innerHTML='<option value="">Nenhuma</option>'+ao;
  // KPI dropdown
  const ko=plKpdefs.filter(k=>k.ativo).map(k=>`<option value="${k.id}">${k.nome}</option>`).join('');
  const frKpi=el('fr-kpi');if(frKpi)frKpi.innerHTML='<option value="">Selecionar KPI...</option>'+ko;
  const ffKr=el('ff-kr-kpi');if(ffKr)ffKr.innerHTML='<option value="">Todos os KPIs</option>'+ko;
}
function loadAcoesForProj(projId,targetId,currId=null){
  const e=el(targetId);if(!e)return;
  const list=projId?plAcoes.filter(a=>String(a.projetoId)===String(projId)):plAcoes;
  e.innerHTML='<option value="">Nenhuma</option>'+list.map(a=>`<option value="${a.id}"${String(a.id)===String(currId)?'selected':''}>${a.nome}</option>`).join('');
}

/* ===== DASHBOARD ===== */
function renderDash(){
  const w=isoWeek(new Date()),y=new Date().getFullYear();
  // Safety checks for elements that may not exist
  const elProj=el('s-proj'),elInov=el('s-inov'),elTsem=el('s-tsem'),elKpi=el('s-kpi');
  const elRag=el('dash-rag'),elKpis=el('dash-kpis'),elTar=el('dash-tarefas');
  if(!elProj||!elInov||!elTsem||!elKpi)return;
  // Projetos ativos = Factorial_Projetos com estado active
  elProj.textContent=plProjetos.filter(p=>p.estado==='active'||p.estado==='Ativo').length;
  // Ações INOV = ações criadas na plataforma com TipoAcao=INOV ou plProjetos do Factorial com (INOV)
  elInov.textContent=plAcoes.filter(a=>a.tipoFJ==='INOV').length+plProjetos.filter(p=>p.tipoFJ==='INOV'||p.nome.includes('(INOV)')).length;
  elTsem.textContent=plTarefas.filter(t=>t.semana==w&&t.anoFJ==y).length;
  elKpi.textContent=plKpdefs.filter(k=>k.ativo).length;
  // RAG
  const ragAcoes=plAcoes.slice(0,8);
  if(elRag)elRag.innerHTML=ragAcoes.length?`<table><colgroup><col style="width:10%"><col style="width:55%"><col style="width:35%"></colgroup><tbody>${ragAcoes.map(a=>{const r=calcRag(a.id);return`<tr><td><span class="rag rag-${r}"></span></td><td style="font-size:12px;font-weight:500">${a.nome}</td><td>${depB(a.dep)}</td></tr>`;}).join('')}</tbody></table>`:'<div class="empty" style="padding:1rem">Sem ações carregadas.</div>';
  // KPIs
  const km=plKpdefs.filter(k=>k.ativo&&k.tipo==='Manual').slice(0,6);
  if(elKpis)elKpis.innerHTML=km.length?`<table><colgroup><col style="width:55%"><col style="width:25%"><col style="width:20%"></colgroup><tbody>${km.map(k=>{const lr=plKpregs.filter(r=>r.kpiId==k.id).sort((a,b)=>b.id-a.id)[0];const pct=lr&&k.meta?Math.round((lr.valor/k.meta)*100):null;const c=pct===null?'var(--tx3)':kpiColor(k,pct);return`<tr><td style="font-size:12px;font-weight:500">${k.nome}</td><td style="font-size:11px">${lr?lr.valor+' '+k.uni:'—'}</td><td style="font-weight:700;color:${c}">${pct!==null?pct+'%':'—'}</td></tr>`;}).join('')}</tbody></table>`:'<div class="empty" style="padding:1rem">Sem KPIs manuais.</div>';
  // Tasks
  const wt=plTarefas.filter(t=>t.semana==w&&t.anoFJ==y).slice(0,8);
  if(elTar)elTar.innerHTML=wt.length?`<table><colgroup><col style="width:45%"><col style="width:25%"><col style="width:15%"><col style="width:15%"></colgroup><tbody>${wt.map(t=>`<tr class="${t.estado==='Concluída'?'done-row':''}"><td style="font-size:12px;font-weight:500">${t.titulo}</td><td style="font-size:11px">${t.responsavel||'—'}</td><td>${estB(t.prioridade)}</td><td>${estB(t.estado)}</td></tr>`).join('')}</tbody></table>`:'<div class="empty" style="padding:1rem">Sem plTarefas para esta semana.</div>';
}
function kpiColor(k,pct){
  if(k.dir==='limite'){return pct<=100?'var(--g5)':pct<=120?'var(--amb2)':'var(--red)';}
  return pct>=80?'var(--g5)':pct>=50?'var(--amb2)':'var(--red)';
}
function kpiEstado(k,pct){
  if(pct===null)return'—';
  if(k.dir==='limite'){return pct<=100?'✓ Atingido':pct<=120?'⚠ Em risco':'✗ Excedido';}
  return pct>=100?'✓ Atingido':pct>=80?'⚠ Em risco':'✗ Não atingido';
}

/* ===== HIERARQUIA ===== */
function renderHier(){
  const dep=gv('ff-h-dep'),tipo=gv('ff-h-tipo');
  const c=el('hier-container');
  const projs=plProjetos.filter(p=>!dep||p.dep===dep);
  if(!projs.length){c.innerHTML='<div class="empty">Sem plProjetos carregados do Factorial.</div>';return;}
  c.innerHTML=projs.map(p=>{
    const pAcoes=plAcoes.filter(a=>String(a.projetoId)===String(p.id)&&(!tipo||a.tipoFJ===tipo));
    // Tarefas não afetas a nenhuma ação mas com projetoId deste projeto, OU sem ação e sem projeto (mostrar em todos)
    const tarSemAcao=plTarefas.filter(t=>!t.acaoId&&(String(t.projetoId)===String(p.id)));
    const tarTotal=plTarefas.filter(t=>String(t.projetoId)===String(p.id)||pAcoes.some(a=>String(t.acaoId)===String(a.id)));
    const concTotal=tarTotal.filter(t=>t.estado==='Concluída').length;
    const pct=tarTotal.length?Math.round((concTotal/tarTotal.length)*100):null;
    const rag=pAcoes.length?(['r','a','v'].reduce((best,r)=>pAcoes.some(a=>calcRag(a.id)===r)?r:best,'c')):'c';
    const tipoFJ=p.tipoFJ||(p.nome.includes('(INOV)')?'INOV':'Recorrente');
    return`<div class="hier-proj">
      <div class="hier-proj-hdr" onclick="toggleHier(${p.id})">
        <span class="rag rag-${rag}"></span>
        <span style="font-size:14px;font-weight:700;color:var(--g1);flex:1">${p.nome}</span>
        ${badgeHTML(tipoFJ==='INOV'?'b-inov':'b-rec',tipoFJ)}
        ${depB(p.dep)}
        <span style="font-size:11px;background:var(--bg3);border-radius:6px;padding:2px 8px;color:var(--tx2)">${pAcoes.length} ações · ${tarTotal.length} plTarefas${pct!==null?' · '+pct+'% conc.':''}</span>
        <button class="btn btn-sm btn-primary" style="font-size:11px" onclick="event.stopPropagation();openAcaoForm(${p.id})">+ Ação</button>
        <svg id="hier-arrow-${p.id}" width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" style="transition:transform .2s"><path d="M4 6l4 4 4-4"/></svg>
      </div>
      <div class="hier-acao-wrap" id="hier-plAcoes-${p.id}">
        ${pAcoes.map(a=>{
          const ts=plTarefas.filter(t=>String(t.acaoId)===String(a.id));
          const conc=ts.filter(t=>t.estado==='Concluída').length;
          const r=calcRag(a.id);
          return`<div class="hier-acao-row" onclick="toggleHierAcao(${a.id})">
            <span class="rag rag-${r}"></span>
            <span class="hier-acao-name">${a.nome}</span>
            ${badgeHTML(a.tipoFJ==='INOV'?'b-inov':'b-rec',a.tipoFJ||'Rec')}
            ${a.dep?depB(a.dep):''}
            <span style="font-size:11px;color:var(--tx3)">${conc}/${ts.length} plTarefas</span>
            <button class="btn btn-sm" onclick="event.stopPropagation();openAcaoForm(${p.id},${a.id})">Editar</button>
            <button class="btn btn-sm btn-danger" onclick="event.stopPropagation();delAcao(${a.id})">✕</button>
            <button class="btn btn-sm btn-or" onclick="event.stopPropagation();openTarFormForAcao(${a.id})">+ Tarefa</button>
          </div>
          <div class="hier-tasks" id="hier-t-${a.id}">
            <div id="hier-t-content-${a.id}"></div>
          </div>`;
        }).join('')}
        ${tarSemAcao.length?`
          <div style="border-top:0.5px solid var(--bd);background:var(--bg2)">
            <div style="padding:8px 16px 6px;font-size:11px;font-weight:600;color:var(--tx3);text-transform:uppercase;letter-spacing:.05em;display:flex;align-items:center;gap:8px">
              Tarefas sem ação atribuída (${tarSemAcao.length})
              <button class="btn btn-sm" onclick="openTarFormForProj(${p.id})" style="font-size:10px">+ Tarefa</button>
            </div>
            <table style="width:100%;border-collapse:collapse">
              <colgroup><col style="width:4%"><col style="width:36%"><col style="width:20%"><col style="width:15%"><col style="width:10%"><col style="width:10%"><col style="width:5%"></colgroup>
              <tbody>
              ${tarSemAcao.map(t=>`<tr class="${t.estado==='Concluída'?'done-row':''}" style="border-top:0.5px solid var(--bd)">
                <td style="padding:7px 11px"><input type="checkbox" ${t.estado==='Concluída'?'checked':''} onchange="toggleTar('${t.id}',this.checked)" style="cursor:pointer;accent-color:var(--g5);width:14px;height:14px"></td>
                <td style="font-size:12px;font-weight:500;padding:7px 4px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${t.titulo}</td>
                <td style="font-size:11px;color:var(--tx2);padding:7px 4px">${t.responsavel||'—'}</td>
                <td style="font-size:11px;padding:7px 4px">${fd(t.dataLimite)}</td>
                <td style="padding:7px 4px">${estB(t.prioridade)}</td>
                <td style="padding:7px 4px">${estB(t.estado)}</td>
                <td style="padding:7px 4px"><button class="btn btn-sm" onclick="openTarFormEdit('${t.id}')">✎</button></td>
              </tr>`).join('')}
              </tbody>
            </table>
          </div>`:''}
        ${!pAcoes.length&&!tarSemAcao.length?`<div style="padding:12px 24px;font-size:13px;color:var(--tx3)">Sem ações nem plTarefas. <button class="btn btn-sm btn-primary" onclick="openAcaoForm(${p.id})">+ Criar ação</button></div>`:''}
      </div>
    </div>`;
  }).join('');
  // Tarefas completamente sem projeto e sem ação
  const tarSemTudo=plTarefas.filter(t=>!t.acaoId&&!t.projetoId);
  if(tarSemTudo.length){
    const div=document.createElement('div');
    div.style.cssText='margin-top:12px;background:var(--bg2);border:0.5px solid var(--bd);border-radius:var(--rl);overflow:hidden';
    div.innerHTML=`<div style="padding:12px 16px;background:var(--bg3);font-size:13px;font-weight:600;color:var(--tx2);display:flex;align-items:center;gap:8px">
      Tarefas sem projeto nem ação (${tarSemTudo.length})
      <button class="btn btn-sm btn-primary" onclick="openTarForm()" style="font-size:11px">+ Nova tarefa</button>
    </div>
    <table style="width:100%;border-collapse:collapse">
      <tbody>${tarSemTudo.map(t=>`<tr class="${t.estado==='Concluída'?'done-row':''}" style="border-top:0.5px solid var(--bd)">
        <td style="padding:8px 11px;width:4%"><input type="checkbox" ${t.estado==='Concluída'?'checked':''} onchange="toggleTar('${t.id}',this.checked)" style="cursor:pointer;accent-color:var(--g5);width:14px;height:14px"></td>
        <td style="font-size:13px;font-weight:500;padding:8px 4px">${t.titulo}</td>
        <td style="font-size:12px;color:var(--tx2);padding:8px 4px">${t.responsavel||'—'}</td>
        <td style="font-size:12px;padding:8px 4px">${fd(t.dataLimite)}</td>
        <td style="padding:8px 4px">${estB(t.prioridade)}</td>
        <td style="padding:8px 4px">${estB(t.estado)}</td>
        <td style="padding:8px 4px"><button class="btn btn-sm" onclick="openTarFormEdit('${t.id}')">✎</button></td>
      </tr>`).join('')}</tbody>
    </table>`;
    c.appendChild(div);
  }
}
function toggleHier(projId){
  const w=el(`hier-plAcoes-${projId}`);const arr=el(`hier-arrow-${projId}`);
  const open=w.classList.toggle('open');
  arr.style.transform=open?'rotate(180deg)':'';
}
function toggleHierAcao(acaoId){
  const panel=el(`hier-t-${acaoId}`);
  const open=panel.classList.toggle('open');
  if(open){
    const ts=plTarefas.filter(t=>String(t.acaoId)===String(acaoId));
    const kps=plKpdefs.filter(k=>String(k.acaoId)===String(acaoId));
    const content=el(`hier-t-content-${acaoId}`);
    content.innerHTML=`
      ${ts.length?`<div style="font-size:11px;font-weight:600;color:var(--tx2);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">Tarefas (${ts.length})</div>
      <table style="width:100%;border-collapse:collapse;margin-bottom:12px">
        <colgroup><col style="width:4%"><col style="width:40%"><col style="width:20%"><col style="width:15%"><col style="width:11%"><col style="width:10%"></colgroup>
        ${ts.map(t=>`<tr class="${t.estado==='Concluída'?'done-row':''}" style="border-bottom:0.5px solid var(--bd)">
          <td><input type="checkbox" ${t.estado==='Concluída'?'checked':''} onchange="toggleTar('${t.id}',this.checked)" style="cursor:pointer;accent-color:var(--g5);width:14px;height:14px"></td>
          <td style="font-size:12px;padding:6px 4px">${t.titulo}</td>
          <td style="font-size:11px;color:var(--tx2);padding:6px 4px">${t.responsavel||'—'}</td>
          <td style="font-size:11px;padding:6px 4px">${fd(t.dataLimite)}</td>
          <td style="padding:6px 4px">${estB(t.prioridade)}</td>
          <td style="padding:6px 4px">${estB(t.estado)}</td>
        </tr>`).join('')}
      </table>`:'<div style="font-size:12px;color:var(--tx3);margin-bottom:8px">Sem plTarefas nesta ação.</div>'}
      ${kps.length?`<div style="font-size:11px;font-weight:600;color:var(--tx2);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">KPIs (${kps.length})</div>
      ${kps.map(k=>{const lr=plKpregs.filter(r=>r.kpiId==k.id).sort((a,b)=>b.id-a.id)[0];const pct=lr&&k.meta?Math.round((lr.valor/k.meta)*100):null;const c=pct!==null?kpiColor(k,pct):'var(--tx3)';return`<div style="display:flex;align-items:center;gap:10px;padding:4px 0;font-size:12px"><span style="flex:1;font-weight:500">${k.nome}</span><span>${lr?lr.valor+' '+k.uni:'—'}</span><span style="font-weight:700;color:${c}">${pct!==null?pct+'%':'—'}</span></div>`;}).join('')}`:''}
    `;
  }
}
function openTarFormForAcao(acaoId){openTarForm();sv2('ftar-acao',acaoId);}
function openTarFormForProj(projId){openTarForm();/* projeto fica vazio, ação vazia — tarefa vai aparecer na secção sem ação */}

/* ===== AÇÕES CRUD ===== */
let editAcaoId=null;
function openAcaoForm(projId,acaoId){
  editAcaoId=acaoId||null;
  el('frm-acao-ov').classList.add('open');
  el('fa-proj').innerHTML='<option value="">Selecionar...</option>'+plProjetos.map(p=>`<option value="${p.id}">${p.nome}</option>`).join('');
  if(acaoId){
    const a=plAcoes.find(x=>x.id===acaoId);if(!a)return;
    el('frm-acao-title').textContent='Editar ação';
    sv2('fa-nome',a.nome);sv2('fa-proj',a.projetoId||'');sv2('fa-tipo',a.tipoFJ||'Recorrente');
    sv2('fa-dep',a.dep||'');sv2('fa-estado',a.estado||'Ativo');
    sv2('fa-inicio',a.dataInicio||'');sv2('fa-fim',a.dataFim||'');
    sv2('fa-resp',a.responsavel||'');sv2('fa-notas',a.notas||'');
  }else{
    el('frm-acao-title').textContent='Nova ação';
    sv2('fa-nome','');sv2('fa-proj',projId||'');sv2('fa-tipo','Recorrente');
    sv2('fa-dep','');sv2('fa-estado','Ativo');sv2('fa-inicio','');sv2('fa-fim','');
    sv2('fa-resp','');sv2('fa-resp-email','');sv2('fa-notas','');
  }
}
function closeAcaoForm(){el('frm-acao-ov').classList.remove('open');editAcaoId=null;}
async function saveAcao(){
  const nome=gv('fa-nome').trim();if(!nome){alert('Nome obrigatório.');return;}
  const pid=gv('fa-proj');if(!pid){alert('Seleciona o projeto do Factorial.');return;}
  const btn=el('save-acao-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const body={__metadata:{type:'SP.Data.FJ_x005f_AcoesListItem'},Title:nome,TipoAcao:gv('fa-tipo'),Departamento:gv('fa-dep'),Estado:gv('fa-estado'),Notas:gv('fa-notas'),ProjetoIdId:parseInt(pid)};
    if(gv('fa-inicio'))body.DataInicio=gv('fa-inicio')+'T00:00:00Z';
    if(gv('fa-fim'))body.DataFim=gv('fa-fim')+'T00:00:00Z';
    const email=gv('fa-resp-email');
    if(email){const uid=await getSpUID(email);if(uid)body.ResponsavelId=uid;}
    if(editAcaoId)await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Acoes')/items(${editAcaoId})`,body,'MERGE','*');
    else await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Acoes')/items`,body,'POST');
    closeAcaoForm();await loadAcoes();renderHier();populateDrops();showNotif('Ação guardada!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}
async function delAcao(id){
  if(!confirm('Eliminar esta ação?'))return;
  try{await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Acoes')/items(${id})`,null,'DELETE');await loadAcoes();renderHier();populateDrops();showNotif('Eliminado.','success');}
  catch(e){showNotif('Erro: '+e.message,'error');}
}

/* ===== GANTT ===== */
const MONTHS=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const GANTT_COLORS=['#1A5E3B','#2E7D52','#0F6E56','#534AB7','#D85A30','#0C447C','#639922','#BA7517'];
function renderGantt(){
  const dep=gv('ff-g-dep'),ano=parseInt(gv('ff-g-ano')||new Date().getFullYear()),modo=gv('ff-g-modo')||'mensal';
  const projs=plProjetos.filter(p=>!dep||p.dep===dep);
  const lblW=220;
  let cols,colW,labels,unitMs,rangeStart,rangeEnd;
  const today=new Date();today.setHours(0,0,0,0);

  if(modo==='mensal'){
    cols=12;colW=70;
    rangeStart=new Date(ano,0,1);rangeEnd=new Date(ano,11,31);
    labels=MONTHS;
    unitMs=(rangeEnd-rangeStart)/cols;
  }else if(modo==='trimestral'){
    cols=4;colW=180;
    rangeStart=new Date(ano,0,1);rangeEnd=new Date(ano,11,31);
    labels=['Q1','Q2','Q3','Q4'];
    unitMs=(rangeEnd-rangeStart)/cols;
  }else if(modo==='semanal'){
    // Mostrar 13 semanas (trimestre) a partir da semana actual
    const w=isoWeek(today);const wStart=weekRange(Math.max(1,w-1),ano).start;
    rangeStart=wStart;cols=13;colW=80;
    rangeEnd=new Date(rangeStart.getTime()+cols*7*86400000);
    labels=Array.from({length:cols},(_,i)=>{const ws=new Date(rangeStart.getTime()+i*7*86400000);return`S${isoWeek(ws)}`;});
    unitMs=7*86400000;
  }else if(modo==='diario'){
    // 30 dias a partir de hoje
    rangeStart=new Date(today);cols=30;colW=38;
    rangeEnd=new Date(rangeStart.getTime()+cols*86400000);
    labels=Array.from({length:cols},(_,i)=>{const d=new Date(rangeStart.getTime()+i*86400000);return`${d.getDate()}/${d.getMonth()+1}`;});
    unitMs=86400000;
  }

  const totalW=lblW+cols*colW;
  const rangeMs=rangeEnd-rangeStart;
  const todayX=lblW+Math.max(0,Math.min(((today-rangeStart)/rangeMs)*(cols*colW),cols*colW));

  let rows=[];
  projs.forEach((p,pi)=>{
    const color=GANTT_COLORS[pi%GANTT_COLORS.length];
    // Calcular span do projeto a partir das ações
    const pAcoes=plAcoes.filter(a=>String(a.projetoId)===String(p.id));
    let ps=pAcoes.reduce((min,a)=>a.dataInicio&&new Date(a.dataInicio)<min?new Date(a.dataInicio):min,new Date(9999,0,1));
    let pe=pAcoes.reduce((max,a)=>a.dataFim&&new Date(a.dataFim)>max?new Date(a.dataFim):max,new Date(0));
    if(ps.getFullYear()===9999)ps=null;if(pe.getFullYear()===1970)pe=null;
    rows.push({type:'proj',label:p.nome,color,start:ps,end:pe,id:p.id});
    pAcoes.forEach(a=>{
      const s=a.dataInicio?new Date(a.dataInicio):null;
      const e=a.dataFim?new Date(a.dataFim):null;
      rows.push({type:'acao',label:a.nome,color,start:s,end:e,id:a.id});
    });
  });

  function barX(d){return Math.max(0,((d-rangeStart)/rangeMs)*(cols*colW));}
  function barW(s,e){return Math.max(4,barX(e)-barX(s));}

  const headerHtml=`<div class="gantt-header">
    <div class="gantt-lbl-col" style="width:${lblW}px">${modo==='semanal'?'Semanas':modo==='diario'?'Dias':ano}</div>
    <div class="gantt-months" style="width:${cols*colW}px">
      ${labels.map((l,i)=>`<div class="gantt-month" style="width:${colW}px;font-size:${colW<50?'9px':'11px'}">${l}</div>`).join('')}
    </div>
  </div>`;

  const rowsHtml=rows.map(row=>{
    let barHtml='';
    if(row.start&&row.end&&row.end>rangeStart&&row.start<rangeEnd){
      const s=Math.max(0,barX(row.start));
      const w2=Math.min(barW(row.start,row.end),cols*colW-s);
      const opacity=row.type==='proj'?0.35:0.8;
      barHtml=`<div class="gantt-bar" style="left:${s}px;width:${w2}px;background:${row.color};opacity:${opacity};height:${row.type==='proj'?'10px':'18px'}" title="${row.label}"></div>`;
    }
    return`<div class="gantt-row ${row.type==='proj'?'proj-row':''}">
      <div class="gantt-row-lbl ${row.type==='acao'?'sub':''}" style="width:${lblW}px" title="${row.label}">${row.label}</div>
      <div class="gantt-cells" style="width:${cols*colW}px">
        ${Array.from({length:cols},(_,i)=>`<div style="position:absolute;left:${i*colW}px;top:0;bottom:0;width:0.5px;background:var(--bd)"></div>`).join('')}
        <div class="gantt-today" style="left:${todayX-lblW}px"></div>
        ${barHtml}
      </div>
    </div>`;
  }).join('');

  el('gantt-container').innerHTML=`<div class="gantt-wrap" style="overflow-x:auto"><div style="min-width:${totalW}px">${headerHtml}${rowsHtml||'<div class="empty">Sem dados. Adiciona ações com datas de início e fim.</div>'}</div></div>`;
}

/* ===== PLANO TRIMESTRAL ===== */
function genTrimTit(){
  const pid=gv('ft-proj');const p=plProjetos.find(x=>String(x.id)===pid);
  sv2('ft-titulo',p?`${gv('ft-trim')} ${gv('ft-ano')} — ${p.nome}`:'');
}
function renderTrim(){
  const tbody=el('tbody-trim');
  if(!tbody)return;
  const q=gv('ff-t-q'),ano=gv('ff-t-ano'),dep=gv('ff-t-dep'),aprov=gv('ff-t-aprov');
  const list=plPlanoTrim.filter(t=>{
    if(q&&t.trim!==q)return false;if(ano&&String(t.ano)!==ano)return false;
    if(dep&&t.dep!==dep)return false;if(aprov&&t.aprovacao!==aprov)return false;return true;
  });
  tbody.innerHTML=list.map(t=>{
    const aprovCls={'Rascunho':'b-rascunho','Submetido':'b-submetido','Aprovado Dir.':'b-aprov-dir','Aprovado Gestão':'b-aprov-gest'}[t.aprovacao]||'b-pfazer';
    const canAprovDir=t.aprovacao==='Submetido'&&isDirOf(t.dep);
    const canAprovGest=t.aprovacao==='Aprovado Dir.'&&isGestao();
    const canSubmit=t.aprovacao==='Rascunho';
    return`<tr>
      <td style="font-weight:500" title="${pn(t.projetoId)}">${pn(t.projetoId)}</td>
      <td title="${an(t.acaoId)}">${an(t.acaoId)}</td>
      <td>${badgeHTML('b-emcurso',t.trim)}</td><td>${t.ano||'—'}</td>
      <td>${depB(t.dep)}</td><td>${estB(t.estado)}</td>
      <td style="text-align:center;font-weight:700">${t.tarPrev||0}</td>
      <td>${badgeHTML(aprovCls,t.aprovacao||'Rascunho')}
        ${canSubmit?`<button class="btn btn-sm" style="margin-left:4px" onclick="aprovTrim(${t.id},'Submetido')">Submeter</button>`:''}
        ${canAprovDir?`<button class="btn btn-sm btn-aprov" onclick="aprovTrim(${t.id},'Aprovado Dir.')">Aprovar</button>`:''}
        ${canAprovGest?`<button class="btn btn-sm btn-aprov" onclick="aprovTrim(${t.id},'Aprovado Gestão')">Aprovar Gestão</button>`:''}
      </td>
      <td><div class="actions"><button class="btn btn-sm" onclick="editTrim(${t.id})">Editar</button><button class="btn btn-sm btn-danger" onclick="delTrim(${t.id})">✕</button></div></td>
    </tr>`;
  }).join('')||`<tr><td colspan="9"><div class="empty">Sem plano trimestral.</div></td></tr>`;
}
async function aprovTrim(id,novoEstado){
  try{
    const body={__metadata:{type:'SP.Data.FJ_x005f_PlanoTrimestralListItem'},EstadoAprov:novoEstado};
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_PlanoTrimestral')/items(${id})`,body,'MERGE','*');
    const t=plPlanoTrim.find(x=>x.id===id);if(t)t.aprovacao=novoEstado;
    renderTrim();showNotif('Estado atualizado!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
}
function openTrimForm(id){
  plEditTrimId=id||null;el('frm-trim-ov').classList.add('open');
  el('ft-proj').innerHTML='<option value="">Selecionar...</option>'+plProjetos.map(p=>`<option value="${p.id}">${p.nome}</option>`).join('');
  if(id){
    const t=plPlanoTrim.find(x=>x.id===id);if(!t)return;
    el('frm-trim-title').textContent='Editar compromisso';
    sv2('ft-proj',t.projetoId||'');loadAcoesForProj(t.projetoId,'ft-acao',t.acaoId);
    sv2('ft-trim',t.trim);sv2('ft-ano',t.ano||'');sv2('ft-dep',t.dep||'');
    sv2('ft-est',t.estado);sv2('ft-tarprev',t.tarPrev||0);sv2('ft-notas',t.notas||'');sv2('ft-titulo',t.titulo||'');
  }else{
    el('frm-trim-title').textContent='Novo compromisso';
    sv2('ft-trim','Q'+Math.ceil((new Date().getMonth()+1)/3));sv2('ft-est','Previsto');
    sv2('ft-tarprev',0);sv2('ft-notas','');sv2('ft-dep','');sv2('ft-titulo','');
    loadAcoesForProj('','ft-acao');
  }
}
function editTrim(id){openTrimForm(id);}
function closeTrimForm(){el('frm-trim-ov').classList.remove('open');plEditTrimId=null;}
async function saveTrim(){
  const pid=gv('ft-proj');if(!pid){alert('Seleciona um projeto.');return;}
  const btn=el('save-trim-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const trim=gv('ft-trim'),ano=parseInt(gv('ft-ano')),pnome=plProjetos.find(p=>String(p.id)===pid)?.nome||'';
    const titulo=gv('ft-titulo')||`${trim} ${ano} — ${pnome}`;
    const body={__metadata:{type:'SP.Data.FJ_x005f_PlanoTrimestralListItem'},Title:titulo,Trimestre:trim,Departamento:gv('ft-dep'),Estado:gv('ft-est'),Notas:gv('ft-notas'),ProjetoIdId:parseInt(pid),EstadoAprov:'Rascunho'};
    if(ano)body.AnoFJ=ano;const aid=gv('ft-acao');if(aid)body.AcaoIdId=parseInt(aid);
    const tp=parseInt(gv('ft-tarprev'));if(tp)body.TarefasPrevistas=tp;
    if(plEditTrimId)await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_PlanoTrimestral')/items(${plEditTrimId})`,body,'MERGE','*');
    else await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_PlanoTrimestral')/items`,body,'POST');
    closeTrimForm();await loadPlanoTrim();renderTrim();showNotif('Guardado!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}
async function delTrim(id){
  if(!confirm('Eliminar?'))return;
  try{await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_PlanoTrimestral')/items(${id})`,null,'DELETE');await loadPlanoTrim();renderTrim();showNotif('Eliminado.','success');}
  catch(e){showNotif('Erro: '+e.message,'error');}
}

/* ===== PLANO SEMANAL ===== */
function setView(v){
  plSemView=v;
  ['pes','eq','col'].forEach(k=>el('vt-'+k)?.classList.toggle('active',k===v[0]+v[1]+(v==='colaborador'?'l':'')));
  el('vt-pes').classList.toggle('active',v==='pessoal');
  el('vt-eq').classList.toggle('active',v==='equipa');
  el('vt-col').classList.toggle('active',v==='colaborador');
  el('ff-s-dep').style.display=v==='equipa'?'block':'none';
  el('col-picker').style.display=v==='colaborador'?'block':'none';
  renderSem();
}
function chgWeek(d){plViewWeek+=d;if(plViewWeek<1){plViewWeek=52;plViewYear--;}else if(plViewWeek>52){plViewWeek=1;plViewYear++;}renderSem();}
function renderSem(){
  const elWl=el('week-label'),elStats=el('stats-sem'),tbody=el('tbody-sem');
  if(!elWl||!elStats||!tbody)return;
  elWl.textContent=weekLabel(plViewWeek,plViewYear);
  const wr=weekRange(plViewWeek,plViewYear);
  const q=(gv('q-sem')||'').toLowerCase(),est=gv('ff-s-est'),dep=gv('ff-s-dep'),orig=gv('ff-s-orig');
  const myName=(user?.name||'').toLowerCase();
  const colEmail=gv('col-email');
  const list=plTarefas.filter(t=>{
    const byW=t.semana==plViewWeek&&t.anoFJ==plViewYear;
    const byD=t.dataLimite&&new Date(t.dataLimite+'T12:00:00')>=wr.start&&new Date(t.dataLimite+'T12:00:00')<=wr.end;
    if(!byW&&!byD)return false;
    if(plSemView==='pessoal'&&myName&&t.responsavel&&!t.responsavel.toLowerCase().split(' ').some(w=>w.length>2&&myName.includes(w)))return false;
    if(plSemView==='colaborador'&&colEmail&&t.responsavelEmail!==colEmail)return false;
    if(dep&&t.dep!==dep)return false;
    if(est==='todo'&&t.estado!=='todo'&&t.estado!=='Por fazer')return false;
    if(est&&est!=='todo'&&t.estado!==est)return false;
    if(orig&&t.notas&&!t.notas.toLowerCase().includes(orig.toLowerCase()))return false;
    if(q&&!t.titulo.toLowerCase().includes(q))return false;
    return true;
  });
  const total=list.length,conc=list.filter(t=>t.estado==='Concluída').length,emcurso=list.filter(t=>t.estado==='Em curso').length;
  const atras=list.filter(t=>t.estado!=='Concluída'&&t.dataLimite&&new Date(t.dataLimite)<new Date()).length;
  elStats.innerHTML=`
    <div class="stat"><span class="stat-n">${total}</span><span class="stat-l">Total</span></div>
    <div class="stat"><span class="stat-n" style="color:var(--g5)">${conc}</span><span class="stat-l">Concluídas</span></div>
    <div class="stat"><span class="stat-n" style="color:var(--bld)">${emcurso}</span><span class="stat-l">Em curso</span></div>
    <div class="stat"><span class="stat-n" style="color:var(--red)">${atras}</span><span class="stat-l">Atrasadas</span></div>`;
  const origCls=o=>o==='Ad-hoc'?'b-adhoc':o==='Reunião'?'b-reuniao':'b-plano';
  tbody.innerHTML=list.map(t=>{
    const conc2=t.estado==='Concluída';
    const acaoNome=an(t.acaoId);
    const hasAtt=t.attachments?.length>0;
    return`<tr class="${conc2?'done-row':''}">
      <td><input type="checkbox" ${conc2?'checked':''} onchange="toggleTar('${t.id}',this.checked)" style="cursor:pointer;accent-color:var(--g5);width:14px;height:14px"></td>
      <td style="font-weight:500" title="${t.titulo}">${t.titulo}${hasAtt?` <span style="font-size:10px;background:var(--bl);color:var(--bld);border-radius:4px;padding:1px 5px">📎${t.attachments.length}</span>`:''}</td>
      <td style="font-size:12px;color:var(--tx2)" title="${acaoNome}">${acaoNome}</td>
      <td>${t.responsavel||'—'}</td>
      <td>${fd(t.dataLimite)}</td>
      <td>${estB(t.prioridade)}</td>
      <td>${badgeHTML(origCls(t.origem),t.origem||'Plano')}</td>
      <td><div class="actions">
        <button class="btn btn-sm" onclick="openTarFormEdit('${t.id}')">Editar</button>
      </div></td>
    </tr>`;
  }).join('')||`<tr><td colspan="8"><div class="empty">Sem plTarefas para esta semana.</div></td></tr>`;
  tbody.innerHTML=tbody.innerHTML||'';
  const elCount=el('tcount-sem');
  if(elCount)elCount.textContent=`${list.length} tarefa${list.length!==1?'s':''}`;
}
async function toggleTar(compId,checked){
  const t=plTarefas.find(x=>x.id===compId);if(!t)return;
  const novoEst=checked?'Concluída':'Em curso';
  try{
    const body={__metadata:{type:`SP.Data.${t.lista==='FJ_Tarefas'?'FJ_x005f_TarefasListItem':'FJ_x005f_TarefasListItem'}`},Estado:novoEst};
    await spWrite(`${SP}/_api/web/lists/getbytitle('${t.lista}')/items(${t.spId})`,body,'MERGE','*');
    t.estado=novoEst;renderSem();renderHier();renderDash();showNotif('✓ Atualizado!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
}
/* Tarefa form */
function checkOrigem(v){el('tar-warn').style.display=v==='Factorial'?'block':'none';}
function openTarForm(){
  plEditTarId=null;el('frm-tar-ov').classList.add('open');
  el('frm-tar-title').textContent='Nova tarefa';
  el('tar-warn').style.display='none';
  sv2('ftar-desc','');sv2('ftar-origem','Ad-hoc');sv2('ftar-acao','');
  sv2('ftar-resp','');sv2('ftar-email','');sv2('ftar-data','');
  sv2('ftar-prior','Média');sv2('ftar-estado','todo');
  el('ftar-existing-attach').innerHTML='';
}
function openTarFormEdit(compId){
  const t=plTarefas.find(x=>x.id===compId);if(!t)return;
  plEditTarId=compId;el('frm-tar-ov').classList.add('open');
  el('frm-tar-title').textContent='Editar tarefa';
  el('tar-warn').style.display='none';
  sv2('ftar-desc',t.titulo);sv2('ftar-origem',t.origem||'Ad-hoc');sv2('ftar-acao',t.acaoId||'');
  sv2('ftar-resp',t.responsavel||'');sv2('ftar-email',t.responsavelEmail||'');
  sv2('ftar-data',t.dataLimite||'');sv2('ftar-prior',t.prioridade||'Média');sv2('ftar-estado',t.estado||'todo');
  el('ftar-existing-attach').innerHTML=t.attachments?.length?t.attachments.map(a=>`<div class="attach-item"><a href="${SP}${a.url}" target="_blank">📎 ${a.name}</a></div>`).join(''):'';
}
function closeTarForm(){el('frm-tar-ov').classList.remove('open');plEditTarId=null;}
async function saveTarefa(){
  if(gv('ftar-origem')==='Factorial'){showNotif('As plTarefas do Factorial devem ser criadas no Factorial.','warn');return;}
  const desc=gv('ftar-desc').trim();
  if(!desc){alert('Descrição obrigatória.');return;}
  const btn=el('save-tar-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const data=gv('ftar-data');
    const semana=data?isoWeek(new Date(data+'T12:00:00')):plViewWeek;
    const anoFJ=data?new Date(data+'T12:00:00').getFullYear():plViewYear;
    const body={__metadata:{type:'SP.Data.FJ_x005f_TarefasListItem'},
      Title:desc,
      Notas:'Origem: '+gv('ftar-origem'),
      Prioridade:gv('ftar-prior'),
      Semana:semana,
      AnoFJ:anoFJ,
      Estado:gv('ftar-estado'),
      Departamento:gv('ftar-dep')||''
    };
    if(data)body.DataLimite=data+'T00:00:00Z';
    const aid=gv('ftar-acao');if(aid)body.AcaoIdId=parseInt(aid);
    let itemId=null;
    if(plEditTarId){
      const t=plTarefas.find(x=>x.id===plEditTarId);
      if(t){await spWrite(`${SP}/_api/web/lists/getbytitle('${t.lista}')/items(${t.spId})`,body,'MERGE','*');itemId=t.spId;}
    }else{
      const r=await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items`,body,'POST');
      try{const d=await r.json();itemId=d.d?.Id;}catch(e){}
    }
    closeTarForm();await loadTarefas();renderSem();renderHier();renderDash();showNotif('Tarefa guardada!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');console.error(e);}
  btn.textContent='Guardar';btn.disabled=false;
}

/* ===== KPIs ===== */
function setKTab(t){
  const elDef=el('tab-def'),elReg=el('tab-reg'),elKpDef=el('kp-def'),elKpReg=el('kp-reg');
  if(elDef)elDef.classList.toggle('active',t==='def');
  if(elReg)elReg.classList.toggle('active',t==='reg');
  if(elKpDef)elKpDef.classList.toggle('active',t==='def');
  if(elKpReg)elKpReg.classList.toggle('active',t==='reg');
}
function renderKdefs(){
  const tbody=el('tbody-kdef');
  if(!tbody)return;
  const proj=gv('ff-kd-proj'),tipo=gv('ff-kd-tipo');
  const list=plKpdefs.filter(k=>{if(proj&&String(k.projetoId)!==proj)return false;if(tipo&&k.tipo!==tipo)return false;return true;});
  tbody.innerHTML=list.map(k=>`<tr>
    <td style="font-weight:500" title="${k.nome}">${k.nome}</td>
    <td style="font-size:12px" title="${an(k.acaoId)}">${an(k.acaoId)}</td>
    <td>${estB(k.tipo)}</td>
    <td>${badgeHTML(k.dir==='limite'?'b-limite':'b-meta',k.dir==='limite'?'↓ Limite':'↑ Meta')}</td>
    <td style="font-weight:700">${k.meta||0}</td>
    <td>${k.uni||'—'}</td>
    <td style="font-size:12px">${k.freq||'—'}</td>
    <td style="font-size:12px">${k.resp||'—'}</td>
    <td><div class="actions"><button class="btn btn-sm" onclick="editKdef(${k.id})">Editar</button><button class="btn btn-sm btn-danger" onclick="delKdef(${k.id})">✕</button></div></td>
  </tr>`).join('')||`<tr><td colspan="9"><div class="empty">Sem KPIs.</div></td></tr>`;
}
function openKdefForm(id){
  plEditKdefId=id||null;el('frm-kdef-ov').classList.add('open');
  el('fk-proj').innerHTML='<option value="">Nenhum</option>'+plProjetos.map(p=>`<option value="${p.id}">${p.nome}</option>`).join('');
  el('fk-proj').onchange=function(){loadAcoesForProj(this.value,'fk-acao');};
  if(id){
    const k=plKpdefs.find(x=>x.id===id);if(!k)return;
    el('frm-kdef-title').textContent='Editar KPI';
    sv2('fk-nome',k.nome);sv2('fk-proj',k.projetoId||'');loadAcoesForProj(k.projetoId,'fk-acao',k.acaoId);
    sv2('fk-tipo',k.tipo);sv2('fk-dir',k.dir||'meta');sv2('fk-uni',k.uni||'');
    sv2('fk-meta',k.meta||0);sv2('fk-freq',k.freq);sv2('fk-resp',k.resp||'');
    sv2('fk-ativo',k.ativo?'1':'0');sv2('fk-desc',k.desc||'');
  }else{
    el('frm-kdef-title').textContent='Novo KPI';
    ['fk-nome','fk-uni','fk-resp','fk-desc'].forEach(i=>sv2(i,''));
    sv2('fk-tipo','Manual');sv2('fk-dir','meta');sv2('fk-meta',0);sv2('fk-freq','Mensal');sv2('fk-ativo','1');
    loadAcoesForProj('','fk-acao');
  }
}
function editKdef(id){openKdefForm(id);}
function closeKdefForm(){el('frm-kdef-ov').classList.remove('open');plEditKdefId=null;}
async function saveKdef(){
  const nome=gv('fk-nome').trim();if(!nome){alert('Nome obrigatório.');return;}
  const btn=el('save-kdef-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const body={__metadata:{type:'SP.Data.FJ_x005f_KPI_x005f_DefinicoesListItem'},Title:nome,TipoKPI:gv('fk-tipo'),Direcao:gv('fk-dir'),Unidade:gv('fk-uni'),Frequencia:gv('fk-freq'),Ativo:gv('fk-ativo')==='1',Descricao:gv('fk-desc')};
    const respEmail=gv('fk-resp');if(respEmail){const uid=await getSpUID(respEmail);if(uid)body.ResponsavelId=uid;}
    const m=parseFloat(gv('fk-meta'));if(!isNaN(m))body.Meta=m;
    const pid=gv('fk-proj');if(pid)body.ProjetoIdId=parseInt(pid);
    const aid=gv('fk-acao');if(aid)body.AcaoIdId=parseInt(aid);
    if(plEditKdefId)await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_KPI_Definicoes')/items(${plEditKdefId})`,body,'MERGE','*');
    else await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_KPI_Definicoes')/items`,body,'POST');
    closeKdefForm();await loadKPIDefs();renderKdefs();populateDrops();showNotif('KPI guardado!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}
async function delKdef(id){
  if(!confirm('Eliminar?'))return;
  try{await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_KPI_Definicoes')/items(${id})`,null,'DELETE');await loadKPIDefs();renderKdefs();showNotif('Eliminado.','success');}
  catch(e){showNotif('Erro: '+e.message,'error');}
}
function renderKregs(){
  const tbody=el('tbody-kreg');
  if(!tbody)return; // Safety check
  const kpiId=gv('ff-kr-kpi'),ano=gv('ff-kr-ano');
  const list=plKpregs.filter(r=>{if(kpiId&&String(r.kpiId)!==kpiId)return false;if(ano&&String(r.ano)!==ano)return false;return true;});
  tbody.innerHTML=list.map(r=>{
    const k=plKpdefs.find(x=>String(x.id)===String(r.kpiId));
    const meta=k?.meta||null;const pct=meta&&r.valor!==null?Math.round((r.valor/meta)*100):null;
    const c=pct!==null?kpiColor(k,pct):'var(--tx)';
    const est=pct!==null?kpiEstado(k,pct):'—';
    return`<tr>
      <td style="font-weight:500;font-size:12px">${kn(r.kpiId)}</td>
      <td>${r.periodo||'—'}</td>
      <td style="font-weight:700">${r.valor} ${k?.uni||''}</td>
      <td>${meta!==null?meta+' '+(k?.uni||''):'—'}</td>
      <td style="font-size:12px">${est}</td>
      <td style="font-weight:700;color:${c}">${pct!==null?pct+'%':'—'}</td>
      <td style="font-size:12px">${fd(r.dataReg)}</td>
      <td style="font-size:12px">${r.inseridoPor||'—'}</td>
    </tr>`;
  }).join('')||`<tr><td colspan="8"><div class="empty">Sem registos.</div></td></tr>`;
}
function openKregForm(){
  el('frm-kreg-ov').classList.add('open');
  el('fr-kpi').innerHTML='<option value="">Selecionar KPI...</option>'+plKpdefs.filter(k=>k.ativo).map(k=>`<option value="${k.id}">${k.nome}</option>`).join('');
  sv2('fr-valor',0);sv2('fr-sem',plViewWeek);sv2('fr-data',today());sv2('fr-notas','');
}
function closeKregForm(){el('frm-kreg-ov').classList.remove('open');}
async function saveKreg(){
  const kpiId=gv('fr-kpi');if(!kpiId){alert('Seleciona um KPI.');return;}
  const btn=el('save-kreg-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const trim=gv('fr-trim'),ano=parseInt(gv('fr-ano')||new Date().getFullYear()),sem=parseInt(gv('fr-sem'));
    const periodo=sem?`${ano}-W${String(sem).padStart(2,'0')}`:`${ano}-${trim}`;
    const body={__metadata:{type:'SP.Data.FJ_x005f_KPI_x005f_RegistosListItem'},Title:periodo,KPIId:parseInt(kpiId),Trimestre:trim,Notas:gv('fr-notas')};
    const v=parseFloat(gv('fr-valor'));if(!isNaN(v))body.Valor=v;
    if(ano)body.AnoFJ=ano;if(sem)body.Semana=sem;
    if(gv('fr-data'))body.DataRegisto=gv('fr-data')+'T00:00:00Z';
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_KPI_Registos')/items`,body,'POST');
    closeKregForm();await loadKPIRegs();renderKregs();renderDash();showNotif('Registo guardado!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}

/* ===== RELATÓRIOS ===== */
function gerarRelS(){
  const ano=parseInt(gv('rl-s-ano')),sem=parseInt(gv('rl-s-sem')),dep=gv('rl-s-dep');
  const wr=weekRange(sem,ano);
  const list=plTarefas.filter(t=>{
    const byW=t.semana==sem&&t.anoFJ==ano;
    const byD=t.dataLimite&&new Date(t.dataLimite+'T12:00:00')>=wr.start&&new Date(t.dataLimite+'T12:00:00')<=wr.end;
    if(!byW&&!byD)return false;if(dep&&t.dep!==dep)return false;return true;
  });
  const conc=list.filter(t=>t.estado==='Concluída'),nc=list.filter(t=>t.estado!=='Concluída');
  const pct=list.length?Math.round((conc.length/list.length)*100):0;
  const pc=pct>=80?'var(--g5)':pct>=50?'var(--amb2)':'var(--red)';
  el('rl-s-content').innerHTML=`
    <div style="color:var(--tx2);font-size:13px;margin-bottom:1rem">${weekLabel(sem,ano)}${dep?' · '+dep:''}</div>
    <div class="stats" style="margin-bottom:1.5rem">
      <div class="stat"><span class="stat-n">${list.length}</span><span class="stat-l">Total plTarefas</span></div>
      <div class="stat"><span class="stat-n" style="color:var(--g5)">${conc.length}</span><span class="stat-l">Concluídas</span></div>
      <div class="stat"><span class="stat-n" style="color:var(--red)">${nc.length}</span><span class="stat-l">Por concluir</span></div>
      <div class="stat"><span class="stat-n" style="color:${pc}">${pct}%</span><span class="stat-l">Cumprimento</span></div>
    </div>
    ${nc.length?`<div style="font-size:13px;font-weight:600;margin-bottom:.5rem">Tarefas não concluídas</div>
    <div class="table-wrap" style="margin-bottom:1.5rem"><table>
      <colgroup><col style="width:38%"><col style="width:18%"><col style="width:18%"><col style="width:14%"><col style="width:12%"></colgroup>
      <thead><tr><th>Tarefa</th><th>Responsável</th><th>Ação</th><th>Data prevista</th><th>Estado</th></tr></thead>
      <tbody>${nc.map(t=>`<tr><td style="font-weight:500">${t.titulo}</td><td>${t.responsavel||'—'}</td><td style="font-size:12px">${an(t.acaoId)}</td><td>${fd(t.dataLimite)}</td><td>${estB(t.estado)}</td></tr>`).join('')}</tbody>
    </table></div>`:`<div style="background:var(--g4);border-radius:var(--r);padding:12px 16px;margin-bottom:1.5rem;font-size:13px;color:var(--g6);font-weight:500">✓ Todas as plTarefas concluídas!</div>`}
    <div style="font-size:13px;font-weight:600;margin-bottom:.5rem">KPIs semanais</div>
    <div class="table-wrap"><table>
      <colgroup><col style="width:40%"><col style="width:20%"><col style="width:20%"><col style="width:20%"></colgroup>
      <thead><tr><th>KPI</th><th>Valor</th><th>Objetivo</th><th>Estado</th></tr></thead>
      <tbody>${plKpdefs.filter(k=>k.ativo&&k.freq==='Semanal').map(k=>{const r=plKpregs.find(x=>x.kpiId==k.id&&x.sem==sem&&x.ano==ano);const pct2=r&&k.meta?Math.round((r.valor/k.meta)*100):null;return`<tr><td>${k.nome}</td><td>${r?r.valor+' '+k.uni:'—'}</td><td>${k.meta} ${k.uni}</td><td style="color:${pct2?kpiColor(k,pct2):'var(--tx3)'}">${pct2!==null?kpiEstado(k,pct2):'—'}</td></tr>`;}).join('')||'<tr><td colspan="4"><div class="empty" style="padding:.75rem">Sem KPIs semanais.</div></td></tr>'}</tbody>
    </table></div>`;
}
function gerarRelT(){
  const ano=parseInt(gv('rl-t-ano')),trim=gv('rl-t-q'),dep=gv('rl-t-dep');
  const planos=plPlanoTrim.filter(p=>p.trim===trim&&p.ano==ano&&(!dep||p.dep===dep));
  const tarPrev=planos.reduce((a,p)=>a+(p.tarPrev||0),0);
  const qMonths={'Q1':[0,1,2],'Q2':[3,4,5],'Q3':[6,7,8],'Q4':[9,10,11]};
  const tarConc=plTarefas.filter(t=>{
    if(t.estado!=='Concluída')return false;const dt=t.dataLimite?new Date(t.dataLimite):null;if(!dt)return false;
    const m=dt.getMonth();const q2=m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
    return q2===trim&&dt.getFullYear()===ano&&(!dep||t.dep===dep);
  }).length;
  const pct=tarPrev?Math.round((tarConc/tarPrev)*100):0;
  const pc=pct>=80?'var(--g5)':pct>=50?'var(--amb2)':'var(--red)';
  el('rl-t-content').innerHTML=`
    <div style="color:var(--tx2);font-size:13px;margin-bottom:1rem">${trim} ${ano}${dep?' · '+dep:''}</div>
    <div class="stats" style="margin-bottom:1.5rem">
      <div class="stat"><span class="stat-n">${planos.length}</span><span class="stat-l">Ações comprometidas</span></div>
      <div class="stat"><span class="stat-n">${tarPrev}</span><span class="stat-l">Tarefas previstas</span></div>
      <div class="stat"><span class="stat-n" style="color:var(--g5)">${tarConc}</span><span class="stat-l">Tarefas concluídas</span></div>
      <div class="stat"><span class="stat-n" style="color:${pc}">${pct}%</span><span class="stat-l">Cumprimento</span></div>
    </div>
    <div style="font-size:13px;font-weight:600;margin-bottom:.5rem">Ações do trimestre</div>
    <div class="table-wrap" style="margin-bottom:1.5rem"><table>
      <colgroup><col style="width:25%"><col style="width:25%"><col style="width:14%"><col style="width:10%"><col style="width:12%"><col style="width:14%"></colgroup>
      <thead><tr><th>Projeto</th><th>Ação</th><th>Departamento</th><th>Tar.Prev.</th><th>Estado</th><th>Aprovação</th></tr></thead>
      <tbody>${planos.map(p=>`<tr><td style="font-weight:500">${pn(p.projetoId)}</td><td>${an(p.acaoId)}</td><td>${depB(p.dep)}</td><td style="text-align:center;font-weight:700">${p.tarPrev||0}</td><td>${estB(p.estado)}</td><td>${estB(p.aprovacao||'Rascunho')}</td></tr>`).join('')||'<tr><td colspan="6"><div class="empty">Sem ações.</div></td></tr>'}</tbody>
    </table></div>
    <div style="font-size:13px;font-weight:600;margin-bottom:.5rem">KPIs do trimestre</div>
    <div class="table-wrap"><table>
      <colgroup><col style="width:30%"><col style="width:15%"><col style="width:15%"><col style="width:10%"><col style="width:30%"></colgroup>
      <thead><tr><th>KPI</th><th>Último valor</th><th>Objetivo</th><th>Estado</th><th>Progresso</th></tr></thead>
      <tbody>${plKpdefs.filter(k=>k.ativo).map(k=>{
        const r=plKpregs.filter(x=>x.kpiId==k.id&&x.trim===trim&&x.ano==ano).sort((a,b)=>b.id-a.id)[0];
        const pct2=r&&k.meta?Math.round((r.valor/k.meta)*100):null;
        const c=pct2!==null?kpiColor(k,pct2):'var(--tx3)';
        const barW=k.dir==='limite'?(pct2!==null?Math.min(100,200-pct2):0):(pct2!==null?Math.min(pct2,100):0);
        return`<tr>
          <td style="font-size:12px;font-weight:500">${k.nome}</td>
          <td>${r?r.valor+' '+k.uni:'—'}</td>
          <td>${k.meta} ${k.uni}</td>
          <td style="font-size:12px;color:${c}">${pct2!==null?kpiEstado(k,pct2):'—'}</td>
          <td><div style="display:flex;align-items:center;gap:8px"><div class="prog-bg" style="flex:1"><div class="prog-fill" style="width:${barW}%;background:${c}"></div></div><span style="font-size:12px;font-weight:700;color:${c};min-width:36px">${pct2!==null?pct2+'%':'—'}</span></div></td>
        </tr>`;
      }).join('')||'<tr><td colspan="5"><div class="empty" style="padding:.75rem">Sem KPIs.</div></td></tr>'}</tbody>
    </table></div>`;
}

/* ===== PEOPLE SEARCH ===== */
async function plSearchPeople(q,sugId,inputId,emailId){
  if(q.length<2){el(sugId)?.classList.remove('open');return;}
  try{
    await getGrTok();
    const r=await fetch(`https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${encodeURIComponent(q)}') or startswith(mail,'${encodeURIComponent(q)}')&$select=displayName,mail&$top=8`,{headers:{Authorization:`Bearer ${graphToken}`}});
    const d=await r.json();const s=el(sugId);if(!s)return;
    if(!d.value?.length){s.classList.remove('open');return;}
    s.innerHTML=d.value.map(p=>`<div class="ppl-item" onclick="plSelPerson('${inputId}','${sugId}','${p.displayName.replace(/'/g,"\\'")}','${emailId||''}','${(p.mail||'').replace(/'/g,"\\'")}')">
      <div class="ppl-av">${p.displayName.split(' ').map(n=>n[0]).slice(0,2).join('').toUpperCase()}</div>
      <div><div style="font-weight:500">${p.displayName}</div><div style="font-size:11px;color:var(--tx3)">${p.mail||''}</div></div>
    </div>`).join('');
    s.classList.add('open');
  }catch(e){console.warn(e);}
}
function plSelPerson(inputId,sugId,name,emailId,email){
  sv2(inputId,name);if(emailId)sv2(emailId,email);
  el(sugId)?.classList.remove('open');
  if(emailId==='col-email')renderSem();
}
document.addEventListener('click',e=>{if(!e.target.closest('.people-wrap'))document.querySelectorAll('.people-sugg').forEach(s=>s.classList.remove('open'));});

/* ===== CLOSE OVERLAYS ON BACKDROP ===== */
document.addEventListener('click',e=>{
  if(e.target===el('frm-tar-ov'))closeTarForm();
  if(e.target===el('frm-trim-ov'))closeTrimForm();
  if(e.target===el('frm-kdef-ov'))closeKdefForm();
  if(e.target===el('frm-kreg-ov'))closeKregForm();
  if(e.target===el('frm-acao-ov'))closeAcaoForm();
});

/* ===== BOOT ===== */
// init() call removed to avoid duplication

function plInit(){
  const now=new Date();
  plViewWeek=isoWeek(now);plViewYear=now.getFullYear();
  // Use shared projetos
  if(sharedProjetos.length) plProjetos=sharedProjetos;
  plLoadAll();
}



