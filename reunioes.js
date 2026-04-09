// ═══ REUNIÕES MODULE ═══
// ASSUNTOS
async function loadItems(){
  showNotif('A carregar...','info');
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Assuntos')/items?$select=*,AttachmentFiles&$expand=AttachmentFiles&$orderby=Created desc&$top=500`);
    items=d.d.results.map(r=>({
      id:r.Id,etag:r.__metadata?.etag||'*',
      titulo:r.Title||'',reuniao:r.TipoReuniao||'',
      data:r.DataReuniao?r.DataReuniao.split('T')[0]:'',
      tipo:r.TipoAssunto||'',responsavel:r.Responsavel?.Title||'',
      estado:r.Estado||'Em aberto',encaminhamento:r.Encaminhamento||'—',
      decisao:r.Decisao||'Pendente',confidencial:r.Confidencial||false,
      atrasado:r.Atrasado||false,notas:r.Notas||'',proposta:r.Proposta||'',
      historicoDecisoes:r.HistoricoDecisoes||'',
      nivelOrigem:r.NivelOrigem||'',nivelAtual:r.NivelAtual||'',
      attachments:(r.AttachmentFiles?.results||[]).map(a=>({name:a.FileName,url:a.ServerRelativeUrl}))
    }));
    renderStats();renderTable();renderCATopics();
    showNotif('Carregado!','success');setTimeout(()=>document.getElementById('notif').classList.remove('show'),1500);
  }catch(e){showNotif('Erro: '+e.message,'error');}
}

function renderCATopics(){
  const panel=el('ca-topics-panel');
  if(!panel)return;
  // Show to Gestao group (not CA, they see everything already)
  if(!userGroups.includes(GROUPS.Gestao)||userGroups.includes(GROUPS.CA)){panel.innerHTML='';return;}
  const caItems=items.filter(i=>i.reuniao==='Reunião com CAdm'&&i.estado!=='Concluído'&&i.estado!=='Cancelado');
  if(!caItems.length){panel.innerHTML='';return;}
  panel.innerHTML=`<div style="margin-top:1.5rem">
    <div style="font-size:13px;font-weight:600;color:var(--text2);margin-bottom:.75rem;display:flex;align-items:center;gap:8px">
      <span class="badge b-ca">CA</span>
      Tópicos em Conselho de Administração
      <span style="font-size:11px;font-weight:400;color:var(--text3)">(só leitura)</span>
    </div>
    <div style="background:var(--bg2);border:0.5px solid var(--border);border-radius:var(--r-lg);overflow:hidden">
      <table style="width:100%;border-collapse:collapse">
        <thead style="background:var(--bg3)">
          <tr>
            <th style="padding:8px 12px;font-size:11px;font-weight:600;color:var(--text2);text-align:left;border-bottom:0.5px solid var(--border)">Tópico</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:600;color:var(--text2);text-align:left;border-bottom:0.5px solid var(--border)">Tipo</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:600;color:var(--text2);text-align:left;border-bottom:0.5px solid var(--border)">Estado</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:600;color:var(--text2);text-align:left;border-bottom:0.5px solid var(--border)">Data</th>
          </tr>
        </thead>
        <tbody>
          ${caItems.map(i=>`<tr style="border-bottom:0.5px solid var(--border)">
            <td style="padding:8px 12px;font-size:13px;font-weight:500">${i.titulo}</td>
            <td style="padding:8px 12px"><span class="badge ${i.tipo==='Deliberativo'?'b-delib':'b-info'}">${i.tipo==='Deliberativo'?'Delib.':'Inform.'}</span></td>
            <td style="padding:8px 12px"><span class="badge ${{'Em aberto':'b-aberto','Em processo':'b-processo','Concluído':'b-concluido'}[i.estado]||'b-na'}">${i.estado}</span></td>
            <td style="padding:8px 12px;font-size:12px;color:var(--text2)">${fd(i.data)}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>`;
}

const RM={'Reunião com CAdm':'b-ca','Reunião Gestão':'b-gestao','Reunião O.Norte':'b-opnorte','Reunião O.Sul':'b-opsul','Reunião Financeiro':'b-fin','Reunião Impacto':'b-imp'};
const EM={'Em aberto':'b-aberto','Em processo':'b-processo','Concluído':'b-concluido','Cancelado':'b-cancelado'};
const DM={Aprovado:'b-aprov','Não aprovado':'b-nap',Pendente:'b-pend','N/A':'b-na'};
const CM={'Presidente Executivo':'b-pe','Segue para CA':'b-segca','Ambos':'b-ambos','—':'b-na'};
const bdg=(map,v,txt)=>`<span class="badge ${map[v]||'b-na'}">${txt||v||'—'}</span>`;

function renderStats(){
  el('s-total').textContent=items.length;
  el('s-delib').textContent=items.filter(i=>i.tipo==='Deliberativo').length;
  el('s-aprov').textContent=items.filter(i=>i.decisao==='Aprovado').length;
  el('s-pend').textContent=items.filter(i=>i.decisao==='Pendente').length;
}
function filteredItems(){
  const q=(el('q')||{value:''}).value.toLowerCase();
  const fr=gv('ff-r'),ft=gv('ff-t'),fe=gv('ff-e');
  return items.filter(i=>{
    if(!canSeeReuniao(i.reuniao))return false;
    if(q&&!(i.titulo.toLowerCase().includes(q)||(i.responsavel||'').toLowerCase().includes(q)))return false;
    if(fr&&i.reuniao!==fr)return false;
    if(ft&&i.tipo!==ft)return false;
    if(fe&&i.estado!==fe)return false;
    return true;
  });
}
function renderTable(){
  const list=filteredItems(),tb=el('tbody');
  if(!list.length){tb.innerHTML=`<tr><td colspan="9"><div class="empty">Nenhum assunto.<br>Clique em "+ Novo assunto".</div></td></tr>`;el('tcount').textContent='';return;}
  tb.innerHTML=list.map(i=>`<tr>
    <td>${fd(i.data)}</td>
    <td title="${i.titulo}" style="font-weight:500">${i.titulo}${i.confidencial?` <span class="badge b-na" style="font-size:10px">&#128274;</span>`:''}${i.attachments?.length?` <span class="badge" style="background:var(--blue-l);color:var(--blue-d);font-size:10px">&#128206; ${i.attachments.length}</span>`:''}</td>
    <td>${bdg(RM,i.reuniao,i.reuniao.replace('Reunião ',''))}</td>
    <td><span class="badge ${i.tipo==='Deliberativo'?'b-delib':'b-info'}">${i.tipo==='Deliberativo'?'Delib.':'Inform.'}</span></td>
    <td>${i.responsavel||'—'}</td>
    <td>${bdg(EM,i.estado)}</td>
    <td>${bdg(CM,i.encaminhamento||'—')}</td>
    <td>${bdg(DM,i.decisao||'Pendente')}</td>
    <td><div class="actions">
      <button class="btn btn-sm" onclick="editItem(${i.id})">Editar</button>
      <button class="btn btn-sm btn-danger" onclick="deleteItem(${i.id})">&#10005;</button>
    </div></td>
  </tr>`).join('');
  el('tcount').textContent=`${list.length} de ${items.length} assunto${items.length!==1?'s':''}`;
}

function openForm(id){
  editId=id||null;
  const p=el('frm-ov');p.classList.add('open');p.scrollIntoView({behavior:'smooth',block:'nearest'});
  el('existing-attachments').innerHTML='';el('f-file').value='';el('hist-dec-panel').innerHTML='';
  if(id){
    const i=items.find(x=>x.id===id);if(!i)return;
    el('frm-title').textContent='Editar assunto';
    sv2('f-titulo',i.titulo);sv2('f-reuniao',i.reuniao);sv2('f-data',i.data);sv2('f-tipo',i.tipo);
    sv2('f-resp',i.responsavel||'');sv2('f-estado',i.estado);sv2('f-enc',i.encaminhamento||'—');
    sv2('f-dec',i.decisao||'Pendente');sv2('f-nivel-origem',i.nivelOrigem||'Departamento');
    sv2('f-nivel-atual',i.nivelAtual||'Departamento');sv2('f-notas',i.notas||'');sv2('f-proposta',i.proposta||'');
    setTog('conf',!!i.confidencial);setTog('atras',!!i.atrasado);
    if(i.attachments?.length){
      el('existing-attachments').innerHTML=i.attachments.map(a=>`
        <div class="attach-item"><a href="${SP}${a.url}" target="_blank">&#128206; ${a.name}</a>
        <button class="attach-remove" onclick="removeAttach('FJ_Assuntos',${id},'${a.name}')">&#10005;</button></div>`).join('');
    }
    if(i.historicoDecisoes){
      const entries=i.historicoDecisoes.split('\n---\n').filter(Boolean);
      el('hist-dec-panel').innerHTML=`<div class="hist-dec-wrap">
        <div class="hist-dec-title">Histórico de decisões</div>
        ${entries.map(e=>`<div class="hist-dec-item">${e.replace(/\n/g,'<br>')}</div>`).join('')}
      </div>`;
    }
  }else{
    el('frm-title').textContent='Novo assunto';
    ['f-titulo','f-resp','f-notas','f-proposta'].forEach(id=>sv2(id,''));
    sv2('f-data',today());sv2('f-tipo','Deliberativo');sv2('f-estado','Em aberto');sv2('f-enc','—');sv2('f-dec','Pendente');
    sv2('f-nivel-origem','Departamento');sv2('f-nivel-atual','Departamento');
    setTog('conf',false);setTog('atras',false);
  }
}
function editItem(id){openForm(id);}
function closeForm(){el('frm-ov').classList.remove('open');editId=null;}

async function removeAttach(list,itemId,fileName){
  if(!confirm(`Remover "${fileName}"?`))return;
  try{
    await getSPToken();const digest=await getDigest();
    await fetch(`${SP}/_api/web/lists/getbytitle('${list}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`,{method:'DELETE',headers:{Authorization:`Bearer ${spToken}`,'X-RequestDigest':digest,'IF-MATCH':'*'}});
    await loadItems();if(editId)openForm(editId);
  }catch(e){showNotif('Erro: '+e.message,'error');}
}

async function saveItem(){
  const titulo=el('f-titulo').value.trim();
  if(!titulo){alert('Por favor indique o assunto.');return;}
  const btn=el('save-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const body={__metadata:{type:'SP.Data.FJ_x005f_AssuntosListItem'},
      Title:titulo,TipoReuniao:gv('f-reuniao'),TipoAssunto:gv('f-tipo'),
      Notas:gv2t('f-notas'),Proposta:gv2t('f-proposta'),Estado:gv('f-estado'),
      Encaminhamento:gv('f-enc'),Decisao:gv('f-dec'),
      NivelOrigem:gv('f-nivel-origem'),NivelAtual:gv('f-nivel-atual'),
      Confidencial:T.conf,Atrasado:T.atras
    };
    if(gv('f-data'))body.DataReuniao=gv('f-data')+'T00:00:00Z';
    let itemId=editId;
    if(editId){
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Assuntos')/items(${editId})`,body,'MERGE','*');
    }else{
      const r=await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Assuntos')/items`,body,'POST');
      const d=await r.json();itemId=d.d.Id;
    }
    const files=el('f-file').files;
    for(const f of files)await uploadAttachment('FJ_Assuntos',itemId,f);
    closeForm();await loadItems();showNotif('Guardado!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}

async function deleteItem(id){
  if(!confirm('Eliminar este assunto?'))return;
  try{
    await getSPToken();const digest=await getDigest();
    await fetch(`${SP}/_api/web/lists/getbytitle('FJ_Assuntos')/items(${id})`,{method:'DELETE',headers:{Authorization:`Bearer ${spToken}`,'X-RequestDigest':digest,'IF-MATCH':'*'}});
    await loadItems();showNotif('Eliminado.','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
}

// REUNIÕES EXTERNAS
async function loadExtItems(){
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_ReunioesExternas')/items?$select=*,AttachmentFiles&$expand=AttachmentFiles&$orderby=DataReuniao desc&$top=500`);
    extItems=d.d.results.map(r=>({
      id:r.Id,etag:r.__metadata?.etag||'*',
      titulo:r.Title||'',entidade:r.Entidade||'',
      data:r.DataReuniao?r.DataReuniao.split('T')[0]:'',
      local:r.Local||'',participantes:r.Participantes||'',
      resumo:r.Resumo||'',proximosPassos:r.ProximosPassos||'',
      departamento:r.Departamento||'',responsavel:r.Responsavel?.Title||'',
      attachments:(r.AttachmentFiles?.results||[]).map(a=>({name:a.FileName,url:a.ServerRelativeUrl}))
    }));
    renderExtTable();
  }catch(e){console.warn('Ext items error',e);}
}
function renderExtTable(){
  const q=(el('qe')||{value:''}).value.toLowerCase();
  const fd_=gv('ffe-dep'),ff=gv('ffe-from'),ft=gv('ffe-to');
  let list=extItems.filter(i=>{
    if(q&&!(i.titulo.toLowerCase().includes(q)||i.entidade.toLowerCase().includes(q)))return false;
    if(fd_&&i.departamento!==fd_)return false;
    if(ff&&i.data&&i.data<ff)return false;
    if(ft&&i.data&&i.data>ft)return false;
    return true;
  });
  const tb=el('tbody-ext');
  if(!list.length){tb.innerHTML=`<tr><td colspan="7"><div class="empty">Nenhuma reunião externa.<br>Clique em "+ Nova reunião".</div></td></tr>`;el('tcount-ext').textContent='';return;}
  tb.innerHTML=list.map(i=>`<tr>
    <td>${fd(i.data)}</td>
    <td style="font-weight:500" title="${i.titulo}">${i.titulo}${i.attachments?.length?` <span class="badge" style="background:var(--blue-l);color:var(--blue-d);font-size:10px">&#128206; ${i.attachments.length}</span>`:''}</td>
    <td>${i.entidade||'—'}</td>
    <td><span class="badge b-gestao" style="font-size:10px">${i.departamento||'—'}</span></td>
    <td>${i.responsavel||'—'}</td>
    <td title="${i.proximosPassos}" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${i.proximosPassos||'—'}</td>
    <td><div class="actions">
      <button class="btn btn-sm" onclick="editExt(${i.id})">Editar</button>
      <button class="btn btn-sm btn-danger" onclick="deleteExt(${i.id})">&#10005;</button>
    </div></td>
  </tr>`).join('');
  el('tcount-ext').textContent=`${list.length} reunião${list.length!==1?'es':''}`;
}
function openExtForm(id){
  editExtId=id||null;
  feInlineTarefas=[];renderInlineTarefas();
  el('frm-ext-ov').classList.add('open');
  el('fe-attachments').innerHTML='';el('fe-file').value='';
  if(id){
    const i=extItems.find(x=>x.id===id);if(!i)return;
    el('frm-ext-title').textContent='Editar reunião externa';
    sv2('fe-titulo',i.titulo);sv2('fe-entidade',i.entidade);sv2('fe-data',i.data);
    sv2('fe-local',i.local||'');sv2('fe-dep',i.departamento||'Gestão');sv2('fe-resp',i.responsavel||'');
    sv2('fe-part',i.participantes||'');sv2('fe-resumo',i.resumo||'');sv2('fe-passos',i.proximosPassos||'');
    if(i.attachments?.length){
      el('fe-attachments').innerHTML=i.attachments.map(a=>`
        <div class="attach-item"><a href="${SP}${a.url}" target="_blank">&#128206; ${a.name}</a>
        <button class="attach-remove" onclick="removeAttach('FJ_ReunioesExternas',${id},'${a.name}')">&#10005;</button></div>`).join('');
    }
  }else{
    el('frm-ext-title').textContent='Nova reunião externa';
    ['fe-titulo','fe-entidade','fe-local','fe-resp','fe-part','fe-resumo','fe-passos'].forEach(id=>sv2(id,''));
    sv2('fe-data',today());sv2('fe-dep','Gestão');
  }
}
function editExt(id){openExtForm(id);}
function closeExtForm(){el('frm-ext-ov').classList.remove('open');editExtId=null;}
async function saveExt(){
  const titulo=el('fe-titulo').value.trim();
  if(!titulo){alert('Por favor indique o título.');return;}
  const btn=el('save-ext-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const body={__metadata:{type:'SP.Data.FJ_x005f_ReunioesExternasListItem'},
      Title:titulo,Entidade:gv('fe-entidade'),Local:gv('fe-local'),
      Participantes:gv2t('fe-part'),Resumo:gv2t('fe-resumo'),ProximosPassos:gv2t('fe-passos'),
      Departamento:gv('fe-dep')
    };
    if(gv('fe-data'))body.DataReuniao=gv('fe-data')+'T00:00:00Z';
    let itemId=editExtId;
    if(editExtId){
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_ReunioesExternas')/items(${editExtId})`,body,'MERGE','*');
    }else{
      const r=await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_ReunioesExternas')/items`,body,'POST');
      const d=await r.json();itemId=d.d.Id;
    }
    const files=el('fe-file').files;
    for(const f of files)await uploadAttachment('FJ_ReunioesExternas',itemId,f);
    if(feInlineTarefas.length)await saveInlineTarefas(itemId,titulo);
    closeExtForm();await loadExtItems();await loadTarefas();showNotif('Reunião externa guardada!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}
async function deleteExt(id){
  if(!confirm('Eliminar esta reunião?'))return;
  try{
    await getSPToken();const digest=await getDigest();
    await fetch(`${SP}/_api/web/lists/getbytitle('FJ_ReunioesExternas')/items(${id})`,{method:'DELETE',headers:{Authorization:`Bearer ${spToken}`,'X-RequestDigest':digest,'IF-MATCH':'*'}});
    await loadExtItems();showNotif('Eliminado.','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
}

// INLINE TASKS FOR EXT MEETINGS
let feInlineTarefas=[];
function addTarefaInline(){
  const titulo=document.getElementById('fe-tar-titulo').value.trim();
  if(!titulo)return;
  const resp=document.getElementById('fe-tar-resp').value.trim();
  const data=document.getElementById('fe-tar-data').value;
  feInlineTarefas.push({titulo,resp,data});
  document.getElementById('fe-tar-titulo').value='';
  document.getElementById('fe-tar-resp').value='';
  document.getElementById('fe-tar-data').value='';
  renderInlineTarefas();
}
function removeInlineTarefa(idx){
  feInlineTarefas.splice(idx,1);
  renderInlineTarefas();
}
function renderInlineTarefas(){
  const list=document.getElementById('fe-rnTarefas-list');
  if(!feInlineTarefas.length){list.innerHTML='';return;}
  list.innerHTML=feInlineTarefas.map((t,i)=>`
    <div style="display:flex;align-items:center;gap:8px;background:var(--bg3);border-radius:8px;padding:8px 10px;margin-bottom:4px">
      <div style="flex:1;font-size:13px;font-weight:500">${t.titulo}</div>
      <div style="font-size:12px;color:var(--text2)">${t.resp||'—'}</div>
      <div style="font-size:12px;color:var(--text2)">${t.data?fd(t.data):'—'}</div>
      <button class="btn btn-sm btn-danger" onclick="removeInlineTarefa(${i})">&#10005;</button>
    </div>`).join('');
}
async function saveInlineTarefas(reuniaoId,reuniaoTitulo){
  for(const t of feInlineTarefas){
    const body={__metadata:{type:'SP.Data.FJ_x005f_TarefasListItem'},
      Title:t.titulo,Departamento:document.getElementById('fe-dep').value,
      Notas:'Origem: Reunião Externa - '+reuniaoTitulo
    };
    if(t.resp)body.Notas=(body.Notas||'')+' | Responsável: '+t.resp;
    if(t.data)body.DataLimite=t.data+'T00:00:00Z';
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items`,body,'POST');
  }
  feInlineTarefas=[];
  renderInlineTarefas();
}

// TAREFAS
async function loadTarefas(){
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items?$select=*&$orderby=DataLimite asc&$top=500`);
    rnTarefas=d.d.results.map(r=>({
      id:r.Id,etag:r.__metadata?.etag||'*',
      titulo:r.Title||'',notas:r.Notas||'',
      responsavel:r.Responsavel||'',
      data:r.DataLimite?r.DataLimite.split('T')[0]:'',
      estado:r.Estado||'Por fazer',prioridade:r.Prioridade||'Média',
      departamento:r.Departamento||'',projetoId:r.ProjetoId||null,
      acaoId:r.AcaoId||null
    }));
    renderTarefasTable();renderTarefasStats();
  }catch(e){console.warn('Tarefas error',e);}
}
function renderTarefasStats(){
  el('t-total').textContent=rnTarefas.length;
  el('t-pfazer').textContent=rnTarefas.filter(t=>t.estado==='Por fazer').length;
  el('t-emcurso').textContent=rnTarefas.filter(t=>t.estado==='Em curso').length;
  const hoje=today();
  el('t-atrasadas').textContent=rnTarefas.filter(t=>t.data&&t.data<hoje&&t.estado!=='Concluída'&&t.estado!=='Cancelada').length;
}
function renderTarefasTable(){
  const q=(el('qt')||{value:''}).value.toLowerCase();
  const fe=gv('fft-est'),fd_=gv('fft-dep'),fp=gv('fft-prior');
  const hoje=today();
  let list=rnTarefas.filter(i=>{
    if(q&&!i.titulo.toLowerCase().includes(q))return false;
    if(fe&&i.estado!==fe)return false;
    if(fd_&&i.departamento!==fd_)return false;
    if(fp&&i.prioridade!==fp)return false;
    return true;
  });
  const tb=el('tbody-tar');
  if(!list.length){tb.innerHTML=`<tr><td colspan="8"><div class="empty">Nenhuma tarefa.<br>Clique em "+ Nova tarefa".</div></td></tr>`;el('tcount-tar').textContent='';return;}
  const PM={'Alta':'b-alta','Média':'b-media','Baixa':'b-baixa'};
  const EM2={'Por fazer':'b-pfazer','Em curso':'b-emcurso','Concluída':'b-concluido','Cancelada':'b-cancelado'};
  tb.innerHTML=list.map(i=>{
    const atrasada=i.data&&i.data<hoje&&i.estado!=='Concluída'&&i.estado!=='Cancelada';
    return `<tr style="${atrasada?'background:#fff8f8':''}">
      <td style="font-weight:500" title="${i.titulo}">${i.titulo}</td>
      <td>${i.responsavel||'—'}</td>
      <td style="${atrasada?'color:var(--red);font-weight:600':''}">${fd(i.data)}${atrasada?' ⚠️':''}</td>
      <td>${bdg(EM2,i.estado)}</td>
      <td>${bdg(PM,i.prioridade)}</td>
      <td>${i.departamento||'—'}</td>
      <td>${i.projeto||'—'}</td>
      <td><div class="actions">
        <button class="btn btn-sm" onclick="editTar(${i.id})">Editar</button>
        <button class="btn btn-sm btn-danger" onclick="deleteTar(${i.id})">&#10005;</button>
      </div></td>
    </tr>`;
  }).join('');
  el('tcount-tar').textContent=`${list.length} tarefa${list.length!==1?'s':''}`;
}
function openTarForm(id){
  rnEditTarId=id||null;
  el('frm-tar-ov').classList.add('open');
  if(id){
    const t=rnTarefas.find(x=>x.id===id);if(!t)return;
    el('frm-tar-title').textContent='Editar tarefa';
    sv2('ft-titulo',t.titulo);sv2('ft-desc',t.descricao||'');sv2('ft-resp',t.responsavel||'');
    sv2('ft-data',t.data);sv2('ft-estado',t.estado);sv2('ft-prior',t.prioridade);
    sv2('ft-dep',t.departamento||'Gestão');sv2('ft-proj',t.projeto||'');
    sv2('ft-origem',t.origem||'Reunião Externa');sv2('ft-reuniao-orig',t.reuniaoOrigem||'');
  }else{
    el('frm-tar-title').textContent='Nova tarefa';
    ['ft-titulo','ft-desc','ft-resp','ft-proj','ft-reuniao-orig'].forEach(id=>sv2(id,''));
    sv2('ft-data','');sv2('ft-estado','Por fazer');sv2('ft-prior','Média');sv2('ft-dep','Gestão');sv2('ft-origem','Reunião Externa');
  }
}
function editTar(id){openTarForm(id);}
function closeTarForm(){el('frm-tar-ov').classList.remove('open');rnEditTarId=null;}
async function saveTarefa(){
  const titulo=el('ft-titulo').value.trim();
  if(!titulo){alert('Por favor indique o título.');return;}
  const btn=el('save-tar-btn');btn.textContent='A guardar...';btn.disabled=true;
  try{
    const body={__metadata:{type:'SP.Data.FJ_x005f_TarefasListItem'},
      Title:titulo,Notas:gv2t('ft-desc'),Estado:gv('ft-estado'),
      Prioridade:gv('ft-prior'),Departamento:gv('ft-dep')
    };
    const projId=gv('ft-proj');if(projId)body.ProjetoId=parseInt(projId);
    if(gv('ft-data'))body.DataLimite=gv('ft-data')+'T00:00:00Z';
    if(rnEditTarId){
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items(${rnEditTarId})`,body,'MERGE','*');
    }else{
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items`,body,'POST');
    }
    closeTarForm();await loadTarefas();showNotif('Tarefa guardada!','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
  btn.textContent='Guardar';btn.disabled=false;
}
async function deleteTar(id){
  if(!confirm('Eliminar esta tarefa?'))return;
  try{
    await getSPToken();const digest=await getDigest();
    await fetch(`${SP}/_api/web/lists/getbytitle('FJ_Tarefas')/items(${id})`,{method:'DELETE',headers:{Authorization:`Bearer ${spToken}`,'X-RequestDigest':digest,'IF-MATCH':'*'}});
    await loadTarefas();showNotif('Eliminado.','success');
  }catch(e){showNotif('Erro: '+e.message,'error');}
}

// PREP MODE & SAVED SESSIONS
let prepItems=[];
let dragSrc=null;

function prepararReuniao(){
  const tipo=gv('r-tipo'),filtro=gv('r-filtro');
  if(!canSeeReuniao(tipo)){
    el('reuniao-setup').innerHTML=`<div class="access-denied"><h3>Sem acesso</h3><p>Não tens permissão para este tipo de reunião.</p><br><button class="btn btn-primary" onclick="location.reload()">← Regressar</button></div>`;
    return;
  }
  let lista=items.filter(i=>i.reuniao===tipo);
  if(filtro==='abertos')lista=lista.filter(i=>i.estado!=='Concluído'&&i.estado!=='Cancelado');
  if(filtro==='pendentes')lista=lista.filter(i=>i.decisao==='Pendente');
  if(!lista.length){alert('Nenhum assunto encontrado.');return;}
  prepItems=lista.map(i=>({...i}));
  el('reuniao-setup').style.display='none';
  el('prep-wrap').style.display='block';
  el('prep-title').textContent='Preparar — '+tipo;
  renderPrepItems();
}

function renderPrepItems(){
  el('prep-count').textContent=`${prepItems.length} assunto${prepItems.length!==1?'s':''}`;
  el('prep-body').innerHTML=prepItems.map((it,i)=>`
    <div class="prep-item" draggable="true" data-idx="${i}"
      ondragstart="dragStart(event,${i})"
      ondragover="dragOver(event,${i})"
      ondrop="dragDrop(event,${i})"
      ondragend="dragEnd(event)">
      <div class="prep-item-num">${i+1}</div>
      <span class="badge ${it.tipo==='Deliberativo'?'b-delib':'b-info'}" style="font-size:10px">${it.tipo==='Deliberativo'?'D':'I'}</span>
      <div class="prep-item-title">${it.titulo}</div>
      ${it.decisao&&it.decisao!=='Pendente'?`<span class="badge ${it.decisao==='Aprovado'?'b-aprov':'b-nap'}" style="font-size:10px">${it.decisao}</span>`:''}
      <button class="prep-item-remove" onclick="removePrepItem(${i})" title="Remover desta sessão">&#10005;</button>
    </div>`).join('');
}

function removePrepItem(idx){
  prepItems.splice(idx,1);
  renderPrepItems();
}

function dragStart(e,idx){dragSrc=idx;e.currentTarget.classList.add('dragging');}
function dragOver(e,idx){e.preventDefault();document.querySelectorAll('.prep-item').forEach((el,i)=>{el.classList.toggle('drag-over',i===idx&&idx!==dragSrc);});}
function dragDrop(e,idx){
  e.preventDefault();
  if(dragSrc===null||dragSrc===idx)return;
  const moved=prepItems.splice(dragSrc,1)[0];
  prepItems.splice(idx,0,moved);
  dragSrc=null;
  renderPrepItems();
}
function dragEnd(e){e.currentTarget.classList.remove('dragging');document.querySelectorAll('.prep-item').forEach(el=>el.classList.remove('drag-over'));}

function fecharPrep(){
  el('prep-wrap').style.display='none';
  el('reuniao-setup').style.display='block';
}

function iniciarComPrep(){
  if(!prepItems.length){alert('Adiciona pelo menos um assunto.');return;}
  carrItems=prepItems;carrIdx=0;carrDec={};currReuniao=gv('r-tipo');
  el('prep-wrap').style.display='none';
  el('carr-wrap').style.display='block';
  el('sumario-wrap').style.display='none';
  el('carr-nome').textContent=currReuniao;
  const cores=getCores(currReuniao);
  el('carr-wrap').querySelector('.carr-header').style.background=cores.bg;
  renderCarr();renderAgendaBar();
}

function guardarSessao(){
  const tipo=gv('r-tipo');
  const sessao={tipo,items:prepItems.map(i=>i.id),data:today()};
  const key='fj_sessao_'+tipo.replace(/\s+/g,'_');
  localStorage.setItem(key,JSON.stringify(sessao));
  showNotif('Sessão guardada! Estará disponível amanhã.','success');
  fecharPrep();
  loadSavedSessions();
}

function loadSavedSessions(){
  const tipo=gv('r-tipo');
  const key='fj_sessao_'+tipo.replace(/\s+/g,'_');
  const wrap=el('saved-sessions-wrap');
  try{
    const saved=localStorage.getItem(key);
    if(!saved){wrap.innerHTML='';return;}
    const sessao=JSON.parse(saved);
    wrap.innerHTML=`<div style="margin-top:1rem;font-size:12px;color:var(--text2);margin-bottom:6px">Sessão guardada:</div>
      <div class="saved-session-item" onclick="retormarSessao('${tipo}')">
        <div>
          <div style="font-weight:600;font-size:13px">&#128203; ${sessao.tipo}</div>
          <div style="font-size:11px;color:var(--text2)">${sessao.items?.length||0} assuntos · Guardada em ${sessao.data||'—'}</div>
        </div>
        <div style="display:flex;gap:6px">
          <button class="btn btn-primary btn-sm" onclick="event.stopPropagation();retormarSessao('${tipo}')">Retomar</button>
          <button class="btn btn-sm btn-danger" onclick="event.stopPropagation();apagarSessao('${tipo}')">&#10005;</button>
        </div>
      </div>`;
  }catch(e){wrap.innerHTML='';}
}

function retormarSessao(tipo){
  const key='fj_sessao_'+tipo.replace(/\s+/g,'_');
  try{
    const sessao=JSON.parse(localStorage.getItem(key));
    const ids=sessao.items||[];
    prepItems=ids.map(id=>items.find(i=>i.id===id)).filter(Boolean);
    if(!prepItems.length){alert('Os assuntos guardados já não existem.');return;}
    el('reuniao-setup').style.display='none';
    el('prep-wrap').style.display='block';
    el('prep-title').textContent='Retomar — '+tipo;
    renderPrepItems();
  }catch(e){showNotif('Erro ao retomar sessão.','error');}
}

function apagarSessao(tipo){
  const key='fj_sessao_'+tipo.replace(/\s+/g,'_');
  localStorage.removeItem(key);
  loadSavedSessions();
}

function renderAgendaBar(){
  const bar=el('agenda-bar');
  if(!bar||!carrItems.length){return;}
  bar.innerHTML=carrItems.map((it,i)=>`
    <div class="agenda-item-chip ${i===carrIdx?'current':i<carrIdx?'done':'pending'}" title="${it.titulo}">
      ${i+1}. ${it.titulo.length>25?it.titulo.slice(0,25)+'…':it.titulo}
    </div>`).join('<span style="color:var(--border2);font-size:10px">›</span>');
}

// CONDUZIR REUNIÃO
function iniciarReuniao(){
  const tipo=gv('r-tipo'),filtro=gv('r-filtro');
  if(!canSeeReuniao(tipo)){
    el('carr-wrap').style.display='none';
    el('reuniao-setup').innerHTML=`<div class="access-denied"><h3>Sem acesso</h3><p>Não tens permissão para conduzir este tipo de reunião.</p><br><button class="btn btn-primary" onclick="location.reload()" style="margin-top:8px">← Regressar</button></div>`;
    return;
  }
  let lista=items.filter(i=>i.reuniao===tipo);
  if(filtro==='abertos')lista=lista.filter(i=>i.estado!=='Concluído'&&i.estado!=='Cancelado');
  if(filtro==='pendentes')lista=lista.filter(i=>i.decisao==='Pendente');
  if(!lista.length){alert('Nenhum assunto encontrado.');return;}
  carrItems=lista;carrIdx=0;carrDec={};currReuniao=tipo;
  el('reuniao-setup').style.display='none';
  el('carr-wrap').style.display='block';
  el('sumario-wrap').style.display='none';
  el('carr-nome').textContent=tipo;
  const cores=getCores(tipo);
  el('carr-wrap').querySelector('.carr-header').style.background=cores.bg;
  el('prog-bar').style.background=cores.bg==='#003087'?'#FF6B00':cores.accent||cores.bg;
  renderCarr();renderAgendaBar();
}

function renderCarr(){
  const it=carrItems[carrIdx],total=carrItems.length,pct=Math.round(((carrIdx+1)/total)*100);
  el('carr-prog-txt').textContent=`Assunto ${carrIdx+1} de ${total}`;
  el('prog-bar').style.width=pct+'%';
  el('btn-prev').disabled=carrIdx===0;
  el('btn-next').textContent=carrIdx===total-1?'Ver sumário →':'Seguinte →';
  el('carr-content').innerHTML=renderContentHTML(it);
  el('carr-side').innerHTML=renderSideHTML(it);
  el('carr-saved-msg').textContent='';
  renderAgendaBar();
}

function renderContentHTML(it){
  const isD=it.tipo==='Deliberativo';
  const attHtml=it.attachments?.length?`<div style="margin-bottom:1.25rem;display:flex;gap:6px;flex-wrap:wrap">${it.attachments.map(a=>`<a href="${SP}${a.url}" target="_blank" class="badge" style="background:var(--blue-l);color:var(--blue-d)">&#128206; ${a.name}</a>`).join('')}</div>`:'';
  const histEnc=it.encaminhamento&&it.encaminhamento!=='—'?`<div class="hist-enc-box">&#128257; Encaminhado anteriormente: <strong>${it.encaminhamento}</strong>${it.historicoDecisoes?`<br><small>${it.historicoDecisoes.split('\n---\n').pop()}</small>`:''}</div>`:'';
  return `
    <div style="display:flex;gap:8px;margin-bottom:1rem;flex-wrap:wrap;align-items:center">
      <span class="badge ${isD?'b-delib':'b-info'}">${it.tipo}</span>
      ${it.confidencial?'<span class="badge b-na">&#128274; Confidencial</span>':''}
      ${it.atrasado?'<span class="badge b-nap">Atrasado</span>':''}
    </div>
    <div class="assunto-titulo">${it.titulo}</div>
    <div class="assunto-meta">
      ${it.responsavel?`<span>&#128100; ${it.responsavel}</span>`:''}
      ${it.data?`<span>&#128197; ${fd(it.data)}</span>`:''}
      ${bdg(EM,it.estado)}
    </div>
    ${attHtml}
    ${histEnc}
    <div class="sec-label">Contextualização</div>
    <div class="content-box ${it.notas?'':'empty'}">${it.notas?(it.notas.replace(/\n/g,'<br>')):'Sem contextualização.'}</div>
    ${isD?`<div class="sec-label">Proposta de Decisão</div>
    <div class="proposta-box ${it.proposta?'':'empty'}">${it.proposta?(it.proposta.replace(/\n/g,'<br>')):'Sem proposta definida.'}</div>`:''}
  `;
}

function renderSideHTML(it){
  const dec=carrDec[it.id]||{decisao:'',encaminhamento:'—',notas:''};
  return `
    <div class="dec-label">Decisão</div>
    <button class="dec-btn ${dec.decisao==='Aprovado'?'sel-aprov':''}" onclick="selDec('Aprovado',this)">&#10003; Aprovado</button>
    <button class="dec-btn ${dec.decisao==='Não aprovado'?'sel-nap':''}" onclick="selDec('Não aprovado',this)">&#10007; Não aprovado</button>
    <button class="dec-btn ${dec.decisao==='Pendente'?'sel-pend':''}" onclick="selDec('Pendente',this)">&#9646; Pendente</button>
    <hr style="border:none;border-top:0.5px solid var(--border);margin:.25rem 0">
    <div class="dec-label">Encaminhamento</div>
    <div class="enc-grid">
      <button class="enc-btn ${dec.encaminhamento==='—'?'sel':''}" onclick="selEnc('—',this)">—</button>
      <button class="enc-btn ${dec.encaminhamento==='Presidente Executivo'?'sel':''}" onclick="selEnc('Presidente Executivo',this)">Pres. Exec.</button>
      <button class="enc-btn ${dec.encaminhamento==='Segue para CA'?'sel':''}" onclick="selEnc('Segue para CA',this)">Segue CA</button>
      <button class="enc-btn ${dec.encaminhamento==='Ambos'?'sel':''}" onclick="selEnc('Ambos',this)">Ambos</button>
    </div>
    <hr style="border:none;border-top:0.5px solid var(--border);margin:.25rem 0">
    <div class="dec-label">Notas da decisão</div>
    <textarea class="notas-dec" id="notas-dec-input" placeholder="Registe as observações...">${dec.notas}</textarea>
    <button class="guardar-dec-btn" onclick="guardarDec()">&#10003; Guardar decisão</button>
    <div class="dec-saved-msg" id="dec-saved-msg"></div>
  `;
}

function selDec(val,btn){
  const it=carrItems[carrIdx];
  if(!carrDec[it.id])carrDec[it.id]={decisao:'',encaminhamento:'—',notas:''};
  carrDec[it.id].decisao=val;
  document.querySelectorAll('.dec-btn').forEach(b=>b.classList.remove('sel-aprov','sel-nap','sel-pend'));
  btn.classList.add(val==='Aprovado'?'sel-aprov':val==='Não aprovado'?'sel-nap':'sel-pend');
  if(fsOpen){const fsB=document.querySelectorAll('#fs-side .dec-btn');fsB.forEach(b=>{const t=b.textContent.includes('Aprovado')?'Aprovado':b.textContent.includes('Não')? 'Não aprovado':'Pendente';b.classList.toggle('sel-aprov',t==='Aprovado'&&val==='Aprovado');b.classList.toggle('sel-nap',t==='Não aprovado'&&val==='Não aprovado');b.classList.toggle('sel-pend',t==='Pendente'&&val==='Pendente');});}
}
function selEnc(val,btn){
  const it=carrItems[carrIdx];
  if(!carrDec[it.id])carrDec[it.id]={decisao:'',encaminhamento:'—',notas:''};
  carrDec[it.id].encaminhamento=val;
  const container=btn.closest('.enc-grid');
  if(container)container.querySelectorAll('.enc-btn').forEach(b=>b.classList.remove('sel'));
  btn.classList.add('sel');
}
function saveNotas(){
  const it=carrItems[carrIdx];
  const n=el('notas-dec-input')||el('fs-notas-input');
  if(n){if(!carrDec[it.id])carrDec[it.id]={decisao:'Pendente',encaminhamento:'—',notas:''};carrDec[it.id].notas=n.value;}
}

async function guardarDec(){
  saveNotas();
  const it=carrItems[carrIdx],dec=carrDec[it.id];
  if(!dec?.decisao){showNotif('Seleciona uma decisão primeiro.','error');return;}
  try{
    // Build history entry
    const now=new Date();
    const quem=user.name||user.username;
    const histEntry=`[${now.toLocaleDateString('pt-PT')} ${now.toLocaleTimeString('pt-PT',{hour:'2-digit',minute:'2-digit'})}] ${dec.decisao} | Encam: ${dec.encaminhamento} | Por: ${quem}${dec.notas?' | Notas: '+dec.notas:''}`;
    const newHist=(it.historicoDecisoes?it.historicoDecisoes+'\n---\n':'')+histEntry;

    // Auto-update TipoReuniao based on encaminhamento
    let novoTipo=it.reuniao;
    if(dec.encaminhamento==='Segue para CA')novoTipo='Reunião com CAdm';
    else if(dec.encaminhamento==='Presidente Executivo'||dec.encaminhamento==='Ambos')novoTipo='Reunião Gestão';

    const body={__metadata:{type:'SP.Data.FJ_x005f_AssuntosListItem'},
      Decisao:dec.decisao,
      Encaminhamento:dec.encaminhamento||'—',
      HistoricoDecisoes:newHist,
      TipoReuniao:novoTipo
    };
    if(dec.notas){body.Notas=(it.notas?it.notas+'\n\n':'')+`[${now.toLocaleDateString('pt-PT')}] ${dec.notas}`;}
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Assuntos')/items(${it.id})`,body,'MERGE','*');

    // Update local item
    const idx=items.findIndex(x=>x.id===it.id);
    if(idx>=0){items[idx].decisao=dec.decisao;items[idx].encaminhamento=dec.encaminhamento;items[idx].historicoDecisoes=newHist;items[idx].reuniao=novoTipo;}
    const cidx=carrItems.findIndex(x=>x.id===it.id);
    if(cidx>=0){carrItems[cidx].historicoDecisoes=newHist;carrItems[cidx].reuniao=novoTipo;}

    const msg=el('dec-saved-msg');if(msg){msg.textContent='✓ Guardado';setTimeout(()=>msg.textContent='',2500);}
    el('carr-saved-msg').textContent='✓ Decisão guardada';
    showNotif('Decisão guardada!','success');
    setTimeout(()=>document.getElementById('notif').classList.remove('show'),2000);
  }catch(e){showNotif('Erro: '+e.message,'error');}
}

function prevA(){saveNotas();if(carrIdx>0){carrIdx--;renderCarr();if(fsOpen)renderFS();}}
function nextA(){
  saveNotas();
  if(carrIdx<carrItems.length-1){carrIdx++;renderCarr();if(fsOpen)renderFS();}
  else terminarReuniao();
}
function terminarReuniao(){saveNotas();guardarReuniao();mostrarSumario();}

async function guardarReuniao(){
  try{
    const now=new Date();
    const aprovados=carrItems.filter(i=>carrDec[i.id]?.decisao==='Aprovado').length;
    const nAprovados=carrItems.filter(i=>carrDec[i.id]?.decisao==='Não aprovado').length;
    const pendentes=carrItems.filter(i=>!carrDec[i.id]?.decisao||carrDec[i.id]?.decisao==='Pendente').length;
    
    // Build detailed notes with all decisions
    const detalhe=carrItems.map((it,i)=>{
      const dec=carrDec[it.id]||{decisao:'Pendente',encaminhamento:'—',notas:''};
      return `${i+1}. ${it.titulo} → ${dec.decisao||'Pendente'}${dec.encaminhamento&&dec.encaminhamento!=='—'?' | Encam: '+dec.encaminhamento:''}${dec.notas?' | '+dec.notas:''}`;
    }).join('\n');

    const titulo=`${currReuniao} — ${now.toLocaleDateString('pt-PT')}`;
    const body={__metadata:{type:'SP.Data.FJ_x005f_ReunioesListItem'},
      Title:titulo,
      TipoReuniao:currReuniao,
      Data:now.toISOString(),
      Participantes:user.name||user.username,
      Estado:'Concluída',
      NotasGerais:`${carrItems.length} assuntos tratados | ${aprovados} aprovados | ${nAprovados} não aprovados | ${pendentes} pendentes\n\n${detalhe}`
    };
    const r=await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Reunioes')/items`,body,'POST');
    const d=await r.json();
    const reuniaoId=d.d.Id;
    
    // Generate and attach PDF summary as text attachment
    const pdfContent=gerarResumoPDF();
    const blob=new Blob([pdfContent],{type:'text/html'});
    const file=new File([blob],`Resumo_${currReuniao.replace(/ /g,'_')}_${now.toISOString().slice(0,10)}.html`,{type:'text/html'});
    try{await uploadAttachment('FJ_Reunioes',reuniaoId,file);}catch(e){console.warn('Anexo PDF falhou',e);}
    
    loadHistReunioes();
    showNotif('Reunião registada no SharePoint!','success');
  }catch(e){
    console.warn('Erro ao guardar reunião',e);
    showNotif('Aviso: reunião não registada — '+e.message,'error');
  }
}

function gerarResumoPDF(){
  const linhas=carrItems.map((it,i)=>{
    const dec=carrDec[it.id]||{decisao:'—',encaminhamento:'—',notas:''};
    const isD=it.tipo==='Deliberativo';
    const dC=dec.decisao==='Aprovado'?'#639922':dec.decisao==='Não aprovado'?'#C00000':'#888780';
    return `<div style="margin-bottom:16px;padding:14px;border:1px solid #dde2ee;border-radius:8px;page-break-inside:avoid">
      <div style="font-size:15px;font-weight:700;margin-bottom:5px;color:#003087">${i+1}. ${it.titulo}</div>
      <div style="font-size:12px;color:#5a6278;margin-bottom:8px">${it.tipo} · ${it.responsavel||'—'} · ${fd(it.data)}</div>
      ${it.notas?`<div style="background:#eef1f6;padding:8px;border-radius:6px;font-size:12px;margin-bottom:8px">${it.notas.replace(/\n/g,'<br>')}</div>`:''}
      ${isD&&it.proposta?`<div style="border:2px dashed #c8d0e0;padding:8px;border-radius:6px;font-size:12px;margin-bottom:8px"><strong>Proposta:</strong> ${it.proposta}</div>`:''}
      ${isD?`<div style="display:flex;gap:8px;font-size:12px;flex-wrap:wrap">
        <span style="background:${dC};color:#fff;padding:2px 10px;border-radius:12px">${dec.decisao||'—'}</span>
        <span style="background:#eef1f6;color:#5a6278;padding:2px 10px;border-radius:12px">${dec.encaminhamento||'—'}</span>
        ${dec.notas?`<span style="color:#5a6278;font-style:italic">${dec.notas}</span>`:''}
      </div>`:''}
    </div>`;
  }).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Arial,sans-serif;padding:40px;max-width:800px;margin:0 auto;color:#1a1a2e}h1{color:#003087;font-size:20px;margin-bottom:4px}.sub{color:#5a6278;font-size:12px;margin-bottom:20px}</style></head><body><h1>Gestão de Reuniões — Fundação da Juventude</h1><div class="sub">${currReuniao} · ${new Date().toLocaleDateString('pt-PT')} · Conduzida por: ${user.name||user.username}</div>${linhas}</body></html>`;
}

async function loadHistReunioes(){
  try{
    const d=await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Reunioes')/items?$select=*&$orderby=Created desc&$top=30`);
    const reunioes=d.d.results.filter(r=>canSeeReuniao(r.TipoReuniao||''));
    const wrap=el('hist-reunioes-wrap');
    if(!reunioes.length){
      wrap.innerHTML=`<div style="text-align:center;color:var(--text3);font-size:13px;padding:1.5rem">Ainda não há reuniões registadas.</div>`;
      return;
    }
    const RM2={'Reunião com CAdm':'b-ca','Reunião Gestão':'b-gestao','Reunião O.Norte':'b-opnorte','Reunião O.Sul':'b-opsul','Reunião Financeiro':'b-fin','Reunião Impacto':'b-imp'};
    wrap.innerHTML=`<div style="font-size:13px;font-weight:600;color:var(--text2);margin-bottom:.75rem;display:flex;align-items:center;gap:8px">
      <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="2" y="2" width="12" height="12" rx="2"/><line x1="5" y1="6" x2="11" y2="6"/><line x1="5" y1="9" x2="9" y2="9"/></svg>
      Histórico de reuniões
    </div>
    ${reunioes.map(r=>{
      const data=r.Data?r.Data.split('T')[0]:'';
      const [y,m,day]=data?data.split('-'):['','',''];
      const dataFmt=data?`${day}/${m}/${y}`:'—';
      return `<div class="hist-item" style="border-left:3px solid var(--fj-blue)">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div style="font-weight:600;font-size:13px">${r.Title||'Reunião'}</div>
          <span class="badge ${RM2[r.TipoReuniao]||'b-na'}" style="font-size:10px">${(r.TipoReuniao||'').replace('Reunião ','')}</span>
        </div>
        <div style="font-size:11px;color:var(--text2);margin-bottom:4px">📅 ${dataFmt} · 👤 ${r.Participantes||'—'}</div>
        ${r.NotasGerais?`<div style="font-size:12px;color:var(--text2);background:var(--bg3);padding:6px 8px;border-radius:6px;white-space:pre-line">${r.NotasGerais.slice(0,300)}${r.NotasGerais.length>300?'…':''}</div>`:''}
      </div>`;
    }).join('')}`;
  }catch(e){console.warn('Hist reunioes error',e);}
}

// FULLSCREEN
function openFS(){
  fsOpen=true;el('fs-wrap').classList.add('open');renderFS();
  document.addEventListener('keydown',fsKey);
}
function closeFS(){
  fsOpen=false;el('fs-wrap').classList.remove('open');
  document.removeEventListener('keydown',fsKey);
}
function fsKey(e){
  if(e.key==='Escape')closeFS();
  else if(e.key==='ArrowRight'||e.key==='ArrowDown')nextA();
  else if(e.key==='ArrowLeft'||e.key==='ArrowUp')prevA();
}
function renderFS(){
  const it=carrItems[carrIdx],total=carrItems.length,pct=Math.round(((carrIdx+1)/total)*100);
  el('fs-nome').textContent=currReuniao;
  const fsCores=getCores(currReuniao);
  el('fs-wrap').querySelector('.fs-topbar').style.background=fsCores.bg;
  el('fs-prog-txt').textContent=`Assunto ${carrIdx+1} de ${total}`;
  el('fs-prog-fill').style.width=pct+'%';
  el('fs-prev').disabled=carrIdx===0;
  el('fs-next').textContent=carrIdx===total-1?'Ver sumário →':'Seguinte →';
  const isD=it.tipo==='Deliberativo';
  const attHtml=it.attachments?.length?`<div style="margin-bottom:1.75rem;display:flex;gap:8px;flex-wrap:wrap">${it.attachments.map(a=>`<a href="${SP}${a.url}" target="_blank" class="badge" style="background:var(--blue-l);color:var(--blue-d)">&#128206; ${a.name}</a>`).join('')}</div>`:'';
  const histEnc=it.encaminhamento&&it.encaminhamento!=='—'?`<div class="hist-enc-box" style="background:rgba(99,59,6,.1);border-color:rgba(99,59,6,.3);color:var(--amber)">&#128257; Encaminhado: <strong>${it.encaminhamento}</strong></div>`:'';
  el('fs-content').innerHTML=`
    <div style="display:flex;gap:8px;margin-bottom:1.25rem;flex-wrap:wrap">
      <span class="badge ${isD?'b-delib':'b-info'}">${it.tipo}</span>
      ${it.confidencial?'<span class="badge b-na">&#128274; Confidencial</span>':''}
    </div>
    <div class="fs-titulo">${it.titulo}</div>
    <div class="fs-meta">
      ${it.responsavel?`<span>&#128100; ${it.responsavel}</span>`:''}
      ${it.data?`<span>&#128197; ${fd(it.data)}</span>`:''}
      ${bdg(EM,it.estado)}
    </div>
    ${attHtml}${histEnc}
    <div class="fs-sec-label">Contextualização</div>
    <div class="fs-content-box ${it.notas?'':'empty'}">${it.notas?(it.notas.replace(/\n/g,'<br>')):'Sem contextualização.'}</div>
    ${isD?`<div class="fs-sec-label">Proposta de Decisão</div>
    <div class="fs-proposta-box ${it.proposta?'':'empty'}">${it.proposta?(it.proposta.replace(/\n/g,'<br>')):'Sem proposta definida.'}</div>`:''}
  `;
  el('fs-side').innerHTML=renderSideHTML(it).replace('notas-dec-input','fs-notas-input').replace("id='dec-saved-msg'","id='fs-dec-saved-msg'");
}

// SUMÁRIO
function mostrarSumario(){
  el('carr-wrap').style.display='none';
  el('sumario-wrap').style.display='block';
  if(fsOpen)closeFS();
  el('sumario-meta').textContent=`${currReuniao} · ${carrItems.length} assunto${carrItems.length!==1?'s':''} · ${new Date().toLocaleDateString('pt-PT')}`;
  el('sumario-body').innerHTML=carrItems.map(it=>{
    const dec=carrDec[it.id]||{decisao:'—',encaminhamento:'—',notas:''};
    return `<div class="sumario-item">
      <div class="sumario-item-title">${it.titulo}</div>
      <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:6px">
        <span class="badge ${it.tipo==='Deliberativo'?'b-delib':'b-info'}">${it.tipo}</span>
        <span class="badge ${DM[dec.decisao]||'b-na'}">${dec.decisao||'—'}</span>
        <span class="badge ${CM[dec.encaminhamento]||'b-na'}">${dec.encaminhamento||'—'}</span>
      </div>
      ${dec.notas?`<div style="font-size:12px;color:var(--text2);font-style:italic">${dec.notas}</div>`:''}
    </div>`;
  }).join('');
}
function novaReuniao(){
  el('reuniao-setup').style.display='block';
  el('carr-wrap').style.display='none';
  el('sumario-wrap').style.display='none';
  loadItems();
}

// PPT EXPORT
async function exportPPT(){
  if(!carrItems.length){alert('Nenhum assunto.');return;}
  const pptx=new PptxGenJS();pptx.layout='LAYOUT_WIDE';
  const W=13.33,H=7.5,WH='FFFFFF',GR='5F5E5A';
  const pptCores=getCores(currReuniao).ppt;
  const B=pptCores.header,O='FF6B00',BLIGHT=pptCores.light;
  const cp=pptx.addSlide();
  cp.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:BLIGHT}});
  cp.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:1.4,fill:{color:B}});
  cp.addShape(pptx.ShapeType.rect,{x:0,y:H-.6,w:W,h:.6,fill:{color:B}});
  cp.addText('Fundação da Juventude',{x:.5,y:0,w:W-1,h:1.4,fontSize:13,color:'AACCEE',valign:'middle',fontFace:'Calibri'});
  cp.addText(currReuniao,{x:.6,y:1.6,w:W-1.2,h:2,fontSize:44,bold:true,color:B,align:'center',valign:'middle',fontFace:'Calibri'});
  cp.addText(new Date().toLocaleDateString('pt-PT'),{x:.6,y:3.8,w:W-1.2,h:.6,fontSize:20,color:GR,align:'center',fontFace:'Calibri'});
  for(const it of carrItems){
    const sl=pptx.addSlide();const dec=carrDec[it.id]||{decisao:'Pendente',encaminhamento:'—',notas:''};const isD=it.tipo==='Deliberativo';
    sl.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:BLIGHT}});
    sl.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:1.1,fill:{color:B}});
    sl.addShape(pptx.ShapeType.rect,{x:0,y:H-.06,w:W,h:.06,fill:{color:O}});
    sl.addText(it.titulo,{x:.3,y:0,w:9.5,h:1.1,fontSize:19,bold:true,color:WH,valign:'middle',fontFace:'Calibri',wrap:true});
    if(it.responsavel)sl.addText(it.responsavel,{x:9.8,y:.1,w:3.2,h:.42,fontSize:12,color:'AACCEE',align:'right',fontFace:'Calibri'});
    if(it.data)sl.addText(fd(it.data),{x:9.8,y:.58,w:3.2,h:.34,fontSize:11,color:'AACCEE',align:'right',fontFace:'Calibri'});
    sl.addText(isD?'Para decisão':'Informativo',{shape:pptx.ShapeType.roundRect,rectRadius:.1,x:.3,y:.76,w:1.6,h:.26,fill:{color:isD?'534AB7':GR},line:{color:isD?'534AB7':GR,width:0},fontSize:10,color:WH,align:'center',valign:'middle',fontFace:'Calibri'});
    sl.addText('Contextualização:',{x:.3,y:1.18,w:2.5,h:.3,fontSize:11,bold:true,color:B,fontFace:'Calibri'});
    sl.addShape(pptx.ShapeType.roundRect,{x:.3,y:1.52,w:W-.6,h:isD?1.9:3.8,fill:{color:WH},line:{color:'DDE2EE',width:.8},rectRadius:.12});
    if(it.notas)sl.addText(it.notas,{x:.5,y:1.62,w:W-1,h:isD?1.7:3.6,fontSize:13,color:'1a1a2e',valign:'top',fontFace:'Calibri',wrap:true});
    if(isD&&it.proposta){
      sl.addText('Proposta:',{x:.3,y:3.52,w:2.8,h:.3,fontSize:11,bold:true,color:B,fontFace:'Calibri'});
      sl.addShape(pptx.ShapeType.roundRect,{x:.3,y:3.86,w:W-.6,h:1.6,fill:{color:WH},line:{color:'C8D0E0',width:1.2,dashType:'dash'},rectRadius:.12});
      if(it.proposta)sl.addText(it.proposta,{x:.5,y:3.96,w:W-1,h:1.4,fontSize:13,color:'1a1a2e',valign:'top',fontFace:'Calibri',wrap:true});
    }
    if(isD){
      const fy=H-.88;const dC=dec.decisao==='Aprovado'?'639922':dec.decisao==='Não aprovado'?'C00000':'888780';
      sl.addText('Decisão:',{x:.3,y:fy,w:1.4,h:.26,fontSize:10,bold:true,color:GR,fontFace:'Calibri'});
      sl.addText(dec.decisao||'Pendente',{shape:pptx.ShapeType.roundRect,rectRadius:.1,x:.3,y:fy+.3,w:1.4,h:.3,fill:{color:dC},line:{color:dC,width:0},fontSize:11,color:WH,bold:true,align:'center',valign:'middle',fontFace:'Calibri'});
      const eC={'Presidente Executivo':'534AB7','Segue para CA':'185FA5','Ambos':'854F0B','—':'888780'}[dec.encaminhamento||'—'];
      sl.addText('Encaminhamento:',{x:1.9,y:fy,w:2.2,h:.26,fontSize:10,bold:true,color:GR,fontFace:'Calibri'});
      sl.addText(dec.encaminhamento||'—',{shape:pptx.ShapeType.roundRect,rectRadius:.1,x:1.9,y:fy+.3,w:2.2,h:.3,fill:{color:eC},line:{color:eC,width:0},fontSize:11,color:WH,bold:true,align:'center',valign:'middle',fontFace:'Calibri'});
      if(dec.notas){sl.addText('Notas:',{x:4.3,y:fy,w:8.7,h:.26,fontSize:10,bold:true,color:GR,fontFace:'Calibri'});sl.addText(dec.notas,{x:4.3,y:fy+.3,w:8.7,h:.3,fontSize:11,color:'1a1a2e',fontFace:'Calibri',wrap:true});}
    }
  }
  await pptx.writeFile({fileName:`${currReuniao.replace(/\s+/g,'_')}_${today()}.pptx`});
}
function exportPDF(){
  const linhas=carrItems.map((it,i)=>{
    const dec=carrDec[it.id]||{decisao:'—',encaminhamento:'—',notas:''};const isD=it.tipo==='Deliberativo';
    const dC=dec.decisao==='Aprovado'?'#639922':dec.decisao==='Não aprovado'?'#C00000':'#888780';
    return `<div style="margin-bottom:20px;padding:16px;border:1px solid #dde2ee;border-radius:8px;page-break-inside:avoid">
      <div style="font-size:16px;font-weight:700;margin-bottom:6px;color:#003087">${i+1}. ${it.titulo}</div>
      <div style="font-size:12px;color:#5a6278;margin-bottom:10px">${it.tipo} · ${it.responsavel||'—'} · ${fd(it.data)}</div>
      ${it.notas?`<div style="background:#eef1f6;padding:10px;border-radius:6px;font-size:13px;margin-bottom:8px">${it.notas}</div>`:''}
      ${isD&&it.proposta?`<div style="border:2px dashed #c8d0e0;padding:10px;border-radius:6px;font-size:13px;margin-bottom:8px"><strong>Proposta:</strong> ${it.proposta}</div>`:''}
      ${isD?`<div style="display:flex;gap:8px;font-size:12px;flex-wrap:wrap">
        <span style="background:${dC};color:#fff;padding:3px 10px;border-radius:12px">${dec.decisao||'—'}</span>
        <span style="background:#eef1f6;color:#5a6278;padding:3px 10px;border-radius:12px">${dec.encaminhamento||'—'}</span>
        ${dec.notas?`<span style="color:#5a6278;font-style:italic">${dec.notas}</span>`:''}
      </div>`:''}
    </div>`;
  }).join('');
  const w=window.open('','_blank');
  w.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Arial,sans-serif;padding:40px;max-width:800px;margin:0 auto;color:#1a1a2e}h1{color:#003087;font-size:22px;margin-bottom:4px}.sub{color:#5a6278;font-size:13px;margin-bottom:24px}@media print{body{padding:20px}}</style></head><body><h1>Gestão de Reuniões — Fundação da Juventude</h1><div class="sub">${currReuniao} · ${new Date().toLocaleDateString('pt-PT')}</div>${linhas}</body></html>`);
  w.document.close();w.print();
}

// UTILS
function showPage(p){
  document.querySelectorAll('.page').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(e=>e.classList.remove('active'));
  el('page-'+p).classList.add('active');
  el('nav-'+p).classList.add('active');
  if(p==='reuniao')loadSavedSessions();
  if(p==='admin')renderAdminUsers();
}

// ═══ ADMINISTRAÇÃO DE UTILIZADORES ═══
const ROLE_LABELS={'pca':'Presidente CA','pe':'Presidente Executivo','df':'Diretor Financeiro','area_director':'Diretor de Área','worker':'Colaborador'};
let _adminEditEmail=null;

function renderAdminUsers(){
  const search=(el('admin-search')?.value||'').toLowerCase();
  const filterRole=el('admin-filter-role')?.value||'';
  const filterDept=el('admin-filter-dept')?.value||'';
  
  let users=orgUsers.filter(u=>{
    const role=userRoles[u.email]?.role||'worker';
    const dept=userRoles[u.email]?.dept||u.dept||'';
    if(search&&!u.name.toLowerCase().includes(search)&&!u.email.toLowerCase().includes(search))return false;
    if(filterRole&&role!==filterRole)return false;
    if(filterDept&&dept!==filterDept)return false;
    return true;
  });
  
  // Estatísticas
  const stats={total:orgUsers.length,df:0,dir:0,colab:0};
  orgUsers.forEach(u=>{
    const role=userRoles[u.email]?.role||'worker';
    if(role==='df')stats.df++;
    else if(['pca','pe','area_director'].includes(role))stats.dir++;
    else stats.colab++;
  });
  el('admin-total').textContent=stats.total;
  el('admin-df').textContent=stats.df;
  el('admin-dir').textContent=stats.dir;
  el('admin-colab').textContent=stats.colab;
  
  const tbody=el('tbody-admin');
  if(!tbody)return;
  tbody.innerHTML=users.map(u=>{
    const role=userRoles[u.email]?.role||'worker';
    const dept=userRoles[u.email]?.dept||u.dept||'—';
    const roleClass={'pca':'b-r','pe':'b-o','df':'b-b','area_director':'b-y','worker':'b-gray'}[role]||'b-gray';
    return`<tr>
      <td><strong>${u.name}</strong></td>
      <td style="color:#5a6278;font-size:13px">${u.email}</td>
      <td>${dept}</td>
      <td><span class="bdg ${roleClass}">${ROLE_LABELS[role]||role}</span></td>
      <td>
        <button class="btn btn-sm" onclick="openAdminEdit('${u.email}')">Editar</button>
      </td>
    </tr>`;
  }).join('');
  
  el('admin-count').textContent=`A mostrar ${users.length} de ${orgUsers.length} utilizadores`;
}

function openAdminEdit(email){
  _adminEditEmail=email;
  const u=orgUsers.find(x=>x.email===email);
  const r=userRoles[email]||{role:'worker',dept:u?.dept||'',projects:[]};
  
  el('admin-edit-name').value=u?.name||email;
  el('admin-edit-role').value=r.role||'worker';
  el('admin-edit-dept').value=r.dept||'';
  el('admin-edit-projects').value=(r.projects||[]).join(', ');
  
  el('admin-modal-ov').style.display='flex';
}

function closeAdminModal(){
  el('admin-modal-ov').style.display='none';
  _adminEditEmail=null;
}

function saveUserPermissions(){
  if(!_adminEditEmail)return;
  
  const role=el('admin-edit-role').value;
  const dept=el('admin-edit-dept').value;
  const projectsStr=el('admin-edit-projects').value;
  const projects=projectsStr?projectsStr.split(',').map(p=>p.trim()).filter(p=>p):[];
  
  userRoles[_adminEditEmail]={role,dept,projects};
  saveUserRoles();
  
  showNotif('Permissões atualizadas com sucesso','success');
  closeAdminModal();
  renderAdminUsers();
}
function showNotif(msg,type='info'){
  const n=el('notif');n.textContent=msg;
  n.className='notif show'+(type==='success'?' success':type==='error'?' error':'');
}

// ── DATE/WEEK UTILITIES (shared) ──
function isoWeek(date){const d=new Date(date);d.setHours(0,0,0,0);d.setDate(d.getDate()+4-(d.getDay()||7));const y=new Date(d.getFullYear(),0,1);return Math.ceil((((d-y)/86400000)+1)/7);}
function weekLabel(w,y){const jan4=new Date(y,0,4);const day=jan4.getDay()||7;const mon=new Date(jan4);mon.setDate(jan4.getDate()-day+1+(w-1)*7);const sun=new Date(mon);sun.setDate(mon.getDate()+6);const fmt=d=>d.toLocaleDateString('pt-PT',{day:'2-digit',month:'short'});return`Semana ${w} · ${fmt(mon)} — ${sun.toLocaleDateString('pt-PT',{day:'2-digit',month:'short',year:'numeric'})}`;}
function weekRange(w,y){const jan4=new Date(y,0,4);const day=jan4.getDay()||7;const mon=new Date(jan4);mon.setDate(jan4.getDate()-day+1+(w-1)*7);const sun=new Date(mon);sun.setDate(mon.getDate()+6);sun.setHours(23,59,59);return{start:mon,end:sun};}

// ── plShowPage alias ──
function plShowPage(p){
  document.querySelectorAll('#module-pl .page').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('#sidebar-pl .nav-item').forEach(e=>e.classList.remove('active'));
  document.getElementById('pl-page-'+p)?.classList.add('active');
  document.getElementById('pl-nav-'+p)?.classList.add('active');
  if(p==='gantt')renderGantt?.();
}

// ── pendingAprovCount (GF) ──
function pendingAprovCount(){
  return (SOLICITACOES||[]).filter(s=>{
    if(_financeRole==='area_director') return s.estado==='solicitada';
    if(_financeRole==='df') return s.estado==='solicitada'||s.estado==='aprovada_dir';
    if(_financeRole==='pe'||_financeRole==='pca') return s.estado==='aprovada';
    return false;
  }).length;
}


function badgeHTML(c,t){return `<span class="bdg ${c}">${t||'—'}</span>`;}
function depB(d){if(!d)return'';return badgeHTML('b-l',d);}
function estB(e){
  const colors={'Concluída':'b-g','Em Curso':'b-b','Atrasada':'b-r','todo':'b-gray','Média':'b-y','Alta':'b-o','Crítica':'b-r','Baixa':'b-gray'};
  return badgeHTML(colors[e]||'b-gray',e);
}

function tog(k){T[k]=!T[k];const b=el('t-'+k),l=el('l-'+k);if(b)b.className='toggle'+(T[k]?' on':'');if(l)l.textContent=T[k]?'Sim':'Não';}
function setTog(k,v){T[k]=v;const b=el('t-'+k),l=el('l-'+k);if(b)b.className='toggle'+(v?' on':'');if(l)l.textContent=v?'Sim':'Não';}
function el(id){return document.getElementById(id);}
function gv(id){const e=el(id);return e?e.value:'';}
function gv2t(id){const e=el(id);return e?e.value:'';}
function sv(id,v){const e=el(id);if(e&&e.tagName==='DIV')e.textContent=v;}
function sv2(id,v){const e=el(id);if(e)e.value=v;}
function fmt(v){return v==null?'—':new Intl.NumberFormat('pt-PT',{minimumFractionDigits:2,maximumFractionDigits:2}).format(v)+' €';}
function pct(v,t){return t>0?Math.min(Math.round(v/t*100),100):0;}
function today(){return new Date().toISOString().slice(0,10);}
function fd(d){if(!d)return'—';try{const[y,m,day]=d.split('-');return`${day}/${m}/${y}`;}catch(e){return d;}}

// init() call removed to avoid duplication

