// ═══ MÓDULO FINANCEIRO COMPLETO ═══
// Integração SharePoint - Versão Final

// ── STATE ──
let GF = {
  screen: 'home',
  year: new Date().getFullYear(),
  project: null,
  rubrica: null,
  solicitacao: null,
  alteracao: null,
  tab: { orc: 'projetos', sol: 'lista', alt: 'lista', aprov: 'pendentes' },
  origens: [] // Para alterações orçamentais
};

let GF_PROJETOS = [];
let GF_RUBRICAS = [];
let GF_SOLICITACOES = [];
let GF_ALTERACOES = [];
let GF_ALTERACOES_ORIGENS = [];
let GF_FATURAS = [];

// ── UTILS ──
const gfFmt = v => v == null ? '—' : new Intl.NumberFormat('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v) + ' €';
const gfDate = d => d ? new Date(d).toLocaleDateString('pt-PT') : '—';
const gfPct = (v, t) => t > 0 ? Math.min(Math.round(v / t * 100), 100) : 0;
const gfEl = id => document.getElementById(id);
const gfVal = id => gfEl(id)?.value || '';
const gfValN = id => parseFloat(gfVal(id)) || 0;

// Gerar código automático
function gfGenCodigo(prefix) {
  const year = new Date().getFullYear();
  const rand = Math.floor(Math.random() * 900) + 100;
  return `${prefix}-${year}-${rand}`;
}

// Badges de estado
const GF_BADGES = {
  // Solicitações
  'Rascunho': '<span class="bdg" style="background:#E5E7EB;color:#374151">Rascunho</span>',
  'Submetido': '<span class="bdg" style="background:#FEF3C7;color:#92400E">Submetido</span>',
  'Aprovado Dir.': '<span class="bdg" style="background:#DBEAFE;color:#1E40AF">Aprovado Dir.</span>',
  'Aprovado DF': '<span class="bdg" style="background:#D1FAE5;color:#065F46">Aprovado DF</span>',
  'Processado': '<span class="bdg" style="background:#10B981;color:#fff">Processado</span>',
  'Rejeitado': '<span class="bdg" style="background:#FEE2E2;color:#991B1B">Rejeitado</span>',
  // Rubricas
  'Ativo': '<span class="bdg" style="background:#D1FAE5;color:#065F46">Ativo</span>',
  'Suspenso': '<span class="bdg" style="background:#FEF3C7;color:#92400E">Suspenso</span>',
  'Fecho': '<span class="bdg" style="background:#E5E7EB;color:#374151">Fecho</span>',
  // Alterações
  'Pendente': '<span class="bdg" style="background:#FEF3C7;color:#92400E">Pendente</span>',
  'Aprovado': '<span class="bdg" style="background:#D1FAE5;color:#065F46">Aprovado</span>'
};

// ── LOAD DATA FROM SHAREPOINT ──
async function gfLoadAll() {
  try {
    // Projetos (usando sharedProjetos ou carregando direto)
    if (!sharedProjetos?.length) await loadSharedProjetos();
    GF_PROJETOS = sharedProjetos.map(p => ({
      id: p.id,
      nome: p.nome,
      codigo: p.codigo || '',
      dep: p.dep || '',
      estado: p.estado || 'Ativo',
      anoOrcamental: p.anoOrcamental || GF.year,
      isBanco: p.isBancoGlobal || false
    }));

    // Rubricas com expand
    const dr = await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Rubricas')/items?$select=*,ProjetoId/Title,ProjetoId/Id&$expand=ProjetoId&$top=2000`);
    GF_RUBRICAS = (dr.d?.results || []).map(r => ({
      id: r.Id,
      titulo: r.Title || '',
      projetoId: r.ProjetoId?.Id || r.ProjetoIdId || null,
      projetoNome: r.ProjetoId?.Title || '',
      tipo: r.TipoRubrica || 'Despesa',
      despOrc: r.DespOrcamentada || 0,
      despSol: r.DespSolicitada || 0,
      despProc: r.DespProcessada || 0,
      recOrc: r.RecOrcamentada || 0,
      recSol: r.RecSolicitada || 0,
      recProc: r.RecProcessada || 0,
      ano: r.AnoOrcamental || GF.year,
      estado: r.Estado || 'Ativo'
    }));

    // Solicitações com expand
    const ds = await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Solicitacoes')/items?$select=*,ProjetoId/Title,RubricaId/Title,Autor/Title,AprovDir/Title,AprovDF/Title&$expand=ProjetoId,RubricaId,Autor,AprovDir,AprovDF&$orderby=Created desc&$top=500`);
    GF_SOLICITACOES = (ds.d?.results || []).map(r => ({
      id: r.Id,
      codigo: r.Codigo || '',
      titulo: r.Title || '',
      projetoId: r.ProjetoId?.Id || r.ProjetoIdId || null,
      projetoNome: r.ProjetoId?.Title || '',
      rubricaId: r.RubricaId?.Id || r.RubricaIdId || null,
      rubricaNome: r.RubricaId?.Title || '',
      valor: r.Valor || 0,
      estado: r.Estado || 'Rascunho',
      fornecedor: r.Fornecedor || '',
      emailForn: r.EmailFornecedor || '',
      justificacao: r.Justificacao || '',
      autorNome: r.Autor?.Title || '',
      dataSubmissao: r.DataSubmissao || null,
      aprovDirNome: r.AprovDir?.Title || null,
      dataAprovDir: r.DataAprovDir || null,
      aprovDFNome: r.AprovDF?.Title || null,
      dataAprovDF: r.DataAprovDF || null,
      motivoRejeicao: r.MotivoRejeicao || ''
    }));

    // Alterações Orçamentais
    const da = await spRead(`${SP}/_api/web/lists/getbytitle('FJ_AlteracoesOrcamentais')/items?$select=*,DestProjetoId/Title,DestRubricaId/Title,Solicitante/Title,AprovDF/Title,AprovPE/Title,AprovPCA/Title&$expand=DestProjetoId,DestRubricaId,Solicitante,AprovDF,AprovPE,AprovPCA&$orderby=Created desc&$top=200`);
    GF_ALTERACOES = (da.d?.results || []).map(r => ({
      id: r.Id,
      titulo: r.Title || '',
      justificacao: r.Justificacao || '',
      valorTotal: r.ValorTotal || 0,
      estado: r.Estado || 'Pendente',
      nivelAprovacao: r.NivelAprovacao || 'df',
      destProjetoId: r.DestProjetoId?.Id || null,
      destProjetoNome: r.DestProjetoId?.Title || '',
      destRubricaId: r.DestRubricaId?.Id || null,
      destRubricaNome: r.DestRubricaId?.Title || '',
      solicitanteNome: r.Solicitante?.Title || '',
      dataSubmissao: r.DataSubmissao || null,
      aprovDFNome: r.AprovDF?.Title || null,
      aprovPENome: r.AprovPE?.Title || null,
      aprovPCANome: r.AprovPCA?.Title || null,
      motivoRejeicao: r.MotivoRejeicao || ''
    }));

    // Origens das Alterações
    const dao = await spRead(`${SP}/_api/web/lists/getbytitle('FJ_AlteracoesOrigens')/items?$select=*,AlteracaoId/Id,OrigProjetoId/Title,OrigRubricaId/Title&$expand=AlteracaoId,OrigProjetoId,OrigRubricaId&$top=1000`);
    GF_ALTERACOES_ORIGENS = (dao.d?.results || []).map(r => ({
      id: r.Id,
      alteracaoId: r.AlteracaoId?.Id || r.AlteracaoIdId || null,
      origProjetoId: r.OrigProjetoId?.Id || null,
      origProjetoNome: r.OrigProjetoId?.Title || '',
      origRubricaId: r.OrigRubricaId?.Id || null,
      origRubricaNome: r.OrigRubricaId?.Title || '',
      valor: r.Valor || 0
    }));

    // Faturas
    const df = await spRead(`${SP}/_api/web/lists/getbytitle('FJ_Faturas')/items?$select=*,SolicitacaoId/Title&$expand=SolicitacaoId&$orderby=Created desc&$top=500`);
    GF_FATURAS = (df.d?.results || []).map(r => ({
      id: r.Id,
      titulo: r.Title || '',
      solicitacaoId: r.SolicitacaoId?.Id || r.SolicitacaoIdId || null,
      solicitacaoNome: r.SolicitacaoId?.Title || '',
      fornecedor: r.Fornecedor || '',
      dataFatura: r.DataFatura || null,
      valor: r.Valor || 0,
      estado: r.Estado || 'Pendente'
    }));

    console.log('GF Data loaded:', { projetos: GF_PROJETOS.length, rubricas: GF_RUBRICAS.length, solicitacoes: GF_SOLICITACOES.length });
    gfRender();
  } catch (e) {
    console.error('gfLoadAll error:', e);
    showNotif('Erro ao carregar dados financeiros: ' + e.message, 'error');
  }
}

// ── NAVIGATION ──
function gfNav(screen, params = {}) {
  GF.screen = screen;
  Object.assign(GF, params);
  gfRender();
  document.querySelectorAll('#sidebar-gf .nav-item').forEach(el => el.classList.remove('active'));
  const navEl = gfEl('gf-nav-' + screen.split('-')[0]);
  if (navEl) navEl.classList.add('active');
}

// ── INIT ──
function gfInit() {
  gfLoadAll();
}

// ── MAIN RENDER ──
function gfRender() {
  const app = gfEl('gf-app');
  if (!app) return;
  
  let html = '';
  switch (GF.screen) {
    case 'home': html = gfRenderHome(); break;
    case 'orcamento': html = gfRenderOrcamento(); break;
    case 'orcamento-rubricas': html = gfRenderRubricas(); break;
    case 'orcamento-rubrica-form': html = gfRenderRubricaForm(); break;
    case 'solicitacoes': html = gfRenderSolicitacoes(); break;
    case 'solicitacao-form': html = gfRenderSolicitacaoForm(); break;
    case 'solicitacao-view': html = gfRenderSolicitacaoView(); break;
    case 'alteracoes': html = gfRenderAlteracoes(); break;
    case 'alteracao-form': html = gfRenderAlteracaoForm(); break;
    case 'aprovacoes': html = gfRenderAprovacoes(); break;
    case 'faturas': html = gfRenderFaturas(); break;
    default: html = gfRenderHome();
  }
  
  app.innerHTML = `<div style="max-width:1200px;margin:0 auto">${html}</div>`;
}

// ══════════════════════════════════════════════════════════════
// HOME
// ══════════════════════════════════════════════════════════════
function gfRenderHome() {
  const pendAprov = GF_SOLICITACOES.filter(s => 
    (_financeRole === 'area_director' && s.estado === 'Submetido') ||
    (_financeRole === 'df' && ['Submetido', 'Aprovado Dir.'].includes(s.estado))
  ).length;
  
  const totOrc = GF_RUBRICAS.filter(r => r.ano === GF.year).reduce((a, r) => a + (r.despOrc || 0), 0);
  const totProc = GF_RUBRICAS.filter(r => r.ano === GF.year).reduce((a, r) => a + (r.despProc || 0), 0);
  
  return `
    <div style="margin-bottom:28px">
      <div style="font-size:22px;font-weight:600;margin-bottom:4px">Gestão Financeira</div>
      <div style="font-size:14px;color:var(--text3)">Ano orçamental: ${GF.year}</div>
    </div>
    
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px">
      <div class="card" style="padding:16px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:6px">Orçamento Total</div>
        <div style="font-size:20px;font-weight:600">${gfFmt(totOrc)}</div>
      </div>
      <div class="card" style="padding:16px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:6px">Processado</div>
        <div style="font-size:20px;font-weight:600;color:var(--blue)">${gfFmt(totProc)}</div>
      </div>
      <div class="card" style="padding:16px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:6px">Execução</div>
        <div style="font-size:20px;font-weight:600">${gfPct(totProc, totOrc)}%</div>
      </div>
      <div class="card" style="padding:16px;${pendAprov > 0 ? 'border-color:#FCD34D;background:#FFFBEB' : ''}">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:6px">Pendentes</div>
        <div style="font-size:20px;font-weight:600;color:${pendAprov > 0 ? '#D97706' : 'inherit'}">${pendAprov}</div>
      </div>
    </div>
    
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px">
      <div class="card" onclick="gfNav('orcamento')" style="cursor:pointer" onmouseover="this.style.boxShadow='0 4px 12px rgba(0,0,0,.1)'" onmouseout="this.style.boxShadow='none'">
        <div style="width:44px;height:44px;background:#EFF6FF;border-radius:10px;display:flex;align-items:center;justify-content:center;margin-bottom:14px">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563EB" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8M12 17v4"/></svg>
        </div>
        <div style="font-size:16px;font-weight:600;margin-bottom:4px">Orçamento</div>
        <div style="font-size:13px;color:var(--text3)">Projetos e rúbricas orçamentais</div>
      </div>
      
      <div class="card" onclick="gfNav('solicitacoes')" style="cursor:pointer" onmouseover="this.style.boxShadow='0 4px 12px rgba(0,0,0,.1)'" onmouseout="this.style.boxShadow='none'">
        <div style="width:44px;height:44px;background:#FFFBEB;border-radius:10px;display:flex;align-items:center;justify-content:center;margin-bottom:14px">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#D97706" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
        </div>
        <div style="font-size:16px;font-weight:600;margin-bottom:4px">Solicitações</div>
        <div style="font-size:13px;color:var(--text3)">Pedidos de efetivação de despesa</div>
      </div>
      
      <div class="card" onclick="gfNav('alteracoes')" style="cursor:pointer" onmouseover="this.style.boxShadow='0 4px 12px rgba(0,0,0,.1)'" onmouseout="this.style.boxShadow='none'">
        <div style="width:44px;height:44px;background:#F0FDF4;border-radius:10px;display:flex;align-items:center;justify-content:center;margin-bottom:14px">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#16A34A" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4z"/></svg>
        </div>
        <div style="font-size:16px;font-weight:600;margin-bottom:4px">Alterações Orçamentais</div>
        <div style="font-size:13px;color:var(--text3)">Transferências entre rúbricas</div>
      </div>
      
      ${gfCanApprove() ? `
      <div class="card" onclick="gfNav('aprovacoes')" style="cursor:pointer;${pendAprov > 0 ? 'border-color:#FCD34D;background:#FFFBEB' : ''}" onmouseover="this.style.boxShadow='0 4px 12px rgba(0,0,0,.1)'" onmouseout="this.style.boxShadow='none'">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
          <div style="width:44px;height:44px;background:#FEF3C7;border-radius:10px;display:flex;align-items:center;justify-content:center">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#D97706" stroke-width="2"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
          </div>
          ${pendAprov > 0 ? `<span class="bdg" style="background:#FCD34D;color:#92400E">${pendAprov}</span>` : ''}
        </div>
        <div style="font-size:16px;font-weight:600;margin-bottom:4px">Aprovações</div>
        <div style="font-size:13px;color:var(--text3)">Pedidos pendentes de aprovação</div>
      </div>
      ` : ''}
    </div>
  `;
}

function gfCanApprove() { return ['area_director', 'df', 'pe', 'pca'].includes(_financeRole); }

// ══════════════════════════════════════════════════════════════
// ORÇAMENTO - PROJETOS (MASTER)
// ══════════════════════════════════════════════════════════════
function gfRenderOrcamento() {
  const projs = GF_PROJETOS.filter(p => !p.anoOrcamental || p.anoOrcamental === GF.year);
  
  // Calcular totais por projeto
  const projsWithTotals = projs.map(p => {
    const rubs = GF_RUBRICAS.filter(r => r.projetoId === p.id && r.ano === GF.year);
    return {
      ...p,
      despOrc: rubs.reduce((a, r) => a + (r.despOrc || 0), 0),
      despSol: rubs.reduce((a, r) => a + (r.despSol || 0), 0),
      despProc: rubs.reduce((a, r) => a + (r.despProc || 0), 0),
      recOrc: rubs.reduce((a, r) => a + (r.recOrc || 0), 0),
      numRubricas: rubs.length
    };
  });
  
  const totD = projsWithTotals.reduce((a, p) => a + p.despOrc, 0);
  const totP = projsWithTotals.reduce((a, p) => a + p.despProc, 0);
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span><span>Orçamento</span>
    </div>
    
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <div>
        <div style="font-size:20px;font-weight:600">Orçamento ${GF.year}</div>
        <div style="font-size:13px;color:var(--text3)">${projs.length} projetos · ${GF_RUBRICAS.filter(r => r.ano === GF.year).length} rúbricas</div>
      </div>
      <div style="display:flex;gap:8px">
        ${[2024, 2025, 2026, 2027].map(y => `
          <button class="btn btn-sm ${y === GF.year ? 'btn-primary' : ''}" onclick="GF.year=${y};gfRender()">${y}</button>
        `).join('')}
      </div>
    </div>
    
    <div class="card" style="padding:0;overflow:hidden">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:var(--bg)">
            <th style="padding:12px 16px;text-align:left;font-weight:500">Projeto</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Rúbricas</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Orçamentado</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Solicitado</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Processado</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Execução</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Ações</th>
          </tr>
        </thead>
        <tbody>
          ${projsWithTotals.length === 0 ? `<tr><td colspan="7" style="padding:40px;text-align:center;color:var(--text3)">Sem projetos</td></tr>` :
            projsWithTotals.map(p => {
              const pct = gfPct(p.despProc, p.despOrc);
              return `
                <tr style="border-bottom:1px solid var(--border)">
                  <td style="padding:14px 16px">
                    <div style="font-weight:500;color:var(--blue)">${p.isBanco ? '⚙ ' : ''}${p.nome}</div>
                    <div style="font-size:11px;color:var(--text3)">${p.dep || '—'}</div>
                  </td>
                  <td style="padding:14px 16px;text-align:center">${p.numRubricas}</td>
                  <td style="padding:14px 16px;text-align:right;font-variant-numeric:tabular-nums">${gfFmt(p.despOrc)}</td>
                  <td style="padding:14px 16px;text-align:right;font-variant-numeric:tabular-nums;color:#D97706">${gfFmt(p.despSol)}</td>
                  <td style="padding:14px 16px;text-align:right;font-variant-numeric:tabular-nums;color:var(--blue)">${gfFmt(p.despProc)}</td>
                  <td style="padding:14px 16px;text-align:center">
                    <div style="display:flex;align-items:center;gap:8px;justify-content:center">
                      <div style="width:60px;height:8px;background:var(--border);border-radius:4px;overflow:hidden">
                        <div style="width:${pct}%;height:100%;background:${pct > 90 ? '#EF4444' : pct > 70 ? '#F59E0B' : '#3B82F6'}"></div>
                      </div>
                      <span style="font-size:12px;font-weight:500">${pct}%</span>
                    </div>
                  </td>
                  <td style="padding:14px 16px;text-align:center">
                    <button class="btn btn-sm" onclick="GF.project=${p.id};gfNav('orcamento-rubricas')">Ver rúbricas</button>
                  </td>
                </tr>
              `;
            }).join('')
          }
        </tbody>
        <tfoot>
          <tr style="background:var(--bg);font-weight:600">
            <td colspan="2" style="padding:14px 16px">Total</td>
            <td style="padding:14px 16px;text-align:right">${gfFmt(totD)}</td>
            <td style="padding:14px 16px;text-align:right;color:#D97706">${gfFmt(projsWithTotals.reduce((a, p) => a + p.despSol, 0))}</td>
            <td style="padding:14px 16px;text-align:right;color:var(--blue)">${gfFmt(totP)}</td>
            <td style="padding:14px 16px;text-align:center">${gfPct(totP, totD)}%</td>
            <td></td>
          </tr>
        </tfoot>
      </table>
    </div>
  `;
}

// ══════════════════════════════════════════════════════════════
// RUBRICAS (DETAIL)
// ══════════════════════════════════════════════════════════════
function gfRenderRubricas() {
  const proj = GF_PROJETOS.find(p => p.id === GF.project);
  if (!proj) return '<p style="padding:40px;text-align:center">Projeto não encontrado</p>';
  
  const rubs = GF_RUBRICAS.filter(r => r.projetoId === proj.id && r.ano === GF.year);
  const totOrc = rubs.reduce((a, r) => a + (r.despOrc || 0), 0);
  const totSol = rubs.reduce((a, r) => a + (r.despSol || 0), 0);
  const totProc = rubs.reduce((a, r) => a + (r.despProc || 0), 0);
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span>
      <a onclick="gfNav('orcamento')" style="color:var(--blue);cursor:pointer">Orçamento</a>
      <span>›</span><span>${proj.nome}</span>
    </div>
    
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <div>
        <div style="font-size:20px;font-weight:600">${proj.nome}</div>
        <div style="font-size:13px;color:var(--text3)">${proj.dep || '—'} · ${rubs.length} rúbricas</div>
      </div>
      <button class="btn btn-primary" onclick="GF.rubrica=null;gfNav('orcamento-rubrica-form')">+ Nova Rúbrica</button>
    </div>
    
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px">
      <div class="card" style="padding:14px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Orçamentado</div>
        <div style="font-size:18px;font-weight:600">${gfFmt(totOrc)}</div>
      </div>
      <div class="card" style="padding:14px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Solicitado</div>
        <div style="font-size:18px;font-weight:600;color:#D97706">${gfFmt(totSol)}</div>
      </div>
      <div class="card" style="padding:14px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Processado</div>
        <div style="font-size:18px;font-weight:600;color:var(--blue)">${gfFmt(totProc)}</div>
      </div>
      <div class="card" style="padding:14px">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Disponível</div>
        <div style="font-size:18px;font-weight:600;color:#10B981">${gfFmt(totOrc - totSol)}</div>
      </div>
    </div>
    
    <div class="card" style="padding:0;overflow:hidden">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:var(--bg)">
            <th style="padding:12px 16px;text-align:left;font-weight:500">Rúbrica</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Tipo</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Orçamentado</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Solicitado</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Processado</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Disponível</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Estado</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Ações</th>
          </tr>
        </thead>
        <tbody>
          ${rubs.length === 0 ? `<tr><td colspan="8" style="padding:40px;text-align:center;color:var(--text3)">Sem rúbricas definidas</td></tr>` :
            rubs.map(r => {
              const orc = r.tipo === 'Receita' ? r.recOrc : r.despOrc;
              const sol = r.tipo === 'Receita' ? r.recSol : r.despSol;
              const proc = r.tipo === 'Receita' ? r.recProc : r.despProc;
              const disp = orc - sol;
              return `
                <tr style="border-bottom:1px solid var(--border)">
                  <td style="padding:14px 16px;font-weight:500">${r.titulo}</td>
                  <td style="padding:14px 16px;text-align:center">
                    <span class="bdg" style="background:${r.tipo === 'Receita' ? '#FFF7ED' : '#EFF6FF'};color:${r.tipo === 'Receita' ? '#EA580C' : '#2563EB'}">${r.tipo}</span>
                  </td>
                  <td style="padding:14px 16px;text-align:right">${gfFmt(orc)}</td>
                  <td style="padding:14px 16px;text-align:right;color:#D97706">${gfFmt(sol)}</td>
                  <td style="padding:14px 16px;text-align:right;color:var(--blue)">${gfFmt(proc)}</td>
                  <td style="padding:14px 16px;text-align:right;color:${disp < 0 ? '#EF4444' : '#10B981'}">${gfFmt(disp)}</td>
                  <td style="padding:14px 16px;text-align:center">${GF_BADGES[r.estado] || r.estado}</td>
                  <td style="padding:14px 16px;text-align:center">
                    <button class="btn btn-sm" onclick="GF.rubrica=${r.id};gfNav('orcamento-rubrica-form')">Editar</button>
                  </td>
                </tr>
              `;
            }).join('')
          }
        </tbody>
      </table>
    </div>
  `;
}

// ══════════════════════════════════════════════════════════════
// RUBRICA FORM
// ══════════════════════════════════════════════════════════════
function gfRenderRubricaForm() {
  const proj = GF_PROJETOS.find(p => p.id === GF.project);
  const rub = GF.rubrica ? GF_RUBRICAS.find(r => r.id === GF.rubrica) : null;
  const isEdit = !!rub;
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span>
      <a onclick="gfNav('orcamento')" style="color:var(--blue);cursor:pointer">Orçamento</a>
      <span>›</span>
      <a onclick="gfNav('orcamento-rubricas')" style="color:var(--blue);cursor:pointer">${proj?.nome || 'Projeto'}</a>
      <span>›</span><span>${isEdit ? 'Editar' : 'Nova'} Rúbrica</span>
    </div>
    
    <div style="font-size:20px;font-weight:600;margin-bottom:20px">${isEdit ? 'Editar' : 'Nova'} Rúbrica</div>
    
    <div class="card" style="max-width:600px">
      <div style="display:flex;flex-direction:column;gap:16px">
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Título <span style="color:#EF4444">*</span></label>
          <input type="text" id="gf-rub-titulo" value="${rub?.titulo || ''}" placeholder="Ex: Formações 2025" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
        </div>
        
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Ano Orçamental <span style="color:#EF4444">*</span></label>
          <select id="gf-rub-ano" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            ${[2024, 2025, 2026, 2027].map(y => `<option value="${y}" ${(rub?.ano || GF.year) === y ? 'selected' : ''}>${y}</option>`).join('')}
          </select>
        </div>
        
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
          <div>
            <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Despesa Orçamentada (€) <span style="color:#EF4444">*</span></label>
            <input type="number" id="gf-rub-desp-orc" value="${rub?.despOrc || 0}" step="0.01" min="0" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
          </div>
          <div>
            <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Receita Orçamentada (€) <span style="color:#EF4444">*</span></label>
            <input type="number" id="gf-rub-rec-orc" value="${rub?.recOrc || 0}" step="0.01" min="0" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
          </div>
        </div>
        
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Estado</label>
          <select id="gf-rub-estado" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            <option value="Ativo" ${rub?.estado === 'Ativo' ? 'selected' : ''}>Ativo</option>
            <option value="Suspenso" ${rub?.estado === 'Suspenso' ? 'selected' : ''}>Suspenso</option>
            <option value="Fecho" ${rub?.estado === 'Fecho' ? 'selected' : ''}>Fecho</option>
          </select>
        </div>
      </div>
      
      <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:24px;padding-top:16px;border-top:1px solid var(--border)">
        <button class="btn" onclick="gfNav('orcamento-rubricas')">Cancelar</button>
        <button class="btn btn-primary" onclick="gfSaveRubrica(${isEdit ? rub.id : 'null'})">${isEdit ? 'Guardar' : 'Criar'}</button>
      </div>
    </div>
  `;
}

async function gfSaveRubrica(id) {
  const titulo = gfVal('gf-rub-titulo').trim();
  const ano = parseInt(gfVal('gf-rub-ano'));
  const estado = gfVal('gf-rub-estado');
  
  if (!titulo) { showNotif('Título obrigatório', 'error'); return; }
  if (!GF.project) { showNotif('Projeto não definido', 'error'); return; }
  
  const body = {
    __metadata: { type: 'SP.Data.FJ_x005f_RubricasListItem' },
    Title: titulo,
    ProjetoIdId: GF.project,
    AnoOrcamental: ano,
    Estado: estado,
    DespOrcamentada: gfValN('gf-rub-desp-orc'),
    RecOrcamentada: gfValN('gf-rub-rec-orc')
  };
  
  try {
    if (id) {
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Rubricas')/items(${id})`, body, 'MERGE');
      showNotif('Rúbrica atualizada', 'success');
    } else {
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Rubricas')/items`, body, 'POST');
      showNotif('Rúbrica criada', 'success');
    }
    await gfLoadAll();
    gfNav('orcamento-rubricas');
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════════
// SOLICITAÇÕES
// ══════════════════════════════════════════════════════════════
function gfRenderSolicitacoes() {
  const tab = GF.tab.sol;
  const lista = tab === 'minhas' ? GF_SOLICITACOES.filter(s => s.autorNome === user?.name) : GF_SOLICITACOES;
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span><span>Solicitações</span>
    </div>
    
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <div style="font-size:20px;font-weight:600">Solicitações de Despesa</div>
      <button class="btn btn-primary" onclick="GF.solicitacao=null;gfNav('solicitacao-form')">+ Nova Solicitação</button>
    </div>
    
    <div style="display:flex;gap:4px;margin-bottom:16px;border-bottom:1px solid var(--border)">
      <button onclick="GF.tab.sol='lista';gfRender()" style="padding:10px 16px;border:none;background:none;cursor:pointer;font-size:13px;font-weight:500;color:${tab === 'lista' ? 'var(--blue)' : 'var(--text3)'};border-bottom:2px solid ${tab === 'lista' ? 'var(--blue)' : 'transparent'}">Todas</button>
      <button onclick="GF.tab.sol='minhas';gfRender()" style="padding:10px 16px;border:none;background:none;cursor:pointer;font-size:13px;font-weight:500;color:${tab === 'minhas' ? 'var(--blue)' : 'var(--text3)'};border-bottom:2px solid ${tab === 'minhas' ? 'var(--blue)' : 'transparent'}">Minhas</button>
    </div>
    
    <div class="card" style="padding:0;overflow:hidden">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:var(--bg)">
            <th style="padding:12px 16px;text-align:left;font-weight:500">Código</th>
            <th style="padding:12px 16px;text-align:left;font-weight:500">Título</th>
            <th style="padding:12px 16px;text-align:left;font-weight:500">Projeto</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Valor</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Estado</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Data</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Ações</th>
          </tr>
        </thead>
        <tbody>
          ${lista.length === 0 ? `<tr><td colspan="7" style="padding:40px;text-align:center;color:var(--text3)">Sem solicitações</td></tr>` :
            lista.map(s => `
              <tr style="border-bottom:1px solid var(--border)">
                <td style="padding:14px 16px;font-family:monospace;font-size:12px">${s.codigo || '—'}</td>
                <td style="padding:14px 16px;font-weight:500">${s.titulo}</td>
                <td style="padding:14px 16px;font-size:12px;color:var(--text3)">${s.projetoNome || '—'}</td>
                <td style="padding:14px 16px;text-align:right;font-weight:500">${gfFmt(s.valor)}</td>
                <td style="padding:14px 16px;text-align:center">${GF_BADGES[s.estado] || s.estado}</td>
                <td style="padding:14px 16px;text-align:center;font-size:12px">${gfDate(s.dataSubmissao)}</td>
                <td style="padding:14px 16px;text-align:center">
                  <button class="btn btn-sm" onclick="GF.solicitacao=${s.id};gfNav('solicitacao-view')">Ver</button>
                </td>
              </tr>
            `).join('')
          }
        </tbody>
      </table>
    </div>
  `;
}

// ══════════════════════════════════════════════════════════════
// SOLICITAÇÃO FORM
// ══════════════════════════════════════════════════════════════
function gfRenderSolicitacaoForm() {
  const sol = GF.solicitacao ? GF_SOLICITACOES.find(s => s.id === GF.solicitacao) : null;
  const isEdit = !!sol;
  const codigo = sol?.codigo || gfGenCodigo('SOL');
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span>
      <a onclick="gfNav('solicitacoes')" style="color:var(--blue);cursor:pointer">Solicitações</a>
      <span>›</span><span>${isEdit ? 'Editar' : 'Nova'}</span>
    </div>
    
    <div style="font-size:20px;font-weight:600;margin-bottom:20px">${isEdit ? 'Editar' : 'Nova'} Solicitação</div>
    
    <div style="display:grid;grid-template-columns:2fr 1fr;gap:20px">
      <div>
        <div class="card" style="margin-bottom:16px">
          <div style="font-size:12px;font-weight:600;color:var(--text3);text-transform:uppercase;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border)">Identificação</div>
          
          <div style="display:grid;grid-template-columns:150px 1fr;gap:16px;margin-bottom:16px">
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Código</label>
              <input type="text" id="gf-sol-codigo" value="${codigo}" readonly style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px;background:#f5f5f5;font-family:monospace">
            </div>
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Título <span style="color:#EF4444">*</span></label>
              <input type="text" id="gf-sol-titulo" value="${sol?.titulo || ''}" placeholder="Ex: Reserva alojamento — FórUM" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            </div>
          </div>
          
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Projeto <span style="color:#EF4444">*</span></label>
              <select id="gf-sol-projeto" onchange="gfSolProjetoChange()" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
                <option value="">Selecionar...</option>
                ${GF_PROJETOS.map(p => `<option value="${p.id}" ${sol?.projetoId === p.id ? 'selected' : ''}>${p.nome}</option>`).join('')}
              </select>
            </div>
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Rúbrica <span style="color:#EF4444">*</span></label>
              <select id="gf-sol-rubrica" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
                <option value="">Selecionar projeto primeiro...</option>
              </select>
            </div>
          </div>
        </div>
        
        <div class="card" style="margin-bottom:16px">
          <div style="font-size:12px;font-weight:600;color:var(--text3);text-transform:uppercase;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border)">Valor e Fornecedor</div>
          
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px">
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Valor (€) <span style="color:#EF4444">*</span></label>
              <input type="number" id="gf-sol-valor" value="${sol?.valor || ''}" min="0" step="0.01" placeholder="0,00" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            </div>
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Fornecedor <span style="color:#EF4444">*</span></label>
              <input type="text" id="gf-sol-fornecedor" value="${sol?.fornecedor || ''}" placeholder="Nome da empresa" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            </div>
            <div>
              <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Email Fornecedor</label>
              <input type="email" id="gf-sol-email" value="${sol?.emailForn || ''}" placeholder="email@empresa.pt" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            </div>
          </div>
          
          <div>
            <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Justificação</label>
            <textarea id="gf-sol-justificacao" placeholder="Descreva o motivo do pedido..." style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px;min-height:100px;resize:vertical">${sol?.justificacao || ''}</textarea>
          </div>
        </div>
        
        <div style="display:flex;gap:10px;justify-content:flex-end">
          <button class="btn" onclick="gfNav('solicitacoes')">Cancelar</button>
          <button class="btn" onclick="gfSaveSolicitacao(${isEdit ? sol.id : 'null'}, 'Rascunho')">Guardar Rascunho</button>
          <button class="btn btn-primary" onclick="gfSaveSolicitacao(${isEdit ? sol.id : 'null'}, 'Submetido')">Submeter</button>
        </div>
      </div>
      
      <div>
        <div class="card" style="background:#F8FAFC">
          <div style="font-size:12px;font-weight:600;color:var(--text3);text-transform:uppercase;margin-bottom:10px">Resumo da Rúbrica</div>
          <div id="gf-sol-rub-info" style="font-size:13px;color:var(--text3)">Selecione projeto e rúbrica</div>
        </div>
      </div>
    </div>
  `;
}

function gfSolProjetoChange() {
  const projId = parseInt(gfVal('gf-sol-projeto'));
  const rubSel = gfEl('gf-sol-rubrica');
  const sol = GF.solicitacao ? GF_SOLICITACOES.find(s => s.id === GF.solicitacao) : null;
  
  if (!projId) {
    rubSel.innerHTML = '<option value="">Selecionar projeto primeiro...</option>';
    return;
  }
  
  const rubs = GF_RUBRICAS.filter(r => r.projetoId === projId && r.estado === 'Ativo');
  rubSel.innerHTML = '<option value="">Selecionar rúbrica...</option>' + 
    rubs.map(r => {
      const disp = (r.despOrc || 0) - (r.despSol || 0);
      return `<option value="${r.id}" ${sol?.rubricaId === r.id ? 'selected' : ''}>${r.titulo} (Disp: ${gfFmt(disp)})</option>`;
    }).join('');
  
  // Trigger rubrica change to show info
  if (sol?.rubricaId) gfSolRubricaInfo(sol.rubricaId);
  rubSel.onchange = () => gfSolRubricaInfo(parseInt(rubSel.value));
}

function gfSolRubricaInfo(rubId) {
  const rub = GF_RUBRICAS.find(r => r.id === rubId);
  const info = gfEl('gf-sol-rub-info');
  if (!rub || !info) return;
  
  const disp = (rub.despOrc || 0) - (rub.despSol || 0);
  info.innerHTML = `
    <div style="margin-bottom:8px"><strong>${rub.titulo}</strong></div>
    <div>Orçamentado: ${gfFmt(rub.despOrc)}</div>
    <div>Solicitado: ${gfFmt(rub.despSol)}</div>
    <div style="font-weight:600;color:${disp < 0 ? '#EF4444' : '#10B981'}">Disponível: ${gfFmt(disp)}</div>
  `;
}

async function gfSaveSolicitacao(id, estado) {
  const codigo = gfVal('gf-sol-codigo');
  const titulo = gfVal('gf-sol-titulo').trim();
  const projetoId = parseInt(gfVal('gf-sol-projeto'));
  const rubricaId = parseInt(gfVal('gf-sol-rubrica'));
  const valor = gfValN('gf-sol-valor');
  const fornecedor = gfVal('gf-sol-fornecedor').trim();
  const emailForn = gfVal('gf-sol-email').trim();
  const justificacao = gfVal('gf-sol-justificacao').trim();
  
  if (!titulo || !projetoId || !rubricaId || valor <= 0 || !fornecedor) {
    showNotif('Preencha todos os campos obrigatórios', 'error');
    return;
  }
  
  // Validar disponível
  const rub = GF_RUBRICAS.find(r => r.id === rubricaId);
  const disp = (rub?.despOrc || 0) - (rub?.despSol || 0);
  if (valor > disp && estado === 'Submetido') {
    showNotif('Valor excede o disponível na rúbrica', 'error');
    return;
  }
  
  const body = {
    __metadata: { type: 'SP.Data.FJ_x005f_SolicitacoesListItem' },
    Codigo: codigo,
    Title: titulo,
    ProjetoIdId: projetoId,
    RubricaIdId: rubricaId,
    Valor: valor,
    Estado: estado,
    Fornecedor: fornecedor,
    EmailFornecedor: emailForn,
    Justificacao: justificacao
  };
  
  if (estado === 'Submetido') {
    body.DataSubmissao = new Date().toISOString();
  }
  
  try {
    if (id) {
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Solicitacoes')/items(${id})`, body, 'MERGE');
      showNotif('Solicitação atualizada', 'success');
    } else {
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Solicitacoes')/items`, body, 'POST');
      showNotif('Solicitação ' + (estado === 'Submetido' ? 'submetida' : 'guardada'), 'success');
    }
    await gfLoadAll();
    gfNav('solicitacoes');
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════════
// SOLICITAÇÃO VIEW
// ══════════════════════════════════════════════════════════════
function gfRenderSolicitacaoView() {
  const sol = GF_SOLICITACOES.find(s => s.id === GF.solicitacao);
  if (!sol) return '<p style="padding:40px;text-align:center">Solicitação não encontrada</p>';
  
  const canApproveDir = _financeRole === 'area_director' && sol.estado === 'Submetido';
  const canApproveDF = _financeRole === 'df' && ['Submetido', 'Aprovado Dir.'].includes(sol.estado);
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span>
      <a onclick="gfNav('solicitacoes')" style="color:var(--blue);cursor:pointer">Solicitações</a>
      <span>›</span><span>${sol.codigo}</span>
    </div>
    
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <div>
        <div style="font-size:20px;font-weight:600">${sol.titulo}</div>
        <div style="display:flex;align-items:center;gap:8px;margin-top:4px">
          <span style="font-family:monospace;font-size:13px;color:var(--text3)">${sol.codigo}</span>
          ${GF_BADGES[sol.estado] || sol.estado}
        </div>
      </div>
      <div style="font-size:24px;font-weight:600;color:var(--blue)">${gfFmt(sol.valor)}</div>
    </div>
    
    <div style="display:grid;grid-template-columns:2fr 1fr;gap:20px">
      <div>
        <div class="card" style="margin-bottom:16px">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
            <div>
              <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Projeto</div>
              <div style="font-weight:500">${sol.projetoNome || '—'}</div>
            </div>
            <div>
              <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Rúbrica</div>
              <div style="font-weight:500">${sol.rubricaNome || '—'}</div>
            </div>
            <div>
              <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Fornecedor</div>
              <div style="font-weight:500">${sol.fornecedor || '—'}</div>
            </div>
            <div>
              <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Email</div>
              <div style="font-weight:500">${sol.emailForn || '—'}</div>
            </div>
          </div>
          ${sol.justificacao ? `
            <div style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border)">
              <div style="font-size:11px;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Justificação</div>
              <div style="font-size:13px">${sol.justificacao}</div>
            </div>
          ` : ''}
        </div>
        
        ${(canApproveDir || canApproveDF) ? `
        <div class="card" style="background:#FFFBEB;border-color:#FCD34D">
          <div style="font-size:14px;font-weight:600;margin-bottom:12px">Ações de Aprovação</div>
          <div style="display:flex;gap:10px">
            <button class="btn" style="background:#10B981;color:white" onclick="gfAprovarSolicitacao(${sol.id})">✓ Aprovar</button>
            <button class="btn" style="background:#EF4444;color:white" onclick="gfRejeitarSolicitacao(${sol.id})">✕ Rejeitar</button>
          </div>
        </div>
        ` : ''}
        
        ${sol.motivoRejeicao ? `
        <div class="card" style="background:#FEF2F2;border-color:#FECACA;margin-top:16px">
          <div style="font-size:12px;font-weight:600;color:#991B1B;margin-bottom:4px">Motivo da Rejeição</div>
          <div style="font-size:13px">${sol.motivoRejeicao}</div>
        </div>
        ` : ''}
      </div>
      
      <div>
        <div class="card">
          <div style="font-size:12px;font-weight:600;color:var(--text3);text-transform:uppercase;margin-bottom:14px">Fluxo de Aprovação</div>
          <div style="display:flex;flex-direction:column;gap:12px">
            ${gfWorkflowStep('Submissão', true, sol.autorNome, sol.dataSubmissao)}
            ${gfWorkflowStep('Aprovação Dir. Área', !!sol.aprovDirNome || sol.estado === 'Aprovado DF' || sol.estado === 'Processado', sol.aprovDirNome, sol.dataAprovDir)}
            ${gfWorkflowStep('Aprovação DF', sol.estado === 'Aprovado DF' || sol.estado === 'Processado', sol.aprovDFNome, sol.dataAprovDF)}
            ${gfWorkflowStep('Processado', sol.estado === 'Processado', null, null)}
          </div>
        </div>
      </div>
    </div>
  `;
}

function gfWorkflowStep(label, done, who, when) {
  return `
    <div style="display:flex;align-items:flex-start;gap:10px">
      <div style="width:20px;height:20px;border-radius:50%;background:${done ? '#10B981' : '#E5E7EB'};display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px">
        ${done ? '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>' : ''}
      </div>
      <div>
        <div style="font-size:13px;font-weight:500;color:${done ? 'var(--text)' : 'var(--text3)'}">${label}</div>
        ${who || when ? `<div style="font-size:11px;color:var(--text3)">${who || ''} ${when ? '· ' + gfDate(when) : ''}</div>` : ''}
      </div>
    </div>
  `;
}

async function gfAprovarSolicitacao(id) {
  if (!confirm('Aprovar esta solicitação?')) return;
  try {
    const novoEstado = _financeRole === 'area_director' ? 'Aprovado Dir.' : 'Aprovado DF';
    const body = {
      __metadata: { type: 'SP.Data.FJ_x005f_SolicitacoesListItem' },
      Estado: novoEstado
    };
    if (_financeRole === 'area_director') {
      body.DataAprovDir = new Date().toISOString();
    } else {
      body.DataAprovDF = new Date().toISOString();
    }
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Solicitacoes')/items(${id})`, body, 'MERGE');
    showNotif('Solicitação aprovada!', 'success');
    await gfLoadAll();
    gfNav('solicitacoes');
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

async function gfRejeitarSolicitacao(id) {
  const motivo = prompt('Motivo da rejeição (obrigatório):');
  if (!motivo || !motivo.trim()) { showNotif('Motivo obrigatório', 'error'); return; }
  try {
    const body = {
      __metadata: { type: 'SP.Data.FJ_x005f_SolicitacoesListItem' },
      Estado: 'Rejeitado',
      MotivoRejeicao: motivo.trim()
    };
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_Solicitacoes')/items(${id})`, body, 'MERGE');
    showNotif('Solicitação rejeitada', 'success');
    await gfLoadAll();
    gfNav('solicitacoes');
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════════
// ALTERAÇÕES ORÇAMENTAIS
// ══════════════════════════════════════════════════════════════
function gfRenderAlteracoes() {
  const tab = GF.tab.alt;
  const pendentes = GF_ALTERACOES.filter(a => a.estado === 'Pendente');
  const historico = GF_ALTERACOES.filter(a => a.estado !== 'Pendente');
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span><span>Alterações Orçamentais</span>
    </div>
    
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <div style="font-size:20px;font-weight:600">Alterações Orçamentais</div>
      <button class="btn btn-primary" onclick="GF.alteracao=null;GF.origens=[];gfNav('alteracao-form')">+ Nova Alteração</button>
    </div>
    
    <div style="display:flex;gap:4px;margin-bottom:16px;border-bottom:1px solid var(--border)">
      <button onclick="GF.tab.alt='lista';gfRender()" style="padding:10px 16px;border:none;background:none;cursor:pointer;font-size:13px;font-weight:500;color:${tab === 'lista' ? 'var(--blue)' : 'var(--text3)'};border-bottom:2px solid ${tab === 'lista' ? 'var(--blue)' : 'transparent'}">Todas (${GF_ALTERACOES.length})</button>
      <button onclick="GF.tab.alt='pendentes';gfRender()" style="padding:10px 16px;border:none;background:none;cursor:pointer;font-size:13px;font-weight:500;color:${tab === 'pendentes' ? 'var(--blue)' : 'var(--text3)'};border-bottom:2px solid ${tab === 'pendentes' ? 'var(--blue)' : 'transparent'}">Pendentes (${pendentes.length})</button>
    </div>
    
    ${gfRenderAltList(tab === 'pendentes' ? pendentes : GF_ALTERACOES)}
  `;
}

function gfRenderAltList(list) {
  if (list.length === 0) {
    return `<div class="card" style="padding:40px;text-align:center;color:var(--text3)">Sem alterações</div>`;
  }
  
  return `
    <div class="card" style="padding:0;overflow:hidden">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:var(--bg)">
            <th style="padding:12px 16px;text-align:left;font-weight:500">Descrição</th>
            <th style="padding:12px 16px;text-align:left;font-weight:500">Destino</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Valor</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Nível</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Estado</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Data</th>
          </tr>
        </thead>
        <tbody>
          ${list.map(a => `
            <tr style="border-bottom:1px solid var(--border)">
              <td style="padding:14px 16px">
                <div style="font-weight:500">${a.titulo || 'Alteração'}</div>
                <div style="font-size:11px;color:var(--text3)">${a.solicitanteNome}</div>
              </td>
              <td style="padding:14px 16px;font-size:12px">${a.destProjetoNome} → ${a.destRubricaNome}</td>
              <td style="padding:14px 16px;text-align:right;font-weight:500">${gfFmt(a.valorTotal)}</td>
              <td style="padding:14px 16px;text-align:center">
                <span class="bdg" style="background:#EFF6FF;color:#2563EB;text-transform:uppercase;font-size:10px">${a.nivelAprovacao}</span>
              </td>
              <td style="padding:14px 16px;text-align:center">${GF_BADGES[a.estado] || a.estado}</td>
              <td style="padding:14px 16px;text-align:center;font-size:12px">${gfDate(a.dataSubmissao)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}

// ══════════════════════════════════════════════════════════════
// ALTERAÇÃO FORM (MASTER-DETAIL)
// ══════════════════════════════════════════════════════════════
function gfRenderAlteracaoForm() {
  const totalOrigens = GF.origens.reduce((a, o) => a + (o.valor || 0), 0);
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span>
      <a onclick="gfNav('alteracoes')" style="color:var(--blue);cursor:pointer">Alterações</a>
      <span>›</span><span>Nova Alteração</span>
    </div>
    
    <div style="font-size:20px;font-weight:600;margin-bottom:20px">Nova Alteração Orçamental</div>
    
    <div class="card" style="margin-bottom:16px">
      <div style="font-size:12px;font-weight:600;color:var(--text3);text-transform:uppercase;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border)">Destino da Transferência</div>
      
      <div style="display:grid;grid-template-columns:1fr 1fr 150px;gap:16px;margin-bottom:16px">
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Projeto Destino <span style="color:#EF4444">*</span></label>
          <select id="gf-alt-dest-proj" onchange="gfAltDestProjChange()" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            <option value="">Selecionar...</option>
            ${GF_PROJETOS.map(p => `<option value="${p.id}">${p.nome}</option>`).join('')}
          </select>
        </div>
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Rúbrica Destino <span style="color:#EF4444">*</span></label>
          <select id="gf-alt-dest-rub" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            <option value="">Selecionar projeto primeiro...</option>
          </select>
        </div>
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Valor Total (€)</label>
          <input type="number" id="gf-alt-valor" value="${totalOrigens || ''}" step="0.01" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px;background:#f5f5f5" readonly>
        </div>
      </div>
      
      <div style="display:grid;grid-template-columns:1fr 200px;gap:16px">
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Justificação <span style="color:#EF4444">*</span></label>
          <textarea id="gf-alt-justificacao" placeholder="Explique o motivo da alteração..." style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px;min-height:60px;resize:vertical"></textarea>
        </div>
        <div>
          <label style="display:block;font-size:13px;font-weight:500;margin-bottom:6px">Nível de Aprovação</label>
          <select id="gf-alt-nivel" style="width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:6px;font-size:14px">
            <option value="df">Apenas DF</option>
            <option value="df_pe">DF + PE</option>
            <option value="df_pe_pca">DF + PE + PCA</option>
          </select>
        </div>
      </div>
    </div>
    
    <div class="card" style="margin-bottom:16px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border)">
        <div style="font-size:12px;font-weight:600;color:var(--text3);text-transform:uppercase">Origens da Transferência</div>
        <button class="btn btn-sm" onclick="gfAddOrigem()">+ Adicionar Origem</button>
      </div>
      
      <div id="gf-alt-origens">
        ${GF.origens.length === 0 ? 
          `<div style="padding:20px;text-align:center;color:var(--text3);font-size:13px">Clique em "Adicionar Origem" para definir de onde virá o valor</div>` :
          GF.origens.map((o, i) => gfOrigemRow(o, i)).join('')
        }
      </div>
      
      ${GF.origens.length > 0 ? `
        <div style="margin-top:14px;padding-top:14px;border-top:1px solid var(--border);display:flex;justify-content:flex-end">
          <div style="font-size:14px"><strong>Total:</strong> <span style="color:var(--blue);font-weight:600">${gfFmt(totalOrigens)}</span></div>
        </div>
      ` : ''}
    </div>
    
    <div style="display:flex;gap:10px;justify-content:flex-end">
      <button class="btn" onclick="gfNav('alteracoes')">Cancelar</button>
      <button class="btn btn-primary" onclick="gfSubmitAlteracao()">Submeter Alteração</button>
    </div>
  `;
}

function gfOrigemRow(o, idx) {
  return `
    <div style="display:grid;grid-template-columns:1fr 1fr 120px 40px;gap:12px;margin-bottom:10px;padding:12px;background:#F8FAFC;border-radius:6px">
      <div>
        <label style="display:block;font-size:11px;color:var(--text3);margin-bottom:4px">Projeto Origem</label>
        <select id="gf-orig-proj-${idx}" onchange="gfOrigProjChange(${idx})" style="width:100%;padding:8px;border:1px solid var(--border);border-radius:4px;font-size:13px">
          <option value="">Selecionar...</option>
          ${GF_PROJETOS.map(p => `<option value="${p.id}" ${o.projetoId === p.id ? 'selected' : ''}>${p.nome}</option>`).join('')}
        </select>
      </div>
      <div>
        <label style="display:block;font-size:11px;color:var(--text3);margin-bottom:4px">Rúbrica Origem</label>
        <select id="gf-orig-rub-${idx}" style="width:100%;padding:8px;border:1px solid var(--border);border-radius:4px;font-size:13px">
          ${o.projetoId ? gfOrigRubOptions(o.projetoId, o.rubricaId) : '<option value="">Selecionar projeto...</option>'}
        </select>
      </div>
      <div>
        <label style="display:block;font-size:11px;color:var(--text3);margin-bottom:4px">Valor (€)</label>
        <input type="number" id="gf-orig-valor-${idx}" value="${o.valor || ''}" step="0.01" onchange="gfOrigValorChange(${idx})" style="width:100%;padding:8px;border:1px solid var(--border);border-radius:4px;font-size:13px">
      </div>
      <div style="display:flex;align-items:flex-end">
        <button onclick="gfRemoveOrigem(${idx})" style="width:36px;height:36px;border:none;background:#FEE2E2;color:#991B1B;border-radius:4px;cursor:pointer;font-size:16px">×</button>
      </div>
    </div>
  `;
}

function gfOrigRubOptions(projId, selectedId) {
  const rubs = GF_RUBRICAS.filter(r => r.projetoId === projId && r.tipo !== 'Receita');
  return '<option value="">Selecionar...</option>' + 
    rubs.map(r => {
      const disp = (r.despOrc || 0) - (r.despSol || 0);
      return `<option value="${r.id}" ${selectedId === r.id ? 'selected' : ''}>${r.titulo} (${gfFmt(disp)})</option>`;
    }).join('');
}

function gfAltDestProjChange() {
  const projId = parseInt(gfVal('gf-alt-dest-proj'));
  const rubSel = gfEl('gf-alt-dest-rub');
  if (!projId) {
    rubSel.innerHTML = '<option value="">Selecionar projeto primeiro...</option>';
    return;
  }
  const rubs = GF_RUBRICAS.filter(r => r.projetoId === projId);
  rubSel.innerHTML = '<option value="">Selecionar rúbrica...</option>' + 
    rubs.map(r => `<option value="${r.id}">${r.titulo}</option>`).join('');
}

function gfAddOrigem() {
  GF.origens.push({ projetoId: null, rubricaId: null, valor: 0 });
  gfRender();
}

function gfRemoveOrigem(idx) {
  GF.origens.splice(idx, 1);
  gfRender();
}

function gfOrigProjChange(idx) {
  const projId = parseInt(gfVal(`gf-orig-proj-${idx}`));
  GF.origens[idx].projetoId = projId;
  GF.origens[idx].rubricaId = null;
  gfEl(`gf-orig-rub-${idx}`).innerHTML = gfOrigRubOptions(projId, null);
}

function gfOrigValorChange(idx) {
  GF.origens[idx].valor = gfValN(`gf-orig-valor-${idx}`);
  const total = GF.origens.reduce((a, o) => a + (o.valor || 0), 0);
  gfEl('gf-alt-valor').value = total.toFixed(2);
}

async function gfSubmitAlteracao() {
  // Collect origem data from form
  GF.origens = GF.origens.map((o, i) => ({
    projetoId: parseInt(gfVal(`gf-orig-proj-${i}`)) || null,
    rubricaId: parseInt(gfVal(`gf-orig-rub-${i}`)) || null,
    valor: gfValN(`gf-orig-valor-${i}`)
  })).filter(o => o.projetoId && o.rubricaId && o.valor > 0);
  
  const destProjId = parseInt(gfVal('gf-alt-dest-proj'));
  const destRubId = parseInt(gfVal('gf-alt-dest-rub'));
  const justificacao = gfVal('gf-alt-justificacao').trim();
  const nivel = gfVal('gf-alt-nivel');
  const valorTotal = GF.origens.reduce((a, o) => a + o.valor, 0);
  
  if (!destProjId || !destRubId || !justificacao || GF.origens.length === 0) {
    showNotif('Preencha todos os campos obrigatórios e adicione pelo menos uma origem', 'error');
    return;
  }
  
  const destProj = GF_PROJETOS.find(p => p.id === destProjId);
  const destRub = GF_RUBRICAS.find(r => r.id === destRubId);
  
  try {
    // 1. Criar registo principal
    const bodyAlt = {
      __metadata: { type: 'SP.Data.FJ_x005f_AlteracoesOrcamentaisListItem' },
      Title: `Alteração: ${destRub?.titulo || 'Rúbrica'} — ${destProj?.nome || 'Projeto'}`,
      Justificacao: justificacao,
      ValorTotal: valorTotal,
      Estado: 'Pendente',
      NivelAprovacao: nivel,
      DestProjetoIdId: destProjId,
      DestRubricaIdId: destRubId,
      DataSubmissao: new Date().toISOString()
    };
    
    const res = await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_AlteracoesOrcamentais')/items`, bodyAlt, 'POST');
    const altId = res.d?.Id;
    
    if (!altId) throw new Error('Não foi possível obter o ID da alteração');
    
    // 2. Criar registos de origem
    for (const o of GF.origens) {
      const bodyOrig = {
        __metadata: { type: 'SP.Data.FJ_x005f_AlteracoesOrigensListItem' },
        Title: `Origem #${altId}`,
        AlteracaoIdId: altId,
        OrigProjetoIdId: o.projetoId,
        OrigRubricaIdId: o.rubricaId,
        Valor: o.valor
      };
      await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_AlteracoesOrigens')/items`, bodyOrig, 'POST');
    }
    
    showNotif('Alteração submetida com sucesso!', 'success');
    GF.origens = [];
    await gfLoadAll();
    gfNav('alteracoes');
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════════
// APROVAÇÕES
// ══════════════════════════════════════════════════════════════
function gfRenderAprovacoes() {
  const solsPend = GF_SOLICITACOES.filter(s => {
    if (_financeRole === 'area_director') return s.estado === 'Submetido';
    if (_financeRole === 'df') return ['Submetido', 'Aprovado Dir.'].includes(s.estado);
    return false;
  });
  
  const altsPend = GF_ALTERACOES.filter(a => {
    if (a.estado !== 'Pendente') return false;
    if (_financeRole === 'df') return true;
    if (_financeRole === 'pe' && ['df_pe', 'df_pe_pca'].includes(a.nivelAprovacao)) return !!a.aprovDFNome;
    if (_financeRole === 'pca' && a.nivelAprovacao === 'df_pe_pca') return !!a.aprovPENome;
    return false;
  });
  
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span><span>Aprovações</span>
    </div>
    
    <div style="font-size:20px;font-weight:600;margin-bottom:20px">Aprovações Pendentes</div>
    
    ${solsPend.length > 0 ? `
      <div style="font-size:14px;font-weight:600;margin-bottom:12px">Solicitações de Despesa (${solsPend.length})</div>
      ${solsPend.map(s => `
        <div class="card" style="margin-bottom:10px">
          <div style="display:flex;align-items:center;justify-content:space-between">
            <div>
              <div style="font-weight:500">${s.titulo}</div>
              <div style="font-size:12px;color:var(--text3)">${s.projetoNome} · ${s.autorNome} · ${gfDate(s.dataSubmissao)}</div>
            </div>
            <div style="display:flex;align-items:center;gap:12px">
              <span style="font-weight:600;color:var(--blue)">${gfFmt(s.valor)}</span>
              <button class="btn btn-sm" style="background:#10B981;color:white" onclick="gfAprovarSolicitacao(${s.id})">Aprovar</button>
              <button class="btn btn-sm" style="background:#EF4444;color:white" onclick="gfRejeitarSolicitacao(${s.id})">Rejeitar</button>
            </div>
          </div>
        </div>
      `).join('')}
    ` : ''}
    
    ${altsPend.length > 0 ? `
      <div style="font-size:14px;font-weight:600;margin:20px 0 12px">Alterações Orçamentais (${altsPend.length})</div>
      ${altsPend.map(a => `
        <div class="card" style="margin-bottom:10px">
          <div style="display:flex;align-items:center;justify-content:space-between">
            <div>
              <div style="font-weight:500">${a.titulo}</div>
              <div style="font-size:12px;color:var(--text3)">${a.solicitanteNome} · ${gfDate(a.dataSubmissao)}</div>
            </div>
            <div style="display:flex;align-items:center;gap:12px">
              <span style="font-weight:600;color:var(--blue)">${gfFmt(a.valorTotal)}</span>
              <button class="btn btn-sm" style="background:#10B981;color:white" onclick="gfAprovarAlteracao(${a.id})">Aprovar</button>
              <button class="btn btn-sm" style="background:#EF4444;color:white" onclick="gfRejeitarAlteracao(${a.id})">Rejeitar</button>
            </div>
          </div>
        </div>
      `).join('')}
    ` : ''}
    
    ${solsPend.length === 0 && altsPend.length === 0 ? 
      `<div class="card" style="padding:40px;text-align:center;color:var(--text3)">Não existem pedidos pendentes de aprovação</div>` : ''
    }
  `;
}

async function gfAprovarAlteracao(id) {
  if (!confirm('Aprovar esta alteração?')) return;
  try {
    const body = { __metadata: { type: 'SP.Data.FJ_x005f_AlteracoesOrcamentaisListItem' } };
    if (_financeRole === 'df') {
      body.AprovDFId = await getCurrentUserId();
      body.DataAprovDF = new Date().toISOString();
    } else if (_financeRole === 'pe') {
      body.AprovPEId = await getCurrentUserId();
    } else if (_financeRole === 'pca') {
      body.AprovPCAId = await getCurrentUserId();
      body.Estado = 'Aprovado';
    }
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_AlteracoesOrcamentais')/items(${id})`, body, 'MERGE');
    showNotif('Alteração aprovada!', 'success');
    await gfLoadAll();
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

async function gfRejeitarAlteracao(id) {
  const motivo = prompt('Motivo da rejeição (obrigatório):');
  if (!motivo || !motivo.trim()) { showNotif('Motivo obrigatório', 'error'); return; }
  try {
    const body = {
      __metadata: { type: 'SP.Data.FJ_x005f_AlteracoesOrcamentaisListItem' },
      Estado: 'Rejeitado',
      MotivoRejeicao: motivo.trim()
    };
    await spWrite(`${SP}/_api/web/lists/getbytitle('FJ_AlteracoesOrcamentais')/items(${id})`, body, 'MERGE');
    showNotif('Alteração rejeitada', 'success');
    await gfLoadAll();
  } catch (e) {
    showNotif('Erro: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════════
// FATURAS
// ══════════════════════════════════════════════════════════════
function gfRenderFaturas() {
  return `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;font-size:13px;color:var(--text3)">
      <a onclick="gfNav('home')" style="color:var(--blue);cursor:pointer">Início</a>
      <span>›</span><span>Faturas</span>
    </div>
    
    <div style="font-size:20px;font-weight:600;margin-bottom:20px">Faturas</div>
    
    <div class="card" style="padding:0;overflow:hidden">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:var(--bg)">
            <th style="padding:12px 16px;text-align:left;font-weight:500">Fornecedor</th>
            <th style="padding:12px 16px;text-align:left;font-weight:500">Solicitação</th>
            <th style="padding:12px 16px;text-align:right;font-weight:500">Valor</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Data</th>
            <th style="padding:12px 16px;text-align:center;font-weight:500">Estado</th>
          </tr>
        </thead>
        <tbody>
          ${GF_FATURAS.length === 0 ? `<tr><td colspan="5" style="padding:40px;text-align:center;color:var(--text3)">Sem faturas registadas</td></tr>` :
            GF_FATURAS.map(f => `
              <tr style="border-bottom:1px solid var(--border)">
                <td style="padding:14px 16px;font-weight:500">${f.fornecedor}</td>
                <td style="padding:14px 16px;font-size:12px">${f.solicitacaoNome || '—'}</td>
                <td style="padding:14px 16px;text-align:right">${gfFmt(f.valor)}</td>
                <td style="padding:14px 16px;text-align:center;font-size:12px">${gfDate(f.dataFatura)}</td>
                <td style="padding:14px 16px;text-align:center">${GF_BADGES[f.estado] || f.estado}</td>
              </tr>
            `).join('')
          }
        </tbody>
      </table>
    </div>
  `;
}

// Helper to get current user ID
async function getCurrentUserId() {
  try {
    const r = await spRead(`${SP}/_api/web/currentuser?$select=Id`);
    return r.d?.Id;
  } catch (e) {
    return null;
  }
}

// ═══════════════ INIT ═══════════════
// gfRender() moved to gfInit() - called when user switches to GF tab
// Auto-add primeira origem quando alterações carrega
// MutationObserver removed - handled by gfInit


