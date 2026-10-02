/* Aba Comercial: Tabela Comercial (tabelas de tarifa + historico tipo chat +
 * simulador) e Regra ICMS. Dados no Supabase do indicador-clientes, so com
 * login de admin (time PortoEx) -- ver supabase/schema_etapa6_comercial.sql
 * naquele repositorio. Script classico, carregado depois do script principal
 * do index.html; showTab() chama comercialAbrir(). */
(function () {
  'use strict';

  const SUPABASE_URL = 'https://fydoatntynvcwudxkhqv.supabase.co';
  const SUPABASE_KEY = 'sb_publishable_lXegUG6Ry1uZI3lqnbKyQg_lv-Ivzv6';
  const STATUS_LISTA = ['Enviada - aguardando aprovação', 'Em negociação', 'Aprovada', 'Reprovada'];
  const LIMITE_LINHAS = 500;

  let _cli = null;
  function cli() {
    // Mesma sessao do login da entrada (acesso.js), pra nao pedir senha duas vezes.
    if (!_cli) _cli = window.ACESSO_CLI || window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { storageKey: 'ansell-comercial-auth' } });
    return _cli;
  }

  const S = {
    perfil: null, email: '',
    tabelas: [], ultStatus: {}, selT: null, dadosT: {}, subaba: 'historico', msgs: [], novaT: null, filtroUf: '', filtroCidade: '',
    regras: [], selR: null, dadosR: {}, novaR: null,
    regraVigente: null, leitura: false,
    vigT: null, simc: { tabelaId: null, sug: {} },
  };

  // ---------- utilitarios ----------
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
  const fmtDH = iso => {
    if (!iso) return '';
    const d = new Date(iso); if (isNaN(d)) return '';
    const p = n => String(n).padStart(2, '0');
    return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  const fmtD = iso => fmtDH(iso).slice(0, 10);
  const brl = v => v == null || isNaN(v) ? '—' : 'R$ ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const pctTxt = f => f == null || isNaN(f) ? '—' : (f * 100).toLocaleString('pt-BR', { maximumFractionDigits: 3 }) + '%';
  function numBR(s) {
    let t = String(s == null ? '' : s).replace(/[^\d,.\-]/g, '');
    if (!t) return NaN;
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, '');
    return parseFloat(t);
  }
  function badgeClasse(status) {
    switch (status) {
      case 'Enviada - aguardando aprovação': return 'cm-b-env';
      case 'Em negociação': return 'cm-b-neg';
      case 'Aprovada': return 'cm-b-apr';
      case 'Reprovada': return 'cm-b-rep';
      case 'Vigente': return 'cm-b-vig';
      default: return 'cm-b-none';
    }
  }
  const container = qual => document.getElementById({ tabelas: 'tab-tabcomercial', regras: 'tab-regraicms', simulacao: 'tab-simulacaocusto' }[qual]);

  // ---------- parsing dos Excel ----------
  function lerExcel(arrayBuffer) {
    return XLSX.read(arrayBuffer, { type: 'array' });
  }
  const aoa = ws => XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const vazia = r => !r || r.every(v => v == null || String(v).trim() === '');

  function limiteDaFaixa(cab) {
    const t = norm(cab);
    if (t.includes('ACIMA')) return null;
    const m = String(cab).match(/R\$\s*([\d.]+(?:,\d+)?)/);
    if (!m) return undefined;
    return parseFloat(m[1].replace(/\./g, '').replace(',', '.'));
  }

  // Layout da tarifa: bloco de FAIXAS (cabecalho "ESTADO,CIDADE,...,FAIXA VALOR ...")
  // e bloco de CIDADES (cabecalho "ESTADO,CIDADE,PRAZO,COD. TARIFA,MINIMO,%,...").
  const simNao = v => {
    if (v == null || String(v).trim() === '') return undefined;
    const t = norm(v);
    if (['S', 'SIM', '1', 'TRUE', 'X', 'VERDADEIRO'].includes(t)) return true;
    if (['N', 'NAO', '0', 'FALSE', 'FALSO'].includes(t)) return false;
    return undefined;
  };

  function parseTabela(arrayBuffer) {
    const wb = lerExcel(arrayBuffer);
    for (const nome of wb.SheetNames) {
      const linhas = aoa(wb.Sheets[nome]);
      const cabs = [];
      linhas.forEach((r, i) => { if (r && norm(r[0]) === 'ESTADO' && norm(r[1]) === 'CIDADE') cabs.push(i); });
      if (!cabs.length) continue;
      const out = { aba: nome, faixas: { cabecalhos: [], linhas: [] }, cidades: [] };
      cabs.forEach(ini => {
        const cab = linhas[ini];
        const tipo = norm(cab[4]).startsWith('FAIXA') ? 'faixas' : 'cidades';
        const idx = {};
        cab.forEach((c, j) => { idx[norm(c)] = j; });
        for (let i = ini + 1; i < linhas.length; i++) {
          const r = linhas[i];
          if (vazia(r) || norm(r[0]) === 'ESTADO') break;
          if (tipo === 'faixas') {
            if (!out.faixas.cabecalhos.length) out.faixas.cabecalhos = cab.slice(4).filter(c => c != null && !norm(c).includes('REVERSO')).map(String);
            const iRevF = cab.findIndex(c => c != null && norm(c).includes('REVERSO'));
            const faixas = [];
            for (let j = 5; j < cab.length; j++) {
              if (cab[j] == null || norm(cab[j]).includes('REVERSO')) continue;
              faixas.push({ ate: limiteDaFaixa(cab[j]) === undefined ? null : limiteDaFaixa(cab[j]), pct: Number(r[j]) });
            }
            out.faixas.linhas.push({ uf: String(r[0]).trim(), cidade: String(r[1]).trim(), prazo: r[2], cod: r[3], minimo: Number(r[4]), faixas, rev: iRevF < 0 ? undefined : simNao(r[iRevF]) });
          } else {
            const col = nome2 => { const k = Object.keys(idx).find(x => x.startsWith(nome2)); return k == null ? -1 : idx[k]; };
            const iMin = col('MINIMO'), iPct = idx['%'], iIcms = col('ICMS'), iIbge = col('IBGE'), iVer = col('VERSAO'), iCod = col('COD');
            const minimo = Number(r[iMin]), pct = Number(r[iPct]);
            if (!r[1] || isNaN(minimo) || isNaN(pct) || !pct) continue;
            const kRev = Object.keys(idx).find(x => x.includes('REVERSO'));
            out.cidades.push({
              rev: kRev == null ? undefined : simNao(r[idx[kRev]]),
              uf: String(r[0]).trim(), cidade: String(r[1]).trim(), prazo: r[2], cod: r[iCod] == null ? '' : String(r[iCod]),
              minimo, pct, icms: r[iIcms] == null ? null : Number(r[iIcms]), ibge: r[iIbge] == null ? null : r[iIbge], versao: r[iVer] == null ? '' : String(r[iVer]),
            });
          }
        }
      });
      if (!out.cidades.length) continue;
      out.temColunaReverso = out.cidades.some(x => x.rev !== undefined);
      return out;
    }
    throw new Error('Não encontrei a tabela de tarifas neste arquivo (esperado: colunas ESTADO, CIDADE, PRAZO, CÓD. TARIFA, MÍNIMO, %, ...).');
  }

  function parseRegraIcms(arrayBuffer) {
    const wb = lerExcel(arrayBuffer);
    for (const nome of wb.SheetNames) {
      const linhas = aoa(wb.Sheets[nome]);
      const ini = linhas.findIndex(r => r && norm(r[0]).startsWith('UF ORIGEM') && norm(r[1]).startsWith('UF DESTINO'));
      if (ini < 0) continue;
      const matriz = {};
      let n = 0;
      for (let i = ini + 1; i < linhas.length; i++) {
        const r = linhas[i];
        if (vazia(r)) continue;
        const o = norm(r[0]), d = norm(r[1]);
        const v = parseFloat(String(r[2]).replace(',', '.'));
        if (!o || !d || isNaN(v) || v < 0 || v > 100) throw new Error(`Linha ${i + 1} inválida (origem "${r[0]}", destino "${r[1]}", alíquota "${r[2]}").`);
        (matriz[o] = matriz[o] || {})[d] = v;
        n++;
      }
      if (!n) continue;
      const ufs = [...new Set([...Object.keys(matriz), ...Object.values(matriz).flatMap(m => Object.keys(m))])].sort();
      return { aba: nome, ufs, matriz, linhas: n };
    }
    throw new Error('Não encontrei a regra de ICMS neste arquivo (esperado: colunas UF ORIGEM, UF DESTINO, ALIQUOTA).');
  }

  // ---------- calculo (por dentro) ----------
  const ORIGEM_PADRAO = { cidade: 'Santo André', uf: 'SP' };
  function origemTabela(d) { return d.origem || ORIGEM_PADRAO; }
  function chaveC(cidade, uf) { return norm(cidade) + '|' + String(uf).trim().toUpperCase(); }
  function linhasTrecho(d) { return [...d.cidades, ...d.faixas.linhas]; }

  function faixaPct(faixas, valor) {
    for (const f of faixas) if (f.ate != null && valor < f.ate) return f.pct;
    return faixas[faixas.length - 1].pct;
  }

  // Toda tabela tem uma origem (ex.: Santo André/SP) e as cidades sao os destinos.
  // Reverso (cidade -> origem) so vale no trecho marcado com rev=true.
  function simularTrecho(d, regra, orig, dest, valor) {
    const o = origemTabela(d), ko = chaveC(o.cidade, o.uf);
    const kOrig = chaveC(orig.cidade, orig.uf), kDest = chaveC(dest.cidade, dest.uf);
    const achar = p => linhasTrecho(d).find(x => chaveC(x.cidade, x.uf) === chaveC(p.cidade, p.uf));
    let linha, reverso;
    if (kOrig === ko) {
      linha = achar(dest); reverso = false;
      if (!linha) return { erro: 'COTAR OPERAÇÃO (destino fora da tabela)' };
    } else if (kDest === ko) {
      linha = achar(orig); reverso = true;
      if (!linha) return { erro: 'COTAR OPERAÇÃO (origem fora da tabela)' };
      if (!linha.rev) return { erro: `Trecho reverso não habilitado: ${orig.cidade}/${orig.uf} → ${o.cidade}/${o.uf} não aceita reverso nesta tabela.` };
    } else {
      return { erro: `Trecho não cadastrado: esta tabela tem origem ${o.cidade}/${o.uf}; a origem ou o destino precisa ser ${o.cidade}/${o.uf}.` };
    }
    const pct = linha.faixas ? faixaPct(linha.faixas, valor) : linha.pct;
    const mUf = regra && regra.matriz[orig.uf.toUpperCase()];
    let icms = mUf && mUf[dest.uf.toUpperCase()] != null ? mUf[dest.uf.toUpperCase()] / 100 : null;
    let fonteIcms = icms != null ? 'Regra ICMS vigente: ' + (S.regraVigente ? S.regraVigente.nome : '') : '';
    if (icms == null && !reverso && linha.icms != null && orig.uf.toUpperCase() === 'SP') { icms = linha.icms; fonteIcms = 'coluna ICMS da tabela (origem SP)'; }
    const frete = Math.max(valor * pct, linha.minimo);
    const r = { linha, reverso, prazo: linha.prazo, cod: linha.cod, minimo: linha.minimo, pct, frete, freteSemMinimo: valor * pct, minimoAplicado: valor * pct < linha.minimo, icms, fonteIcms };
    if (icms != null) { r.total = frete / (1 - icms); r.icmsValor = r.total - frete; }
    return r;
  }

  // ---------- login ----------
  async function carregarSessao() {
    const { data: { session } } = await cli().auth.getSession();
    if (!session) { S.perfil = null; return false; }
    const { data: p, error } = await cli().from('profiles').select('nome,role').eq('id', session.user.id).maybeSingle();
    if (error || !p) { S.perfil = null; return false; }
    S.email = session.user.email;
    S.perfil = p;
    S.leitura = p.role === 'comercial_leitura';
    return p.role === 'admin' || S.leitura;
  }

  function renderLogin(qual, msg) {
    const el = container(qual);
    el.innerHTML = `<div class="cm-card cm-login">
      <h3>Comercial — acesso da equipe PortoEx</h3>
      ${msg ? `<div class="cm-erro">${esc(msg)}</div>` : ''}
      <form data-act="login" data-qual="${qual}">
        <label>E-mail</label><input type="text" name="email" autocomplete="username" required>
        <label>Senha</label><input type="password" name="senha" autocomplete="current-password" required>
        <button class="cm-btn prim" type="submit" style="width:100%">Entrar</button>
      </form></div>`;
  }

  function topo(titulo, extra) {
    return `<div class="cm-topo"><h2>${titulo}</h2><div class="cm-user">${esc(S.perfil.nome)} (${esc(S.email)})
      ${extra || ''}<button class="cm-btn peq" data-act="sair">Sair</button></div></div>`;
  }

  // ---------- Tabela Comercial ----------
  async function carregarTabelas() {
    const t = await cli().from('comercial_tabelas').select('id,nome,arquivo_origem,criada_em,criada_por_nome,vigente').order('criada_em', { ascending: false });
    if (t.error) throw t.error;
    S.tabelas = t.data || [];
    if (S.leitura) { S.ultStatus = {}; return; }
    const h = await cli().from('comercial_historico').select('tabela_id,status,autor_nome,criado_em').not('status', 'is', null).order('criado_em', { ascending: false });
    if (h.error) throw h.error;
    S.ultStatus = {};
    // "Vigente" e lancamento automatico do botao de vigencia -- a vigencia
    // tem coluna propria, entao o selo de status mostra so os de negociacao.
    (h.data || []).forEach(x => { if (x.status !== 'Vigente' && !S.ultStatus[x.tabela_id]) S.ultStatus[x.tabela_id] = x; });
  }

  async function carregarRegraVigente() {
    const r = await cli().from('comercial_regras_icms').select('id,nome,dados').eq('vigente', true).maybeSingle();
    S.regraVigente = r.data ? { id: r.data.id, nome: r.data.nome, ...r.data.dados } : null;
  }

  function badgeStatus(id) {
    const u = S.ultStatus[id];
    const txt = u ? u.status : 'Sem status';
    return `<button class="cm-badge ${badgeClasse(u && u.status)}" data-act="abrir-chat" data-id="${id}" title="${u ? esc(u.autor_nome || '') + ' — ' + fmtDH(u.criado_em) : 'Clique pra registrar o primeiro status'}">${esc(txt)}</button>`;
  }

  function renderListaTabelas() {
    const linhas = S.tabelas.map(t => `<tr class="${t.id === S.selT ? 'sel' : ''}">
      <td><b>${esc(t.nome)}</b>${t.arquivo_origem ? `<div style="font-size:.7rem;color:#94a3b8">${esc(t.arquivo_origem)}</div>` : ''}</td>
      <td>${fmtDH(t.criada_em)}</td><td>${esc(t.criada_por_nome || '—')}</td>
      <td>${badgeStatus(t.id)}</td>
      <td>${t.vigente ? '<span class="cm-badge cm-b-vig">VIGENTE</span>' : `<button class="cm-btn peq" data-act="vigente-t" data-id="${t.id}">Marcar como vigente</button>`}</td>
      <td><button class="cm-btn peq" data-act="abrir-t" data-id="${t.id}">Abrir</button></td></tr>`).join('');
    return `<div class="cm-card"><div class="cm-topo" style="margin-bottom:10px"><h3 style="margin:0">Tabelas comerciais</h3>
      <button class="cm-btn prim" data-act="nova-t">+ Nova Tabela</button></div>
      ${S.tabelas.length ? `<div class="cm-scroll"><table class="cm-tabela"><thead><tr><th>Nome</th><th>Criada em</th><th>Criada por</th><th>Status</th><th>Vigente</th><th></th></tr></thead><tbody>${linhas}</tbody></table></div>`
        : '<div class="cm-vazio">Nenhuma tabela ainda. Clique em "+ Nova Tabela" e anexe o Excel.</div>'}</div>`;
  }

  function iniciarRev(n) {
    n.revSet = new Set(); n.padraoRev = null; n.ajustar = false; n.fUf = ''; n.fQ = '';
    if (n.dados.temColunaReverso) {
      linhasTrecho(n.dados).forEach(x => { if (x.rev) n.revSet.add(chaveC(x.cidade, x.uf)); });
      n.padraoRev = 'arquivo';
    }
  }

  function renderNovaTabela() {
    if (!S.novaT) return '';
    const n = S.novaT, r = n.resumo, d = n.dados;
    let reverso = '';
    if (d) {
      const total = linhasTrecho(d).length;
      const ufs = ufsDaTabela(d);
      reverso = `<h3 style="margin:16px 0 8px">Origem da tabela e trechos reversos</h3>
        <div class="cm-info">Toda cidade da tabela é cotada a partir da origem abaixo. <b>Aceita reverso</b> = o trecho também pode ser calculado no sentido contrário (da cidade para a origem).</div>
        <div class="cm-sim"><div><label>Origem da tabela — cidade *</label><input type="text" id="cm-nt-ocid" value="${esc(n.origemCidade)}"></div>
          <div><label>Origem da tabela — UF *</label><input type="text" id="cm-nt-ouf" maxlength="2" value="${esc(n.origemUf)}" style="text-transform:uppercase"></div></div>
        ${d.temColunaReverso ? '<div class="cm-info">O arquivo traz a coluna ACEITA REVERSO: usei os valores dela. Você pode ajustar abaixo antes de salvar.</div>'
          : `<label>Aceita reverso *</label><div style="margin-bottom:12px;font-size:.84rem">
            <label style="display:inline;font-weight:400"><input type="radio" name="nt-rev" value="todos" ${n.padraoRev === 'todos' ? 'checked' : ''} style="width:auto;margin:0 5px 0 0"> Todos os trechos aceitam reverso</label> &nbsp;&nbsp;
            <label style="display:inline;font-weight:400"><input type="radio" name="nt-rev" value="nenhum" ${n.padraoRev === 'nenhum' ? 'checked' : ''} style="width:auto;margin:0 5px 0 0"> Nenhum aceita (marco um a um)</label></div>`}
        <div style="font-size:.82rem;margin-bottom:8px"><b id="cm-rev-cont">${n.revSet.size.toLocaleString('pt-BR')} de ${total.toLocaleString('pt-BR')}</b> trechos aceitam reverso.
          <button type="button" class="cm-btn peq" data-act="nt-ajustar">${n.ajustar ? 'Ocultar ajuste por trecho' : 'Ajustar trecho a trecho'}</button></div>
        ${n.ajustar ? `<div class="cm-filtros cm-form">
          <div><label>Estado</label><select id="nt-f-uf"><option value="">Todos</option>${ufs.map(u => `<option ${u === n.fUf ? 'selected' : ''}>${u}</option>`).join('')}</select></div>
          <div style="flex:1;min-width:180px"><label>Buscar cidade</label><input type="text" id="nt-f-q" value="${esc(n.fQ)}" placeholder="Digite parte do nome..."></div>
          <div><label>&nbsp;</label><button type="button" class="cm-btn peq" data-act="nt-marcar">Marcar filtrados</button> <button type="button" class="cm-btn peq" data-act="nt-desmarcar">Desmarcar filtrados</button></div></div>
          <div class="cm-scroll" style="max-height:320px"><table class="cm-tabela"><thead><tr><th>Aceita reverso</th><th>UF</th><th>Cidade</th></tr></thead><tbody id="cm-rev-lista"></tbody></table></div>
          <div id="cm-rev-aviso" style="font-size:.74rem;color:#64748b;margin:6px 0 12px"></div>` : ''}`;
    }
    return `<div class="cm-card cm-form"><h3>Nova tabela comercial</h3>
      ${n.erro ? `<div class="cm-erro">${esc(n.erro)}</div>` : ''}
      <label>Nome da tabela</label><input type="text" id="cm-nt-nome" placeholder="Ex.: Tarifa Ansell Jul/2026 - V8" value="${esc(n.nome || '')}">
      <label>Arquivo Excel (mesmo layout da planilha de tarifa)</label><input type="file" id="cm-nt-arq" accept=".xlsx,.xls">
      ${r ? `<div class="cm-info">Arquivo lido: <b>${r.cidades}</b> cidades em <b>${r.ufs}</b> estados e <b>${r.faixas}</b> linha(s) de faixa por valor (${esc(r.nomesFaixa)}).</div>` : ''}
      ${reverso}
      <button class="cm-btn prim" data-act="salvar-t" ${n.salvando ? 'disabled' : ''}>${n.salvando ? 'Salvando...' : 'Salvar como nova tabela'}</button>
      <button class="cm-btn" data-act="cancelar-nova-t">Cancelar</button></div>`;
  }

  function trechosFiltradosNova() {
    const n = S.novaT, termo = norm(n.fQ || '');
    return linhasTrecho(n.dados).filter(x => (!n.fUf || x.uf === n.fUf) && (!termo || norm(x.cidade).includes(termo)));
  }
  function atualizarContRev() {
    const el = document.getElementById('cm-rev-cont');
    if (el && S.novaT && S.novaT.dados) el.textContent = `${S.novaT.revSet.size.toLocaleString('pt-BR')} de ${linhasTrecho(S.novaT.dados).length.toLocaleString('pt-BR')}`;
  }
  function atualizarRevLista() {
    const n = S.novaT, tb = document.getElementById('cm-rev-lista');
    if (!n || !n.dados || !tb) return;
    const f = trechosFiltradosNova();
    tb.innerHTML = f.slice(0, 300).map(x => { const k = chaveC(x.cidade, x.uf); return `<tr><td><input type="checkbox" class="nt-chk" data-k="${esc(k)}" ${n.revSet.has(k) ? 'checked' : ''}></td><td>${esc(x.uf)}</td><td>${esc(x.cidade)}</td></tr>`; }).join('') || '<tr><td colspan="3" class="cm-vazio">Nenhum resultado</td></tr>';
    const av = document.getElementById('cm-rev-aviso');
    if (av) av.textContent = f.length > 300 ? `Mostrando 300 de ${f.length.toLocaleString('pt-BR')} — refine o filtro; "Marcar/Desmarcar filtrados" vale para todos os ${f.length.toLocaleString('pt-BR')}.` : `${f.length.toLocaleString('pt-BR')} trecho(s).`;
  }
  function aplicarPadraoRev(valor) {
    const n = S.novaT;
    n.padraoRev = valor; n.revSet = new Set();
    if (valor === 'todos') linhasTrecho(n.dados).forEach(x => n.revSet.add(chaveC(x.cidade, x.uf)));
  }

  function renderDetalheTabela() {
    const t = S.tabelas.find(x => x.id === S.selT);
    if (!t) return '';
    const sub = (id, nome) => `<button class="${S.subaba === id ? 'ativo' : ''}" data-act="sub" data-sub="${id}">${nome}</button>`;
    return `<div class="cm-card"><div class="cm-topo" style="margin-bottom:0"><h3 style="margin:0">${esc(t.nome)} ${t.vigente ? '<span class="cm-badge cm-b-vig">VIGENTE</span>' : ''}</h3>
      <div style="font-size:.76rem;color:#64748b">Criada em ${fmtDH(t.criada_em)} por ${esc(t.criada_por_nome || '—')}</div></div>
      <div class="cm-sub">${S.leitura ? '' : sub('historico', 'Histórico / Status')}${sub('tarifas', 'Tarifas')}</div>
      <div id="cm-corpo-t"></div></div>`;
  }

  function renderChat() {
    const msgs = S.msgs.length ? S.msgs.map(m => `<div class="cm-msg"><div class="cab"><b>${esc(m.autor_nome || m.autor_email || '—')}</b>
      <span>${fmtDH(m.criado_em)}</span>${m.status ? `<span class="cm-badge ${badgeClasse(m.status)}">${esc(m.status)}</span>` : ''}</div>
      ${m.comentario ? `<div class="txt">${esc(m.comentario)}</div>` : ''}</div>`).join('')
      : '<div class="cm-vazio">Sem lançamentos ainda. Registre abaixo, por exemplo: "Enviada - aguardando aprovação".</div>';
    return `<div class="cm-chat" id="cm-chat">${msgs}</div>
      <div class="cm-form"><label>Novo lançamento (autor e data/hora ficam registrados automaticamente)</label>
      <select id="cm-st"><option value="">Somente comentário</option>${STATUS_LISTA.map(s => `<option>${esc(s)}</option>`).join('')}</select>
      <textarea id="cm-txt" rows="3" placeholder="Ex.: Enviada para a Juliana (Ansell) aguardando aprovação..."></textarea>
      <button class="cm-btn prim" data-act="enviar-msg">Registrar</button></div>`;
  }

  function ufsDaTabela(d) { return [...new Set([...d.cidades.map(c => c.uf), ...d.faixas.linhas.map(c => c.uf)])].sort(); }

  function renderTarifas(d) {
    const ufs = ufsDaTabela(d);
    const termo = norm(S.filtroCidade);
    const filtradas = d.cidades.filter(c => (!S.filtroUf || c.uf === S.filtroUf) && (!termo || norm(c.cidade).includes(termo)));
    const faixas = d.faixas.linhas.filter(c => (!S.filtroUf || c.uf === S.filtroUf) && (!termo || norm(c.cidade).includes(termo)));
    const cabF = d.faixas.cabecalhos, o = origemTabela(d);
    const totalRev = linhasTrecho(d).filter(x => x.rev).length;
    const blocoFaixas = faixas.length ? `<h3 style="margin:14px 0 8px">Cidades com tarifa por faixa de valor da mercadoria</h3>
      <div class="cm-scroll" style="max-height:none"><table class="cm-tabela"><thead><tr><th>UF</th><th>Cidade</th><th>Prazo</th><th>Cód.</th>${cabF.map(c => `<th>${esc(c)}</th>`).join('')}<th>Aceita reverso</th></tr></thead><tbody>
      ${faixas.map(f => `<tr><td>${esc(f.uf)}</td><td>${esc(f.cidade)}</td><td>${esc(f.prazo)}</td><td>${esc(f.cod)}</td><td>${brl(f.minimo)}</td>${f.faixas.map(x => `<td>${pctTxt(x.pct)}</td>`).join('')}<td>${f.rev ? 'Sim' : 'Não'}</td></tr>`).join('')}</tbody></table></div>` : '';
    const mostrar = filtradas.slice(0, LIMITE_LINHAS);
    return `<div class="cm-info">Origem da tabela: <b>${esc(o.cidade)}/${esc(o.uf)}</b> · trechos que aceitam reverso: <b>${totalRev.toLocaleString('pt-BR')}</b> de ${linhasTrecho(d).length.toLocaleString('pt-BR')}</div>
      <div class="cm-filtros cm-form">
        <div><label>Estado</label><select id="cm-f-uf"><option value="">Todos</option>${ufs.map(u => `<option ${u === S.filtroUf ? 'selected' : ''}>${u}</option>`).join('')}</select></div>
        <div style="flex:1;min-width:200px"><label>Buscar cidade</label><input type="text" id="cm-f-cid" value="${esc(S.filtroCidade)}" placeholder="Digite parte do nome..."></div>
        <div><label>&nbsp;</label><button class="cm-btn verde" data-act="export-t">⬇ Exportar Excel</button></div></div>
      ${blocoFaixas}
      <h3 style="margin:14px 0 8px">Tarifas por cidade <span style="font-weight:400;color:#64748b;font-size:.78rem">(${filtradas.length.toLocaleString('pt-BR')} de ${d.cidades.length.toLocaleString('pt-BR')})</span></h3>
      <div class="cm-scroll"><table class="cm-tabela"><thead><tr><th>UF</th><th>Cidade</th><th>Prazo</th><th>Cód. Tarifa</th><th>Mínimo</th><th>%</th><th>Valor Mínimo</th><th>ICMS (tabela)</th><th>IBGE</th><th>Aceita reverso</th></tr></thead><tbody>
      ${mostrar.map(c => `<tr><td>${esc(c.uf)}</td><td>${esc(c.cidade)}</td><td>${esc(c.prazo)}</td><td>${esc(c.cod)}</td><td>${brl(c.minimo)}</td><td>${pctTxt(c.pct)}</td><td>${brl(c.minimo / c.pct)}</td><td>${pctTxt(c.icms)}</td><td>${esc(c.ibge == null ? '' : c.ibge)}</td><td>${c.rev ? 'Sim' : 'Não'}</td></tr>`).join('') || '<tr><td colspan="10" class="cm-vazio">Nenhum resultado</td></tr>'}</tbody></table></div>
      ${filtradas.length > LIMITE_LINHAS ? `<div style="font-size:.74rem;color:#64748b;margin-top:6px">Mostrando ${LIMITE_LINHAS} de ${filtradas.length.toLocaleString('pt-BR')} — refine o filtro. O Excel leva todas as linhas filtradas.</div>` : ''}`;
  }

  async function renderCorpoT() {
    const el = document.getElementById('cm-corpo-t');
    if (!el) return;
    if (S.subaba === 'historico') { el.innerHTML = renderChat(); const c = document.getElementById('cm-chat'); if (c) c.scrollTop = c.scrollHeight; return; }
    el.innerHTML = '<div class="cm-vazio">Carregando tabela...</div>';
    const d = await dadosTabela(S.selT);
    el.innerHTML = renderTarifas(d);
  }

  async function dadosTabela(id) {
    if (!S.dadosT[id]) {
      const r = await cli().from('comercial_tabelas').select('dados').eq('id', id).single();
      if (r.error) throw r.error;
      S.dadosT[id] = r.data.dados;
    }
    return S.dadosT[id];
  }

  async function carregarMsgs() {
    const r = await cli().from('comercial_historico').select('*').eq('tabela_id', S.selT).order('criado_em', { ascending: true });
    if (r.error) throw r.error;
    S.msgs = r.data || [];
  }

  async function renderTabelas() {
    const el = container('tabelas');
    el.innerHTML = topo('Tabela Comercial') + '<div id="cm-lista"></div><div id="cm-nova"></div><div id="cm-detalhe"></div>';
    if (S.leitura) { S.selT = S.tabelas.length ? S.tabelas[0].id : null; S.subaba = 'tarifas'; }
    await redesenharTabelas();
  }
  async function redesenharTabelas() {
    document.getElementById('cm-lista').innerHTML = S.leitura
      ? (S.tabelas.length ? '' : '<div class="cm-erro">Nenhuma tabela vigente disponível no momento.</div>')
      : renderListaTabelas();
    document.getElementById('cm-nova').innerHTML = renderNovaTabela();
    if (S.novaT && S.novaT.ajustar) atualizarRevLista();
    document.getElementById('cm-detalhe').innerHTML = renderDetalheTabela();
    if (S.selT) await renderCorpoT();
  }

  // ---------- Simulacao Custo (usa a tabela VIGENTE) ----------
  const SEM_TABELA = 'Trecho sem Tabela Negociada';
  function pontoTxt(p) { return p.cidade + '/' + p.uf; }

  async function carregarTabelaVigente() {
    const r = await cli().from('comercial_tabelas').select('id,nome,criada_em,criada_por_nome,dados').eq('vigente', true).maybeSingle();
    if (r.error) throw r.error;
    S.vigT = r.data || null;
    if (S.vigT && S.simc.tabelaId !== S.vigT.id) {
      const o = origemTabela(S.vigT.dados);
      S.simc = { tabelaId: S.vigT.id, orig: { texto: pontoTxt(o), sel: { cidade: o.cidade, uf: o.uf } }, dest: { texto: '', sel: null }, qtd: '', peso: '', valor: '', res: null, erro: '', sug: {} };
    }
  }

  // Cidades da tabela + a propria origem; "rev" = aceita reverso.
  function pontosDaTabela(d) {
    const o = origemTabela(d), ko = chaveC(o.cidade, o.uf), m = new Map();
    linhasTrecho(d).forEach(x => m.set(chaveC(x.cidade, x.uf), { cidade: x.cidade, uf: x.uf, rev: !!x.rev }));
    const atual = m.get(ko);
    m.set(ko, { cidade: o.cidade, uf: o.uf, rev: atual ? atual.rev : false, origem: true });
    return [...m.values()];
  }

  // Origem: so a origem da tabela e as cidades com reverso. Destino: tudo,
  // exceto quando a origem escolhida nao e a da tabela (ai so volta pra origem).
  function sugestoes(qual, termo) {
    const d = S.vigT.dados, s = S.simc, o = origemTabela(d);
    let pts = pontosDaTabela(d);
    if (qual === 'orig') pts = pts.filter(p => p.origem || p.rev);
    else if (s.orig.sel && chaveC(s.orig.sel.cidade, s.orig.sel.uf) !== chaveC(o.cidade, o.uf)) pts = pts.filter(p => p.origem);
    const q = norm(termo).replace(/\s*\/\s*/g, '/');
    pts = pts.filter(p => !q || norm(pontoTxt(p)).includes(q));
    pts.sort((a, b) => ((norm(a.cidade).startsWith(q) ? 0 : 1) - (norm(b.cidade).startsWith(q) ? 0 : 1)) || a.cidade.localeCompare(b.cidade, 'pt-BR'));
    return pts.slice(0, 12);
  }

  function mostrarSugestoes(qual) {
    const lista = document.getElementById('sc-' + qual + '-l');
    if (!lista || !S.vigT) return;
    const sug = sugestoes(qual, S.simc[qual].texto);
    S.simc.sug[qual] = sug;
    (S.simc.ativo = S.simc.ativo || { orig: -1, dest: -1 })[qual] = sug.length ? 0 : -1;
    lista.innerHTML = sug.length
      ? sug.map((p, i) => `<div class="cm-ac-item${i === 0 ? ' ativo' : ''}" data-qual="${qual}" data-i="${i}">${esc(pontoTxt(p))}${p.origem ? ' <span style="color:#94a3b8;font-size:.7rem">(origem da tabela)</span>' : ''}</div>`).join('')
      : '<div class="vazio">Sem tabela negociada para esta cidade</div>';
    lista.style.display = 'block';
  }
  function marcarAtivo(qual) {
    const itens = document.querySelectorAll('#sc-' + qual + '-l .cm-ac-item');
    const a = (S.simc.ativo || {})[qual];
    itens.forEach((el, i) => {
      el.classList.toggle('ativo', i === a);
      if (i === a) el.scrollIntoView({ block: 'nearest' });
    });
  }
  function esconderSugestoes(qual) { const l = document.getElementById('sc-' + qual + '-l'); if (l) l.style.display = 'none'; }
  function escolherPonto(qual, i) {
    const p = (S.simc.sug[qual] || [])[i];
    if (!p) return;
    S.simc[qual] = { texto: pontoTxt(p), sel: { cidade: p.cidade, uf: p.uf } };
    const inp = document.getElementById('sc-' + qual + '-q');
    if (inp) inp.value = S.simc[qual].texto;
    esconderSugestoes(qual);
    const prox = document.getElementById(qual === 'orig' ? 'sc-dest-q' : 'sc-qtd');
    if (prox) prox.focus();
  }
  function resolverPonto(qual) {
    const f = S.simc[qual];
    if (f.sel) return f.sel;
    // texto digitado a mao: vale se bater exatamente com uma cidade da tabela
    // (a validade do trecho, reverso etc., quem decide e o simularTrecho)
    const ex = pontosDaTabela(S.vigT.dados).filter(p => norm(pontoTxt(p)) === norm(f.texto));
    return ex.length === 1 ? { cidade: ex[0].cidade, uf: ex[0].uf } : null;
  }

  function lerCamposSimc() {
    const v = id => { const e = document.getElementById(id); return e ? e.value : ''; };
    const s = S.simc;
    s.qtd = v('sc-qtd'); s.peso = v('sc-peso'); s.valor = v('sc-val');
  }

  function renderSimulacao() {
    const el = container('simulacao');
    const t = S.vigT, s = S.simc;
    let corpo;
    if (!t) {
      corpo = '<div class="cm-erro">Nenhuma tabela está marcada como vigente. Em Tabela Comercial, abra a tabela correta e use "Marcar como vigente".</div>';
    } else {
      const regra = S.regraVigente, o = origemTabela(t.dados);
      let res = '';
      if (s.erro) res = `<div class="cm-erro"><b>${esc(s.erro)}</b>${s.detalhe ? `<div style="font-size:.76rem;margin-top:3px">${esc(s.detalhe)}</div>` : ''}</div>`;
      else if (s.res) {
        const r = s.res, peso = numBR(s.peso), qtd = numBR(s.qtd), og = s.orig.sel, ds = s.dest.sel, valorNf = numBR(s.valor);
        // ICMS fora do percentual: o cliente credita, entao conta so o valor do frete.
        const pctNf = (r.frete / valorNf * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
        res = `<div class="cm-card"><h3>Resultado da simulação ${r.reverso ? '<span class="cm-badge cm-b-neg">trecho reverso</span>' : ''}</h3>
          <div class="cm-resw"><div class="cm-res">
            <div class="lin"><span>Trecho</span><span>${esc(pontoTxt(og))} → ${esc(pontoTxt(ds))}</span></div>
            <div class="lin"><span>Valor da mercadoria</span><span>${brl(valorNf)}</span></div>
            <div class="lin"><span>Quantidade</span><span>${qtd.toLocaleString('pt-BR')}</span></div>
            <div class="lin"><span>Peso</span><span>${peso.toLocaleString('pt-BR')} kg</span></div></div>
          <div class="cm-res">
            <div class="lin"><span>Valor do frete ${r.minimoAplicado ? '(mínimo da tabela aplicado)' : '(valor da mercadoria × ' + pctTxt(r.pct) + ')'}</span><span>${brl(r.frete)}</span></div>
            ${r.minimoAplicado ? `<div class="lin"><span>Valor × ${pctTxt(r.pct)} = ${brl(r.freteSemMinimo)}, abaixo do mínimo de ${brl(r.minimo)}</span><span></span></div>` : ''}
            <div class="lin"><span>ICMS ${r.icms != null ? '(' + pctTxt(r.icms) + ' por dentro: frete ÷ (1 − ' + pctTxt(r.icms) + '))' : ''}</span><span>${r.icms != null ? brl(r.icmsValor) : '—'}</span></div>
            <div class="lin tot"><span>FRETE TOTAL</span><span>${r.total != null ? brl(r.total) : 'importe a Regra ICMS'}</span></div>
            <div class="lin"><span>Valor do frete (sem ICMS) sobre o valor da nota fiscal</span><span><b>${pctNf}</b></span></div>
            <div class="lin"><span>Cidade/UF e prazo</span><span>${esc(pontoTxt(r.reverso ? og : ds))} — ${esc(r.prazo)} dia(s) útil(eis)</span></div></div></div>
          <div style="font-size:.76rem;color:#64748b;margin-top:10px;line-height:1.7">${esc(pontoTxt(og))} &gt; ${esc(pontoTxt(ds))}<br>Tabela: % de ${pctTxt(r.pct)} · valor mínimo de ${brl(r.minimo)}<br>Valor de carga mínima de ${brl(r.minimo / r.pct)}<br>${r.frete / valorNf > r.pct * (1 + 1e-9) ? '<b style="color:#b91c1c">Venda Abaixo do Valor Mínimo</b>' : '<b style="color:#166534">Venda Acima do Valor Mínimo</b>'}</div>
          <div style="margin-top:14px"><button class="cm-btn prim" data-act="sc-limpar">Limpar e fazer nova simulação</button></div></div>`;
      }
      corpo = `<div class="cm-card cm-form"><label>Tabela vigente (usada no cálculo)</label>
        <div style="margin-bottom:14px"><b>${esc(t.nome)}</b> <span class="cm-badge cm-b-vig">VIGENTE</span>
        <span style="font-size:.74rem;color:#64748b"> · origem ${esc(pontoTxt(o))} · criada em ${fmtDH(t.criada_em)} por ${esc(t.criada_por_nome || '—')}</span></div>
        <div class="cm-sim">
          <div style="position:relative"><label>Origem (cidade/UF) *</label><input type="text" id="sc-orig-q" autocomplete="off" value="${esc(s.orig.texto)}" placeholder="Digite a cidade ou UF..."><div class="cm-ac" id="sc-orig-l"></div></div>
          <div style="position:relative"><label>Destino (cidade/UF) *</label><input type="text" id="sc-dest-q" autocomplete="off" value="${esc(s.dest.texto)}" placeholder="Digite a cidade ou UF..."><div class="cm-ac" id="sc-dest-l"></div></div>
          <div><label>Quantidade (volumes) *</label><input type="text" id="sc-qtd" value="${esc(s.qtd)}" inputmode="numeric" placeholder="Ex.: 12"></div>
          <div><label>Peso (kg) *</label><input type="text" id="sc-peso" value="${esc(s.peso)}" inputmode="decimal" placeholder="Ex.: 350,5"></div>
          <div><label>Valor da mercadoria (R$) *</label><input type="text" id="sc-val" value="${esc(s.valor)}" inputmode="decimal" placeholder="Ex.: 150.000,00"></div></div>
        <button class="cm-btn prim" data-act="sc-calcular">Calcular</button> <button class="cm-btn" data-act="sc-limpar">Limpar</button>
        <div style="font-size:.72rem;color:#94a3b8;margin-top:8px">* obrigatórios. Digite e escolha na lista. A origem ou o destino precisa ser ${esc(pontoTxt(o))} (reverso só nos trechos que aceitam). Quantidade e peso são informativos (não entram na conta).</div></div>${res}`;
    }
    el.innerHTML = topo('Simulação Custo') + corpo;
  }

  function calcularSimulacao() {
    lerCamposSimc();
    const s = S.simc, d = S.vigT.dados;
    s.res = null; s.erro = ''; s.detalhe = '';
    const og = resolverPonto('orig'), ds = resolverPonto('dest');
    const qtd = numBR(s.qtd), peso = numBR(s.peso), valor = numBR(s.valor);
    const falta = [];
    if (!(s.orig.texto || '').trim()) falta.push('Origem');
    if (!(s.dest.texto || '').trim()) falta.push('Destino');
    if (!(qtd >= 1)) falta.push('Quantidade');
    if (!(peso > 0)) falta.push('Peso');
    if (!(valor > 0)) falta.push('Valor da mercadoria');
    if (falta.length) s.erro = 'Preencha corretamente: ' + falta.join(', ') + '.';
    else if (!og || !ds) { s.erro = SEM_TABELA; s.detalhe = 'A cidade informada não consta na tabela vigente.'; }
    else {
      s.orig = { texto: pontoTxt(og), sel: og }; s.dest = { texto: pontoTxt(ds), sel: ds };
      const r = simularTrecho(d, S.regraVigente, og, ds, valor);
      if (r.erro) { s.erro = SEM_TABELA; s.detalhe = r.erro; } else s.res = r;
    }
    renderSimulacao();
  }

  // ---------- Regra ICMS ----------
  async function carregarRegras() {
    const r = await cli().from('comercial_regras_icms').select('id,nome,arquivo_origem,criada_em,criada_por_nome,vigente').order('criada_em', { ascending: false });
    if (r.error) throw r.error;
    S.regras = r.data || [];
  }

  function renderRegras() {
    const lista = S.regras.map(t => `<tr class="${t.id === S.selR ? 'sel' : ''}"><td><b>${esc(t.nome)}</b>${t.arquivo_origem ? `<div style="font-size:.7rem;color:#94a3b8">${esc(t.arquivo_origem)}</div>` : ''}</td>
      <td>${fmtDH(t.criada_em)}</td><td>${esc(t.criada_por_nome || '—')}</td>
      <td>${t.vigente ? '<span class="cm-badge cm-b-vig">VIGENTE</span>' : `<button class="cm-btn peq" data-act="vigente-r" data-id="${t.id}">Marcar como vigente</button>`}</td>
      <td><button class="cm-btn peq" data-act="abrir-r" data-id="${t.id}">Ver matriz</button></td></tr>`).join('');
    let nova = '';
    if (S.novaR) {
      const r = S.novaR.resumo;
      nova = `<div class="cm-card cm-form"><h3>Nova regra de ICMS</h3>${S.novaR.erro ? `<div class="cm-erro">${esc(S.novaR.erro)}</div>` : ''}
        <label>Nome da regra</label><input type="text" id="cm-nr-nome" placeholder="Ex.: Regra ICMS Out/2026" value="${esc(S.novaR.nome || '')}">
        <label>Arquivo Excel (colunas UF ORIGEM, UF DESTINO, ALIQUOTA)</label><input type="file" id="cm-nr-arq" accept=".xlsx,.xls">
        ${r ? `<div class="cm-info">Arquivo lido: <b>${r.linhas}</b> combinações entre <b>${r.ufs}</b> UFs.</div>` : ''}
        <button class="cm-btn prim" data-act="salvar-r" ${S.novaR.salvando ? 'disabled' : ''}>${S.novaR.salvando ? 'Salvando...' : 'Salvar como nova regra'}</button>
        <button class="cm-btn" data-act="cancelar-nova-r">Cancelar</button></div>`;
    }
    let matriz = '';
    if (S.selR && S.dadosR[S.selR]) {
      const d = S.dadosR[S.selR];
      const nm = (S.regras.find(x => x.id === S.selR) || {}).nome;
      matriz = `<div class="cm-card"><h3>Matriz de ICMS — ${esc(nm)} <span style="font-weight:400;color:#64748b;font-size:.78rem">(alíquota %, origem nas linhas, destino nas colunas; cálculo por dentro)</span></h3>
        <div class="cm-scroll" style="max-height:640px"><table class="cm-matriz"><thead><tr><th>Orig \\ Dest</th>${d.ufs.map(u => `<th>${u}</th>`).join('')}</tr></thead><tbody>
        ${d.ufs.map(o => `<tr><th>${o}</th>${d.ufs.map(x => { const v = d.matriz[o] && d.matriz[o][x]; return `<td class="${v === 7 ? 'l7' : v === 12 ? 'l12' : v == null ? '' : 'out'}">${v == null ? '' : v}</td>`; }).join('')}</tr>`).join('')}</tbody></table></div></div>`;
    }
    container('regras').innerHTML = topo('Regra ICMS') + `<div class="cm-card"><div class="cm-topo" style="margin-bottom:10px"><h3 style="margin:0">Regras de ICMS</h3>
      <button class="cm-btn prim" data-act="nova-r">+ Nova Regra</button></div>
      ${S.regras.length ? `<div class="cm-scroll"><table class="cm-tabela"><thead><tr><th>Nome</th><th>Criada em</th><th>Criada por</th><th>Vigente</th><th></th></tr></thead><tbody>${lista}</tbody></table></div>` : '<div class="cm-vazio">Nenhuma regra ainda. Clique em "+ Nova Regra" e anexe o Excel.</div>'}</div>${nova}${matriz}`;
  }

  // ---------- exportacao ----------
  async function exportarTabela() {
    const d = await dadosTabela(S.selT);
    const t = S.tabelas.find(x => x.id === S.selT);
    const termo = norm(S.filtroCidade);
    const ok = c => (!S.filtroUf || c.uf === S.filtroUf) && (!termo || norm(c.cidade).includes(termo));
    const o = origemTabela(d);
    const cidades = d.cidades.filter(ok).map(c => ({ ORIGEM: o.cidade + '/' + o.uf, ESTADO: c.uf, CIDADE: c.cidade, PRAZO: c.prazo, 'CÓD. TARIFA': c.cod, 'MÍNIMO': c.minimo, '%': c.pct, 'VALOR MÍNIMO': c.minimo / c.pct, ICMS: c.icms, IBGE: c.ibge, 'ACEITA REVERSO': c.rev ? 'S' : 'N' }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cidades), 'Tarifas');
    const faixas = d.faixas.linhas.filter(ok).map(f => {
      const o = { ESTADO: f.uf, CIDADE: f.cidade, PRAZO: f.prazo, 'CÓD. TARIFA': f.cod, 'MÍNIMO': f.minimo, 'ACEITA REVERSO': f.rev ? 'S' : 'N' };
      f.faixas.forEach((x, i) => { o[d.faixas.cabecalhos[i + 1] || 'FAIXA ' + (i + 1)] = x.pct; });
      return o;
    });
    if (faixas.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(faixas), 'Faixas por valor');
    XLSX.writeFile(wb, `tabela_comercial_${(t ? t.nome : 'tabela').replace(/[^\w\-]+/g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // ---------- eventos ----------
  async function erroTela(qual, e) {
    const el = container(qual);
    const faltaSql = e && /relation|does not exist|schema cache/i.test(e.message || '');
    el.innerHTML = topo({ tabelas: 'Tabela Comercial', regras: 'Regra ICMS', simulacao: 'Simulação Custo' }[qual]) + `<div class="cm-erro">${faltaSql
      ? 'As tabelas do Comercial ainda não existem no Supabase. Rode o arquivo supabase/schema_etapa6_comercial.sql no SQL Editor do projeto indicador-clientes.'
      : 'Erro: ' + esc((e && e.message) || e)}</div>`;
  }

  async function abrir(qual) {
    const el = container(qual);
    if (!el) return;
    el.innerHTML = '<div class="cm-vazio">Carregando...</div>';
    let ok;
    try { ok = await carregarSessao(); } catch (e) { renderLogin(qual, 'Não foi possível verificar o login: ' + e.message); return; }
    if (!ok) { renderLogin(qual, S.perfil ? 'Seu usuário não tem permissão (precisa ser admin da PortoEx).' : ''); return; }
    try {
      if (qual === 'tabelas') { await Promise.all([carregarTabelas(), carregarRegraVigente()]); await renderTabelas(); }
      else if (qual === 'simulacao') { await Promise.all([carregarTabelaVigente(), carregarRegraVigente()]); renderSimulacao(); }
      else if (S.leitura) { container('regras').innerHTML = topo('Regra ICMS') + '<div class="cm-erro">Sem permissão para esta tela.</div>'; }
      else { await carregarRegras(); renderRegras(); }
    } catch (e) { await erroTela(qual, e); }
  }
  window.comercialAbrir = qual => abrir(qual);

  function qualAtual() {
    if (document.getElementById('tab-tabcomercial').classList.contains('active')) return 'tabelas';
    return document.getElementById('tab-simulacaocusto').classList.contains('active') ? 'simulacao' : 'regras';
  }

  async function abrirTabela(id, sub) {
    S.selT = id; S.subaba = sub || 'historico'; S.filtroUf = ''; S.filtroCidade = '';
    if (S.subaba === 'historico') await carregarMsgs();
    await redesenharTabelas();
    document.getElementById('cm-detalhe').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function aoClicar(ev) {
    const b = ev.target.closest('[data-act]');
    if (!b || !(b.closest('#tab-tabcomercial') || b.closest('#tab-regraicms') || b.closest('#tab-simulacaocusto'))) return;
    const act = b.dataset.act;
    if (S.leitura && ['nova-t', 'salvar-t', 'vigente-t', 'enviar-msg', 'nova-r', 'salvar-r', 'vigente-r'].includes(act)) return;
    const qual = qualAtual();
    try {
      if (act === 'sair') { await cli().auth.signOut(); location.reload(); }
      else if (act === 'abrir-t') await abrirTabela(b.dataset.id, 'historico');
      else if (act === 'abrir-chat') await abrirTabela(b.dataset.id, 'historico');
      else if (act === 'sub') { S.subaba = b.dataset.sub; if (S.subaba === 'historico') await carregarMsgs(); await redesenharTabelas(); }
      else if (act === 'nova-t') { S.novaT = { nome: '', resumo: null, dados: null, arquivo: '', origemCidade: ORIGEM_PADRAO.cidade, origemUf: ORIGEM_PADRAO.uf, revSet: new Set(), padraoRev: null, ajustar: false, fUf: '', fQ: '' }; await redesenharTabelas(); }
      else if (act === 'nt-ajustar') { S.novaT.ajustar = !S.novaT.ajustar; await redesenharTabelas(); }
      else if (act === 'nt-marcar' || act === 'nt-desmarcar') {
        trechosFiltradosNova().forEach(x => { const k = chaveC(x.cidade, x.uf); if (act === 'nt-marcar') S.novaT.revSet.add(k); else S.novaT.revSet.delete(k); });
        if (S.novaT.padraoRev === 'todos' || S.novaT.padraoRev === 'nenhum') S.novaT.padraoRev = 'ajustado';
        atualizarRevLista(); atualizarContRev();
      }
      else if (act === 'cancelar-nova-t') { S.novaT = null; await redesenharTabelas(); }
      else if (act === 'salvar-t') await salvarTabela();
      else if (act === 'vigente-t') {
        const t = S.tabelas.find(x => x.id === b.dataset.id);
        if (!confirm(`Marcar "${t.nome}" como a tabela VIGENTE? A que está vigente hoje deixa de ser.`)) return;
        const r = await cli().rpc('comercial_marcar_tabela_vigente', { p_id: b.dataset.id });
        if (r.error) throw r.error;
        await carregarTabelas(); if (S.selT) await carregarMsgs(); await redesenharTabelas();
      }
      else if (act === 'enviar-msg') {
        const st = document.getElementById('cm-st').value, txt = document.getElementById('cm-txt').value.trim();
        if (!st && !txt) { alert('Escolha um status ou escreva um comentário.'); return; }
        const r = await cli().from('comercial_historico').insert({ tabela_id: S.selT, status: st || null, comentario: txt || null });
        if (r.error) throw r.error;
        await Promise.all([carregarMsgs(), carregarTabelas()]); await redesenharTabelas();
      }
      else if (act === 'export-t') await exportarTabela();
      else if (act === 'sc-calcular') calcularSimulacao();
      else if (act === 'sc-limpar') { const o = origemTabela(S.vigT.dados); S.simc = { tabelaId: S.vigT.id, orig: { texto: pontoTxt(o), sel: { cidade: o.cidade, uf: o.uf } }, dest: { texto: '', sel: null }, qtd: '', peso: '', valor: '', res: null, erro: '', sug: {} }; renderSimulacao(); const d1 = document.getElementById('sc-dest-q'); if (d1) { d1.scrollIntoView({ block: 'center' }); d1.focus(); } }
      else if (act === 'nova-r') { S.novaR = { nome: '', resumo: null, dados: null, arquivo: '' }; renderRegras(); }
      else if (act === 'cancelar-nova-r') { S.novaR = null; renderRegras(); }
      else if (act === 'salvar-r') await salvarRegra();
      else if (act === 'abrir-r') {
        if (!S.dadosR[b.dataset.id]) {
          const r = await cli().from('comercial_regras_icms').select('dados').eq('id', b.dataset.id).single();
          if (r.error) throw r.error;
          S.dadosR[b.dataset.id] = r.data.dados;
        }
        S.selR = b.dataset.id; renderRegras();
      }
      else if (act === 'vigente-r') {
        const t = S.regras.find(x => x.id === b.dataset.id);
        if (!confirm(`Marcar "${t.nome}" como a regra de ICMS VIGENTE? O simulador passa a usar esta.`)) return;
        const r = await cli().rpc('comercial_marcar_regra_icms_vigente', { p_id: b.dataset.id });
        if (r.error) throw r.error;
        await carregarRegras(); renderRegras();
      }
    } catch (e) { alert('Erro: ' + (e.message || e)); }
  }

  async function salvarTabela() {
    const n = S.novaT;
    n.nome = (document.getElementById('cm-nt-nome').value || '').trim();
    if (!n.nome) { n.erro = 'Informe o nome da tabela.'; await redesenharTabelas(); return; }
    if (!n.dados) { n.erro = 'Anexe o arquivo Excel da tarifa.'; await redesenharTabelas(); return; }
    n.origemCidade = (n.origemCidade || '').trim(); n.origemUf = (n.origemUf || '').trim().toUpperCase();
    if (!n.origemCidade || n.origemUf.length !== 2) { n.erro = 'Informe a origem da tabela (cidade e UF com 2 letras).'; await redesenharTabelas(); return; }
    if (!n.padraoRev) { n.erro = 'Escolha se os trechos aceitam reverso (todos ou nenhum) e, se precisar, ajuste trecho a trecho.'; await redesenharTabelas(); return; }
    n.erro = ''; n.salvando = true; await redesenharTabelas();
    const d = n.dados;
    const marcar = x => ({ ...x, rev: n.revSet.has(chaveC(x.cidade, x.uf)) });
    const dadosFinal = { aba: d.aba, origem: { cidade: n.origemCidade, uf: n.origemUf }, faixas: { cabecalhos: d.faixas.cabecalhos, linhas: d.faixas.linhas.map(marcar) }, cidades: d.cidades.map(marcar) };
    const r = await cli().from('comercial_tabelas').insert({ nome: n.nome, arquivo_origem: n.arquivo, dados: dadosFinal }).select('id').single();
    if (r.error) { n.salvando = false; n.erro = 'Não foi possível salvar: ' + r.error.message; await redesenharTabelas(); return; }
    S.novaT = null;
    await carregarTabelas();
    await abrirTabela(r.data.id, 'historico');
  }

  async function salvarRegra() {
    const n = S.novaR;
    n.nome = (document.getElementById('cm-nr-nome').value || '').trim();
    if (!n.nome) { n.erro = 'Informe o nome da regra.'; renderRegras(); return; }
    if (!n.dados) { n.erro = 'Anexe o arquivo Excel da regra.'; renderRegras(); return; }
    n.erro = ''; n.salvando = true; renderRegras();
    const r = await cli().from('comercial_regras_icms').insert({ nome: n.nome, arquivo_origem: n.arquivo, dados: { ufs: n.dados.ufs, matriz: n.dados.matriz } }).select('id').single();
    if (r.error) { n.salvando = false; n.erro = 'Não foi possível salvar: ' + r.error.message; renderRegras(); return; }
    S.novaR = null; S.selR = null;
    await carregarRegras(); renderRegras();
  }

  async function aoMudar(ev) {
    const t = ev.target;
    try {
      if (t.id === 'cm-nt-arq' || t.id === 'cm-nr-arq') {
        const f = t.files[0]; if (!f) return;
        const nova = t.id === 'cm-nt-arq' ? S.novaT : S.novaR;
        nova.nome = (document.getElementById(t.id === 'cm-nt-arq' ? 'cm-nt-nome' : 'cm-nr-nome').value || '').trim();
        try {
          const buf = await f.arrayBuffer();
          if (t.id === 'cm-nt-arq') {
            const d = parseTabela(buf);
            nova.dados = d; nova.arquivo = f.name; nova.erro = ''; iniciarRev(nova);
            nova.resumo = { cidades: d.cidades.length, ufs: new Set(d.cidades.map(c => c.uf)).size, faixas: d.faixas.linhas.length, nomesFaixa: d.faixas.linhas.map(x => x.cidade).join(', ') || 'nenhuma' };
          } else {
            const d = parseRegraIcms(buf);
            nova.dados = d; nova.arquivo = f.name; nova.erro = '';
            nova.resumo = { linhas: d.linhas, ufs: d.ufs.length };
          }
        } catch (e) { nova.dados = null; nova.resumo = null; nova.erro = e.message; }
        if (t.id === 'cm-nt-arq') await redesenharTabelas(); else renderRegras();
      } else if (t.name === 'nt-rev') { aplicarPadraoRev(t.value); await redesenharTabelas(); }
      else if (t.classList && t.classList.contains('nt-chk')) {
        if (t.checked) S.novaT.revSet.add(t.dataset.k); else S.novaT.revSet.delete(t.dataset.k);
        if (S.novaT.padraoRev === 'todos' || S.novaT.padraoRev === 'nenhum') S.novaT.padraoRev = 'ajustado';
        atualizarContRev();
      }
      else if (t.id === 'nt-f-uf') { S.novaT.fUf = t.value; atualizarRevLista(); }
      else if (t.id === 'cm-f-uf') { S.filtroUf = t.value; await renderCorpoT(); }
    } catch (e) { alert('Erro: ' + (e.message || e)); }
  }

  let _deb = null;
  function aoDigitar(ev) {
    const t = ev.target;
    if (t.id === 'sc-orig-q' || t.id === 'sc-dest-q') {
      const qual = t.id === 'sc-orig-q' ? 'orig' : 'dest';
      S.simc[qual] = { texto: t.value, sel: null };
      mostrarSugestoes(qual);
      return;
    }
    if (S.novaT) {
      if (t.id === 'cm-nt-nome') { S.novaT.nome = t.value; return; }
      if (t.id === 'cm-nt-ocid') { S.novaT.origemCidade = t.value; return; }
      if (t.id === 'cm-nt-ouf') { S.novaT.origemUf = t.value; return; }
      if (t.id === 'nt-f-q') { S.novaT.fQ = t.value; atualizarRevLista(); return; }
    }
    if (t.id === 'cm-f-cid') {
      const campo = t.id, pos = t.selectionStart;
      S.filtroCidade = t.value;
      clearTimeout(_deb);
      _deb = setTimeout(async () => {
        await renderCorpoT();
        const novo = document.getElementById(campo);
        if (novo) { novo.focus(); try { novo.setSelectionRange(pos, pos); } catch (e) { /* tipo sem selecao */ } }
      }, 250);
    }
  }

  async function aoEnviar(ev) {
    const f = ev.target.closest('form[data-act="login"]');
    if (!f) return;
    ev.preventDefault();
    const qual = f.dataset.qual;
    const r = await cli().auth.signInWithPassword({ email: f.email.value.trim(), password: f.senha.value });
    if (r.error) { renderLogin(qual, 'E-mail ou senha incorretos.'); return; }
    await abrir(qual);
  }

  document.addEventListener('click', aoClicar);
  document.addEventListener('change', aoMudar);
  document.addEventListener('input', aoDigitar);
  document.addEventListener('submit', aoEnviar);
  document.addEventListener('keydown', ev => {
    const id = ev.target.id;
    if (id !== 'sc-orig-q' && id !== 'sc-dest-q') return;
    const qual = id === 'sc-orig-q' ? 'orig' : 'dest';
    const lista = document.getElementById('sc-' + qual + '-l');
    const aberta = lista && lista.style.display === 'block';
    const sug = S.simc.sug[qual] || [], n = sug.length;
    const at = S.simc.ativo = S.simc.ativo || { orig: -1, dest: -1 };
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
      ev.preventDefault();
      if (!aberta) { mostrarSugestoes(qual); return; }
      if (!n) return;
      at[qual] = ev.key === 'ArrowDown' ? (at[qual] + 1) % n : (at[qual] - 1 + n) % n;
      marcarAtivo(qual);
    } else if (ev.key === 'Enter') {
      if (aberta && n && at[qual] >= 0) { ev.preventDefault(); escolherPonto(qual, at[qual]); }
    } else if (ev.key === 'Escape') {
      esconderSugestoes(qual);
    }
  });
  document.addEventListener('focusin', ev => {
    const id = ev.target.id;
    if (S.vigT && (id === 'sc-orig-q' || id === 'sc-dest-q')) mostrarSugestoes(id === 'sc-orig-q' ? 'orig' : 'dest');
  });
  document.addEventListener('focusout', ev => {
    const id = ev.target.id;
    if (id === 'sc-orig-q' || id === 'sc-dest-q') setTimeout(() => esconderSugestoes(id === 'sc-orig-q' ? 'orig' : 'dest'), 150);
  });
  document.addEventListener('mousedown', ev => {
    const it = ev.target.closest && ev.target.closest('.cm-ac-item');
    if (it) { ev.preventDefault(); escolherPonto(it.dataset.qual, Number(it.dataset.i)); }
  });

  // Pagina aberta direto numa aba do Comercial (F5 ou link com #tabcomercial):
  // showTab() rodou antes deste script existir, entao abre aqui.
  if (window.COMERCIAL_PAGINA_PROPRIA) { /* comercial.html abre a aba por conta propria */ }
  else if (location.hash === '#tabcomercial') abrir('tabelas');
  else if (location.hash === '#regraicms') abrir('regras');
  else if (location.hash === '#simulacaocusto') abrir('simulacao');

  // Pra testes: expoe o parser e o calculo (sem rede).
  window._comercialTeste = { parseTabela, parseRegraIcms, simularTrecho, numBR };
})();
