/* Login na entrada do dashboard Ansell + menus liberados por perfil.
 * Carregado no <head> (antes do resto da pagina) pra tampar o dashboard
 * desde o primeiro instante ate o login ser validado.
 *
 * - admin (time PortoEx): todos os menus, incluindo Comercial.
 * - cliente vinculado a Ansell e ativo: so os paineis marcados no cadastro
 *   do cliente Ansell (admin.html do indicador-clientes -> paineis_permitidos).
 *   Comercial nunca aparece pra ele.
 * - comercial_leitura (convidado): so Tabela Comercial (apenas a vigente) e Simulacao Custo.
 * - qualquer outro: sem acesso.
 *
 * ATENCAO: os dados do dashboard estao dentro do proprio HTML, entao isto e
 * um bloqueio de tela, nao protecao de dados -- a protecao real continua
 * sendo o Cloudflare Access na frente do dominio.
 */
(function () {
  'use strict';

  const SUPABASE_URL = 'https://fydoatntynvcwudxkhqv.supabase.co';
  const SUPABASE_KEY = 'sb_publishable_lXegUG6Ry1uZI3lqnbKyQg_lv-Ivzv6';
  const ANSELL_CLIENTE_ID = '1b2b560e-e184-495f-bb06-c5041c3b33ab';
  const SO_ADMIN = new Set(['tabcomercial', 'regraicms', 'simulacaocusto']);

  const ACESSO = { pronto: false, admin: false, leitura: false, permitidos: null, perfil: null, email: '' };
  window.ACESSO = ACESSO;
  // "Posicao NF" acompanha o painel "Status NF" (statusnf) do cadastro.
  ACESSO.podeAbrir = function (id) {
    if (!ACESSO.pronto || ACESSO.admin) return true;
    // convidado comercial_leitura: so a tabela vigente e o simulador
    if (ACESSO.leitura) return id === 'tabcomercial' || id === 'simulacaocusto';
    if (SO_ADMIN.has(id)) return false;
    const chave = id === 'posicaonf' ? 'statusnf' : id;
    return !!(ACESSO.permitidos && ACESSO.permitidos.has(chave));
  };

  // ---- tampa a pagina ja (antes do body existir) ----
  const css = document.createElement('style');
  css.textContent = `
    html.acesso-pendente body { visibility: hidden !important; }
    #acesso-gate { position: fixed; inset: 0; z-index: 99999; background: #1a3a5c; display: flex; align-items: center; justify-content: center; font-family: 'Segoe UI', sans-serif; }
    #acesso-gate .card { background: #fff; border-radius: 12px; padding: 34px 30px; width: 100%; max-width: 360px; box-shadow: 0 10px 40px rgba(0,0,0,.25); }
    #acesso-gate h1 { font-size: 1.2rem; color: #1a3a5c; margin: 0 0 4px; }
    #acesso-gate p.sub { font-size: .82rem; color: #64748b; margin: 0 0 22px; }
    #acesso-gate label { display: block; font-size: .78rem; font-weight: 600; color: #334155; margin-bottom: 5px; }
    #acesso-gate input { width: 100%; box-sizing: border-box; border: 1px solid #e2e8f0; border-radius: 7px; padding: 9px 11px; font-size: .9rem; margin-bottom: 16px; }
    #acesso-gate button { width: 100%; background: #1a3a5c; color: #fff; border: 0; border-radius: 8px; padding: 11px; font-size: .92rem; font-weight: 700; cursor: pointer; }
    #acesso-gate button:disabled { opacity: .6; cursor: not-allowed; }
    #acesso-gate .erro { background: #fdecec; color: #b91c1c; border-radius: 7px; padding: 9px 11px; font-size: .82rem; margin-bottom: 16px; }
    #acesso-user { display: inline-flex; align-items: center; gap: 8px; color: #fff; font-size: .74rem; margin-left: 10px; }
    #acesso-user button { background: #e2e8f0; color: #334155; border: 0; border-radius: 6px; padding: 4px 10px; font-size: .72rem; font-weight: 600; cursor: pointer; }
  `;
  document.documentElement.appendChild(css);
  document.documentElement.classList.add('acesso-pendente');

  const gate = document.createElement('div');
  gate.id = 'acesso-gate';
  document.documentElement.appendChild(gate);

  function mostrarLogin(msg) {
    gate.innerHTML = `<div class="card"><h1>Indicador Ansell — PortoEx</h1><p class="sub">Entre com seu e-mail e senha.</p>
      ${msg ? `<div class="erro">${msg}</div>` : ''}
      <form id="acesso-form"><label>E-mail</label><input type="text" name="email" autocomplete="username" required>
      <label>Senha</label><input type="password" name="senha" autocomplete="current-password" required>
      <button type="submit">Entrar</button></form></div>`;
    document.getElementById('acesso-form').addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const f = ev.target, b = f.querySelector('button');
      b.disabled = true; b.textContent = 'Entrando…';
      const r = await cli.auth.signInWithPassword({ email: f.email.value.trim(), password: f.senha.value });
      if (r.error) { mostrarLogin('E-mail ou senha incorretos.'); return; }
      iniciar();
    });
  }

  function mostrarSemAcesso(msg) {
    gate.innerHTML = `<div class="card"><h1>Acesso não liberado</h1><p class="sub">${msg}</p>
      <button id="acesso-sair">Sair</button></div>`;
    document.getElementById('acesso-sair').addEventListener('click', async () => { await cli.auth.signOut(); location.reload(); });
  }

  if (!window.supabase) {
    gate.innerHTML = '<div class="card"><h1>Não foi possível carregar o login</h1><p class="sub">Verifique sua conexão com a internet e atualize a página.</p></div>';
    return;
  }
  // sessionStorage: a sessao vale so enquanto a aba/janela estiver aberta --
  // ao fechar e abrir de novo, pede login de novo.
  const cli = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { storage: window.sessionStorage, storageKey: 'ansell-auth' } });
  window.ACESSO_CLI = cli;

  async function iniciar() {
    gate.innerHTML = '<div class="card"><p class="sub" style="margin:0">Verificando acesso…</p></div>';
    const { data: { session } } = await cli.auth.getSession();
    if (!session) { mostrarLogin(''); return; }
    const { data: p, error } = await cli.from('profiles').select('nome,role,cliente_id,clientes(ativo,paineis_permitidos)').eq('id', session.user.id).maybeSingle();
    if (error || !p) { mostrarSemAcesso('Seu login existe, mas ainda não foi vinculado a um perfil. Fale com a PortoEx.'); return; }
    ACESSO.perfil = p; ACESSO.email = session.user.email;
    if (p.role === 'admin') { ACESSO.admin = true; }
    else if (p.role === 'comercial_leitura') { ACESSO.leitura = true; }
    else if (p.role === 'cliente' && p.cliente_id === ANSELL_CLIENTE_ID && p.clientes && p.clientes.ativo) {
      ACESSO.permitidos = new Set(p.clientes.paineis_permitidos || []);
    } else { mostrarSemAcesso('Seu usuário não tem acesso a este painel. Fale com a PortoEx.'); return; }
    if (window.ACESSO_SO_COMERCIAL && !ACESSO.admin && !ACESSO.leitura) { mostrarSemAcesso('Esta página é restrita à equipe comercial. Fale com a PortoEx.'); return; }
    ACESSO.pronto = true;
    gate.remove();
    document.documentElement.classList.remove('acesso-pendente');
    aplicarMenus();
  }

  function aplicarMenus() {
    const rodar = () => {
      const tabs = [...document.querySelectorAll('.tab[data-tab]')];
      tabs.forEach(t => { if (!ACESSO.podeAbrir(t.dataset.tab)) t.style.display = 'none'; });
      document.querySelectorAll('.tgrp').forEach(g => {
        const itens = [...g.querySelectorAll('.tmenu .tab[data-tab]')];
        if (itens.length && itens.every(t => t.style.display === 'none')) g.style.display = 'none';
      });
      // aba aberta no momento nao liberada -> vai pra primeira liberada
      const ativa = document.querySelector('.content.active');
      const idAtiva = ativa ? ativa.id.replace(/^tab-/, '') : '';
      if (!ACESSO.podeAbrir(idAtiva)) {
        const primeira = tabs.find(t => t.style.display !== 'none');
        if (primeira && window.showTab) window.showTab(primeira.dataset.tab);
      }
      const barra = document.querySelector('.tabs');
      if (barra && !document.getElementById('acesso-user')) {
        const u = document.createElement('div');
        u.id = 'acesso-user';
        u.innerHTML = `<span>${(ACESSO.perfil.nome || ACESSO.email).replace(/</g, '&lt;')}</span><button type="button">Sair</button>`;
        u.querySelector('button').addEventListener('click', async () => { await cli.auth.signOut(); location.reload(); });
        barra.appendChild(u);
      }
    };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', rodar); else rodar();
  }

  iniciar();
})();
