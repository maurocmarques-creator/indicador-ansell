/* Aba "Conferência de Cubagem": o usuario importa os DANFE (PDF) e a aba cruza cada item da nota
 * com a tabela de medidas (cubagem_medidas.json, importada de "Medidas Ansell.xls": medidas em cm
 * ja arredondadas, ",001" para baixo e o resto para cima) e mostra um grid por nota com
 * destinatario, municipio, UF e a cubagem. Cubagem do item = quantidade da nota x cubagem de 1
 * unidade da tabela (teto: nao considera que as unidades possam ir empacotadas em menos espaco).
 * Tudo roda no navegador (o PDF nao sai da maquina). Script classico, carregado depois do
 * script principal do index.html; showTab() chama cubagemAbrir(). IDs/classes com prefixo "cb". */
(function () {
  'use strict';

  const PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  const PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const UNID = '(?:PC|UN|UND|CX|PCT|CJ|PAR|PR|RL|MT|KG|LT|FD|PCS)';
  const RE_QTD = new RegExp('\\b' + UNID + '\\s+([\\d.]+,\\d{2,4})\\s+([\\d.]+,\\d{2,10})', 'g');
  const RE_TOK = /[A-Za-z0-9][A-Za-z0-9.\-/]{2,}/g;

  const S = { medidas: null, idx: null, notas: [], abertas: new Set(), ordem: { col: null, dir: 1 }, msg: '', lendo: false, pronto: false };
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fmt = (v, d) => v == null || isNaN(v) ? '—' : Number(v).toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d });
  const brl = v => v == null || isNaN(v) ? '—' : 'R$ ' + fmt(v, 2);
  const br = s => parseFloat(String(s).replace(/\./g, '').replace(',', '.'));
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();

  function carregarScript(url) {
    return new Promise((ok, erro) => {
      const s = document.createElement('script'); s.src = url; s.onload = ok; s.onerror = () => erro(new Error('Não consegui carregar ' + url));
      document.head.appendChild(s);
    });
  }

  async function prepararBase() {
    if (!window.pdfjsLib) await carregarScript(PDFJS);
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
    if (!S.medidas) {
      const r = await fetch('cubagem_medidas.json', { cache: 'no-cache' });
      if (!r.ok) throw new Error('Tabela de medidas (cubagem_medidas.json) não encontrada.');
      S.medidas = await r.json();
      S.idx = new Map();
      S.medidas.m.forEach(([cli, item, peg, c, l, a, pb, un, cubUn, conf]) => {
        const reg = { cliente: cli, item, pegada: peg, c, l, a, pb, un, cubUn, conf };
        new Set([String(item).toUpperCase(), String(peg).toUpperCase()]).forEach(k => { if (!S.idx.has(k)) S.idx.set(k, []); S.idx.get(k).push(reg); });
      });
    }
  }

  // ---------- leitura do PDF ----------
  // Junta os pedacos de texto de cada pagina em linhas (mesmo y) para o DANFE ficar parecido com o texto corrido.
  async function linhasDoPdf(buf) {
    const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
    const linhas = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const tc = await (await pdf.getPage(p)).getTextContent();
      const it = tc.items.filter(i => i.str && i.str.trim()).map(i => ({ s: i.str, x: i.transform[4], y: i.transform[5], w: i.width || 0 }));
      it.sort((a, b) => b.y - a.y || a.x - b.x);
      let grupo = [], yRef = null;
      const fecha = () => {
        if (!grupo.length) return;
        grupo.sort((a, b) => a.x - b.x);
        let t = '', fim = null;
        grupo.forEach(g => { t += (fim != null && g.x - fim > 1 ? ' ' : (fim != null ? '' : '')) + g.s; fim = g.x + g.w; });
        linhas.push(t.replace(/\s+/g, ' ').trim());
        grupo = [];
      };
      it.forEach(i => { if (yRef == null || Math.abs(i.y - yRef) > 2.5) { fecha(); yRef = i.y; } grupo.push(i); });
      fecha();
    }
    return linhas;
  }

  function cabecalhoDaNota(linhas, nomeArq) {
    const txt = linhas.join('\n');
    const cab = { arquivo: nomeArq };
    let chave = (nomeArq.match(/\d{44}/) || [])[0];
    if (!chave) { const m = txt.match(/(?:\d{4}\s){10}\d{4}/); if (m) chave = m[0].replace(/\s/g, ''); }
    let m = txt.match(/N[º°o.]*\s*(\d{3}\.\d{3}\.\d{3})/);
    cab.nf = chave ? String(parseInt(chave.slice(25, 34), 10)) : (m ? String(parseInt(m[1].replace(/\./g, ''), 10)) : '');
    m = txt.match(/Emiss.o:\s*(\d{2}\/\d{2}\/\d{4})/i); cab.emissao = m ? m[1] : '';
    m = txt.match(/Dest\/Reme:\s*(.+?)\s+Valor Total:\s*R\$\s*([\d.,]+)/i);
    cab.destinatario = m ? m[1].trim() : ''; cab.valor = m ? br(m[2]) : null;
    if (cab.valor == null) { m = txt.match(/VALOR TOTAL:\s*R\$\s*([\d.,]+)/i); if (m) cab.valor = br(m[1]); }
    if (!cab.destinatario) { const i = linhas.findIndex(l => /^NOME \/ RAZ.O SOCIAL CNPJ/.test(l)); if (i >= 0 && linhas[i + 1]) cab.destinatario = linhas[i + 1].replace(/\s+\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}.*$/, '').trim(); }
    let i = linhas.findIndex(l => /^MUNIC.PIO\s+UF/i.test(l));
    m = i >= 0 && linhas[i + 1] ? linhas[i + 1].match(/^(.+?)\s+([A-Z]{2})\s+\(?\d/) : null;
    cab.municipio = m ? m[1].trim() : ''; cab.uf = m ? m[2] : '';
    i = linhas.findIndex(l => /^QUANTIDADE\s/i.test(l));
    m = i >= 0 && linhas[i + 1] ? linhas[i + 1].match(/^(\d+)\s+\D.*?(\d[\d.]*,\d{3})/) : null;   // 1o numero com 3 decimais = peso bruto (com ou sem 'KG')
    cab.volumes = m ? parseInt(m[1], 10) : null; cab.peso = m ? br(m[2]) : null;
    m = txt.match(/Recebemos de (.+?) os produtos/i); cab.emitente = m ? m[1].trim() : '';
    const a = txt.toUpperCase().indexOf('DADOS DOS PRODUTOS'), b = txt.toUpperCase().indexOf('DADOS ADICIONAIS');
    cab.secao = a >= 0 && b > a ? txt.slice(a, b) : txt;
    return cab;
  }

  // Codigos da tabela que aparecem na secao de produtos (em ordem), cada um com a quantidade da proxima
  // linha de produto. O DANFE as vezes quebra o codigo no fim da linha ("HECV1P30AH" + "220" embaixo).
  function itensDaNota(secao) {
    const linhas = secao.split('\n'), achados = [];
    let pos = 0;
    linhas.forEach((ln, n) => {
      // primeiros "pedacos" das proximas 3 linhas (onde o resto de um codigo quebrado costuma cair)
      const prox = [1, 2, 3].map(d => n + d < linhas.length ? (linhas[n + d].trim().split(/\s+/)[0] || '').toUpperCase() : '').filter(Boolean);
      for (const m of ln.matchAll(RE_TOK)) {
        const tok = m[0].toUpperCase().replace(/\.+$/, '');
        if (S.idx.has(tok)) { achados.push({ pos: pos + m.index, cod: tok, aviso: '' }); continue; }
        if (tok.length < 8) continue;
        const inteiro = prox.map(p => tok + p).find(c => S.idx.has(c));
        if (inteiro) achados.push({ pos: pos + m.index, cod: inteiro, aviso: 'código quebrado em duas linhas no PDF' });
      }
      pos += ln.length + 1;
    });
    const qtds = [...secao.matchAll(RE_QTD)].map(m => ({ pos: m.index, q: br(m[1]) }));
    const usadas = new Set();
    const itens = achados.map(a => {
      const reg = S.idx.get(a.cod), x = reg[0], k = qtds.find(k => k.pos > a.pos), q = k ? k.q : undefined;
      if (k) usadas.add(k.pos);
      const obs = [];
      if (a.aviso) obs.push(a.aviso);
      if (new Set(reg.map(r => r.c + 'x' + r.l + 'x' + r.a)).size > 1) obs.push(reg.length + ' medidas na tabela; usada a 1ª');
      if (q == null) obs.push('quantidade não encontrada');
      if (x.conf === 'INCOERENTE') obs.push('medida por unidade do relatório pegada não confere com a caixa; usada caixa ÷ un. por caixa');
      const volCx = x.c * x.l * x.a / 1e6, cx = q == null ? null : Math.ceil(q / x.un - 1e-9);
      return { cod: a.cod, qtd: q == null ? null : q, c: x.c, l: x.l, a: x.a, un: x.un, cx, volCx, obs: obs.join('; ') };
    });
    // linhas de produto cuja quantidade nao foi usada por nenhum item da tabela: o codigo nao esta na tabela de medidas
    qtds.filter(k => !usadas.has(k.pos)).forEach(k => {
      const ini = secao.lastIndexOf('\n', k.pos) + 1, cod1 = (secao.slice(ini, k.pos).trim().split(/\s+/)[0] || '').replace(/\.+$/, '');
      const cod = /^[A-Za-z0-9][A-Za-z0-9.\-/]{3,}$/.test(cod1) && !/^\d{8}$/.test(cod1) ? cod1 : '(código não identificado)';
      itens.push({ cod, qtd: k.q, c: null, l: null, a: null, un: null, cx: null, volCx: null, semMedida: true, obs: 'sem medida na tabela de medidas' });
    });
    return { itens, linhasProduto: qtds.length };
  }

  async function lerPdf(arq) {
    const linhas = await linhasDoPdf(await arq.arrayBuffer());
    const cab = cabecalhoDaNota(linhas, arq.name);
    const { itens, linhasProduto } = itensDaNota(cab.secao);
    delete cab.secao;
    const obs = [], sem = itens.filter(i => i.semMedida);
    if (sem.length) obs.push('SEM MEDIDA na tabela: ' + sem.map(i => i.cod).join(', ') + ' — a cubagem da nota está incompleta');
    else if (!itens.length) obs.push('nenhum item da nota encontrado na tabela de medidas');
    if (!cab.nf) obs.push('não consegui ler o número da nota');
    itens.forEach(i => { i.cub = (i.cx == null || i.volCx == null) ? 0 : i.cx * i.volCx; });
    cab.caixas = itens.reduce((s, i) => s + (i.cx || 0), 0);
    if (!sem.length && itens.length && cab.volumes != null && cab.caixas !== cab.volumes) obs.push('caixas calculadas (' + cab.caixas + ') diferem dos volumes da nota (' + cab.volumes + ')');
    cab.itens = itens; cab.obs = obs.join('; ');
    cab.qtd = itens.reduce((s, i) => s + (i.qtd || 0), 0);
    cab.cub = itens.reduce((s, i) => s + i.cub, 0);
    return cab;
  }

  async function importar(arquivos) {
    const pdfs = [...arquivos].filter(f => /\.pdf$/i.test(f.name));
    if (!pdfs.length) { S.msg = 'Selecione arquivos PDF (DANFE).'; return render(); }
    S.lendo = true; S.msg = 'Lendo ' + pdfs.length + ' PDF(s)...'; render();
    try {
      await prepararBase();
      const falhas = [];
      for (const f of pdfs) {
        try {
          const n = await lerPdf(f);
          const k = S.notas.findIndex(x => x.nf && x.nf === n.nf);
          if (k >= 0) S.notas[k] = n; else S.notas.push(n);   // mesma nota importada de novo: substitui
        } catch (e) { falhas.push(f.name + ': ' + (e.message || e)); }
      }
      S.msg = falhas.length ? 'Não consegui ler: ' + falhas.join(' | ') : '';
    } catch (e) { S.msg = 'Erro: ' + (e.message || e); }
    S.lendo = false; render();
  }

  // ---------- tela ----------
  const COLS = [
    ['nf', 'NF', 'l'], ['emissao', 'Emissão', 'l'], ['destinatario', 'Destinatário', 'l'], ['municipio', 'Município', 'l'], ['uf', 'UF', 'l'],
    ['valor', 'Valor NF', 'r'], ['volumes', 'Volumes (nota)', 'r'], ['caixas', 'Caixas (calc.)', 'r'], ['peso', 'Peso (kg)', 'r'], ['qtd', 'Qtd (un)', 'r'], ['cub', 'Cubagem (m³)', 'r'], ['obs', 'Situação', 'l'],
  ];

  function ordenadas() {
    const o = S.ordem, v = [...S.notas];
    if (!o.col) return v;
    return v.sort((a, b) => {
      const x = a[o.col], y = b[o.col];
      const r = (typeof x === 'number' && typeof y === 'number') ? x - y : String(x == null ? '' : x).localeCompare(String(y == null ? '' : y), 'pt-BR', { numeric: true });
      return r * o.dir;
    });
  }

  function render() {
    const el = $('tab-cubagem'); if (!el) return;
    const ns = ordenadas();
    const tot = ns.reduce((t, n) => ({ v: t.v + (n.valor || 0), vol: t.vol + (n.volumes || 0), cx: t.cx + n.caixas, p: t.p + (n.peso || 0), q: t.q + n.qtd, c: t.c + n.cub }), { v: 0, vol: 0, cx: 0, p: 0, q: 0, c: 0 });
    const seta = c => S.ordem.col === c ? (S.ordem.dir > 0 ? ' ▲' : ' ▼') : '';
    const linhas = ns.map(n => {
      const aberta = S.abertas.has(n.nf), ok = !n.obs;
      let h = `<tr class="cb-lin" onclick="cubagemAlternar('${esc(n.nf)}')" style="cursor:pointer">
        <td><span style="color:#64748b">${aberta ? '▾' : '▸'}</span> <b>${esc(n.nf)}</b></td><td>${esc(n.emissao)}</td><td>${esc(n.destinatario)}</td><td>${esc(n.municipio)}</td><td>${esc(n.uf)}</td>
        <td class="r">${brl(n.valor)}</td><td class="r">${fmt(n.volumes, 0)}</td><td class="r">${fmt(n.caixas, 0)}${n.volumes != null && n.caixas === n.volumes ? ' <span style="color:#166534" title="bate com os volumes da nota">✔</span>' : ''}</td><td class="r">${fmt(n.peso, 1)}</td><td class="r">${fmt(n.qtd, 0)}</td>
        <td class="r"><b>${fmt(n.cub, 4)}</b></td>
        <td>${ok ? '<span style="color:#166534">✔ ok</span>' : '<span style="color:#b91c1c" title="' + esc(n.obs) + '">⚠ ' + esc(n.obs.length > 70 ? n.obs.slice(0, 70) + '…' : n.obs) + '</span>'}</td></tr>`;
      if (aberta) {
        h += `<tr><td colspan="${COLS.length}" style="background:#f8fafc;padding:8px 14px 12px 34px">
          <div style="font-size:.74rem;color:#64748b;margin-bottom:4px">${esc(n.emitente)} → ${esc(n.destinatario)} · ${esc(n.municipio)}/${esc(n.uf)} · arquivo ${esc(n.arquivo)}</div>
          ${n.obs ? '<div style="font-size:.76rem;color:#b91c1c;margin-bottom:6px">' + esc(n.obs) + '</div>' : ''}
          <table style="width:auto"><thead><tr><th>Item</th><th class="r">Qtd (nota)</th><th class="r">Un. por caixa</th><th class="r">Caixas (qtd ÷ un.)</th><th class="r">Caixa C × L × A (cm)</th><th class="r">Cubagem da caixa (m³)</th><th class="r">Cubagem do item (m³)</th><th>Observação</th></tr></thead><tbody>` +
          n.itens.map(i => `<tr style="${i.semMedida ? 'background:#fef2f2' : ''}"><td>${esc(i.cod)}</td><td class="r">${fmt(i.qtd, 0)}</td><td class="r">${i.semMedida ? '—' : fmt(i.un, 0)}</td><td class="r">${i.semMedida ? '—' : fmt(i.cx, 0)}</td><td class="r">${i.semMedida ? '—' : i.c + ' × ' + i.l + ' × ' + i.a}</td><td class="r">${i.semMedida ? '—' : fmt(i.volCx, 6)}</td><td class="r"><b>${i.semMedida ? '—' : fmt(i.cub, 4)}</b></td><td style="${i.semMedida ? 'color:#b91c1c' : ''}">${esc(i.obs)}</td></tr>`).join('') +
          `</tbody></table></td></tr>`;
      }
      return h;
    }).join('');
    el.innerHTML = `
      <p style="font-size:.78rem;color:#64748b;margin:0 0 14px">Importe os <b>DANFE em PDF</b> das notas: a aba lê o destinatário, o município/UF e os itens, e cruza cada item com a <b>tabela de medidas</b>
        (${S.medidas ? fmt(S.medidas.m.length, 0) + ' itens, de ' + esc(S.medidas.fonte) + ', importados em ' + esc(S.medidas.gerado) : 'Relatório pegada'}; medidas da caixa em cm arredondadas para cima, exceto ",001", que vai para baixo).
        <b>Quantidade da nota ÷ unidades por caixa = nº de caixas (arredondado para cima); cubagem = nº de caixas × cubagem da caixa.</b> A coluna "Caixas (calc.)" é conferida com os volumes da própria nota. O PDF é lido no seu navegador e não é enviado a lugar nenhum.</p>
      <div class="table-card" style="margin-bottom:14px">
        <div id="cb-drop" style="border:2px dashed #94a3b8;border-radius:10px;padding:22px;text-align:center;background:#f8fafc;cursor:pointer"
             ondragover="event.preventDefault();this.style.background='#e0f2fe'" ondragleave="this.style.background='#f8fafc'"
             ondrop="event.preventDefault();this.style.background='#f8fafc';cubagemImportar(event.dataTransfer.files)" onclick="document.getElementById('cb-arq').click()">
          📄 <b>Arraste os PDFs aqui</b> ou clique para escolher (pode selecionar vários)
          <input id="cb-arq" type="file" accept="application/pdf,.pdf" multiple style="display:none" onchange="cubagemImportar(this.files);this.value=''">
        </div>
        ${S.msg ? '<div style="margin-top:10px;font-size:.82rem;color:' + (S.lendo ? '#1a3a5c' : '#b91c1c') + '">' + esc(S.msg) + '</div>' : ''}
      </div>
      <div class="kpi-row">
        <div class="kpi"><label>Notas</label><div class="value">${fmt(ns.length, 0)}</div><div class="sub">${fmt(tot.q, 0)} unidades</div></div>
        <div class="kpi"><label>Volumes (nota)</label><div class="value">${fmt(tot.vol, 0)}</div><div class="sub">${fmt(tot.cx, 0)} caixas calculadas · ${fmt(tot.p, 1)} kg</div></div>
        <div class="kpi orange"><label>Valor das Notas</label><div class="value">${brl(tot.v)}</div><div class="sub">&nbsp;</div></div>
        <div class="kpi red"><label>Cubagem Total</label><div class="value">${fmt(tot.c, 4)} m³</div><div class="sub">caixas × cubagem da caixa</div></div>
      </div>
      <div class="table-card">
        <div style="display:flex;gap:8px;margin-bottom:10px;align-items:center;flex-wrap:wrap">
          <h3 style="margin:0;flex:1">Notas importadas <span style="font-weight:400;font-size:.76rem;color:#64748b">— clique na linha para ver os itens</span></h3>
          <button onclick="cubagemExportar()" ${ns.length ? '' : 'disabled'} style="white-space:nowrap;border:1px solid var(--border);border-radius:6px;padding:7px 14px;font-size:.8rem;font-weight:600;color:#166534;background:#f0fdf4;cursor:pointer">⬇ Exportar Excel</button>
          <button onclick="cubagemLimpar()" ${ns.length ? '' : 'disabled'} style="white-space:nowrap;border:1px solid var(--border);border-radius:6px;padding:7px 14px;font-size:.8rem;font-weight:600;color:#475569;background:#fff;cursor:pointer">Limpar</button>
        </div>
        ${ns.length ? `<div style="overflow-x:auto"><table><thead><tr>${COLS.map(([k, t, a]) => `<th class="${a === 'r' ? 'r' : ''}" style="cursor:pointer" onclick="cubagemOrdenar('${k}')">${t}${seta(k)}</th>`).join('')}</tr></thead><tbody>${linhas}</tbody>
          <tfoot><tr style="font-weight:700;background:#f1f5f9"><td colspan="5">TOTAL (${ns.length} nota${ns.length > 1 ? 's' : ''})</td><td class="r">${brl(tot.v)}</td><td class="r">${fmt(tot.vol, 0)}</td><td class="r">${fmt(tot.cx, 0)}</td><td class="r">${fmt(tot.p, 1)}</td><td class="r">${fmt(tot.q, 0)}</td><td class="r">${fmt(tot.c, 4)}</td><td></td></tr></tfoot></table></div>`
        : '<div style="padding:18px;color:#64748b;font-size:.85rem">Nenhuma nota importada ainda.</div>'}
      </div>`;
  }

  function exportar() {
    if (!S.notas.length || !window.XLSX) return;
    const wn = [['NF', 'Emissão', 'Emitente', 'Destinatário', 'Município', 'UF', 'Valor NF R$', 'Volumes (nota)', 'Peso bruto kg (nota)', 'Qtd de itens (un)', 'Caixas (calc.)', 'CUBAGEM TOTAL (m³)', 'Situação']];
    const wi = [['NF', 'Destinatário', 'Município', 'UF', 'Item', 'Qtd (nota)', 'Un. por caixa', 'Caixas (qtd ÷ un.)', 'Caixa C cm', 'Caixa L cm', 'Caixa A cm', 'Cubagem da caixa (m³)', 'Cubagem do item (m³)', 'Observação']];
    S.notas.forEach(n => {
      const r = wn.length + 1;
      wn.push([n.nf, n.emissao, n.emitente, n.destinatario, n.municipio, n.uf, n.valor, n.volumes, n.peso,
        { t: 'n', v: n.qtd, f: `SUMIFS(Itens!$F:$F,Itens!$A:$A,A${r})` }, { t: 'n', v: n.caixas, f: `SUMIFS(Itens!$H:$H,Itens!$A:$A,A${r})` }, { t: 'n', v: n.cub, f: `SUMIFS(Itens!$M:$M,Itens!$A:$A,A${r})` }, n.obs]);
      n.itens.forEach(i => {
        const q = wi.length + 1;
        wi.push(i.semMedida ? [n.nf, n.destinatario, n.municipio, n.uf, i.cod, i.qtd, '', '', '', '', '', '', '', i.obs]
          : [n.nf, n.destinatario, n.municipio, n.uf, i.cod, i.qtd, i.un, { t: 'n', v: i.cx, f: `ROUNDUP(F${q}/G${q},0)` }, i.c, i.l, i.a, { t: 'n', v: i.volCx, f: `I${q}*J${q}*K${q}/1000000` }, { t: 'n', v: i.cub, f: `H${q}*L${q}` }, i.obs]);
      });
    });
    const wb = XLSX.utils.book_new();
    const a = XLSX.utils.aoa_to_sheet(wn), b = XLSX.utils.aoa_to_sheet(wi);
    a['!cols'] = [11, 11, 32, 36, 16, 5, 14, 10, 12, 12, 12, 14, 50].map(w => ({ wch: w }));
    b['!cols'] = [11, 36, 16, 5, 16, 10, 10, 12, 10, 10, 10, 16, 16, 40].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, a, 'Por Nota'); XLSX.utils.book_append_sheet(wb, b, 'Itens');
    XLSX.writeFile(wb, 'Conferencia_Cubagem_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  }

  window.cubagemImportar = importar;
  window.cubagemAlternar = nf => { S.abertas.has(nf) ? S.abertas.delete(nf) : S.abertas.add(nf); render(); };
  window.cubagemOrdenar = col => { S.ordem = S.ordem.col === col ? { col, dir: -S.ordem.dir } : { col, dir: 1 }; render(); };
  window.cubagemLimpar = () => { S.notas = []; S.abertas.clear(); S.msg = ''; render(); };
  window.cubagemExportar = exportar;
  window.cubagemAbrir = () => { render(); if (!S.pronto) { S.pronto = true; prepararBase().then(render).catch(e => { S.msg = 'Erro: ' + (e.message || e); render(); }); } };
  const painel = document.getElementById('tab-cubagem');
  if (painel && painel.classList.contains('active')) window.cubagemAbrir();   // pagina aberta direto em #cubagem
  window._cubagemTeste = { importar, lerPdf, prepararBase, linhasDoPdf, itensDaNota, S };   // so para teste manual no console
})();
