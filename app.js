/* App — Diretrizes para a Gestão de Projetos UNIALFA */
(function () {
  'use strict';

  const root = document.getElementById('app');

  function esc(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function artChip(code) {
    const a = ARTEFATOS[code];
    if (!a) return '';
    if (a.arquivo) {
      return `<a class="art-chip ready" href="${encodeURI(a.arquivo)}" target="_blank" rel="noopener">${esc(a.nome)}<span class="arrow">↗</span></a>`;
    }
    return `<span class="art-chip missing" title="Ainda não construído">${esc(a.nome)} · a construir</span>`;
  }

  function gateBadgeHtml(gateId) {
    const g = GATES[gateId];
    const parts = g.nome.split(' — ');
    return `<button class="gate-badge" data-gate="${g.id}" title="${esc(g.nome)}">
      <span class="gate-diamond"><svg viewBox="0 0 24 24" class="gate-icon">${GATE_ICONS[gateId] || ''}</svg></span>
      <span class="g-lbl">${esc(parts[0])}<br>${esc(parts[1] || '')}</span>
    </button>`;
  }

  const DIAGRAM_ICONS = {
    D01: '<g class="icon-svg"><line x1="24" y1="4" x2="24" y2="24"></line><path d="M14 15 L24 25 L34 15"></path><path d="M5 27 L5 43 L43 43 L43 27"></path></g>',
    D02: '<g class="icon-svg"><circle cx="24" cy="24" r="20"></circle><circle cx="24" cy="24" r="12"></circle><circle cx="24" cy="24" r="4" fill="var(--red-deep)" stroke="none"></circle></g>',
    D03: '<g class="icon-svg"><circle cx="24" cy="24" r="20"></circle><path d="M19 15 L35 24 L19 33 Z"></path></g>',
    D04: '<g class="icon-svg" transform="translate(2,3)"><circle cx="22" cy="22" r="21"></circle><line x1="22" y1="22" x2="33" y2="11"></line><circle cx="22" cy="22" r="3.2" fill="var(--red-deep)" stroke="none"></circle></g>',
    D05: '<g class="icon-svg"><rect x="6" y="20" width="36" height="22" rx="3"></rect><rect x="3" y="9" width="42" height="9" rx="3"></rect><line x1="18" y1="13.5" x2="30" y2="13.5"></line></g>',
    D06: '<g class="icon-svg-tv" transform="translate(2,4)"><path d="M22 0 L44 14 L0 14 Z"></path><line x1="4" y1="18" x2="4" y2="34"></line><line x1="14" y1="18" x2="14" y2="34"></line><line x1="30" y1="18" x2="30" y2="34"></line><line x1="40" y1="18" x2="40" y2="34"></line><line x1="-2" y1="38" x2="46" y2="38"></line></g>',
    D07: '<g class="icon-svg-tv" transform="translate(1,3)"><circle cx="16" cy="10" r="8"></circle><path d="M2 34 C2 22 9 16 16 16 C23 16 30 22 30 34"></path><circle cx="33" cy="14" r="6.5"></circle><path d="M24 34 C24 25 29 22 34 22 C39 22 44 26 45 33"></path></g>',
  };
  const GATE_ICONS = {
    gate1: '<path d="M3 4 L21 4 L13 14 L13 20 L11 20 L11 14 Z"></path>',
    gate2: '<path d="M4 12 L10 18 L20 6"></path>',
  };

  function gateCardHtml(gateId) {
    const g = GATES[gateId];
    return `<div class="gate-card" id="gate-${g.id}">
      <div class="gc-head"><span class="gc-diamond"></span><h3>${esc(g.nome)}</h3></div>
      <div class="gc-grid">
        <div class="gc-field"><span>Quando</span><b>${esc(g.quando)}</b></div>
        <div class="gc-field"><span>Quem decide</span><b>${esc(g.quemDecide)}</b></div>
        <div class="gc-field"><span>Sobre qual artefato</span><b>${esc(g.sobreArtefato)}</b></div>
        <div class="gc-field"><span>Efeito</span><b>${esc(g.efeito)}</b></div>
      </div>
    </div>`;
  }

  function tag(d) {
    return d.id.replace('D', '');
  }

  // ---------------- Sidebar ----------------
  function renderSidebar(activeId) {
    const seqItems = SEQUENCIAIS.map(d => `
      <a class="nav-item ${activeId === d.id ? 'active' : ''}" href="#/${d.id}">
        <span class="nav-icon"><svg viewBox="0 0 48 48">${DIAGRAM_ICONS[d.id] || ''}</svg></span><span class="lbl">${d.id} ${esc(d.titulo)}</span>
      </a>`).join('');
    const tvItems = TRANSVERSAIS.map(d => `
      <a class="nav-item ${activeId === d.id ? 'active' : ''}" href="#/${d.id}">
        <span class="nav-icon"><svg viewBox="0 0 48 48">${DIAGRAM_ICONS[d.id] || ''}</svg></span><span class="lbl">${d.id} ${esc(d.titulo)}</span><span class="pill">Transv.</span>
      </a>`).join('');

    return `
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-brand">
          <div class="logo">SGP<small>SISTEMA DE GESTÃO DE PROJETOS</small></div>
          <p>Diretrizes, Estratégias e Práticas</p>
        </div>
        <nav class="sidebar-nav">
          <a class="nav-item ${activeId === 'home' ? 'active' : ''}" href="#/">
            <span class="tag">◎</span><span class="lbl">Visão geral</span>
          </a>
          <div class="sidebar-section-label">Fluxo sequencial (D01–D05)</div>
          ${seqItems}
          <div class="sidebar-section-label">Transversais</div>
          ${tvItems}
          <div class="sidebar-section-label">Ferramentas</div>
          <a class="nav-item" href="14 - validador-projetos.html" target="_blank" rel="noopener">
            <span class="tag">◈</span><span class="lbl">Validador de Projetos</span>
          </a>
          <a class="nav-item" href="15 - central-ajuda.html?v=78" target="_blank" rel="noopener">
            <span class="tag">?</span><span class="lbl">Central de Ajuda</span>
          </a>
          <a class="nav-item" href="apresentacao.html">
            <span class="tag">★</span><span class="lbl">Sobre este projeto</span>
          </a>
          <div class="sidebar-section-label">Gestão</div>
          <a class="nav-item" href="20 - meu-painel.html" target="_blank" rel="noopener">
            <span class="tag">◐</span><span class="lbl">Meu Painel</span>
          </a>
          <a class="nav-item" href="21 - gestao-atividades.html" target="_blank" rel="noopener">
            <span class="tag">☰</span><span class="lbl">Gestão de Atividades</span>
          </a>
          <a class="nav-item" href="16 - painel-executivo.html" target="_blank" rel="noopener">
            <span class="tag">▤</span><span class="lbl">Painel Executivo</span>
          </a>
          <a class="nav-item" href="19 - painel-prazos.html" target="_blank" rel="noopener">
            <span class="tag">⏱</span><span class="lbl">Painel de Prazos</span>
          </a>
          <a class="nav-item" href="17 - relatorio-resultados.html" target="_blank" rel="noopener">
            <span class="tag">◆</span><span class="lbl">Relatório de Resultados</span>
          </a>
          <a class="nav-item" href="13 - administracao-usuarios.html" target="_blank" rel="noopener">
            <span class="tag">⚙</span><span class="lbl">Administração</span>
          </a>
        </nav>
        <div class="sidebar-foot">12 artefatos institucionais · 2 gates de autorização</div>
      </aside>
      <div class="scrim" id="scrim"></div>
      <div class="mobile-bar">
        <button id="menuBtn" aria-label="Abrir menu">☰</button>
        <span class="mb-title">SGP · DIRETRIZES</span>
      </div>`;
  }

  // ---------------- Home / Map ----------------
  function renderHome() {
    const nodeColsHtml = SEQUENCIAIS.map((d, i) => {
      const spacer = (i === 0 || i === 1) ? ' style="margin-right:66px"' : '';
      return `<a class="node-col" href="#/${d.id}"${spacer}>
        <span class="icon-circle"><svg viewBox="0 0 48 48">${DIAGRAM_ICONS[d.id] || ''}</svg></span>
        <span class="tag">${d.id}</span>
        <h4>${esc(d.titulo)}</h4>
        <p>${esc(d.resumo)}</p>
      </a>`;
    }).join('');

    const tvCardHtml = (d, pos) => {
      const chips = d.grupos[0].estrategias.map(e => `<span class="tv-chip">${esc(e.id)}</span>`).join('');
      const icon = `<span class="icon-circle"><svg viewBox="0 0 48 48">${DIAGRAM_ICONS[d.id] || ''}</svg></span>`;
      const body = `
        <span class="tag">${d.id} · Transversal</span>
        <h3>${esc(d.titulo)}</h3>
        <p>${esc(d.resumo)}</p>
        <span class="tv-chips">${chips}</span>`;
      return `<a class="tv-card ${pos}" href="#/${d.id}">${pos === 'top' ? icon + body : body + icon}</a>`;
    };

    const d06 = TRANSVERSAIS.find(d => d.id === 'D06');
    const d07 = TRANSVERSAIS.find(d => d.id === 'D07');

    const mNode = d => `<a class="dm-node" href="#/${d.id}">
      <span class="icon-circle"><svg viewBox="0 0 48 48">${DIAGRAM_ICONS[d.id] || ''}</svg></span>
      <span class="dm-body"><span class="tag">${d.id}</span><h4>${esc(d.titulo)}</h4><p>${esc(d.resumo)}</p></span>
    </a>`;
    const mGate = gateId => {
      const g = GATES[gateId];
      const parts = g.nome.split(' — ');
      return `<button class="dm-gate" data-gate="${g.id}">
        <span class="gate-diamond"><svg viewBox="0 0 24 24" class="gate-icon">${GATE_ICONS[gateId] || ''}</svg></span>
        <span class="dm-body"><b>${esc(parts[0])} — ${esc(parts[1] || '')}</b></span>
      </button>`;
    };
    const mTv = d => {
      const chips = d.grupos[0].estrategias.map(e => `<span class="tv-chip">${esc(e.id)}</span>`).join('');
      return `<a class="dm-node dm-tv" href="#/${d.id}">
        <span class="icon-circle"><svg viewBox="0 0 48 48">${DIAGRAM_ICONS[d.id] || ''}</svg></span>
        <span class="dm-body"><span class="tag">${d.id} · Transversal</span><h4>${esc(d.titulo)}</h4><p>${esc(d.resumo)}</p><span class="tv-chips">${chips}</span></span>
      </a>`;
    };

    return `
      <div class="hero">
        <h1>Diretrizes para a Gestão de Projetos</h1>
        <p>7 diretrizes (D01–D07), 61 estratégias e 12 artefatos institucionais que orientam o ciclo de vida de um projeto na instituição — da demanda ao encerramento, com governança e liderança presentes o tempo todo. Clique em qualquer etapa do mapa para ver o detalhe.</p>
      </div>

      <div class="comm-band">Comunicação <span>— atravessa todas as diretrizes, do D01 ao D07</span></div>

      <div class="map-wrap">
        <div class="diagram-desktop">
          <div class="diagram-scroll" id="diagramScroll">
            <div class="diagram-stage" id="diagramStage">
              <svg class="arc-svg" id="arcSvg" preserveAspectRatio="none" aria-hidden="true">
                <path class="arc"></path><path class="arc"></path><path class="arc"></path><path class="arc"></path>
                <g class="chevron" id="chevronGroup"></g>
              </svg>
              ${d06 ? tvCardHtml(d06, 'top') : ''}
              <div class="node-row" id="nodeRow">
                ${gateBadgeHtml('gate1').replace('class="gate-badge"', 'class="gate-badge" style="left:228px"')}
                ${gateBadgeHtml('gate2').replace('class="gate-badge"', 'class="gate-badge" style="left:518px"')}
                ${nodeColsHtml}
              </div>
              ${d07 ? tvCardHtml(d07, 'bottom') : ''}
            </div>
          </div>
        </div>
        <div class="diagram-mobile">
          ${d06 ? mTv(d06) : ''}
          ${mNode(SEQUENCIAIS[0])}
          ${mGate('gate1')}
          ${mNode(SEQUENCIAIS[1])}
          ${mGate('gate2')}
          ${SEQUENCIAIS.slice(2).map(mNode).join('')}
          ${d07 ? mTv(d07) : ''}
        </div>
        <div class="legend">
          <span><i style="background:var(--red)"></i> Diretriz sequencial (D01–D05)</span>
          <span><i style="background:var(--ink)"></i> Diretriz transversal (D06–D07)</span>
          <span><i style="background:var(--gold);border-radius:3px;transform:rotate(45deg)"></i> Gate de autorização institucional</span>
        </div>
      </div>

      <div class="home-grid">
        <div class="info-card">
          <h3>Os 2 Gates de autorização</h3>
          <ul>
            <li><b>Gate 1 — Triagem</b> (em D01.4/D01.5): o Gestor Responsável aprova ou rejeita a entrada da demanda no portfólio.</li>
            <li><b>Gate 2 — Pactuação</b> (fim do D02): o mais crítico — o Dono do Negócio pactua o Relatório de Entregas e Benefícios com a Alta Gestão, autorizando o início da execução.</li>
          </ul>
        </div>
        <div class="info-card">
          <h3>Artefatos sempre acessíveis</h3>
          <ul>
            <li><b>Ata de Reunião</b> (SGP-07) — qualquer interação formal, do D01 ao D07.</li>
            <li><b>SMP</b> (SGP-06) — solicitação de mudança, a qualquer momento da execução/controle.</li>
            <li><b>Relatório de Situação de Projetos</b> (SGP-11) — cobre também o Reporte de Resultados, recorrente, ligado à governança (D06.5–D06.7).</li>
          </ul>
          <p style="margin-top:10px;font-size:11px;color:var(--muted-2)">Esses três não pertencem à esteira sequencial D01→D05 — não bloqueiam nem são bloqueados por ela.</p>
        </div>
      </div>`;
  }

  // ---------------- Detail ----------------
  function estrategiaHtml(e) {
    const nomes = (e.artefatos || []).map(code => ARTEFATOS[code] && ARTEFATOS[code].nome).filter(Boolean);
    const formsLabel = nomes.length ? `<div class="acc-forms">Formulários (${nomes.map(esc).join(', ')})</div>` : '';
    return `<div class="acc-item" data-acc>
      <div class="acc-head">
        <button class="acc-trigger" data-toggle>
          <span class="acc-id">${esc(e.id)}</span>
          <span class="acc-title" title="${esc(e.titulo)}">${esc(e.titulo)}</span>
          <span class="acc-chev">▾</span>
        </button>
        ${formsLabel}
      </div>
      <div class="acc-body">
        <div class="acc-body-in">
          <ul>${e.pontos.map(p => `<li>${esc(p)}</li>`).join('')}</ul>
        </div>
      </div>
    </div>`;
  }

  function renderDetail(id) {
    const d = getDiretriz(id);
    if (!d) return renderHome();

    const entregaLi = d.entrega.map(x => `<li>${esc(x)}</li>`).join('');
    const pessoasLi = d.pessoas.map(x => `<li>${esc(x)}</li>`).join('');

    const gateBefore = Object.values(GATES).find(g => g.diretriz === d.id && false); // reserved

    const artCodes = [];
    d.grupos.forEach(g => g.estrategias.forEach(e => (e.artefatos || []).forEach(c => {
      if (!artCodes.includes(c)) artCodes.push(c);
    })));
    const formsSummaryHtml = artCodes.length ? `
      <div class="forms-summary">
        <h4>Formulários utilizados nesta diretriz</h4>
        <div class="art-row">${artCodes.map(artChip).join('')}</div>
      </div>` : '';

    let strategiesHtml = '';

    d.grupos.forEach(grp => {
      if (grp.titulo) {
        strategiesHtml += `<div class="strat-group-title">${esc(grp.id)} — ${esc(grp.titulo)}</div>`;
      } else {
        strategiesHtml += `<div class="strategies-title">Estratégias recomendadas</div>`;
      }
      strategiesHtml += `<div class="accordion">${grp.estrategias.map(estrategiaHtml).join('')}</div>`;

      // insert gate card right after the strategy it follows
      Object.values(GATES).forEach(g => {
        if (grp.estrategias.some(e => e.id === g.apos)) {
          strategiesHtml += gateCardHtml(g.id);
        }
      });
    });

    return `
      <div class="crumb"><a href="#/">Visão geral</a> <span>/</span> <span>${esc(d.id)} — ${esc(d.titulo)}</span></div>

      <div class="d-head">
        <div class="d-head-band">
          <div class="d-num">${tag(d)}</div>
          <div class="d-title">
            <h1>${esc(d.id)} — ${esc(d.titulo)}</h1>
            <div class="d-badge">${d.tipo === 'transversal' ? 'Diretriz transversal · presente do D01 ao D05' : `Diretriz sequencial · etapa ${d.ordem} de 5`}</div>
          </div>
          <div class="badge-pill">${d.tipo === 'transversal' ? 'Transversal' : 'Sequencial'}</div>
        </div>
        <div class="d-head-body">
          <div><h4>O que é?</h4><p>${esc(d.oQueE)}</p></div>
          <div><h4>Por que é importante?</h4><p>${esc(d.porQue)}</p></div>
        </div>
        <div class="ficha">
          <div><h4>O que a diretriz entrega?</h4><ul>${entregaLi}</ul></div>
          <div><h4>Quem está envolvido?</h4><ul>${pessoasLi}</ul></div>
        </div>
        ${formsSummaryHtml}
      </div>

      ${strategiesHtml}
    `;
  }

  // ---------------- Router ----------------
  function currentId() {
    const h = location.hash.replace('#/', '').trim();
    return h || 'home';
  }

  // ---------------- Diagram layout (arcs + responsive scale) ----------------
  function fitDiagramScale() {
    const scrollEl = document.getElementById('diagramScroll');
    const stage = document.getElementById('diagramStage');
    if (!scrollEl || !stage || stage.offsetParent === null) return; // oculto (versão mobile empilhada ativa)
    const naturalW = 1228;
    const scale = Math.min(1, scrollEl.clientWidth / naturalW);
    stage.style.transform = `scale(${scale})`;
    scrollEl.style.height = (stage.offsetHeight * scale) + 'px';
  }

  function layoutDiagram() {
    const stage = document.getElementById('diagramStage');
    if (!stage || stage.offsetParent === null) return; // oculto (versão mobile empilhada ativa)
    const svg = document.getElementById('arcSvg');
    const chevronGroup = document.getElementById('chevronGroup');
    const nodeRow = document.getElementById('nodeRow');
    const cols = nodeRow ? nodeRow.querySelectorAll('.node-col') : [];
    const d06Icon = stage.querySelector('.tv-card.top .icon-circle');
    const d07Icon = stage.querySelector('.tv-card.bottom .icon-circle');
    if (!cols.length || !d06Icon || !d07Icon) return;

    stage.style.transform = 'none'; // measure natural (unscaled) layout first

    const sr = stage.getBoundingClientRect();
    const rel = el => {
      const r = el.getBoundingClientRect();
      return { cx: r.left + r.width / 2 - sr.left, cy: r.top + r.height / 2 - sr.top, l: r.left - sr.left, rr: r.right - sr.left, t: r.top - sr.top, b: r.bottom - sr.top };
    };

    const d06 = rel(d06Icon), d07 = rel(d07Icon);
    const col1El = cols[0], col5El = cols[cols.length - 1];
    const c1 = rel(col1El.querySelector('.icon-circle'));
    const c5 = rel(col5El.querySelector('.icon-circle'));
    const c1Col = rel(col1El); // whole column box (icon + tag + title + paragraph)
    const c5Col = rel(col5El);
    const R = 24;

    svg.setAttribute('viewBox', `0 0 ${stage.scrollWidth} ${stage.scrollHeight}`);
    const paths = [
      `M ${d06.l} ${d06.cy} L ${c1.cx + R} ${d06.cy} Q ${c1.cx} ${d06.cy} ${c1.cx} ${d06.cy + R} L ${c1.cx} ${c1.t}`,
      `M ${d06.rr} ${d06.cy} L ${c5.cx - R} ${d06.cy} Q ${c5.cx} ${d06.cy} ${c5.cx} ${d06.cy + R} L ${c5.cx} ${c1.t}`,
      `M ${d07.l} ${d07.cy} L ${c1.cx + R} ${d07.cy} Q ${c1.cx} ${d07.cy} ${c1.cx} ${d07.cy - R} L ${c1.cx} ${c1Col.b}`,
      `M ${d07.rr} ${d07.cy} L ${c5.cx - R} ${d07.cy} Q ${c5.cx} ${d07.cy} ${c5.cx} ${d07.cy - R} L ${c5.cx} ${c5Col.b}`,
    ];
    svg.querySelectorAll('.arc').forEach((p, i) => p.setAttribute('d', paths[i] || ''));

    chevronGroup.innerHTML = '';
    for (let i = 0; i < cols.length - 1; i++) {
      const a = cols[i].querySelector('.icon-circle');
      const b = cols[i + 1].querySelector('.icon-circle');
      const ra = rel(a), rb = rel(b);
      const midX = (ra.rr + rb.l) / 2, cy = ra.cy;
      const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      p.setAttribute('d', `M ${midX - 8} ${cy - 12} L ${midX + 8} ${cy} L ${midX - 8} ${cy + 12} L ${midX - 8} ${cy + 3} L ${midX - 20} ${cy + 3} L ${midX - 20} ${cy - 3} L ${midX - 8} ${cy - 3} Z`);
      chevronGroup.appendChild(p);
    }

    fitDiagramScale();
  }

  function render() {
    const id = currentId();
    const body = id === 'home' ? renderHome() : renderDetail(id);
    root.innerHTML = `
      ${renderSidebar(id)}
      <main class="main"><div class="main-inner">${body}</div></main>
    `;
    wireEvents();
    window.scrollTo(0, 0);
    if (id === 'home') layoutDiagram();
  }

  function wireEvents() {
    root.querySelectorAll('[data-toggle]').forEach(btn => {
      btn.addEventListener('click', () => {
        btn.closest('.acc-item').classList.toggle('open');
      });
    });
    root.querySelectorAll('[data-gate]').forEach(btn => {
      btn.addEventListener('click', () => {
        const g = GATES[btn.getAttribute('data-gate')];
        location.hash = `#/${g.diretriz}`;
        setTimeout(() => {
          const el = document.getElementById(`gate-${g.id}`);
          if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 60);
      });
    });
    const menuBtn = document.getElementById('menuBtn');
    const sidebar = document.getElementById('sidebar');
    const scrim = document.getElementById('scrim');
    if (menuBtn) {
      menuBtn.addEventListener('click', () => {
        sidebar.classList.add('open');
        scrim.classList.add('show');
      });
      scrim.addEventListener('click', () => {
        sidebar.classList.remove('open');
        scrim.classList.remove('show');
      });
      root.querySelectorAll('.nav-item').forEach(a => a.addEventListener('click', () => {
        sidebar.classList.remove('open');
        scrim.classList.remove('show');
      }));
    }
  }

  window.addEventListener('hashchange', render);
  window.addEventListener('DOMContentLoaded', render);
  window.addEventListener('resize', () => { if (currentId() === 'home') fitDiagramScale(); });
})();
