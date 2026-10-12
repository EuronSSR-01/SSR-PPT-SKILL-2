(() => {
  'use strict';
  const deck = document.getElementById('ssr-deck');
  if (!deck) throw new Error('SSR player requires #ssr-deck.');
  const player = deck.closest('.ssr-player');
  const stage = deck.closest('.ssr-stage');
  const slides = Array.from(deck.children).filter(el => el.classList.contains('ssr-slide'));
  if (!player || !stage || !slides.length) {
    throw new Error('SSR player requires .ssr-player > .ssr-stage > #ssr-deck with at least one .ssr-slide.');
  }
  const W = Number(deck.dataset.width || 1280);
  const H = Number(deck.dataset.height || 720);
  if (!(W > 0 && H > 0 && Number.isFinite(W) && Number.isFinite(H))) {
    throw new Error('SSR canvas dimensions must be positive finite numbers.');
  }
  if (!deck.dataset.transition) deck.dataset.transition = 'fade';
  player.style.setProperty('--ssr-width', `${W}px`);
  player.style.setProperty('--ssr-height', `${H}px`);
  const pageSize = document.createElement('style');
  pageSize.textContent = `@page { size: ${W / 96}in ${H / 96}in; margin: 0; }`;
  document.head.appendChild(pageSize);
  const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)');

  // [data-step] groups per slide, ordered by number; an empty value means document order.
  const steps = slides.map(slide => {
    const items = Array.from(slide.querySelectorAll('[data-step]'), (el, i) => {
      const n = parseFloat(el.dataset.step);
      return { el, key: Number.isFinite(n) ? n : i + 1 };
    });
    const keys = [...new Set(items.map(item => item.key))].sort((a, b) => a - b);
    return keys.map(key => items.filter(item => item.key === key).map(item => item.el));
  });
  const totalSteps = steps.reduce((sum, groups) => sum + groups.length, 0);

  slides.forEach(slide => {
    slide.querySelectorAll('[data-stagger]').forEach(group => {
      const gap = Number(group.dataset.stagger) || 80;
      group.querySelectorAll('[data-anim]').forEach((el, i) => {
        if (!el.style.getPropertyValue('--d')) el.style.setProperty('--d', `${i * gap}ms`);
      });
    });
    slide.querySelectorAll('[data-anim="draw"]').forEach(el => {
      const shapes = el instanceof SVGGeometryElement ? [el]
        : el.querySelectorAll('path, line, polyline, polygon, circle, ellipse, rect');
      shapes.forEach(shape => {
        shape.setAttribute('pathLength', '1');
        shape.classList.add('ssr-draw');
      });
    });
  });

  // UI is generated here so decks only need the stage markup.
  player.querySelectorAll('.ssr-controls').forEach(el => el.remove());
  const ICONS = {
    prev: '<path d="M15 6l-6 6 6 6"/>',
    next: '<path d="M9 6l6 6-6 6"/>',
    overview: '<rect x="4" y="4" width="7" height="7" rx="1"/><rect x="13" y="4" width="7" height="7" rx="1"/>'
      + '<rect x="4" y="13" width="7" height="7" rx="1"/><rect x="13" y="13" width="7" height="7" rx="1"/>',
    notes: '<path d="M6 4h9l3 3v13H6z"/><path d="M9 11h6M9 15h6"/>',
    full: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
    exitFull: '<path d="M9 4v5H4M15 4v5h5M9 20v-5H4M15 20v-5h5"/>',
  };
  const icon = name => `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[name]}</svg>`;
  const make = (tag, className, attrs = {}) => {
    const el = document.createElement(tag);
    el.className = className;
    Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
    return el;
  };
  const controls = make('nav', 'ssr-controls ssr-ui', { 'aria-label': '演示控制' });
  controls.innerHTML = `
    <button type="button" data-act="prev" aria-label="上一步" title="上一步（←）">${icon('prev')}</button>
    <output class="ssr-count" aria-live="polite"></output>
    <button type="button" data-act="next" aria-label="下一步" title="下一步（→ / 空格）">${icon('next')}</button>
    <span class="ssr-sep" aria-hidden="true"></span>
    <button type="button" data-act="overview" aria-label="概览" aria-pressed="false" title="概览（O）">${icon('overview')}</button>
    <button type="button" data-act="notes" aria-label="备注" aria-pressed="false" title="备注（N）">${icon('notes')}</button>
    <button type="button" data-act="fullscreen" aria-label="全屏" title="全屏（F）">${icon('full')}</button>`;
  const progress = make('div', 'ssr-progress ssr-ui', { 'aria-hidden': 'true' });
  progress.appendChild(document.createElement('i'));
  const notes = make('aside', 'ssr-notes-panel ssr-ui', { 'aria-label': '演讲备注' });
  notes.hidden = true;
  const blackout = make('div', 'ssr-blackout ssr-ui');
  blackout.hidden = true;
  const status = make('div', 'ssr-status ssr-ui', { role: 'status' });
  player.append(progress, controls, notes, blackout, status);
  const button = act => controls.querySelector(`[data-act="${act}"]`);
  const count = controls.querySelector('.ssr-count');

  let current = 0;
  let revealed = 0;
  let overview = false;
  let leaveTimer = 0;
  let statusTimer = 0;
  const hits = [];

  function setRevealed(k) {
    revealed = Math.max(0, Math.min(steps[current].length, k));
    steps[current].forEach((group, g) => group.forEach(el => el.classList.toggle('ssr-on', g < revealed)));
  }

  function go(index, { reveal = 0, animate = true } = {}) {
    index = Math.max(0, Math.min(slides.length - 1, index));
    const from = current;
    current = index;
    if (overview) {
      if (hits[index]) hits[index].focus({ preventScroll: true });
    } else {
      clearTimeout(leaveTimer);
      const transition = animate && from !== index && deck.dataset.transition !== 'none' && !reducedMotion.matches;
      slides.forEach((slide, i) => {
        const active = i === index;
        const leaving = transition && i === from;
        slide.hidden = !(active || leaving);
        slide.inert = !active;
        slide.setAttribute('aria-hidden', String(!active));
        slide.classList.toggle('ssr-active', active);
        slide.classList.toggle('ssr-leaving', leaving);
        slide.classList.remove('ssr-enter');
      });
      if (transition) {
        deck.style.setProperty('--ssr-dir', index > from ? '1' : '-1');
        void slides[index].offsetWidth; // restart the entrance animation
        slides[index].classList.add('ssr-enter');
        leaveTimer = setTimeout(() => {
          slides[from].hidden = from !== current;
          slides[from].classList.remove('ssr-leaving');
        }, 700);
      }
    }
    slides.forEach((slide, i) => slide.classList.toggle('ssr-current', i === index));
    setRevealed(reveal === 'all' ? Infinity : reveal);
    try {
      history.replaceState(null, '', `#${index + 1}`);
    } catch (error) {
      // Some browsers refuse history updates for file: URLs; the deck still works.
    }
    update();
  }

  function next() {
    if (overview) go(current + 1);
    else if (revealed < steps[current].length) { setRevealed(revealed + 1); update(); }
    else if (current < slides.length - 1) go(current + 1);
  }

  function prev() {
    if (overview) go(current - 1);
    else if (revealed > 0) { setRevealed(revealed - 1); update(); }
    else if (current > 0) go(current - 1, { reveal: 'all' });
  }

  function update() {
    button('prev').disabled = current === 0 && (overview || revealed === 0);
    button('next').disabled = current === slides.length - 1 && (overview || revealed === steps[current].length);
    count.textContent = `${current + 1} / ${slides.length}`;
    player.style.setProperty('--ssr-progress', String((current + 1) / slides.length));
    if (!notes.hidden) renderNotes();
  }

  function renderNotes() {
    const source = slides[current].querySelector('.ssr-notes');
    const head = make('div', 'ssr-notes-head');
    head.textContent = `第 ${current + 1} 页备注`;
    const body = document.createElement('div');
    if (source) body.innerHTML = source.innerHTML;
    else body.textContent = '本页没有备注。';
    notes.replaceChildren(head, body);
  }

  function toggleNotes(force) {
    notes.hidden = !(force ?? notes.hidden);
    button('notes').setAttribute('aria-pressed', String(!notes.hidden));
    if (!notes.hidden) renderNotes();
  }

  function layoutOverview() {
    const n = slides.length;
    const gap = W * 0.06;
    let best = { scale: 0 };
    for (let cols = 1; cols <= n; cols++) {
      const rows = Math.ceil(n / cols);
      const dw = cols * W + (cols - 1) * gap;
      const dh = rows * H + (rows - 1) * gap;
      const scale = Math.min(stage.clientWidth * 0.92 / dw, stage.clientHeight * 0.84 / dh);
      if (scale > best.scale) best = { scale, cols, dw, dh };
    }
    player.style.setProperty('--ssr-deck-w', `${best.dw}px`);
    player.style.setProperty('--ssr-deck-h', `${best.dh}px`);
    player.style.setProperty('--ssr-scale', String(best.scale));
    slides.forEach((slide, i) => {
      const offset = `translate(${(i % best.cols) * (W + gap)}px, ${Math.floor(i / best.cols) * (H + gap)}px)`;
      slide.style.transform = offset;
      hits[i].style.transform = offset;
    });
  }

  function toggleOverview(force) {
    const on = force ?? !overview;
    if (on === overview) return;
    overview = on;
    player.classList.toggle('ssr-is-overview', overview);
    button('overview').setAttribute('aria-pressed', String(overview));
    clearTimeout(leaveTimer);
    if (overview) {
      slides.forEach((slide, i) => {
        slide.hidden = false;
        slide.inert = true;
        slide.classList.remove('ssr-active', 'ssr-leaving', 'ssr-enter');
        const hit = make('button', 'ssr-ov-hit', { type: 'button', 'aria-label': `转到第 ${i + 1} 页` });
        hit.addEventListener('click', () => {
          current = i;
          toggleOverview(false);
        });
        hit.addEventListener('focus', () => { if (overview && current !== i) go(i); }); // highlight follows Tab focus
        deck.appendChild(hit);
        hits[i] = hit;
      });
      layoutOverview();
      hits[current].focus({ preventScroll: true });
      update();
    } else {
      hits.splice(0).forEach(hit => hit.remove());
      slides.forEach(slide => { slide.style.transform = ''; });
      player.style.removeProperty('--ssr-deck-w');
      player.style.removeProperty('--ssr-deck-h');
      resize();
      go(current, { animate: false });
    }
  }

  async function toggleFullscreen() {
    if (!document.fullscreenEnabled) return;
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await player.requestFullscreen();
    } catch (error) {
      flash('当前浏览器无法进入全屏，可使用浏览器自带的全屏或演示模式。');
    }
  }

  function flash(message) {
    status.textContent = message;
    clearTimeout(statusTimer);
    statusTimer = setTimeout(() => { status.textContent = ''; }, 4000);
  }

  function resize() {
    if (overview) layoutOverview();
    else player.style.setProperty('--ssr-scale', String(Math.min(stage.clientWidth / W, stage.clientHeight / H)));
  }

  // Controls appear on pointer activity and fade out when idle.
  let idleTimer = 0;
  function wake() {
    player.classList.remove('ssr-idle');
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => {
      if (!overview && !controls.querySelector(':focus-visible')) player.classList.add('ssr-idle');
    }, 2500);
  }
  player.addEventListener('pointermove', wake);
  player.addEventListener('pointerdown', wake);
  controls.addEventListener('focusin', wake);

  controls.addEventListener('click', event => {
    const act = event.target.closest('button')?.dataset.act;
    if (act === 'prev') prev();
    else if (act === 'next') next();
    else if (act === 'overview') toggleOverview();
    else if (act === 'notes') toggleNotes();
    else if (act === 'fullscreen') toggleFullscreen();
  });

  document.addEventListener('keydown', event => {
    if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey) return;
    const target = event.target instanceof Element ? event.target : document.body;
    if (target.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"])')) return;
    const onButton = Boolean(target.closest('button, a[href], summary'));
    const key = event.key;
    let handled = true;
    if (['ArrowRight', 'ArrowDown', 'PageDown'].includes(key) || (key === ' ' && !event.shiftKey && !onButton)) next();
    else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(key) || (key === ' ' && event.shiftKey && !onButton)) prev();
    else if (key === 'Home') go(0);
    else if (key === 'End') go(slides.length - 1, { reveal: 'all' });
    else if ((key === 'Escape' || (key === 'Enter' && !onButton)) && overview) toggleOverview(false);
    else if (key === 'Escape' && !notes.hidden) toggleNotes(false);
    else if (key === 'o' || key === 'O') toggleOverview();
    else if (key === 'n' || key === 'N') toggleNotes();
    else if (key === 'f' || key === 'F') toggleFullscreen();
    else if (key === 'b' || key === 'B' || key === '.') blackout.hidden = !blackout.hidden;
    else handled = false;
    if (handled) event.preventDefault();
  });

  let touchStart = null;
  stage.addEventListener('pointerdown', event => {
    touchStart = event.pointerType === 'mouse' ? null : { x: event.clientX, y: event.clientY };
  });
  stage.addEventListener('pointerup', event => {
    if (!touchStart) return;
    const dx = event.clientX - touchStart.x;
    const dy = event.clientY - touchStart.y;
    touchStart = null;
    if (Math.abs(dx) > 48 && Math.abs(dx) > Math.abs(dy) * 1.5) (dx < 0 ? next : prev)();
  });
  stage.addEventListener('pointercancel', () => { touchStart = null; });

  const fullscreenButton = button('fullscreen');
  fullscreenButton.disabled = !document.fullscreenEnabled;
  document.addEventListener('fullscreenchange', () => {
    const on = Boolean(document.fullscreenElement);
    fullscreenButton.innerHTML = icon(on ? 'exitFull' : 'full');
    fullscreenButton.setAttribute('aria-label', on ? '退出全屏' : '全屏');
    fullscreenButton.title = on ? '退出全屏（F）' : '全屏（F）';
    resize();
  });

  const slideFromHash = () => {
    const match = /^#(\d+)$/.exec(location.hash);
    return match ? Number(match[1]) - 1 : null;
  };
  window.addEventListener('hashchange', () => {
    const index = slideFromHash();
    if (index !== null && index !== current) go(index);
  });

  window.SSRPlayer = Object.freeze({
    get index() { return current; },
    get count() { return slides.length; },
    get step() { return revealed; },
    get totalSteps() { return totalSteps; },
    go: (index, options) => go(index, { animate: false, ...options }),
    next,
    prev,
    overview: force => toggleOverview(force),
    notes: force => toggleNotes(force),
  });

  new ResizeObserver(resize).observe(stage);
  resize();
  go(slideFromHash() ?? 0, { animate: false });
  wake();
})();
