(() => {
  'use strict';

  const CONFIG = {
    // Replace FILE_ID with the "fileId" returned by your Apps Script Web App.
    DATA_URLS: [
      'https://drive.usercontent.google.com/download?id=FILE_ID&export=download'
      // Most reliable CORS option (needs a Drive-API-restricted key):
      // 'https://www.googleapis.com/drive/v3/files/FILE_ID?alt=media&key=YOUR_API_KEY'
    ],
    TTL_MS: 10 * 60 * 1000,        // match CONFIG.REFRESH_MINUTES in Code.gs
    CACHE_KEY: 'newsline:v2',
    PAGE: 24                       // stories per "Show more" step
  };

  const $ = (id) => document.getElementById(id);
  const container = $('news-container'), chips = $('chips'), cats = $('cats'), search = $('search');
  const fresh = $('fresh'), more = $('more'), updated = $('updated');
  const state = { data: { articles: [] }, source: 'All', category: 'All', query: '', shown: CONFIG.PAGE };
  let pending = null, lastCheck = Date.now();

  /* ---------- data ---------- */
  const readCache = () => { try { return JSON.parse(sessionStorage.getItem(CONFIG.CACHE_KEY)); } catch { return null; } };
  const writeCache = (data) => { try { sessionStorage.setItem(CONFIG.CACHE_KEY, JSON.stringify({ t: Date.now(), data })); } catch {} };

  async function fetchData() {
    // force-cache never revalidates, so the URL changes every TTL window: same window = instant
    // HTTP-cache hit, new window = fresh download.
    const bucket = Math.floor(Date.now() / CONFIG.TTL_MS);
    let lastError;
    for (const base of CONFIG.DATA_URLS) {
      try {
        const res = await fetch(base + (base.includes('?') ? '&' : '?') + 'v=' + bucket, { cache: 'force-cache' });
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return await res.json();
      } catch (err) { lastError = err; }
    }
    throw lastError;
  }

  async function getData() {
    const cached = readCache();
    if (cached && Date.now() - cached.t < CONFIG.TTL_MS) return cached.data;
    try { const data = await fetchData(); writeCache(data); return data; }
    catch (err) { if (cached) return cached.data; throw err; }
  }

  /* ---------- helpers (textContent only: feed text can never inject HTML) ---------- */
  const el = (tag, cls, text) => { const n = document.createElement(tag); if (cls) n.className = cls; if (text) n.textContent = text; return n; };
  const safeUrl = (u) => (/^https?:\/\//i.test(u || '') ? u.replace(/^http:\/\//i, 'https://') : '');
  const hue = (s) => [...s].reduce((h, c) => (h * 31 + c.charCodeAt(0)) % 360, 7);
  const rtf = new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' });
  function ago(s) {
    const t = Date.parse(s);
    if (!t) return '';
    const m = Math.min(0, Math.round((t - Date.now()) / 60000));
    if (m > -60) return rtf.format(m, 'minute');
    const h = Math.round(m / 60);
    if (h > -24) return rtf.format(h, 'hour');
    const d = Math.round(h / 24);
    return d > -7 ? rtf.format(d, 'day') : new Date(t).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  /* ---------- rendering ---------- */
  function buildCard(a, hero) {
    const card = el('article', hero ? 'card card--hero' : 'card');
    const src = safeUrl(a.image_url);
    if (src) {
      const img = document.createElement('img');
      img.loading = 'lazy'; img.decoding = 'async';
      img.referrerPolicy = 'no-referrer';          // many publisher CDNs reject hotlinks that send a referrer
      img.width = 640; img.height = 360;           // reserves space, prevents layout shift
      img.alt = '';                                // decorative: the headline carries the meaning
      img.src = src;
      img.addEventListener('error', () => img.remove(), { once: true });
      card.append(img);
    }

    const meta = el('div', 'card__meta');
    const badge = el('span', 'card__src', a.source || '');
    badge.style.setProperty('--h', hue(a.source || ''));
    meta.append(badge);
    const when = ago(a.published_date);
    if (when) { const t = el('time', null, when); t.dateTime = new Date(a.published_date).toISOString(); meta.append(t); }

    const title = el('h2', 'card__title');
    const href = safeUrl(a.url);
    if (href) {
      const link = el('a', null, a.title);
      link.href = href; link.target = '_blank'; link.rel = 'noopener noreferrer';
      title.append(link);
    } else title.textContent = a.title;

    const foot = el('div', 'card__foot');
    foot.append(el('span', null, [a.author, a.category].filter(Boolean).join(', ')));
    if (!href && (a.content || '').length > 220) {   // own stories expand in place
      const btn = el('button', 'card__more', 'Read more');
      btn.type = 'button'; btn.setAttribute('aria-expanded', 'false');
      foot.append(btn);
    }

    const body = el('div', 'card__body');
    body.append(meta, title);
    if (a.content) body.append(el('p', 'card__text', a.content));
    body.append(foot);
    card.append(body);
    return card;
  }

  function showMessage(text, retry) {
    const box = el('div', 'msg', text);
    if (retry) {
      const btn = el('button', null, 'Try again');
      btn.type = 'button'; btn.addEventListener('click', retry);
      box.append(document.createElement('br'), btn);
    }
    container.replaceChildren(box);
    container.removeAttribute('aria-busy');
    more.hidden = true;
  }

  function filtered() {
    const q = state.query.trim().toLowerCase();
    return state.data.articles.filter((a) =>
      (state.source === 'All' || a.source === state.source) &&
      (state.category === 'All' || a.category === state.category) &&
      (!q || (a.title + ' ' + (a.content || '') + ' ' + a.source).toLowerCase().includes(q)));
  }

  function renderChips() {
    const counts = {};
    state.data.articles.forEach((a) => { counts[a.source] = (counts[a.source] || 0) + 1; });
    chips.replaceChildren(...['All', ...Object.keys(counts)].map((name) => {
      const b = el('button', 'chip', name);
      b.type = 'button'; b.dataset.source = name;
      b.setAttribute('aria-pressed', String(name === state.source));
      return b;
    }));    renderCats();
  }

  // Category chips follow the selected publisher; hidden when there is only one category.
  function renderCats() {
    const names = [...new Set(state.data.articles
      .filter((a) => state.source === 'All' || a.source === state.source)
      .map((a) => a.category).filter(Boolean))];
    if (!names.includes(state.category)) state.category = 'All';
    cats.hidden = names.length < 2;
    cats.replaceChildren(...['All', ...names].map((name) => {
      const b = el('button', 'chip', name === 'All' ? 'All categories' : name);
      b.type = 'button'; b.dataset.category = name;
      b.setAttribute('aria-pressed', String(name === state.category));
      return b;
    }));
  }

  function render() {
    const list = filtered();
    if (!list.length) return showMessage(state.data.articles.length ? 'No stories match your search.' : 'No stories yet. Check back soon.');
    const frag = document.createDocumentFragment();
    list.slice(0, state.shown).forEach((a, i) => frag.append(buildCard(a, i === 0 && !!safeUrl(a.image_url))));
    container.replaceChildren(frag);
    container.removeAttribute('aria-busy');
    more.hidden = list.length <= state.shown;
    if (state.data.generatedAt) {
      updated.textContent = 'Updated ' + new Date(state.data.generatedAt).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
    }
  }

  const reset = () => { state.shown = CONFIG.PAGE; render(); };

  /* ---------- events ---------- */
  chips.addEventListener('click', (e) => {
    const b = e.target.closest('.chip');
    if (!b) return;
    state.source = b.dataset.source; state.category = 'All';
    renderChips(); reset();
  });
  cats.addEventListener('click', (e) => {
    const b = e.target.closest('.chip');
    if (!b) return;
    state.category = b.dataset.category;
    renderCats(); reset();
  });
  let timer;
  search.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => { state.query = search.value; reset(); }, 150); });
  more.addEventListener('click', () => { state.shown += CONFIG.PAGE; render(); });
  container.addEventListener('click', (e) => {
    const btn = e.target.closest('.card__more');
    if (!btn) return;
    const open = btn.closest('.card').classList.toggle('is-open');
    btn.textContent = open ? 'Show less' : 'Read more';
    btn.setAttribute('aria-expanded', String(open));
  });

  /* ---------- live refresh: check for new stories while the tab is open ---------- */
  async function checkForNew() {
    lastCheck = Date.now();
    try {
      const data = await fetchData();
      if (data.generatedAt && data.generatedAt !== state.data.generatedAt) { pending = data; fresh.hidden = false; }
    } catch { /* stay quiet; try again next interval */ }
  }
  fresh.addEventListener('click', () => {
    state.data = pending; writeCache(pending); pending = null; fresh.hidden = true;
    renderChips(); reset(); window.scrollTo({ top: 0 });
  });
  setInterval(checkForNew, CONFIG.TTL_MS);
  document.addEventListener('visibilitychange', () => { if (!document.hidden && Date.now() - lastCheck > CONFIG.TTL_MS) checkForNew(); });

  function init() {
    getData().then((data) => {
      state.data = { ...data, articles: Array.isArray(data) ? data : data.articles || [] };
      renderChips(); render();
    }).catch(() => showMessage("Couldn't load stories. Check your connection and try again.", init));
  }
  init();
})();