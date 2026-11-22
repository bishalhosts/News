/**
 * Newsline backend: RSS feeds from authorized publishers + your own "Articles" sheet
 * -> one data.json in Drive, rebuilt on a timer. Headlines, short excerpts, images and
 * links only; readers click through to the publisher for the full story.
 */
const CONFIG = {
  SPREADSHEET_ID: '',                       // '' when the script is bound to the spreadsheet
  SHEET_NAME: 'Articles',                   // optional: your own stories
  OWN_SOURCE: 'Newsline Desk',              // label for stories from the sheet
  SOURCES_SHEET: 'Sources',                 // one row per publisher category (see getFeeds_)
  FOLDER_ID: 'PASTE_DRIVE_FOLDER_ID_HERE',
  FILE_NAME: 'data.json',
  FIELDS: ['id', 'title', 'content', 'author', 'category', 'image_url', 'published_date'],
  PER_FEED: 15,                             // newest items kept per publisher
  MAX_ARTICLES: 200,
  SUMMARY_CHARS: 240,
  REFRESH_MINUTES: 10                       // allowed values: 1, 5, 10, 15, 30
};

// Fallback list, used only when the "Sources" sheet is missing or empty.
const FEEDS = [
  { name: 'BBC News',     url: 'https://feeds.bbci.co.uk/news/rss.xml',        category: 'World' },
  { name: 'The Guardian', url: 'https://www.theguardian.com/world/rss',        category: 'World' },
  { name: 'Al Jazeera',   url: 'https://www.aljazeera.com/xml/rss/all.xml',    category: 'World' },
  { name: 'NPR',          url: 'https://feeds.npr.org/1001/rss.xml',           category: 'News' }
];

const NS = {
  media: XmlService.getNamespace('http://search.yahoo.com/mrss/'),
  content: XmlService.getNamespace('http://purl.org/rss/1.0/modules/content/'),
  dc: XmlService.getNamespace('http://purl.org/dc/elements/1.1/')
};

/** Fetches every feed + the sheet, merges, and overwrites data.json (same file ID). */
function generateJSON() {
  const existing = findFile_();
  let previous = [];
  try { previous = JSON.parse(existing.getBlob().getDataAsString()).articles || []; } catch (e) {}

  let articles = readSheet_();
  getFeeds_().forEach(feed => {
    try {
      articles = articles.concat(fetchFeed_(feed));
    } catch (err) {
      console.warn('Feed failed: ' + feed.name + ' (' + err.message + '). Keeping previous items.');
      articles = articles.concat(previous.filter(a => a.source === feed.name && (!feed.category || a.category === feed.category)));
    }
  });

  const seen = {};
  articles = articles
    .filter(a => { const k = a.url || 'sheet:' + a.id; if (seen[k]) return false; seen[k] = true; return true; })
    .sort((a, b) => (Date.parse(b.published_date) || 0) - (Date.parse(a.published_date) || 0))
    .slice(0, CONFIG.MAX_ARTICLES);

  const json = JSON.stringify({ generatedAt: new Date().toISOString(), count: articles.length, articles });
  let file = existing;
  if (file) file.setContent(json);
  else file = DriveApp.getFolderById(CONFIG.FOLDER_ID).createFile(CONFIG.FILE_NAME, json, 'application/json');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

function findFile_() {
  const it = DriveApp.getFolderById(CONFIG.FOLDER_ID).getFilesByName(CONFIG.FILE_NAME);
  return it.hasNext() ? it.next() : null;
}

/* ---------- Google Sheet (optional) ---------- */
function readSheet_() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss && ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift().map(h => String(h).trim().toLowerCase());
  const idx = CONFIG.FIELDS.map(f => header.indexOf(f));
  if (idx.indexOf(-1) !== -1) throw new Error('Articles sheet is missing a required column');
  return rows.filter(r => r[idx[0]] !== '' && r[idx[1]] !== '').map(r => {
    const a = { source: CONFIG.OWN_SOURCE, url: '' };
    CONFIG.FIELDS.forEach((f, i) => {
      let v = r[idx[i]];
      if (v instanceof Date) v = v.toISOString();
      a[f] = typeof v === 'string' ? v.trim() : v;
    });
    return a;
  });
}

/* ---------- "Sources" sheet: paste a category link and it is picked up automatically ----------
 * Add a tab named "Sources" with this header row:  source | category | url | feed_url | enabled
 *   source   = publisher name (e.g. BBC News)       category = e.g. Sport
 *   url      = the category page OR its RSS link    feed_url = filled in for you the first time
 *   enabled  = type "no" to pause a row
 */
function getFeeds_() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss && ss.getSheetByName(CONFIG.SOURCES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return FEEDS;
  const rows = sheet.getDataRange().getValues();
  const head = rows.shift().map(h => String(h).trim().toLowerCase());
  const [cS, cC, cU, cF, cE] = ['source', 'category', 'url', 'feed_url', 'enabled'].map(n => head.indexOf(n));
  if (cS < 0 || cU < 0) throw new Error('Sources sheet needs columns: source, category, url, feed_url, enabled');
  const feeds = [];
  rows.forEach((r, i) => {
    const url = String(r[cU] || '').trim();
    if (!url || (cE >= 0 && String(r[cE]).trim().toLowerCase() === 'no')) return;
    let feedUrl = cF >= 0 ? String(r[cF] || '').trim() : '';
    if (!feedUrl) {
      feedUrl = discoverFeed_(url);
      if (feedUrl && cF >= 0) sheet.getRange(i + 2, cF + 1).setValue(feedUrl);   // remember it
    }
    if (feedUrl) feeds.push({ name: String(r[cS]).trim(), category: cC >= 0 ? String(r[cC]).trim() : '', url: feedUrl });
    else console.warn('No RSS feed found for ' + url + ' (paste the feed link into feed_url)');
  });
  return feeds;
}

/** Turns a category page link into its RSS/Atom feed URL ('' if none is found). */
function discoverFeed_(pageUrl) {
  const get = u => {
    try {
      const r = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (compatible; NewslineBot/1.0)' } });
      return r.getResponseCode() === 200 ? r.getContentText() : '';
    } catch (e) { return ''; }
  };
  const isFeed = t => /<(rss|feed|rdf:RDF)[\s>]/i.test(t.slice(0, 3000));
  const page = get(pageUrl);
  if (isFeed(page)) return pageUrl;                        // they pasted the feed itself

  const origin = (/^https?:\/\/[^\/]+/i.exec(pageUrl) || [''])[0];
  const path = pageUrl.slice(origin.length).replace(/[?#].*$/, '').replace(/\/$/, '');
  const found = [];
  (page.match(/<link\b[^>]*>/gi) || []).forEach(tag => {   // <link rel="alternate" type="application/rss+xml" ...>
    const href = /href=["']([^"']+)["']/i.exec(tag);
    if (href && /type=["']application\/(rss|atom)\+xml/i.test(tag)) {
      let u = href[1].replace(/&amp;/g, '&');
      if (u.indexOf('//') === 0) u = 'https:' + u; else if (u[0] === '/') u = origin + u;
      found.push(u);
    }
  });
  const match = path && found.filter(u => u.indexOf(path) !== -1)[0];   // prefer the category's own feed
  if (match) return match;
  const base = pageUrl.replace(/[?#].*$/, '').replace(/\/$/, '');
  for (const g of [base + '/feed', base + '/feed/', base + '/rss', base + '/rss.xml', base + '/feed.xml']) {
    if (isFeed(get(g))) return g;
  }
  return found[0] || '';
}

/* ---------- RSS / Atom ---------- */
function fetchFeed_(feed) {
  const res = UrlFetchApp.fetch(feed.url, {
    muteHttpExceptions: true, followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; NewslineBot/1.0)' }
  });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  return parseFeed_(res.getContentText(), feed);
}

function parseFeed_(xml, feed) {
  const root = XmlService.parse(xml).getRootElement();
  const atom = root.getName() === 'feed';
  const an = atom ? root.getNamespace() : null;
  const items = atom ? root.getChildren('entry', an) : root.getChild('channel').getChildren('item');
  const kid = (el, name, ns) => (ns ? el.getChild(name, ns) : el.getChild(name));
  const txt = (el, name, ns) => { const c = kid(el, name, ns); return c ? c.getText() : ''; };

  return items.slice(0, CONFIG.PER_FEED).map(it => {
    let link = '';
    if (atom) { const l = kid(it, 'link', an); link = l && l.getAttribute('href') ? l.getAttribute('href').getValue() : ''; }
    else link = txt(it, 'link') || txt(it, 'guid');
    const html = atom ? (txt(it, 'content', an) || txt(it, 'summary', an))
                      : (txt(it, 'encoded', NS.content) || txt(it, 'description'));
    const date = new Date(atom ? (txt(it, 'published', an) || txt(it, 'updated', an))
                               : (txt(it, 'pubDate') || txt(it, 'date', NS.dc)));
    const author = atom ? txt(kid(it, 'author', an) || it, 'name', an) : (txt(it, 'creator', NS.dc) || txt(it, 'author'));
    const cat = txt(it, 'category');
    return {
      id: Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, link)).slice(0, 12),
      title: clean_(txt(it, 'title', an), 300),
      content: clean_(html, CONFIG.SUMMARY_CHARS),
      author: clean_(author, 80),
      category: feed.category || clean_(cat, 40),
      image_url: bestImage_(it, html),
      published_date: (isNaN(date) ? new Date() : date).toISOString(),
      source: feed.name,
      url: link
    };
  }).filter(a => a.title && a.url);
}

function bestImage_(it, html) {
  let best = '', bestW = -1;
  ['content', 'thumbnail'].forEach(name => {
    it.getChildren(name, NS.media).forEach(m => {
      const u = m.getAttribute('url'), t = m.getAttribute('type'), w = m.getAttribute('width');
      if (!u || (t && t.getValue().indexOf('image') !== 0)) return;   // skip video/audio
      const width = w ? Number(w.getValue()) || 0 : 0;
      if (width > bestW) { best = u.getValue(); bestW = width; }
    });
  });
  if (!best) {
    const enc = it.getChild('enclosure'), t = enc && enc.getAttribute('type');
    if (enc && t && t.getValue().indexOf('image') === 0) best = enc.getAttribute('url').getValue();
  }
  if (!best) { const m = /<img[^>]+src=["']([^"']+)["']/i.exec(html || ''); if (m) best = m[1]; }
  return best;
}

function clean_(s, max) {
  s = String(s || '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(n)).replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
  return s.length > max ? s.slice(0, max - 1).replace(/\s+\S*$/, '') + '…' : s;
}

/* ---------- Web App + trigger ---------- */
function doGet(e) {
  const file = findFile_() || generateJSON();
  if (e && e.parameter && e.parameter.data === '1') {
    return ContentService.createTextOutput(file.getBlob().getDataAsString()).setMimeType(ContentService.MimeType.JSON);
  }
  const id = file.getId();
  return ContentService.createTextOutput(JSON.stringify({
    fileId: id,
    url: 'https://drive.usercontent.google.com/download?id=' + id + '&export=download',
    lastUpdated: file.getLastUpdated().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/** Run ONCE. Replaces old triggers and refreshes every CONFIG.REFRESH_MINUTES. */
function createTimeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'generateJSON') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('generateJSON').timeBased().everyMinutes(CONFIG.REFRESH_MINUTES).create();
}