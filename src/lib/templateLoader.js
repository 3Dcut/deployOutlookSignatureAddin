// templateLoader.js - Template loading with localStorage cache

const TEMPLATE_CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours

function _cacheKey(lang, style, format) {
  return 'acadon_tmpl_' + lang + '_' + style + '_' + format;
}

function _getCached(lang, style, format) {
  try {
    const key = _cacheKey(lang, style, format);
    const raw = localStorage.getItem(key);
    if (!raw) return null;

    const entry = JSON.parse(raw);
    if (Date.now() - entry.timestamp > TEMPLATE_CACHE_TTL) {
      localStorage.removeItem(key);
      return null;
    }
    return entry.content;
  } catch (e) {
    return null;
  }
}

function _setCache(lang, style, format, content) {
  try {
    const key = _cacheKey(lang, style, format);
    localStorage.setItem(key, JSON.stringify({
      content: content,
      timestamp: Date.now()
    }));
  } catch (e) {
    // localStorage full or unavailable - continue without cache
  }
}

async function getTemplate(lang, style, format) {
  // Check cache first
  const cached = _getCached(lang, style, format);
  if (cached) return cached;

  // Fetch from server (relative to add-in origin)
  const url = '/templates/' + lang + '/' + style + '.' + format;
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error('Template nicht gefunden: ' + lang + '/' + style + '.' + format);
  }

  let content = await response.text();

  // Fix charset for HTML templates
  if (format === 'htm') {
    content = fixCharsetMeta(content);
  }

  _setCache(lang, style, format, content);
  return content;
}

async function preloadTemplates(lang) {
  const styles = ['acadon_long', 'acadon_short'];
  const formats = ['htm', 'txt'];
  const promises = [];

  for (const style of styles) {
    for (const format of formats) {
      promises.push(getTemplate(lang, style, format));
    }
  }

  await Promise.all(promises);
}

function clearTemplateCache() {
  const keysToRemove = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith('acadon_tmpl_')) {
      keysToRemove.push(key);
    }
  }
  keysToRemove.forEach(function(key) { localStorage.removeItem(key); });
}
