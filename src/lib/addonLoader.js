// addonLoader.js - Load optional building blocks (Zusatz-Bausteine)

var ADDONS_REGISTRY_URL = '../../addons/addons.json';
var ADDON_CACHE_KEY = 'acadon_addons_registry';
var ADDON_CONTENT_PREFIX = 'acadon_addon_';
var ADDON_CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours

async function getAddonsRegistry() {
  // Check cache
  try {
    var raw = localStorage.getItem(ADDON_CACHE_KEY);
    if (raw) {
      var entry = JSON.parse(raw);
      if (Date.now() - entry.timestamp < ADDON_CACHE_TTL) {
        return entry.data;
      }
    }
  } catch (e) { /* continue */ }

  // Fetch from server
  var response = await fetch(ADDONS_REGISTRY_URL);
  if (!response.ok) return [];
  var data = await response.json();

  // Cache
  try {
    localStorage.setItem(ADDON_CACHE_KEY, JSON.stringify({
      data: data,
      timestamp: Date.now()
    }));
  } catch (e) { /* continue */ }

  return data;
}

async function getAddonHtml(addonId, registry) {
  var addon = null;
  for (var i = 0; i < registry.length; i++) {
    if (registry[i].id === addonId) {
      addon = registry[i];
      break;
    }
  }
  if (!addon) return '';

  // Check content cache
  var cacheKey = ADDON_CONTENT_PREFIX + addonId;
  try {
    var raw = localStorage.getItem(cacheKey);
    if (raw) {
      var entry = JSON.parse(raw);
      if (Date.now() - entry.timestamp < ADDON_CACHE_TTL) {
        return entry.content;
      }
    }
  } catch (e) { /* continue */ }

  // Fetch HTML
  var url = '../../addons/' + addon.htmlFile;
  var response = await fetch(url);
  if (!response.ok) return '';
  var content = await response.text();

  // Cache content
  try {
    localStorage.setItem(cacheKey, JSON.stringify({
      content: content,
      timestamp: Date.now()
    }));
  } catch (e) { /* continue */ }

  return content;
}

async function composeAddonsHtml(enabledAddonIds) {
  if (!enabledAddonIds || enabledAddonIds.length === 0) return '';

  var registry = await getAddonsRegistry();
  var parts = [];

  for (var i = 0; i < enabledAddonIds.length; i++) {
    var html = await getAddonHtml(enabledAddonIds[i], registry);
    if (html) parts.push(html);
  }

  return parts.join('\n');
}
