// blockLoader.js - Block registry and HTML loading with localStorage cache
// Replaces templateLoader.js and addonLoader.js

var BLOCK_REGISTRY_URL = '../../blocks/blocks.json';
var BLOCK_CACHE_KEY = 'acadon_block_registry';
var BLOCK_CONTENT_PREFIX = 'acadon_block_';
var BLOCK_CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours

var _registryCache = null;

async function getBlockRegistry() {
  // Return in-memory cache if available
  if (_registryCache) return _registryCache;

  // Check localStorage cache
  try {
    var raw = localStorage.getItem(BLOCK_CACHE_KEY);
    if (raw) {
      var entry = JSON.parse(raw);
      if (Date.now() - entry.timestamp < BLOCK_CACHE_TTL) {
        _registryCache = entry.data;
        return _registryCache;
      }
    }
  } catch (e) { /* continue */ }

  // Fetch from server
  try {
    var response = await fetch(BLOCK_REGISTRY_URL);
    if (!response.ok) return { version: 1, blocks: [], presets: [] };
    var data = await response.json();

    // Cache in localStorage
    try {
      localStorage.setItem(BLOCK_CACHE_KEY, JSON.stringify({
        data: data,
        timestamp: Date.now()
      }));
    } catch (e) { /* continue */ }

    _registryCache = data;
    return data;
  } catch (e) {
    console.error('Failed to load block registry:', e.message);
    return { version: 1, blocks: [], presets: [] };
  }
}

async function getBlockHtml(blockId, format) {
  var registry = await getBlockRegistry();
  var blockDef = _getBlockDef(registry, blockId);
  if (!blockDef) {
    console.warn('Block not found in registry:', blockId);
    return '';
  }

  var fileKey = format === 'txt' ? 'txtFile' : 'htmlFile';
  var filePath = blockDef[fileKey];
  if (!filePath) return '';

  // Check localStorage cache
  var cacheKey = BLOCK_CONTENT_PREFIX + blockId + '_' + format;
  try {
    var raw = localStorage.getItem(cacheKey);
    if (raw) {
      var entry = JSON.parse(raw);
      if (Date.now() - entry.timestamp < BLOCK_CACHE_TTL) {
        return entry.content;
      }
    }
  } catch (e) { /* continue */ }

  // Fetch from server
  try {
    var url = '../../blocks/' + filePath;
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
  } catch (e) {
    console.warn('Failed to load block:', blockId, e.message);
    return '';
  }
}

function _getBlockDef(registry, blockId) {
  if (!registry || !registry.blocks) return null;
  for (var i = 0; i < registry.blocks.length; i++) {
    if (registry.blocks[i].id === blockId) return registry.blocks[i];
  }
  return null;
}

function getBlocksForLanguage(registry, lang) {
  if (!registry || !registry.blocks) return [];
  return registry.blocks.filter(function(b) {
    return b.language === null || b.language === lang;
  });
}

function getPreset(registry, presetId) {
  if (!registry || !registry.presets) return null;
  for (var i = 0; i < registry.presets.length; i++) {
    if (registry.presets[i].id === presetId) return registry.presets[i];
  }
  return null;
}

function getPresetsForLanguage(registry, lang) {
  if (!registry || !registry.presets) return [];
  return registry.presets.filter(function(p) {
    return p.language === lang;
  });
}

function clearBlockCache() {
  _registryCache = null;
  try {
    var keysToRemove = [];
    for (var i = 0; i < localStorage.length; i++) {
      var key = localStorage.key(i);
      if (key === BLOCK_CACHE_KEY || key.indexOf(BLOCK_CONTENT_PREFIX) === 0) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(function(k) { localStorage.removeItem(k); });
  } catch (e) { /* continue */ }
}
