// settingsManager.js - Roaming Settings for user preferences (v4 - block-based with per-signature language/type)

var SETTINGS_KEY = 'acadon_signature_prefs';
var ASSET_BASE_URL = 'https://3dcut.github.io/deployOutlookSignatureAddin/assets';

function _defaultPreferences(officeLocation) {
  var lang = resolveLanguage(officeLocation);
  var langLower = lang.toLowerCase();
  return {
    version: 5,
    autoInsertEnabled: true,

    overrides: {
      phone: null,
      jobTitle: null,
      address: null
    },
    signatures: [
      {
        id: 'sig_default_long',
        name: 'Standard (lang)',
        language: lang,
        type: 'long',
        blocks: [
          { blockId: 'greeting_' + langLower },
          { blockId: 'layout_logo_start' },
          { blockId: 'nameblock_full' },
          { blockId: 'layout_logo_end' },
          { blockId: 'address_' + langLower },
          { blockId: 'legal_' + langLower }
        ],
        customBlocks: []
      },
      {
        id: 'sig_default_short',
        name: 'Kompakt (kurz)',
        language: lang,
        type: 'short',
        blocks: [
          { blockId: 'greeting_' + langLower },
          { blockId: 'nameblock_compact' }
        ],
        customBlocks: []
      }
    ],
    assignments: {
      newMessage: 'sig_default_long',
      reply: 'sig_default_short'
    },
    lastUpdated: new Date().toISOString()
  };
}

function getPreferences() {
  try {
    var stored = Office.context.roamingSettings.get(SETTINGS_KEY);
    if (stored && stored.version) {
      return stored;
    }
  } catch (e) {
    // roamingSettings not available (e.g. during first load)
  }
  return null;
}

function getPreferencesOrDefaults(officeLocation) {
  var prefs = getPreferences();
  if (!prefs) return _defaultPreferences(officeLocation);

  // Auto-migrate v2 -> v3 -> v4
  if (prefs.version < 3) {
    prefs = _migrateV2toV3(prefs);
  }
  if (prefs.version < 4) {
    prefs = _migrateV3toV4(prefs);
  }

  if (prefs.version < 5) {
    prefs = _migrateV4toV5(prefs);
  }
  if (prefs._migrated) {
    delete prefs._migrated;
    savePreferences(prefs);
  }

  return prefs;
}

function savePreferences(prefs, callback) {
  prefs.lastUpdated = new Date().toISOString();
  Office.context.roamingSettings.set(SETTINGS_KEY, prefs);
  Office.context.roamingSettings.saveAsync(function(result) {
    if (callback) {
      callback(result.status === Office.AsyncResultStatus.Succeeded);
    }
  });
}

function clearPreferences(callback) {
  Office.context.roamingSettings.remove(SETTINGS_KEY);
  Office.context.roamingSettings.saveAsync(function(result) {
    if (callback) {
      callback(result.status === Office.AsyncResultStatus.Succeeded);
    }
  });
}

function mergeUserData(graphData, overrides, language) {
  var companyInfo = resolveCompanyInfo(language || 'DE');
  var merged = {
    givenName: graphData.givenName || '',
    surname: graphData.surname || '',
    jobTitle: graphData.jobTitle || '',
    phone: graphData.phone || '',
    mail: graphData.mail || '',
    address: graphData.address || '',
    companyName: companyInfo.companyName,
    websiteUrl: companyInfo.websiteUrl,
    assetBaseUrl: ASSET_BASE_URL
  };

  if (overrides) {
    if (overrides.phone) merged.phone = overrides.phone;
    if (overrides.jobTitle) merged.jobTitle = overrides.jobTitle;
    if (overrides.address) merged.address = overrides.address;
  }

  return merged;
}

// --- Signature CRUD ---

function getSignatureById(prefs, sigId) {
  if (!prefs || !prefs.signatures) return null;
  for (var i = 0; i < prefs.signatures.length; i++) {
    if (prefs.signatures[i].id === sigId) return prefs.signatures[i];
  }
  return null;
}

function addSignature(prefs, signature) {
  if (!prefs.signatures) prefs.signatures = [];
  // Ensure new signatures have v4 fields
  if (!signature.language) signature.language = 'DE';
  if (!signature.type) signature.type = 'long';
  if (!signature.customBlocks) signature.customBlocks = [];
  prefs.signatures.push(signature);
}

function removeSignature(prefs, sigId) {
  if (!prefs.signatures) return;
  prefs.signatures = prefs.signatures.filter(function(s) {
    return s.id !== sigId;
  });
  // Reset assignments if they point to the deleted signature
  if (prefs.assignments) {
    if (prefs.assignments.newMessage === sigId) {
      prefs.assignments.newMessage = prefs.signatures.length > 0 ? prefs.signatures[0].id : null;
    }
    if (prefs.assignments.reply === sigId) {
      prefs.assignments.reply = prefs.signatures.length > 0 ? prefs.signatures[0].id : null;
    }
  }
}

function updateSignature(prefs, sigId, updates) {
  var sig = getSignatureById(prefs, sigId);
  if (!sig) return;
  if (updates.name !== undefined) sig.name = updates.name;
  if (updates.blocks !== undefined) sig.blocks = updates.blocks;
  if (updates.language !== undefined) sig.language = updates.language;
  if (updates.type !== undefined) sig.type = updates.type;
}

// --- Block operations within a signature ---

function addBlockToSignature(prefs, sigId, blockId, position) {
  var sig = getSignatureById(prefs, sigId);
  if (!sig) return;
  var entry = { blockId: blockId };
  if (position !== undefined && position >= 0 && position <= sig.blocks.length) {
    sig.blocks.splice(position, 0, entry);
  } else {
    sig.blocks.push(entry);
  }
}

function removeBlockFromSignature(prefs, sigId, blockIndex) {
  var sig = getSignatureById(prefs, sigId);
  if (!sig || blockIndex < 0 || blockIndex >= sig.blocks.length) return;

  // If removing a custom block, also remove from customBlocks array
  var removedBlockId = sig.blocks[blockIndex].blockId;
  if (removedBlockId && removedBlockId.indexOf('custom_') === 0) {
    removeCustomBlock(prefs, sigId, removedBlockId);
  }

  sig.blocks.splice(blockIndex, 1);
}

function moveBlockInSignature(prefs, sigId, fromIndex, toIndex) {
  var sig = getSignatureById(prefs, sigId);
  if (!sig) return;
  if (fromIndex < 0 || fromIndex >= sig.blocks.length) return;
  if (toIndex < 0 || toIndex >= sig.blocks.length) return;
  var item = sig.blocks.splice(fromIndex, 1)[0];
  sig.blocks.splice(toIndex, 0, item);
}

// --- Custom Block operations ---

function addCustomBlock(prefs, sigId, customBlock) {
  var sig = getSignatureById(prefs, sigId);
  if (!sig) return null;
  if (!sig.customBlocks) sig.customBlocks = [];

  var block = {
    id: customBlock.id || ('custom_' + Date.now()),
    name: customBlock.name || 'Eigener Baustein',
    htmlContent: customBlock.htmlContent || ''
  };
  sig.customBlocks.push(block);

  // Also add to blocks list
  sig.blocks.push({ blockId: block.id });

  return block;
}

function removeCustomBlock(prefs, sigId, customBlockId) {
  var sig = getSignatureById(prefs, sigId);
  if (!sig || !sig.customBlocks) return;
  sig.customBlocks = sig.customBlocks.filter(function(cb) {
    return cb.id !== customBlockId;
  });
}

function getCustomBlockHtml(signature, customBlockId) {
  if (!signature || !signature.customBlocks) return '';
  for (var i = 0; i < signature.customBlocks.length; i++) {
    if (signature.customBlocks[i].id === customBlockId) {
      return signature.customBlocks[i].htmlContent || '';
    }
  }
  return '';
}

// --- Migration v3 -> v4 ---

function _migrateV3toV4(v3Prefs) {
  var globalLang = v3Prefs.language || 'DE';

  // Add language/type/customBlocks to each signature
  if (v3Prefs.signatures) {
    v3Prefs.signatures.forEach(function(sig) {
      if (!sig.language) sig.language = globalLang;
      if (!sig.type) {
        // Derive type from name or blocks
        var nameLower = (sig.name || '').toLowerCase();
        if (nameLower.indexOf('kompakt') >= 0 || nameLower.indexOf('kurz') >= 0 ||
            nameLower.indexOf('compact') >= 0 || nameLower.indexOf('short') >= 0) {
          sig.type = 'short';
        } else {
          sig.type = 'long';
        }
      }
      if (!sig.customBlocks) sig.customBlocks = [];

      // Replace branding_logo_social with layout markers in long signatures
      if (sig.type === 'long') {
        var newBlocks = [];
        var hasBranding = false;
        for (var i = 0; i < sig.blocks.length; i++) {
          var bid = sig.blocks[i].blockId;
          if (bid === 'branding_logo_social') {
            hasBranding = true;
            // Skip it - we'll add layout markers instead
          } else {
            newBlocks.push(sig.blocks[i]);
          }
        }
        if (hasBranding) {
          // Insert layout_logo_start before nameblock, layout_logo_end after nameblock
          var nameIdx = -1;
          for (var j = 0; j < newBlocks.length; j++) {
            if (newBlocks[j].blockId.indexOf('nameblock') === 0) {
              nameIdx = j;
              break;
            }
          }
          if (nameIdx >= 0) {
            newBlocks.splice(nameIdx, 0, { blockId: 'layout_logo_start' });
            newBlocks.splice(nameIdx + 2, 0, { blockId: 'layout_logo_end' });
          }
          sig.blocks = newBlocks;
        }
      }
    });
  }

  // Remove global language field
  delete v3Prefs.language;
  v3Prefs.version = 4;
  v3Prefs._migrated = true;

  return v3Prefs;
}

// --- Migration v4 -> v5 ---

function _migrateV4toV5(v4Prefs) {
  if (v4Prefs.autoInsertEnabled === undefined) {
    v4Prefs.autoInsertEnabled = true;
  }
  v4Prefs.version = 5;
  v4Prefs._migrated = true;
  return v4Prefs;
}

// --- Migration v2 -> v3 ---

function _migrateV2toV3(v2Prefs) {
  var lang = v2Prefs.language || 'DE';
  var langLower = lang.toLowerCase();

  var signatures = [];
  var assignments = { newMessage: null, reply: null };

  var newMsgStyle = v2Prefs.templateStyle || 'acadon_long';
  var replyStyle = v2Prefs.templateStyleReply || 'acadon_short';

  // Skip profile-based styles for base signatures
  var isNewMsgProfile = newMsgStyle.indexOf('custom_') === 0;
  var isReplyProfile = replyStyle.indexOf('custom_') === 0;

  if (!isNewMsgProfile) {
    var sig1 = _createSignatureFromOldStyle(newMsgStyle, langLower, 'sig_migrated_1', 'Standard (lang)');
    signatures.push(sig1);
    assignments.newMessage = sig1.id;
  }

  if (!isReplyProfile) {
    if (replyStyle === newMsgStyle && !isNewMsgProfile) {
      assignments.reply = 'sig_migrated_1';
    } else {
      var sig2 = _createSignatureFromOldStyle(replyStyle, langLower, 'sig_migrated_2', 'Kompakt (kurz)');
      signatures.push(sig2);
      assignments.reply = sig2.id;
    }
  }

  // Migrate custom profiles as additional signatures
  if (v2Prefs.profiles && v2Prefs.profiles.length > 0) {
    v2Prefs.profiles.forEach(function(profile) {
      var baseBlocks = _getBlocksForStyle(profile.baseTemplate || 'acadon_long', langLower);
      var migrated = {
        id: profile.id,
        name: profile.name || 'Migriertes Profil',
        blocks: baseBlocks
      };
      signatures.push(migrated);

      if (isNewMsgProfile && newMsgStyle === profile.id) assignments.newMessage = profile.id;
      if (isReplyProfile && replyStyle === profile.id) assignments.reply = profile.id;
    });
  }

  // Migrate enabled addons: append to all signatures
  if (v2Prefs.enabledAddons && v2Prefs.enabledAddons.length > 0) {
    signatures.forEach(function(sig) {
      v2Prefs.enabledAddons.forEach(function(addonId) {
        sig.blocks.push({ blockId: addonId });
      });
    });
  }

  // Fallback if no assignments were set
  if (!assignments.newMessage && signatures.length > 0) assignments.newMessage = signatures[0].id;
  if (!assignments.reply && signatures.length > 0) assignments.reply = signatures[0].id;

  var v3 = {
    version: 3,
    language: lang,
    overrides: v2Prefs.overrides || { phone: null, jobTitle: null, address: null },
    signatures: signatures,
    assignments: assignments,
    lastUpdated: new Date().toISOString(),
    _migrated: true
  };

  return v3;
}

function _createSignatureFromOldStyle(style, langLower, id, name) {
  return {
    id: id,
    name: name,
    blocks: _getBlocksForStyle(style, langLower)
  };
}

function _getBlocksForStyle(style, langLower) {
  if (style === 'acadon_long') {
    return [
      { blockId: 'greeting_' + langLower },
      { blockId: 'nameblock_full' },
      { blockId: 'branding_logo_social' },
      { blockId: 'address_' + langLower },
      { blockId: 'legal_' + langLower }
    ];
  } else {
    return [
      { blockId: 'greeting_' + langLower },
      { blockId: 'nameblock_compact' }
    ];
  }
}
