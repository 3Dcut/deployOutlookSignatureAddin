// settingsManager.js - Roaming Settings for user preferences (v3 - block-based)

var SETTINGS_KEY = 'acadon_signature_prefs';

function _defaultPreferences(officeLocation) {
  var lang = resolveLanguage(officeLocation);
  var langLower = lang.toLowerCase();
  return {
    version: 3,
    language: lang,
    overrides: {
      phone: null,
      jobTitle: null,
      address: null
    },
    signatures: [
      {
        id: 'sig_default_long',
        name: 'Standard (lang)',
        blocks: [
          { blockId: 'greeting_' + langLower },
          { blockId: 'nameblock_full' },
          { blockId: 'branding_logo_social' },
          { blockId: 'address_' + langLower },
          { blockId: 'legal_' + langLower }
        ]
      },
      {
        id: 'sig_default_short',
        name: 'Kompakt (kurz)',
        blocks: [
          { blockId: 'greeting_' + langLower },
          { blockId: 'nameblock_compact' }
        ]
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

  // Auto-migrate from v1/v2
  if (prefs.version < 3) {
    prefs = _migrateV2toV3(prefs);
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
    websiteUrl: companyInfo.websiteUrl
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

  return {
    version: 3,
    language: lang,
    overrides: v2Prefs.overrides || { phone: null, jobTitle: null, address: null },
    signatures: signatures,
    assignments: assignments,
    lastUpdated: new Date().toISOString()
  };
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
