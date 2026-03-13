// taskpane.js - Block-based signature management UI (v4)

var PANELS = ['loading', 'error-panel', 'main-form'];
var currentUserData = null;
var savedPrefs = null;
var blockRegistry = null;
var editingSignatureId = null;
var previewDialog = null;

// --- Init ---

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

async function init() {
  show('loading');

  try {
    currentUserData = await getUserData();
    savedPrefs = getPreferencesOrDefaults(currentUserData.officeLocation);
    blockRegistry = await getBlockRegistry();

    // Populate Graph Data
    document.getElementById('firstName').value = currentUserData.givenName || '';
    document.getElementById('lastName').value = currentUserData.surname || '';
    document.getElementById('email').value = currentUserData.mail || '';
    document.getElementById('office').value = currentUserData.officeLocation || 'Krefeld';
    document.getElementById('address').value = currentUserData.address || '';

    // Populate overrides
    document.getElementById('jobTitle').value =
      (savedPrefs.overrides && savedPrefs.overrides.jobTitle) || currentUserData.jobTitle || '';
    document.getElementById('phone').value =
      (savedPrefs.overrides && savedPrefs.overrides.phone) || currentUserData.phone || '';

    // Auto-insert toggle
    document.getElementById('autoInsert').checked = savedPrefs.autoInsertEnabled !== false;

    // Render signature UI
    renderSignatureList();
    renderAssignmentDropdowns(savedPrefs);

    // Select first signature for editing
    if (savedPrefs.signatures && savedPrefs.signatures.length > 0) {
      selectSignatureForEditing(savedPrefs.signatures[0].id);
    }

    updateStorageUsage();
    show('main-form');
    updatePreview();

    // Fallback: Auto-insert if enabled
    if (savedPrefs.autoInsertEnabled !== false) {
      setTimeout(function() {
        console.log("Auto-insert triggered from taskpane load");
        insertSignatureFromTaskpane();
      }, 500); // Small delay to ensure items are ready
    }
  } catch (err) {
    showError('Fehler beim Laden: ' + err.message);
  }
}

// --- Signature List ---

function renderSignatureList() {
  var container = document.getElementById('sig-list');
  container.innerHTML = '';

  if (!savedPrefs.signatures || savedPrefs.signatures.length === 0) {
    container.innerHTML = '<div class="sig-list-empty">Keine Signaturen vorhanden</div>';
    return;
  }

  // Get filter values
  var langFilter = document.getElementById('filterLang').value;
  var typeFilter = document.getElementById('filterType').value;

  var filtered = savedPrefs.signatures.filter(function(sig) {
    if (langFilter !== 'all' && sig.language !== langFilter) return false;
    if (typeFilter !== 'all' && sig.type !== typeFilter) return false;
    return true;
  });

  if (filtered.length === 0) {
    container.innerHTML = '<div class="sig-list-empty">Keine Signaturen f\u00fcr diesen Filter</div>';
    return;
  }

  filtered.forEach(function(sig) {
    var item = document.createElement('div');
    item.className = 'sig-list-item';
    if (sig.id === editingSignatureId) {
      item.className += ' active';
    }
    item.setAttribute('data-sig-id', sig.id);

    var nameSpan = document.createElement('span');
    nameSpan.className = 'sig-name';
    nameSpan.textContent = sig.name;

    var metaSpan = document.createElement('span');
    metaSpan.className = 'sig-meta';
    metaSpan.textContent = (sig.language || '?') + ' / ' + (sig.type === 'short' ? 'kurz' : 'lang');

    item.appendChild(nameSpan);
    item.appendChild(metaSpan);

    item.addEventListener('click', function() {
      selectSignatureForEditing(sig.id);
    });

    container.appendChild(item);
  });
}

function renderAssignmentDropdowns(prefs) {
  var newMsgSel = document.getElementById('assignNewMessage');
  var replySel = document.getElementById('assignReply');

  [newMsgSel, replySel].forEach(function(sel) {
    sel.innerHTML = '';
    if (prefs.signatures) {
      prefs.signatures.forEach(function(sig) {
        var opt = document.createElement('option');
        opt.value = sig.id;
        opt.textContent = sig.name + ' (' + (sig.language || '?') + ')';
        sel.appendChild(opt);
      });
    }
  });

  if (prefs.assignments) {
    newMsgSel.value = prefs.assignments.newMessage || '';
    replySel.value = prefs.assignments.reply || '';
  }
}

// --- Signature Editor ---

function selectSignatureForEditing(sigId) {
  editingSignatureId = sigId;
  var sig = getSignatureById(savedPrefs, sigId);

  if (!sig) {
    document.getElementById('sig-editor-section').classList.add('hidden');
    return;
  }

  document.getElementById('sigName').value = sig.name;
  document.getElementById('sigLang').value = sig.language || 'DE';
  document.getElementById('sigType').value = sig.type || 'long';
  document.getElementById('sig-editor-section').classList.remove('hidden');
  document.getElementById('preset-section').classList.add('hidden');
  document.getElementById('block-picker').classList.add('hidden');
  document.getElementById('custom-block-creator').classList.add('hidden');

  renderSignatureList();
  renderBlockList(sig);
  updatePreview();
}

function renderBlockList(sig) {
  var container = document.getElementById('block-list');
  container.innerHTML = '';

  if (!sig || !sig.blocks || sig.blocks.length === 0) {
    container.innerHTML = '<p class="info-text">Keine Bausteine. Klicke "+ Baustein".</p>';
    return;
  }

  sig.blocks.forEach(function(blockRef, index) {
    var blockId = blockRef.blockId;
    var blockDef = _getBlockDefFromRegistry(blockId);
    var isLayout = blockId === 'layout_logo_start' || blockId === 'layout_logo_end';
    var isCustom = blockId.indexOf('custom_') === 0;

    var name;
    if (blockDef) {
      name = blockDef.name;
    } else if (isCustom) {
      var customBlock = _getCustomBlockDef(sig, blockId);
      name = customBlock ? customBlock.name : blockId;
    } else {
      name = blockId;
    }

    var item = document.createElement('div');
    item.className = 'block-item';
    if (isLayout) item.className += ' block-item-layout';
    if (isCustom) item.className += ' block-item-custom';

    var controls = document.createElement('div');
    controls.className = 'block-controls';

    var btnUp = document.createElement('button');
    btnUp.className = 'btn-icon';
    btnUp.innerHTML = '&#9650;';
    btnUp.title = 'Nach oben';
    btnUp.disabled = index === 0;
    btnUp.addEventListener('click', function() { handleMoveBlock(index, index - 1); });

    var btnDown = document.createElement('button');
    btnDown.className = 'btn-icon';
    btnDown.innerHTML = '&#9660;';
    btnDown.title = 'Nach unten';
    btnDown.disabled = index === sig.blocks.length - 1;
    btnDown.addEventListener('click', function() { handleMoveBlock(index, index + 1); });

    controls.appendChild(btnUp);
    controls.appendChild(btnDown);

    var label = document.createElement('span');
    label.className = 'block-label';
    label.textContent = name;
    if (blockDef && blockDef.language) {
      label.textContent += ' [' + blockDef.language + ']';
    }
    if (isCustom) {
      label.textContent += ' *';
    }

    var btnRemove = document.createElement('button');
    btnRemove.className = 'btn-icon btn-remove';
    btnRemove.innerHTML = '&times;';
    btnRemove.title = 'Entfernen';
    btnRemove.addEventListener('click', function() { handleRemoveBlock(index); });

    item.appendChild(controls);
    item.appendChild(label);
    item.appendChild(btnRemove);
    container.appendChild(item);
  });
}

function _getCustomBlockDef(sig, customBlockId) {
  if (!sig || !sig.customBlocks) return null;
  for (var i = 0; i < sig.customBlocks.length; i++) {
    if (sig.customBlocks[i].id === customBlockId) return sig.customBlocks[i];
  }
  return null;
}

function handleMoveBlock(fromIndex, toIndex) {
  if (!editingSignatureId) return;
  moveBlockInSignature(savedPrefs, editingSignatureId, fromIndex, toIndex);
  var sig = getSignatureById(savedPrefs, editingSignatureId);
  renderBlockList(sig);
  updateStorageUsage();
  updatePreview();
}

function handleRemoveBlock(blockIndex) {
  if (!editingSignatureId) return;
  removeBlockFromSignature(savedPrefs, editingSignatureId, blockIndex);
  var sig = getSignatureById(savedPrefs, editingSignatureId);
  renderBlockList(sig);
  updateStorageUsage();
  updatePreview();
}

function handleAddBlock(blockId) {
  if (!editingSignatureId) return;
  addBlockToSignature(savedPrefs, editingSignatureId, blockId);
  var sig = getSignatureById(savedPrefs, editingSignatureId);
  renderBlockList(sig);
  document.getElementById('block-picker').classList.add('hidden');
  updateStorageUsage();
  updatePreview();
}

// --- Block Picker ---

function toggleBlockPicker() {
  var picker = document.getElementById('block-picker');
  var customCreator = document.getElementById('custom-block-creator');
  customCreator.classList.add('hidden');

  if (picker.classList.contains('hidden')) {
    picker.classList.remove('hidden');
    renderBlockPicker();
  } else {
    picker.classList.add('hidden');
  }
}

function renderBlockPicker() {
  var container = document.getElementById('picker-list');
  container.innerHTML = '';

  if (!blockRegistry || !blockRegistry.blocks) return;

  // Use the signature's language for filtering
  var sig = editingSignatureId ? getSignatureById(savedPrefs, editingSignatureId) : null;
  var lang = sig ? (sig.language || 'DE') : 'DE';
  var category = document.getElementById('pickerCategory').value;

  var availableBlocks = getBlocksForLanguage(blockRegistry, lang);

  // Filter out deprecated blocks
  availableBlocks = availableBlocks.filter(function(b) {
    return !b.tags || b.tags.indexOf('deprecated') < 0;
  });

  if (category !== 'all') {
    availableBlocks = availableBlocks.filter(function(b) {
      return b.category === category;
    });
  }

  // Sort by sortOrder
  availableBlocks.sort(function(a, b) { return (a.sortOrder || 0) - (b.sortOrder || 0); });

  if (availableBlocks.length === 0) {
    container.innerHTML = '<p class="info-text">Keine Bausteine in dieser Kategorie.</p>';
    return;
  }

  availableBlocks.forEach(function(block) {
    var row = document.createElement('div');
    row.className = 'picker-item';

    var info = document.createElement('div');
    info.className = 'picker-info';

    var nameSpan = document.createElement('span');
    nameSpan.className = 'picker-name';
    nameSpan.textContent = block.name;

    var desc = document.createElement('span');
    desc.className = 'picker-desc';
    desc.textContent = block.description || '';

    info.appendChild(nameSpan);
    info.appendChild(desc);

    var btnAdd = document.createElement('button');
    btnAdd.className = 'btn-icon btn-add';
    btnAdd.textContent = '+';
    btnAdd.title = 'Hinzuf\u00fcgen';
    btnAdd.addEventListener('click', function() { handleAddBlock(block.id); });

    row.appendChild(info);
    row.appendChild(btnAdd);
    container.appendChild(row);
  });
}

// --- Custom Block Creator ---

function toggleCustomBlockCreator() {
  var creator = document.getElementById('custom-block-creator');
  var picker = document.getElementById('block-picker');
  picker.classList.add('hidden');

  if (creator.classList.contains('hidden')) {
    creator.classList.remove('hidden');
    document.getElementById('customBlockName').value = '';
    document.getElementById('customBlockHtml').value = '';
  } else {
    creator.classList.add('hidden');
  }
}

function createCustomBlock() {
  if (!editingSignatureId) return;

  var name = document.getElementById('customBlockName').value.trim();
  var html = document.getElementById('customBlockHtml').value.trim();

  if (!name) {
    showErrorMessage('Bitte einen Namen eingeben.');
    return;
  }
  if (!html) {
    showErrorMessage('Bitte HTML-Inhalt eingeben.');
    return;
  }

  var block = addCustomBlock(savedPrefs, editingSignatureId, {
    name: name,
    htmlContent: html
  });

  if (block) {
    var sig = getSignatureById(savedPrefs, editingSignatureId);
    renderBlockList(sig);
    document.getElementById('custom-block-creator').classList.add('hidden');
    updateStorageUsage();
    updatePreview();
  }
}

// --- Create / Delete Signatures ---

function createNewSignature() {
  var sig = editingSignatureId ? getSignatureById(savedPrefs, editingSignatureId) : null;
  var defaultLang = sig ? sig.language : 'DE';

  var newId = 'sig_' + Date.now();
  var newSig = {
    id: newId,
    name: 'Neue Signatur',
    language: defaultLang,
    type: 'long',
    blocks: [],
    customBlocks: []
  };
  addSignature(savedPrefs, newSig);
  renderSignatureList();
  renderAssignmentDropdowns(savedPrefs);
  selectSignatureForEditing(newId);
  updateStorageUsage();

  // Show preset section for quick start
  renderPresetOptions(defaultLang);
  document.getElementById('preset-section').classList.remove('hidden');
}

function deleteCurrentSignature() {
  if (!editingSignatureId) return;
  if (savedPrefs.signatures && savedPrefs.signatures.length <= 1) {
    showErrorMessage('Mindestens eine Signatur muss vorhanden sein.');
    return;
  }
  if (!confirm('Soll diese Signatur wirklich gel\u00f6scht werden?')) return;

  removeSignature(savedPrefs, editingSignatureId);
  editingSignatureId = null;
  renderSignatureList();
  renderAssignmentDropdowns(savedPrefs);

  if (savedPrefs.signatures && savedPrefs.signatures.length > 0) {
    selectSignatureForEditing(savedPrefs.signatures[0].id);
  } else {
    document.getElementById('sig-editor-section').classList.add('hidden');
  }
  updateStorageUsage();
  updatePreview();
}

// --- Presets ---

function renderPresetOptions(lang) {
  var sel = document.getElementById('presetSelect');
  sel.innerHTML = '';

  if (!blockRegistry || !blockRegistry.blocks) return;

  var presets = getPresetsForLanguage(blockRegistry, lang);
  var otherPresets = blockRegistry.presets.filter(function(p) { return p.language !== lang; });

  presets.forEach(function(p) {
    var opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = p.name;
    sel.appendChild(opt);
  });

  if (otherPresets.length > 0) {
    var group = document.createElement('optgroup');
    group.label = 'Andere Sprachen';
    otherPresets.forEach(function(p) {
      var opt = document.createElement('option');
      opt.value = p.id;
      opt.textContent = p.name + ' (' + p.language + ')';
      group.appendChild(opt);
    });
    sel.appendChild(group);
  }
}

function applyPreset() {
  if (!editingSignatureId || !blockRegistry) return;
  var presetId = document.getElementById('presetSelect').value;
  var preset = getPreset(blockRegistry, presetId);
  if (!preset) return;

  var sig = getSignatureById(savedPrefs, editingSignatureId);
  if (!sig) return;

  sig.name = preset.name;
  sig.language = preset.language || sig.language;
  sig.type = preset.type || sig.type;
  sig.blocks = preset.blockIds.map(function(id) { return { blockId: id }; });

  document.getElementById('sigName').value = sig.name;
  document.getElementById('sigLang').value = sig.language;
  document.getElementById('sigType').value = sig.type;
  renderBlockList(sig);
  renderSignatureList();
  document.getElementById('preset-section').classList.add('hidden');
  updateStorageUsage();
  updatePreview();
}

function cancelPreset() {
  document.getElementById('preset-section').classList.add('hidden');
}

// --- Preview ---

async function updatePreview() {
  var previewContainer = document.getElementById('preview-container');
  previewContainer.innerHTML = '<p class="info-text">Vorschau wird geladen...</p>';

  try {
    var sig = editingSignatureId ? getSignatureById(savedPrefs, editingSignatureId) : null;

    if (!sig || !sig.blocks || sig.blocks.length === 0) {
      previewContainer.innerHTML = '<p class="info-text">Keine Bausteine in der Signatur.</p>';
      return;
    }

    var userData = _getCurrentFormData();
    var sigLang = sig.language || 'DE';
    var mergedData = mergeUserData(userData, null, sigLang);

    var html = await composeSignature(sig, 'htm', mergedData);

    previewContainer.innerHTML = '';
    var frame = document.createElement('iframe');
    frame.sandbox = 'allow-same-origin';
    frame.style.width = '100%';
    frame.style.border = 'none';
    frame.style.minHeight = '200px';
    previewContainer.appendChild(frame);

    frame.contentDocument.open();
    frame.contentDocument.write(html);
    frame.contentDocument.close();

    frame.onload = function() {
      try {
        var height = frame.contentDocument.body.scrollHeight;
        frame.style.height = (height + 20) + 'px';
      } catch (e) {
        frame.style.height = '400px';
      }
    };
  } catch (err) {
    previewContainer.innerHTML =
      '<p class="error-text">Vorschau konnte nicht geladen werden: ' + err.message + '</p>';
  }
}

// --- Pop-out Preview Dialog ---

async function openPreviewDialog() {
  var sig = editingSignatureId ? getSignatureById(savedPrefs, editingSignatureId) : null;
  if (!sig || !sig.blocks || sig.blocks.length === 0) {
    showErrorMessage('Keine Bausteine in der Signatur.');
    return;
  }

  // Close existing dialog if open
  if (previewDialog) {
    previewDialog.close();
    previewDialog = null;
  }

  var userData = _getCurrentFormData();
  var sigLang = sig.language || 'DE';
  var mergedData = mergeUserData(userData, null, sigLang);
  var html = await composeSignature(sig, 'htm', mergedData);

  var dialogUrl = new URL('preview-dialog.html', window.location.href).href;

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 50, width: 60, displayInIframe: false },
    function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showErrorMessage('Dialog konnte nicht ge\u00f6ffnet werden: ' + asyncResult.error.message);
        return;
      }

      previewDialog = asyncResult.value;

      previewDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        function(arg) {
          handleDialogMessage(arg, html);
        }
      );

      previewDialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        function() {
          previewDialog = null;
        }
      );

      // Send HTML after dialog has initialized Office.js
      setTimeout(function() {
        if (previewDialog) {
          previewDialog.messageChild(JSON.stringify({
            type: 'preview',
            html: html
          }));
        }
      }, 1000);
    }
  );
}

function handleDialogMessage(arg, composedHtml) {
  try {
    var message = JSON.parse(arg.message);

    if (message.action === 'insert') {
      var htmlToInsert = message.html || composedHtml;

      if (Office.context.mailbox.item) {
        Office.context.mailbox.item.body.setSignatureAsync(
          htmlToInsert,
          { coercionType: Office.CoercionType.Html },
          function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              showSuccessMessage('Signatur eingef\u00fcgt.');
            } else {
              showErrorMessage('Fehler: ' + asyncResult.error.message);
            }
          }
        );
      } else {
        showErrorMessage('Keine E-Mail ge\u00f6ffnet.');
      }
    }

    if (previewDialog) {
      previewDialog.close();
      previewDialog = null;
    }
  } catch (e) {
    if (previewDialog) {
      previewDialog.close();
      previewDialog = null;
    }
  }
}

// --- Save Preferences ---

function savePreferencesFromForm() {
  // Update signature attributes if editing
  if (editingSignatureId) {
    var nameInput = document.getElementById('sigName').value.trim();
    var langInput = document.getElementById('sigLang').value;
    var typeInput = document.getElementById('sigType').value;

    if (nameInput) {
      updateSignature(savedPrefs, editingSignatureId, {
        name: nameInput,
        language: langInput,
        type: typeInput
      });
    }
  }

  if (!savedPrefs.overrides) savedPrefs.overrides = {};

  var phoneInput = document.getElementById('phone').value.trim();
  if (phoneInput && currentUserData && phoneInput !== currentUserData.phone) {
    savedPrefs.overrides.phone = phoneInput;
  } else {
    savedPrefs.overrides.phone = null;
  }

  var jobTitleInput = document.getElementById('jobTitle').value.trim();
  if (jobTitleInput && currentUserData && jobTitleInput !== currentUserData.jobTitle) {
    savedPrefs.overrides.jobTitle = jobTitleInput;
  } else {
    savedPrefs.overrides.jobTitle = null;
  }

  // Save assignments
  if (!savedPrefs.assignments) savedPrefs.assignments = {};
  savedPrefs.assignments.newMessage = document.getElementById('assignNewMessage').value;
  savedPrefs.assignments.reply = document.getElementById('assignReply').value;

  // Save auto-insert setting
  savedPrefs.autoInsertEnabled = document.getElementById('autoInsert').checked;

  savePreferences(savedPrefs, function(success) {
    if (success) {
      renderSignatureList();
      renderAssignmentDropdowns(savedPrefs);
      updateStorageUsage();
      showSuccessMessage('Einstellungen gespeichert. Die Signatur wird ab der n\u00e4chsten E-Mail aktualisiert.');
    } else {
      showErrorMessage('Fehler beim Speichern der Einstellungen.');
    }
  });
}

// --- Insert Signature ---

async function insertSignatureFromTaskpane() {
  if (!Office.context.mailbox.item) {
    showErrorMessage('Keine E-Mail ge\u00f6ffnet.');
    return;
  }

  try {
    var sig = editingSignatureId ? getSignatureById(savedPrefs, editingSignatureId) : null;
    if (!sig) {
      showErrorMessage('Keine Signatur ausgew\u00e4hlt.');
      return;
    }

    var userData = _getCurrentFormData();
    var sigLang = sig.language || 'DE';
    var mergedData = mergeUserData(userData, null, sigLang);

    var htmlSignature = await composeSignature(sig, 'htm', mergedData);

    Office.context.mailbox.item.body.setSignatureAsync(
      htmlSignature,
      { coercionType: Office.CoercionType.Html },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          showSuccessMessage('Signatur eingef\u00fcgt.');
        } else {
          showErrorMessage('Fehler: ' + asyncResult.error.message);
        }
      }
    );
  } catch (err) {
    showErrorMessage('Fehler beim Einf\u00fcgen: ' + err.message);
  }
}

// --- Helpers ---

function _getCurrentFormData() {
  return {
    givenName: document.getElementById('firstName').value,
    surname: document.getElementById('lastName').value,
    jobTitle: document.getElementById('jobTitle').value,
    phone: document.getElementById('phone').value,
    mail: document.getElementById('email').value,
    address: document.getElementById('address').value
  };
}

function _getBlockDefFromRegistry(blockId) {
  if (!blockRegistry || !blockRegistry.blocks) return null;
  for (var i = 0; i < blockRegistry.blocks.length; i++) {
    if (blockRegistry.blocks[i].id === blockId) return blockRegistry.blocks[i];
  }
  return null;
}

function show(panelId) {
  PANELS.forEach(function(id) {
    document.getElementById(id).classList.add('hidden');
  });
  document.getElementById(panelId).classList.remove('hidden');
}

function showError(msg) {
  document.getElementById('error-message').textContent = msg;
  show('error-panel');
}

function showSuccessMessage(msg) {
  var statusEl = document.getElementById('save-status');
  var msgEl = document.getElementById('save-message');
  msgEl.textContent = msg;
  msgEl.className = 'success-text';
  statusEl.classList.remove('hidden');
  setTimeout(function() { statusEl.classList.add('hidden'); }, 4000);
}

function showErrorMessage(msg) {
  var statusEl = document.getElementById('save-status');
  var msgEl = document.getElementById('save-message');
  msgEl.textContent = msg;
  msgEl.className = 'error-text';
  statusEl.classList.remove('hidden');
  setTimeout(function() { statusEl.classList.add('hidden'); }, 4000);
}

// --- Storage Usage ---

function updateStorageUsage() {
  if (!savedPrefs) return;

  try {
    var serialized = JSON.stringify(savedPrefs);
    // Rough estimate of byte size (UTF-8)
    var bytes = unescape(encodeURIComponent(serialized)).length;
    var kb = (bytes / 1024).toFixed(1);
    var limitKb = 32;
    var percent = Math.min(100, (bytes / (limitKb * 1024)) * 100);

    var usedEl = document.getElementById('storage-used');
    var fillEl = document.getElementById('storage-fill');

    if (usedEl) usedEl.textContent = kb;
    if (fillEl) {
      fillEl.style.width = percent + '%';
      
      // Update colors
      fillEl.classList.remove('warning', 'danger');
      if (percent > 95) {
        fillEl.classList.add('danger');
      } else if (percent > 80) {
        fillEl.classList.add('warning');
      }
    }
    
    console.log("Storage usage updated: " + bytes + " bytes (" + percent.toFixed(1) + "%)");
  } catch (e) {
    console.warn("Could not update storage usage display:", e);
  }
}

// --- Debounced preview ---
var _previewTimeout = null;
function debouncedPreview() {
  clearTimeout(_previewTimeout);
  _previewTimeout = setTimeout(updatePreview, 300);
}

// --- Event Listeners ---

// Filter controls
document.getElementById('filterLang').addEventListener('change', renderSignatureList);
document.getElementById('filterType').addEventListener('change', renderSignatureList);

// Signature editor - language/type changes
document.getElementById('sigLang').addEventListener('change', function() {
  if (editingSignatureId) {
    updateSignature(savedPrefs, editingSignatureId, { language: this.value });
    renderSignatureList();
    updateStorageUsage();
    updatePreview();
  }
});

document.getElementById('sigType').addEventListener('change', function() {
  if (editingSignatureId) {
    updateSignature(savedPrefs, editingSignatureId, { type: this.value });
    renderSignatureList();
    updateStorageUsage();
    updatePreview();
  }
});

document.getElementById('new-sig-btn').addEventListener('click', createNewSignature);
document.getElementById('delete-sig-btn').addEventListener('click', deleteCurrentSignature);

document.getElementById('add-block-btn').addEventListener('click', toggleBlockPicker);
document.getElementById('add-custom-block-btn').addEventListener('click', toggleCustomBlockCreator);
document.getElementById('pickerCategory').addEventListener('change', renderBlockPicker);

document.getElementById('create-custom-btn').addEventListener('click', createCustomBlock);
document.getElementById('cancel-custom-btn').addEventListener('click', function() {
  document.getElementById('custom-block-creator').classList.add('hidden');
});

document.getElementById('apply-preset-btn').addEventListener('click', applyPreset);
document.getElementById('cancel-preset-btn').addEventListener('click', cancelPreset);

document.getElementById('sigName').addEventListener('input', function() {
  if (editingSignatureId) {
    updateSignature(savedPrefs, editingSignatureId, { name: this.value.trim() });
    renderSignatureList();
    renderAssignmentDropdowns(savedPrefs);
    updateStorageUsage();
  }
});

document.getElementById('phone').addEventListener('change', debouncedPreview);
document.getElementById('jobTitle').addEventListener('change', debouncedPreview);

document.getElementById('popout-preview-btn').addEventListener('click', openPreviewDialog);
document.getElementById('insert-btn').addEventListener('click', insertSignatureFromTaskpane);

// Local save buttons in each accordion
document.querySelectorAll('.save-btn').forEach(function(btn) {
  btn.addEventListener('click', savePreferencesFromForm);
});

document.getElementById('retry-btn').addEventListener('click', init);

// Also update preview when address changes (it's read-only but just in case)
document.getElementById('address').addEventListener('change', debouncedPreview);
