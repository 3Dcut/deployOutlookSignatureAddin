// taskpane.js - Block-based signature management UI

var PANELS = ['loading', 'error-panel', 'main-form'];
var currentUserData = null;
var savedPrefs = null;
var blockRegistry = null;
var editingSignatureId = null;

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

    document.getElementById('language').value = savedPrefs.language;

    // Render signature UI
    renderSignatureList(savedPrefs);
    renderAssignmentDropdowns(savedPrefs);

    // Select first signature for editing
    if (savedPrefs.signatures && savedPrefs.signatures.length > 0) {
      selectSignatureForEditing(savedPrefs.signatures[0].id);
    }

    show('main-form');
    updatePreview();
  } catch (err) {
    showError('Fehler beim Laden: ' + err.message);
  }
}

// --- Signature List ---

function renderSignatureList(prefs) {
  var sel = document.getElementById('signatureSelect');
  sel.innerHTML = '';

  if (!prefs.signatures || prefs.signatures.length === 0) {
    var opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '-- Keine Signaturen --';
    sel.appendChild(opt);
    return;
  }

  prefs.signatures.forEach(function(sig) {
    var opt = document.createElement('option');
    opt.value = sig.id;
    opt.textContent = sig.name;
    sel.appendChild(opt);
  });

  if (editingSignatureId) {
    sel.value = editingSignatureId;
  }
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
        opt.textContent = sig.name;
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

  document.getElementById('signatureSelect').value = sigId;
  document.getElementById('sigName').value = sig.name;
  document.getElementById('sig-editor-section').classList.remove('hidden');
  document.getElementById('preset-section').classList.add('hidden');

  renderBlockList(sig);
  updatePreview();
}

function renderBlockList(sig) {
  var container = document.getElementById('block-list');
  container.innerHTML = '';

  if (!sig || !sig.blocks || sig.blocks.length === 0) {
    container.innerHTML = '<p class="info-text">Keine Bausteine. Klicke "+ Baustein hinzuf\u00fcgen".</p>';
    return;
  }

  sig.blocks.forEach(function(blockRef, index) {
    var blockDef = _getBlockDefFromRegistry(blockRef.blockId);
    var name = blockDef ? blockDef.name : blockRef.blockId;

    var item = document.createElement('div');
    item.className = 'block-item';

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

function handleMoveBlock(fromIndex, toIndex) {
  if (!editingSignatureId) return;
  moveBlockInSignature(savedPrefs, editingSignatureId, fromIndex, toIndex);
  var sig = getSignatureById(savedPrefs, editingSignatureId);
  renderBlockList(sig);
  updatePreview();
}

function handleRemoveBlock(blockIndex) {
  if (!editingSignatureId) return;
  removeBlockFromSignature(savedPrefs, editingSignatureId, blockIndex);
  var sig = getSignatureById(savedPrefs, editingSignatureId);
  renderBlockList(sig);
  updatePreview();
}

function handleAddBlock(blockId) {
  if (!editingSignatureId) return;
  addBlockToSignature(savedPrefs, editingSignatureId, blockId);
  var sig = getSignatureById(savedPrefs, editingSignatureId);
  renderBlockList(sig);
  document.getElementById('block-picker').classList.add('hidden');
  updatePreview();
}

// --- Block Picker ---

function toggleBlockPicker() {
  var picker = document.getElementById('block-picker');
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

  var lang = document.getElementById('language').value;
  var category = document.getElementById('pickerCategory').value;

  var availableBlocks = getBlocksForLanguage(blockRegistry, lang);

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

// --- Create / Delete Signatures ---

function createNewSignature() {
  var newId = 'sig_' + Date.now();
  var sig = {
    id: newId,
    name: 'Neue Signatur',
    blocks: []
  };
  addSignature(savedPrefs, sig);
  renderSignatureList(savedPrefs);
  renderAssignmentDropdowns(savedPrefs);
  selectSignatureForEditing(newId);

  // Show preset section for quick start
  var lang = document.getElementById('language').value;
  renderPresetOptions(lang);
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
  renderSignatureList(savedPrefs);
  renderAssignmentDropdowns(savedPrefs);

  if (savedPrefs.signatures && savedPrefs.signatures.length > 0) {
    selectSignatureForEditing(savedPrefs.signatures[0].id);
  } else {
    document.getElementById('sig-editor-section').classList.add('hidden');
  }
  updatePreview();
}

// --- Presets ---

function renderPresetOptions(lang) {
  var sel = document.getElementById('presetSelect');
  sel.innerHTML = '';

  if (!blockRegistry || !blockRegistry.presets) return;

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
  sig.blocks = preset.blockIds.map(function(id) { return { blockId: id }; });

  document.getElementById('sigName').value = sig.name;
  renderBlockList(sig);
  renderSignatureList(savedPrefs);
  document.getElementById('preset-section').classList.add('hidden');
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
    var language = document.getElementById('language').value;
    var mergedData = mergeUserData(userData, null, language);

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

// --- Save Preferences ---

function savePreferencesFromForm() {
  // Update signature name if editing
  if (editingSignatureId) {
    var nameInput = document.getElementById('sigName').value.trim();
    if (nameInput) {
      updateSignature(savedPrefs, editingSignatureId, { name: nameInput });
    }
  }

  savedPrefs.language = document.getElementById('language').value;

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

  savePreferences(savedPrefs, function(success) {
    if (success) {
      renderSignatureList(savedPrefs);
      renderAssignmentDropdowns(savedPrefs);
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
    var language = document.getElementById('language').value;
    var mergedData = mergeUserData(userData, null, language);

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

// --- Debounced preview ---
var _previewTimeout = null;
function debouncedPreview() {
  clearTimeout(_previewTimeout);
  _previewTimeout = setTimeout(updatePreview, 300);
}

// --- Event Listeners ---

document.getElementById('language').addEventListener('change', function() {
  renderBlockPicker();
  updatePreview();
});

document.getElementById('signatureSelect').addEventListener('change', function() {
  selectSignatureForEditing(this.value);
});

document.getElementById('new-sig-btn').addEventListener('click', createNewSignature);
document.getElementById('delete-sig-btn').addEventListener('click', deleteCurrentSignature);

document.getElementById('add-block-btn').addEventListener('click', toggleBlockPicker);
document.getElementById('pickerCategory').addEventListener('change', renderBlockPicker);

document.getElementById('apply-preset-btn').addEventListener('click', applyPreset);
document.getElementById('cancel-preset-btn').addEventListener('click', cancelPreset);

document.getElementById('sigName').addEventListener('input', function() {
  if (editingSignatureId) {
    updateSignature(savedPrefs, editingSignatureId, { name: this.value.trim() });
    renderSignatureList(savedPrefs);
    renderAssignmentDropdowns(savedPrefs);
  }
});

document.getElementById('phone').addEventListener('change', debouncedPreview);
document.getElementById('jobTitle').addEventListener('change', debouncedPreview);

document.getElementById('insert-btn').addEventListener('click', insertSignatureFromTaskpane);
document.getElementById('save-btn').addEventListener('click', savePreferencesFromForm);
document.getElementById('retry-btn').addEventListener('click', init);
