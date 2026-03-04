// taskpane.js - Task pane initialization, preview, and settings management

var PANELS = ['loading', 'error-panel', 'main-form'];
var currentUserData = null;

// --- Init ---

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

async function init() {
  show('loading');

  try {
    // Fetch user data from Graph API
    currentUserData = await getUserData();

    // Load saved preferences (if any)
    var prefs = getPreferencesOrDefaults(currentUserData.officeLocation);

    // Populate form with Graph data
    document.getElementById('firstName').value = currentUserData.givenName || '';
    document.getElementById('lastName').value = currentUserData.surname || '';
    document.getElementById('email').value = currentUserData.mail || '';
    document.getElementById('office').value = currentUserData.officeLocation || 'Krefeld';
    document.getElementById('address').value = currentUserData.address || '';

    // Populate editable fields - prefer overrides, fallback to Graph data
    document.getElementById('jobTitle').value =
      (prefs.overrides && prefs.overrides.jobTitle) || currentUserData.jobTitle || '';
    document.getElementById('phone').value =
      (prefs.overrides && prefs.overrides.phone) || currentUserData.phone || '';

    // Set saved preferences in dropdowns
    document.getElementById('language').value = prefs.language;
    document.getElementById('templateStyle').value = prefs.templateStyle;
    document.getElementById('templateStyleReply').value = prefs.templateStyleReply;

    show('main-form');

    // Load preview
    updatePreview();
  } catch (err) {
    showError('Fehler beim Laden: ' + err.message);
  }
}

// --- Preview ---

async function updatePreview() {
  var previewContainer = document.getElementById('preview-container');
  previewContainer.innerHTML = '<p class="info-text">Vorschau wird geladen...</p>';

  try {
    var userData = _getCurrentFormData();
    var language = document.getElementById('language').value;
    var style = document.getElementById('templateStyle').value;

    var html = await composeSignature(style, language, 'htm', userData);

    // Render preview in a sandboxed container
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

    // Auto-resize iframe to content height
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
  var prefs = {
    version: 1,
    templateStyle: document.getElementById('templateStyle').value,
    templateStyleReply: document.getElementById('templateStyleReply').value,
    language: document.getElementById('language').value,
    overrides: {
      phone: null,
      jobTitle: null,
      address: null
    }
  };

  // Save overrides only if they differ from Graph data
  var phoneInput = document.getElementById('phone').value.trim();
  if (phoneInput && currentUserData && phoneInput !== currentUserData.phone) {
    prefs.overrides.phone = phoneInput;
  }

  var jobTitleInput = document.getElementById('jobTitle').value.trim();
  if (jobTitleInput && currentUserData && jobTitleInput !== currentUserData.jobTitle) {
    prefs.overrides.jobTitle = jobTitleInput;
  }

  savePreferences(prefs, function(success) {
    var statusEl = document.getElementById('save-status');
    var msgEl = document.getElementById('save-message');

    if (success) {
      msgEl.textContent = 'Einstellungen gespeichert. Die Signatur wird ab der naechsten E-Mail aktualisiert.';
      msgEl.className = 'success-text';
    } else {
      msgEl.textContent = 'Fehler beim Speichern der Einstellungen.';
      msgEl.className = 'error-text';
    }

    statusEl.classList.remove('hidden');
    setTimeout(function() {
      statusEl.classList.add('hidden');
    }, 4000);
  });
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

// --- Event Listeners ---

document.getElementById('language').addEventListener('change', updatePreview);
document.getElementById('templateStyle').addEventListener('change', updatePreview);
document.getElementById('phone').addEventListener('change', updatePreview);
document.getElementById('jobTitle').addEventListener('change', updatePreview);

document.getElementById('save-btn').addEventListener('click', savePreferencesFromForm);
document.getElementById('retry-btn').addEventListener('click', init);
