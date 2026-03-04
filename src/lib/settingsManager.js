// settingsManager.js - Roaming Settings for user preferences

var SETTINGS_KEY = 'acadon_signature_prefs';

function _defaultPreferences(officeLocation) {
  return {
    version: 1,
    templateStyle: 'acadon_long',
    templateStyleReply: 'acadon_short',
    language: resolveLanguage(officeLocation),
    overrides: {
      phone: null,
      jobTitle: null,
      address: null
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
  return getPreferences() || _defaultPreferences(officeLocation);
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

function mergeUserData(graphData, overrides) {
  var merged = {
    givenName: graphData.givenName || '',
    surname: graphData.surname || '',
    jobTitle: graphData.jobTitle || '',
    phone: graphData.phone || '',
    mail: graphData.mail || '',
    address: graphData.address || ''
  };

  if (overrides) {
    if (overrides.phone) merged.phone = overrides.phone;
    if (overrides.jobTitle) merged.jobTitle = overrides.jobTitle;
    if (overrides.address) merged.address = overrides.address;
  }

  return merged;
}
