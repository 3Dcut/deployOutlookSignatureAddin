// graphClient.js - SSO token acquisition + Microsoft Graph API

var GRAPH_CACHE_KEY = 'acadon_graph_userdata';
var GRAPH_CACHE_TTL = 4 * 60 * 60 * 1000; // 4 hours

function _getCachedUserData() {
  try {
    var raw = localStorage.getItem(GRAPH_CACHE_KEY);
    if (!raw) return null;

    var entry = JSON.parse(raw);
    if (Date.now() - entry.timestamp > GRAPH_CACHE_TTL) {
      localStorage.removeItem(GRAPH_CACHE_KEY);
      return null;
    }
    return entry.data;
  } catch (e) {
    return null;
  }
}

function _setCachedUserData(data) {
  try {
    localStorage.setItem(GRAPH_CACHE_KEY, JSON.stringify({
      data: data,
      timestamp: Date.now()
    }));
  } catch (e) {
    // Cache unavailable
  }
}

async function getAccessToken() {
  try {
    var token = await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      forMSGraphAccess: true
    });
    return token;
  } catch (err) {
    console.warn('SSO token acquisition failed:', err.code, err.message);

    // 13001 = SSO not supported on this platform
    // 13003 = consent required
    // 13005 = invalid resource
    // 13007 = user not signed in
    // 13012 = runtime not supported
    if (err.code === 13003) {
      // Try with consent prompt
      try {
        var tokenWithConsent = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true
        });
        return tokenWithConsent;
      } catch (retryErr) {
        console.error('SSO with consent failed:', retryErr.code, retryErr.message);
        return null;
      }
    }

    return null;
  }
}

function _transformGraphResponse(graphData) {
  var phone = '';
  if (graphData.businessPhones && graphData.businessPhones.length > 0) {
    phone = graphData.businessPhones[0];
  } else if (graphData.mobilePhone) {
    phone = graphData.mobilePhone;
  }

  var officeLocation = graphData.officeLocation || 'Krefeld';
  var address = resolveAddress(officeLocation);

  return {
    givenName: graphData.givenName || '',
    surname: graphData.surname || '',
    jobTitle: graphData.jobTitle || '',
    phone: phone,
    mail: graphData.mail || '',
    officeLocation: officeLocation,
    address: address
  };
}

async function getUserData() {
  // Check cache first
  var cached = _getCachedUserData();
  if (cached) return cached;

  // Try SSO + Graph API
  var token = await getAccessToken();

  if (token) {
    try {
      var response = await fetch(
        'https://graph.microsoft.com/v1.0/me?$select=givenName,surname,jobTitle,businessPhones,mobilePhone,mail,officeLocation',
        {
          headers: { 'Authorization': 'Bearer ' + token }
        }
      );

      if (response.ok) {
        var graphData = await response.json();
        var userData = _transformGraphResponse(graphData);
        _setCachedUserData(userData);
        return userData;
      }

      console.warn('Graph API request failed:', response.status);
    } catch (fetchErr) {
      console.warn('Graph API fetch error:', fetchErr.message);
    }
  }

  // Fallback: use Office.context.mailbox.userProfile for basic data
  try {
    var profile = Office.context.mailbox.userProfile;
    var fallbackData = {
      givenName: '',
      surname: '',
      jobTitle: '',
      phone: '',
      mail: profile.emailAddress || '',
      officeLocation: 'Krefeld',
      address: resolveAddress('Krefeld')
    };

    // Try to split display name into first/last
    var displayName = profile.displayName || '';
    var parts = displayName.split(' ');
    if (parts.length >= 2) {
      fallbackData.givenName = parts[0];
      fallbackData.surname = parts.slice(1).join(' ');
    } else if (parts.length === 1) {
      fallbackData.givenName = parts[0];
    }

    return fallbackData;
  } catch (profileErr) {
    console.error('Fallback userProfile also failed:', profileErr.message);
    return {
      givenName: '',
      surname: '',
      jobTitle: '',
      phone: '',
      mail: '',
      officeLocation: 'Krefeld',
      address: resolveAddress('Krefeld')
    };
  }
}

function clearUserDataCache() {
  localStorage.removeItem(GRAPH_CACHE_KEY);
}
