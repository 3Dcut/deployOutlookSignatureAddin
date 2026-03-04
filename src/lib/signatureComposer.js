// signatureComposer.js - Orchestrates template loading, data merging and signature injection

async function composeSignature(style, language, format, userData) {
  var template = await getTemplate(language, style, format);
  return applyPlaceholders(template, userData);
}

async function injectSignature(event, isReply) {
  try {
    // 1. Load preferences
    var userData = await getUserData();
    var prefs = getPreferencesOrDefaults(userData.officeLocation);

    // 2. Merge user data with any overrides
    var mergedData = mergeUserData(userData, prefs.overrides);

    // 3. Select template style based on new message vs reply
    var style = isReply ? prefs.templateStyleReply : prefs.templateStyle;
    var language = prefs.language;

    // 4. Compose the HTML signature
    var htmlSignature = await composeSignature(style, language, 'htm', mergedData);

    // 5. Inject via setSignatureAsync
    var item = Office.context.mailbox.item;
    item.body.setSignatureAsync(
      htmlSignature,
      { coercionType: Office.CoercionType.Html },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error('setSignatureAsync failed:', asyncResult.error.message);
        }
        event.completed();
      }
    );
  } catch (err) {
    console.error('Signature injection error:', err.message);
    event.completed();
  }
}
