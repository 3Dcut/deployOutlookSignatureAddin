// signatureComposer.js - Block-based signature composition

async function composeSignature(signatureObj, format, userData) {
  // signatureObj: { id, name, blocks: [{blockId}, ...] }
  // format: 'htm' or 'txt'
  // userData: merged user data with overrides applied

  var parts = [];

  for (var i = 0; i < signatureObj.blocks.length; i++) {
    var blockRef = signatureObj.blocks[i];
    var blockHtml = await getBlockHtml(blockRef.blockId, format);

    if (blockHtml) {
      var processed = applyPlaceholders(blockHtml, userData);
      parts.push(processed);
    }
  }

  if (format === 'htm') {
    return '<div style="font-family:\'Arial\',sans-serif;">\n' + parts.join('\n') + '\n</div>';
  } else {
    return parts.join('\n\n');
  }
}

async function injectSignature(event, isReply) {
  try {
    // 1. Load user data and preferences
    var userData = await getUserData();
    var prefs = getPreferencesOrDefaults(userData.officeLocation);

    // 2. Get the assigned signature
    var assignmentKey = isReply ? 'reply' : 'newMessage';
    var sigId = prefs.assignments[assignmentKey];
    var signature = getSignatureById(prefs, sigId);

    if (!signature) {
      console.warn('No signature found for assignment:', assignmentKey);
      event.completed();
      return;
    }

    // 3. Merge user data with overrides and company info
    var mergedData = mergeUserData(userData, prefs.overrides, prefs.language);

    // 4. Compose the HTML signature from blocks
    var htmlSignature = await composeSignature(signature, 'htm', mergedData);

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
