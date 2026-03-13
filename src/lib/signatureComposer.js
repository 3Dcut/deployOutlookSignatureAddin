// signatureComposer.js - Block-based signature composition with layout marker support

var LAYOUT_LOGO_START = 'layout_logo_start';
var LAYOUT_LOGO_END = 'layout_logo_end';

async function composeSignature(signatureObj, format, userData) {
  // signatureObj: { id, name, language, type, blocks: [{blockId}], customBlocks: [{id, name, htmlContent}] }
  // format: 'htm' or 'txt'
  // userData: merged user data with overrides applied

  var beforeParts = [];
  var rightColumnParts = [];
  var afterParts = [];
  var logoHtml = '';
  var inLogoSection = false;
  var hasLogoSection = false;

  for (var i = 0; i < signatureObj.blocks.length; i++) {
    var blockRef = signatureObj.blocks[i];
    var blockId = blockRef.blockId;

    // Handle layout markers
    if (blockId === LAYOUT_LOGO_START) {
      inLogoSection = true;
      hasLogoSection = true;
      // Load the logo block HTML (contains logo, brand, social icons)
      var logoBlockHtml = await _getBlockContent(blockId, format, signatureObj);
      if (logoBlockHtml) {
        logoHtml = applyPlaceholders(logoBlockHtml, userData);
      }
      continue;
    }

    if (blockId === LAYOUT_LOGO_END) {
      inLogoSection = false;
      continue;
    }

    // Load block content
    var blockContent = await _getBlockContent(blockId, format, signatureObj);
    if (!blockContent) continue;

    var processed = applyPlaceholders(blockContent, userData);

    if (hasLogoSection && inLogoSection) {
      rightColumnParts.push(processed);
    } else if (hasLogoSection && !inLogoSection && rightColumnParts.length > 0) {
      afterParts.push(processed);
    } else if (!hasLogoSection) {
      beforeParts.push(processed);
    } else {
      beforeParts.push(processed);
    }
  }

  // Assemble final output
  if (format === 'htm') {
    return _assembleHtml(beforeParts, logoHtml, rightColumnParts, afterParts, hasLogoSection);
  } else {
    return _assembleText(beforeParts, rightColumnParts, afterParts);
  }
}

async function _getBlockContent(blockId, format, signatureObj) {
  // Custom blocks: read from signatureObj.customBlocks
  if (blockId.indexOf('custom_') === 0) {
    if (format === 'txt') return '';
    return getCustomBlockHtml(signatureObj, blockId);
  }
  // Server blocks: use blockLoader
  return await getBlockHtml(blockId, format);
}

function _assembleHtml(beforeParts, logoHtml, rightColumnParts, afterParts, hasLogoSection) {
  var parts = [];

  // Before parts (e.g. greeting)
  if (beforeParts.length > 0) {
    parts.push(beforeParts.join('\n'));
  }

  // Two-column layout section
  if (hasLogoSection) {
    var twoColumn = '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin-left:10px;">\n';
    twoColumn += '  <tr>\n';
    // Left column: logo/brand/social
    twoColumn += '    <td valign="top" style="width:170px; padding:0 23px 0 0;">\n';
    twoColumn += '      ' + logoHtml + '\n';
    twoColumn += '    </td>\n';
    // Right column: content blocks
    twoColumn += '    <td valign="top" style="padding:0;">\n';
    twoColumn += '      ' + rightColumnParts.join('\n      ') + '\n';
    twoColumn += '    </td>\n';
    twoColumn += '  </tr>\n';
    twoColumn += '</table>';
    parts.push(twoColumn);
  }

  // After parts (e.g. address, legal)
  if (afterParts.length > 0) {
    parts.push(afterParts.join('\n'));
  }

  return '<div style="font-family:\'Arial\',sans-serif;">\n' + parts.join('\n') + '\n</div>';
}

function _assembleText(beforeParts, rightColumnParts, afterParts) {
  var all = beforeParts.concat(rightColumnParts).concat(afterParts);
  return all.join('\n\n');
}

async function injectSignature(event, isReply, triggerSource) {
  try {
    console.log('acadon Signatur: Triggered by', triggerSource);

    // 1. Get current item to check subject for fallback
    var item = Office.context.mailbox.item;

    // Fallback: If not marked as reply but subject looks like one
    if (!isReply) {
      const subjectPromise = new Promise((resolve) => {
        item.subject.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            resolve("");
          }
        });
      });
      
      var subject = await subjectPromise;
      if (subject) {
        var lowerSubj = subject.toLowerCase();
        // Check for common reply/forward prefixes
        if (lowerSubj.indexOf('aw:') === 0 || lowerSubj.indexOf('re:') === 0 || lowerSubj.indexOf('fw:') === 0 || lowerSubj.indexOf('wg:') === 0) {
          console.log('acadon Signatur: Fallback detected reply/forward via subject:', subject);
          isReply = true;
          triggerSource += " (Fallback)";
        }
      }
    }

    // 2. Load user data and preferences
    var userData = await getUserData();
    var prefs = getPreferencesOrDefaults(userData.officeLocation);

    // 2.5. Check if auto-insert is enabled
    if (prefs.autoInsertEnabled === false) {
      console.log('acadon Signatur: Auto-insert is disabled in preferences.');
      event.completed();
      return;
    }

    // 3. Get the assigned signature
    var assignmentKey = isReply ? 'reply' : 'newMessage';
    var sigId = prefs.assignments[assignmentKey];
    var signature = getSignatureById(prefs, sigId);

    if (!signature) {
      console.warn('acadon Signatur: No signature found for assignment:', assignmentKey);
      event.completed();
      return;
    }

    console.log('acadon Signatur: Inserting signature:', signature.name, ' (Source:', triggerSource, ')');

    // 4. Merge user data with overrides and company info (use signature's language)
    var sigLang = signature.language || 'DE';
    var mergedData = mergeUserData(userData, prefs.overrides, sigLang);

    // 5. Compose the HTML signature from blocks
    var htmlSignature = await composeSignature(signature, 'htm', mergedData);

    // 6. Inject via setSignatureAsync
    item.body.setSignatureAsync(
      htmlSignature,
      { coercionType: Office.CoercionType.Html },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error('acadon Signatur: setSignatureAsync failed:', asyncResult.error.message);
        }
        event.completed();
      }
    );
  } catch (err) {
    console.error('acadon Signatur: Injection error:', err.message);
    event.completed();
  }
}
