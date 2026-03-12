// preview-dialog.js - Pop-out preview dialog for signature preview

var receivedHtml = '';

Office.onReady(function() {
  // Register handler for messages from the taskpane
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    onMessageFromParent
  );

  // Button handlers
  document.getElementById('cancel-btn').addEventListener('click', function() {
    Office.context.ui.messageParent(JSON.stringify({ action: 'cancel' }));
  });

  document.getElementById('adopt-btn').addEventListener('click', function() {
    Office.context.ui.messageParent(JSON.stringify({
      action: 'insert',
      html: receivedHtml
    }));
  });
});

function onMessageFromParent(arg) {
  try {
    var message = JSON.parse(arg.message);

    if (message.type === 'preview' && message.html) {
      receivedHtml = message.html;
      renderPreview(message.html);
    }
  } catch (e) {
    document.getElementById('dialog-preview').innerHTML =
      '<p class="info-text">Fehler beim Laden der Vorschau.</p>';
  }
}

function renderPreview(html) {
  var container = document.getElementById('dialog-preview');
  container.innerHTML = '';

  var frame = document.createElement('iframe');
  frame.sandbox = 'allow-same-origin';
  frame.style.width = '100%';
  frame.style.border = 'none';
  frame.style.minHeight = '200px';
  container.appendChild(frame);

  frame.contentDocument.open();
  frame.contentDocument.write(html);
  frame.contentDocument.close();

  frame.onload = function() {
    try {
      var height = frame.contentDocument.body.scrollHeight;
      frame.style.height = (height + 20) + 'px';
    } catch (e) {
      frame.style.height = '600px';
    }
  };
}
