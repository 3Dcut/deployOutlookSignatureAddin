// commands.js - Event-based activation handlers for OnNewMessageCompose / OnReplyCompose

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    // Register event handlers
    Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
    Office.actions.associate("onReplyCompose", onReplyCompose);
  }
});

function onNewMessageCompose(event) {
  injectSignature(event, false);
}

function onReplyCompose(event) {
  injectSignature(event, true);
}
