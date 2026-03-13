// commands.js - Event-based activation handlers for OnNewMessageCompose / OnReplyCompose

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    // Register event handlers
    Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
  }
});

function onNewMessageCompose(event) {
  injectSignature(event, false);
}
