/* SIGNOVA Signature Spike – Technik-Test
   Setzt beim Verfassen einer neuen E-Mail automatisch eine Signatur.
   Nutzt echte Benutzerdaten aus dem Postfach-Profil (Name, E-Mail).
   Im späteren Produkt kommen jobTitle, Telefon usw. per Microsoft Graph aus Entra ID. */

function buildSignatureHtml(displayName, emailAddress) {
  return (
    '<table cellpadding="0" cellspacing="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:10pt; color:#222222;">' +
    '<tr><td style="padding-bottom:6px;"><strong style="font-size:11pt; color:#1F3864;">' + displayName + '</strong><br/>' +
    '<span style="color:#595959;">Pilotphase &ndash; zentral verwaltete Signatur</span></td></tr>' +
    '<tr><td style="border-top:2px solid #1F3864; padding-top:6px;">' +
    'E-Mail: <a href="mailto:' + emailAddress + '" style="color:#1F3864;">' + emailAddress + '</a><br/>' +
    'Telefon: {{businessPhone}} &nbsp;|&nbsp; Mobil: {{mobilePhone}}<br/>' +
    '<span style="color:#595959;">{{department}} &ndash; {{companyName}}</span>' +
    '</td></tr>' +
    '<tr><td style="padding-top:8px; font-size:8pt; color:#8A8A8A;">' +
    'Diese Signatur wurde automatisch durch SIGNOVA gesetzt (Technik-Test). ' +
    'Die {{Platzhalter}} werden im n&auml;chsten Schritt automatisch aus Entra ID bef&uuml;llt.' +
    '</td></tr>' +
    '</table>'
  );
}

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile;
  var signatureHtml = buildSignatureHtml(profile.displayName, profile.emailAddress);

  item.body.setSignatureAsync(
    signatureHtml,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("SIGNOVA: Signatur konnte nicht gesetzt werden: " + JSON.stringify(asyncResult.error));
      }
      event.completed();
    }
  );
}

// Registrierung des Handlers (erforderlich fuer Outlook Windows classic)
if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
