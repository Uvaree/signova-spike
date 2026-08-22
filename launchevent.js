/* SIGNOVA Spike v0.2 – Signatur aus zentraler Vorlage (template.json auf GitHub Pages)
   Automatik-Handler (aktiv bei Admin-Deployment) + gemeinsamer Baukasten fuer den Taskpane-Button. */

var SIGNOVA_TEMPLATE_URL = "https://uvaree.github.io/signova-spike/template.json";

function signovaFallbackTemplate() {
  return { version: "fallback", firma: "SIGNOVA Pilot", farbe: "#1F3864", webseite: "",
           banner_aktiv: false, banner_titel: "", banner_text: "",
           hinweis: "Zentral verwaltet mit SIGNOVA." };
}

function signovaFetchTemplate(callback) {
  fetch(SIGNOVA_TEMPLATE_URL + "?v=" + Date.now())
    .then(function (r) { return r.json(); })
    .then(function (tpl) { callback(tpl); })
    .catch(function () { callback(signovaFallbackTemplate()); });
}

function signovaBuildHtml(tpl, profile) {
  var farbe = tpl.farbe || "#1F3864";
  var html =
    '<table cellpadding="0" cellspacing="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:10pt; color:#222;">' +
    '<tr><td style="padding-bottom:6px;"><strong style="font-size:11pt; color:' + farbe + ';">' + profile.displayName + '</strong><br/>' +
    '<span style="color:#595959;">' + (tpl.firma || "") + '</span></td></tr>' +
    '<tr><td style="border-top:2px solid ' + farbe + '; padding-top:6px;">' +
    'E-Mail: <a href="mailto:' + profile.emailAddress + '" style="color:' + farbe + ';">' + profile.emailAddress + '</a>' +
    (tpl.webseite ? '<br/>' + tpl.webseite : '') +
    '</td></tr>';
  if (tpl.banner_aktiv) {
    html += '<tr><td style="padding-top:8px;"><table cellpadding="10" cellspacing="0" style="background:' + farbe +
      '; border-radius:6px;"><tr><td style="color:#ffffff; font-family:Segoe UI, Arial, sans-serif; font-size:9.5pt;">' +
      '<strong>' + (tpl.banner_titel || "") + '</strong><br/>' + (tpl.banner_text || "") + '</td></tr></table></td></tr>';
  }
  html += '<tr><td style="padding-top:8px; font-size:8pt; color:#8A8A8A;">' + (tpl.hinweis || "") +
    ' &nbsp;(Vorlage: ' + (tpl.version || "?") + ')</td></tr></table>';
  return html;
}

function signovaApply(done) {
  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile;
  signovaFetchTemplate(function (tpl) {
    item.body.setSignatureAsync(
      signovaBuildHtml(tpl, profile),
      { coercionType: Office.CoercionType.Html },
      function (res) { done(res); }
    );
  });
}

/* Automatik: startet bei neuer Mail – erfordert Admin-Deployment (M365 Admin Center) */
function onNewMessageComposeHandler(event) {
  signovaApply(function () { event.completed(); });
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
