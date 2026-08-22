/* SIGNOVA Spike v0.3 – Vorlage (template.json) + zentrale Benutzerdaten (users.json = simuliertes Entra ID)
   Neu: Titel, Abteilung, Telefon/Mobil pro Person aus zentraler Liste; leere Felder blenden die Zeile aus. */

var SIGNOVA_BASE = "https://uvaree.github.io/signova-spike/";

function signovaFallbackTemplate() {
  return { version: "fallback", firma: "SIGNOVA Pilot", farbe: "#1F3864", webseite: "",
           banner_aktiv: false, banner_titel: "", banner_text: "",
           hinweis: "Zentral verwaltet mit SIGNOVA." };
}

function signovaFetchJson(file, fallback, callback) {
  fetch(SIGNOVA_BASE + file + "?v=" + Date.now())
    .then(function (r) { return r.json(); })
    .then(function (data) { callback(data); })
    .catch(function () { callback(fallback); });
}

function signovaFindUser(users, email) {
  if (!users || !users.benutzer) return null;
  var mail = (email || "").toLowerCase();
  for (var i = 0; i < users.benutzer.length; i++) {
    if ((users.benutzer[i].email || "").toLowerCase() === mail) return users.benutzer[i];
  }
  return null;
}

function signovaBuildHtml(tpl, profile, person) {
  var farbe = tpl.farbe || "#1F3864";
  var titel = person && person.titel ? person.titel : "";
  var abteilung = person && person.abteilung ? person.abteilung : "";
  var telefon = person && person.telefon ? person.telefon : "";
  var mobil = person && person.mobil ? person.mobil : "";

  var untertitel = titel;                                 /* Conditional: Zeile nur wenn vorhanden */
  var firmenzeile = (abteilung ? abteilung + " – " : "") + (tpl.firma || "");
  var telzeile = "";
  if (telefon) telzeile += "Tel. " + telefon;
  if (mobil) telzeile += (telzeile ? " · " : "") + "Mobil " + mobil;

  var html =
    '<table cellpadding="0" cellspacing="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:10pt; color:#222;">' +
    '<tr><td style="padding-bottom:6px;"><strong style="font-size:11pt; color:' + farbe + ';">' + profile.displayName + '</strong>' +
    (untertitel ? '<br/><span style="color:#595959;">' + untertitel + '</span>' : '') +
    '</td></tr>' +
    '<tr><td style="border-top:2px solid ' + farbe + '; padding-top:6px;">' +
    firmenzeile +
    (telzeile ? '<br/>' + telzeile : '') +
    '<br/>E-Mail: <a href="mailto:' + profile.emailAddress + '" style="color:' + farbe + ';">' + profile.emailAddress + '</a>' +
    (tpl.webseite ? ' · ' + tpl.webseite : '') +
    '</td></tr>';
  if (tpl.banner_aktiv) {
    html += '<tr><td style="padding-top:8px;"><table cellpadding="10" cellspacing="0" style="background:' + farbe +
      '; border-radius:6px;"><tr><td style="color:#ffffff; font-family:Segoe UI, Arial, sans-serif; font-size:9.5pt;">' +
      '<strong>' + (tpl.banner_titel || "") + '</strong><br/>' + (tpl.banner_text || "") + '</td></tr></table></td></tr>';
  }
  html += '<tr><td style="padding-top:8px; font-size:8pt; color:#8A8A8A;">' + (tpl.hinweis || "") +
    ' &nbsp;(Vorlage: ' + (tpl.version || "?") + (person ? ' · Personendaten: zentral' : ' · Personendaten: nicht gefunden') + ')</td></tr></table>';
  return html;
}

function signovaApply(done) {
  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile;
  signovaFetchJson("template.json", signovaFallbackTemplate(), function (tpl) {
    signovaFetchJson("users.json", null, function (users) {
      var person = signovaFindUser(users, profile.emailAddress);
      item.body.setSignatureAsync(
        signovaBuildHtml(tpl, profile, person),
        { coercionType: Office.CoercionType.Html },
        function (res) { done(res); }
      );
    });
  });
}

/* Automatik: startet bei neuer Mail – erfordert Admin-Deployment (M365 Admin Center) */
function onNewMessageComposeHandler(event) {
  signovaApply(function () { event.completed(); });
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
