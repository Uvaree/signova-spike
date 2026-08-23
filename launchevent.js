/* SIGNOVA Spike v0.4 – Mehrere Vorlagen (templates.json) + Zuweisung pro Person (users.json)
   Die Mini-"Rules Engine": users.json -> Feld 'vorlage' bestimmt die Vorlage; ohne Zuweisung gilt 'standard'. */

var SIGNOVA_BASE = "https://signova-app-eta.vercel.app/api/addin/";

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

function signovaPickTemplate(templates, person) {
  if (!templates || !templates.vorlagen) return null;
  var wunsch = person && person.vorlage ? person.vorlage : "standard";
  return templates.vorlagen[wunsch] || templates.vorlagen["standard"] || null;
}

function signovaBuildHtml(tpl, profile, person, vorlagenName) {
  var farbe = tpl.farbe || "#1F3864";
  var titel = person && person.titel ? person.titel : "";
  var abteilung = person && person.abteilung ? person.abteilung : "";
  var telefon = person && person.telefon ? person.telefon : "";
  var mobil = person && person.mobil ? person.mobil : "";

  var firmenzeile = (abteilung ? abteilung + " – " : "") + (tpl.firma || "");
  var telzeile = "";
  if (telefon) telzeile += "Tel. " + telefon;
  if (mobil) telzeile += (telzeile ? " · " : "") + "Mobil " + mobil;

  var html =
    '<table cellpadding="0" cellspacing="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:10pt; color:#222;">' +
    '<tr><td style="padding-bottom:6px;"><strong style="font-size:11pt; color:' + farbe + ';">' + profile.displayName + '</strong>' +
    (titel ? '<br/><span style="color:#595959;">' + titel + '</span>' : '') +
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
    ' &nbsp;(' + (vorlagenName ? 'Vorlage: ' + vorlagenName + ' – ' : '') + (tpl.version || "?") +
    (person ? ' · Personendaten: zentral' : ' · Personendaten: nicht gefunden') + ')</td></tr></table>';
  return html;
}

function signovaApply(done) {
  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile;
  signovaFetchJson("templates.json", null, function (templates) {
    signovaFetchJson("users.json", null, function (users) {
      var person = signovaFindUser(users, profile.emailAddress);
      var tpl = signovaPickTemplate(templates, person);
      var name = person && person.vorlage ? person.vorlage : "standard";
      if (!tpl) { tpl = signovaFallbackTemplate(); name = ""; }   /* Notfall: alte/keine templates.json */
      item.body.setSignatureAsync(
        signovaBuildHtml(tpl, profile, person, name),
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
