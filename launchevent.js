/* SIGNIDENT Spike v0.4 – Mehrere Vorlagen (templates.json) + Zuweisung pro Person (users.json)
   Die Mini-"Rules Engine": users.json -> Feld 'vorlage' bestimmt die Vorlage; ohne Zuweisung gilt 'standard'. */

var SIGNOVA_BASE = "https://signova-app-eta.vercel.app/api/addin/";

/* Zugriffstoken fuer die SIGNIDENT-API. Muss identisch mit der ENV-Variable
   ADDIN_TOKEN im Vercel-Projekt sein, sonst antwortet die API mit 401.
   Hinweis: Das ist Basisschutz, keine Authentifizierung - das Token steht
   hier im Klartext. Vor dem Piloten mit echten Kanzleidaten wird es durch
   Entra ID (Nested App Authentication) ersetzt. */
var SIGNOVA_TOKEN = "hkaVWOSgspki6qXdi2lVUqLtHb9cEzkJB6Tj8YVwtbY";

function signovaFallbackTemplate() {
  return { version: "fallback", firma: "SIGNIDENT Pilot", farbe: "#1F3864", webseite: "",
           logo_url: "", banner_aktiv: false, banner_titel: "", banner_text: "",
           banner_image_url: "", hinweis: "Zentral verwaltet mit SIGNIDENT." };
}

function signovaFetchJson(file, fallback, callback) {
  /* token = Zugriff, v = Cache-Buster (beides als Query-Parameter) */
  var url = SIGNOVA_BASE + file +
            "?token=" + encodeURIComponent(SIGNOVA_TOKEN) +
            "&v=" + Date.now();
  fetch(url)
    .then(function (r) {
      /* 401 = Token fehlt oder ist falsch. Ohne diese Pruefung wuerde die
         Fehlerantwort als gueltiges JSON durchgehen und zu einer leeren
         Signatur fuehren statt zum Fallback. */
      if (!r.ok) throw new Error("SIGNIDENT API " + r.status);
      return r.json();
    })
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

/* Masse, an die sich Add-in UND Dashboard-Vorschau halten muessen.
   Gegenstueck: LOGO_MAX_HEIGHT_PX / BANNER_MAX_WIDTH_PX in signova-app,
   src/lib/signature.ts */
var SIGNOVA_LOGO_MAX_HEIGHT = 40;
var SIGNOVA_BANNER_MAX_WIDTH = 600;

/* Outlook rendert nur absolute Bild-Adressen. Relative Angaben werden
   deshalb gegen den Ursprung der SIGNIDENT-API aufgeloest. */
function signovaAbsoluteUrl(url) {
  var wert = (url || "").trim();
  if (!wert) return "";
  if (/^(https?:)?\/\//i.test(wert) || /^data:/i.test(wert)) return wert;
  try {
    return new URL(wert, SIGNOVA_BASE).href;
  } catch (e) {
    return wert;
  }
}

/* Minimales Escaping fuer Werte, die in Attribute geschrieben werden. */
function signovaAttr(wert) {
  return String(wert || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* Mini-Template-Sprache fuer den HTML-Modus einer Vorlage.
   Gegenstueck: renderHtmlTemplate() in signova-app/src/lib/html-template.ts.
   Beide Seiten muessen zeichengenau dasselbe tun.

     {{feld}}                 Platzhalter
     {{#if feld}}...{{/if}}   Block entfaellt bei leerem Feld

   Erst die Bloecke aufloesen, dann die Platzhalter ersetzen - andersherum
   wuerde ein Wert, der zufaellig {{/if}} enthaelt, die Struktur zerstoeren. */
function signovaRenderHtmlTemplate(html, kontext) {
  var mitBloecken = String(html || "").replace(
    /\{\{#if\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}([\s\S]*?)\{\{\/if\}\}/g,
    function (treffer, feld, inhalt) {
      var wert = kontext[feld];
      return !wert || String(wert).trim() === "" ? "" : inhalt;
    }
  );

  return mitBloecken.replace(
    /\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}/g,
    function (treffer, feld) {
      var wert = kontext[feld];
      /* Unbekannte Platzhalter bleiben stehen - so faellt ein Tippfehler auf,
         statt stillschweigend zu verschwinden. */
      if (wert === undefined) return treffer;
      return signovaAttr(wert);
    }
  );
}

function signovaHtmlKontext(tpl, profile, person) {
  return {
    displayName: profile.displayName,
    titel: person && person.titel ? person.titel : "",
    abteilung: person && person.abteilung ? person.abteilung : "",
    telefon: person && person.telefon ? person.telefon : "",
    mobil: person && person.mobil ? person.mobil : "",
    email: profile.emailAddress,
    firma: tpl.firma || "",
    webseite: tpl.webseite || "",
    logo_url: signovaAbsoluteUrl(tpl.logo_url)
  };
}

/* Banner einer laufenden Kampagne. Hat Vorrang vor dem Banner der Vorlage -
   die Auswahl (Zeitraum, Zielgruppe) trifft der Server, hier kommt nur noch
   ein fertiges Ergebnis an oder gar keins. */
function signovaKampagnenBannerHtml(person, farbe) {
  var k = person && person.kampagnen_banner ? person.kampagnen_banner : null;
  if (!k) return "";

  var bild = signovaAbsoluteUrl(k.image_url);
  if (bild) {
    return '<tr><td style="padding-top:8px;">' +
      '<img src="' + signovaAttr(bild) + '" alt="" style="width:100%; max-width:' +
      SIGNOVA_BANNER_MAX_WIDTH + 'px; height:auto; display:block; border:0;" /></td></tr>';
  }

  if (!k.titel && !k.text) return "";

  return '<tr><td style="padding-top:8px;"><table cellpadding="10" cellspacing="0" style="background:' +
    farbe + '; border-radius:6px;"><tr><td style="color:#ffffff; font-family:Segoe UI, Arial, sans-serif; font-size:9.5pt;">' +
    '<strong>' + (k.titel || "") + '</strong><br/>' + (k.text || "") + '</td></tr></table></td></tr>';
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

  var logo = signovaAbsoluteUrl(tpl.logo_url);
  var bannerBild = signovaAbsoluteUrl(tpl.banner_image_url);
  var kampagne = signovaKampagnenBannerHtml(person, farbe);

  /* HTML-Modus: Die Vorlage IST die Signatur - keine Felder, keine Fussnote.
     Genauso verhaelt sich die Vorschau im Dashboard. Ein laufendes
     Kampagnen-Banner wird darunter angehaengt, weil eine freie HTML-Vorlage
     keinen definierten Banner-Platz hat. */
  if (tpl.mode === "html") {
    var eigen = signovaRenderHtmlTemplate(tpl.html_content, signovaHtmlKontext(tpl, profile, person));
    if (!kampagne) return eigen;
    return eigen +
      '<table cellpadding="0" cellspacing="0" style="max-width:' + SIGNOVA_BANNER_MAX_WIDTH + 'px;">' +
      kampagne + '</table>';
  }

  var html =
    '<table cellpadding="0" cellspacing="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:10pt; color:#222;">';

  /* Logo ueber dem Namen (Ausbaupaket 1) */
  if (logo) {
    html += '<tr><td style="padding-bottom:8px;">' +
      '<img src="' + signovaAttr(logo) + '" alt="" style="max-height:' + SIGNOVA_LOGO_MAX_HEIGHT +
      'px; max-width:100%; display:block; border:0;" /></td></tr>';
  }

  html +=
    '<tr><td style="padding-bottom:6px;"><strong style="font-size:11pt; color:' + farbe + ';">' + profile.displayName + '</strong>' +
    (titel ? '<br/><span style="color:#595959;">' + titel + '</span>' : '') +
    '</td></tr>' +
    '<tr><td style="border-top:2px solid ' + farbe + '; padding-top:6px;">' +
    firmenzeile +
    (telzeile ? '<br/>' + telzeile : '') +
    '<br/>E-Mail: <a href="mailto:' + profile.emailAddress + '" style="color:' + farbe + ';">' + profile.emailAddress + '</a>' +
    (tpl.webseite ? ' · ' + tpl.webseite : '') +
    '</td></tr>';
  /* Banner-Rangfolge: laufende Kampagne schlaegt das Vorlagen-Banner.
     Ohne Kampagne gilt buildBanner() aus signova-app/src/lib/signature.ts:
     Banner aus -> nichts; Bild gesetzt -> Bild; sonst Text-Box, sofern gefuellt. */
  if (kampagne) {
    html += kampagne;
  } else if (tpl.banner_aktiv && bannerBild) {
    html += '<tr><td style="padding-top:8px;">' +
      '<img src="' + signovaAttr(bannerBild) + '" alt="" style="width:100%; max-width:' +
      SIGNOVA_BANNER_MAX_WIDTH + 'px; height:auto; display:block; border:0;" /></td></tr>';
  } else if (tpl.banner_aktiv && (tpl.banner_titel || tpl.banner_text)) {
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
