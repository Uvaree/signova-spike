/* SIGNIDENT Spike v0.5 – Vorlagen (templates.json), Zuweisung pro Person
   (users.json), Empfaenger-Regeln und Signatur-Panel.

   Arbeitsteilung mit dem Server (Paket 10):
     - WER welche Vorlage bekommt, entscheidet der Server. Das Add-in kennt
       keine Verzeichnisdaten und keine Regelbedingungen.
     - AN WEN die Mail geht, entscheidet das Add-in. Der Server sieht die
       Empfaenger nie - deshalb liefert er in `empfaenger_regeln` nur die
       Empfaenger-Seite der Regeln mit, die fuer diese Person ueberhaupt in
       Frage kommen. */

var SIGNOVA_BASE = "https://signova-app-eta.vercel.app/api/addin/";

/* Zugriffstoken fuer die SIGNIDENT-API. Muss identisch mit der ENV-Variable
   ADDIN_TOKEN im Vercel-Projekt sein, sonst antwortet die API mit 401.
   Hinweis: Das ist Basisschutz, keine Authentifizierung - das Token steht
   hier im Klartext. Vor dem Piloten mit echten Kanzleidaten wird es durch
   Entra ID (Nested App Authentication) ersetzt. */
/* Version des Add-ins. Wird in der Taskpane angezeigt und hilft beim
   Zuordnen von Fehlerberichten aus der Kanzlei. */
var SIGNOVA_VERSION = "1.3.0";

/* Nach 8 Sekunden ohne Antwort abbrechen. Ohne Timeout haengt das Add-in im
   Zug oder Hotel-WLAN unbegrenzt und der Nutzer sieht nur "laedt". */
var SIGNOVA_TIMEOUT_MS = 8000;

/* ------------------------------------------------------------------
   Nested App Authentication (NAA) - VORBEREITET, STANDARD AUS.

   Solange SIGNOVA_NAA_ENABLED false ist, passiert hier nichts: Das Add-in
   benutzt weiter das statische Token. Das ist Absicht - der laufende Betrieb
   soll sich nicht aendern, bevor ein Tenant angebunden ist.

   Zum Einschalten (Reihenfolge zwingend, siehe README "Microsoft 365
   anbinden"):
     1. In Entra eine App-Registrierung mit SPA-Redirect
        brk-multihub://<domain> anlegen und eine Application-ID-URI vergeben.
     2. ENTRA_ADDIN_AUDIENCE und ENTRA_ADDIN_AUTH_ENABLED in Vercel setzen.
     3. msal-browser einbinden (Skript-Tag in taskpane.html und commands.html)
        und SIGNOVA_MSAL_CLIENT_ID eintragen.
     4. Erst dann SIGNOVA_NAA_ENABLED auf true.
   ------------------------------------------------------------------ */
var SIGNOVA_NAA_ENABLED = false;
var SIGNOVA_MSAL_CLIENT_ID = "";
/* Scope der eigenen API, z.B. "api://<application-id-uri>/Signatures.Read" */
var SIGNOVA_NAA_SCOPE = "";

var signovaMsalInstanz = null;

/* Holt ein Entra-Token still im Hintergrund.
   Gibt null zurueck, wenn NAA aus ist, msal fehlt oder etwas schiefgeht -
   der Aufrufer faellt dann auf das statische Token zurueck. */
function signovaHoleEntraToken(callback) {
  if (!SIGNOVA_NAA_ENABLED || !SIGNOVA_MSAL_CLIENT_ID || !SIGNOVA_NAA_SCOPE) {
    callback(null);
    return;
  }

  /* msal-browser wird per Skript-Tag geladen. Fehlt es, ist das kein Fehler,
     sondern schlicht "NAA nicht verfuegbar". */
  if (typeof msal === "undefined" || !msal.createNestablePublicClientApplication) {
    callback(null);
    return;
  }

  try {
    var weiter = function (instanz) {
      signovaMsalInstanz = instanz;
      instanz
        .acquireTokenSilent({ scopes: [SIGNOVA_NAA_SCOPE] })
        .then(function (ergebnis) {
          callback((ergebnis && ergebnis.accessToken) || null);
        })
        .catch(function () {
          /* Kein interaktiver Login an dieser Stelle: Der Handler laeuft beim
             Verfassen einer Mail, ein Anmeldedialog waere dort unzumutbar.
             Wir fallen still auf das statische Token zurueck. */
          callback(null);
        });
    };

    if (signovaMsalInstanz) {
      weiter(signovaMsalInstanz);
      return;
    }

    msal
      .createNestablePublicClientApplication({
        auth: { clientId: SIGNOVA_MSAL_CLIENT_ID }
      })
      .then(weiter)
      .catch(function () { callback(null); });
  } catch (e) {
    callback(null);
  }
}

var SIGNOVA_TOKEN = "hkaVWOSgspki6qXdi2lVUqLtHb9cEzkJB6Tj8YVwtbY";

function signovaFallbackTemplate() {
  return { version: "fallback", firma: "SIGNIDENT Pilot", farbe: "#1F3864", webseite: "",
           schriftart: "", bloecke: [], logo_url: "", logo_alt: "", banner_aktiv: false, banner_titel: "", banner_text: "",
           banner_image_url: "", banner_alt: "", hinweis: "Zentral verwaltet mit SIGNIDENT.",
           reply_mode: "short", short_html_content: "" };
}

/* Uebersetzt eine technische Ursache in einen Satz, den ein Mensch versteht. */
function signovaFehlertext(ursache) {
  if (ursache === "timeout") {
    return "Zeitueberschreitung: Die SIGNIDENT-Plattform hat nicht innerhalb von " +
      (SIGNOVA_TIMEOUT_MS / 1000) + " Sekunden geantwortet. Bitte Netzwerkverbindung pruefen.";
  }
  if (ursache === "offline") {
    return "Keine Netzwerkverbindung. Die Signatur kann erst gesetzt werden, wenn Sie wieder online sind.";
  }
  if (ursache === "401") {
    return "Zugriff verweigert (401). Das Add-in-Token stimmt nicht mehr mit der Plattform ueberein - bitte die IT informieren.";
  }
  if (/^5\d\d$/.test(String(ursache))) {
    return "Die SIGNIDENT-Plattform meldet einen Serverfehler (" + ursache + "). Bitte spaeter erneut versuchen.";
  }
  if (ursache) {
    return "Die SIGNIDENT-Plattform hat mit Status " + ursache + " geantwortet.";
  }
  return "Die SIGNIDENT-Plattform ist nicht erreichbar.";
}

function signovaFetchJson(file, fallback, callback) {
  signovaHoleEntraToken(function (entraToken) {
    signovaFetchMitToken(file, entraToken || SIGNOVA_TOKEN, Boolean(entraToken), fallback, callback);
  });
}

function signovaFetchMitToken(file, token, alsHeader, fallback, callback) {
  /* Ein Entra-Token gehoert in den Authorization-Header, nicht in die URL:
     Query-Parameter landen in Server- und Proxy-Logs. Das statische Token
     bleibt aus Kompatibilitaetsgruenden im Query-Parameter. */
  var url = SIGNOVA_BASE + file +
            (alsHeader ? "?" : "?token=" + encodeURIComponent(token) + "&") +
            "v=" + Date.now();
  var optionen = alsHeader ? { headers: { Authorization: "Bearer " + token } } : undefined;

  /* Offline gar nicht erst versuchen - so kommt sofort eine verstaendliche
     Meldung statt eines Timeouts nach acht Sekunden. */
  if (typeof navigator !== "undefined" && navigator.onLine === false) {
    callback(fallback, "offline");
    return;
  }

  var erledigt = false;
  var uhr = setTimeout(function () {
    if (erledigt) return;
    erledigt = true;
    callback(fallback, "timeout");
  }, SIGNOVA_TIMEOUT_MS);

  fetch(url, optionen)
    .then(function (r) {
      /* 401 = Token fehlt oder ist falsch. Ohne diese Pruefung wuerde die
         Fehlerantwort als gueltiges JSON durchgehen und zu einer leeren
         Signatur fuehren statt zum Fallback. */
      if (!r.ok) throw new Error(String(r.status));
      return r.json();
    })
    .then(function (data) {
      if (erledigt) return;
      erledigt = true;
      clearTimeout(uhr);
      callback(data, null);
    })
    .catch(function (e) {
      if (erledigt) return;
      erledigt = true;
      clearTimeout(uhr);
      callback(fallback, (e && e.message) || "netzwerk");
    });
}

function signovaFindUser(users, email) {
  if (!users || !users.benutzer) return null;
  var mail = (email || "").toLowerCase();
  for (var i = 0; i < users.benutzer.length; i++) {
    if ((users.benutzer[i].email || "").toLowerCase() === mail) return users.benutzer[i];
  }
  return null;
}

/* Beschriftungen je Sprache. Gegenstueck: BESCHRIFTUNGEN in
   signova-app/src/lib/sprachen.ts - beide Seiten muessen gleich bleiben. */
var SIGNOVA_BESCHRIFTUNGEN = {
  de: { telefon: "Tel.", mobil: "Mobil", email: "E-Mail", vorlage: "Vorlage" },
  en: { telefon: "Phone", mobil: "Mobile", email: "Email", vorlage: "Template" }
};

function signovaSprache(person) {
  return person && person.sprache === "en" ? "en" : "de";
}

function signovaBeschriftungen(person) {
  return SIGNOVA_BESCHRIFTUNGEN[signovaSprache(person)] || SIGNOVA_BESCHRIFTUNGEN.de;
}

function signovaPickTemplate(templates, person) {
  /* Testbenutzer bekommen den Entwurf ihrer eigenen Vorlage direkt in
     users.json mitgeliefert. Die Auswahl trifft der Server - ein normaler
     Benutzer bekommt dieses Feld nie zu sehen. */
  if (person && person.vorlage_entwurf) return person.vorlage_entwurf;

  /* Ab Ausbaupaket 6 liefert der Server die effektive Vorlage fertig mit -
     inklusive bereits ersetzter Textbausteine in der Sprache der Person.
     Damit braucht das Add-in keine Baustein-Logik. Fehlt das Feld (aeltere
     API), greift der bisherige Weg ueber templates.json. */
  if (person && person.vorlage_aufgeloest) return person.vorlage_aufgeloest;

  if (!templates || !templates.vorlagen) return null;
  var wunsch = person && person.vorlage ? person.vorlage : "standard";
  return templates.vorlagen[wunsch] || templates.vorlagen["standard"] || null;
}

/* Kennzeichnung der eigenen Signatur.

   setSignatureAsync ERSETZT den Signaturbereich des Entwurfs - eine zweite
   Signatur kann dadurch gar nicht entstehen, egal wie oft gesetzt wird. Die
   Marke dient deshalb nicht dem Ersetzen selbst, sondern dem Wiedererkennen:
   Das Panel kann damit sagen, ob im Entwurf bereits eine SIGNIDENT-Signatur
   steht, und ein Blick in den Quelltext einer Mail zeigt, aus welcher
   Add-in-Version sie stammt. */
var SIGNOVA_MARKE = "signident-signatur";

function signovaMarkiere(html) {
  if (!html) return "";
  return '<div id="' + SIGNOVA_MARKE + '" data-signident="' + SIGNOVA_VERSION + '">' +
    html + '</div>';
}

/* Steht im Entwurf schon eine Signatur von uns? Rein zur Anzeige - fuer das
   Setzen selbst ist die Antwort ohne Belang (siehe oben). */
function signovaHatEigeneSignatur(item, callback) {
  try {
    if (!item || !item.body || typeof item.body.getAsync !== "function") {
      callback(false);
      return;
    }
    item.body.getAsync(Office.CoercionType.Html, function (res) {
      var html = res && res.status === Office.AsyncResultStatus.Succeeded ? res.value : "";
      callback(String(html || "").indexOf(SIGNOVA_MARKE) >= 0);
    });
  } catch (e) {
    callback(false);
  }
}

/* Ist die gerade verwendete Vorlage ein unveroeffentlichter Entwurf? */
function signovaIstEntwurf(person) {
  return Boolean(person && person.vorlage_entwurf);
}

/* Sichtbare Kennzeichnung, damit ein Entwurf im Postausgang nicht mit dem
   Live-Stand verwechselt wird. */
function signovaEntwurfHinweisHtml() {
  return '<tr><td style="padding-top:8px;">' +
    '<span style="display:inline-block; padding:3px 8px; border-radius:4px; ' +
    'background:#FEF3C7; color:#92400E; font-family:Segoe UI, Arial, sans-serif; ' +
    'font-size:8pt; font-weight:600;">Entwurf – nur für Testbenutzer sichtbar</span>' +
    '</td></tr>';
}

/* Masse, an die sich Add-in UND Dashboard-Vorschau halten muessen.
   Gegenstueck: LOGO_MAX_HEIGHT_PX / BANNER_MAX_WIDTH_PX in signova-app,
   src/lib/signature.ts */
/* Schriftarten. Spiegel von signova-app/src/lib/schriftarten.ts - wer dort
   etwas ergaenzt, muss es hier mitziehen, sonst faellt die betroffene
   Vorlage im Add-in stillschweigend auf die Vorgabe zurueck.
   Leere oder unbekannte Kennung = Vorgabe, damit bestehende Vorlagen
   unveraendert aussehen. */
var SIGNOVA_STANDARD_SCHRIFT = "'Segoe UI', Arial, sans-serif";
var SIGNOVA_SCHRIFTEN = {
  segoe:     "'Segoe UI', Arial, sans-serif",
  arial:     "Arial, Helvetica, sans-serif",
  helvetica: "Helvetica, Arial, sans-serif",
  verdana:   "Verdana, Geneva, sans-serif",
  tahoma:    "Tahoma, Geneva, sans-serif",
  trebuchet: "'Trebuchet MS', Tahoma, sans-serif",
  georgia:   "Georgia, 'Times New Roman', serif",
  times:     "'Times New Roman', Times, serif",
  garamond:  "Garamond, Georgia, 'Times New Roman', serif"
};

function signovaSchrift(tpl) {
  var id = tpl && tpl.schriftart ? String(tpl.schriftart) : "";
  return SIGNOVA_SCHRIFTEN[id] || SIGNOVA_STANDARD_SCHRIFT;
}

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
/* Loest {{#if feld}}...{{/if}} auf - auch INEINANDER verschachtelt.

   Spiegel von loeseBedingungen() in signova-app/src/lib/html-template.ts.
   Der frueher hier stehende regulaere Ausdruck schloss beim ersten
   {{/if}} und liess bei verschachtelten Bedingungen ein {{/if}} als Text
   in der Signatur stehen. Der visuelle Baukasten erzeugt Verschachtelung
   zwangslaeufig (bedingter Block mit Social-Links darin). */
function signovaLoeseBedingungen(html, kontext) {
  var OEFFNEN = /\{\{#if\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}/g;
  var SCHLIESSEN = "{{/if}}";
  var stapel = [{ feld: null, teile: [] }];
  var position = 0;

  function anhaengen(text) {
    if (text) stapel[stapel.length - 1].teile.push(text);
  }
  function leer(wert) {
    return !wert || String(wert).trim() === "";
  }

  while (position < html.length) {
    OEFFNEN.lastIndex = position;
    var auf = OEFFNEN.exec(html);
    var zu = html.indexOf(SCHLIESSEN, position);
    if (!auf && zu === -1) break;

    var aufAn = auf ? auf.index : Number.MAX_SAFE_INTEGER;
    var zuAn = zu === -1 ? Number.MAX_SAFE_INTEGER : zu;

    if (aufAn < zuAn && auf) {
      anhaengen(html.slice(position, auf.index));
      stapel.push({ feld: auf[1], teile: [] });
      position = auf.index + auf[0].length;
      continue;
    }

    anhaengen(html.slice(position, zuAn));
    position = zuAn + SCHLIESSEN.length;
    if (stapel.length === 1) continue;

    var block = stapel.pop();
    anhaengen(leer(kontext[block.feld]) ? "" : block.teile.join(""));
  }

  anhaengen(html.slice(position));

  while (stapel.length > 1) {
    var offen = stapel.pop();
    stapel[stapel.length - 1].teile.push(
      leer(kontext[offen.feld]) ? "" : offen.teile.join("")
    );
  }

  return stapel[0].teile.join("");
}

function signovaRenderHtmlTemplate(html, kontext) {
  var mitBloecken = signovaLoeseBedingungen(String(html || ""), kontext);

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
    logo_url: signovaAbsoluteUrl(tpl.logo_url),
    /* Spiegel von signova-app/src/lib/html-template.ts - dort ebenfalls
       zuletzt in der Liste PLATZHALTER. */
    farbe: tpl.farbe || ""
  };
}

/* ==================================================================
   Visueller Baukasten

   Spiegel von signova-app/src/lib/baukasten.ts. Wer dort einen
   Blocktyp ergaenzt, muss ihn HIER mitziehen - sonst verschwindet der
   Block in der echten Signatur stillschweigend, waehrend die Vorschau
   im Dashboard ihn zeigt.

   Ausgabe ist bewusst tabellenbasiert mit Inline-Stilen: Outlook auf
   Windows rendert mit der Word-Engine, dort gibt es kein Flexbox, kein
   Grid und keine externen Stilangaben.
   ================================================================== */

var SIGNOVA_AUSRICHTUNG = { links: "left", mitte: "center", rechts: "right" };

function signovaBlockStil(stil, kontext) {
  stil = stil || {};
  var teile = [];
  var oben = stil.abstandOben || 0;
  var unten = stil.abstandUnten || 0;
  if (oben || unten) teile.push("padding:" + oben + "px 0 " + unten + "px 0");

  teile.push(
    "font-family:" +
      (stil.schriftart
        ? SIGNOVA_SCHRIFTEN[stil.schriftart] || SIGNOVA_STANDARD_SCHRIFT
        : kontext.schriftStack)
  );
  teile.push("font-size:" + (stil.groesse || 10.5) + "pt");
  teile.push("color:" + (stil.farbe || "#333333"));
  if (stil.fett) teile.push("font-weight:bold");
  if (stil.kursiv) teile.push("font-style:italic");
  if (stil.ausrichtung) teile.push("text-align:" + SIGNOVA_AUSRICHTUNG[stil.ausrichtung]);
  teile.push("line-height:1.45");
  return teile.join("; ");
}

/* Nur http(s) und mailto - verhindert javascript: in einer Signatur. */
function signovaSichereUrl(roh) {
  var wert = String(roh || "").trim();
  if (!wert) return "";
  if (/^(https?:|mailto:)/i.test(wert)) return wert;
  if (/^[\w.-]+\.[a-z]{2,}(\/|$)/i.test(wert)) return "https://" + wert;
  return "";
}

function signovaBlockZeile(block, kontext) {
  var inhalt = signovaBlockInhalt(block, kontext);
  if (!inhalt) return "";
  return '<tr><td style="' + signovaBlockStil(block.stil, kontext) + '">' + inhalt + "</td></tr>";
}

function signovaBlockInhalt(block, kontext) {
  var farbe = (block.stil && block.stil.farbe) || kontext.farbe;
  var breite;

  switch (block.typ) {
    case "text":
      return String(block.text || "").split("\n").join("<br/>");

    case "platzhalter":
      return block.feld ? "{{" + block.feld + "}}" : "";

    case "bedingung": {
      if (!block.wennFeld) return "";
      var innen = (block.kinder || [])
        .map(function (kind) { return signovaBlockZeile(kind, kontext); })
        .filter(Boolean)
        .join("");
      if (!innen) return "";
      return (
        "{{#if " + block.wennFeld + "}}" +
        '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' +
        innen + "</table>{{/if}}"
      );
    }

    case "bild": {
      var url = signovaSichereUrl(block.url);
      if (!url) return "";
      breite = block.breite || 200;
      return (
        '<img src="' + signovaAttr(url) + '" alt="' + signovaAttr(block.alt || "") +
        '" width="' + breite + '" style="display:block; border:0; width:' + breite +
        'px; max-width:100%; height:auto;" />'
      );
    }

    case "logo":
      breite = block.breite || 160;
      return (
        '{{#if logo_url}}<img src="{{logo_url}}" alt="{{firma}}" width="' + breite +
        '" style="display:block; border:0; width:' + breite +
        'px; max-width:100%; height:auto;" />{{/if}}'
      );

    case "tabelle": {
      var links = (block.links || []).map(function (k) { return signovaBlockZeile(k, kontext); }).join("");
      var rechts = (block.rechts || []).map(function (k) { return signovaBlockZeile(k, kontext); }).join("");
      if (!links && !rechts) return "";
      return (
        '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse; width:100%;"><tr>' +
        '<td width="50%" valign="top" style="width:50%; vertical-align:top; padding-right:12px;">' +
        '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' + links + "</table></td>" +
        '<td width="50%" valign="top" style="width:50%; vertical-align:top;">' +
        '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' + rechts + "</table></td>" +
        "</tr></table>"
      );
    }

    case "link": {
      var ziel = signovaSichereUrl(block.url);
      if (!ziel) return "";
      var text = (block.beschriftung && block.beschriftung.trim()) || ziel;
      return '<a href="' + signovaAttr(ziel) + '" style="color:' + farbe + '; text-decoration:underline;">' + text + "</a>";
    }

    case "social": {
      var link = function (feld, name) {
        return '<a href="{{' + feld + '}}" style="color:' + farbe + '; text-decoration:none;">' + name + "</a>";
      };
      var trenner = '<span style="color:#B5AFA6;"> · </span>';
      return (
        "{{#if linkedin_url}}" + link("linkedin_url", "LinkedIn") + "{{/if}}" +
        "{{#if xing_url}}{{#if linkedin_url}}" + trenner + "{{/if}}" + link("xing_url", "XING") + "{{/if}}"
      );
    }

    case "meeting": {
      var beschriftung = (block.beschriftung && block.beschriftung.trim()) || "Termin vereinbaren";
      return (
        '{{#if meeting_url}}<a href="{{meeting_url}}" style="color:' + farbe +
        '; text-decoration:underline;">' + beschriftung + "</a>{{/if}}"
      );
    }

    case "qr": {
      /* Die Adresse kommt fertig vom Server (users.json -> qr_url). Ein
         data:-URI funktioniert in Outlook auf Windows nicht, und ein
         fremder QR-Dienst wuerde melden, wer die Mail oeffnet. */
      if (!kontext.qrUrl) return "";
      breite = block.breite || 96;
      return (
        '<img src="' + signovaAttr(kontext.qrUrl) + '" alt="' +
        signovaAttr(block.alt || "Kontaktdaten als QR-Code") + '" width="' + breite +
        '" height="' + breite + '" style="display:block; border:0; width:' + breite +
        "px; height:" + breite + 'px;" />'
      );
    }

    case "trenner": {
      var linienfarbe = (block.stil && block.stil.farbe) || "#E9E5E0";
      return (
        '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse; width:100%;"><tr>' +
        '<td style="height:1px; line-height:1px; font-size:0; background:' + linienfarbe + ';">&nbsp;</td>' +
        "</tr></table>"
      );
    }

    case "banner":
      return "";

    case "baustein":
      return block.bausteinKey ? "{{baustein:" + block.bausteinKey + "}}" : "";

    case "html":
      return block.text || "";

    default:
      return "";
  }
}

function signovaBaukastenHtml(bloecke, kontext) {
  if (!bloecke || !bloecke.length) return "";
  var zeilen = bloecke
    .map(function (block) { return signovaBlockZeile(block, kontext); })
    .filter(Boolean)
    .join("");
  if (!zeilen) return "";
  return (
    '<table cellpadding="0" cellspacing="0" border="0" role="presentation" ' +
    'style="border-collapse:collapse; max-width:600px;">' + zeilen + "</table>"
  );
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
      '<img src="' + signovaAttr(bild) + '" alt="' + signovaAttr(k.alt || "") + '" style="width:100%; max-width:' +
      SIGNOVA_BANNER_MAX_WIDTH + 'px; height:auto; display:block; border:0;" /></td></tr>';
  }

  if (!k.titel && !k.text) return "";

  return '<tr><td style="padding-top:8px;"><table cellpadding="10" cellspacing="0" style="background:' +
    farbe + '; border-radius:6px;"><tr><td style="color:#ffffff; font-family:Segoe UI, Arial, sans-serif; font-size:9.5pt;">' +
    '<strong>' + (k.titel || "") + '</strong><br/>' + (k.text || "") + '</td></tr></table></td></tr>';
}

/* Kurzform fuer Antworten und Weiterleitungen.
   Gegenstueck: buildShortSignature() in signident/src/lib/signature.ts.
   Bewusst reduziert - kein Logo, kein Banner, keine Kampagne. */
function signovaBuildShortHtml(tpl, profile, person) {
  var farbe = tpl.farbe || "#1F3864";
  var titel = person && person.titel ? person.titel : "";
  var telefon = person && person.telefon ? person.telefon : "";
  var mobil = person && person.mobil ? person.mobil : "";

  /* Eigenes HTML fuer die Kurzform hat Vorrang - dieselbe
     Platzhalter-Maschine wie im HTML-Modus, kein zweiter Dialekt. */
  var eigen = tpl.short_html_content ? String(tpl.short_html_content).trim() : "";
  if (eigen) {
    return signovaRenderHtmlTemplate(eigen, signovaHtmlKontext(tpl, profile, person));
  }

  var telzeile = "";
  var l = signovaBeschriftungen(person);
  if (telefon) telzeile += l.telefon + " " + telefon;
  if (mobil) telzeile += (telzeile ? " · " : "") + l.mobil + " " + mobil;
  return '<table cellpadding="0" cellspacing="0" style="font-family:' + signovaSchrift(tpl) + '; font-size:10pt; color:#222;">' +
    '<tr><td><strong style="color:' + farbe + ';">' + profile.displayName + '</strong>' +
    (titel ? '<br/><span style="color:#595959;">' + titel + '</span>' : '') +
    (telzeile ? '<br/>' + telzeile : '') +
    '<br/>' + profile.emailAddress +
    '</td></tr>' +
    (signovaIstEntwurf(person) ? signovaEntwurfHinweisHtml() : '') +
    '</table>';
}

/* ==================================================================
   Empfaenger-Regeln - CLIENTSEITIG.

   Spiegel von pruefeEmpfaenger() in signova-app/src/lib/rules.ts. Beide
   Seiten muessen dieselbe Antwort geben, sonst zeigt der Signaturtester in
   der Verwaltung etwas anderes als Outlook.

   Der Server hat bereits entschieden, welche Regeln fuer DIESE PERSON
   ueberhaupt in Frage kommen; hier steht nur noch die Empfaenger-Seite.
   ================================================================== */

function signovaDomainVon(adresse) {
  var teile = String(adresse || "").trim().toLowerCase().split("@");
  return teile.length > 1 ? teile.pop() : "";
}

/* '*' steht fuer beliebig viele Zeichen. Enthaelt das Muster ein '@', wird
   die ganze Adresse verglichen statt nur der Domain. */
function signovaPasstMuster(adresse, muster) {
  var wert = String(muster || "").trim().toLowerCase();
  if (!wert) return false;

  var ziel = wert.indexOf("@") >= 0
    ? String(adresse || "").trim().toLowerCase()
    : signovaDomainVon(adresse);
  if (!ziel) return false;

  /* Alles maskieren, was in einem regulaeren Ausdruck Bedeutung hat - nur der
     Stern bleibt als Platzhalter uebrig. Spiegel von passtDomainMuster(). */
  var regex = wert.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\\\*/g, ".*");
  try {
    return new RegExp("^" + regex + "$").test(ziel);
  } catch (e) {
    return false;
  }
}

function signovaTrifftEinMuster(adressen, muster) {
  for (var i = 0; i < adressen.length; i++) {
    for (var j = 0; j < muster.length; j++) {
      if (signovaPasstMuster(adressen[i], muster[j])) return adressen[i];
    }
  }
  return null;
}

function signovaEmpfaengerTrifft(regel, adressen, eigeneDomain) {
  var ein = regel.einschluss || [];
  var aus = regel.ausschluss || [];
  var modus = regel.modus || "alle";
  var ohneBedingung = modus === "alle" && ein.length === 0 && aus.length === 0;

  /* Ohne eingetragene Empfaenger kann nur eine Regel greifen, die gar nichts
     ueber sie aussagt - genau so entscheidet auch der Server. */
  if (adressen.length === 0) return ohneBedingung;

  /* Ausschluss zuerst und ohne Gegenrede: Wer "nie an *.gov" einstellt,
     meint eine harte Grenze, keine Stimme unter mehreren. */
  if (aus.length && signovaTrifftEinMuster(adressen, aus)) return false;
  if (ein.length && !signovaTrifftEinMuster(adressen, ein)) return false;

  if (modus === "intern" || modus === "extern") {
    var fremd = false;
    for (var i = 0; i < adressen.length; i++) {
      if (signovaDomainVon(adressen[i]) !== String(eigeneDomain || "").toLowerCase()) {
        fremd = true;
        break;
      }
    }
    return modus === "intern" ? !fremd : fremd;
  }

  return true;
}

/* Welche Vorlage gilt bei DIESEN Empfaengern?
   null = keine Empfaenger-Regel greift; dann bleibt es bei der Vorlage, die
   der Server ohnehin ausgeliefert hat. */
function signovaVorlageNachEmpfaengern(person, adressen) {
  var regeln = person && person.empfaenger_regeln;
  if (!regeln || !regeln.length || !person.vorlagen) return null;

  var treffer = null;
  for (var i = 0; i < regeln.length; i++) {
    if (!signovaEmpfaengerTrifft(regeln[i], adressen, person.absender_domain)) continue;
    treffer = regeln[i];
    /* "Nach Treffer stoppen" ist die Vorgabe; ohne sie darf eine spaetere
       Regel ueberschreiben - dieselbe Logik wie werteRegelnAus(). */
    if (treffer.stopp) break;
  }

  if (!treffer) return null;
  var tpl = person.vorlagen[treffer.vorlage_id];
  return tpl ? { tpl: tpl, name: treffer.template_key } : null;
}

/* Liest An und Cc aus dem Entwurf. Im Lesemodus gibt es diese Objekte nicht -
   dann bleibt die Liste leer und es gilt die Vorlage des Servers. */
function signovaEmpfaenger(item, callback) {
  var adressen = [];

  function adressenAus(res) {
    if (!res || res.status !== Office.AsyncResultStatus.Succeeded || !res.value) return;
    for (var i = 0; i < res.value.length; i++) {
      var eintrag = res.value[i];
      if (eintrag && eintrag.emailAddress) adressen.push(eintrag.emailAddress);
    }
  }

  try {
    if (!item || !item.to || typeof item.to.getAsync !== "function") {
      callback([]);
      return;
    }

    item.to.getAsync(function (res) {
      adressenAus(res);
      if (item.cc && typeof item.cc.getAsync === "function") {
        item.cc.getAsync(function (res2) {
          adressenAus(res2);
          callback(adressen);
        });
      } else {
        callback(adressen);
      }
    });
  } catch (e) {
    callback([]);
  }
}

/* Gilt fuer diesen Verfassen-Typ die Kurzform?
   Gegenstueck: verwendeKurzform() in signident/src/lib/signature.ts. */
function signovaVerwendeKurzform(tpl, composeTyp) {
  return composeTyp === "antwort" && tpl.reply_mode === "short";
}

/* Gar keine Signatur?

   reply_mode 'first_only' heisst "nur in der ersten Mail eines Verlaufs".
   Erkannt wird das allein am Compose-Typ: Outlook meldet reply/forward,
   sobald eine Nachricht aus einer anderen heraus entsteht. Ein echtes "ist
   das der erste Beitrag dieses Threads?" gaebe es nur, indem das Add-in das
   Postfach durchsieht - genau das soll es nicht.

   Gegenstueck: unterdrueckeSignatur() in signova-app/src/lib/signature.ts. */
function signovaUnterdrueckt(tpl, composeTyp) {
  return composeTyp === "antwort" && tpl && tpl.reply_mode === "first_only";
}

/* Ermittelt, ob gerade eine neue Mail oder eine Antwort/Weiterleitung
   verfasst wird. getComposeTypeAsync gibt es erst ab Mailbox 1.10 und nur
   im Compose-Modus - faellt es aus, gilt "neu". Eine falsche Kurzform waere
   schlimmer als eine volle Signatur an der falschen Stelle. */
function signovaComposeTyp(item, callback) {
  try {
    if (!item || typeof item.getComposeTypeAsync !== "function") {
      callback("neu");
      return;
    }
    item.getComposeTypeAsync(function (res) {
      if (res && res.status === Office.AsyncResultStatus.Succeeded && res.value) {
        var typ = res.value.composeType;
        callback(typ === "reply" || typ === "forward" ? "antwort" : "neu");
      } else {
        callback("neu");
      }
    });
  } catch (e) {
    callback("neu");
  }
}

/* Telemetrie: meldet nach erfolgreichem Setzen, WER die Signatur mit WELCHEM
   CLIENT und WELCHER VERSION bekommen hat. Fire-and-forget - die Antwort wird
   nicht abgewartet und Fehler werden verschluckt. Ein Ausfall der Telemetrie
   darf niemals eine Signatur verhindern. */
function signovaClientKennung() {
  var host = "Outlook";
  var plattform = "";

  /* Beide Zugriffe einzeln absichern: hostName liegt unter mailbox.diagnostics,
     platform unter Office.context.diagnostics. Je nach Outlook-Version und
     Requirement-Set kann das eine da sein und das andere fehlen. */
  try {
    host = Office.context.mailbox.diagnostics.hostName || host;
  } catch (e) { /* Standardwert behalten */ }

  try {
    plattform = Office.context.diagnostics.platform || "";
  } catch (e) { /* ohne Plattform weitermachen */ }

  return plattform ? host + " / " + plattform : host;
}

function signovaPing(profile, tpl, vorlagenName, modus) {
  try {
    fetch(SIGNOVA_BASE + "ping?token=" + encodeURIComponent(SIGNOVA_TOKEN), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email: profile.emailAddress,
        client: signovaClientKennung(),
        template_key: vorlagenName || "",
        template_version: (tpl && tpl.version) || "",
        mode: modus
      })
    })["catch"](function () { /* bewusst ignoriert */ });
  } catch (e) { /* bewusst ignoriert */ }
}

function signovaBuildHtml(tpl, profile, person, vorlagenName, composeTyp) {
  var farbe = tpl.farbe || "#1F3864";
  var titel = person && person.titel ? person.titel : "";
  var abteilung = person && person.abteilung ? person.abteilung : "";
  var telefon = person && person.telefon ? person.telefon : "";
  var mobil = person && person.mobil ? person.mobil : "";

  var firmenzeile = (abteilung ? abteilung + " – " : "") + (tpl.firma || "");
  var telzeile = "";
  var l = signovaBeschriftungen(person);
  if (telefon) telzeile += l.telefon + " " + telefon;
  if (mobil) telzeile += (telzeile ? " · " : "") + l.mobil + " " + mobil;

  var logo = signovaAbsoluteUrl(tpl.logo_url);
  var bannerBild = signovaAbsoluteUrl(tpl.banner_image_url);
  var kampagne = signovaKampagnenBannerHtml(person, farbe);

  /* Antwort/Weiterleitung mit reply_mode 'short': Kurzform, sonst nichts.
     Auch die Kampagne entfaellt hier bewusst. */
  if (signovaVerwendeKurzform(tpl, composeTyp)) {
    return signovaBuildShortHtml(tpl, profile, person);
  }

  /* Baukasten-Modus: Die Bloecke ergeben HTML MIT Platzhaltern, das
     anschliessend durch dieselbe Maschine laeuft wie eine handgeschriebene
     HTML-Vorlage. Kein zweiter Dialekt, keine zweite Maskierung.
     Banner bzw. laufende Kampagne haengen darunter - genau wie im
     HTML-Modus, wo es ebenfalls keinen festen Bannerplatz gibt. */
  if (tpl.mode === "baukasten") {
    var geruest = signovaBaukastenHtml(tpl.bloecke || [], {
      farbe: farbe,
      schriftStack: signovaSchrift(tpl),
      qrUrl: person && person.qr_url ? person.qr_url : ""
    });
    var gebaut = signovaRenderHtmlTemplate(geruest, signovaHtmlKontext(tpl, profile, person));

    var bannerZeile = kampagne;
    if (!bannerZeile && tpl.banner_aktiv && bannerBild) {
      bannerZeile = '<tr><td style="padding-top:12px;">' +
        '<img src="' + signovaAttr(bannerBild) + '" alt="' + signovaAttr(tpl.banner_alt || "") +
        '" style="width:100%; max-width:' + SIGNOVA_BANNER_MAX_WIDTH +
        'px; height:auto; display:block; border:0;" /></td></tr>';
    } else if (!bannerZeile && tpl.banner_aktiv && (tpl.banner_titel || tpl.banner_text)) {
      bannerZeile = '<tr><td style="padding-top:12px;"><table cellpadding="10" cellspacing="0" style="background:' +
        farbe + '; border-radius:4px;"><tr><td style="color:#ffffff; font-family:Segoe UI, Arial, sans-serif;">' +
        (tpl.banner_titel ? '<strong style="font-size:10.5pt;">' + tpl.banner_titel + '</strong>' : '') +
        (tpl.banner_text ? '<br/><span style="font-size:9.5pt;">' + tpl.banner_text + '</span>' : '') +
        '</td></tr></table></td></tr>';
    }

    var anhangB = bannerZeile + (signovaIstEntwurf(person) ? signovaEntwurfHinweisHtml() : '');
    if (!anhangB) return gebaut;
    return gebaut +
      '<table cellpadding="0" cellspacing="0" style="max-width:' + SIGNOVA_BANNER_MAX_WIDTH + 'px;">' +
      anhangB + '</table>';
  }

  /* HTML-Modus: Die Vorlage IST die Signatur - keine Felder, keine Fussnote.
     Genauso verhaelt sich die Vorschau im Dashboard. Ein laufendes
     Kampagnen-Banner wird darunter angehaengt, weil eine freie HTML-Vorlage
     keinen definierten Banner-Platz hat. */
  if (tpl.mode === "html") {
    var eigen = signovaRenderHtmlTemplate(tpl.html_content, signovaHtmlKontext(tpl, profile, person));
    var anhang = kampagne + (signovaIstEntwurf(person) ? signovaEntwurfHinweisHtml() : '');
    if (!anhang) return eigen;
    return eigen +
      '<table cellpadding="0" cellspacing="0" style="max-width:' + SIGNOVA_BANNER_MAX_WIDTH + 'px;">' +
      anhang + '</table>';
  }

  var html =
    '<table cellpadding="0" cellspacing="0" style="font-family:' + signovaSchrift(tpl) + '; font-size:10pt; color:#222;">';

  /* Logo ueber dem Namen (Ausbaupaket 1) */
  if (logo) {
    html += '<tr><td style="padding-bottom:8px;">' +
      '<img src="' + signovaAttr(logo) + '" alt="' + signovaAttr(tpl.logo_alt || "") + '" style="max-height:' + SIGNOVA_LOGO_MAX_HEIGHT +
      'px; max-width:100%; display:block; border:0;" /></td></tr>';
  }

  html +=
    '<tr><td style="padding-bottom:6px;"><strong style="font-size:11pt; color:' + farbe + ';">' + profile.displayName + '</strong>' +
    (titel ? '<br/><span style="color:#595959;">' + titel + '</span>' : '') +
    '</td></tr>' +
    '<tr><td style="border-top:2px solid ' + farbe + '; padding-top:6px;">' +
    firmenzeile +
    (telzeile ? '<br/>' + telzeile : '') +
    '<br/>' + l.email + ': <a href="mailto:' + profile.emailAddress + '" style="color:' + farbe + ';">' + profile.emailAddress + '</a>' +
    (tpl.webseite ? ' · ' + tpl.webseite : '') +
    '</td></tr>';
  /* Banner-Rangfolge: laufende Kampagne schlaegt das Vorlagen-Banner.
     Ohne Kampagne gilt buildBanner() aus signova-app/src/lib/signature.ts:
     Banner aus -> nichts; Bild gesetzt -> Bild; sonst Text-Box, sofern gefuellt. */
  if (kampagne) {
    html += kampagne;
  } else if (tpl.banner_aktiv && bannerBild) {
    html += '<tr><td style="padding-top:8px;">' +
      '<img src="' + signovaAttr(bannerBild) + '" alt="' + signovaAttr(tpl.banner_alt || "") + '" style="width:100%; max-width:' +
      SIGNOVA_BANNER_MAX_WIDTH + 'px; height:auto; display:block; border:0;" /></td></tr>';
  } else if (tpl.banner_aktiv && (tpl.banner_titel || tpl.banner_text)) {
    html += '<tr><td style="padding-top:8px;"><table cellpadding="10" cellspacing="0" style="background:' + farbe +
      '; border-radius:6px;"><tr><td style="color:#ffffff; font-family:Segoe UI, Arial, sans-serif; font-size:9.5pt;">' +
      '<strong>' + (tpl.banner_titel || "") + '</strong><br/>' + (tpl.banner_text || "") + '</td></tr></table></td></tr>';
  }
  html += '<tr><td style="padding-top:8px; font-size:8pt; color:#8A8A8A;">' + (tpl.hinweis || "") +
    ' &nbsp;(' + (vorlagenName ? l.vorlage + ': ' + vorlagenName + ' – ' : '') + (tpl.version || "?") +
    (person ? ' · Personendaten: zentral' : ' · Personendaten: nicht gefunden') + ')</td></tr>' +
    (signovaIstEntwurf(person) ? signovaEntwurfHinweisHtml() : '') +
    '</table>';
  return html;
}

/* ==================================================================
   Daten holen und Signatur setzen
   ================================================================== */

/* Zwischenspeicher der letzten Antwort. Das Panel zeigt daraus "aktualisiert
   vor X Minuten" und muss beim Umschalten zwischen Signaturen nicht jedes
   Mal erneut laden. Der Handler beim Verfassen holt bewusst frisch: Dort
   zaehlt der aktuelle Stand, nicht die Geschwindigkeit. */
var signovaZwischenspeicher = null;

function signovaLadeDaten(callback) {
  var ursache = null;
  signovaFetchJson("templates.json", null, function (templates, fehlerT) {
    if (fehlerT) ursache = fehlerT;
    signovaFetchJson("users.json", null, function (users, fehlerU) {
      if (fehlerU && !ursache) ursache = fehlerU;

      var daten = {
        templates: templates,
        users: users,
        person: signovaFindUser(users, Office.context.mailbox.userProfile.emailAddress),
        geholtAm: Date.now(),
        ursache: ursache
      };
      /* Bei einem Fehler den letzten guten Stand behalten - eine kurz
         gestoerte Verbindung soll das Panel nicht leerraeumen. */
      if (!ursache || !signovaZwischenspeicher) signovaZwischenspeicher = daten;
      callback(daten);
    });
  });
}

/* Baut das HTML fuer eine bestimmte Auswahl aus dem Panel.
   `eintrag` ist ein Element aus person.signaturen; ohne Eintrag gilt die
   automatisch bestimmte Vorlage. */
function signovaHtmlFuerAuswahl(person, profile, eintrag, composeTyp) {
  var tpl = person && person.vorlagen ? person.vorlagen[eintrag.vorlage_id] : null;
  if (!tpl) return "";

  /* Die Sprachvariante rendert dieselbe Vorlage mit den Beschriftungen der
     anderen Sprache. Dafuer bekommt signovaBuildHtml eine Kopie der Person
     mit geaenderter Sprache - keine zweite Renderfunktion. */
  var wer = person;
  if (eintrag.sprache && eintrag.sprache !== signovaSprache(person)) {
    wer = {};
    for (var k in person) if (Object.prototype.hasOwnProperty.call(person, k)) wer[k] = person[k];
    wer.sprache = eintrag.sprache;
  }

  var typ = eintrag.art === "antwort" ? "antwort" : composeTyp;
  if (signovaUnterdrueckt(tpl, typ)) return "";
  return signovaBuildHtml(tpl, profile, wer, eintrag.vorlage_id, typ);
}

/**
 * Setzt die Signatur im Entwurf.
 *
 * setSignatureAsync ERSETZT den Signaturbereich - deshalb kann weder ein
 * zweiter Aufruf noch ein Empfaengerwechsel eine zweite Signatur erzeugen.
 * Genau darauf beruht die Zusage "ersetzen, nie duplizieren"; das Anhaengen
 * an den Nachrichtentext (setSelectedDataAsync) taete das Gegenteil und wird
 * hier bewusst nicht benutzt.
 *
 * Leeres HTML ist ein gueltiger Fall (reply_mode 'first_only'): Es loescht
 * eine zuvor gesetzte Signatur, statt sie stehen zu lassen.
 */
function signovaSetzeSignatur(item, html, callback) {
  item.body.setSignatureAsync(
    signovaMarkiere(html),
    { coercionType: Office.CoercionType.Html },
    callback
  );
}

/**
 * Bestimmt und setzt die Signatur.
 *
 * `auswahl` ist optional und kommt aus dem Panel ("Diese Signatur
 * verwenden"). Ohne sie entscheiden Server-Vorlage und Empfaenger-Regeln.
 */
function signovaApply(done, modus, auswahl) {
  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile;

  signovaLadeDaten(function (daten) {
    var person = daten.person;
    var ursache = daten.ursache;

    signovaComposeTyp(item, function (composeTyp) {
      signovaEmpfaenger(item, function (adressen) {
        var tpl = null;
        var name = person && person.vorlage ? person.vorlage : "standard";

        if (auswahl) {
          tpl = person && person.vorlagen ? person.vorlagen[auswahl.vorlage_id] : null;
          name = auswahl.vorlage_id;
        } else {
          /* Empfaenger-Regeln zuerst: Sie sind die einzige Entscheidung, die
             der Server nicht treffen konnte. Greift keine, bleibt es bei
             seiner Vorlage. */
          var nachEmpfaenger = signovaVorlageNachEmpfaengern(person, adressen);
          if (nachEmpfaenger) {
            tpl = nachEmpfaenger.tpl;
            name = nachEmpfaenger.name;
          } else {
            tpl = signovaPickTemplate(daten.templates, person);
          }
        }

        if (!tpl) { tpl = signovaFallbackTemplate(); name = ""; }   /* Notfall: alte/keine templates.json */

        var html = auswahl
          ? signovaHtmlFuerAuswahl(person, profile, auswahl, composeTyp)
          : (signovaUnterdrueckt(tpl, composeTyp)
              ? ""
              : signovaBuildHtml(tpl, profile, person, name, composeTyp));

        signovaSetzeSignatur(item, html, function (res) {
          if (res && res.status === Office.AsyncResultStatus.Succeeded) {
            /* Ein Entwurf ist kein veroeffentlichter Stand - die Version
               bekommt deshalb einen Zusatz, sonst zaehlte die Verteilung ihn
               als "aktuell". */
            var gemeldet = signovaIstEntwurf(person)
              ? { version: (tpl.version || "") + " (Entwurf)" }
              : tpl;
            signovaPing(profile, gemeldet, name, modus === "button" ? "button" : "auto");
          }
          /* Die Ursache wird mitgegeben, damit die Taskpane sagen kann, WARUM
             die Notfall-Vorlage gegriffen hat, statt nur "erfolgreich". */
          done(res, ursache, { html: html, vorlage: name });
        });
      });
    });
  });
}

/* Automatik: startet bei neuer Mail – erfordert Admin-Deployment (M365 Admin Center) */
function onNewMessageComposeHandler(event) {
  signovaApply(function () { event.completed(); }, "auto");
}

/**
 * Automatik: Die Empfaenger haben sich geaendert.
 *
 * Nur dann neu setzen, wenn fuer diese Person ueberhaupt eine Regel von den
 * Empfaengern abhaengt. Sonst waere jedes Tippen im An-Feld ein Schreibzugriff
 * auf den Entwurf - und wuerde eine im Panel getroffene Auswahl ueberschreiben,
 * ohne dass sich am Ergebnis etwas aendert.
 *
 * Braucht Mailbox 1.13. Wo Outlook das Ereignis nicht kennt, passiert schlicht
 * nichts; die Signatur bleibt die beim Oeffnen gesetzte.
 */
function onMessageRecipientsChangedHandler(event) {
  var person = signovaZwischenspeicher && signovaZwischenspeicher.person;

  if (person && (!person.empfaenger_regeln || !person.empfaenger_regeln.length)) {
    event.completed();
    return;
  }

  signovaApply(function () { event.completed(); }, "auto");
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate(
    "onMessageRecipientsChangedHandler",
    onMessageRecipientsChangedHandler
  );
}
