/*
  Seitenleisten-Test für klassisches Outlook.
  Zeigt nur Betreff und Absender der ausgewählten Mail, Prüfergebnisse und einen Mikrofonpegel.
  Kein Mailtext, kein Speichern, keine eigene Netzwerkanfrage, keine KI.
  Bewusst in einfachem JavaScript geschrieben, damit die Seite auch in einer alten Webansicht ihr Ergebnis zeigen kann.
*/
(function () {
  "use strict";

  var MAILBOX_VERSIONEN = [
    "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9",
    "1.10", "1.11", "1.12", "1.13", "1.14", "1.15", "1.16", "1.17"
  ];
  var PEGEL_TAKT_MS = 100;
  var MIKROFON_HOECHSTDAUER_MS = 60000;
  var OUTLOOK_WARTEZEIT_MS = 15000;

  var wechselAnzahl = 0;
  var freigabeLauf = 0;
  var mikrofon = null;
  var mikrofonWartet = false;
  var seiteVerlassen = false;

  function el(id) {
    return document.getElementById(id);
  }

  function schreibe(id, text) {
    el(id).textContent = text;
  }

  function fehlerText(fehler) {
    if (fehler && fehler.name) {
      return fehler.name;
    }
    if (fehler && fehler.code) {
      return "Fehlercode " + fehler.code;
    }
    return "Unbekannter Fehler";
  }

  function sicher(aufgabe) {
    try {
      aufgabe();
    } catch (fehler) {
      schreibe("verbindung", "Mit Outlook verbunden, aber ein Prüfschritt schlug fehl: " + fehlerText(fehler));
    }
  }

  /* Ausgewählte Mail: liest nur Betreff und Absender. */

  function absenderText(von) {
    if (!von) {
      return "Nicht verfügbar";
    }
    if (von.displayName && von.emailAddress) {
      return von.displayName + " <" + von.emailAddress + ">";
    }
    return von.displayName || von.emailAddress || "Nicht verfügbar";
  }

  function pruefeFreigabe(item, lauf) {
    if (typeof item.getSharedPropertiesAsync !== "function") {
      schreibe("ergebnis-freigabe", "Nicht prüfbar (Funktion fehlt)");
      return;
    }
    schreibe("ergebnis-freigabe", "Wird geprüft");
    item.getSharedPropertiesAsync(function (ergebnis) {
      /* Antwort einer früheren Mail nicht der jetzt ausgewählten zuordnen. Inhalt der Antwort wird nicht gelesen. */
      if (lauf !== freigabeLauf) {
        return;
      }
      schreibe("ergebnis-freigabe", ergebnis.status === Office.AsyncResultStatus.Succeeded ? "Ja" : "Nein");
    });
  }

  function zeigeMail() {
    var item = Office.context.mailbox.item;
    freigabeLauf += 1;
    if (!item) {
      schreibe("betreff", "Keine einzelne Mail ausgewählt");
      schreibe("absender", "Keine einzelne Mail ausgewählt");
      schreibe("ergebnis-freigabe", "Keine Mail ausgewählt");
      return;
    }
    try {
      schreibe("betreff", typeof item.subject === "string" && item.subject ? item.subject : "(ohne Betreff)");
      schreibe("absender", absenderText(item.from));
      pruefeFreigabe(item, freigabeLauf);
    } catch (fehler) {
      schreibe("betreff", "Nicht lesbar: " + fehlerText(fehler));
      schreibe("absender", "Nicht lesbar: " + fehlerText(fehler));
    }
  }

  /* Mitgehen und Anheften */

  function beiMailwechsel() {
    /* Microsoft meldet ItemChanged nur, solange die Leiste angeheftet ist. Das Ereignis belegt beides. */
    wechselAnzahl += 1;
    schreibe("ergebnis-mitgehen", "Ja (" + wechselAnzahl + "-mal)");
    schreibe("ergebnis-anheften", "Ja (Leiste blieb offen)");
    zeigeMail();
  }

  function meldeWechselAn(mailbox) {
    if (!Office.EventType || !Office.EventType.ItemChanged || typeof mailbox.addHandlerAsync !== "function") {
      schreibe("ergebnis-mitgehen", "Nein (Ereignis fehlt)");
      return;
    }
    mailbox.addHandlerAsync(Office.EventType.ItemChanged, beiMailwechsel, function (ergebnis) {
      if (ergebnis.status === Office.AsyncResultStatus.Failed) {
        schreibe("ergebnis-mitgehen", "Nein (" + fehlerText(ergebnis.error) + ")");
      }
    });
  }

  /* Schnittstellen */

  function outlookArtText(name) {
    if (name === "Outlook") {
      return "Klassisch (" + name + ")";
    }
    if (name === "newOutlookWindows") {
      return "Neues Outlook (" + name + ")";
    }
    if (name === "OutlookWebApp") {
      return "Im Browser (" + name + ")";
    }
    return name;
  }

  function zeigeSchnittstellen() {
    var kontext = Office.context;
    var diagnose = kontext.diagnostics;
    var mailDiagnose = kontext.mailbox.diagnostics;
    var ja = [];
    var nein = [];
    var i;

    schreibe("host", diagnose && diagnose.host ? String(diagnose.host) : "Nicht gemeldet");
    schreibe("outlook-art", mailDiagnose && mailDiagnose.hostName ? outlookArtText(mailDiagnose.hostName) : "Nicht gemeldet");
    schreibe("plattform", diagnose && diagnose.platform ? String(diagnose.platform) : "Nicht gemeldet");
    if (diagnose && diagnose.version) {
      schreibe("version", diagnose.version);
    } else {
      schreibe("version", mailDiagnose && mailDiagnose.hostVersion ? mailDiagnose.hostVersion : "Nicht gemeldet");
    }

    for (i = 0; i < MAILBOX_VERSIONEN.length; i += 1) {
      if (kontext.requirements.isSetSupported("Mailbox", MAILBOX_VERSIONEN[i])) {
        ja.push(MAILBOX_VERSIONEN[i]);
      } else {
        nein.push(MAILBOX_VERSIONEN[i]);
      }
    }
    schreibe("mailbox-ja", ja.length ? ja.join(", ") : "Keine");
    schreibe("mailbox-nein", nein.length ? nein.join(", ") : "Keine");

    if (!kontext.requirements.isSetSupported("Mailbox", "1.5")) {
      schreibe("ergebnis-anheften", "Nein (Mailbox 1.5 fehlt)");
    }
  }

  function zeigeStartOhneMail(mailbox) {
    if (!mailbox.item) {
      schreibe("ergebnis-ohne-mail", "Ja");
    } else if (!Office.context.requirements.isSetSupported("Mailbox", "1.13")) {
      schreibe("ergebnis-ohne-mail", "Nein (Mailbox 1.13 fehlt)");
    }
  }

  /* Mikrofontest: nur auf Knopfdruck, nur Pegel, keine Aufnahme. */

  function zeigePegel(prozent) {
    el("pegel-balken").style.width = prozent + "%";
    el("pegel").setAttribute("aria-valuenow", String(prozent));
  }

  function setzeKnopf(zustand) {
    var knopf = el("mikrofon-knopf");
    if (zustand === "wartet") {
      knopf.textContent = "Warte auf Freigabe";
      knopf.disabled = true;
    } else if (zustand === "laeuft") {
      knopf.textContent = "Mikrofontest stoppen";
      knopf.disabled = false;
    } else {
      knopf.textContent = "Mikrofontest starten";
      knopf.disabled = false;
    }
  }

  function beendeSpuren(strom) {
    var spuren = strom.getTracks();
    var i;
    for (i = 0; i < spuren.length; i += 1) {
      spuren[i].stop();
    }
  }

  function stoppeMikrofon(meldung) {
    var lauf = mikrofon;
    if (!lauf) {
      return;
    }
    mikrofon = null;
    window.clearInterval(lauf.takt);
    window.clearTimeout(lauf.frist);
    beendeSpuren(lauf.strom);
    if (lauf.quelle) {
      lauf.quelle.disconnect();
    }
    if (lauf.kontext && lauf.kontext.close) {
      lauf.kontext.close();
    }
    zeigePegel(0);
    setzeKnopf("bereit");
    schreibe("mikrofon-status", meldung);
  }

  function starteAnzeige(lauf) {
    var AudioKontext = window.AudioContext || window.webkitAudioContext;
    var analyse;
    var puffer;
    if (!AudioKontext) {
      schreibe("mikrofon-status", "Mikrofon ist da, aber diese Webansicht kann keinen Pegel anzeigen.");
      return;
    }
    lauf.kontext = new AudioKontext();
    if (lauf.kontext.state === "suspended" && lauf.kontext.resume) {
      lauf.kontext.resume();
    }
    lauf.quelle = lauf.kontext.createMediaStreamSource(lauf.strom);
    analyse = lauf.kontext.createAnalyser();
    analyse.fftSize = 1024;
    /* Nur zur Analyse verbunden. Nicht an die Lautsprecher und nicht an eine Aufnahme. */
    lauf.quelle.connect(analyse);
    puffer = new Uint8Array(analyse.fftSize);
    lauf.takt = window.setInterval(function () {
      var spitze = 0;
      var i;
      var abstand;
      analyse.getByteTimeDomainData(puffer);
      for (i = 0; i < puffer.length; i += 1) {
        abstand = Math.abs(puffer[i] - 128);
        if (abstand > spitze) {
          spitze = abstand;
        }
      }
      zeigePegel(Math.min(100, Math.round(Math.sqrt(spitze / 128) * 100)));
    }, PEGEL_TAKT_MS);
  }

  function beiMikrofonStrom(strom) {
    var spuren;
    var lauf;
    var i;
    mikrofonWartet = false;
    if (seiteVerlassen) {
      beendeSpuren(strom);
      return;
    }
    lauf = { strom: strom, kontext: null, quelle: null, takt: null, frist: null };
    mikrofon = lauf;
    schreibe("ergebnis-mikrofon", "Ja");
    schreibe("mikrofon-status", "Bitte sprechen. Der Balken sollte sich bewegen.");
    setzeKnopf("laeuft");
    lauf.frist = window.setTimeout(function () {
      stoppeMikrofon("Nach 60 Sekunden automatisch beendet. Das Mikrofon ist wieder frei.");
    }, MIKROFON_HOECHSTDAUER_MS);
    spuren = strom.getTracks();
    for (i = 0; i < spuren.length; i += 1) {
      if (spuren[i].addEventListener) {
        spuren[i].addEventListener("ended", function () {
          stoppeMikrofon("Die Verbindung zum Mikrofon wurde beendet.");
        });
      }
    }
    try {
      starteAnzeige(lauf);
    } catch (fehler) {
      schreibe("mikrofon-status", "Mikrofon ist da, Pegel nicht anzeigbar: " + fehlerText(fehler));
    }
  }

  function beiMikrofonFehler(fehler) {
    mikrofonWartet = false;
    setzeKnopf("bereit");
    schreibe("ergebnis-mikrofon", "Nein (" + fehlerText(fehler) + ")");
    schreibe("mikrofon-status", "");
  }

  function starteMikrofon() {
    var medien = navigator.mediaDevices;
    if (mikrofon || mikrofonWartet) {
      return;
    }
    if (!medien || typeof medien.getUserMedia !== "function") {
      schreibe("ergebnis-mikrofon", "Nein (keine Schnittstelle in dieser Webansicht)");
      return;
    }
    mikrofonWartet = true;
    setzeKnopf("wartet");
    try {
      medien.getUserMedia({ audio: true, video: false }).then(beiMikrofonStrom, beiMikrofonFehler);
    } catch (fehler) {
      beiMikrofonFehler(fehler);
    }
  }

  function beiSeitenende() {
    seiteVerlassen = true;
    stoppeMikrofon("");
  }

  /* Stimme und Spracherkennung: nur prüfen, was da ist. Die Spracherkennung wird nicht gestartet, weil sie Ton ins Netz schicken kann. */

  function stimmenListe() {
    var ausgabe = window.speechSynthesis;
    if (!ausgabe || typeof ausgabe.getVoices !== "function") {
      return null;
    }
    return ausgabe.getVoices() || [];
  }

  function istDeutsch(stimme) {
    return Boolean(stimme && stimme.lang && stimme.lang.toLowerCase().indexOf("de") === 0);
  }

  function zeigeStimmen() {
    var alle = stimmenListe();
    var lokal = [];
    var netz = [];
    var i;
    if (alle === null) {
      schreibe("ergebnis-stimmen", "Nein (keine Sprachausgabe in dieser Webansicht)");
      schreibe("ergebnis-stimmen-netz", "Nicht prüfbar");
      return;
    }
    for (i = 0; i < alle.length; i += 1) {
      if (!istDeutsch(alle[i])) {
        continue;
      }
      if (alle[i].localService) {
        lokal.push(alle[i].name);
      } else {
        netz.push(alle[i].name);
      }
    }
    schreibe("ergebnis-stimmen", lokal.length ? lokal.join(", ") : "Keine");
    schreibe("ergebnis-stimmen-netz", netz.length ? netz.join(", ") : "Keine");
  }

  function testeStimme() {
    var ausgabe = window.speechSynthesis;
    var alle = stimmenListe();
    var stimme = null;
    var satz;
    var i;
    if (!ausgabe || alle === null || typeof window.SpeechSynthesisUtterance !== "function") {
      schreibe("ergebnis-stimme", "Nein (keine Sprachausgabe in dieser Webansicht)");
      return;
    }
    for (i = 0; i < alle.length; i += 1) {
      if (alle[i].localService && istDeutsch(alle[i])) {
        stimme = alle[i];
        break;
      }
    }
    if (!stimme) {
      schreibe("ergebnis-stimme", "Nein (keine deutsche Stimme auf diesem Rechner)");
      return;
    }
    satz = new window.SpeechSynthesisUtterance("Das ist ein Test der Stimme.");
    satz.voice = stimme;
    satz.lang = stimme.lang;
    satz.onend = function () {
      schreibe("ergebnis-stimme", "Ja, mit " + stimme.name + ". Ob sie zu hören war, bitte in die Testliste eintragen.");
    };
    satz.onerror = function () {
      schreibe("ergebnis-stimme", "Nein (Fehler bei der Ausgabe)");
    };
    schreibe("ergebnis-stimme", "Spricht gerade");
    ausgabe.cancel();
    ausgabe.speak(satz);
  }

  function zeigeErkennung() {
    var erkennung = window.SpeechRecognition;
    var da = typeof erkennung === "function" || typeof window.webkitSpeechRecognition === "function";
    var anfrage;
    schreibe("ergebnis-erkennung", da ? "Ja (vorhanden, nicht gestartet)" : "Nein");
    /* Nur nachfragen, nicht installieren: available() lädt nichts und nimmt nichts auf. */
    if (typeof erkennung !== "function" || typeof erkennung.available !== "function") {
      schreibe("ergebnis-erkennung-lokal", "Nein (Abfrage fehlt in dieser Webansicht)");
      return;
    }
    try {
      anfrage = erkennung.available({ langs: ["de-DE"], processLocally: true });
    } catch (fehler) {
      schreibe("ergebnis-erkennung-lokal", "Nein (" + fehlerText(fehler) + ")");
      return;
    }
    Promise.resolve(anfrage).then(function (stand) {
      var texte = { available: "Ja, das Modell ist schon da", downloadable: "Ja, nach einmaligem Laden des Modells", downloading: "Ja, das Modell lädt gerade", unavailable: "Nein" };
      schreibe("ergebnis-erkennung-lokal", texte[stand] || ("Unklar (" + String(stand) + ")"));
    }, function (fehler) {
      schreibe("ergebnis-erkennung-lokal", "Nein (" + fehlerText(fehler) + ")");
    });
  }

  /* Start */

  zeigeErkennung();
  zeigeStimmen();
  if (window.speechSynthesis && "onvoiceschanged" in window.speechSynthesis) {
    window.speechSynthesis.onvoiceschanged = zeigeStimmen;
  }
  el("stimme-knopf").addEventListener("click", testeStimme);

  el("mikrofon-knopf").addEventListener("click", function () {
    if (mikrofon) {
      stoppeMikrofon("Test beendet. Das Mikrofon ist wieder frei.");
    } else {
      starteMikrofon();
    }
  });
  window.addEventListener("pagehide", beiSeitenende);
  window.addEventListener("beforeunload", beiSeitenende);

  if (typeof Office === "undefined" || !Office.onReady) {
    schreibe("verbindung", "Office.js wurde nicht geladen. Netzwerkfreigabe prüfen. Der Mikrofontest geht trotzdem.");
    return;
  }

  var wartefrist = window.setTimeout(function () {
    schreibe("verbindung", "Keine Antwort von Outlook. Der Mikrofontest geht trotzdem.");
  }, OUTLOOK_WARTEZEIT_MS);

  Office.onReady(function (info) {
    window.clearTimeout(wartefrist);
    if (!info || info.host !== Office.HostType.Outlook) {
      schreibe("verbindung", "Diese Seite läuft nicht in Outlook. Die Mailzeilen bleiben leer.");
      return;
    }
    schreibe("verbindung", "Mit Outlook verbunden.");
    /* Jeder Schritt läuft für sich: Ein Fehler soll sichtbar sein und die übrigen Ergebnisse nicht verhindern. */
    sicher(zeigeSchnittstellen);
    sicher(function () {
      zeigeStartOhneMail(Office.context.mailbox);
    });
    sicher(function () {
      meldeWechselAn(Office.context.mailbox);
    });
    sicher(zeigeMail);
  });
}());
