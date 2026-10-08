/**
 * Sagra di Penco — prenotazioni dal sito sagradipenco.it
 * ======================================================
 *
 * INSTALLAZIONE (una volta sola) — tutto dall'account sagradipenco@gmail.com
 * 1. Crea un foglio Google vuoto nel Drive della sagra e chiamalo "Prenotazioni Sagra di Penco".
 * 2. Estensioni → Apps Script. Cancella il contenuto e incolla questo file. Salva (icona dischetto).
 * 3. In alto scegli la funzione "setup" e premi ▶ Esegui. Concedi i permessi quando li chiede
 *    (Avanzate → Vai a … (non sicuro) → Consenti). Il foglio si riempie con i tab:
 *      Config        → capienze, turni, chiusure, email: QUI si regolano i limiti
 *      Destinatari   → chi riceve il recap del mattino e gli avvisi "esaurito" (un indirizzo per riga)
 *      Riepilogo     → prenotati / posti rimasti per ogni evento e turno (formule, si aggiorna da solo)
 *      un tab per ogni evento con form online → l'elenco delle prenotazioni
 * 4. Distribuisci → Nuova distribuzione → tipo "Applicazione web":
 *      Descrizione: prenotazioni · Esegui come: Me · Chi ha accesso: Chiunque
 *    Copia l'URL dell'app web (finisce con /exec) e incollalo in FORM_ENDPOINT nel file index.html.
 * 5. Condividi il foglio (bottone Condividi) con chi deve vedere o inserire le prenotazioni.
 * 6. Avvisi automatici: scrivi gli indirizzi nel tab Destinatari, poi esegui una volta la funzione
 *    "installaRecap". Attiva due cose:
 *      - ogni mattina all'ora ORA_RECAP una mail con il punto delle prenotazioni e quelle arrivate ieri
 *        (per provarlo subito: esegui "recapGiornaliero");
 *      - un avviso immediato quando un evento o un turno raggiunge la capienza, sia per le prenotazioni
 *        dal sito sia per quelle scritte a mano nel foglio (per provarlo: esegui "controllaLimiti").
 *
 * SHOP: nello stesso progetto c'è anche il file "shop" (ordini della pagina /shop). Questo file
 * smista già gli ordini allo shop all'inizio di doPost: sostituendo questo file non serve aggiungere nulla.
 *
 * SE MODIFICHI QUESTO CODICE: Distribuisci → Gestisci distribuzioni → ✎ → Versione: Nuova → Distribuisci.
 * (Senza questo passaggio il sito continua a usare la versione vecchia.)
 *
 * PRENOTAZIONI TELEFONICHE: gli organizzatori le scrivono a mano nel tab dell'evento, una per riga:
 * Nome, Telefono, le due quantità (Adulti/Bambini, oppure Menu/Menu veg per il pranzo cinese),
 * Totale (obbligatorio!) e, per il pranzo cinese, l'Orario (es. 12:30).
 * Contano nei posti esattamente come quelle del sito. Per annullare una prenotazione scrivere
 * "annullata" nella colonna Stato: la riga resta ma non conta più.
 */

// ---- Impostazioni ---------------------------------------------------------
const NOME_MITTENTE   = "Sagra di Penco";          // nome che compare nelle email
const EMAIL_IN_COPIA  = "";                         // eventuale indirizzo in copia a ogni conferma; vuoto = nessuna copia (le conferme restano nella posta dell'account che esegue lo script)
const FUSO            = "Europe/Rome";
const ORA_RECAP       = 8;                          // ora (0-23) del recap del mattino (riporta le prenotazioni di ieri)

// Valori iniziali del tab Config (poi si modificano nel foglio, non qui).
// id = deve coincidere con l'id dell'evento nel sito.
const CONFIG_INIZIALE = [
  // id                   evento                     capienza  turni                        capienza_turno  chiusura            email_avviso            attivo  campi
  ["ven-cena-francese",  "Cena Francese con Penco",  100,      "",                          "",             "2026-10-14 23:59", "",                     "SI",   "Adulti € 20 | Bambini (menu bambini € 10)"],
  ["sab-pranzo-cinese",  "Mezzogiorno d'Oriente",    100,      "12:00, 12:30, 13:00, 13:30", 25,            "2026-10-14 23:59", "",                     "SI",   "Menu € 15 | Menu veg € 10"],
  ["sab-penco-trekking", "Penco Trekking",           "",       "",                          "",             "2026-10-14 23:59", "",                     "SI",   "Adulti | Bambini <11"],
  ["dom-sellero-penco",  "A piedi da S. Ellero",     "",       "",                          "",             "2026-10-14 23:59", "circolo.sellero@gmail.com", "SI", "Adulti | Bambini"]
];
const CONFIG_INTESTAZIONE = ["id", "evento", "capienza", "turni", "capienza_turno", "chiusura", "email_avviso", "attivo", "campi"];
const CONFIG_NOTE = [
  "Identificativo usato dal sito. Non cambiarlo.",
  "Nome del tab con le prenotazioni (creato da setup).",
  "Posti totali. Vuoto = senza limite.",
  "Orari separati da virgola (es. 12:00, 12:30). Vuoto = nessun turno.",
  "Posti per ogni turno. Vuoto = senza limite per turno.",
  "Data e ora oltre cui il form si chiude (AAAA-MM-GG HH:MM). Vuoto = mai.",
  "Chi riceve l'avviso di ogni nuova prenotazione (più indirizzi separati da virgola). Vuoto = nessun avviso, basta il recap del mattino.",
  "SI = form attivo sul sito. NO = chiuso a mano.",
  "Nomi delle due quantità chieste (separati da |): diventano le intestazioni delle colonne E ed F del tab."
];

// Colonne dei tab evento
const COL = ["Data", "Canale", "Nome", "Telefono", "Q1", "Q2", "Totale", "Orario", "Email", "Note", "Stato"];  // Q1/Q2 prendono il nome da Config.campi
const C = Object.fromEntries(COL.map((n, i) => [n, i]));   // C.Totale → indice colonna

// ---- Endpoint --------------------------------------------------------------

function doGet(e) {
  try {
    const azione = (e && e.parameter && e.parameter.azione) || "";
    if (azione === "disponibilita") {
      // cache di 20 secondi: evita di rileggere tutto il foglio a ogni apertura della pagina
      const cache = CacheService.getScriptCache();
      let json = cache.get("disponibilita");
      if (!json) { json = JSON.stringify(disponibilita_()); cache.put("disponibilita", json, 20); }
      return ContentService.createTextOutput('{"ok":true,"eventi":' + json + '}').setMimeType(ContentService.MimeType.JSON);
    }
    if (azione === "esito") {
      // il sito chiede se una prenotazione (identificata dal codice rid) è stata registrata
      const rid = String(e.parameter.rid || "");
      return risposta_({ ok: true, registrata: !!(rid && CacheService.getScriptCache().get("rid_" + rid)) });
    }
    return risposta_({ ok: true, info: "Prenotazioni Sagra di Penco" });
  } catch (err) { return risposta_({ ok: false, errore: String(err.message || err) }); }
}

function doPost(e) {
  // Ordini dello shop dei volontari (file "shop" nello stesso progetto, vedi apps-script/shop.gs):
  // usano lo stesso indirizzo /exec e vengono smistati qui. NON TOGLIERE questa riga.
  if (typeof shopEOrdine_ === "function" && shopEOrdine_(e)) return shopOrdine_(e);

  const lock = LockService.getScriptLock();
  let preso = false;
  try {
    preso = lock.tryLock(20000);              // evita che due prenotazioni contemporanee superino il limite
    if (!preso) throw new Error("Troppe prenotazioni in questo momento: riprova tra qualche secondo.");
    const d = JSON.parse(e.postData.contents);
    const cache = CacheService.getScriptCache();
    const rid = String(d.rid || "").slice(0, 40);
    // stesso codice già registrato (es. l'utente ha ripremuto Invia): non si scrive una seconda volta
    if (rid && cache.get("rid_" + rid)) return risposta_({ ok: true, duplicata: true });
    const cfg = config_()[d.id];
    if (!cfg) throw new Error("Evento non riconosciuto.");
    if (!d.nome || !d.telefono || !d.email) throw new Error("Compila nome, telefono ed email.");

    const q1 = Number(d.q1) || 0, q2 = Number(d.q2) || 0, totale = q1 + q2;
    if (totale < 1) throw new Error("Indica almeno una persona.");
    if (cfg.turni.length && !cfg.turni.includes(d.orario)) throw new Error("Scegli un orario.");

    const stato = statoEvento_(cfg);
    if (!stato.aperto) return risposta_({ ok: false, esaurito: true, errore: "Le prenotazioni online sono chiuse: prova per telefono." });
    if (stato.esaurito) return risposta_({ ok: false, esaurito: true, errore: "Posti esauriti: prova per telefono." });
    if (d.orario && stato.turni[d.orario].esaurito)
      return risposta_({ ok: false, esaurito: true, errore: "Turno delle " + d.orario + " esaurito: scegli un altro orario." });

    const riga = [];
    riga[C.Data] = new Date(); riga[C.Canale] = "sito";
    riga[C.Nome] = String(d.nome).trim(); riga[C.Telefono] = String(d.telefono).trim();
    riga[C.Q1] = q1; riga[C.Q2] = q2; riga[C.Totale] = totale;
    riga[C.Orario] = d.orario || ""; riga[C.Email] = String(d.email).trim();
    riga[C.Note] = String(d.note || "").trim(); riga[C.Stato] = "";
    tabEvento_(cfg.evento, cfg.campi).appendRow(riga);
    SpreadsheetApp.flush();
    cache.remove("disponibilita");
    if (rid) cache.put("rid_" + rid, "1", 21600);   // ricordato per 6 ore

    // da qui la prenotazione è registrata: un problema con le mail non deve farla risultare fallita
    d.q1 = q1; d.q2 = q2; d.totale = totale; d.evento = cfg.evento;
    d.dettaglio = `${q1} ${cfg.campi[0]}, ${q2} ${cfg.campi[1]}`;
    let mail = true;
    try { confermaOspite_(d, cfg); } catch (err) { mail = false; console.error("Conferma non inviata: " + err); }
    try { avvisaOrganizzatori_(d, cfg, stato); } catch (err) { console.error("Avviso non inviato: " + err); }
    try { controllaLimiti_(cfg.id); } catch (err) { console.error("Controllo limiti: " + err); }
    return risposta_({ ok: true, mail: mail });
  } catch (err) {
    return risposta_({ ok: false, errore: String(err.message || err) });
  } finally { if (preso) lock.releaseLock(); }
}

// ---- Setup ------------------------------------------------------------------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Config
  let cfg = ss.getSheetByName("Config");
  if (!cfg) {
    cfg = ss.insertSheet("Config", 0);
    cfg.appendRow(CONFIG_INTESTAZIONE);
    cfg.appendRow(CONFIG_NOTE);
    CONFIG_INIZIALE.forEach(r => cfg.appendRow(r));
    cfg.getRange(1, 1, 1, CONFIG_INTESTAZIONE.length).setFontWeight("bold").setBackground("#204f7a").setFontColor("#ffffff");
    cfg.getRange(2, 1, 1, CONFIG_INTESTAZIONE.length).setFontStyle("italic").setFontColor("#5b544b").setWrap(true);
    cfg.getRange(3, 6, CONFIG_INIZIALE.length, 1).setNumberFormat("@");   // chiusura come testo
    cfg.setFrozenRows(2); cfg.setColumnWidths(1, 9, 170); cfg.setColumnWidth(2, 220);
  }
  // Tab eventi
  Object.values(config_()).forEach(c => tabEvento_(c.evento, c.campi));
  // Destinatari del recap
  if (!ss.getSheetByName("Destinatari")) {
    const d = ss.insertSheet("Destinatari", 1);
    d.appendRow(["email", "nome (facoltativo)"]);
    d.appendRow(["Un indirizzo per riga: ricevono il recap del mattino e gli avvisi di posti esauriti.", ""]);
    d.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#204f7a").setFontColor("#ffffff");
    d.getRange(2, 1).setFontStyle("italic").setFontColor("#5b544b");
    d.setFrozenRows(2); d.setColumnWidth(1, 280);
  }
  // Riepilogo
  riepilogo_();
  // Rimuove il foglio vuoto iniziale
  const f1 = ss.getSheetByName("Foglio1") || ss.getSheetByName("Sheet1");
  if (f1 && f1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(f1);
  Logger.log("Setup completato. Ora: Distribuisci → Nuova distribuzione → Applicazione web.");
}

function tabEvento_(nome, campi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let t = ss.getSheetByName(nome);
  if (!t) {
    t = ss.insertSheet(nome);
    const testata = COL.slice(); testata[C.Q1] = campi[0]; testata[C.Q2] = campi[1];
    t.appendRow(testata);
    t.getRange(1, 1, 1, COL.length).setFontWeight("bold").setBackground("#c6352c").setFontColor("#ffffff");
    t.setFrozenRows(1);
    t.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
    t.setColumnWidth(C.Nome + 1, 200); t.setColumnWidth(C.Email + 1, 220); t.setColumnWidth(C.Note + 1, 260);
    const nota = "Prenotazioni telefoniche: aggiungi una riga con Nome, Telefono, " + campi[0] + ", " + campi[1] + ", TOTALE e (se previsto) Orario. " +
                 "Per annullare scrivi 'annullata' in Stato.";
    t.getRange(1, 1).setNote(nota);
  }
  return t;
}

function riepilogo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let r = ss.getSheetByName("Riepilogo");
  if (r) ss.deleteSheet(r);
  r = ss.insertSheet("Riepilogo", 1);
  const righe = [["Evento", "Turno", "Capienza", "Prenotati", "Posti rimasti"]];
  Object.values(config_()).forEach(c => {
    const t = "'" + c.evento.replace(/'/g, "''") + "'";
    const somma = orario => `=SUMIFS(${t}!G:G, ${t}!K:K, "<>annullata"` + (orario ? `, ${t}!H:H, "${orario}"` : "") + `)`;
    righe.push([c.evento, "totale", c.capienza || "", somma(""), c.capienza ? `=C${righe.length + 1}-D${righe.length + 1}` : "senza limite"]);
    c.turni.forEach(o => {
      righe.push(["", o, c.capienzaTurno || "", somma(o), c.capienzaTurno ? `=C${righe.length + 1}-D${righe.length + 1}` : "senza limite"]);
    });
  });
  r.getRange(1, 1, righe.length, 5).setValues(righe);
  r.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#204f7a").setFontColor("#ffffff");
  r.setFrozenRows(1); r.setColumnWidth(1, 240);
  r.getRange(2, 5, righe.length - 1, 1).setFontWeight("bold");
}

// ---- Letture ------------------------------------------------------------------

function config_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!sh) throw new Error("Manca il tab Config: esegui setup.");
  const out = {};
  sh.getDataRange().getValues().slice(2).forEach(r => {
    if (!r[0]) return;
    out[String(r[0]).trim()] = {
      id: String(r[0]).trim(), evento: String(r[1]).trim(),
      capienza: Number(r[2]) || 0,
      turni: String(r[3] || "").split(",").map(s => s.trim()).filter(Boolean),
      capienzaTurno: Number(r[4]) || 0,
      chiusura: parseChiusura_(r[5]),
      email: String(r[6] || "").trim(),
      attivo: String(r[7] || "SI").trim().toUpperCase() !== "NO",
      campi: (String(r[8] || "Adulti | Bambini <11").split("|").map(s => s.trim()).concat(["", ""])).slice(0, 2)
    };
  });
  return out;
}

// Conta le persone prenotate leggendo il tab dell'evento (sito + telefono, escluse le annullate)
function conteggi_(cfg) {
  const t = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.evento);
  const tot = { totale: 0, turni: {} };
  cfg.turni.forEach(o => tot.turni[o] = 0);
  if (!t || t.getLastRow() < 2) return tot;
  t.getRange(2, 1, t.getLastRow() - 1, COL.length).getValues().forEach(r => {
    if (String(r[C.Stato]).trim().toLowerCase() === "annullata") return;
    let n = Number(r[C.Totale]);
    if (!n) n = (Number(r[C.Q1]) || 0) + (Number(r[C.Q2]) || 0);
    if (!n) return;
    tot.totale += n;
    const o = normalizzaOrario_(r[C.Orario]);
    if (o in tot.turni) tot.turni[o] += n;
  });
  return tot;
}

// Regola: un turno/evento è "esaurito" quando i prenotati raggiungono la capienza.
// L'ultima prenotazione può superare il limite (se restano 7 posti e ne chiedono 8, passa).
function statoEvento_(cfg) {
  const c = conteggi_(cfg);
  const ora = new Date();
  const aperto = cfg.attivo && !(cfg.chiusura && ora > cfg.chiusura);
  const st = {
    aperto: aperto,
    chiusura: cfg.chiusura ? Utilities.formatDate(cfg.chiusura, FUSO, "d MMMM") : "",   // mostrata sul sito ("entro il …")
    capienza: cfg.capienza, prenotati: c.totale,
    disponibili: cfg.capienza ? Math.max(0, cfg.capienza - c.totale) : null,
    esaurito: cfg.capienza ? c.totale >= cfg.capienza : false,
    turni: null
  };
  if (cfg.turni.length) {
    st.turni = {};
    cfg.turni.forEach(o => {
      const p = c.turni[o] || 0;
      st.turni[o] = {
        capienza: cfg.capienzaTurno, prenotati: p,
        disponibili: cfg.capienzaTurno ? Math.max(0, cfg.capienzaTurno - p) : null,
        esaurito: st.esaurito || (cfg.capienzaTurno ? p >= cfg.capienzaTurno : false)
      };
    });
  }
  return st;
}

function disponibilita_() {
  const out = {};
  Object.values(config_()).forEach(cfg => out[cfg.id] = statoEvento_(cfg));
  return out;
}

// ---- Email -----------------------------------------------------------------------

function confermaOspite_(d, cfg) {
  const persone = `${d.totale} ${d.totale === 1 ? "persona" : "persone"} (${d.dettaglio})`;
  const msg = {
    to: d.email, name: NOME_MITTENTE,
    subject: `Prenotazione ricevuta — ${cfg.evento}`,
    body:
`Ciao ${d.nome},

abbiamo ricevuto la tua prenotazione per "${cfg.evento}" (${d.giorno || ""}${d.orario ? ", ore " + d.orario : ""}) per ${persone}.
${d.note ? "Note: " + d.note + "\n" : ""}
Ti ricontattiamo al ${d.telefono} solo se serve una conferma.

Ci vediamo a Leccio!
Sagra di Penco — sagradipenco.it`
  };
  if (EMAIL_IN_COPIA) msg.cc = EMAIL_IN_COPIA;
  MailApp.sendEmail(msg);
}

function avvisaOrganizzatori_(d, cfg, stato) {
  if (!cfg.email) return;   // nessun avviso se la cella email_avviso in Config è vuota
  const posti = cfg.capienza ? `\nPrenotati finora: ${stato.prenotati + d.totale} su ${cfg.capienza}` : "";
  MailApp.sendEmail({
    to: cfg.email, name: NOME_MITTENTE,
    subject: `Nuova prenotazione: ${cfg.evento} — ${d.nome} (${d.totale})`,
    body:
`Nuova prenotazione dal sito.

Evento:   ${cfg.evento}${d.orario ? " — ore " + d.orario : ""}
Nome:     ${d.nome}
Telefono: ${d.telefono}
Persone:  ${d.totale} (${d.dettaglio})
Email:    ${d.email}
Note:     ${d.note || "-"}${posti}

Foglio: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`
  });
}

// ---- Recap del mattino e avvisi "esaurito" -------------------------------------------

function installaRecap() {
  const gestite = ["recapGiornaliero", "controllaLimiti"];
  ScriptApp.getProjectTriggers().forEach(t => { if (gestite.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("recapGiornaliero").timeBased().everyDays(1).atHour(ORA_RECAP).inTimezone(FUSO).create();
  // scatta quando qualcuno modifica il foglio (es. prenotazione telefonica scritta a mano)
  ScriptApp.newTrigger("controllaLimiti").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  controllaLimiti_(null, true);   // registra la situazione attuale senza mandare avvisi per i pieni già noti
  Logger.log("Recap attivo ogni mattina alle " + ORA_RECAP + ":00 (" + FUSO + ") e avvisi di posti esauriti attivi.");
}

// Funzione chiamata dall'attivatore di modifica del foglio (e utilizzabile a mano per una prova).
function controllaLimiti() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try { controllaLimiti_(null); } finally { lock.releaseLock(); }
}

// Confronta eventi e turni pieni con quelli già segnalati; manda un avviso solo per i nuovi pieni.
// soloId: controlla solo quell'evento (dopo una prenotazione dal sito). silenzioso: aggiorna senza mandare mail.
function controllaLimiti_(soloId, silenzioso) {
  const props = PropertiesService.getScriptProperties();
  const segnalati = new Set(JSON.parse(props.getProperty("esauriti") || "[]"));
  const nuovi = [];
  Object.values(config_()).forEach(cfg => {
    if (soloId && cfg.id !== soloId) return;
    const st = statoEvento_(cfg);
    const voci = [{ chiave: cfg.id, pieno: st.esaurito, testo: cfg.evento, n: st.prenotati, cap: cfg.capienza }];
    if (st.turni) Object.entries(st.turni).forEach(([o, x]) =>
      voci.push({ chiave: cfg.id + "|" + o, pieno: !!(x.capienza && x.prenotati >= x.capienza), testo: `${cfg.evento} — turno delle ${o}`, n: x.prenotati, cap: x.capienza }));
    voci.forEach(v => {
      if (v.pieno && !segnalati.has(v.chiave)) { segnalati.add(v.chiave); nuovi.push(v); }
      if (!v.pieno && segnalati.has(v.chiave)) segnalati.delete(v.chiave);   // riaperto (es. una disdetta): potrà essere segnalato di nuovo
    });
  });
  props.setProperty("esauriti", JSON.stringify([...segnalati]));
  if (silenzioso || !nuovi.length) return;
  const dest = destinatari_();
  if (!dest.length) return;
  const ora = Utilities.formatDate(new Date(), FUSO, "dd/MM 'alle' HH:mm");
  MailApp.sendEmail({
    to: dest.join(","), name: NOME_MITTENTE,
    subject: `Posti esauriti: ${nuovi.map(v => v.testo).join("; ")}`,
    body:
`Raggiunta la capienza (${ora}):

${nuovi.map(v => `- ${v.testo}: ${v.n} prenotati su ${v.cap}`).join("\n")}

Il form sul sito non accetta più prenotazioni per ${nuovi.length === 1 ? "questa voce" : "queste voci"}.
Per riaprire: aumenta la capienza nel tab Config oppure segna "annullata" qualche prenotazione.

Foglio: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`
  });
}

function destinatari_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Destinatari");
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(2).map(r => String(r[0] || "").trim()).filter(e => /@/.test(e));
}

function recapGiornaliero() {
  const dest = destinatari_();
  if (!dest.length) { Logger.log("Nessun destinatario nel tab Destinatari."); return; }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oggi = Utilities.formatDate(new Date(), FUSO, "dd/MM/yyyy");
  // prenotazioni arrivate ieri (da mezzanotte a mezzanotte): ognuna compare in un solo recap
  const fine = new Date(); fine.setHours(0, 0, 0, 0);
  const inizio = new Date(fine); inizio.setDate(inizio.getDate() - 1);
  const GIORNI_IT = ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"];
  const ieri = GIORNI_IT[Number(Utilities.formatDate(inizio, FUSO, "u")) % 7] + " " + Utilities.formatDate(inizio, FUSO, "dd/MM");
  let corpo = `Punto prenotazioni al ${oggi}. Le "nuove" sono quelle arrivate ieri (${ieri}).\n\n`;
  let totaleNuove = 0;

  Object.values(config_()).forEach(cfg => {
    const st = statoEvento_(cfg);
    corpo += `== ${cfg.evento} ==\n`;
    corpo += cfg.capienza
      ? `Prenotati: ${st.prenotati} su ${cfg.capienza} (rimasti ${st.disponibili})${st.esaurito ? " — ESAURITO" : ""}\n`
      : `Prenotati: ${st.prenotati}\n`;
    if (st.turni) Object.entries(st.turni).forEach(([o, x]) =>
      corpo += `  ${o}: ${x.prenotati}${x.capienza ? " su " + x.capienza : ""}${x.esaurito ? " — esaurito" : ""}\n`);
    if (!st.aperto) corpo += "Form online chiuso.\n";

    const t = ss.getSheetByName(cfg.evento);
    const nuove = (t && t.getLastRow() > 1 ? t.getRange(2, 1, t.getLastRow() - 1, COL.length).getValues() : [])
      .filter(r => r[C.Data] instanceof Date && r[C.Data] >= inizio && r[C.Data] < fine && String(r[C.Stato]).toLowerCase() !== "annullata");
    totaleNuove += nuove.length;
    if (nuove.length) {
      corpo += `Nuove di ieri (${nuove.length}):\n`;
      nuove.forEach(r => {
        const n = Number(r[C.Totale]) || (Number(r[C.Q1]) || 0) + (Number(r[C.Q2]) || 0);
        corpo += `  - ${r[C.Nome]} · ${n} ${n === 1 ? "persona" : "persone"} (${r[C.Q1] || 0} ${cfg.campi[0]}, ${r[C.Q2] || 0} ${cfg.campi[1]})` +
                 `${r[C.Orario] ? " · ore " + normalizzaOrario_(r[C.Orario]) : ""} · ${r[C.Telefono]}${r[C.Canale] === "sito" ? "" : " · " + r[C.Canale]}` +
                 `${r[C.Note] ? " · " + r[C.Note] : ""}\n`;
      });
    } else corpo += "Nessuna nuova prenotazione ieri.\n";
    corpo += "\n";
  });
  corpo += `Foglio completo: ${ss.getUrl()}\n\nSagra di Penco — sagradipenco.it`;

  MailApp.sendEmail({
    to: dest.join(","), name: NOME_MITTENTE,
    subject: `Prenotazioni Sagra di Penco — ${oggi} (${totaleNuove} nuove ieri)`,
    body: corpo
  });
}

// ---- Utilità -------------------------------------------------------------------

function parseChiusura_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim().replace("T", " ");
  try { return Utilities.parseDate(s.length > 10 ? s : s + " 23:59", FUSO, "yyyy-MM-dd HH:mm"); }
  catch (e) { return null; }
}

function normalizzaOrario_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, FUSO, "HH:mm");
  const s = String(v || "").trim();
  const m = s.match(/^(\d{1,2})[:.,](\d{2})/);
  return m ? m[1].padStart(2, "0") + ":" + m[2] : s;
}

function risposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Prova dall'editor: inserisce una prenotazione finta nel Pranzo Cinese (poi cancellala dal tab). */
function test() {
  const finto = { postData: { contents: JSON.stringify({
    id: "sab-pranzo-cinese", giorno: "Sabato 17 ottobre", orario: "12:30",
    nome: "Prova Prova", telefono: "000", q1: 2, q2: 1,
    email: Session.getActiveUser().getEmail(), note: "riga di test, cancellare"
  }) } };
  Logger.log(doPost(finto).getContent());
  Logger.log(JSON.stringify(disponibilita_(), null, 1));
}
