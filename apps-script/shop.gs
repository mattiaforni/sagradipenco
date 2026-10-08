/**
 * SHOP DEI VOLONTARI — ordini dalla pagina sagradipenco.it/shop
 * File da aggiungere allo stesso progetto Apps Script delle prenotazioni.
 *
 * INSTALLAZIONE (una volta sola, dall'account sagradipenco@gmail.com)
 *  1. Apri il foglio delle prenotazioni → Estensioni → Apps Script.
 *  2. Nel menu a sinistra, "+" accanto a File → Script → chiamalo "shop" → incolla tutto questo file → salva.
 *  3. Niente da fare: il file delle prenotazioni (apps-script/prenotazioni.gs, versione dell'8/10 o successiva)
 *     smista già gli ordini dello shop all'inizio di doPost. Se usi una versione più vecchia, aggiornala.
 *  4. Seleziona la funzione shopSetup ed esegui (concedi i permessi richiesti): crea i tab "Shop ordini"
 *     e "Shop totali" e attiva l'aggiornamento automatico dei totali quando modifichi il foglio.
 *  5. Esegui shopTest: nel tab "Shop ordini" compare un ordine di prova → controlla e cancella le sue righe
 *     (la prova non consuma numeri: il primo ordine vero sarà S-001).
 *  6. Distribuisci → Gestisci distribuzioni → matita sulla distribuzione attiva → Versione: "Nuova versione" → Distribuisci.
 *     L'indirizzo /exec resta lo stesso, quindi né il sito né lo shop vanno modificati.
 *
 * USO
 *  - "Shop ordini": una riga per ogni articolo ordinato. Le caselle Pagato e Consegnato servono allo stand.
 *    Per annullare un ordine cancellane le righe.
 *  - "Shop totali": pezzi da ordinare a Forme Grafiche per articolo e variante (taglia, colore, tipo).
 *    Lo riscrive lo script a ogni ordine e a ogni modifica del foglio (anche quando cancelli righe).
 *    Per forzare l'aggiornamento: esegui shopAggiornaTotali.
 *  - Se chi ordina lascia l'email, riceve un riepilogo dell'ordine (mittente: l'account che esegue lo script).
 *  - Se cambi un prezzo, cambialo qui in SHOP_CATALOGO e nella pagina shop/index.html.
 */

const SHOP_CATALOGO = {
  maglietta: { nome: "Maglietta", prezzo: 9.50,
               varianti: ["3–4 anni","5–6 anni","7–8 anni","9–11 anni","12–13 anni","S","M","L","XL","XXL","XXXL"] },
  agenda:    { nome: "Agendina", prezzo: 3.00, varianti: ["Rossa","Nera","Bianca"] },
  borraccia: { nome: "Borraccia termica", prezzo: 8.50, varianti: [""] },
  spilla:    { nome: "Spilla", prezzo: 1.50, varianti: ["Coccarda","Penco e paesaggio"] },
  gomma:     { nome: "Gomma", prezzo: 2.50, varianti: [""] },
  matita:    { nome: "Matita", prezzo: 0.80, varianti: [""] }
};
const SHOP_TAB = "Shop ordini";
const SHOP_TAB_TOTALI = "Shop totali";
const SHOP_COLONNE = ["Data","N. ordine","Nome","Cognome","Telefono","Email","Articolo","Variante",
                      "Quantità","Prezzo","Importo","Note","Pagato","Consegnato","Ordinamento"];
const SHOP_FUSO = "Europe/Rome";

/** true se la richiesta POST arriva dallo shop (campo azione = "ordine") */
function shopEOrdine_(e) {
  try { return JSON.parse(e.postData.contents).azione === "ordine"; } catch (err) { return false; }
}

/** registra l'ordine: una riga per articolo nel tab "Shop ordini" */
function shopOrdine_(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.sito) return shopRisposta_({ ok: true, numero: "—" });          // campo trappola anti-spam
    const t = s => String(s || "").trim().slice(0, 200);
    const nome = t(d.nome), cognome = t(d.cognome), telefono = t(d.telefono), email = t(d.email), note = t(d.note);
    if (!nome || !cognome) throw new Error("Mancano nome o cognome");
    if (!Array.isArray(d.righe) || !d.righe.length) throw new Error("Carrello vuoto");

    const ordineArticoli = Object.keys(SHOP_CATALOGO);
    const righe = d.righe.map(r => {
      const p = SHOP_CATALOGO[r.id];
      if (!p) throw new Error("Articolo sconosciuto: " + r.id);
      const variante = String(r.variante || "");
      const iv = p.varianti.indexOf(variante);
      if (iv < 0) throw new Error("Variante non valida per " + p.nome + ": " + variante);
      const qt = Math.floor(Number(r.qt));
      if (!(qt >= 1 && qt <= 99)) throw new Error("Quantità non valida");
      return { p, variante, qt, chiave: (ordineArticoli.indexOf(r.id) + 1) * 100 + iv };
    });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let numero, totale = 0;
    try {
      const sh = shopTab_();
      const props = PropertiesService.getScriptProperties();
      const n = Number(props.getProperty("SHOP_ULTIMO") || 0) + 1;
      props.setProperty("SHOP_ULTIMO", String(n));
      numero = "S-" + ("00" + n).slice(-3);
      const ora = new Date();
      const valori = righe.map(r => {
        totale += r.qt * r.p.prezzo;
        return [ora, numero, nome, cognome, telefono ? "'" + telefono : "", email, r.p.nome, r.variante,
                r.qt, r.p.prezzo, r.qt * r.p.prezzo, note, false, false, r.chiave];
      });
      const inizio = sh.getLastRow() + 1;
      sh.getRange(inizio, 1, valori.length, SHOP_COLONNE.length).setValues(valori);
      sh.getRange(inizio, 13, valori.length, 2).insertCheckboxes();
      sh.getRange(inizio, 1, valori.length, 1).setNumberFormat("dd/MM/yyyy HH:mm");
      sh.getRange(inizio, 10, valori.length, 2).setNumberFormat("€ #,##0.00");
      SpreadsheetApp.flush();
      shopAggiornaTotali();
    } finally {
      lock.releaseLock();
    }
    const mail = shopMailRiepilogo_(email, nome, numero, righe, totale);
    return shopRisposta_({ ok: true, numero: numero, totale: totale, mail: mail });
  } catch (err) {
    return shopRisposta_({ ok: false, errore: String(err && err.message || err) });
  }
}

/** riepilogo via email a chi ordina, solo se ha lasciato un indirizzo valido. Un errore qui non blocca l'ordine. */
function shopMailRiepilogo_(email, nome, numero, righe, totale) {
  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return false;
  try {
    const h = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
    const euro = n => "€ " + n.toFixed(2).replace(".", ",");
    const testo = righe.map(r => `${r.qt} × ${r.p.nome}${r.variante ? " (" + r.variante + ")" : ""} — ${euro(r.qt * r.p.prezzo)}`);
    const html =
      `<div style="font-family:Georgia,serif;color:#2a2622;max-width:520px">` +
      `<p style="font-size:22px;color:#c6352c;font-style:italic;margin:0 0 4px">Sagra di Penco</p>` +
      `<p style="margin:0 0 16px">Ciao ${h(nome)}, abbiamo ricevuto il tuo ordine <b>${numero}</b> dallo shop dei volontari.</p>` +
      `<table style="border-collapse:collapse;width:100%;font-size:15px">` +
      righe.map(r => `<tr><td style="padding:6px 0;border-bottom:1px solid #cdbb8d">${r.qt} × ${h(r.p.nome)}` +
        `${r.variante ? ` <i style="color:#5b544b">(${h(r.variante)})</i>` : ""}</td>` +
        `<td style="padding:6px 0;border-bottom:1px solid #cdbb8d;text-align:right">${euro(r.qt * r.p.prezzo)}</td></tr>`).join("") +
      `<tr><td style="padding:10px 0;font-weight:bold">Totale</td><td style="padding:10px 0;text-align:right;font-weight:bold;color:#204f7a">${euro(totale)}</td></tr>` +
      `</table><p style="margin:16px 0 0">Paghi e ritiri allo stand della Sagra.</p></div>`;
    MailApp.sendEmail({
      to: email, name: "Sagra di Penco", subject: "Sagra di Penco · riepilogo ordine " + numero,
      body: `Ciao ${nome}, abbiamo ricevuto il tuo ordine ${numero} dallo shop dei volontari.\n\n` +
            testo.join("\n") + `\n\nTotale: ${euro(totale)}\n\nPaghi e ritiri allo stand della Sagra.`,
      htmlBody: html
    });
    return true;
  } catch (err) {
    console.error("Mail riepilogo non inviata: " + err);
    return false;
  }
}

function shopRisposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function shopTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHOP_TAB);
  if (!sh) {
    sh = ss.insertSheet(SHOP_TAB);
    sh.getRange(1, 1, 1, SHOP_COLONNE.length).setValues([SHOP_COLONNE]).setFontWeight("bold").setBackground("#f1e6c8");
    sh.setFrozenRows(1);
    sh.hideColumns(SHOP_COLONNE.length);                    // Ordinamento: serve solo ai totali
    sh.setColumnWidth(12, 220);
  }
  return sh;
}

/** crea i due tab dello shop e il trigger che tiene aggiornati i totali. Si può rieseguire: non tocca gli ordini. */
function shopSetup() {
  shopTab_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "shopAggiornaTotali")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("shopAggiornaTotali").forSpreadsheet(ss).onChange().create();
  shopAggiornaTotali();
}

/** riscrive il tab "Shop totali" contando le righe di "Shop ordini" */
function shopAggiornaTotali() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHOP_TAB);
  if (!sh) return;
  const tot = ss.getSheetByName(SHOP_TAB_TOTALI) || ss.insertSheet(SHOP_TAB_TOTALI);

  // ordine di presentazione: articoli e varianti nell'ordine del listino
  const pos = {};
  Object.keys(SHOP_CATALOGO).forEach((id, i) => {
    const p = SHOP_CATALOGO[id];
    p.varianti.forEach((v, j) => { pos[p.nome + "|" + v] = (i + 1) * 100 + j; });
  });

  const n = sh.getLastRow() - 1;
  const dati = n > 0 ? sh.getRange(2, 1, n, SHOP_COLONNE.length).getValues() : [];
  const gruppi = {}, ordini = {};
  dati.forEach(r => {
    const articolo = String(r[6] || "").trim(); if (!articolo) return;
    const variante = String(r[7] || "").trim(), k = articolo + "|" + variante;
    if (!gruppi[k]) gruppi[k] = { articolo: articolo, variante: variante, pezzi: 0, importo: 0, pos: pos[k] || 9999 };
    gruppi[k].pezzi += Number(r[8]) || 0;
    gruppi[k].importo += Number(r[10]) || 0;
    if (r[1]) ordini[r[1]] = true;
  });
  const righe = Object.keys(gruppi).map(k => gruppi[k]).sort((a, b) => a.pos - b.pos);

  tot.clear();
  tot.getRange("A1").setValue("Pezzi da ordinare a Forme Grafiche").setFontWeight("bold").setFontSize(12);
  tot.getRange("A2").setValue("Aggiornato il " + Utilities.formatDate(new Date(), SHOP_FUSO, "dd/MM/yyyy HH:mm") +
                              " · ordini: " + Object.keys(ordini).length).setFontStyle("italic");
  tot.getRange(4, 1, 1, 4).setValues([["Articolo", "Variante", "Pezzi", "Importo"]])
     .setFontWeight("bold").setBackground("#f1e6c8");
  if (righe.length) {
    tot.getRange(5, 1, righe.length, 4).setValues(righe.map(g => [g.articolo, g.variante, g.pezzi, g.importo]));
    const fine = 5 + righe.length;
    tot.getRange(fine, 1, 1, 4).setValues([["Totale", "",
      righe.reduce((s, g) => s + g.pezzi, 0), righe.reduce((s, g) => s + g.importo, 0)]])
      .setFontWeight("bold").setBorder(true, null, null, null, null, null);
    tot.getRange(5, 4, righe.length + 1, 1).setNumberFormat("€ #,##0.00");
  } else {
    tot.getRange("A5").setValue("Nessun ordine.").setFontStyle("italic");
  }
  tot.setColumnWidth(1, 170); tot.setColumnWidth(2, 160);
  tot.setFrozenRows(4);
}

/** ordine di prova: dopo averlo controllato, cancellane le righe nel tab "Shop ordini" */
function shopTest() {
  const props = PropertiesService.getScriptProperties();
  const prima = props.getProperty("SHOP_ULTIMO");          // la prova non consuma numeri d'ordine
  const e ={ postData: { contents: JSON.stringify({
    azione: "ordine", nome: "Prova", cognome: "Shop", telefono: "333 0000000",
    email: Session.getEffectiveUser().getEmail(), note: "ordine di prova",          // il riepilogo di prova arriva a te
    righe: [ { id: "maglietta", variante: "M", qt: 2 }, { id: "agenda", variante: "Nera", qt: 1 },
             { id: "spilla", variante: "Coccarda", qt: 3 }, { id: "matita", variante: "", qt: 1 } ] }) } };
  Logger.log(shopEOrdine_(e));
  Logger.log(shopOrdine_(e).getContent());
  if (prima === null) props.deleteProperty("SHOP_ULTIMO"); else props.setProperty("SHOP_ULTIMO", prima);
}
