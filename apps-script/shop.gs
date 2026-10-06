/**
 * SHOP DEI VOLONTARI — ordini dalla pagina sagradipenco.it/shop
 * File da aggiungere allo stesso progetto Apps Script delle prenotazioni.
 *
 * INSTALLAZIONE (una volta sola, dall'account sagradipenco@gmail.com)
 *  1. Apri il foglio delle prenotazioni → Estensioni → Apps Script.
 *  2. Nel menu a sinistra, "+" accanto a File → Script → chiamalo "shop" → incolla tutto questo file → salva.
 *  3. Apri il file delle prenotazioni e, nella funzione doPost, aggiungi come PRIMA riga dentro le graffe:
 *
 *         if (shopEOrdine_(e)) return shopOrdine_(e);
 *
 *     così:   function doPost(e) {
 *               if (shopEOrdine_(e)) return shopOrdine_(e);
 *               ...resto invariato...
 *     Salva.
 *  4. Seleziona la funzione shopSetup ed esegui: crea i tab "Shop ordini" e "Shop totali".
 *  5. Esegui shopTest: nel tab "Shop ordini" compare un ordine di prova → controlla e cancella le sue righe
 *     (la prova non consuma numeri: il primo ordine vero sarà S-001).
 *  6. Distribuisci → Gestisci distribuzioni → matita sulla distribuzione attiva → Versione: "Nuova versione" → Distribuisci.
 *     L'indirizzo /exec resta lo stesso, quindi né il sito né lo shop vanno modificati.
 *
 * USO
 *  - "Shop ordini": una riga per ogni articolo ordinato. Le caselle Pagato e Consegnato servono allo stand.
 *    Per annullare un ordine cancellane le righe.
 *  - "Shop totali": pezzi da ordinare a Forme Grafiche per articolo e variante (taglia, colore, tipo).
 *    Si aggiorna da solo.
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
    if (!nome || !cognome || !telefono) throw new Error("Mancano nome, cognome o telefono");
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
        return [ora, numero, nome, cognome, "'" + telefono, email, r.p.nome, r.variante,
                r.qt, r.p.prezzo, r.qt * r.p.prezzo, note, false, false, r.chiave];
      });
      const inizio = sh.getLastRow() + 1;
      sh.getRange(inizio, 1, valori.length, SHOP_COLONNE.length).setValues(valori);
      sh.getRange(inizio, 13, valori.length, 2).insertCheckboxes();
      sh.getRange(inizio, 1, valori.length, 1).setNumberFormat("dd/MM/yyyy HH:mm");
      sh.getRange(inizio, 10, valori.length, 2).setNumberFormat("€ #,##0.00");
    } finally {
      lock.releaseLock();
    }
    return shopRisposta_({ ok: true, numero: numero, totale: totale });
  } catch (err) {
    return shopRisposta_({ ok: false, errore: String(err && err.message || err) });
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

/** crea (o ricrea) i due tab dello shop. Si può rieseguire: non tocca gli ordini già registrati. */
function shopSetup() {
  shopTab_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tot = ss.getSheetByName(SHOP_TAB_TOTALI);
  if (!tot) tot = ss.insertSheet(SHOP_TAB_TOTALI);
  tot.clear();
  tot.getRange("B1").setValue("Pezzi da ordinare a Forme Grafiche").setFontWeight("bold").setFontSize(12);
  tot.getRange("F1").setValue("Totale ordini").setFontWeight("bold");
  tot.getRange("G1").setFormula("=SUM('" + SHOP_TAB + "'!K2:K)").setNumberFormat("€ #,##0.00").setFontWeight("bold");
  tot.getRange("A3").setFormula(
    "=QUERY('" + SHOP_TAB + "'!A2:O," +
    "\"select O, G, H, sum(I), sum(K) where G is not null group by O, G, H order by O " +
    "label O '', G 'Articolo', H 'Variante', sum(I) 'Pezzi', sum(K) 'Importo'\", 0)");
  tot.getRange("E4:E200").setNumberFormat("€ #,##0.00");
  tot.getRange("A3:E3").setFontWeight("bold").setBackground("#f1e6c8");
  tot.hideColumns(1);
  tot.setColumnWidth(2, 160); tot.setColumnWidth(3, 160);
  tot.setFrozenRows(3);
}

/** ordine di prova: dopo averlo controllato, cancellane le righe nel tab "Shop ordini" */
function shopTest() {
  const props = PropertiesService.getScriptProperties();
  const prima = props.getProperty("SHOP_ULTIMO");          // la prova non consuma numeri d'ordine
  const e ={ postData: { contents: JSON.stringify({
    azione: "ordine", nome: "Prova", cognome: "Shop", telefono: "333 0000000", email: "", note: "ordine di prova",
    righe: [ { id: "maglietta", variante: "M", qt: 2 }, { id: "agenda", variante: "Nera", qt: 1 },
             { id: "spilla", variante: "Coccarda", qt: 3 }, { id: "matita", variante: "", qt: 1 } ] }) } };
  Logger.log(shopEOrdine_(e));
  Logger.log(shopOrdine_(e).getContent());
  if (prima === null) props.deleteProperty("SHOP_ULTIMO"); else props.setProperty("SHOP_ULTIMO", prima);
}
