// ========================================================================
//          CONFIGURAZIONE - **ADATTARE QUESTI VALORI!**
// ========================================================================
// --- Nomi dei Fogli Google Sheet ---
const NOME_FOGLIO_FATTURE = "Fatture";
const NOME_FOGLIO_APPARTAMENTI = "Appartamenti";
const NOME_FOGLIO_LETTURE_CONSUMI = "Letture";
const NOME_FOGLIO_IMPOSTAZIONI = "Impostazioni";
// --- ID del Foglio di Calcolo ---
const SPREADSHEET_ID = "1GKqcZWPSi7mq7s8Ht-IRwCK641p82xoIplMwixIxfFw";
// --- Tipi Utenza Tecnici ---
const TIPO_ELETTRICO = "elettrico";
const TIPO_ACQUA = "acqua";
const TIPO_GAS = "gas";
// --- Indici Colonne Foglio 'Condomini' (partendo da 1) ---
const NOME_FOGLIO_CONDOMINI = "Condomini"; // Sostituisci con il nome reale del tuo foglio "Condomini"




// --- Indici Colonne Foglio 'Appartamenti' (partendo da 1) ---
const COL_APP_ID = 1;             // Col A: ID Univoco Appartamento
const COL_APP_CONDOMINIO_ID = 2;  // Col B: ID o Nome del Condominio
const COL_APP_INTERNO = 3;        // Col C: Numero/interno (per display)
const COL_APP_PIANO = 4;          // Col D: Piano (per display/indirizzo)
const COL_APP_MQ = 5;             // Col E: Metri Quadri (non usato qui)
const COL_APP_PROPRIETARIO = 6;   // Col F: Proprietario (usato come Nome Display e Intestatario)
const COL_APP_CF_PIVA = 7;        // Col G: Codice Fiscale
const COL_APP_EMAIL = 8;          // Col H: Email (non usato qui)
const COL_APP_TELEFONO = 9;       // Col I: Telefono (non usato qui)
const COL_APP_RESIDENTI = 10;     // Col J: Residenti (non usato qui)
const COL_APP_CONDOMINIO_NOME = 11; // Col K: Nome del Condominio (per visualizzazione)

// --- Indici Colonne Foglio 'Letture' ---
const COL_LETT_APP_ID = 2; 
const COL_LETT_TIPO_CONT = 3; 
const COL_LETT_DATA = 4; 
const COL_LETT_CONSUMO = 6;

// --- Indici Colonne Foglio 'Impostazioni' ---
const COL_IMP_TAR_TIPO = 1; 
const COL_IMP_TAR_PREZZO_UNIT = 2; 
const COL_IMP_TAR_UNITA = 3; 
const COL_IMP_TAR_COSTO_FISSO = 4; 
const COL_IMP_TAR_IVA = 5; 
const COL_IMP_TAR_NOME_DISPLAY = 6;
const COL_IMP_GEN_DESCRIZIONE = 1; 
const COL_IMP_GEN_VALORE = 2;

// --- Chiavi per Impostazioni Generali ---
const KEY_NOME_AZIENDA = "Nome Azienda/Cond."; 
const KEY_INDIRIZZO_AZIENDA = "Indirizzo Azienda"; 
const KEY_PIVA_CF_AZIENDA = "P.IVA / CF Azienda"; 
const KEY_EMAIL = "Email Contatto"; 
const KEY_TELEFONO = "Telefono Contatto"; 
const KEY_IBAN = "IBAN Pagamenti"; 
const KEY_GIORNI_SCADENZA = "Giorni Scadenza Fatt."; 
const KEY_NOTE_FATTURA = "Note Predefinite Fatt."; 
const KEY_STATO_DEFAULT = "Stato Default Fattura"; 
const KEY_PREFISSO_ID = "Prefisso ID Fattura";

// --- Indici Colonne Foglio 'Fatture' ---
const COL_FATT_ID = 1; 
const COL_FATT_APP_ID = 2; 
const COL_FATT_DATA_INIZIO = 3; 
const COL_FATT_DATA_FINE = 4; 
const COL_FATT_DATA_EMISSIONE = 5; 
const COL_FATT_DATA_SCADENZA = 6; 
const COL_FATT_CONS_ELETTRICO = 7; 
const COL_FATT_IMP_ELETTRICO = 8; 
const COL_FATT_CONS_ACQUA = 9; 
const COL_FATT_IMP_ACQUA = 10; 
const COL_FATT_CONS_GAS = 11; 
const COL_FATT_IMP_GAS = 12; 
const COL_FATT_IMP_TOTALE = 13; 
const COL_FATT_STATO = 14; 
const COL_FATT_DATA_CREAZIONE = 15;
const COL_FATT_METODO_PAGAMENTO = 16; // Nuova colonna per il metodo di pagamento

// ========================================================================
//                      CODICE APPLICAZIONE WEB
// ========================================================================
/**
 * Gestisce il routing delle pagine dell'applicazione web
 */
function doGet(e) {
  try {
    // Gestione parametri URL
    const params = e.parameter;
    
    // Visualizzazione PDF o HTML
    if (params.view && params.id) {
      if (params.view === 'pdf') {
        // Mantieni la funzionalità PDF esistente
        return generaFatturaPdf(params.id);
      } else if (params.view === 'html') {
        // Nuova modalità di visualizzazione HTML
        // Utilizza la stessa funzione che genererebbe il PDF ma senza convertirlo
        const htmlOutput = generaHtmlFattura(params.id);
        return htmlOutput;
      }
    }
    
    // Pagina di elenco fatture
    if (params.page && params.page === 'list') {
      return HtmlService.createTemplateFromFile('ElencaFatture')
        .evaluate()
        .setTitle('Elenco Fatture')
        .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/forms/favicon_qp2.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    // Pagina nuova fattura (default)
    return HtmlService.createTemplateFromFile('Fatture')
      .evaluate()
      .setTitle('Generatore Fatture')
      .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/forms/favicon_qp2.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    Logger.log(`Errore in doGet: ${error}`);
    return HtmlService.createHtmlOutput(`<h1>Errore</h1><p>${error.message || error}</p>`)
      .setTitle('Errore')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * Aggiorna lo stato delle fatture scadute
 * Da impostare come trigger giornaliero
 */
function aggiornaStatoFattureScadute() {
  try {
    const ss = getSpreadsheet();
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    
    if (!sheetFatture) {
      Logger.log("Foglio fatture non trovato");
      return;
    }
    
    const dati = sheetFatture.getDataRange().getValues();
    
    // Indice delle colonne da intestazioni
    const colDataScadenza = 5; // Colonna F - Data Scadenza
    const colStato = 14; // Colonna O - Stato
    
    const oggi = new Date();
    oggi.setHours(0, 0, 0, 0); // Normalizza la data di oggi
    
    // Conta quante fatture sono state aggiornate
    let conteggioAggiornate = 0;
    
    // Controlla ogni fattura
    for (let i = 1; i < dati.length; i++) {
      const dataScadenza = dati[i][colDataScadenza];
      const statoAttuale = dati[i][colStato];
      
      // Verifica se la fattura è scaduta e lo stato è "Da Pagare"
      if (dataScadenza instanceof Date && 
          !isNaN(dataScadenza.getTime()) &&
          dataScadenza < oggi && 
          statoAttuale === "Da Pagare") {
        
        // Aggiorna lo stato a "Scaduta"
        sheetFatture.getRange(i + 1, colStato + 1).setValue("Scaduta");
        conteggioAggiornate++;
      }
    }
    
    Logger.log(`Aggiornamento completato. ${conteggioAggiornate} fatture sono diventate "Scadute"`);
    return conteggioAggiornate;
  } catch (error) {
    Logger.log("Errore nell'aggiornamento degli stati: " + error);
    return 0;
  }
}

/**
 * Genera l'HTML della fattura per l'anteprima usando lo stesso template della generazione automatica
 * @param {string} fatturaId - ID della fattura
 * @returns {string} - HTML della fattura
 */
function getHtmlFatturaAnteprima(fatturaId) {
  try {
    // Ottieni i dati della fattura dal foglio
    const ss = getSpreadsheet();
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    
    if (!sheetFatture) throw new Error("Foglio fatture non trovato");
    
    // Cerca la fattura per ID
    const dati = sheetFatture.getDataRange().getValues();
    let fatturaRow = null;
    for (let i = 1; i < dati.length; i++) {
      if (String(dati[i][0]) === String(fatturaId)) {
        fatturaRow = dati[i];
        break;
      }
    }
    
    if (!fatturaRow) throw new Error(`Fattura con ID ${fatturaId} non trovata`);
    
    // Recupera tutti i dati necessari per la fattura, inclusi campi aggiuntivi
    const idAppartamento = fatturaRow[1];
    const datiAppartamento = getDatiAppartamentoById(ss, idAppartamento);
    const impostazioni = leggiImpostazioniComplete();
    
    // Preparazione dati completi per includere tutti i campi come nella fattura mostrata
    const datiCompleti = {
      id: fatturaId,
      idAppartamento: idAppartamento,
      intestatario: datiAppartamento?.intestatario || "N/D",
      indirizzo: datiAppartamento?.indirizzo || "N/D",
      cfPiva: datiAppartamento?.cfPiva || "N/D",
      dataInizio: fatturaRow[2],
      dataFine: fatturaRow[3],
      dataEmissione: fatturaRow[4],
      dataScadenza: fatturaRow[5],
      consumoElettrico: fatturaRow[6] || 0,
      importoElettrico: fatturaRow[7] || 0,
      consumoAcqua: fatturaRow[8] || 0,
      importoAcqua: fatturaRow[9] || 0,
      consumoGas: fatturaRow[10] || 0,
      importoGas: fatturaRow[11] || 0,
      importoTotale: fatturaRow[12] || 0,
      stato: fatturaRow[14] || "Da Pagare",
      metodoPagamento: fatturaRow[15] || "Bonifico Bancario",
      noteImportanti: fatturaRow[16] || "",
      scontoPercentuale: fatturaRow[17] || 0,
      motivoSconto: fatturaRow[18] || "",
      aziendaInfo: impostazioni.generali
    };
    
    // Genera l'HTML esattamente come nella generazione automatica
    // Usa la stessa funzione che genera la fattura come mostrato nell'immagine
    const htmlContent = creaHtmlFattura(datiCompleti);
    
    return htmlContent;
  } catch (error) {
    Logger.log("Errore in getHtmlFatturaAnteprima: " + error);
    return `<div class="alert alert-danger">Errore: ${error.message || "Errore sconosciuto"}</div>`;
  }
}

/**
 * Ottiene il foglio di calcolo, con gestione robusta degli errori
 */
function getSpreadsheet() {
  try {
    // Prima tenta di ottenere il foglio attivo (se lo script è container-bound)
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Se fallisce, usa l'ID specifico
    if (!ss) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    if (!ss) {
      throw new Error("Impossibile accedere al foglio di calcolo");
    }
    
    return ss;
  } catch (error) {
    Logger.log("Errore nell'accesso al foglio di calcolo: " + error);
    throw error;
  }
}

// ========================================================================
//           FUNZIONI CHIAMABILI DAL JAVASCRIPT DEL FRONTEND
// ========================================================================

/**
 * Restituisce l'elenco univoco dei condomini con ID e nome.
 * @returns {Array<Object>} Array di oggetti {id, nome} per i condomini
 */
function getCondominiList() {
    try {
        const ss = getSpreadsheet();
        const sheet = ss.getSheetByName(NOME_FOGLIO_APPARTAMENTI);
        if (!sheet) throw new Error(`Foglio "${NOME_FOGLIO_APPARTAMENTI}" non trovato.`);

        // Leggiamo sia ID che Nome del condominio
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return []; // Solo intestazione o foglio vuoto
        
        const idRange = sheet.getRange(2, COL_APP_CONDOMINIO_ID, lastRow - 1, 1);
        const idValues = idRange.getValues();
        
        // Creiamo una mappa per tenere traccia dei condomini univoci
        const condominiMap = new Map();
        
        for (let i = 0; i < idValues.length; i++) {
            const id = idValues[i][0];
            const nome = getCondominioNomeById(id) || `Condominio ${id}`; // Se il nome è vuoto, usiamo ID con prefisso
            
            if (id && !condominiMap.has(id)) {
                condominiMap.set(id, {
                    id: id,
                    nome: nome
                });
            }
        }
        
        // Convertiamo la mappa in array
        const condomini = Array.from(condominiMap.values());
        
        // Ordiniamo per nome
        condomini.sort((a, b) => String(a.nome).localeCompare(String(b.nome)));
        
        Logger.log(`Trovati ${condomini.length} Condomini univoci.`);
        return condomini;
    } catch (error) {
        Logger.log(`Errore GRAVE in getCondominiList: ${error}`);
        return [];
    }
}
function getCondominioNomeById(condominioId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(NOME_FOGLIO_CONDOMINI); // Assicurati di avere questa costante definita

    if (!sheet) throw new Error(`Foglio "${NOME_FOGLIO_CONDOMINI}" non trovato.`);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null; // Foglio vuoto o solo intestazione

    // Trova la colonna dove si trova l'ID (ipotizziamo sia la colonna A, quindi 1)
    const idColumn = 1; // Controlla che sia corretto nel foglio Condomini
    const nomeColumn = 2; // Ipotizziamo che il nome sia nella colonna B (2) // Controlla che sia corretto nel foglio Condomini

    // Leggi tutti gli ID dal foglio Condomini
    const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
    const idValues = idRange.getValues();

    // Leggi tutti i nomi dal foglio Condomini
    const nomeRange = sheet.getRange(2, nomeColumn, lastRow - 1, 1);
    const nomeValues = nomeRange.getValues();


    // Cerca l'ID corrispondente e restituisci il nome
    for (let i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0]) === String(condominioId)) { // Confronta come stringhe per sicurezza
        return nomeValues[i][0];
      }
    }

    return null; // Condominio non trovato
  } catch (error) {
    Logger.log(`Errore in getCondominioNomeById: ${error}`);
    return null;
  }
}
/**
 * Restituisce TUTTI i dati rilevanti degli appartamenti per il filtraggio frontend.
 * @returns {Array<Object>} Array di oggetti {id, condominioId, nomeDisplay} o array vuoto.
 */
function getAppartamentiData() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(NOME_FOGLIO_APPARTAMENTI);
    if (!sheet) throw new Error(`Foglio "${NOME_FOGLIO_APPARTAMENTI}" non trovato.`);

    // Legge le colonne necessarie: ID, Condominio ID, Interno, Proprietario
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(COL_APP_ID, COL_APP_CONDOMINIO_ID, COL_APP_INTERNO, COL_APP_PROPRIETARIO));
    const values = range.getValues();
    const appartamenti = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const idApp = row[COL_APP_ID - 1];
      const condoId = row[COL_APP_CONDOMINIO_ID - 1];
      const interno = row[COL_APP_INTERNO - 1];
      const proprietario = row[COL_APP_PROPRIETARIO - 1];

      if (idApp && condoId) { // Richiede ID Appartamento e ID Condominio
         // Costruisce il nome da visualizzare (es. "Proprietario - Int. X")
         let nomeDisplay = proprietario || `ID: ${idApp}`; // Fallback se manca proprietario
         if(interno) {
             nomeDisplay += ` (Int. ${interno})`;
         }

        appartamenti.push({
          id: idApp,
          condominioId: condoId, // Aggiunge ID Condominio
          nome: nomeDisplay      // Nome da mostrare nel secondo dropdown
        });
      }
    }
    Logger.log(`Recuperati ${appartamenti.length} appartamenti con dati completi.`);
    return appartamenti;
  } catch (error) {
    Logger.log(`Errore GRAVE in getAppartamentiData: ${error}`);
    return [];
  }
}

/**
 * Recupera tutte le impostazioni e tariffe necessarie.
 * Ristrutturato in un formato più facilmente utilizzabile dal frontend.
 */
function leggiImpostazioniComplete() {
  try {
    const ss = getSpreadsheet();
    const sheetImpostazioni = ss.getSheetByName(NOME_FOGLIO_IMPOSTAZIONI);
    if (!sheetImpostazioni) { 
      Logger.log(`Foglio ${NOME_FOGLIO_IMPOSTAZIONI} non trovato`); 
      return { 
        error: `Foglio "${NOME_FOGLIO_IMPOSTAZIONI}" non trovato`,
        tariffe: { elettrico: {}, acqua: {}, gas: {} },
        generali: {}
      };
    }
    
    // Legge tutti i dati dal foglio impostazioni
    const valuesAll = sheetImpostazioni.getDataRange().getValues();
    
    // Oggetto per memorizzare risultati
    const result = {
      tariffe: {
        elettrico: { tipo: TIPO_ELETTRICO },
        acqua: { tipo: TIPO_ACQUA },
        gas: { tipo: TIPO_GAS }
      },
      generali: {}
    };
    
    // Elaborazione dati foglio
    for (let i = 1; i < valuesAll.length; i++) { // Parte da 1 per saltare intestazione
      const row = valuesAll[i];
      
      // Se è una riga di tariffe (riconoscibile dal tipo in prima colonna)
      if (row[COL_IMP_TAR_TIPO - 1] === TIPO_ELETTRICO || 
          row[COL_IMP_TAR_TIPO - 1] === TIPO_ACQUA || 
          row[COL_IMP_TAR_TIPO - 1] === TIPO_GAS) {
        
        const tipoTariffa = row[COL_IMP_TAR_TIPO - 1];
        result.tariffe[tipoTariffa] = {
          tipo: tipoTariffa,
          prezzoUnitario: row[COL_IMP_TAR_PREZZO_UNIT - 1] || 0,
          unitaMisura: row[COL_IMP_TAR_UNITA - 1] || "",
          costoFisso: row[COL_IMP_TAR_COSTO_FISSO - 1] || 0,
          iva: row[COL_IMP_TAR_IVA - 1] || 0,
          nomeDisplay: row[COL_IMP_TAR_NOME_DISPLAY - 1] || tipoTariffa
        };
      }
      // Se è una riga di impostazioni generali (no tipo ma valore presente in seconda colonna)
      else if (row[COL_IMP_GEN_DESCRIZIONE - 1] && row[COL_IMP_GEN_VALORE - 1] !== undefined) {
        const chiave = row[COL_IMP_GEN_DESCRIZIONE - 1];
        const valore = row[COL_IMP_GEN_VALORE - 1];
        
        switch(chiave) {
          case KEY_NOME_AZIENDA:
            result.generali.nomeAzienda = valore;
            break;
          case KEY_INDIRIZZO_AZIENDA:
            result.generali.indirizzoAzienda = valore;
            break;
          case KEY_PIVA_CF_AZIENDA:
            result.generali.pivaCfAzienda = valore;
            break;
          case KEY_EMAIL:
            result.generali.emailContatto = valore;
            break;
          case KEY_TELEFONO:
            result.generali.telefonoContatto = valore;
            break;
          case KEY_IBAN:
            result.generali.iban = valore;
            break;
          case KEY_GIORNI_SCADENZA:
            result.generali.giorniScadenza = valore;
            break;
          case KEY_NOTE_FATTURA:
            result.generali.notePredefinite = valore;
            break;
          case KEY_STATO_DEFAULT:
            result.generali.statoDefaultFattura = valore;
            break;
          case KEY_PREFISSO_ID:
            result.generali.prefissoIdFattura = valore;
            break;
          default:
            // Memorizza anche chiavi non standard
            result.generali[chiave] = valore;
        }
      }
    }
    
    return result;
  } catch (error) {
    Logger.log(`Errore in leggiImpostazioniComplete: ${error}`);
    return { 
      error: `Errore: ${error.message || error}`,
      tariffe: { elettrico: {}, acqua: {}, gas: {} },
      generali: {}
    };
  }
}

/**
 * Genera una nuova fattura a partire dall'ID appartamento e dalla data fine periodo
 * @param {string} idAppartamento - ID dell'appartamento
 * @param {string} dataFineStr - Data fine periodo in formato stringa
 * @param {string} metodoPagamento - Metodo di pagamento selezionato
 * @returns {Object} - Dati fattura generata
 */
function generaERegistraFattura(idAppartamento, dataFineStr, metodoPagamento, noteImportanti, scontoPercentuale, motivoSconto) {
  try {
    // 0. Controllo parametri di input
    if (!idAppartamento) throw new Error("ID Appartamento non specificato");
    if (!dataFineStr) throw new Error("Data Fine Periodo non specificata");
    
    // Converti string to date e formatta data
    let dataFine = new Date(dataFineStr);
    if (isNaN(dataFine.getTime())) throw new Error(`Data Fine Periodo non valida: ${dataFineStr}`);
    const oggi = new Date();
    
    // 1. Recupera SpreadsheetApp
    const ss = getSpreadsheet();
    if (!ss) throw new Error("Foglio di calcolo non trovato");
    
    // 2. Recupera dati appartamento dall'ID
    const datiAppartamento = getDatiAppartamentoById(ss, idAppartamento);
    if (!datiAppartamento) throw new Error(`Appartamento con ID ${idAppartamento} non trovato`);
    
    // 3. Recupera impostazioni (tariffe, giorni scadenza, ecc.)
    const impostazioni = leggiImpostazioniComplete();
    if (!impostazioni) throw new Error("Impossibile leggere le impostazioni necessarie");
    
    // 4. Trova la data di inizio (ultima fattura o 1 mese indietro)
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    if (!sheetFatture) throw new Error(`Foglio ${NOME_FOGLIO_FATTURE} non trovato`);
    
    // Trova l'ultima fattura per l'appartamento specificato
    const valoriFatture = sheetFatture.getDataRange().getValues();
    let dataInizio = new Date(dataFine);
    dataInizio.setMonth(dataInizio.getMonth() - 1); // Default: 1 mese prima
    
    // Cerca l'ultima fattura per questo appartamento
    for (let i = 1; i < valoriFatture.length; i++) {
      if (valoriFatture[i][COL_FATT_APP_ID - 1] == idAppartamento) {
        const dataFinePrecedente = new Date(valoriFatture[i][COL_FATT_DATA_FINE - 1]);
        // Utilizza la data successiva all'ultima fattura come inizio
        if (!isNaN(dataFinePrecedente.getTime())) {
          dataInizio = new Date(dataFinePrecedente);
          dataInizio.setDate(dataInizio.getDate() + 1);
          break;
        }
      }
    }
    
    // 5. Trova letture consumo per il periodo specificato
    const sheetLetture = ss.getSheetByName(NOME_FOGLIO_LETTURE_CONSUMI);
    if (!sheetLetture) throw new Error(`Foglio ${NOME_FOGLIO_LETTURE_CONSUMI} non trovato`);
    const valuesLetture = sheetLetture.getDataRange().getValues();
    
    // Trovare tutti i consumi nel periodo
    let tutteLetture = [];
    for (let i = 1; i < valuesLetture.length; i++) {
      // Includi solo letture per questo appartamento
      if (valuesLetture[i][COL_LETT_APP_ID - 1] == idAppartamento) {
        const dataLettura = new Date(valuesLetture[i][COL_LETT_DATA - 1]);
        tutteLetture.push({
          tipo: valuesLetture[i][COL_LETT_TIPO_CONT - 1],
          data: dataLettura,
          consumo: valuesLetture[i][COL_LETT_CONSUMO - 1] || 0
        });
      }
    }
    
    // 6. Calcola consumi
    const consumoElettrico = trovaConsumoPrecalcolato(idAppartamento, TIPO_ELETTRICO, dataFine, tutteLetture);
    const consumoAcqua = trovaConsumoPrecalcolato(idAppartamento, TIPO_ACQUA, dataFine, tutteLetture);
    const consumoGas = trovaConsumoPrecalcolato(idAppartamento, TIPO_GAS, dataFine, tutteLetture);
    
   // 7. Calcola importi in base alle tariffe e applica lo sconto se presente
    const sconto = (scontoPercentuale / 100) || 0;

// Importi base
    const importoElettricoBase = calcolaImportoSingolaUtenza(consumoElettrico, impostazioni.tariffe.elettrico);
    const importoAcquaBase = calcolaImportoSingolaUtenza(consumoAcqua, impostazioni.tariffe.acqua);
    const importoGasBase = calcolaImportoSingolaUtenza(consumoGas, impostazioni.tariffe.gas);

// Importi con sconto applicato
    const importoElettrico = importoElettricoBase * (1 - sconto);
    const importoAcqua = importoAcquaBase * (1 - sconto);
    const importoGas = importoGasBase * (1 - sconto);
    const importoTotale = importoElettrico + importoAcqua + importoGas;
    
    // 8. Genera data scadenza
    const dataEmissione = oggi;
    const dataScadenza = new Date(oggi);
    dataScadenza.setDate(dataScadenza.getDate() + (parseInt(impostazioni.generali.giorniScadenza) || 30));
    
    // 9. Genera nuovo ID fattura
    const nuovoId = (impostazioni.generali.prefissoIdFattura || "FATT-") + generaProssimoIdFatturaNumerico(ss);
    
    // 10. Registra fattura in foglio 'Fatture'
   const nuovaRiga = [
  nuovoId,                                // ID Fattura
  idAppartamento,                         // ID Appartamento
  Utilities.formatDate(dataInizio, "GMT+1", "yyyy-MM-dd"), // Data Inizio
  Utilities.formatDate(dataFine, "GMT+1", "yyyy-MM-dd"),   // Data Fine
  Utilities.formatDate(dataEmissione, "GMT+1", "yyyy-MM-dd"), // Data Emissione
  Utilities.formatDate(dataScadenza, "GMT+1", "yyyy-MM-dd"),  // Data Scadenza
  consumoElettrico || 0,                  // Consumo Elettrico
  importoElettrico || 0,                  // Importo Elettrico
  consumoAcqua || 0,                      // Consumo Acqua
  importoAcqua || 0,                      // Importo Acqua
  consumoGas || 0,                        // Consumo Gas
  importoGas || 0,                        // Importo Gas
  importoTotale || 0,                     // Importo Totale
  impostazioni.generali.statoDefaultFattura || "Da Pagare", // Stato
  Utilities.formatDate(oggi, "GMT+1", "yyyy-MM-dd HH:mm:ss"), // Data Creazione
  metodoPagamento || "Non specificato",   // Metodo di Pagamento
  noteImportanti || "",                   // Note Importanti
  scontoPercentuale || 0,                 // Sconto Percentuale
  motivoSconto || ""                      // Motivo Sconto
];
    
    // Assicurati che la prima riga abbia l'intestazione per il nuovo campo
    // Se il foglio è vuoto o ha solo l'intestazione, aggiungi/aggiorna l'intestazione
    if (sheetFatture.getLastRow() <= 1) {
      // Ottieni l'intestazione corrente o crea una nuova
      let headers = [];
      if (sheetFatture.getLastRow() === 1) {
        headers = sheetFatture.getRange(1, 1, 1, sheetFatture.getLastColumn()).getValues()[0];
      } else {
        headers = [
          "ID Fattura", "ID Appartamento", "Data Inizio", "Data Fine", 
          "Data Emissione", "Data Scadenza", "Consumo Elettrico", "Importo Elettrico",
          "Consumo Acqua", "Importo Acqua", "Consumo Gas", "Importo Gas",
          "Importo Totale", "Stato", "Data Creazione", "Note Importanti"
        ];
      }
      
      // Aggiungi l'intestazione per il metodo di pagamento se non esiste
      if (headers.length < COL_FATT_METODO_PAGAMENTO || !headers[COL_FATT_METODO_PAGAMENTO - 1]) {
        headers[COL_FATT_METODO_PAGAMENTO - 1] = "Metodo Pagamento";
        sheetFatture.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
      // Aggiungi intestazioni per Note Importanti, Sconto e Motivo Sconto se non esistono
  const COL_FATT_NOTE_IMPORTANTI = COL_FATT_METODO_PAGAMENTO + 1;
  const COL_FATT_SCONTO_PERCENTUALE = COL_FATT_METODO_PAGAMENTO + 2;
  const COL_FATT_MOTIVO_SCONTO = COL_FATT_METODO_PAGAMENTO + 3;
  
  if (headers.length < COL_FATT_NOTE_IMPORTANTI || !headers[COL_FATT_NOTE_IMPORTANTI - 1]) {
    headers[COL_FATT_NOTE_IMPORTANTI - 1] = "Note Importanti";
  }
  
  if (headers.length < COL_FATT_SCONTO_PERCENTUALE || !headers[COL_FATT_SCONTO_PERCENTUALE - 1]) {
    headers[COL_FATT_SCONTO_PERCENTUALE - 1] = "Sconto Percentuale";
  }
  
  if (headers.length < COL_FATT_MOTIVO_SCONTO || !headers[COL_FATT_MOTIVO_SCONTO - 1]) {
    headers[COL_FATT_MOTIVO_SCONTO - 1] = "Motivo Sconto";
    sheetFatture.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
    }
    
    sheetFatture.appendRow(nuovaRiga);
    
    // 11. Prepara i dati da restituire per l'anteprima
    return {
      id: nuovoId,
      idAppartamento: idAppartamento,
      intestatario: datiAppartamento.intestatario,
      indirizzo: datiAppartamento.indirizzo,
      cfPiva: datiAppartamento.cfPiva,
      email: datiAppartamento.email,
      dataInizio: Utilities.formatDate(dataInizio, "GMT+1", "yyyy-MM-dd"),
      dataFine: Utilities.formatDate(dataFine, "GMT+1", "yyyy-MM-dd"),
      dataEmissione: Utilities.formatDate(dataEmissione, "GMT+1", "yyyy-MM-dd"),
      dataScadenza: Utilities.formatDate(dataScadenza, "GMT+1", "yyyy-MM-dd"),
      consumoElettrico: consumoElettrico || 0,
      importoElettrico: importoElettrico || 0,
      consumoAcqua: consumoAcqua || 0,
      importoAcqua: importoAcqua || 0,
      consumoGas: consumoGas || 0,
      importoGas: importoGas || 0,
      importoTotale: importoTotale || 0,
      stato: impostazioni.generali.statoDefaultFattura || "Da Pagare",
      metodoPagamento: metodoPagamento,
      noteImportanti: noteImportanti || "",
      scontoPercentuale: scontoPercentuale || 0,
      motivoSconto: motivoSconto || "",
      aziendaInfo: {
        nome: impostazioni.generali.nomeAzienda,
        indirizzo: impostazioni.generali.indirizzoAzienda,
        pivaCf: impostazioni.generali.pivaCfAzienda,
        email: impostazioni.generali.emailContatto,
        telefono: impostazioni.generali.telefonoContatto,
        iban: impostazioni.generali.iban,
        note: impostazioni.generali.notePredefinite
      }
    };
    
  } catch (error) {
    Logger.log(`Errore in generaERegistraFattura: ${error}`);
    return { error: `Errore: ${error.message || error}` };
  }
}

/**
 * Ottiene l'elenco completo delle fatture con dati relativi.
 * Versione modificata per compatibilità con il foglio dati mostrato.
 * @returns {Array<Object>} Lista di fatture formattate
 */
function getFattureList() {
  try {
    const ss = getSpreadsheet();
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    if (!sheetFatture) {
      Logger.log(`Foglio ${NOME_FOGLIO_FATTURE} non trovato`);
      return [];
    }
    
    // Leggi tutte le righe (esclusa intestazione)
    const lastRow = sheetFatture.getLastRow();
    if (lastRow <= 1) {
      Logger.log("Solo intestazione o foglio vuoto");
      return []; // Solo intestazione o foglio vuoto
    }
    
    // Adattamento per compatibilità: leggi TUTTE le colonne dal foglio
    Logger.log(`Lettura dati fatture da ${lastRow-1} righe`);
    const values = sheetFatture.getRange(2, 1, lastRow - 1, sheetFatture.getLastColumn()).getValues();
    const fatture = [];
    
    for (let i = 0; i < values.length; i++) {
      try {
        const row = values[i];
        
        // Si aspetta ID Fattura nella prima colonna
        const idFattura = row[0]; // Prima colonna nel foglio
        if (!idFattura) {
          Logger.log(`Riga ${i+2} senza ID fattura, saltata`);
          continue; // Salta righe senza ID
        }

        // Ottieni gli indici delle colonne necessarie dal foglio mostrato nell'immagine
        // NOTA: adatto agli indici mostrati nel tuo foglio, non quelli definiti nelle costanti
        const COL_APP_ID = 1; // Colonna appartamento (B)
        const COL_DATA_INIZIO = 2; // Data inizio (C)
        const COL_DATA_FINE = 3;   // Data fine (D)
        const COL_DATA_EMISSIONE = 4; // Data emissione (E)
        const COL_DATA_SCADENZA = 5;  // Data scadenza (F)
        const COL_TOTALE = 12;       // Importo totale (M)
        const COL_STATO = 14;        // Stato (O)
        
        // In base all'immagine, prendi i dati dalle colonne corrette
        const idAppartamento = row[COL_APP_ID]; 
        
        // Recupera dati appartamento
        const datiAppartamento = getDatiAppartamentoById(ss, idAppartamento);
        
        // Converti le date in formato stringa se sono oggetti Date
        const formatDate = (date) => {
          if (!date) return "";
          if (typeof date === "string") return date;
          if (date instanceof Date && !isNaN(date.getTime())) {
            return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          return "";
        };
        
        // Crea oggetto fattura
        const importoTotale = typeof row[COL_TOTALE] === 'number' ? row[COL_TOTALE] : 0;
        const stato = row[COL_STATO] || "Da Pagare";
        
        fatture.push({
          id: idFattura.toString(),
          idAppartamento: idAppartamento ? idAppartamento.toString() : "",
          intestatario: datiAppartamento ? datiAppartamento.intestatario : "Intestatario sconosciuto",
          indirizzo: datiAppartamento ? datiAppartamento.indirizzo : "Indirizzo sconosciuto",
          email: datiAppartamento ? datiAppartamento.email : "",
          dataInizio: formatDate(row[COL_DATA_INIZIO]),
          dataFine: formatDate(row[COL_DATA_FINE]),
          dataEmissione: formatDate(row[COL_DATA_EMISSIONE]),
          dataScadenza: formatDate(row[COL_DATA_SCADENZA]),
          importoElettrico: 0, // Fallback
          importoAcqua: 0,     // Fallback 
          importoGas: 0,       // Fallback
          importoTotale: importoTotale,
          stato: stato,
          metodoPagamento: "Bonifico Bancario" // Valore predefinito
        });
      } catch (rowError) {
        Logger.log(`Errore nell'elaborazione della riga ${i+2}: ${rowError}`);
        // Continua con la prossima riga
      }
    }
    
    Logger.log(`Recuperate ${fatture.length} fatture con successo.`);
    return fatture;
  } catch (error) {
    Logger.log(`Errore GRAVE in getFattureList: ${error}`);
    return [];
  }
}

/**
 * Aggiorna lo stato di una fattura
 * @param {Object} statoData - Dati per l'aggiornamento dello stato
 * @returns {Object} Risultato dell'operazione
 */
function aggiornaStatoFattura(statoData) {
  try {
    // Validazione
    if (!statoData || !statoData.fatturaId) {
      return { success: false, error: "ID Fattura non specificato" };
    }
    
    if (!statoData.stato) {
      return { success: false, error: "Nuovo stato non specificato" };
    }
    
    // Accedi al foglio
    const ss = getSpreadsheet();
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    if (!sheetFatture) {
      return { success: false, error: `Foglio ${NOME_FOGLIO_FATTURE} non trovato` };
    }
    
    // Trova la riga della fattura
    const values = sheetFatture.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][COL_FATT_ID - 1] === statoData.fatturaId) {
        rowIndex = i + 1; // +1 perché gli indici di riga partono da 1
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: `Fattura con ID ${statoData.fatturaId} non trovata` };
    }
    
    // Aggiorna lo stato
    sheetFatture.getRange(rowIndex, COL_FATT_STATO).setValue(statoData.stato);
    
    // Opzionale: registra le note se presenti
    if (statoData.note) {
      // Aggiungi una riga nel foglio delle note o aggiungi una colonna per le note nel foglio fatture
      Logger.log(`Note per fattura ${statoData.fatturaId}: ${statoData.note}`);
    }
    
    return { success: true };
  } catch (error) {
    Logger.log(`Errore in aggiornaStatoFattura: ${error}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Invia una fattura via email come PDF
 * @param {Object} emailData - Dati dell'email e della fattura
 * @returns {Object} - Risultato dell'operazione
 */
function inviaFatturaViaEmail(emailData) {
  try {
    // Valida i dati
    if (!emailData || !emailData.destinatario) {
      return { success: false, error: "Destinatario email non specificato" };
    }
    
    if (!emailData.fatturaId) {
      return { success: false, error: "ID Fattura non specificato" };
    }
    
    // 1. Prima genera il PDF della fattura
    const pdfBlob = generaPdfFattura(emailData);
    if (!pdfBlob) {
      return { success: false, error: "Impossibile generare il PDF della fattura" };
    }
    
    // 2. Prepara i parametri dell'email
    const oggetto = emailData.oggetto || `Fattura ${emailData.fatturaId}`;
    const messaggio = emailData.messaggio || "In allegato la fattura richiesta.";
    const ccList = emailData.cc ? emailData.cc.split(',').map(email => email.trim()) : [];
    
    // 3. Invia l'email con l'allegato
    GmailApp.sendEmail(
      emailData.destinatario,
      oggetto,
      messaggio,
      {
        attachments: [pdfBlob],
        cc: ccList.join(','),
        name: "Amministrazione Condominio" // Nome mittente personalizzabile
      }
    );
    
    // 4. Registra l'invio nel foglio
    registraInvioEmail(emailData.fatturaId, emailData.destinatario);
    
    return { success: true };
  } catch (error) {
    Logger.log(`Errore in inviaFatturaViaEmail: ${error}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Genera una risposta HTTP con il PDF della fattura
 * @param {string} fatturaId - ID della fattura
 * @returns {HtmlOutput} - Output HTML con PDF integrato
 */
function generaFatturaPdf(fatturaId) {
  try {
    // Genera il PDF
    const pdfBlob = generaPdfFattura({ fatturaId: fatturaId });
    if (!pdfBlob) {
      throw new Error("Impossibile generare il PDF della fattura");
    }
    
    // Converti in base64
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    // Crea pagina HTML che mostra il PDF
    const htmlOutput = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Fattura ${fatturaId}</title>
        <style>
          body, html { margin: 0; padding: 0; height: 100%; }
          embed { width: 100%; height: 100%; }
        </style>
      </head>
      <body>
        <embed src="data:application/pdf;base64,${pdfBase64}" type="application/pdf" />
      </body>
      </html>
    `);
    
    return htmlOutput
      .setTitle(`Fattura ${fatturaId}`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    Logger.log(`Errore in generaFatturaPdf: ${error}`);
    return HtmlService.createHtmlOutput(`<h1>Errore</h1><p>${error.message || error}</p>`)
      .setTitle('Errore')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}
/**
 * Genera l'HTML della fattura per la visualizzazione
 * @param {string} fatturaId - ID della fattura
 * @returns {HtmlOutput} - Output HTML della fattura
 */
function generaHtmlFattura(fatturaId) {
  try {
    // Questo è simile a generaFatturaPdf ma restituisce l'HTML invece del PDF
    
    // Ottieni i dati della fattura
    const ss = getSpreadsheet();
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    
    if (!sheetFatture) throw new Error("Foglio fatture non trovato");
    
    // Cerca la fattura per ID
    const valori = sheetFatture.getDataRange().getValues();
    let fatturaRow = null;
    for (let i = 1; i < valori.length; i++) {
      if (valori[i][0] === fatturaId) {
        fatturaRow = valori[i];
        break;
      }
    }
    
    if (!fatturaRow) throw new Error(`Fattura con ID ${fatturaId} non trovata`);
    
    // Ottieni i dati dell'appartamento
    const idAppartamento = fatturaRow[1];
    const datiAppartamento = getDatiAppartamentoById(ss, idAppartamento);
    
    // Ottieni i dati generali (impostazioni, tariffe, etc.)
    const impostazioni = leggiImpostazioniComplete();
    
    // Crea l'oggetto dati per la fattura
    const datiCompleti = {
      id: fatturaId,
      idAppartamento: idAppartamento,
      intestatario: datiAppartamento?.intestatario,
      indirizzo: datiAppartamento?.indirizzo,
      cfPiva: datiAppartamento?.cfPiva,
      dataInizio: fatturaRow[2],
      dataFine: fatturaRow[3],
      dataEmissione: fatturaRow[4],
      dataScadenza: fatturaRow[5],
      consumoElettrico: fatturaRow[6],
      importoElettrico: fatturaRow[7],
      consumoAcqua: fatturaRow[8],
      importoAcqua: fatturaRow[9],
      consumoGas: fatturaRow[10],
      importoGas: fatturaRow[11],
      importoTotale: fatturaRow[12],
      stato: fatturaRow[14],
      metodoPagamento: fatturaRow[15] || "Non specificato",
      scontoPercentuale: fatturaRow[16] || 0,  // Campo per lo sconto
      motivoSconto: fatturaRow[17] || "",  // Motivo dello sconto
      aziendaInfo: impostazioni.generali
    };
    
    // Usa la stessa funzione che genera l'HTML per il PDF
    const htmlContent = creaHtmlFattura(datiCompleti);
    
    // Restituisci l'HTML senza convertirlo in PDF
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle("Fattura " + fatturaId)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (error) {
    Logger.log(`Errore in generaHtmlFattura: ${error}`);
    return HtmlService.createHtmlOutput(`
      <h1>Errore</h1>
      <p>${error.message || "Si è verificato un errore durante la generazione della visualizzazione"}</p>
      <p><a href="javascript:window.history.back();">Torna indietro</a></p>
    `)
    .setTitle('Errore')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * Genera un PDF della fattura
 * @param {Object} dati - Dati della fattura
 * @returns {Blob} - PDF come Blob
 */
function generaPdfFattura(dati) {
  try {
    // 1. Ottieni dati completi della fattura dal foglio
    const ss = getSpreadsheet();
    const sheetFatture = ss.getSheetByName(NOME_FOGLIO_FATTURE);
    if (!sheetFatture) throw new Error(`Foglio ${NOME_FOGLIO_FATTURE} non trovato`);
    
    // Cerca la fattura per ID
    const valori = sheetFatture.getDataRange().getValues();
    let fatturaRow = null;
    for (let i = 1; i < valori.length; i++) {
      if (valori[i][COL_FATT_ID - 1] === dati.fatturaId) {
        fatturaRow = valori[i];
        break;
      }
    }
    
    if (!fatturaRow) throw new Error(`Fattura con ID ${dati.fatturaId} non trovata`);
    
    // 2. Ottieni i dati dell'appartamento
    const idAppartamento = fatturaRow[COL_FATT_APP_ID - 1];
    const datiAppartamento = getDatiAppartamentoById(ss, idAppartamento);
    if (!datiAppartamento) throw new Error(`Appartamento con ID ${idAppartamento} non trovato`);
    
    // 3. Ottieni i dati generali (impostazioni, tariffe, etc.)
    const impostazioni = leggiImpostazioniComplete();
    if (!impostazioni || !impostazioni.generali) throw new Error("Impossibile leggere le impostazioni");
    
    // 4. Crea il contenuto HTML per il PDF
    const htmlContent = creaHtmlFattura({
      id: dati.fatturaId,
      idAppartamento: idAppartamento,
      intestatario: datiAppartamento.intestatario,
      indirizzo: datiAppartamento.indirizzo,
      cfPiva: datiAppartamento.cfPiva,
      dataInizio: fatturaRow[COL_FATT_DATA_INIZIO - 1],
      dataFine: fatturaRow[COL_FATT_DATA_FINE - 1],
      dataEmissione: fatturaRow[COL_FATT_DATA_EMISSIONE - 1],
      dataScadenza: fatturaRow[COL_FATT_DATA_SCADENZA - 1],
      consumoElettrico: fatturaRow[COL_FATT_CONS_ELETTRICO - 1],
      importoElettrico: fatturaRow[COL_FATT_IMP_ELETTRICO - 1],
      consumoAcqua: fatturaRow[COL_FATT_CONS_ACQUA - 1],
      importoAcqua: fatturaRow[COL_FATT_IMP_ACQUA - 1],
      consumoGas: fatturaRow[COL_FATT_CONS_GAS - 1],
      importoGas: fatturaRow[COL_FATT_IMP_GAS - 1],
      importoTotale: fatturaRow[COL_FATT_IMP_TOTALE - 1],
      stato: fatturaRow[COL_FATT_STATO - 1],
      metodoPagamento: fatturaRow[COL_FATT_METODO_PAGAMENTO - 1] || "Non specificato",
      aziendaInfo: impostazioni.generali
    });
    
    // 5. Converti in PDF
    const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
      .getAs('application/pdf')
      .setName(`Fattura_${dati.fatturaId}.pdf`);
      
    return pdfBlob;
  } catch (error) {
    Logger.log(`Errore in generaPdfFattura: ${error}`);
    return null;
  }
}

/**
 * Crea l'HTML per la fattura
 * @param {Object} dati - Dati della fattura
 * @returns {string} - HTML completo
 */
function creaHtmlFattura(dati) {
  // Formatta le date in formato italiano
  const formatDate = (dateStr) => {
    if (!dateStr) return "N/D";
    if (typeof dateStr === 'string') {
      const dateParts = dateStr.split('-');
      if (dateParts.length === 3) {
        return `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
      }
      return dateStr;
    }
    // Se è un oggetto Date
    return Utilities.formatDate(new Date(dateStr), "GMT+1", "dd/MM/yyyy");
  };
  
  // Formatta numeri in formato italiano
  const formatNumber = (num, decimals = 0) => {
    if (num === undefined || num === null) return "0";
    return num.toLocaleString('it-IT', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
  };
  
  // Crea variabili formattate per l'output
  const formattedData = {
    id: dati.id || "N/D",
    intestatario: dati.intestatario || "N/D",
    indirizzo: dati.indirizzo || "N/D",
    cfPiva: dati.cfPiva || "N/D",
    dataInizio: formatDate(dati.dataInizio),
    dataFine: formatDate(dati.dataFine),
    dataEmissione: formatDate(dati.dataEmissione),
    dataScadenza: formatDate(dati.dataScadenza),
    consumoElettrico: formatNumber(dati.consumoElettrico) + " kWh",
    importoElettrico: formatNumber(dati.importoElettrico, 2) + " €",
    consumoAcqua: formatNumber(dati.consumoAcqua) + " m³",
    importoAcqua: formatNumber(dati.importoAcqua, 2) + " €",
    consumoGas: formatNumber(dati.consumoGas) + " m³",
    importoGas: formatNumber(dati.importoGas, 2) + " €",
    importoTotale: formatNumber(dati.importoTotale, 2) + " €",
    stato: dati.stato || "Non Definito",
    metodoPagamento: dati.metodoPagamento || "Non specificato",
    aziendaInfo: dati.aziendaInfo || {},
  };
  
  // Stile CSS per il PDF
  const styles = `
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 20px; font-size: 12px; }
      .container { max-width: 800px; margin: 0 auto; }
      table { width: 100%; border-collapse: collapse; margin: 20px 0; }
      th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
      th { background-color: #f2f2f2; }
      .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
      .header h1 { color: #3498db; margin-bottom: 5px; }
      .info-block { margin-bottom: 20px; }
      .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
      .box { border: 1px solid #eee; padding: 15px; border-radius: 5px; }
      .box h3 { margin-top: 0; border-bottom: 1px solid #eee; padding-bottom: 5px; }
      .footer { margin-top: 50px; padding-top: 10px; border-top: 1px dashed #ccc; font-size: 10px; color: #666; }
      .payment-method { background-color: #f8f9fa; padding: 10px; border-radius: 5px; margin-bottom: 20px; border-left: 3px solid #3498db; }
      .total-row { font-weight: bold; background-color: #f0f0f0; }
    </style>
  `;
  
  // HTML del documento
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Fattura ${formattedData.id}</title>
      ${styles}
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>FATTURA ${formattedData.id}</h1>
          <p>${formattedData.aziendaInfo.nomeAzienda || ""}<br>
          ${formattedData.aziendaInfo.indirizzoAzienda || ""}<br>
          P.IVA/CF: ${formattedData.aziendaInfo.pivaCfAzienda || ""}</p>
        </div>
        
        <div class="info-grid">
          <div class="box">
            <h3>Intestatario</h3>
            <p><strong>${formattedData.intestatario}</strong><br>
            ${formattedData.indirizzo}<br>
            CF/P.IVA: ${formattedData.cfPiva}</p>
          </div>
          <div class="box">
            <h3>Dettagli Fattura</h3>
            <p>
            <strong>Stato:</strong> ${formattedData.stato}<br>
            <strong>Data Emissione:</strong> ${formattedData.dataEmissione}<br>
            <strong>Data Scadenza:</strong> ${formattedData.dataScadenza}<br>
            <strong>Periodo:</strong> ${formattedData.dataInizio} - ${formattedData.dataFine}</p>
          </div>
        </div>
        
        <div class="payment-method">
          <h3>Metodo di Pagamento</h3>
          <p><strong>${formattedData.metodoPagamento}</strong></p>
        </div>
        
        <h3>Consumi e Importi</h3>
        <table>
          <thead>
            <tr>
              <th>Servizio</th>
              <th>Consumo</th>
              <th>Importo</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Energia Elettrica</td>
              <td>${formattedData.consumoElettrico}</td>
              <td>${formattedData.importoElettrico}</td>
            </tr>
            <tr>
              <td>Acqua</td>
              <td>${formattedData.consumoAcqua}</td>
              <td>${formattedData.importoAcqua}</td>
            </tr>
            <tr>
              <td>Gas</td>
              <td>${formattedData.consumoGas}</td>
              <td>${formattedData.importoGas}</td>
            </tr>
            <tr class="total-row">
              <td colspan="2"><strong>TOTALE</strong></td>
              <td><strong>${formattedData.importoTotale}</strong></td>
            </tr>
          </tbody>
        </table>
        
        <div class="info-block">
          <h3>Modalità di Pagamento</h3>
          <p><strong>IBAN:</strong> ${formattedData.aziendaInfo.iban || "N/D"}<br>
          <strong>Intestato a:</strong> ${formattedData.aziendaInfo.nomeAzienda || "N/D"}</p>
        </div>
        
        <div class="footer">
          <p><strong>Note:</strong> ${formattedData.aziendaInfo.notePredefinite || ""}</p>
          <p>Email: ${formattedData.aziendaInfo.emailContatto || ""} - Tel: ${formattedData.aziendaInfo.telefonoContatto || ""}</p>
          <p style="text-align: center; margin-top: 10px;">
            Documento generato automaticamente dal sistema di gestione condomini.
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Registra l'invio di una email nel foglio di calcolo
 * @param {string} fatturaId - ID della fattura
 * @param {string} destinatario - Email del destinatario
 */
function registraInvioEmail(fatturaId, destinatario) {
  try {
    // Potremmo creare un nuovo foglio "Email_Inviate" o aggiungere una colonna al foglio Fatture
    // Per ora, aggiungiamo solo un log
    Logger.log(`Registrato invio fattura ${fatturaId} a ${destinatario} il ${new Date()}`);
    
    // Se vuoi implementare il registro, puoi creare un foglio "Email_Inviate"
    // e aggiungere una riga con data, id fattura, destinatario, etc.
  } catch (error) {
    Logger.log(`Errore in registraInvioEmail: ${error}`);
  }
}

/**
 * Helper per ottenere l'URL dello script
 * Utilizzato nei template HTML
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ========================================================================
//                      FUNZIONI HELPER INTERNE
// ========================================================================
function trovaConsumoPrecalcolato(idAppartamento, tipoContatore, dataFine, tutteLetture) {
  if (!tutteLetture || tutteLetture.length === 0) return 0;
  
  // Filtra per tipo di contatore
  const lettureContatore = tutteLetture.filter(l => l.tipo === tipoContatore);
  if (lettureContatore.length === 0) return 0;
  
  // Trova la lettura più vicina alla data di fine, ma non oltre
  let letturaFinePeriodo = null;
  const dataFineObj = new Date(dataFine);
  
  for (let i = 0; i < lettureContatore.length; i++) {
    const lettura = lettureContatore[i];
    if (lettura.data <= dataFineObj) {
      if (!letturaFinePeriodo || lettura.data > letturaFinePeriodo.data) {
        letturaFinePeriodo = lettura;
      }
    }
  }
  
  // Se abbiamo trovato una lettura, usa il suo valore di consumo
  if (letturaFinePeriodo) {
    return letturaFinePeriodo.consumo;
  }
  
  return 0;
}

function calcolaImportoSingolaUtenza(consumo, tariffa) {
  if (!consumo || !tariffa) return 0;
  
  // Calcola importo base
  let importoBase = (parseFloat(consumo) * parseFloat(tariffa.prezzoUnitario)) + parseFloat(tariffa.costoFisso || 0);
  
  // Aggiungi IVA se presente
  if (tariffa.iva) {
    importoBase = importoBase * (1 + (parseFloat(tariffa.iva) / 100));
  }
  
  return Math.round(importoBase * 100) / 100; // Arrotonda a 2 decimali
}

function getDatiAppartamentoById(ss, idAppartamento) {
  const sheet = ss.getSheetByName(NOME_FOGLIO_APPARTAMENTI);
  if (!sheet) { Logger.log(`Foglio ${NOME_FOGLIO_APPARTAMENTI} non trovato`); return null; }
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
     if (values[i][COL_APP_ID - 1] == idAppartamento) {
       return { // Restituisce i dati necessari per l'anteprima
         id: values[i][COL_APP_ID - 1],
         intestatario: values[i][COL_APP_PROPRIETARIO - 1], // Usa Proprietario come Intestatario
         indirizzo: `Piano ${values[i][COL_APP_PIANO - 1] || 'N/D'}`, // Usa Piano come Indirizzo/Riferimento
         cfPiva: values[i][COL_APP_CF_PIVA - 1],
         email: values[i][COL_APP_EMAIL - 1] || "" // Include email per invio fatture
       };
     }
  }
  Logger.log(`Appartamento ID ${idAppartamento} non trovato in getDatiAppartamentoById.`);
  return null;
}

function generaProssimoIdFatturaNumerico(ss) {
  const sheet = ss.getSheetByName(NOME_FOGLIO_FATTURE);
  if (!sheet) { Logger.log(`Foglio ${NOME_FOGLIO_FATTURE} non trovato`); return "00001"; }
  
  // Conta il numero di righe con dati (esclusa intestazione)
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "00001"; // Solo intestazione
  
  // Conta le righe effettive con dati
  const valuesIDs = sheet.getRange(2, COL_FATT_ID, lastRow - 1, 1).getValues();
  let maxNumber = 0;
  
  // Tenta di estrarre numeri dalle stringhe ID esistenti
  for (let i = 0; i < valuesIDs.length; i++) {
    const idFattura = String(valuesIDs[i][0]);
    // Estrai numeri (cerca sequenze di cifre)
    const matches = idFattura.match(/\d+/g);
    if (matches) {
      // Prende l'ultimo gruppo di numeri trovato
      const numPart = parseInt(matches[matches.length - 1], 10);
      if (!isNaN(numPart) && numPart > maxNumber) {
        maxNumber = numPart;
      }
    }
  }
  
  // Incrementa di 1 e formatta con padding di zeri
  maxNumber++;
  return maxNumber.toString().padStart(5, "0");
}
