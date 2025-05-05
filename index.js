let excelLoaded = false;
let pdfLoaded = false;
let excelCategories;
let pdfFields;
let originalPdfBytes;

/* Event Listener sull'elemento con id "ExcelFile", la funzione si attiva quando
l'utente seleziona il file excel e lo carica */
document.getElementById('excelFile').addEventListener('change', convertExcelToCSV);

async function convertExcelToCSV() {
  // Recupero l'elemento HTML, ovvero il file Excel che l'utente inserisce
  const excelFileInput = document.getElementById('excelFile');
  const excelFile = excelFileInput.files[0];

  // Utilizzo un FileReader per leggere il contenuto del file
  const reader = new FileReader();

  // Gestore di eventi, quando il lettore di file carica con successo il file Excel, l'evento load
  // viene attivato. Riceve un parametro 'e' che rappresenta l'evento di caricamento.
  reader.onload = async (e) => {
    try {
      // Estraggo il contenuto del file Excel caricato nell'evento load e li memorizzo in un buffer di array
      const arrayBuffer = e.target.result;
      // I dati del file Excel vengono convertiti in un oggetto Uint8Array per poterli leggere.
      const data = new Uint8Array(arrayBuffer);

      // Leggo i dati del file Excel
      const workBook = XLSX.read(data, { type: 'array' });

      // In caso non fossero presenti fogli di lavoro la funzione si interrompe
      if (!workBook.SheetNames.length) {
        alert('Il file Excel non contiene fogli di lavoro.');
        return;
      }

      const firstSheetName = workBook.SheetNames[0];
      const workSheet = workBook.Sheets[firstSheetName];

      // Converto il foglio di lavoro in formato CSV tramite libreria 'xlsx'
      const csvData = XLSX.utils.sheet_to_csv(workSheet);

      // Analizzo la prima riga della stringa CSV che contiene i nomi delle categorie
      const columnNames = Papa.parse(csvData.split('\n')[0]).data;

      // Rimuovo la prima riga (intestazione) dal CSV prima di analizzare i dati restanti
      const csvWithoutHeader = csvData.split('\n').slice(1).join('\n');

      Papa.parse(csvWithoutHeader, {
        header: false, // Ignoro la prima riga che contiene i nomi delle categorie
        dynamicTyping: true,
        skipEmptyLines: true,
        complete: (result) => {
          // result.data conterrà solo i valori delle celle, senza i nomi delle colonne
          excelValues = result.data;
          excelCategories = columnNames[0]; // Si assume che i nomi delle colonne siano nella prima riga

          // Mostro le informazioni utili sul file Excel
          const excelInfoContainer = document.querySelector('.infoExcel p');
          // Costruisce il contenuto con i nomi delle categorie in grassetto e il numero di righe in grassetto
          const categoriesText = excelCategories.map(category => `${category}`).join(', ');
          const rowsText = `• <strong>Numero di righe</strong>: ${excelValues.length}`;

          // Aggiunge il contenuto all'elemento HTML
          excelInfoContainer.innerHTML = `• <strong>Categorie</strong>: ${categoriesText}<br>${rowsText}`;


          excelLoaded = true;
          checkAndCreateInterface();
        },
        error: (error) => {
          alert(`Errore durante il parsing CSV: ${error.message}`);
        },
      });
    } catch (error) {
      alert(`Errore durante la lettura di Excel: ${error.message}`);
    }
  };
  reader.readAsArrayBuffer(excelFile);
}

let pdfDoc;
document.getElementById('pdfFile').addEventListener('change', async (event) => {
  const file = event.target.files[0];
  originalPdfBytes = await file.arrayBuffer();
  pdfDoc = await PDFLib.PDFDocument.load(originalPdfBytes);
  console.log('PDF caricato con successo:', pdfDoc);
  usePdfDoc();

  // Mostro il contenitore delle informazioni quando carico il PDF
  document.querySelector('.infoPdf').style.display = 'block';

  // Estrazione dei nomi dei campi PDF e dei loro costruttori
  const fields = pdfDoc.getForm().getFields();
  const pdfInfoContainer = document.querySelector('.infoPdf p');
  pdfInfoContainer.textContent = ''; // Pulisco il contenuto precedente

  // Stampo in pagina le info sul PDF
  fields.forEach(field => {
    const info = document.createElement('div');
    info.innerHTML = `• Il campo <strong>${field.getName()}</strong> ha come costruttore il tipo "<strong>${field.constructor.name}</strong>"`;

    // Situazione particolare in caso di una Radio
    if (field.constructor.name === 'PDFRadioGroup') {
      const options = field.getOptions();
      const numOptions = options.length;
      const optionValues = options.join(', ');
      // Aggiungo al messaggio di info iniziale
      info.innerHTML += ` e ha <strong>${numOptions}</strong> opzioni: "<strong>${optionValues}</strong>"`;
    }

    pdfInfoContainer.appendChild(info);
  });

  // Estrazione dei nomi dei campi PDF
  pdfFields = fields.map(field => field.getName());

  pdfLoaded = true;
  checkAndCreateInterface();
});

// Controlla se pdfDoc è definito
function usePdfDoc() {
  if (pdfDoc) {
    console.log('Stai utilizzando pdfDoc globalmente');
  } else {
    console.error('pdfDoc non è definito.');
  }
}


// Se entrambe le flag sono true e excelCategories e pdfFields hanno un contenuto, chiamo la funzione per l'associazione
function checkAndCreateInterface() {
  // <--- Controlli temporanei ---> //
  console.log('La funzione checkAndCreateInterface è stata chiamata con successo.');
  console.log('excelCategories:', excelCategories);
  console.log('pdfFields:', pdfFields);
  if (excelLoaded && pdfLoaded && excelCategories && pdfFields) {
    createAssociationInterface(excelCategories, pdfFields);
  }
}


// <--- GENERAZIONE INTERFACCIA ---> //
function createAssociationInterface() {
  // Recupero l'elemento HTML dove verranno inseriti gli elementi UI
  const associationContainer = document.getElementById('associationContainer');
  const compileButton = document.getElementById('compileButton');
  associationContainer.innerHTML = '';
  compileButton.innerHTML = '';

  // Scorro attraverso ogni campo del PDF presente nell'array pdfFields, per ognuno:
  pdfFields.forEach(pdfField => {
    // Creo il label
    const pdfFieldLabel = document.createElement('label');
    pdfFieldLabel.textContent = pdfField;

    // Creo una select per la selezione delle categorie Excel
    const selectExcelCategory = document.createElement('select');

    // Creo un'opzione vuota 
    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = '';
    selectExcelCategory.appendChild(emptyOption);

    /* Scorro attraverso ogni categoria Excel presente nell'array excelCategories e per ognuna
    creo un opzione nella select con testo e valore corrispondenti */
    excelCategories.forEach(excelCategory => {
      const option = document.createElement('option'); // per ogni cat. Excel creo una <option>
      //IMPORTANTE: Imposto il valore dell'opzione con il valore della cat. Excel, quando un'opzione viene selezionata, il valore viene utilizzato
      option.value = excelCategory;
      // Imposto il testo dell'opzione con il testo della cat. Excel, questo è quello che l'utente vede
      option.textContent = excelCategory;

      // Verifico se il nome della categoria Excel è presente o parzialmente presente nel nome del campo PDF
      if (pdfField.toLowerCase().includes(excelCategory.toLowerCase())) {
        // Se è presente, seleziono automaticamente questa opzione
        option.selected = true;
      }

      // Aggiungo l'elemento alla select con appendChild()
      selectExcelCategory.appendChild(option);
    });

    /* Aggiungo gli elementi al contenitore: il label del campo PDF, la select delle 
    categorie Excel e un'elemento <br> per separare i campi */
    associationContainer.appendChild(pdfFieldLabel);
    associationContainer.appendChild(selectExcelCategory);
  });

  // Creo l'elemento button per il pulsante 'Compila PDF'
  const compilePdfBtn = document.createElement('button');
  compilePdfBtn.textContent = 'Compila PDF';
  // Al click sul pulsante eseguo la funzione 'compilePdf'
  compilePdfBtn.addEventListener('click', compilePdf);
  // Aggiungo il bottone al contenitore
  compileButton.appendChild(compilePdfBtn);
  document.querySelector('#compileButton').style.display = 'block'
}

// <--- CREAZIONE ASSOCIAZIONI --->
// Funzione per compilare il PDF utilizzando le associazioni tra categorie Excel e campi PDF
async function compilePdf() {

  // Recupero il contenitore e prendo le select al suo interno 
  const associationContainer = document.getElementById('associationContainer');
  const selects = associationContainer.querySelectorAll('select');

  // Inizializzo la variabile associations
  let associations = [];

  // Recupero tutte le righe di valori dal file Excel
  const excelRows = excelValues;

  /* ciclo for...of scorre attraverso ogni riga del file Excel, 'excelRows' contiene i valori della riga corrente 
  Funzionamento: dopo aver definito il ciclo, ad ogni iterazione la variabile excelRow assume il valore dell'elemento 
  corrente nell'array excelRows */
  for (const excelRow of excelRows) {
    // Ciclo foreach attraverso ogni elemento select nell'interfaccia, per ognuno:
    selects.forEach(select => {
      // Recuperiamo il nome del campo PDF associato alla select, prendendo il testo del nodo precedente
      const pdfField = select.previousSibling.textContent;
      // Recuperiamo il valore selezionato dall'utente
      const excelCategory = select.value;
      // Recuperiamo il valore corrispondente alla categoria Excel per la riga corrente
      const excelValue = excelRow[excelCategories.indexOf(excelCategory)];

      /* le info recuperate vengono utilizzate per creare l'oggetto associations, che rappresenta 
      l'associazione tra campo PDF, categoria Excel e il rispettivo valore */
      associations.push({
        pdfField: pdfField,
        excelCategory: excelCategory,
        excelValue: excelValue
      });
    });

    // Chiamo la funzione che compila il PDF per questa riga di valori passandogli associations
    await compilePdfField(associations);
    // svuoto l'array associations per la prossima riga di valori
    associations = [];
  }
}

let pdfCounter = 1;
let errors = [];
async function compilePdfField(associations) {

  //IMPORTANTE: Ricarico il PDF originale prima di ogni compilazione
  pdfDoc = await PDFLib.PDFDocument.load(originalPdfBytes);

  /* Ciclo attraverso ogni associazione per questa riga di valori, ogni associazione
  rappresenta un campo PDF da compilare
  Funzionamento: dopo aver definito il ciclo, ad ogni iterazione la variabile association assume il valore dell'elemento 
  corrente nell'array associations */
  for (const association of associations) {
    // Ottengo excelCategory, pdfField e excelValue da association
    const { excelCategory, pdfField, excelValue } = association;
    // Trim() Funzionamento: metodo delle stringhe che rimuove eventuali spazi bianchi dall'inizio e dalla fine di una stringa
    const excelCategoryValue = excelCategory.trim();

    // Verifico se il valore della categoria Excel non è vuoto
    if (excelCategoryValue !== "") {
      console.log(`Tentativo di compilare il campo PDF "${pdfField}" con valore "${excelValue}"`);

      // Ottengo il nome del campo da compilare che corrisponde a quello dell'associazione tramite il metodo find()
      const pdfFieldObject = pdfDoc.getForm().getFields().find(field => field.getName() === pdfField);

      // Verifico se il campo è stato trovato
      if (pdfFieldObject) {
        // Utilizzo uno switch per gestire diversi tipi di campi
        switch (pdfFieldObject.constructor.name) {

          case 'PDFTextField':
            if (excelValue !== null) {
              // Converto in stringa in caso dovessi avere dei numeri
              pdfFieldObject.setText(excelValue.toString());
              console.log(`Campo PDF "${pdfField}" compilato con valore: "${excelValue}"`);
            } else {
              console.error(`Il valore per il campo PDF "${pdfField}" è null.`);
            }
            break;

          case 'PDFCheckBox':
            let boolValue;
            // Verifica se il valore del campo Excel è nullo o vuoto
            if (excelValue === null || excelValue.trim() === '') {
              boolValue = false;
              pdfFieldObject.uncheck();
            } else if (excelValue.trim().toLowerCase() === 'x' || // toLowerCase(): converte tutti i caratteri in lettere minuscole
              excelValue.trim().toLowerCase() === 'true' ||
              excelValue.trim().toLowerCase() === 'yes' ||
              excelValue.trim().toLowerCase() === 'si' ||
              excelValue.trim() === '1') {
              boolValue = true;
              pdfFieldObject.check();
            } else {
              pdfFieldObject.uncheck();
              // Memorizzo l'errore nell'array errors dove verrà poi stampato
              errors.push(`<span class="boldError">Attenzione:</span> Nel PDF compilato numero_${pdfCounter} nel campo ${pdfField}, il valore <strong>"${excelValue}" non è valido.</strong>`);
            }
            console.log(`Campo checkbox PDF "${pdfField}" compilato con valore: "${boolValue}"`);
            break;

          case 'PDFDropdown':
            // Verifica se il valore del campo Excel è nullo o vuoto o indefinito
            if (excelValue === null || excelValue === undefined || excelValue.toString().trim() === '') {
              // Codice precedente per gestire il caso in cui il valore Excel è vuoto
            } else {
              /* Creo un indice che rappresenta la posizione dell'opzione nel menù a discesa nel PDF corrispondente al valore Excel,
              se questa corrispondenza è presente otterrò sicuramente un indice >=0 */
              const optionIndex = pdfFieldObject.getOptions().findIndex(option => option === excelValue.toString());
              // Verifica per determinare se l'opzione fornita dal file Excel esiste già o meno
              if (optionIndex !== -1) {
                // Se corrisponde a una delle opzioni, seleziona quella opzione
                pdfFieldObject.select(pdfFieldObject.getOptions()[optionIndex]);
                console.log(`Campo combobox PDF "${pdfField}" selezionato con opzione "${excelValue}"`);
              } else {
                // Se non corrisponde a nessuna opzione esistente, aggiungi una nuova opzione
                pdfFieldObject.addOptions(excelValue.toString());
                pdfFieldObject.select(excelValue.toString()); // Seleziono la nuova opzione appena aggiunta

                // Stampo un avviso
                errors.push(`<span class="boldError">Nota:</span> Nel PDF ${pdfCounter} nel campo ${pdfField}, il valore "${excelValue}" non è presente, è stato quindi aggiunto alla lista.`)
              }
            }
            break;

          case 'PDFRadioGroup':
            // Mi assicuro che nessuna opzione sia selezionata prima di eseguire il codice
            pdfFieldObject.clear();
            // Verifico se il valore è nullo o indefinito o vuoto
            if (excelValue === null || excelValue === undefined || excelValue.toString().trim() === '') {
              pdfFieldObject.clear();
              console.log(`Campo radio PDF "${pdfField}" lasciato senza selezione`);
            } else {
              // Ragionamento uguale alla Dropdown ma senza l'aggiunta dell'opzione
              const optionIndex = pdfFieldObject.getOptions().findIndex(option => option === excelValue.toString());
              if (optionIndex !== -1) {
                // Se corrisponde a una delle opzioni, seleziona quella opzione
                pdfFieldObject.select(pdfFieldObject.getOptions()[optionIndex]);
                console.log(`Campo radio PDF "${pdfField}" selezionato con opzione "${excelValue}"`);
              } else {
                // Se non c'è corrispondenza stampo un avviso
                errors.push(`<span class="boldError">Attenzione:</span> Nel PDF compilato numero_${pdfCounter} nel campo ${pdfField}, il valore <strong>"${excelValue}" non è valido.</strong>`);
              }
            }
            break;

          default:
            console.warn(`Il tipo di campo PDF "${pdfField}" non è supportato per la compilazione.`);
        }
      } else {
        console.warn(`Campo PDF non trovato: "${pdfField}"`);
      }
    } else {
      console.log(`La categoria Excel per il campo "${pdfField}" è vuota, quindi non viene compilato.`);
    }
  }

  // Crea il nuovo PDF compilato per la riga di valori corrente
  const pdfBytes = await pdfDoc.save();
  const blob = new Blob([pdfBytes], { type: 'application/pdf' });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `PDF compilato numero_${pdfCounter}.pdf`; // Utilizzo un nome univoco per il PDF
  link.click();
  window.URL.revokeObjectURL(url);

  // Incrementa il contatore per i nomi dei PDF
  pdfCounter++;

  // Metodo per stampare gli errori
  const errorContainer = document.getElementById('errorContainer');
  errors.forEach(errorMessage => {
    const errorDiv = document.createElement('div');
    errorDiv.innerHTML = errorMessage;
    errorContainer.appendChild(errorDiv);

    // faccio comparire il container solo se è presente almeno un errore
    if (errors.length > 0) {
      errorContainer.style.display = 'block';
    }
    // Svuoto l'array per il prossimo PDF
    errors = [];
  });
}