README per il compilatore PDF da Excel

<---Istruzioni per l'installazione delle dipendenze utilizzate nel programma--->

Premessa: il programma non presenta l'obbligo di installare alcuna dipendenza sul proprio PC dato che tutto il necessario è caricato direttamente nel file HTML tramite CDN (Content Delivery Network).

Tuttavia, nel caso (futuro) in cui si presentasse l'obbligo di installare le dipendenze per poter utilizzare il programma correttamente, il seguente manuale vi guiderà passo passo.

1. Assicurarsi di avere Node.js e npm installati sul proprio pc. Puoi farlo da qui https://nodejs.org/en.

2. Caricare la cartella in un editor di codice come Visual Studio Code.

3. Una volta nella directory del progetto, tramite il terminale eseguire il comando "npm install" per installare le dipendenze necessarie.

4. In alternativa eseguire il comando "npm install papaparse pdf-lib"

Le dipendenze aggiornate che il programma utilizza sono:

	xlsx v.0.18.5 Permette la manipolazione di file Excel
	sito web: https://sheetjs.com

	Papaparse v.5.4.1 Permette il parsing di file CSV
	sito web: https://www.papaparse.com

	PDF-Lib v.1.17.1 Permette la manipolazione di file PDF
	sito web: https://pdf-lib.js.org

NB: xlsx può essere caricata solo ed esclusivamente tramite CDN, per maggiori informazioni visitare:
https://github.com/SheetJS/sheetjs/issues/2822
oppure
https://git.sheetjs.com/sheetjs/sheetjs/issues/2667

5. eseguire il comando "npm install -g live-server"

6. eseguire il comando "live-server" per ottenere un server locale e aprire il sito web automaticamente sul browser predefinito.

