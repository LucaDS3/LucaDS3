16/11/2023

- con pulsante F9 do sempre il comando "fai un calcolo".
    

![](https://www.evernote.com/shard/s566/res/2ac2f99c-f6f3-7ce8-7555-5928c8e6a6cb)

  

- **Valuta formula** permette di fare i calcoli delle formule vedendo i risultati parziali e seguire passo passo tutti i passaggi
    

  

![](https://www.evernote.com/shard/s566/res/2d780805-caca-c2ec-1b82-f384908411fe)![](https://www.evernote.com/shard/s566/res/3e7fac91-7504-dde2-767f-98d9b40d4e27)![](https://www.evernote.com/shard/s566/res/2b95dc50-1335-cc5d-0fa7-ed0d65249f3b)

- I nomi alternativi alle celle devono essere senza spazi. possono essere utilizzati all'interno delle funzioni per renderle più leggibili
    

  

CONVALIDA DATI

![](https://www.evernote.com/shard/s566/res/b61bf972-dde3-2865-f46d-f3c32e4415b0)![](https://www.evernote.com/shard/s566/res/57094294-4153-58d2-25d4-de705edb47b5)

permette di definire che valori possono essere inseriti all'interno di una cella. serve a limitare i dati che possono essere inseriti

  

SCARTO

permette la ricerca all'interno di un campo in riferimento al numero di righe e di colonne e da un risultato

![](https://www.evernote.com/shard/s566/res/d5017f0a-2039-a52b-2a3e-00699720b329)

  

la funzione scarto può essere anche utilizzata per definire un elenco in modo dinamico

![](https://www.evernote.com/shard/s566/res/3979ac4e-8f88-4563-ab30-eab2d5423840)![](https://www.evernote.com/shard/s566/res/ba9a2ca5-d368-e1c7-e210-55a3f8dc70af)

  

la prima data utilizzata dal sistema è 01/01/1900. tutte le date precedenti non vengono considerate

le date sono i numeri dei giorni passati dal 01/01/1900 (1,00).

  

23/11/2023

posso richiamare un foglio da un altro scrivendolo tra virgolette e con punto esclamativo alla fine:

=SOMMA('Esempio IVA'!B2:b5)

gli apici vanno inseriti solo in caso di spazio vuoto nel nome del file. se non ci sono spazi, gli apici non servono.

  

ERRORI

#VALORE --> appare se faccio formule con numeri e lettere (ad esempio =8+ciao)

#NOME --> appare se sbaglio a scrivere il nome di una cella (ad esempio =A2+B)

#DIV/0 ---> non è un errore, ma va elaborato

se utilizzo la cella dove inserisco il risultato come elemento, creo un riferimento cicolare. lo posso trovare nella voce funzioni, controlla errori, riferimento circolare

#RIF! --> significa che è stata eliminata una cella o delle righe di riferimento per una funzione

  

per sapere quali celle sono state convalidate, basta cercare il valore nel menu home in "cerca"

![](https://www.evernote.com/shard/s566/res/7aafa5b7-a71f-91df-5e4c-9501b00ba22c)

  

Analisi di simulazione:

![](https://www.evernote.com/shard/s566/res/b515745e-87d0-adbe-28f7-7ce14e82f12c)![](https://www.evernote.com/shard/s566/res/395a218e-5033-d9ef-62e5-bb2bc33c5f62)

excel modifica la cella richiesta per ottenere

  

per copiare un foglio, basta che lo trascino e premo ctrl

  

Analisi di simulazione: tabella dati.

scrivere i valori di simulazione sempre nella colonna prima della cella della funzione da analizzare

![](https://www.evernote.com/shard/s566/res/36a2d491-7ce6-5bdf-cf0c-a1ba81d20c9c)![](https://www.evernote.com/shard/s566/res/544d20d3-485a-a3a7-2612-689727418b92)

può essere fatto ad una o a due variabili.

se ci sono 3 variabili, si passa a Scenari

![](https://www.evernote.com/shard/s566/res/6ba60f0c-429c-56f2-440f-6f9c5f0fe129)![](https://www.evernote.com/shard/s566/res/44f9aa6c-67d9-ca5f-227c-4de534c9a775)

se clicco mostra, mi modifica i valori in modo da simulare i risultati per i dati che cambiano

poi può costruire una tabella di riepilogo

![](https://www.evernote.com/shard/s566/res/f34d275f-2743-ba8c-4417-861ea1ca7d57)![](https://www.evernote.com/shard/s566/res/9df4273f-9710-ee86-51a6-ae7b49efd519)

  

![](https://www.evernote.com/shard/s566/res/21b8b199-386c-64d2-dc01-d2de00e450f8)

permette di avere il riferimento a quale riga appartiene una cella. può essere poi usato a sua volta in altre funzioni

  

funzione CONTA.VALORI serve a contare quanti elementi sono presenti in un riferimento

  

![](https://www.evernote.com/shard/s566/res/3f0e3a68-1728-89a5-2a44-b73f850265ef)

Vari elementi di formattazione celle:

#.##0,0 - Numero con un decimale dopo la virgola ed eventuale separatore delle migliaia (es. se scrivo 3 leggerò 3,0; se scrivo 1000 leggerò 1.000,0)

00000  - Numero in cui vedo sempre fino alle decine di migliaia (es. se scrivo 4 leggerò 00004)

(0,00) "ton" - Numero tra parentesi di cui vedo fino i centesimi dopo la virgola e la scritta "ton" (se scrivo 2, leggerò (2,00 ton))

0. "K"- Numero in migliaia seguito da K (es. scrivo 2000, leggerò 2 K)

aaaa (mmm) gg (ggg) - Data con nome del mese e del giorno tra parentesi (es. se scrivo 30/11/2023 leggerò 2023 (nov) 30 (gio))

[Rosso] [hh]:mm - ora e minuti in rosso (es. scrivo 15:25, leggerò 15:25 in rosso) se metto hh tra parentesi quadre, elimino il reset dell'ora dopo 24 ore

L. #.##0; [Rosso] L. #.##0;;[Blu] @ - Positivo in lire, negativo in lire rosso, zero non visibile, testo in blu - @ fa da place holder per i testi

**_positivi; Negativi; zero; testi_**

;;; - non far vedere niente

  

formattazione condizionale

![[Pasted image 20231214085300.png]]  

funzione [RANGO.UG](http://RANGO.UG) - permette di identificare una classifica in base ai numeri, dando un riferimento al numero più alto

![[Pasted image 20231214085239.png]]


