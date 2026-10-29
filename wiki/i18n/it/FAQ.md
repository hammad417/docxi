<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# Domande frequenti su docxi: documenti Word in Ruby e Rails

Risposte brevi alle domande più comuni. Ogni risposta rimanda a una spiegazione più approfondita.

**Devo avere installato Microsoft Word o LibreOffice?**
No. docxi scrive da solo l'XML e l'archivio zip.

**docxi richiede Rails?**
No. Funziona in qualsiasi programma Ruby. I test vengono eseguiti anche con Rails dalla 5.2 alla 8.1 caricato. Consulta [Integrazioni](Integrations).

**Come genero un documento Word in Ruby on Rails?**
Costruiscilo in un controller o in un job in background e invia `word.render.string` con `send_data`. Consulta [Integrazioni](Integrations).

**Quali versioni di Ruby e di Rails sono supportate?**
Ruby 2.5 e successive. Rails è facoltativo, e la CI testa Rails dalla 5.2 alla 8.1. Consulta [Installazione](Installation).

**docxi può modificare un .docx esistente o compilare un modello di Word?**
No, crea solo documenti nuovi. Per i modelli con campi unione, usa una gem basata su modelli, come sablon.

**Il sommario viene generato dai titoli?**
No. Sei tu ad aggiungere ogni voce e il relativo numero di pagina. Consulta [Funzionalità](Features#table-of-contents).

**Posso usare elenchi numerati o stili titolo?**
Non nella 0.1.0. I paragrafi possono essere voci di elenco puntato (`ul: true`), e puoi formattare i titoli con opzioni di testo come `size:` e `bold:`.

**Posso cambiare le dimensioni o l'orientamento della pagina?**
Non ancora. Le pagine sono in formato US Letter con orientamento verticale, e puoi modificare solo i margini. Consulta [Configurazione](Configuration).

**Quali formati di immagine funzionano?**
PNG. Ogni immagine viene memorizzata come parte PNG.

**Posso aggiungere un'immagine da un URL?**
Sì. `add_media` accetta URL `http`, `https` e `ftp`, e li scarica quando chiami `render`. Non passare URL provenienti dagli utenti senza controllarli. Consulta [Sicurezza](Security).

**Posso scrivere il documento direttamente in un file?**
`render` restituisce uno `StringIO`. Scrivi il suo `.string` con `File.binwrite("out.docx", word.render.string)`.

**docxi può generare documenti molto grandi in streaming?**
No. L'intero documento viene costruito in memoria quando chiami `render`. Consulta [Deployment](Deployment).

**Posso impostare paragrafi da destra a sinistra?**
Nella 0.1.0 non ci sono opzioni per la scrittura da destra a sinistra. Il testo viene scritto così come lo passi.

**Dove sono elencati i problemi noti?**
Nella [Roadmap](Roadmap) e nella sezione «Limitations and known issues» (limitazioni e problemi noti) del README.

## <a id="related-pages"></a>Pagine correlate

- [Risoluzione dei problemi](Troubleshooting)
- [Funzionalità](Features)
- [Primi passi](Getting-Started)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
