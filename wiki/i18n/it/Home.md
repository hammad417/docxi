<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# Wiki di docxi: generare documenti Word (.docx) con Ruby

docxi è una gem Ruby che genera file `.docx` di Microsoft Word a partire dal codice. Offre paragrafi, testo formattato, tabelle, immagini, intestazioni e piè di pagina, numeri di pagina, collegamenti ipertestuali e un sommario semplice. Non richiede Microsoft Word, LibreOffice né un modello, e funziona sia in Ruby puro sia nelle applicazioni Ruby on Rails.

Il [README](https://github.com/hammad417/docxi#readme) è l'introduzione breve. Questa wiki copre ogni opzione, offre ricette pronte da copiare, spiega come funziona docxi al suo interno e ti aiuta quando qualcosa non va.

Versione attuale: **0.1.1**, per Ruby 2.5 o successivo. Consulta il [registro delle modifiche](Changelog).

## <a id="start-here"></a>Da dove iniziare

**Se docxi è nuovo per te**

- [Primi passi](Getting-Started): installa la gem e crea il tuo primo documento
- [Installazione](Installation): requisiti, Bundler e aggiornamento dalla 0.0.x
- [Esempi d'uso](Usage-Examples): report, fatture, lettere e tabelle a partire dai dati, pronti da copiare

**Creare documenti**

- [Funzionalità](Features): ogni elemento e le opzioni che accetta
- [Configurazione](Configuration): margini, impostazioni di pagina e unità di misura
- [Riferimento API](API-Reference): classi e metodi a colpo d'occhio
- [Buone pratiche](Best-Practices): abitudini che evitano le sorprese più comuni

**Quando qualcosa non va**

- [Risoluzione dei problemi](Troubleshooting): messaggi di errore e documenti che non appaiono come dovrebbero
- [Domande frequenti](FAQ): risposte brevi alle domande più comuni
- [Sicurezza](Security): segnalare problemi e gestire input non attendibili

**Per approfondire**

- [Architettura](Architecture): come le chiamate Ruby diventano un pacchetto `.docx`
- [Integrazioni](Integrations): controller Rails, applicazioni Rack e job in background
- [Deployment](Deployment): usare docxi in produzione e come vengono pubblicate le versioni
- [Contribuire](Contributing) e [Ambiente di sviluppo](Development-Setup)
- [Roadmap](Roadmap), [Registro delle modifiche](Changelog) e [Riconoscimenti](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>Cosa fa docxi e cosa non fa

docxi crea documenti nuovi. Non può aprire né modificare un `.docx` esistente, compilare un modello di Word o eseguire una stampa unione. Non ha elenchi numerati, stili titolo né un sommario automatico, e le pagine sono sempre in formato US Letter con orientamento verticale. I problemi noti previsti per la 0.2.0 sono elencati nella [Roadmap](Roadmap).

## <a id="credits"></a>Riconoscimenti

La manutenzione di docxi è curata da [Hammad Maqbool](https://github.com/hammad417) presso [Phaedra Solutions](https://www.phaedrasolutions.com/). Il progetto è nato nel 2016 a partire da [docxer](https://github.com/gtimti/docxer) di Timur Gabdrakipov. Consulta i [Riconoscimenti](Credits) per i dettagli.

## <a id="languages"></a>Lingue

La versione inglese è quella di riferimento per questa wiki. I link alle traduzioni si trovano nella barra laterale e in cima a ogni pagina.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Segnala un problema](https://github.com/hammad417/docxi/issues)
