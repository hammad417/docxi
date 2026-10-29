<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# Contribuire a docxi

Segnalazioni di bug, correzioni, documentazione e traduzioni sono tutte benvenute. Questa pagina spiega come segnalare un problema e di cosa ha bisogno una pull request. Per preparare una copia di lavoro, consulta [Ambiente di sviluppo](Development-Setup).

## <a id="reporting-a-bug"></a>Segnalare un bug

Apri una issue su [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) indicando:

- un breve script Ruby che riproduca il problema
- le versioni di Ruby, docxi, nokogiri e rubyzip (e di Rails, se lo usi)
- cosa ti aspettavi e cosa hai ottenuto. Per i documenti difettosi, indica con quale programma li hai aperti

Segnala invece i problemi di sicurezza in privato. Consulta [Sicurezza](Security).

## <a id="pull-requests"></a>Pull request

1. Fai un fork del repository e crea un branch.
2. Apporta la modifica rispettando lo stile del codice circostante.
3. Aggiungi o aggiorna i test ed esegui `bundle exec rake test`.
4. Se l'XML generato cambia di proposito, esegui `bundle exec rake golden` e spiega la modifica nella pull request. I file di riferimento proteggono l'output su cui contano gli utenti attuali.
5. Aggiorna `CHANGELOG.md` e la documentazione se il comportamento cambia.
6. Apri la pull request. La CI esegue i test su Ruby dalla 2.5 alla 4.0 e su Rails dalla 5.2 alla 8.1, e verifica che LibreOffice apra i documenti di esempio.

Le modifiche che alterano documenti esistenti, come le correzioni previste per la 0.2.0, vanno in una pull request a parte, così la differenza nei file di riferimento è facile da esaminare.

## <a id="documentation-and-translations"></a>Documentazione e traduzioni

- Il README in inglese è la fonte di 15 traduzioni in `docs/i18n/`. Quando lo modifichi, segui `docs/i18n/TRANSLATION_GUIDE.md` nel repository.
- Questa wiki viene scritta nella cartella `wiki/` del repository e pubblicata da lì. `wiki/README.md` spiega come.

## <a id="related-pages"></a>Pagine correlate

- [Ambiente di sviluppo](Development-Setup)
- [Architettura](Architecture)
- [Roadmap](Roadmap)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
