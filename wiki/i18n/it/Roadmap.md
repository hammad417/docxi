<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# Roadmap e problemi noti di docxi

docxi non ha un calendario delle versioni né una roadmap con impegni precisi. L'unico lavoro pianificato e documentato è la correzione dei problemi noti qui sotto. Le correzioni sono previste per la 0.2.0, senza una data. Sono state tenute fuori dalla 0.1.0 perché cambiano i documenti che gli utenti attuali già generano.

## <a id="known-issues-planned-for-020"></a>Problemi noti previsti per la 0.2.0

| Problema | Soluzione alternativa nella 0.1.0 |
|---|---|
| Le immagini posizionate, quelle nelle intestazioni e quelle nelle celle delle tabelle calcolano le dimensioni come `width * height * 14.92` EMU per entrambe le dimensioni | Usa immagini in linea dimensionate con `style:`, oppure consulta [Buone pratiche](Best-Practices#images) |
| I `borders:` a livello di tabella vengono applicati a ogni cella invece che al contorno della tabella | Per il solo contorno, imposta `borders:` sui lati esterni delle celle del perimetro |
| Nel codice del campo numero di pagina manca la barra rovesciata di `\* MERGEFORMAT` | Nessuna |
| I numeri di pagina nelle celle delle tabelle sono preceduti da «GlobalOptions» | Metti i numeri di pagina in un'intestazione o in un piè di pagina |
| Le immagini vengono sempre memorizzate come parti PNG, e `[Content_Types].xml` dichiara solo le prime due intestazioni e i primi due piè di pagina | Usa immagini PNG e al massimo due intestazioni e due piè di pagina |
| Passare solo alcuni margini fa scrivere vuoti gli altri | Passa `top`, `right`, `bottom` e `left` insieme |
| `docProps/core.xml` registra l'ora locale con un indicatore UTC e non scrive mai l'autore | Nessuna |
| `Docxi.test` solleva `NoMethodError` | Non chiamarlo. È un metodo di supporto rimasto da versioni precedenti |
| `document.br` e `document.tab` senza `page: true` producono XML non valido | Usa `p.br` e `p.tab` dentro un paragrafo |

## <a id="not-planned"></a>Non pianificato

Questi punti non fanno parte dell'elenco qui sopra, e non è stato deciso nulla in merito:

- modificare documenti esistenti o compilare modelli
- elenchi numerati, stili titolo o un sommario automatico
- altri formati di pagina o pagine orizzontali

## <a id="related-pages"></a>Pagine correlate

- [Registro delle modifiche](Changelog)
- [Contribuire](Contributing)
- [Domande frequenti](FAQ)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
