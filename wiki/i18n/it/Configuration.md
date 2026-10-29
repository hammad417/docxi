<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# Configurare docxi: margini, impostazioni di pagina e unità di misura

docxi non ha file di configurazione, variabili d'ambiente né impostazioni globali. Tutto si imposta per documento o per elemento. Questa pagina descrive le impostazioni a livello di documento e le unità di misura usate dalle opzioni. Le opzioni di ciascun elemento sono elencate in [Funzionalità](Features).

## <a id="page-margins"></a>Margini della pagina

I margini si passano quando crei il documento:

```ruby
require "docxi"

# Tutti i valori sono in ventesimi di punto (1440 = 1 pollice).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Opzione | Predefinito | Significato |
|---|---|---|
| `top` | `720` (0,5 pollici) | Margine superiore |
| `right` | `1440` (1 pollice) | Margine destro |
| `bottom` | `720` (0,5 pollici) | Margine inferiore |
| `left` | `1440` (1 pollice) | Margine sinistro |
| `header` | `190` | Distanza dal bordo superiore della pagina all'intestazione |
| `footer` | `200` | Distanza dal bordo inferiore della pagina al piè di pagina |

L'hash che passi sostituisce i valori predefiniti, quindi imposta `top`, `right`, `bottom` e `left` insieme. Qualsiasi di questi quattro valori tu ometta viene scritto come valore vuoto. `header` e `footer` tornano a 190 e 200 quando mancano o sono vuoti.

## <a id="page-size-and-orientation"></a>Dimensioni e orientamento della pagina

Le pagine sono in formato US Letter con orientamento verticale: 12240 × 15840 ventesimi di punto, cioè 8,5 × 11 pollici. Nella 0.1.0 dimensioni e orientamento non si possono cambiare.

## <a id="first-page"></a>Prima pagina

docxi attiva sempre l'opzione di Word «Diversi per la prima pagina». Le intestazioni e i piè di pagina aggiunti con `add_header` e `add_footer` si usano dalla pagina 2 in poi. La pagina 1 usa `add_first_header` e `add_first_footer`, quindi imposta anche questi se la prima pagina ha bisogno di un'intestazione o di un piè di pagina. Consulta [Funzionalità](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>Cosa contiene ogni documento

- Il carattere predefinito è Arial 12 pt.
- Stili definiti: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` e `NoList`.
- Le voci di elenco puntato (`ul: true`) usano un trattino come punto elenco.
- Le proprietà del documento indicano «Docxi Ruby Gem» come nome dell'applicazione e registrano l'ora in cui il documento è stato generato.

## <a id="units"></a>Unità di misura

| Opzione | Unità |
|---|---|
| `size:` del testo | Punti (docxi li raddoppia per ottenere i mezzi punti di Word) |
| Margini, `left:` e `right:` del paragrafo, `iwidth:` della tabella, `height:` della riga | Ventesimi di punto (1440 per pollice) |
| Spessore dei `borders:` della cella | Ottavi di punto |
| `columns_width:` e `width:` della cella | Moltiplicati per 14.92 per ottenere ventesimi di punto, valori vicini ai pixel a 96 dpi |
| `style:` dell'immagine in linea | Lunghezze CSS come `"120pt"` |
| `pH:` e `pV:` dell'immagine posizionata | EMU (914400 per pollice) |
| Colori del testo e del riempimento | Esadecimale senza `#`, per esempio `"8B0000"` |
| `fill_color:` dell'immagine | Colore VML con `#`, per esempio `"#00FF00"` |

## <a id="related-pages"></a>Pagine correlate

- [Funzionalità](Features)
- [Riferimento API](API-Reference)
- [Buone pratiche](Best-Practices)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
