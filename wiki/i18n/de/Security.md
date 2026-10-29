<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Sicherheit: Probleme melden und nicht vertrauenswürdige Eingaben behandeln

Diese Seite erklärt, wie Sie ein Sicherheitsproblem melden, und was Sie zur Sicherheit wissen sollten, wenn Ihre Anwendung Benutzereingaben an docxi weitergibt.

## <a id="reporting-a-vulnerability"></a>Eine Sicherheitslücke melden

Bitte melden Sie Sicherheitsprobleme nicht in öffentlichen Issues. Schreiben Sie stattdessen eine E-Mail an die Adressen für die Projektpflege in [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Geben Sie die docxi-Version an und beschreiben Sie kurz, wie sich das Problem reproduzieren lässt.

## <a id="upgrade-from-00x"></a>Von 0.0.x aktualisieren

docxi 0.0.x übergibt Bildpfad-Strings an `Kernel#open`. `Kernel#open` führt einen String, der mit `|` beginnt, als Shell-Befehl aus. Wer einen Bildpfad kontrolliert, könnte so Befehle auf Ihrem Server ausführen. 0.1.0 öffnet Dateien mit `File.open` und lädt URLs mit `URI#open` herunter. Keines von beiden führt Befehle aus.

## <a id="untrusted-input"></a>Nicht vertrauenswürdige Eingaben

**Text kann gefahrlos übergeben werden.** Text, Tabelleninhalte und andere Strings werden als XML-Text oder Attributwerte geschrieben, die nokogiri maskiert. Benutzereingaben können dem Dokument keine Elemente hinzufügen.

**Bildquellen erfordern Vorsicht.** `add_media` akzeptiert:

| Quelle | Was docxi tut |
|---|---|
| Ein IO, das Sie geöffnet haben | Liest es |
| Ein Dateipfad | Liest diese Datei von der Festplatte Ihres Servers |
| Eine `http`-, `https`- oder `ftp`-URL | Lädt sie von Ihrem Server aus herunter |

Wenn Benutzer den Pfad wählen können, können sie jede lesbare Datei Ihres Servers, etwa Konfiguration oder Schlüssel, in das Dokument einbetten lassen. Wenn Benutzer die URL wählen können, ruft Ihr Server sie ab, was auch interne Adressen erreichen kann (Server-Side Request Forgery, SSRF). Übergeben Sie ein IO, das Sie selbst geöffnet haben, oder erlauben Sie nur bekannte Pfade und Hosts.

**Hyperlinks** zeigen dorthin, wohin die URL zeigt. Prüfen Sie URLs von Benutzern, bevor Sie sie an `add_hyperlink` übergeben.

## <a id="how-releases-are-protected"></a>Wie Versionen geschützt sind

- Versionen werden aus GitHub Actions mit dem Trusted Publishing von RubyGems veröffentlicht, daher ist im Repository kein API-Schlüssel gespeichert.
- Die Gem-Metadaten setzen `rubygems_mfa_required`, daher brauchen Pushes mit einem persönlichen API-Schlüssel eine Multi-Faktor-Authentifizierung.

## <a id="related-pages"></a>Verwandte Seiten

- [Best Practices](Best-Practices)
- [Deployment](Deployment)
- [Mitwirken](Contributing)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
