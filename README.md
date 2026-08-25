# Ersatzteildatenbank (HTA)

Eine eigenständige Windows-Desktop-App (HTA / HTML Application) zur Verwaltung von
Ersatzteilen. Keine Installation nötig – `Ersatzteildatenbank.hta` per Doppelklick
öffnen (läuft über `mshta.exe`, auf jedem Windows vorinstalliert).

## Funktionen

- **Einbuchen** (`+ Ein`) / **Ausbuchen** (`- Aus`): Menge auf die **IstMenge**
  buchen. Fällt die IstMenge beim Ausbuchen **unter die MinMenge (Grenzwert)**,
  erscheint sofort eine **Warnung** und es kann direkt auf „Bestellt" gesetzt
  werden (inkl. Hinweis auf den Nachfolger).
- **Warnung unter Grenzwert**: betroffene Zeilen orange, Zähler in der Fußzeile,
  Filter „nur Warnungen". Warnung greift, wenn **IstMenge < MinMenge**
  (IstMenge = MinMenge gilt noch als OK, z. B. 1 / 1).
- **Ausblenden**: Artikel können ausgeblendet werden (z. B. wenn ein Nachfolger
  existiert). Ausgeblendete sind standardmäßig nicht sichtbar; mit dem Toolbar-
  Häkchen „ausgeblendete anzeigen" einblendbar. Spalte `Ausgeblendet` (1/ja/x).
- **Status „Bestellt"**: Button `Best.` schaltet um. Beim Einbuchen über die
  MinMenge wird „Bestellt" automatisch zurückgesetzt.
- **Nachfolger** (wie in der Excel): Artikel-Nr. des Nachfolge-Teils; ist es
  angelegt, ist der Eintrag anklickbar und springt dorthin.
- **Schlagwort- & Artikelnummer-Suche**: ein Suchfeld über Artikelnummer,
  Erweiterte Art.Nr, Bezeichnung, Schlagworte, Bauteilgruppe, Bin, Hersteller,
  Nachfolger und Position/Typ. Mehrere Begriffe = UND-verknüpft.
- **Zuordnung (Lieferant)**: eigenes Feld/Spalte `Zuordnung` (z. B. Kardex,
  Syncore, Mosca, Assa Abloy), durchsuchbar und unter der Artikel-Nr. sichtbar.
- **Platz / Einbauort**: Feld/Spalte `Platz` (z. B. `DC30.1`), durchsuchbar –
  man kann direkt nach einem Platz suchen. Kommt ein Teil an mehreren Plätzen
  vor, stehen alle in der Zelle, mit `; ` getrennt (z. B. `DC30.1; DC30.2`).
  Die Werte stammen aus dem Kardex-Ersatzteilkatalog (Einbauort je Förderer-
  Abschnitt) und wurden über die Artikelnummer mit den Stammdaten verknüpft.
- **Passwortschutz**: Bearbeiten, Löschen, Neu anlegen und CSV-Import sind
  passwortgeschützt. Das Passwort wird im HTA-Kopf gesetzt
  (`var ADMIN_PW = "1234";` – **bitte ändern**). Ein-/Ausbuchen und
  Bestellt-Umschalten bleiben ohne Passwort möglich. Button oben rechts
  entsperrt/sperrt die Sitzung.
- **Angeheftete Kopfzeile**: die Spaltenüberschriften bleiben beim Scrollen oben.
- **Buchungsprotokoll (Audit-Trail)**: jede Ein-/Ausbuchung wird in
  `buchungen_log.csv` (neben der HTA) protokolliert – mit Zeitpunkt, Aktion,
  Artikelnummer, Bezeichnung, Menge, Bestand vorher/nachher und Windows-Benutzer.
  Über den Button „Protokoll" einsehbar und filterbar. Die Datei kann direkt in
  Excel geöffnet werden.
- Sortierung per Spaltenklick, CSV-Export/-Import.

## Datenmodell (an die Excel „Ersatzteilmatrix" angelehnt)

| Excel-Spalte        | Bedeutung in der App                     |
|---------------------|------------------------------------------|
| Position            | Positionsnummer (Anzeige/Sortierung)     |
| Menge               | **Empfohlene Menge**                      |
| Artikelnummer       | Artikelnummer (Suche, Duplikatprüfung)   |
| Erweiterte Art.Nr   | zusätzliche/erweiterte Artikelnummer     |
| Nachfolger          | Artikel-Nr. des Nachfolge-Teils          |
| Artikelbezeichnung  | Bezeichnung                              |
| Bin                 | Lagerort / Fach                          |
| Position/Typ        | Position bzw. Typ                        |
| IstMenge            | **Bestand** (Ein-/Ausbuchen)             |
| MinMenge            | **Grenzwert** für die Warnung            |
| UnterMinMeng        | wird von der App berechnet (IstMenge ≤ MinMenge) |
| Bestellt            | Status „Bestellt" (1/ja/x = ja)          |
| Bauteilgruppe       | Zuordnung Teil → Baugruppe               |
| Hersteller          | Hersteller                              |
| Schlagworte         | zusätzliche Suchbegriffe (optional)      |
| Bemerkung           | Freitext (optional)                      |

## Daten

Gespeichert in `ersatzteile.json` **im selben Ordner** wie die HTA (UTF-8).

## Excel-Import

1. In Excel das Blatt so aufbauen, dass die **Kopfzeile** die obigen Spaltennamen
   enthält (Reihenfolge egal, unbekannte Spalten werden ignoriert). Vorlage:
   `import_vorlage.csv`.
2. Als **CSV (Trennzeichen Semikolon)** speichern, z. B. `import.csv` im Ordner
   der HTA. `Bestellt`: `1`, `ja` oder `x` = bestellt.
3. In der App **CSV Import** klicken, Pfad bestätigen.
   Vorhandene Artikelnummern werden aktualisiert, neue angelegt.

> Die Zuordnung Teile → Baugruppen läuft über die Spalte **Bauteilgruppe** –
> einfach in der Excel bzw. beim Import mitgeben, oder später in der App ergänzen.

## Hinweise

- Falls Windows beim ersten `ADODB.Stream`/`FileSystemObject`-Zugriff nachfragt:
  „Ja/Zulassen".
- Bewusst als **eine Datei** gehalten (leicht per Netzlaufwerk/USB weiterzugeben).
