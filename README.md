# Ersatzteildatenbank (HTA)

Eine eigenständige Windows-Desktop-App (HTA / HTML Application) zur Verwaltung von
Ersatzteilen. Keine Installation nötig – einfach `Ersatzteildatenbank.hta` per
Doppelklick öffnen (läuft über `mshta.exe`, ist auf jedem Windows vorinstalliert).

## Funktionen

- **Einbuchen** (`+ Ein`): Menge zum Bestand hinzufügen.
- **Ausbuchen** (`- Aus`): Menge vom Bestand abziehen. Fällt der Bestand dabei
  **unter den Grenzwert (Mindestbestand)**, erscheint sofort eine **Warnung**
  (inkl. Hinweis auf den Nachfolger, falls hinterlegt).
- **Warnung unter Grenzwert**: Betroffene Zeilen sind orange markiert, der Zähler
  in der Fußzeile zeigt, wie viele Teile unter dem Grenzwert liegen. Filter
  „nur Warnungen" blendet nur diese ein.
- **Status auf „Bestellt"**: Button `Best.` schaltet den Status um. Beim Einbuchen
  über den Grenzwert wird „Bestellt" automatisch wieder auf „Vorhanden" gesetzt.
- **Nachfolger** (wie in der Excel): Artikelnummer des Nachfolge-Teils; ist das
  Teil angelegt, ist der Eintrag anklickbar und springt dorthin.
- **Schlagwortsuche + Artikelnummer-Suche**: Ein Suchfeld durchsucht
  Artikelnummer, Bezeichnung, Schlagworte, Bauteil, Lagerort, Hersteller und
  Nachfolger. Mehrere Begriffe (durch Leerzeichen getrennt) werden UND-verknüpft.
- **Bauteil / Zuordnung**: Feld für die Zuordnung Teil → Bauteil. Kann später
  über CSV-Import oder direkt in der App ergänzt werden.
- Sortierung per Klick auf die Spaltenüberschrift, CSV-Export/-Import.

## Daten

Die Daten werden in `ersatzteile.json` **im selben Ordner** wie die HTA-Datei
gespeichert (UTF-8). Beim ersten Start wird ein Beispiel-Datensatz angelegt –
diesen einfach löschen.

## Import der bestehenden Excel-Daten

1. In Excel die relevanten Spalten in ein Tabellenblatt bringen mit **Kopfzeile**
   und diesen Feldnamen (Reihenfolge egal, nur passende Spalten werden übernommen):

   ```
   artikelnummer;bezeichnung;schlagworte;bauteil;lagerort;bestand;mindestbestand;status;nachfolger;hersteller;bemerkung
   ```

2. Als **CSV (Trennzeichen Semikolon)** speichern, z. B. `import.csv` in den
   Ordner der HTA.
3. In der App **CSV Import** klicken und den Pfad bestätigen.
   - Vorhandene Artikelnummern werden aktualisiert, neue angelegt.

> Sobald du mir die Excel-Spalten / die Zuordnung Teile→Bauteile lieferst, passe
> ich die Feldnamen bzw. eine fertige Import-Vorlage genau an deine Datei an.

## Hinweise

- Falls Windows beim Start des `ADODB.Stream`/`FileSystemObject` nachfragt: mit
  „Ja/Zulassen" bestätigen (HTAs laufen mit lokalen Rechten).
- Die App ist bewusst als **eine einzige Datei** gehalten, damit sie sich leicht
  per Netzlaufwerk oder USB weitergeben lässt.
