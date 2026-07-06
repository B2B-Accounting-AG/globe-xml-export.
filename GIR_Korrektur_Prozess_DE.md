# GIR-Korrektur (Rektifikat) — So funktioniert der Prozess

**Stand:** 2026-07-06 · Tool: GloBE XML Export App v2.7.0 (Korrekturmodus)

## Worum geht es?

Wurde eine GIR-Meldung von der ESTV **akzeptiert** und stellt sich später heraus, dass
Zahlen oder Angaben falsch waren, wird KEINE neue Meldung eingereicht, sondern eine
**Korrekturmeldung (Rektifikat)**. Die ESTV ersetzt damit gezielt die fehlerhaften Teile
des Berichts — der Rest der akzeptierten Meldung bleibt unverändert bestehen.

## Was braucht es zwingend?

| # | Voraussetzung | Warum |
|---|---------------|-------|
| 1 | Das **korrigierte, vollständige Excel-Template** | Gleiche Vorlage wie immer — einfach die falschen Werte berichtigen. Kein "Teil-Excel". |
| 2 | Das **original eingereichte Roh-XML** (aus "Roh-XML herunterladen") | Die Korrektur muss auf die internen Kennungen (DocRefIds) der akzeptierten Meldung verweisen. Diese stehen NUR im Original-XML. |

**Achtung:**
- Die **`_encrypted.zip` nützt nichts** — sie ist mit dem öffentlichen ESTV-Schlüssel
  verschlüsselt und kann von niemandem ausser der ESTV geöffnet werden.
- Die **Statusmeldung der Annahme enthält KEINE DocRefIds** (geprüft). Ist das Roh-XML
  unauffindbar: (1) bei der Person nachfragen, die die Datei generiert hat,
  (2) prüfen, ob das ePortal die eingereichte Datei erneut zum Download anbietet,
  (3) ESTV-Support um die DocRefIds bitten — unter Angabe der MessageRefId aus der
  Statusmeldung. **Ohne DocRefIds ist keine Korrektur (und auch kein Storno) möglich.**
- Wurde bereits einmal korrigiert, gilt das XML der **letzten akzeptierten Korrektur**
  als "Original" (Korrekturketten).

## Der Ablauf in der App (5 Schritte)

1. **Excel korrigieren** — der Kunde berichtigt die falschen Werte im vollständigen
   GIR-Template (die Datei bleibt komplett, auch die unveränderten Blätter).
2. **App öffnen → Schritt 1 → Meldungsart "Korrektur (Rektifikat)"** wählen.
   Beide Dateien hochladen: korrigiertes Excel + Original-Roh-XML.
3. **Vergleichstabelle prüfen** — die App vergleicht automatisch jedes Berichtselement
   (pro Jurisdiktion) und hakt die **geänderten** an. Die Auswahl mit der
   Steuerfachperson prüfen; unveränderte Elemente bleiben weg (ESTV-Vorschrift).
4. **"Datei validieren"** — die App baut die Korrekturmeldung und prüft sie doppelt:
   - **Prüfung 1:** die Korrekturdatei selbst (alle ESTV-Korrekturregeln), und
   - **Prüfung 2:** der **Gesamtbericht nach der Korrektur** (die App simuliert, wie der
     Bericht bei der ESTV nach Annahme aussieht, und validiert ihn komplett neu —
     so fallen Folgefehler in abhängigen Abschnitten VOR der Einreichung auf).
   Erst wenn ALLES grün ist, werden die Downloads freigeschaltet.
5. **"Verschlüsseln & herunterladen" → Upload in myESTV → GIR-Applikation** —
   gleicher Upload-Weg wie bei einer normalen Meldung.

## Was bekommt die ESTV?

**Nur die geänderten Teile — nicht das Excel, nicht den ganzen Bericht.**
Die Korrekturdatei enthält:

- die **FilingInfo** (Pflicht in jeder Korrektur, unverändert erneut gesendet),
- jedes **geänderte Berichtselement als Ganzes** (auch wenn nur eine Zahl anders ist),
  mit Verweis auf das zu ersetzende Original.

Beispiel: Original mit 28 Abschnitten, ein Fehler in Serbien →
die Korrekturdatei enthält nur 3 Abschnitte (FilingInfo + Konzernstruktur + Serbien).

## Test vs. Produktion

- **Test** (Abnahmeportal eportal-a.admin.ch): Datei erhält automatisch den
  `Test_`-Dateinamen und die Test-Kennzeichnungen.
- **Produktion** (eportal.admin.ch): Modus in Schritt 2 → "Erweiterte Optionen"
  auf **Produktion** stellen — die App setzt alles Übrige automatisch.
- Empfehlung: Ablauf zuerst im Testportal durchspielen (sobald es wieder verfügbar
  ist), erst dann produktiv korrigieren.

## Wichtige Regeln (macht die App automatisch)

- Korrekturmeldung = eigener Meldungstyp **GIR102** mit neuer MessageRefId.
- Jede Korrektur erhält eine **neue Kennung** und verweist auf die **letzte
  akzeptierte** Version des Elements (Korrekturkette — nie auf eine ältere).
- Das Empfängerland (RecJurCode) darf sich in einer Korrektur **nicht ändern**.
- **Neue** Berichtselemente dürfen NICHT in eine Korrektur — sie brauchen eine
  separate Neumeldung (die App blockiert das und weist darauf hin).
- Ein **Storno** (Löschung) beendet die Kette endgültig — ein gelöschtes Element
  kann nur über eine Neumeldung wieder eingereicht werden (deshalb: Korrektur
  immer dem Storno vorziehen).

## Häufige Stolpersteine

| Stolperstein | Lösung |
|--------------|--------|
| Original-XML nicht auffindbar | Rückfall-Leiter oben — ohne DocRefIds geht nichts |
| Falsches Excel als Basis | Immer vom Excel ausgehen, das der eingereichten Meldung zugrunde lag — sonst zeigt die Vergleichstabelle alles als "geändert" |
| TINs fehlen im Excel | Wie bei der Neumeldung: im TIN-Editor (Schritt 2) ergänzen — sonst bleibt der Download gesperrt |
| Test-Datei im Produktionsportal | Die App kennzeichnet Test-Dateien (`Test_…`) — Modus vor der echten Einreichung auf Produktion stellen |
