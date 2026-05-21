
<!-- Page 1 -->

Eidgenössisches Finanzdepartement EFD
Eidgenössische Steuerverwaltung ESTV
Hauptabteilung Direkte Bundessteuer,
Verrechnungssteuer, Stempelabgaben
Bern, März 2026
Technische Wegleitung
GloBE Information Return

<!-- Page 2 -->

Dokumentengeschichte
Datum Änderungen
18.03.2026 Initialversion
2

<!-- Page 3 -->

Inhaltsverzeichnis
Technische Wegleitung ......................................................................................................... 1
GloBE Information Return ..................................................................................................... 1
1. Einleitung ....................................................................................................................... 7
1.1 Zweck der Wegleitung ............................................................................................ 7
1.2 Zielpublikum ........................................................................................................... 7
1.3 Grundlagen des GIR .............................................................................................. 7
1.3.1 Internationale Grundlagen .................................................................................. 7
1.3.2 Innerstaatliche Grundlagen ................................................................................ 7
2. Prozesse ........................................................................................................................ 7
2.1 Einmalige Prozesse ............................................................................................... 8
2.1.1 Aufschaltung Partnerstaat .................................................................................. 8
2.1.2 Registrierung ...................................................................................................... 8
2.1.3 Abmeldung ......................................................................................................... 8
2.2 Jährlich wiederholende Prozesse ........................................................................... 9
2.2.1 Datenübermittlung .............................................................................................. 9
3. Datensicherheit und Datenschutz ................................................................................... 9
3.1 Datensicherheit ...................................................................................................... 9
3.2 Datenschutz ........................................................................................................... 9
3.3 Verschlüsselung der Daten .................................................................................... 9
3.3.1 Verschlüsselung mittels eigenem Tool .......................................................... 9
3.3.2 Verschlüsselung mittels ESTV-Encryptor .....................................................10
3.4 Datenintegrität.......................................................................................................11
4. Datenübermittlung .........................................................................................................11
4.1 GIR-Meldungen einreichen ...................................................................................11
4.1.1 Hochladen einer XML-Datei ..............................................................................11
4.2 Meldepflicht und Bestätigung der Einreichung ......................................................12
5. OECD GIR-XML-Schema ..............................................................................................12
5.1 Dateivalidierung ....................................................................................................12
5.2 Schemavalidierung ...............................................................................................12
5.3 Erweiterte Validierung ...........................................................................................13
5.3.1 Message Header ...............................................................................................13
3

<!-- Page 4 -->

5.3.2 GLOBEBody .....................................................................................................16
5.3.3 FilingInfo ...........................................................................................................17
5.3.4 GeneralSection, JurisdictionSection, UTPRAttribution.......................................18
5.3.5 DocSpec ...........................................................................................................19
6. Meldesequenzen (Storno / Korrekturen) ........................................................................22
6.1 Neumeldungen......................................................................................................22
6.2 Stornierung ganzer Meldungen .............................................................................22
6.3 Korrekturmeldungen .............................................................................................22
6.3.1 Grundsätze .......................................................................................................22
6.3.2 Aufbau einer Korrekturmeldung .........................................................................22
6.3.3 Korrekturketten .................................................................................................23
6.4 Beispiele ...............................................................................................................23
6.4.1 Korrektur eines Berichtselements ......................................................................23
6.4.2 Hinzufügen von Berichtselementen zu einer bestehenden Meldung..................24
6.4.3 Stornierung eines Berichtselements und anschliessende Neuübermittlung .......25
Korrektur der FilingInfo ..................................................................................................26
7. Anhang ..........................................................................................................................28
7.1 Zulässiger Zeichensatz .........................................................................................28
4

<!-- Page 5 -->

Abkürzungen und Begriffe
CE Constituent Entity (Konstituierender Rechtsträger)
CHF Schweizer Franken
DFE Designated Filing Entity (Benannte Einreichungsstelle)
ESTV Eidgenössische Steuerverwaltung
GIR GloBE Information Return
GloBE Global Anti-Base Erosion
ID Identifikationsnummer
Inbound Von Partnerstaaten eingehende Informationen
MCAA Multilateral Competent Authority Agreement
OECD Organisation für wirtschaftliche Zusammenarbeit und Entwicklung
Outbound Von der Schweiz ausgehende Informationen
Partnerstaat Staat oder Hoheitsgebiet, mit welchem die Schweiz den
Datenaustausch vereinbart hat
SIN Steueridentifikationsnummer; auch TIN
TIN Taxpayer Identification Number; auch SIN
UID Unternehmens-Identifikationsnummer
UPE Ultimate Parent Entity (Konzernobergesellschaft)
UUID Universally Unique Identifier
XML Extensible Markup Language
5

<!-- Page 6 -->

Referenzen
Nr. Dokument / Link
[1] Tax Challenges Arising from the Digitalisation of the Economy – GloBE Information
Return (January 2025)
https://www.oecd.org/en/publications/tax-challenges-arising-from-the-digitalisation-of-
the-economy-globe-information-return-january-2025_a05ec99a-en.html
[2] Multilateral Competent Authority Agreement on the Exchange of GloBE Information
(January 2025)
https://www.oecd.org/en/topics/sub-issues/global-minimum-tax/global-anti-base-
erosion-model-rules-pillar-two.html#globe-information-return
[3] GloBE Information Return (Pillar Two) XML Schema
https://www.oecd.org/en/publications/globe-information-return-pillar-two-xml-
schema_c594935a-en.html
[4] GloBE Information Return (Pillar Two) Status Message XML Schema
https://www.oecd.org/en/publications/globe-information-return-pillar-two-status-
message-xml-schema_449e3cc3-en.html
Geschlechtsneutrale Formulierung
Aus Gründen der besseren Lesbarkeit wird im Text nur die männliche Form verwendet.
Gemeint ist stets sowohl die weibliche als auch die männliche Form.
6

<!-- Page 7 -->

1. Einleitung
1.1 Zweck der Wegleitung
Die vorliegende Wegleitung beschreibt und konkretisiert die Prozesse und Abläufe, die sich
bei den ergänzungssteuerpflichtigen Geschäftseinheiten und der ESTV in Bezug auf die
technische Umsetzung des GIR ergeben.
1.2 Zielpublikum
Die vorliegende Wegleitung richtet sich an grosse Unternehmensgruppen, die verpflichtet
sind, bei der ESTV eine GloBE Information Return (GIR) einzureichen.
1.3 Grundlagen des GIR
1.3.1 Internationale Grundlagen
Am 14. Dezember 2021 hat die OECD im Rahmen des BEPS-Projekts den Bericht über die
Einführung einer globalen effektiven Mindestbesteuerung («Pillar Two») veröffentlicht. Ziel
dieser Initiative ist es, sicherzustellen, dass grosse multinationale Unternehmensgruppen
unabhängig von ihrem Sitzstaat einem Mindeststeuersatz von 15 % unterliegen. Zur
Umsetzung dieser Mindestbesteuerung und zum automatischen Austausch relevanter Daten
wurde der GloBE Information Return (GIR) [vgl. Referenz Nr. 1, 2 und 3 hievor] entwickelt.
Der GIR dient dem Informationsaustausch zwischen den Steuerverwaltungen der
teilnehmenden Staaten. Er enthält detaillierte Angaben zu den weltweit erzielten Gewinnen,
gezahlten Steuern sowie zu den auf Konzernebene berechneten effektiven Steuersätzen.
Die Multilaterale Vereinbarung der zuständigen Behörden über den Austausch von Global-
Anti-Base-Erosion-Informationen stellt die internationale Grundlage für den automatischen
Austausch der GIR dar.
1.3.2 Innerstaatliche Grundlagen
Die Verordnung über die Mindestbesteuerung grosser Unternehmensgruppen
(Mindestbesteuerungsverordnung, MindStV) ist am 1. Januar 2024 in Kraft getreten und am
1. Januar 2026 wurde die MindStV um die innerstaatlichen Bestimmungen betreffend GIR
ergänzt.
2. Prozesse
Beim GIR-Austausch wird zwischen Inbound- und Outbound-Prozessen unterschieden.
• Inbound-Prozesse: Die Daten werden von den Partnerstaaten an die ESTV
übermittelt und können bei dieser von den kantonalen Steuerverwaltungen abgerufen
werden.
• Outbound-Prozesse: Die Daten werden von den ergänzungssteuerpflichtigen
Geschäftseinheiten sowie steuerlich der Schweiz zugehörigen Geschäftseinheiten an
die ESTV übermittelt und von dieser an die Partnerstaaten weitergeleitet. Auch die
Outbound-Daten können von den kantonalen Steuerverwaltungen bei der ESTV
abgerufen werden.
Da sich die vorliegende Wegleitung an ergänzungssteuerpflichtige Geschäftseinheiten sowie
steuerlich der Schweiz zugehörige Geschäftseinheiten gemäss Artikel 28c Absatz 2
Buchstabe a MindStV richtet, wird im Folgenden nur auf den Outbound-Prozess
eingegangen.
7

<!-- Page 8 -->

Es kann zwischen einmaligen und sich jährlich wiederholenden Prozessen unterschieden
werden.
Die GIR ist spätestens 15 Monate (18 Monate für die erste GIR) nach dem Ende der
Berichtsteuerperiode an die ESTV zu übermitteln (vgl. Art. 28d Abs. 1 i.V.m. Art. 20 Abs. 1
MindstV).
Beispiel: Steuerperiode 1.1.2025 – 31.12.2025;
→ Die GIR muss spätestens am 31.03.2027 bei der ESTV eingereicht werden.
2.1 Einmalige Prozesse
2.1.1 Aufschaltung Partnerstaat
Wenn der GIR-Austausch mit einem Partnerstaat vereinbart wurde, wird das Land in die
Liste der Partnerstaaten aufgenommen. Die Liste wird auf der Internetseite des
Staatssekretariats für internationale Finanzfragen aufgeschaltet:
https://www.sif.admin.ch/de/globe-information-return-gir
2.1.2 Registrierung
Die ergänzungssteuerpflichtige Geschäftseinheit oder die steuerlich der Schweiz zugehörige
Geschäftseinheit gemäss Artikel 28c Absatz 2 Buchstabe a MindStV muss sich innert der
Fristen nach Artikel 28d i.V.m. Art. 20 MindStV unaufgefordert bei der ESTV für den GIR-
Austausch anmelden.
Die Registrierung als ergänzungssteuerpflichtige Geschäftseinheit oder steuerlich der
Schweiz zugehörige Geschäftseinheit gemäss Artikel 28c Absatz 2 Buchstabe a MindStV hat
über die GIR-Anwendung im ePortal zu erfolgen. Diese wird von der ESTV unter
https://myestv.estv.admin.ch/home zur Verfügung gestellt.
Bei Fragen bezüglich des Vorgehens stehen Ihnen Erklärvideos zum Verwalten von
Unternehmen und Personen auf myESTV (Klick auf das Fragezeichen oben rechts) und zum
Authentifizierungsdienst der Schweizer Behörden AGOV (https://www.agov.admin.ch/de/info-
d-video) sowie der Support unter +41 58 461 61 11 zur Verfügung.
Der Registrierungsprozess gilt als abgeschlossen, sobald ein Administrator der
ergänzungssteuerpflichtigen Geschäftseinheit oder der steuerlich der Schweiz zugehörigen
Geschäftseinheit gemäss Artikel 28c Absatz 2 Buchstabe a MindstV im ePortal freigeschaltet
und die Registrierungsinformationen in der GIR-Applikation vervollständigt wurden. Es wird
keine Registrierungsbestätigung versendet.
2.1.3 Abmeldung
Die ergänzungssteuerpflichtige Geschäftseinheit oder die steuerlich der Schweiz zugehörige
Geschäftseinheit gemäss Artikel 28c Absatz 2 Buchstabe a MindstV muss sich
unaufgefordert im Informationssystem abmelden, wenn ihre Pflicht, eine GIR einzureichen,
endet (vgl. Art. 28m Abs. 3 MindstV). Die Abmeldung muss schriftlich per Brief oder per E-
Mail beantragt werden.
Adresse:
Eidgenössische Steuerverwaltung, Abteilung Informationsaustausch in Steuersachen, Team
AIA, Eigerstrasse 65, 3003 Bern.
8

<!-- Page 9 -->

E-Mail:
info-gir@estv.admin.ch
2.2 Jährlich wiederholende Prozesse
2.2.1 Datenübermittlung
Die GIR ist spätestens 15 Monate (18 Monate für die erste GIR) nach dem Ende der
Berichtsteuerperiode an die ESTV zu übermitteln (vgl. Art. 28d, Abs. 1 MindstV).
Versäumt die ergänzungssteuerpflichtige Geschäftseinheit oder die steuerlich der Schweiz
zugehörige Geschäftseinheit gemäss Artikel 28c Absatz 2 Buchstabe a MindstV die
Einreichungsfrist, so wird die ergänzungssteuerpflichtige Geschäftseinheit für jeden Tag
zwischen dem Ende der Frist und der Einreichung der GIR mit einem Betrag von
200 Franken belastet, höchstens jedoch mit 50 000 Franken (vgl. Art. 28e MindstV).
Die Verwaltungssanktion bei Säumnis nach Artikel 28e sowie die Strafbarkeit bei fahrlässiger
Verletzung von Verfahrenspflichten oder fahrlässiger Steuerhinterziehung nach den Artikeln
29 und 30 MindStV entfallen für alle Geschäftsjahre nach Artikel 10.1 der GloBE-
Mustervorschriften, die bis am 31. Dezember 2026 beginnen und bis am 30. Juni 2028
enden (vgl. Art. 40 Abs. 3 MindStV).
Der Vollständigkeit halber sei auf die Überprüfungen verwiesen, welche die ESTV gestützt
auf Artikel 28o MindstV vornimmt. Dabei handelt es sich weder um einen einmaligen noch
um einen sich im Jahresrhythmus wiederholenden Prozess. Vielmehr werden die Kontrollen
periodisch gemäss Einschätzung der ESTV durchgeführt.
3. Datensicherheit und Datenschutz
3.1 Datensicherheit
Die Datensicherheit ist umfassend gewährleistet. Bei der Dateneinlieferung via XML-Datei-
Upload wird die Meldung durch die ergänzungssteuerpflichtige Geschäftseinheit
verschlüsselt und der Transport erfolgt über einen sicheren, verschlüsselten Kanal.
In der GIR-Applikation werden lediglich die Metadaten der erfolgten Dateneinlieferungen
angezeigt. Nach erfolgter Dateneinlieferung (XML-Datei-Upload) werden die Daten
entschlüsselt, validiert und erneut verschlüsselt und sicher abgelegt.
Die GIR-Applikation wird periodisch einem Sicherheitspenetrationstest durch eine externe,
unabhängige Firma unterzogen.
3.2 Datenschutz
Bei der Konzeption und Umsetzung der GIR-Applikation wurden alle relevanten
Anforderungen betreffend den Datenschutz entsprechend berücksichtigt.
3.3 Verschlüsselung der Daten
3.3.1 Verschlüsselung mittels eigenem Tool
Um eine GIR im Portal hochzuladen, muss die GIR-XML-Datei komprimiert und verschlüsselt
werden. Dazu müssen die folgenden Schritte durchgeführt werden:
9

<!-- Page 10 -->

Prozessbeschreibung Ergebnis (Dateiname)
1. Komprimieren der GIR-XML-Datei Payload.zip
• Die GIR-XML-Datei muss „Payload.xml“ genannt
werden
• Erstellen einer Zip-Datei mit Inhalt "Payload.xml"
2. Verschlüsseln der komprimierten Datei Payload
• Die Datei Payload.zip wird mittels AES-256
verschlüsselt
• Erzeugen eines AES-256 Schlüssels
• Cipher mode: CBC (Cipher Block Chaining)
• Initialization Vector (IV): 16 byte IV
• Key size: 256 bits/32 bytes
• Encoding: None
• Padding: PKCS#7 oder PKCS#5
Hinweise:
Für die Sicherheit der Verschlüsselung ist es wichtig, dass
der Initialisierungsvektor jedes Mal neu erzeugt wird.
Für Implementierungen basierend auf Java: Die Sun
Implementierung kennt kein PKCS#7, hier sollte PKCS#5
verwendet werden.
3. Verschlüsseln des AES-Schlüssels und IV-Parameters mit Key
dem Public Key aus dem ESTV-Zertifikat
• AES-Schlüssel und IV werden vor der
Verschlüsselung zusammengesetzt (48 bytes total
- 32 byte AES-Schlüssel und 16 byte IV)
• Verschlüsseln dieser 48 bytes:
• Algorithmus: RSA
• Padding: PKCS#1 v1.5
4. Übertragungspaket erstellen Zip-Datei mit beliebigem
• Erstellen einer Zip-Datei mit folgendem Inhalt Dateinamen und
• "Payload" (Datei aus Schritt 2) Endung .zip
• "Key" (Datei aus Schritt 3)
Das öffentliche ESTV-Zertifikat für die Verschlüsselung im 3. Schritt kann im Portal
heruntergeladen werden.
3.3.2 Verschlüsselung mittels ESTV-Encryptor
a) Vorbereitung
Laden Sie das Archiv «estv-encryptor-*.zip» von der Applikation herunter. Dieses Archiv
enthält den ESTV-Encryptor sowie den Public Key (ESTV-PublicKey.pem). Entpacken Sie
das Archiv und speichern Sie alle Dateien in einem gemeinsamen, neuen Ordner auf Ihrem
Computer. Dies vereinfacht den Prozess und verhindert Fehler.
b) Dateien vorbereiten
Kopieren oder verschieben Sie alle zu verschlüsselnden XML-Dateien in denselben Ordner,
in dem sich der Encryptor befindet. Der Encryptor verarbeitet nur Dateien in seinem eigenen
Verzeichnis.
10

<!-- Page 11 -->

c) Auswahl der Anwendung
Öffnen Sie Ihren Datei-Explorer (Windows) oder Finder (macOS/Linux) und navigieren Sie
zum vorbereiteten Ordner. Starten Sie die für Ihr Betriebssystem vorgesehene Anwendung
durch einen Doppelklick:
• Windows: estv-encryptor-win.exe
• macOS: estv-encryptor-macos (Unter Umständen müssen Sie in den Sicherheitseinstellungen
die Ausführung erst erlauben)
• Linux: estv-encryptor-linux (Stellen Sie sicher, dass die Datei ausführbar ist: chmod +x estv-
encryptor-linux)
d) Verschlüsselung
Die Anwendung verschlüsselt automatisch alle XML-Dateien (.xml-Endung) im Ordner.
Für jede erfolgreich verarbeitete XML-Datei wird eine verschlüsselte ZIP-Archivdatei mit
identischem Basisnamen erstellt.
Beispiel: muster_2023.xml → muster_2023.zip
Wichtig: Die originalen XML-Dateien bleiben unverändert erhalten.
e) Erfolgskontrolle
Überprüfen Sie nach Abschluss des Vorgangs den Ordner. Für jede XML-Datei sollte nun
eine entsprechende .zip-Datei vorhanden sein. Sie können das Terminalfenster schliessen.
f) Weiterverarbeitung
Die generierten .zip-Dateien sind nun verschlüsselt und bereit für den Upload. Laden Sie
diese Dateien in die Anwendung der ESTV hoch. Stellen Sie vor dem Upload sicher, dass
jede .zip-Datei nur die beiden erforderlichen Komponenten enthält: den verschlüsselten
Sitzungsschlüssel (key) und die verschlüsselte Nutzlast (payload).
3.4 Datenintegrität
Nach dem Übermitteln einer Meldung per Upload lässt sich in der GIR-Meldungsübersicht
überprüfen, ob die Meldung korrekt übertragen wurde. Nur wenn die Meldung den Status
«Akzeptiert» erhält, hat die ergänzungssteuerpflichtige Geschäftseinheit ihre Meldepflicht für
die Berichtsperiode erfüllt.
4. Datenübermittlung
4.1 GIR-Meldungen einreichen
Die GIR-Meldungen werden durch Hochladen einer XML-Datei via GIR-Applikation (XML-
Datei-Upload) eingereicht.
4.1.1 Hochladen einer XML-Datei
Beim XML-Datei-Upload können Dateien im GIR-XML-Format hochgeladen und der ESTV
übermittelt werden. Die Erstellung der GIR-XML-Datei erfolgt durch die
ergänzungssteuerpflichtige Geschäftseinheit.
Daten, die über die GIR-Applikation hochgeladen werden, müssen als maximal 100 MB
grosse Dateien im XML-Format vorliegen. Die Dateien müssen für den Upload komprimiert
und verschlüsselt (vgl. Ziffer 3.3.) werden und dürfen komprimiert maximal 10 MB gross sein.
11

<!-- Page 12 -->

Der Transport zu den Systemen der ESTV wird zusätzlich durch eine verschlüsselte
Verbindung (HTTPS) abgesichert.
Per Upload können sowohl Neumeldungen als auch Korrektur- und Stornomeldungen
übermittelt werden. Die übermittelten Dateien müssen nach den Vorgaben in Ziffer 5 erstellt
werden, ansonsten wird die komplette Meldung als fehlerhaft zurückgewiesen.
4.2 Meldepflicht und Bestätigung der Einreichung
Die ergänzungssteuerpflichtige Geschäftseinheit hat die GIR jährlich spätestens 15 Monate
nach dem Ende der Berichtssteuerperiode an die ESTV zu übermitteln (vgl. Art. 28d Abs. 1
i.V.m. Art. 20 Abs. 1 MindStV). Es obliegt der ergänzungssteuerpflichtigen Geschäftseinheit,
in der GIR-Meldungsübersicht der GIR-Applikation zu überprüfen, dass sie für jeden
eingereichten Bericht eine positive Validierungsbestätigung erhalten hat. Es kann nach dem
Einreichen einige Minuten dauern, bis das Validierungsergebnis vorliegt.
5. OECD GIR-XML-Schema
Eine detaillierte Beschreibung des GIR-XML-Schemas findet sich in Referenz Nr. [3] hievor.
Die Validierungsregeln und Fehlercodes richten sich nach den Definitionen im GIR Status
Message Guide [vgl. Referenz Nr. 4 hievor].
Im Folgenden wird beschrieben, welchen Vorgaben und Validierungsregeln eine GIR-XML-
Datei entsprechen muss, um von der ESTV entgegengenommen und verarbeitet werden zu
können. Das Ergebnis der Validierung kann in der GIR-Applikation in der GIR-
Meldungsübersicht abgerufen werden.
5.1 Dateivalidierung
In einem ersten Schritt wird die übermittelte Datei überprüft. Falls bereits in diesem ersten
Schritt Fehler auftreten, wird direkt eine Fehlermeldung angezeigt. In diesem Fall wird kein
Eintrag in der Liste der Meldungen erzeugt und auch keine Statusmeldung erstellt.
In diese Kategorie fallen die folgenden Fehlercodes:
Fehler Fehlercode Beschreibung
Failed Download 50001 Die Datei wurde nicht korrekt übertragen oder ist fehlerhaft und
k ann nicht geöffnet werden.
Failed Decryption 50002 Die Datei konnte nicht entschlüsselt werden.
Failed Decompression 50003 Die Datei konnte nicht dekomprimiert werden.
Failed Threat Scan 50005 Es wurde eine potenzielle Bedrohung in der Datei entdeckt.
Failed Virus Scan 50006 Es wurde ein Virus in der Datei entdeckt.
Auch die Validierungen mit 50’000er-Fehlercodes im Message Header (vgl. Ziffer 5.3.1)
gelten als Dateivalidierungen.
5.2 Schemavalidierung
Wenn die vorherigen Prüfungen erfolgreich waren, wird in einem zweiten Schritt die Datei
geöffnet und mit dem GIR-XML-Schema [vgl. Referenz Nr. 3 hievor] verglichen.
Die Schemavalidierung überprüft, ob die Meldung dem GIR-XML-Schema entspricht. Falls
nicht, wird die Meldung als Ganzes zurückgewiesen.
12

<!-- Page 13 -->

Fehler Fehlercode Beschreibung
Failed Schema Validation 50007 Die Datei entspricht nicht dem GIR-XML-Schema
Hinweis: Die GIR-Meldungen dürfen nicht signiert werden. Eine Signatur des XML führt
ebenfalls dazu, dass die Meldung zurückgewiesen wird.
5.3 Erweiterte Validierung
Nach der Datei- und Schemavalidierung wird der Inhalt einzelner Elemente geprüft. Im
Folgenden werden nicht alle Regeln beschrieben, es gelten grundsätzlich die Regeln und
Fehlercodes der OECD [vgl. Referenz Nr. 4]. Beschrieben sind nur die wichtigsten Regeln
sowie die Regeln, welche für den innerstaatlichen Datenaustausch von der OECD-
Spezifikation abweichen.
Im Gegensatz zur OECD-Spezifikation gelten für den Datenaustausch zwischen den
ergänzungssteuerpflichtigen Geschäftseinheiten und der ESTV alle Regeln als verpflichtend.
Jeder Fehler führt also zur Ablehnung der kompletten Meldung.
Zusätzlich wurden ergänzende Regeln definiert (Fehlercodes 98000-98999), welche nur für
den innerstaatlichen Austausch innerhalb der Schweiz gelten. Diese Regeln werden im
folgenden Abschnitt im Detail erläutert.
5.3.1 Message Header
Die im Message Header (MessageSpec) angegebenen Daten werden nicht an die
Partnerstaaten übermittelt. Die ESTV generiert bei der Erstellung der Meldungen an die
Partnerstaaten einen neuen MessageSpec. Dennoch müssen hier einige Daten erfasst
werden, damit die Meldungen durch die ESTV korrekt verarbeitet werden können.
SendingEntityIN
Als Identifikationsnummer der ergänzungssteuerpflichtigen Geschäftseinheit muss die UID
oder die ESTV-ID eingetragen werden. Basierend auf dieser Angabe wird die Meldung der
ergänzungssteuerpflichtigen Geschäftseinheit zugeordnet. Der Benutzer muss über eine
gültige Berechtigung für den entsprechenden Partner verfügen. Da im Fehlerfall die Meldung
keinem Partner zugeordnet werden kann, verhält sich diese Validierung ebenfalls wie eine
Dateivalidierung, es wird also eine Fehlermeldung angezeigt und keine Statusmeldung
erzeugt.
Die Nummer darf sowohl formatiert (CHE-123.456.789 bzw. 052.1234.5678) als auch
unformatiert (123456789 bzw. 5212345678) übermittelt werden. Andere Buchstaben ausser
"CHE" dürfen nicht enthalten sein.
Regel Validierung Fehlercode
Muss die UID der Wert = UID oder ESTV-ID 98001
ergänzungssteuerpflichtigen
Geschäftseinheit enthalten.
TransmittingCountry
Der ISO-Ländercode des Senderstaates ist hier immer die Schweiz.
Regel Validierung Fehlercode
ISO-Ländercode der Schweiz Wert = „CH“ 98002
13

<!-- Page 14 -->

ReceivingCountry
ISO-Ländercode des Empfängerstaates. Da die Meldungen der ergänzungssteuerpflichtigen
Geschäftseinheit an die ESTV übermittelt werden und Daten für mehrere Empfängerstaaten
enthalten können, muss hier als Empfängerstaat „CH“ eingetragen werden.
Regel Validierung Fehlercode
ISO-Ländercode des Empfängerstaates Wert = „CH“ 50010
MessageType
Kennzeichnet die Art der Meldung. Hier muss immer der Wert „GIR“ stehen. Dies ist durch
das GIR-XML-Schema vorgegeben und der korrekte Wert wird bereits bei der
Schemavalidierung geprüft.
Warning
Dieses Element wird für die Übermittlung zwischen ergänzungssteuerpflichtigen
Geschäftseinheiten und der ESTV nicht verwendet. Daten in diesem Element werden von
der ESTV weder validiert noch ausgewertet oder weitergeleitet.
Contact
Dieses Element wird für die Übermittlung zwischen ergänzungssteuerpflichtigen
Geschäftseinheiten und der ESTV nicht verwendet. Daten in diesem Element werden von
der ESTV weder validiert noch ausgewertet oder weitergeleitet.
MessageRefId
Dies ist der Unique Identifier für die gesamte Meldung. Die MessageRefId ist
zusammenzusetzen aus:
TransmittingCountry + Jahr + ReceivingCountry + Eindeutige ID
TransmittingCountry und ReceivingCountry sind immer “CH” (siehe oben).
Als Jahr muss das Jahr aus dem Feld «ReportingPeriod» verwendet werden, welches dem
Ende der Berichtssteuerperiode entspricht.
Die Berichtssteuerperiode entspricht dem Geschäftsjahr, für welches die Angaben im GIR-
Bericht festgehalten werden. Weicht die Berichtssteuerperiode vom Kalenderjahr ab, ist der
letzte Tag der Berichtssteuerperiode entscheidend.
Beispiel: Berichtssteuerperiode 1.7.2024 - 30.6.2025; es muss das Jahr 2025 angegeben
werden.
Um den GIR-Bericht bei der ESTV einreichen zu können, muss die
ergänzungssteuerpflichtige Geschäftseinheit für das betreffende Geschäftsjahr registriert
sein. Es ist deshalb wichtig, bei der Registrierung korrekt anzugeben, ab welchem Jahr
Berichte eingereicht werden müssen.
Die MessageRefId muss global eindeutig sein, daher muss nach den vorgeschriebenen
Elementen eine eindeutige ID folgen, um sicherzustellen, dass weder eine frühere Meldung
noch eine andere ergänzungssteuerpflichtige Geschäftseinheit die gleiche ID verwendet. Wir
empfehlen die Verwendung einer UUID nach RFC 4122.
Beispiel: CH2024CH8b0f7048-e2ff-11e6-bf01-fe55135034f3
14

<!-- Page 15 -->

Hinweis: Als MessageRefId dürfen keine vertraulichen Informationen verwendet werden, da
die MessageRefId in Fehlermeldungen und Validierungsbestätigungen sowie in den
Metadaten des Berichts unverschlüsselt gespeichert wird.
Folgende Einschränkungen gelten für die MessageRefId:
• Die MessageRefId darf maximal 170 Zeichen lang sein
• Der Ländercode muss in Grossbuchstaben geschrieben werden
• Zulässig sind alle Zeichen gemäss Ziffer 7.1
Als regulärer Ausdruck: CH[0-9]{4}CH.{1,162}
Regel Validierung Fehlercode
Die Struktur der MessageRefId muss dem Wert = "CH" & Jahr der 60001
vorgegebenen Schema entsprechen. ReportingPeriod & "CH" &
UUID
Darf nicht gleich der MessageRefId einer Wert ≠ frühere MessageRefId 60002
früheren Meldung sein.
MessageTypeIndic
Um die Kategorisierung der Meldungen zu erleichtern, muss dieses Element für die
Übermittlung zwischen ergänzungssteuerpflichtigen Geschäftseinheiten und ESTV immer
ausgefüllt werden.
Erlaubte Werte sind: "GIR101" (Neumeldung) / "GIR102" (Korrektur) / "GIR103"
(Nullmeldung).
Da gemäss GIR User Guide [vgl. Referenz Nr. 3 hievor] die Mischung von Neu- und
Korrekturmeldungen nicht erlaubt ist, dürfen je nach Wert in diesem Element im GLOBEBody
nur entweder neue Daten (DocTypeIndic "OECD1") oder Korrekturen/Löschungen
(DocTypeIndic "OECD2" oder "OECD3") vorhanden sein.
Rechtsträger, welche sich als Einheit eines Konzerns registriert haben, welcher in einem
anderen Staat den GIR-Bericht einreicht, können Nullmeldungen mit MessageTypeIndic
"GIR103" einreichen. In diesem Fall darf die Meldung keinen GIR-Bericht enthalten, d.h. der
GLOBEBody darf dann keine Berichtselemente ausser FilingInfo enthalten.
Regel Validierung Fehlercode
Eine Neumeldung darf keine Korrekturen Falls Wert = „GIR101“: Kein 60004
oder Stornos enthalten. DocTypeIndic in der ganzen
Meldung darf den Wert
(„OECD2“ oder „OECD3“)
haben.
Eine Korrekturmeldung darf keine neuen Falls Wert = „GIR102“: Kein 60004
Daten enthalten. DocTypeIndic in der ganzen
Meldung darf den Wert
(„OECD1“) haben.
Eine Nullmeldung darf keinen GIR-Bericht Falls Wert = "GIR103", darf 98009
enthalten. Es darf auch zuvor noch kein GLOBEBody keine Elemente
GIR-Bericht für das Berichtsjahr eingereicht ausser "FilingInfo" enthalten.
worden sein. Ausserdem dürfen keine
früheren Meldungen für die
gleiche ReportingPeriod
vorliegen, die nicht gelöscht
wurden.
15

<!-- Page 16 -->

Regel Validierung Fehlercode
Die ergänzungssteuerpflichtige Wert = "GIR103" nur für CE 98010
Geschäftseinheit muss als CE registriert und nur, wenn die Schweiz den
sein, und der GIR muss der Schweiz von GIR von einem Partnerstaat
einem anderen Staat zugestellt werden. erhält.
ReportingPeriod
Dieses Datenelement gibt den letzten Tag der Berichtssteuerperiode an, auf den sich die
Meldung bezieht. Es können nur Meldungen für vergangene Berichtssteuerperioden
eingereicht werden.
Das Datum muss im Format JJJJ-MM-TT angegeben werden. Beispiel:
Berichtssteuerperiode 1.4.2024 – 31.3.2025; ReportingPeriod = 2025-03-31.
Regel Validierung Fehlercode
Es können keine Meldungen für Wert <= aktuelles Datum 60003
Berichtssteuerperioden eingereicht werden,
die noch nicht zu Ende sind.
Der letzte Tag der Berichtssteuerperiode Wert = Enddatum des 98004
muss dem Ende des Geschäftsjahres der Geschäftsjahres
ergänzungssteuerpflichtigen
Geschäftseinheit entsprechen.
Falls das Geschäftsjahr an einem anderen
Datum endet als im Vorjahr, muss dies der
ESTV mitgeteilt werden (info-
gir@estv.admin.ch).
Timestamp
Der Zeitstempel gibt an, wann die Meldung erstellt wurde. Dies soll eine sinnvolle Angabe
sein, der Wert soll also nicht in der Zukunft liegen – mit einer gewissen Toleranz, da die
Systemzeiten in IT-Systemen nicht immer völlig synchron laufen. Der Wert sollte zudem
auch nicht allzu weit in der Vergangenheit liegen.
Regel Validierung Fehlercode
Der Wert darf nicht mehr als einen Tag in Aktuelles Datum und Zeit 98008
der Zukunft und nicht mehr als ein Jahr in − 1 Jahr ≤ Wert ≤ aktuelles
der Vergangenheit liegen. Datum und Zeit + 1 Tag
5.3.2 GLOBEBody
Der GLOBEBody enthält den eigentlichen GIR-Bericht, welcher an die Partnerstaaten
weitergeleitet wird. Jedes Berichtselement enthält einen RecJurCode ("Receiving Jurisdiction
Code"), welcher anzeigt, an welchen Partnerstaat das Element weitergeleitet werden soll.
Aufgrund dieser Information sortiert die ESTV die Daten in ausgehende Meldungen, welche
an die jeweiligen Partnerstaaten übermittelt werden.
Der Inhalt des Berichts besteht aus einem Element FilingInfo, welches Daten über die
ergänzungssteuerpflichtige Geschäftseinheit enthält. Anschliessend folgt eine
GeneralSection, danach beliebig viele Elemente Summary, JurisdictionSection und
UTPRAttribution.
Im internationalen Datenaustausch kann das Element GLOBEBody wiederholt werden, um
die Daten mehrerer ergänzungssteuerpflichtiger Geschäftseinheiten an einen anderen
Partnerstaat zu senden. Bei der Übermittlung zwischen ergänzungssteuerpflichtigen
Geschäftseinheiten und ESTV kann eine Meldung jedoch immer nur die Daten genau einer
16

<!-- Page 17 -->

ergänzungssteuerpflichtigen Geschäftseinheit enthalten, entsprechend darf es nur einen
GLOBEBody geben.
Regel Validierung Fehlercode
GLOBEBody darf für Übermittlungen Element darf nur einmal 98100
zwischen ergänzungssteuerpflichtigen vorkommen.
Geschäftseinheiten und ESTV nicht
wiederholt werden.
5.3.3 FilingInfo
Im Abschnitt FilingInfo müssen die Daten der ergänzungssteuerpflichtigen Geschäftseinheit
angegeben werden.
ResCountryCode
Hier wird das Land angegeben, in dem die ergänzungssteuerpflichtige Geschäftseinheit
steuerlich ansässig ist. Dieser Wert muss dem TransmittingCountry in der MessageSpec
entsprechen, also der Schweiz. Ein ergänzungssteuerpflichtige Geschäftseinheit, der nicht in
der Schweiz ansässig ist, muss der ESTV keine Daten melden.
Regel Validierung Fehlercode
Der ergänzungssteuerpflichtige Der Wert muss 60023
Geschäftseinheit muss seine steuerliche TransmittingCountry
Ansässigkeit in der Schweiz haben. entsprechen.
TIN
Die Steueridentifikationsnummer (deutsch: SIN; englisch: TIN) der
ergänzungssteuerpflichtigen Geschäftseinheit entspricht seiner UID. Hier wird geprüft, dass
der Wert vorhanden ist und dass er der UID der ergänzungssteuerpflichtigen
Geschäftseinheit entspricht, welche die Meldung übermittelt hat. Ausserdem muss der Typ
korrekt als „TIN“ (GIR3001) und als Ausgabeland die Schweiz angegeben sein.
Regel Validierung Fehlercode
Als TIN muss die korrekte UID der Wert = UID 98101
ergänzungssteuerpflichtigen
Geschäftseinheit angegeben werden.
Die TIN muss den Typ "TIN" (GIR3001) TypeOfTIN = "GIR3001" 98103
haben.
Die TIN der FilingCE muss von der Schweiz issuedBy = “CH” 98104
ausgegeben worden sein.
Role
Hier muss die Rolle der ergänzungssteuerpflichtigen Geschäftseinheit angegeben werden.
Diese muss der Rolle entsprechen, welche bei der Registrierung des Rechtsträgers
ausgewählt wurde.
Rolle bei der Registrierung ReportingRole
UPE GIR401
(Konzernobergesellschaft) Die übliche Rolle, wenn die Konzernobergesellschaft
den Bericht selber einreicht.
DFE GIR402
(Benannte Einreichungsstelle) Wenn die Konzernobergesellschaft eine andere
Konzerneinheit benennt, welche die GIR-Berichte
17

<!-- Page 18 -->

Rolle bei der Registrierung ReportingRole
einreicht, muss die benannte Einheit sich mit dieser
Rolle registrieren.
CE GIR404
(Konstituierende Einheit) Diese Rolle dient der Deklaration, dass die
ergänzungssteuerpflichtige Konzerneinheit zu einem
Konzern gehört, welcher in einem anderen Staat
ansässig ist. Falls in diesem Staat eine GIR-Meldung
eingereicht wird, welche mit der Schweiz ausgetauscht
wird, muss die CE nur eine Nullmeldung einreichen.
Falls er in einem Nicht-Partnerstaat ansässig ist, muss
die CE eine GIR-Meldung einreichen, welche aber
nicht international ausgetauscht wird.
Regel Validierung Fehlercode
Es muss dieselbe Rolle angegeben FilingCE.Role = Rolle bei 98106
werden, welche bei der Registrierung Registrierung
angegeben wurde.
5.3.4 GeneralSection, JurisdictionSection, UTPRAttribution
In diesem Bereich werden nur noch einzelne Regeln beschrieben, welche von der OECD-
Spezifikation abweichen. Es gelten aber grundsätzlich alle Validierungsregeln, welche von
der OECD vorgegeben sind [vgl. Referenz Nr. 4 hievor].
RecJurCode
Der Ländercode des Empfängerlandes (RecJurCode) definiert, an welche Partnerstaaten die
jeweilige Sektion des Berichts übermittelt wird. Daher gelten hier einige spezielle Regeln.
Die OECD gibt in der Regel 60018 vor, dass der RecJurCode dem ReceivingCountry der
Meldung entsprechen muss. Da wir im innerstaatlichen Austausch zwischen der
ergänzungssteuerpflichtigen Geschäftseinheit und der ESTV immer "CH" als
ReceivingCountry verwenden, funktioniert das natürlich nicht. Stattdessen muss der
RecJurCode einem Partnerstaat der Schweiz entsprechen, mit welchem für den
Berichtszeitraum ein Abkommen für den Datenaustausch nach GIR besteht.
Um die Verarbeitung von Korrekturmeldungen zu erleichtern, wurde ausserdem die
zusätzliche Regel eingeführt, dass die Korrektur eines Berichtselements die gleichen
RecJurCodes enthalten muss wie die Originalmeldung. Falls der RecJurCode selbst
korrigiert werden soll, muss daher zunächst eine Löschmeldung für das betreffende
Berichtselement übermittelt und anschliessend das korrigierte Element als Neumeldung neu
eingereicht werden (vgl. Kapitel 6.4.3).
Regel Validierung Fehlercode
Es können nur GIR-Berichte für Empfänger Muss einem im 60018
eingereicht werden, welche im Berichtszeitraum gültigen
Berichtszeitraum als Partnerstaaten der Partnerstaat entsprechen
Schweiz registriert sind.
Bei einer Korrekturmeldung muss der Muss dem RecJurCode der zu 98200
RecJurCode gleich dem RecJurCode der Originalmeldung entsprechen
zu korrigierenden Originalmeldung sein
Die GeneralSection muss an alle In der GeneralSection müssen 98201
Partnerstaaten übermittelt werden, denen sämtliche RecJurCodes
18

<!-- Page 19 -->

Regel Validierung Fehlercode
mindestens eine der anderen Sections vorkommen, die in Summary,
gesendet wird. JurisdictionSection und
UTPRAttribution vorkommen
Abschnittsübergreifende Validierungsregeln
Neben der oben erwähnten Regel 98201 gibt es noch eine ganze Reihe von Regeln, welche
Werte aus mehreren Berichtselementen vergleichen. Beispielsweise werden Werte im
Summary mit Werten aus den JurisdictionSections oder der UTPR-Attribution verglichen. Da
die einzelnen Berichtselemente jedoch unterschiedliche Empfänger haben können, werden
diese Regeln nur innerhalb der Berichtselemente mit dem gleichen RecJurCode geprüft, da
die Regeln nach der Ländersortierung der Berichte immer noch erfüllt sein müssen.
Dies betrifft die folgenden Regeln:
• 60022
• 70008
• 70036
• 70037
• 70040
• 70041
• 70045
• 70047
• 70048
• 70049
• 70050
• 70051
• 70052
• 70053
• 70099
• 70100
Zu beachten ist dies nicht nur beim Erstellen einer Meldung, sondern auch bei Korrekturen.
Wird ein Berichtselement korrigiert, kann es aufgrund der Abhängigkeiten nötig sein, andere
Berichtselemente ebenfalls anzupassen.
5.3.5 DocSpec
Jedes Berichtselement im GIR-XML-Schema muss das Element DocSpec enthalten, das die
Metadaten zum übermittelten Element enthält.
Der Begriff „Berichtselement“ wird hier als Oberbegriff für die Elemente FilingInfo,
GeneralSection, Summary, JurisdictionSection und UTPRAttribution benutzt.
Der Unterschied zwischen FilingInfo und den übrigen Berichtselementen besteht darin, dass
in einer Korrektur- oder bei erneuter Übermittlung einer Neumeldung das Element FilingInfo
erneut gesendet werden muss, die anderen Berichtselementen hingegen dürfen nie als
Resend (DocTypeIndic „OECD0“) geschickt werden.
DocTypeIndic
Mit dem DocTypeIndic wird angezeigt, ob es sich um ein neues Berichtselement oder um
eine Korrektur- oder Stornomeldung handelt. Neumeldungen und Korrekturen/Stornos dürfen
in einer Meldung nicht gemischt werden. Wird das Element FilingInfo erneut übermittelt,
ohne geändert zu werden, soll gemäss OECD-Vorgaben „Resent Data“ ("OECD0“)
verwendet werden.
19

<!-- Page 20 -->

Die Resend-Option darf in den folgenden Fällen verwendet werden:
• Neue Daten: Falls neue, zusätzliche Berichtselemente gesendet werden sollen,
nachdem bereits eine Meldung für die Berichtsperiode gesendet wurde
• Korrektur/Storno: Falls Berichtselemente korrigiert oder storniert werden, wobei das
ReportingEntity-Element nicht korrigiert werden muss
Es gilt darauf hinzuweisen, dass das FilingInfo-Element nicht gelöscht werden kann, ohne
alle zugehörigen Berichtselemente zu löschen.
Regel Validierung Fehlercode
Eine produktive Meldung (Dateiname Wert = „OECD10“, „OECD11“, 50009
beginnt nicht mit „Test“) darf keine Test- „OECD12“ oder „OECD13“ und
DocTypeIndics enthalten. Dateiname ≠ „Test*.zip“
FilingInfo darf nur gelöscht werden, wenn Wert = „OECD3“ nur, wenn alle 60010
zuvor oder gleichzeitig alle zugehörigen Berichtselemente gelöscht sind.
Berichts-Elemente gelöscht werden.
DocRefId
Die DocRefId ist der Unique Identifier eines Berichtselements. Kein anderes Berichtselement
darf die gleiche DocRefId nochmals enthalten, weder in dieser noch in irgendeiner anderen
Meldung, auch nicht von einer anderen ergänzungssteuerpflichtigen Geschäftseinheit. Dies
wird über die eindeutige ID sichergestellt. Die einzige Ausnahme ist das erneute Senden des
Elements „ReportingEntity“ in einer Korrekturmeldung.
Die DocRefId ist wie folgt zusammenzusetzen:
Ländercode des Senderstaates & Berichtssteuerperiode & Eindeutige ID
Für die Übermittlung zwischen ergänzungssteuerpflichtigen Geschäftseinheiten und ESTV
muss der Ländercode des Senderstaates „CH“ sein.
Für die DocRefId gelten die folgenden Einschränkungen:
• Die DocRefId darf maximal 200 Zeichen lang sein
• Der Ländercode muss in Grossbuchstaben geschrieben werden
• Zulässig sind alle Zeichen gemäss Ziffer 7.1.
Als regulärer Ausdruck: CH[0-9]{4}.{1,194}
20

<!-- Page 21 -->

Regel Validierung Fehlercode
Darf nicht gleich einer anderen DocRefId in Wert ≠ frühere DocRefId, falls 60007
dieser oder einer früher erhaltenen DocTypeIndic ≠ „OECD0“
Meldung sein, ausser wenn die
ReportingEntity erneut gesendet wird.
Die Struktur der DocRefId muss dem Wert =„CH“ & 60011
vorgegebenen Schema entsprechen. Die Berichtssteuerperiode & 1-194
Berichtssteuerperiode muss dabei dem Ziffern, Buchstaben,
Wert aus der MessageRefId entsprechen. Bindestriche, Unterstriche oder
Punkte
Wird ein Berichtselement erneut gesendet Wenn DocTypeIndic = 60014
(„Resent Data“), muss das letzte zuvor "OECD0", muss die gleiche
übermittelte Berichtselement die gleiche DocRefId in einer früheren
DocRefId haben und das zuvor übermittelte Meldung vorhanden sein und
Berichtselement darf nicht gelöscht oder sie darf noch nicht korrigiert
korrigiert worden sein (nur bei FilingInfo). oder gelöscht worden sein (darf
keiner früheren CorrDocRefId
entsprechen)
CorrDocRefId
Jedes Berichtselement kann auch korrigiert werden, in diesem Fall muss eine CorrDocRefId
angegeben werden, welche auf die DocRefId des zu korrigierenden Berichtselements
verweist. Dabei müssen beide die gleichen Berichtselemente sein, eine JurisdictionSection
beispielsweise kann nur durch eine neue JurisdictionSection ersetzt werden.
Jedes Berichtselement darf nur einmal korrigiert werden, eine zweite Korrektur muss auf die
letzte Korrektur verweisen, nicht auf das initiale Berichtselement. Damit darf also auch jede
CorrDocRefId nur ein einziges Mal verwendet werden.
Eine CorrDocRefId darf nur bei Korrektur- und Storno-Berichtselementen (DocTypeIndic
"OECD2"/"OECD3 ") angegeben werden. Bei neuen Berichtselementen ("OECD1") oder
einem Resend einer FilingInfo ("OECD0") darf keine CorrDocRefId vorhanden sein.
Regel Validierung Fehlercode
Das referenzierte Berichtselement muss Wenn DocTypeIndic = 60005
vom gleichen Typ (FilingInfo, General "OECD2" oder "OECD3", muss
section, Summary, JuridictionSection oder das Berichtselement der
UTPRAttribution) sein wie die Korrektur gleiche Typ sein wie der in der
CorrDocRefId referenzierte
Die gleiche CorrDocRefId darf nicht Wert ≠ andere CorrDocRefId in 60006
mehrfach in der gleichen Meldung der gleichen Meldung
verwendet werden.
Die CorrDocRefId muss der DocRefId einer Wert = DocRefId in einer 60008
früheren, akzeptierten Meldung früheren Meldung
entsprechen
Das korrigierte Berichtselement darf nicht Wert ≠ frühere CorrDocRefId 60009
bereits früher korrigiert oder gelöscht
worden sein.
Eine Neumeldung oder Resend darf keine Falls Wert = „OECD0“ oder 60012
CorrDocRefId enthalten (vgl. Regel bei „OECD1“, darf DocSpec keine
CorrDocRefId). CorrDocRefId enthalten
Eine Korrektur- oder Stornomeldung muss Falls Wert = „OECD2“ oder 60015
eine CorrDocRefId enthalten (vgl. Regel bei „OECD3“, muss CorrDocRefId
CorrDocRefId). ausgefüllt sein
21

<!-- Page 22 -->

6. Meldesequenzen (Storno / Korrekturen)
6.1 Neumeldungen
Eine Neumeldung ist der Normalfall, d.h. ein Bericht wird erstmalig übermittelt. Jede
Neumeldung darf dabei nur Berichtselemente enthalten, die zuvor noch nicht übermittelt
wurden. In einer Neumeldung dürfen also keine Korrekturen oder Stornos (DocTypeIndic
„OECD2“ oder „OECD3“) vorkommen.
6.2 Stornierung ganzer Meldungen
Meldungen als Ganzes können nicht storniert werden. Um eine Meldung komplett zu
stornieren, muss eine Korrekturmeldung übermittelt werden, die sämtliche Berichtselemente
der ursprünglichen Meldung storniert.
In der Praxis dürfte es allerdings nicht nötig sein, komplette Meldungen zu stornieren. Fehler
in einzelnen Berichtselementen lassen sich über die im Folgenden beschriebenen
Korrekturmechanismen einfacher beheben. Ausserdem geht bei einer Stornierung und
anschliessenden Neumeldung der Bezug zu den vorherigen Daten verloren, daher ist einer
Korrektur immer der Vorzug zu geben.
6.3 Korrekturmeldungen
6.3.1 Grundsätze
Im GIR gibt es fünf korrigierbare Elemente: FilingInfo, GeneralSection, Summary,
JurisdictionSection und UTPRAttribution.
Ein Element kann nur als Ganzes ersetzt werden, selbst wenn nur ein Teilelement korrigiert
werden soll. Auch wenn also beispielsweise nur eine einzige Angabe im Summary-Element
korrigiert werden soll, muss das komplette Element neu übermittelt werden. Das neu
übermittelte Element ersetzt das vorherige vollständig.
Falls Elemente der ursprünglichen Meldung nicht geändert wurden, müssen sie in einer
Korrekturmeldung nicht erneut übermittelt werden. Es reicht, wenn die Korrekturmeldung die
geänderten Elemente enthält. Das FilingInfo-Element muss jedoch in jeder Korrekturmeldung
mitgeliefert werden. Es wird dazu als „Resent Data“ markiert und mit der gleichen DocRefID
erneut übermittelt.
6.3.2 Aufbau einer Korrekturmeldung
Eine Korrekturmeldung ist grundsätzlich gleich aufgebaut wie eine Neumeldung. Sie besteht
aus den Elementen FilingInfo, GeneralSection, Summary, JurisdictionSection und
UTPRAttribution, wobei nicht alle Elemente verwendet werden müssen.
Das Element MessageTypeIndic im MessageSpec einer Korrekturmeldung muss den Wert
„GIR102“ enthalten (GIR102 = „The message contains corrections for previously sent
information“).
Eine Korrekturmeldung muss ebenso wie eine Neumeldung eine eindeutige MessageRefId
enthalten. Keinesfalls darf eine MessageRefId einer früheren Meldung wiederverwendet
werden, auch nicht diejenige der zu korrigierenden Meldung.
Eine Korrekturmeldung darf keine neuen Berichtselemente enthalten, sondern nur Korrekturen
und Stornos. Der DocTypeIndic jedes Berichtselements in der Korrekturmeldung muss also
den Wert „OECD2“ für Korrektur oder „OECD3“ für Storno enthalten.
22

<!-- Page 23 -->

Jedes Korrektur- oder Storno-Berichtselement muss eine neue DocRefId enthalten. Es darf
auch hier keine bereits früher verwendete DocRefId wiederverwendet werden, auch nicht
diejenige der zu korrigierenden Meldung.
6.3.3 Korrekturketten
Die Verbindung zwischen einer Korrektur und dem zu korrigierenden Berichtselement wird
über das Element CorrDocRefId hergestellt. Die CorrDocRefId verweist auf ein bestehendes
Berichtselement, der korrigiert werden soll, muss also der DocRefId eines früheren
Berichtselements entsprechen.
Dabei ist zu beachten, dass ein Berichtselement nicht mehrfach korrigiert werden darf. Jede
CorrDocRefId darf daher ebenso wie die DocRefId nur einmal übermittelt und nicht
wiederverwendet werden.
Falls ein Berichtselement nach der Korrektur immer noch nicht korrekt ist, ist es hingegen
erlaubt, die Korrektur erneut zu korrigieren. Es kann dann eine Korrektur erstellt werden,
deren CorrDocRefId auf die DocRefId der vorherigen Korrektur verweist. Auf diesem Weg
entsteht eine Korrekturkette, bei der immer nur das letzte Glied gültig ist.
Wird ein Berichtselement hingegen storniert, endet die Kette. Ein storniertes Berichtselement
kann über eine weitere Korrektur nicht wieder hinzugefügt werden. Um ein fälschlicherweise
storniertes Berichtselement erneut zu melden, muss es wieder als neues Berichtselement in
einer Neumeldung geschickt werden.
Abbildung 1: Korrekturkette
Die Abbildung zeigt eine Korrekturkette am Beispiel einer GeneralSection, nach dem
gleichen Muster können auch die anderen Berichtselemente korrigiert werden.
6.4 Beispiele
Die nachfolgenden Beispiele dienen der Illustration und Konkretisierung des
Korrekturprozesses.
In den Beispielen sind die korrigierten und die zu korrigierenden Elemente jeweils rot
dargestellt. Das Element ReportingEntity wird grün markiert, wenn es unverändert erneut
gesendet wird.
6.4.1 Korrektur eines Berichtselements
Das erste Beispiel stellt den Fall dar, dass ein Konzern eine Neumeldung mit einer
GeneralSection und einem Summary übermittelt hat. Zuerst wird ein Element der
GeneralSection korrigiert. Anschliessend wird eine zweite Korrektur der gleichen
GeneralSection vorgenommen.
23

<!-- Page 24 -->

Die CorrDocRefId der GeneralSection verweist immer auf die direkt vorangehende Meldung,
nicht auf die initiale Meldung. Der DocTypeIndic der GeneralSection wechselt von „OECD1“
in der initialen Meldung zu „OECD2“ in der Korrekturmeldung.
Das Element FilingInfo muss auch in der Korrekturmeldung immer mitgeschickt werden,
selbst wenn es nicht verändert wird. Der DocTypeIndic wird dann auf „OECD0“ gesetzt und
die DocRefId bleibt unverändert.
In der Korrekturmeldung wird nur die veränderte GeneralSection geschickt. Unkorrigierte
Elemente wie das Summary oder allfällige weitere Berichtselemente (im Beispiel nicht
dargestellt) sind in der Korrekturmeldung nicht zu wiederholen.
Abbildung 2: Zweimalige Korrektur einer GeneralSection
6.4.2 Hinzufügen von Berichtselementen zu einer bestehenden Meldung
Wurden in der ersten Neumeldung nicht alle Berichtselemente eines Konzerns übermittelt,
können diese in weiteren Meldungen ergänzt werden. Auf diesem Weg kann ein Konzern
seine Datenlieferung auf mehrere Meldungen aufteilen oder fehlende Berichtselemente
nachliefern.
Die zweite und jede weitere Meldung sind genau wie die erste Meldung Neumeldungen. Der
MessageTypeIndic ist also „GIR101“, der DocTypeIndic aller Berichtselemente muss
„OECD1“ sein.
Das Element FilingInfo muss wiederverwendet und mit dem DocTypeIndic "OECD0“ erneut
gesendet werden. Die DocRefId der FilingInfo der zweiten Meldung muss identisch mit der
DocRefId in der ersten Meldung sein.
24

<!-- Page 25 -->

Abbildung 3: Hinzufügen von Berichtselementen für eine bestehende ergänzungssteuerpflichtige
Geschäftseinheit
6.4.3 Stornierung eines Berichtselements und anschliessende Neuübermittlung
Soll ein Berichtselement gelöscht werden, muss eine Korrekturmeldung erstellt werden
(MessageTypeIndic="GIR102"), die eine Stornomeldung für das betreffende Berichtselement
enthält. DocTypeIndic des Berichtselements ist dann „OECD3“.
Auch im Storno-Berichtselement sind die Musselemente auszufüllen. Daher ist es am
einfachsten, das ursprüngliche Berichtselement nochmals zu senden und nur den
DocTypeIndic auf „OECD3“ zu ändern.
Wichtig ist, dass auch das stornierte Berichtselement die gleichen RecJurCodes enthält, wie
in der vorangegangenen Meldung (vgl. Validierungsregel 98200 in Kapitel 5.3.4). Damit wird
sichergestellt, dass die Stornierung an dieselben Partnerstaaten übermittelt wird, welche die
vorangegangene Meldung erhalten haben.
Wird ein Berichtselement storniert, kann diese Stornierung nicht zurückgenommen oder
korrigiert werden, da die Korrekturkette mit einer Stornierung („OECD3“) beendet wird (vgl.
Ziffer 6.3.3). Um es doch wieder hinzuzufügen, muss es als neues Berichtselement mit einer
neuen DocRefId und ohne CorrDocRefId übermittelt werden. Dabei kann dann ein neuer
RecJurCode übermittelt werden, um das Berichtselement neu einem anderen Partnerstaat
zu übermitteln.
Im Beispiel wird zunächst eine Neumeldung mit einer GeneralSection und einem Summary
gesendet. Das Summary wird anschliessend storniert. Um es dann (allenfalls in korrigierter
Form) doch zu übermitteln, wird es mit DocTypeIndic „OECD1“ erneut gesendet.
Die DocRefId des ursprünglichen Summarys („SU1“) kann dabei nicht wiederverwendet
werden, es muss für die erneute Übermittlung eine neue DocRefId („SU3“) benutzt werden,
selbst wenn es sich inhaltlich wieder um die gleichen Daten handeln sollte.
25

<!-- Page 26 -->

Abbildung 4: Stornierung und anschliessende Neumeldung
Zu beachten ist, dass manche Berichtselemente aufgrund gegenseitiger Abhängigkeiten in
den Validierungsregeln allenfalls nicht einzeln gelöscht werden können. Soll beispielsweise
die GeneralSection gelöscht werden, muss zwangsläufig der gesamte vorherige Bericht
gelöscht werden, da der Zustand ohne GeneralSection gegen verschiedene Regeln
verstösst. Auch im Fall des gelöschten Summarys kann es nötig sein, zumindest die
JurisdictionSections und UTPRAttribution mit dem gleichen RecJurCode ebenfalls zu
löschen.
Korrektur der FilingInfo
Auch das Element FilingInfo kann korrigiert werden, falls die Daten des Konzerns selbst
Fehler enthalten. In diesem Fall reicht es, eine Meldung mit der korrigierten FilingInfo zu
übermitteln, weitere Berichtselemente können weggelassen werden.
Die korrigierte FilingInfo ersetzt das vorhergehende Berichtselement in der Korrekturkette,
sodass alle nachfolgenden Neu- oder Korrekturmeldungen nicht mehr die ursprüngliche,
sondern die neue, korrigierte FilingInfo enthalten müssen, selbst wenn Berichtselemente aus
der ersten Meldung korrigiert werden sollen.
26

<!-- Page 27 -->

Abbildung 5: Korrektur der FilingInfo und anschliessende weitere Korrektur
Ebenso erlaubt wäre es in diesem Fall, die Korrektur der FilingInfo sowie von anderen
Berichtselementen in einer Meldung zu übermitteln. Es ist jedoch nicht erlaubt, in der
gleichen Meldung auch neue Berichtselemente zu senden.
27

<!-- Page 28 -->

7. Anhang
7.1 Zulässiger Zeichensatz
Die Datenelemente in einer GIR-XML-Datei dürfen nur Zeichen aus der ISO 8859-1
Codepage mit Ausnahme der folgenden Zeichen enthalten:
Zeichen Beschreibung UTF-8 ISO 8859-1
Code Code
! Ausrufezeichen U+0021 0x21
" Anführungszeichen U+0022 0x22
# Doppelkreuz U+0023 0x23
$ Dollarzeichen U+0024 0x24
< Kleiner-als-Zeichen U+003C 0x3C
> Größer-als-Zeichen U+003E 0x3E
^ Zirkumflex U+005E 0x5E
~ Tilde U+007E 0x7E
£ Pfundzeichen U+00A3 0xA3
¤ Allg. Währungssymbol U+00A4 0xA4
¥ Yen-Zeichen U+00A5 0xA5
¦ Unterbrochener Strich U+00A6 0xA6
§ Paragraphenzeichen U+00A7 0xA7
¨ Trema U+00A8 0xA8
© Copyrightzeichen U+00A9 0xA9
ª Feminines Ordinalzeichen U+00AA 0xAA
« Nach links zeigendes doppeltes spitzes Anführungszeichen U+00AB 0xAB
¬ Nicht-Zeichen U+00AC 0xAC
Weiches Trennzeichen U+00AD 0xAD
® Zeichen für ein registriertes Warenzeichen U+00AE 0xAE
¯ Makron U+00AF 0xAF
° Gradzeichen U+00B0 0xB0
± Plusminuszeichen U+00B1 0xB1
² Hochgestellte Zwei U+00B2 0xB2
³ Hochgestellte Drei U+00B3 0xB3
´ Akut U+00B4 0xB4
μ Mikro-Zeichen U+00B5 0xB5
· Mittelpunkt U+00B7 0xB7
¸ Cedille U+00B8 0xB8
¹ Hochgestellte Eins U+00B9 0xB9
º Maskulines Ordinalzeichen U+00BA 0xBA
» Nach rechts zeigendes doppeltes spitzes U+00BB 0xBB
Anführungszeichen
¼ Bruch ein Viertel U+00BC 0xBC
½ Bruch ein Halb U+00BD 0xBD
¾ Bruch drei Viertel U+00BE 0xBE
¿ Umgekehrtes Fragezeichen U+00BF 0xBF
÷ Divisionszeichen U+00F7 0xF7
28

<!-- Page 29 -->

Zudem sind die folgenden Zeichenfolgen nicht erlaubt:
Zeichen Beschreibung UTF-8 Code ISO 8859-1 Code
-- Minuszeichen Minuszeichen U+002DU+002D 0x2D0x2D
/* Bruchstrichzeichen Sternzeichen U+002FU+002A 0x2F0x2A
&# Kaufmännisches Und Doppelkreuz U+0026U+0023 0x260x23
29