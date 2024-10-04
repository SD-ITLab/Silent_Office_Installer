# Silent_Office_Installer | Unattended Script

Version: 2.0.0

## Beschreibung

Der Silent Office Installer ist ein unbeaufsichtigtes Skript, das die Installation von Microsoft Office Pakete
automatisiert.
Es bietet eine menügesteuerte Oberfläche, die es Ihnen ermöglicht, die zu installierende Software ohne jegliche
Benutzerinteraktion auszuwählen.
Das Skript unterstützt die Installation der folgenden Pakete:

1. Microsoft Office 2021 Pro Plus
2. Microsoft Office 2024 Home and Student
3. Microsoft Office 2024 Home and Business
4. Microsoft Office 2024 Pro Plus

Das Skript prüft, ob die jeweilige Software bereits auf dem System installiert ist, und überspringt die Installation, wenn sie erkannt wird.
Andernfalls sucht es im angegebenen Verzeichnis nach den Installationsdateien und installiert die Software unbemerkt.

## Verwendung
1. Führen Sie die Silent_Office_Installer.exe aus.
2. Das Skript zeigt ein Menü mit Optionen für verschiedene Installationskonfigurationen an.
3. Wählen Sie die gewünschte Option aus in dem Sie den entsprechend Button betätigen.
4. Das Skript installiert die ausgewählte Softwarepaket im Hintergrund, ohne dass der Benutzer eingreifen muss.
5. Sobald der Installationsvorgang abgeschlossen ist, wird eine Bestätigungsmeldung angezeigt.

## Hinweise
- Das Skript geht davon aus, dass sich die Grund-Installationsdateien für im angegebenen Verzeichnis (`.\_Office-Installer\`) befinden.
- Wenn eine Installationsdatei nicht gefunden wird, wird eine Fehlermeldung angezeigt.
- Das Skript prüft anhand der Registrierungsinformationen, ob die jeweilige Software bereits installiert ist."
