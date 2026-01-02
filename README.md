# CTX-Wartungs-Tools

Modulares PowerShell-Wartungstool fuer Citrix/Multiuser-Umgebungen.

## Struktur (Kurzfassung)
- /host: optionaler Host (spaeter GUI)
- /tools/<tool>: autarke Sub-Tools mit eigenem SDK-Vendor, Config und Runner
- /shared/SDK: Source-of-truth fuer das SDK
- /shared/build/Sync-SDK.ps1: verteilt SDK in alle Tools

## Quickstart (Tool lokal testen)
1. Beispiel-Configs kopieren:
   - tools/benutzerprofil-wartung/customer.json.example -> customer.json
   - tools/benutzerprofil-wartung/policy.json.example -> policy.json
2. Aktionen und UI-Stubs anpassen.
3. Start:
   - User-UI: tools/benutzerprofil-wartung/Start.ps1
   - Runner (silent): tools/benutzerprofil-wartung/Runners/Runner.ps1 -Trigger Logon

## Bestehende Skripte importieren
- Ersetze die Stubs in `tools/<tool>/Actions/*.ps1` durch deine vorhandenen Aktionen.
- Ersetze bzw. erweitere die UI in `tools/<tool>/UI/*.ps1`.
- Passe `tools/<tool>/tool.json`, `customer.json` und `policy.json` an.
- Lasse `shared/SDK` als Source-of-truth bestehen und fuehre danach `shared/build/Sync-SDK.ps1` aus, damit die vendor copy im Tool aktualisiert wird.
