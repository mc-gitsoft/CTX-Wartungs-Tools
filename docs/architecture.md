# Architektur

## Ueberblick
- Ein optionaler Host startet und verwaltet Sub-Tools.
- Jedes Sub-Tool ist autark (eigene Configs, Runner, SDK-Vendor copy).

## Konfiguration
- `customer.json` und `policy.json` liegen im Tool-Root (neben `Start.ps1`).
- Keine systemweite Ablage (kein ProgramData, keine Registry).

## State
- Once-Policies speichern ihren Status pro Benutzer unter
  `%LOCALAPPDATA%\CTX-Wartungs-Tools\State\<toolId>\`.

## Runner Trigger
- GPO-Wrapper rufen `Runner.ps1` mit `-Trigger Logon|Logoff` auf.
- Aktionen werden ausschliesslich im Silent-Mode ausgefuehrt.

## Silent vs. Interactive
- Runner und GPO-Skripte oeffnen nie GUI.
- UI-Skripte sind getrennt und nur fuer User/Admin.

## Offline Konzept
- Offline-Betrieb ist optional und wird ueber customer.json gesteuert.
