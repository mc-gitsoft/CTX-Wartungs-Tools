# Benutzerprofil-Wartung

Autarkes Sub-Tool fuer die Wartung von Benutzerprofilen. Kann direkt gestartet werden oder ueber den Host.

## Einstieg
- `Start.ps1` fuer UI (User/Admin).
- `Runners/Runner.ps1` fuer silent Ausfuehrung via GPO.

## Konfiguration
- `customer.json` und `policy.json` liegen im Tool-Root.
- Beispiel-Dateien: `customer.json.example` und `policy.json.example`.

## Standalone Start
- `Start.ps1 -UiMode User` fuer Benutzeroberflaeche.
- `Start.ps1 -UiMode Admin` fuer Policy-Editor (logon.once).

## Policy Pflege
- `UI/Wartung_Admin.ps1` schreibt `policy.json` im Tool-Root.
- `policy.json` steuert `logon.every`, `logon.once`, `logoff.every`, `logoff.once`.

## Runner / GPO
- `Runners/GPO_Logon.ps1` und `Runners/GPO_Logoff.ps1` rufen `Runner.ps1` auf.
- Aktionen werden im Runner immer im Silent-Mode ausgefuehrt.
