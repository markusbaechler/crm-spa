# CLAUDE.md — bbz CRM

Projektgedächtnis für Claude Code. Lies das vor jeder Änderung.

## Was das ist
Framework-freie Vanilla-JS-PWA. CRM für bbz Beratung: Firmen, Kontakte,
Aktivitäten, Events. Datenbackend ist SharePoint (MS Graph v1.0, MSAL-Auth).
Gehostet auf GitHub Pages: https://markusbaechler.github.io/crm-spa/

**Kein Build-Step, kein Paketmanager, kein Bundler.** Die drei Dateien werden
1:1 ausgeliefert. Kein `npm install`, kein Transpiling.

## Dateikarte
| Datei | Inhalt |
|---|---|
| `index.html` | App-Shell + gesamtes CSS (`:root`-Design-Tokens ab Z. ~22) + Nav |
| `app.js` | Gesamte Logik (~5800 Z.): CONFIG, SCHEMA, state, helpers, dataModel, views, controller |
| `io.js` | Import/Export via SheetJS (`window.bbzIO`) |
| `service-worker.js` | Network-first für Navigationen, cacht nur Offline-URL |
| `manifest.json` | PWA-Manifest, scope `/crm-spa/` |

## Architektur (app.js)
- `CONFIG` (Z. ~4): Graph tenantId/clientId, SharePoint-Site, Listennamen, defaults.
- `SCHEMA` (Z. ~37): Mapping App-Feld -> SharePoint-Feldname pro Liste.
- `state` (Z. ~104): auth, meta, data (roh), enriched (angereichert), filters, selection, modal.
- `helpers`: escapeHtml, Datums-Utils, `firmSignal`, `leadbbzBadgeHtml`, `debounce` etc.
- `dataModel.enrich()` (Z. ~1694): join firms/contacts/history/tasks, leitet ab:
  firms -> `contactsCount, contacts[], tasks[], history[], openTasksCount, nextDeadline, latestActivity`;
  history -> `contactId, contactName, firmId, firmTitle, projektbezugBool`;
  tasks -> `isOpen, isOverdue`.
- `views`: reine String-Builder pro Route. `renderRoute()` (Z. ~1812) dispatcht per `state.filters.route`.
- `controller.render()` (Z. ~5681): `ui.renderShell()` + `ui.renderView(views.renderRoute())` -> `innerHTML` in `#view-root`.

## Datenmodell (SharePoint-Listen)
- **CRMFirms**: Title, Adresse, PLZ, Ort, Land, Hauptnummer, Klassifizierung (A/B/C), VIP
- **CRMContacts**: Title(=Nachname), Vorname, Anrede, Firma(Lookup), Funktion, Email1/2, Direktwahl, Mobile, Rolle, Leadbbz0, SGF, Geburtstag, Kommentar, Event, Eventhistory, Archiviert
- **CRMHistory**: Title, Nachname(Lookup), Datum, Kontaktart(=typ), Notizen, Projektbezug, Leadbbz
- **CRMTasks**: Title, Name(Lookup), Deadline, Status, Leadbbz

`firmSignal(firm)` -> `overdue | never | cold | ok | ""` (Basis für Pflege-Radar).

## Konventionen — strikt einhalten
1. **Kein Framework/Build einführen.** Bleibt Vanilla. Keine neuen Dependencies ohne expliziten Auftrag.
2. **Rendering:** Views geben HTML-Strings zurück. Interaktion ausschliesslich über
   `data-action="..."`-Attribute + zentrale Event-Delegation.
3. **Escaping:** Jeder Nutzer-/SP-Wert im HTML MUSS durch `helpers.escapeHtml()`.
4. **Styling:** Nur bestehende CSS-Variablen (`--blue`, `--muted`, `--green`, `--amber`,
   `--red`, `--line`, `--panel`, `--r-md`, `--shadow` ...) und `bbz-*`-Klassen. Keine neuen Ad-hoc-Farben.
5. **State:** Filter/Selektion in `state.filters.<route>` bzw. `state.selection`. Nach Mutation `controller.render()`.
6. **SharePoint-Write-Eigenheit:** POST speichert zuverlässig nur `Title` + Lookups. Restfelder per
   separatem PATCH auf die neue Item-ID nachschreiben (Muster ab Z. ~5229 nicht umgehen).
7. **Deutsch (CH):** UI-Texte auf Deutsch, "ß" immer als "ss".

## Lokal ausführen
Static-Server aus dem Repo-Root: `python -m http.server 8080`, dann http://localhost:8080/ öffnen.
**Auth-Gotcha:** MSAL `redirectUri` ist auf die Prod-URL fixiert. Login gegen echtes SharePoint
funktioniert lokal nur, wenn `http://localhost:8080/` in der Azure-AD-App als Redirect-URI registriert ist.

## Deploy (autonom)
Trigger: Push auf `main`. Workflow `.github/workflows/deploy.yml`:
1. `node --check` auf app.js/io.js/service-worker.js — bricht bei Syntaxfehler ab (kein Deploy).
2. Stampt kurze Commit-SHA als Cache-Buster in Script-Tags + SW-CACHE-Namen.
3. Publiziert nach GitHub Pages.

**Regeln:** Nie mit rotem `node --check` pushen. app.js-Änderungen sind global (eine Datei = ganze App);
nach Änderung lokal laden und Zielroute klicken. Ein Commit = eine logische Änderung.

## Aktueller Sprint
Aktivitäten-Seite (`history`-Route, `views.historyView()` ab Z. ~3852) zum Cockpit umbauen.
Brief in `COCKPIT-SPEC.md`.
