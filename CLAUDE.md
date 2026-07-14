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
- **CRMFirms**: Title, Adresse, PLZ, Ort, Land, Hauptnummer, Klassifizierung (A/B/C), VIP,
  Kategorie (Choice: Kunde/Lieferant/Übrige; internes Feld `Kategorie` -> `firm.kategorie`)
- **CRMContacts**: Title(=Nachname), Vorname, Anrede, Firma(Lookup), Funktion, Email1/2, Direktwahl, Mobile, Rolle, Leadbbz0, SGF, Geburtstag, Kommentar, Event, Eventhistory, Archiviert
- **CRMHistory**: Title, Nachname(Lookup), Datum, Kontaktart(=typ), Notizen, Projektbezug, Leadbbz
- **CRMTasks**: Title, Name(Lookup), Deadline, Status, Leadbbz

`firmSignal(firm)` -> `overdue | never | cold | ok | ""`. **Gate:** nur Firmen mit
`kategorie === "Kunde"` erhalten ein Signal/Dot; Lieferant/Übrige/leer -> `""` (kein Dot).
Klassifizierung (A/B/C) spielt **keine** Rolle. Stufen: `overdue` (offene überfällige Task) >
`never` (keine History) > `cold` (letzte Aktivität > 12 Monate, exakte Monatsdifferenz) >
`ok` (on track). Basis für die Signal-Dots im **Firmen-Cockpit** (grüner `ok`-Dot bleibt,
nur bei Kunden); die `history`-Route nutzt kein `firmSignal`. Kein Pflege-Radar mehr —
die Cockpit-Tabelle führt Deadline+Aktivität in **einer** Spalte „Status/Aktivität"
zusammen (Precedence überfällige Task > laufende Task > Aktivitätsalter), sortierbar nach Dringlichkeit.

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
8. **Klickbare Aktivitäts-Charts (`history`):** Aktiv/selektiert = `var(--amber)`
   (`.bbz-actbar-fill.bbz-on`), inaktiv `--blue-mid`, Hover **nur** auf `:not(.bbz-on)`
   (`--blue`). **Keine `background`-Transition** — der Aktivton erscheint sofort beim
   Klick (kein Warten auf `mouseleave`). `fensterTage` (aus `granularitaet`+`periode`)
   ist die gemeinsame Fensterlogik für Band-Kacheln UND beide Charts.
9. **Doku im selben Commit:** Jede Verhaltensänderung an einer Route aktualisiert
   den entsprechenden Abschnitt in dieser CLAUDE.md im **selben Commit**.

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

**Genau EIN Pages-Workflow:** nur `deploy.yml`. Keinen zweiten (z.B. `static.yml`)
anlegen — zwei Workflows laden dasselbe `github-pages`-Artefakt hoch -> Kollision.
**"Deployment failed / try again later"** ist meist ein transienter GitHub-Pages-
Backend-Fehler (githubstatus.com), **kein** Code-Fix: Deploy neu auslösen mit
`gh workflow run deploy.yml --ref main` (nicht `gh run rerun --failed` — das dupliziert Artefakte).

## Aktivitäten-Route (`history`) — Ist-Stand
`views.historyView()` (app.js ~Z. 3918). **Reiner Aktivitäts-Bericht** über tatsächliche
Kontakte/Aktivitäten — **kein** Pflege-Radar, **kein** Handlungszentrum, **keine**
Tasks/Ziele (Tasks leben in `planning`, Pflege-Signale in `firms`).

**Zwei Ansichten** via `filters.history.viewMode`:
- `"firms"` (Default): **Firmen-Bericht** — nur Firmen mit ≥1 Aktivität im aktiven
  Zeitfenster, sortiert nach letzter Aktivität (neueste zuerst), aufklappbar
  (`toggle-firm-expand`) -> `renderCard` der Firmen-Aktivitäten. Schnellerfassung
  (Buttons je Kontaktart) über dem Inhalt.
- `"timeline"`: chronologische, datumsgruppierte Timeline (`renderCard`) + Kontaktart-`<select>`.

**Steuerung (visuell/klickbar, keine Zeit-/Lead-Dropdowns):** Slim-Actionbar
(Ansicht-Umschalter `Firmen|Chronologisch` + Suche + `+ Aktivität`) plus zwei
klickbare Balkengrafiken als Filter (Inline-SVG/Divs, keine Lib):
- **Lead BBZ** (horizontal): Count je `leadbbz0`, case-insensitiv dedupliziert
  (kanonisch = häufigste Variante) -> `filters.history.lens` (`data-action="filter-lens"`, Toggle).
- **Zeitraum** (vertikal): Granularität `filters.history.granularitaet`
  (`"monat"|"quartal"|"jahr"`, Umschalter `filter-granularitaet`) mit rollierender
  Spanne (12 Monate / 8 Quartale / max. 3 Jahre mit Daten). Balkenklick ->
  `filters.history.periode` (`filter-periode`, Toggle); Reset-Chip zeigt die aktive Spanne.

**Fensterlogik:** `activeWindow` = Aktivitäten der Spanne ∩ `periode`; der Lead-Chart
aggregiert über dasselbe `activeWindow`. `scoped = activeWindow ∩ lens` speist Band,
Firmen-Bericht und Timeline. `fensterTage` (aus `granularitaet`+`periode`) ist die
gemeinsame Fensterlänge für Band-Kacheln UND Ø/Woche.

**Instrumenten-Band** = reine Anzeige, **3 Kacheln**: Aktivitäten total /
Aktive Firmen (≥1 Aktivität) / **Ø pro Woche** = `total / Math.max(1, fensterTage/7)`
(kein fixes `/52`). Zentrale Helper `helpers.periodKey(date, gran)` /
`helpers.periodLabel(key)` überall wiederverwenden.

`state.filters.history` (real existierende Felder): `search, kontaktart, viewMode,
lens, granularitaet, periode, expandedFirms`. (Entfernt: `radarMode`, `monat`,
`wochenziel`, `zeitfenster`, `leadbbz`, `groupBy`.)
