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
  Kategorie (Choice: Kunde/Lieferant/Übrige; SP-internes Feld `Kategorie`, gemappt in
  `normalizer.firm()` -> `firm.kategorie`)
- **CRMContacts**: Title(=Nachname), Vorname, Anrede, Firma(Lookup), Funktion, Email1/2, Direktwahl, Mobile, Rolle, Leadbbz0, SGF, Geburtstag, Kommentar, Event, Eventhistory, Archiviert
- **CRMHistory**: Title, Nachname(Lookup), Datum, Kontaktart(=typ), Notizen, Projektbezug, Leadbbz
- **CRMTasks**: Title, Name(Lookup), Deadline, Status, Leadbbz

> **Lead-Semantik (Route `aktivitaeten`):** „Lead bbz" = **Record-Lead** (`history.leadbbz` /
> `task.leadbbz`), **kein** Kontakt-Fallback auf `contact.leadbbz0`. Bewusste Abweichung von
> der alten `history`-Route, die nach `contact.leadbbz0` filterte.

`firmSignal(firm)` -> **fünf Rückgaben** `overdue | never | cold | ok | ""`. **Gate:**
nur Firmen mit `firm.kategorie === "Kunde"` erhalten ein Signal; Lieferant/Übrige/leer
-> `""` (kein Dot). Klassifizierung (A/B/C) spielt **keine** Rolle; **VIP ist ein
separates Flag ohne Gate-Einfluss**. Stufen in Priorität: `overdue` (≥1 offene,
überfällige Task) > `never` (keine History — gilt für **alle** Kunden, nicht nur A/B)
> `cold` (letzte Aktivität > 12 Monate, exakte Monatsdifferenz via getFullYear/getMonth)
> `ok` (on track, sonst). Die `history`-Route nutzt kein `firmSignal`.

**Signal-Dot-Farben** (Rendering an zwei Stellen — Desktop-Tabelle + Mobile-Card,
identisch halten): `overdue` = **rot**, `never` = **rot**, `cold` = **amber**,
`ok` = **grün**, `""` = **kein Dot** (Lieferant/Übrige). Farben ausschliesslich über
`--red`/`--amber`/`--green` (siehe `--red`-Warnung in Konvention #4).

## Konventionen — strikt einhalten
1. **Kein Framework/Build einführen.** Bleibt Vanilla. Keine neuen Dependencies ohne expliziten Auftrag.
2. **Rendering:** Views geben HTML-Strings zurück. Interaktion ausschliesslich über
   `data-action="..."`-Attribute + zentrale Event-Delegation.
3. **Escaping:** Jeder Nutzer-/SP-Wert im HTML MUSS durch `helpers.escapeHtml()`.
4. **Styling:** Nur bestehende CSS-Variablen (`--blue`, `--muted`, `--green`, `--amber`,
   `--red`, `--line`, `--panel`, `--r-md`, `--shadow` ...) und `bbz-*`-Klassen. Keine neuen Ad-hoc-Farben.
   **`--red` MUSS echtes Rot bleiben (`#a4161a`).** Ein versehentlich gesetztes Teal
   (`#0d6e6a`) liess sämtliche Danger-Elemente inkl. Signal-Dots grünlich rendern
   (Logik/Mapping waren korrekt — nur der Token-Wert falsch). Nie auf einen Nicht-Rot-Wert setzen.
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

## Firmen-Route (`firms`) — Ist-Stand
`views.firms()`. **Reines Firmenboard** — **kein** Instrumenten-Band, **keine**
KPI-Kacheln, **keine** Geburtstage-Kachel, **keine** rechte „Offene Tasks"-Liste,
**kein** Pflege-Radar (alles ersatzlos entfernt). Aufbau: Header („Firmen" + „+ Firma")
-> Filterbereich -> Suche -> aufklappbare Legende -> Tabelle.

**Zweistufiger Filter** — `state.filters.firms = { kategorie, klassifizierung, vip,
search, legendeOffen, sortBy, sortDir }`:
- Stufe 1 **Kategorie** (`kpi-filter`/`firms-kategorie`): Chips Kunden(Default) |
  Lieferanten | Übrige, filtert `firm.kategorie`, Zähler je Chip.
- Stufe 2 **Klassifizierung/VIP** — **nur bei Kunden sichtbar**: `Alle|A|B|C`
  (`firms-klassifizierung`, exklusiv) + separat `♛ VIP` (`firms-vip`, boolescher
  Toggle, additiv zu A/B/C). Beim Kategorie-Wechsel weg von „Kunde" werden
  `klassifizierung` + `vip` zurückgesetzt.

**Tabelle (6 Spalten):** Dot · Firma · Ort · Klassifizierung · Kontakte ·
Status/Aktivität. **Keine** VIP-Spalte, **keine** Tasks-Spalte — Task-Info steckt nur
im kombinierten Feld „Status/Aktivität" (Precedence überfällige Task > laufende Task >
Aktivitätsalter; farblos ausser rotem „seit X fällig"; sortierbar nach Dringlichkeit).
Dot = `firmSignal` (siehe oben), Nicht-Kunden ohne Dot.

**Aufklappbare Signal-Legende** (`data-action="toggle-firm-legende"` ->
`filters.firms.legendeOffen`): grün „Aktiv gepflegt" / amber „Aufmerksamkeit" /
rot „Nicht aktiv gepflegt" + Fusszeile „Lieferanten und Übrige tragen keinen Punkt."

**Vorgehalten fürs geplante Dashboard — bewusst NICHT gelöscht:**
`helpers.upcomingBirthdays` / `helpers.birthdayLabel` werden weiter in `firmDetail`
und der Kontakte-Route genutzt — **nicht als toten Code entfernen**. Die aus dem Board
entfernte KPI-/Geburtstags-/Task-Fälligkeits-Anzeige war **inline** (kein benannter
Helper) und ist bei Bedarf aus der Git-History wiederherstellbar.

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

## Aktivitäten-Route (`aktivitaeten`) — Ist-Stand  ⟵ zusammengeführt aus `planning` + `history`

`views.aktivitaeten()` (app.js ~Z. 3681). **Ersetzt** die Screens `planning` (Aufgaben)
und `history` (Aktivitäten) durch **einen** Screen. Aktivität (Vergangenheit) und Aufgabe
(Zukunft) verschmelzen um die **Kunden-Beziehung** herum, bleiben aber typografisch
getrennt: **blau = Aufgabe, rot = überfällig, grau = Aktivität/erledigt**. Ampelfarben
(grün/amber/rot) sind **ausschliesslich** dem Firmen-Signal vorbehalten.

**Routing/Redirect:** Nav-Buttons (Desktop + Bottom) zeigen auf `data-route="aktivitaeten"`.
`renderRoute()` lenkt `planning` und `history` **auf `aktivitaeten` um** (alte Bookmarks/
Deep-Links `#planning`/`#history` überleben; `knownRoutes` enthält alle drei). Die alten
View-Funktionen `planning()` / `historyView()` bleiben im Code (nicht gelöscht), sind aber
nicht mehr über die Nav erreichbar. Cross-Links aus dem Firmenboard
(`navigate-planning` / `navigate-planning-filtered`) zeigen jetzt auf `aktivitaeten`
(`week`/`rest` → `month`, Achse `chrono`).

**State:** `state.filters.aktivitaeten = { segment:"kunden", axis:"firm", search, lead,
faelligkeit, expandedFirms[], bucketOpen{}, moreOpen{}, legendeOffen }`.
- **segment** (`kpi-filter`/`akt-segment`, exklusiv): `kunden` (Default, = Banken/
  Versicherungen, Gate `kategorie==="Kunde"`) | `alle`. Wechsel setzt `lead`+`faelligkeit` zurück.
- **axis** (`akt-axis`): `firm` (Default) | `chrono`.
- **lead** (`kpi-filter`/`akt-lead`, Toggle, case-insensitiv): Record-Lead-Filter.
- **faelligkeit** (`kpi-filter`/`akt-faelligkeit`, Toggle): `""|overdue|month|later`.
- **expandedFirms/bucketOpen/moreOpen**: UI-Zustand der Achsen (getrennt von `history`).
- **legendeOffen** (`akt-legende`): aufklappbare Signal-Legende (nur Firma-Achse).
- Suche: `data-filter="akt-search"`.

**Monats-Raster (kein Woche-Bucket):** `mo = heute+30`, `moP = heute−30`. Aufgaben-Fenster:
`overdue` (offen & überfällig) · `month` (offen, heute…+30) · `later` (offen, >+30). Aktivitäts-
Zähler: total · 30 Tage · 365 Tage. Kontaktart-Split-Bar aus `choices[CRMHistory].Kontaktart`
(Reihenfolge SP-Choices, sonst alphabetisch).

**Zwei Achsen:**
- `firm` (Default): **Bank-Cadence-Karten** je Firma mit ≥1 Aktivität/Aufgabe, **gruppiert nach
  `firmSignal`** (NICHT nach Aufgaben-Fälligkeit): `akt-f-rot` „Nicht aktiv gepflegt"
  (overdue+never, offen) / `akt-f-amber` „Aufmerksamkeit" (cold, offen) / `akt-f-gruen`
  „Aktiv gepflegt" (ok, zu) / `akt-f-kein` „Ohne Signal" (Nicht-Kunden, zu).
  **Nicht auf Aufgaben-Buckets zurückbauen** — das verwirrte, v.a. der Sammel-Bucket
  „Ohne offene Aufgabe". Buckets klappbar (`akt-bucket`), Cap 8 + „+N weitere".
  Sortierung innerhalb: **längster Kontaktabstand zuerst**; nie kontaktierte Firmen ganz oben.
  Achtung: `helpers.compareDateAsc` sortiert fehlende Daten ans ENDE — daher eigene
  `byLastTouch`-Sortierung. Kopf zeigt Signal-Dot,
  „Letzter Touch" + „Nächste Aufgabe" (Farbe rot/amber/neutral nach Dringlichkeit) + Lead-Tag.
  Aufklappbar (`akt-firm-expand`) → gemergte Timeline (Aktivitäten+Aufgaben, neueste zuerst) +
  „+ Aktivität"/„+ Aufgabe" (reuse `open-history-form`/`open-task-form` mit `data-firm-id`).
- `chrono`: **Zweispaltig** (`.bbz-akt-split`, eigene CSS-Klasse in index.html; stapelt <900px).
  Die Spalten trennen **Objekttyp**, NICHT Zeit: **links ausschliesslich Aktivitäten**
  (`akt-p-month` offen / `akt-p-old` zu), **rechts ausschliesslich Aufgaben**
  (`akt-c-over`, `akt-c-month` offen / `akt-c-later` zu / `akt-c-done` „Erledigt" zu).
  Jede Gruppe klappbar (`akt-bucket`), **Cap 8** + „+N weitere" (`akt-more`).
  **Erledigte Aufgaben gehören NIE in den Verlauf.** Sie werden nach *Deadline* einsortiert,
  die in der Zukunft liegen kann — im „Verlauf" ergäbe das „in 2 Tagen". Sie stehen im
  Bucket `akt-c-done` (default eingeklappt = ausgeblendet, Zähler sichtbar).
  **Wichtig:** NICHT `.bbz-history-split` wiederverwenden — die blendet Spalte 2 mobil aus
  (altes Tab-Bar-Konzept, index.html Z. ~321).

**Wiederverwendete Aktionen (nicht duplizieren):** `open-firm`, `open-contact`,
`open-history-form`, `open-task-form`, `edit-task`, `edit-history`, `complete-task`,
`task-status-change`. Neue Aktionen alle mit `akt-`-Präfix.

**Zeilen-Aktionen (jede Event-Zeile, `iconBtn`-Helper):** Aktivität — Klick öffnet das
Detail-Modal, ✎ öffnet Bearbeiten (`edit-history`). Aufgabe — Klick/✎ öffnet Bearbeiten
(`edit-task`), ✓ markiert erledigt (`complete-task`, nur offene).

**Löschen NUR im Bearbeiten-Modus** (gegen versehentliches Löschen): kein ✕ in Zeilen, kein
Löschen im read-only Detail-Modal. `renderHistoryForm` hatte den Löschen-Button im
`mode === "edit"` schon; `renderTaskForm` wurde entsprechend nachgezogen (wirkt auch in
firmDetail/planning, die dieselbe Form nutzen). **Nicht wieder ✕ in die Zeilen bauen.**

**Zeilen-Darstellung:** In beiden Achsen ist die **Firma prominent** (fett, 13px) — nicht die
Kontaktart. Aktivität: Firma · Kontaktart(blau) · relatives Datum, Notiz einzeilig darunter.
Aufgabe: Firma · Fälligkeit, Titel darunter.

**Aktivitäts-Detail-Modal** (`state.modal.type === "history-detail"`, `views.renderHistoryDetail`,
`controller.openHistoryDetail(id)`, Aktion `open-history-detail`): read-only Vollansicht mit
**ungekürzten Notizen**, Firma/Kontakt verlinkt. Footer: Schliessen / Bearbeiten
(→ öffnet `history`-Formular; dort erst ist Löschen möglich). Klick auf eine Aktivitäts-Zeile öffnet dieses Modal (nicht die
Firma). Schliessen via `[data-close-modal]` oder Backdrop.

**Lead bbz** ist eine **dezente Chip-Zeile** (`bbz-kpi-chip`, Label „Filter · Lead bbz",
aktiver Chip + „✕ Filter aufheben") — bewusst KEIN Balkendiagramm mehr (war optisch zu
dominant und nicht als Filter erkennbar).

**Styling:** nur bestehende Tokens/`bbz-*`-Klassen; Event-Zeilen als Inline-Styles mit
CSS-Variablen (Muster wie `historyView`). Neue CSS-Klassen in `index.html` nur wenn Media-Queries nötig sind (z. B. `.bbz-akt-split`).
