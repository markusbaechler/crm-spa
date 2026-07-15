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

**Zwei Achsen** (`akt-axis`), **Default `chrono`**:
- `chrono` **„Agenda" = Hauptansicht.** Zweispaltig (`.bbz-akt-split`, stapelt <900px). Die
  Spalten trennen **Objekttyp**, nicht Zeit: **links nur Aktivitäten** (`akt-p-week` /
  `akt-p-month` offen, `akt-p-old` „Früher" zu), **rechts nur Aufgaben** (`akt-c-over`,
  `akt-c-month` offen; `akt-c-later`, `akt-c-done` „Erledigt" zu).
  **Erledigte Aufgaben gehören NIE in den Verlauf** — sie werden nach *Deadline* sortiert, die
  in der Zukunft liegen kann („in 2 Tagen" im Verlauf = Unsinn).
- `firm` **„Firmencockpit".** **Signal-FILTER statt Rubriken** (`akt-sig`, `F.sig`):
  `gruen` „Aktiv gepflegt" (Default) / `amber` „Beobachten" / `rot` „Brauchen Pflege" /
  `kein` „Ohne Signal" (nur bei Segment `alle`). **Immer genau EINE Kategorie sichtbar** —
  Zweck: keine Kategorie darf die andere erschlagen. **Nicht auf gestapelte Signal-Buckets
  zurückbauen.** Darin Gliederung nach letztem Kontakt: `akt-f-wk` „Diese Woche" / `akt-f-mon`
  „Diesen Monat" offen, `akt-f-alt` „Übrige" zu — **gleiche Richtung wie die Agenda (neu→alt)**;
  Drift fangen die Kategorien `amber`/`rot` ab, nicht die Sortierung.
  Kacheln im `.bbz-akt-fgrid` (auto-fill, 3/2/1 Spalten), aufgeklappte Kachel spannt voll
  (`.is-open`) und zeigt Aktivitäten|Aufgaben der Firma (`.bbz-akt-fsplit`).
  **Der Signal-Punkt entfällt in der Kachel** — im gefilterten Cockpit trägt er keine
  Information. **Kein Platzhalter „keine offene Aufgabe"** — der Text war reines Rauschen;
  fehlt die nächste Aufgabe, bleibt die Stelle leer. Nicht wieder einbauen. `firmRows` umfasst **alle** Segment-Firmen, auch nie kontaktierte (Signal
  `never` → `rot`); nicht auf „nur Firmen mit Daten" filtern, sonst fehlen die dringendsten.

**Visuelle Grammatik (Kern gegen Verwechslung):** Aktivität und Aufgabe haben
**unterschiedliche FORMEN**, nicht nur Farben.
- **Aktivität = Timeline** (`.bbz-akt-tl` mit Schiene, `.bbz-akt-ev`, **kein Rahmen**) → „lesen".
  Punktfarbe = **Kanal**, identisch mit der Mix-Bar im Panel. Klick = Detail-Modal, ✎ bei Hover.
- **Aufgabe = Karte** (Rahmen, Schatten, linker Akzent) mit **Checkbox** (`.bbz-akt-cb`) und
  Fälligkeits-Pille → „handeln".
Firma ist in beiden Zeilen **prominent** (13px fett), Kontaktart/Titel sekundär, Notiz-Vorschau
einzeilig in `--subtle`.

**Panels** (`bbz-kpis`, links Aktivitäten 1.55fr, rechts Aufgaben 1fr — spiegelt die Agenda):
- **Aktivitäten:** Anzahl **im laufenden Monat** + **Delta vs. Vormonat** + **6-Monats-Balken**
  + Ø/Monat + **Kanalmix in %** (12 Mt.). Bewusst **kein „total"** — beantwortet keine Frage.
  Balken und Mix **reagieren auf Segment/Lead-Filter** (messen die gefilterte Bearbeitung).
- **Die Balken sind ein Filter** (`akt-monat`, `F.monat` = Monatsschlüssel `year*12+month`,
  Toggle). Klick setzt den Monat und erzwingt `axis="chrono"`. **Wirkt NUR auf die
  Agenda-Aktivitäten** — nicht auf Aufgaben und **nicht auf das Firmencockpit**: dort würde ein
  Monatsfilter „Letzter Touch" verfälschen. Bei aktivem Filter zeigt die Aktivitäten-Spalte
  **eine** Gruppe `akt-p-sel` (Monatsname), weil Woche/Monat/Früher dann sinnlos wäre
  (alles fiele in „Früher"). Aufheben über den ✕-Chip im Panelkopf.
- **Aufgaben:** offen + erledigt-Zähler + Fälligkeits-Chips + **älteste überfällige Aufgabe**.
  `cDone` ist ein **Gesamtzähler**, kein Monatswert: CRMTasks hat **kein Erledigt-Datum**.

**Wiederverwendete Aktionen (nicht duplizieren):** `open-firm`, `open-contact`,
`open-history-form`, `open-task-form`, `edit-task`, `edit-history`, `complete-task`,
`task-status-change`. Neue Aktionen alle mit `akt-`-Präfix.

**Zeilen-Aktionen:** Aktivität — Klick auf die Zeile öffnet das Detail-Modal
(`open-history-detail`), ✎ öffnet Bearbeiten (`edit-history`). Aufgabe — ✓ erledigt
(`complete-task`), ✎ bearbeiten (`edit-task`).

> **Handler-Reihenfolge beachten:** `edit-history` MUSS vor `open-history-detail` geprüft
> werden. Der ✎-Button liegt *innerhalb* der klickbaren Zeile; sonst gewinnt `closest()` den
> äusseren Detail-Handler und der Stift öffnet das falsche Modal.

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


## Firmen-Screen (`views.firms()`) — Filter & Fallen

**Header** = Titel + Zähler-Badge + **Suche** + „+ Firma" in EINER Zeile. Die Suche stand
früher unter den Chips und wurde übersehen. Untertitel zeigt die **aktive Filterkette**
(`activeFilterLabel`), nicht mehr „X Firmen in dieser Ansicht". Clear-Button:
`firms-search-clear`.

**Filter-Hierarchie — drei Ebenen, drei visuelle Gewichte:**
1. **Kategorie** (`bbz-chip-lg`, 32px) — `Alle` (**Default**, `filters.kategorie === ""`) /
   Kunden / Lieferanten / Übrige.
2. **+3. im `.bbz-subfilter`-Panel** — **nur wenn `kategorie === "Kunde"`**: eingerückt,
   getönt, mit **blauer Linkskante** = „hängt an Kunden". Chips dort `bbz-chip-md` (26px),
   also **kleiner als Ebene 1**. Drei gleich schwere Chip-Reihen wirkten unaufgeräumt.
   Innerhalb durch `.bbz-subfilter-sep` getrennt, weil es **zwei verschiedene Ebenen** sind:
   - **Klassifizierung + VIP** — Label „Klassifizierung / Stammdaten": **Eigenschaft** aus
     dem SP-Feld. Chips **eckig** (`.bbz-chip-sq`) = Etikett, das an der Firma klebt.
     VIP ist ein **additiver** Toggle, unabhängig von A/B/C.
   - **Pflege-Status** — Label „Pflege-Status / errechnet": **abgeleiteter Zustand**
     (`helpers.pflegePredicate`), nicht in SP gespeichert. Chips als **Pille mit Farbpunkt**,
     Zeile im abgesetzten Block (`.bbz-subfilter-state`).
   Unterschieden wird über die **Form**, nicht über mehr Farbe — Farbe wäre hier Rauschen.
   Beide Zeilen haben einen `Alle`-Chip (`data-value=""`).

> Die Unterscheidung Stammdaten vs. errechnet ist der Grund für die Trennlinie — nicht
> Dekoration. Wer sie entfernt, verliert die Aussage.

> Stufe 2+3 sind **doppelt abgesichert**: sie werden nur gerendert *und* nur angewendet
> (`!isKunde || ...`), wenn Kunden aktiv ist; zusätzlich setzt der `firms-kategorie`-Handler
> `klassifizierung`/`vip`/`pflege` beim Verlassen von „Kunde" zurück. Ohne beides würde ein
> unsichtbarer Filter als **Geisterfilter** weiterwirken. Die Zähler in Stufe 2+3 zählen
> ebenfalls nur Kunden.

> **⚠ Klassifizierung NIE hardcoden.** Früher `["A","B","C"]` + `startsWith(k)` →
> `"Akquisition".startsWith("A") === true`, d.h. Akquisitions-Firmen zählten und filterten
> stillschweigend als **A**. Jetzt: Werte aus `state.meta.choices[CRMFirms].Klassifizierung`
> (Fallback: distinct aus dem Datenbestand) + **exakter Match**. Gleicher Bug steckte in
> `ui.detailBandClass` (`v.includes("A")`) → dort wird **Akquisition zuerst** geprüft.
> **App-weit behoben:** `helpers.klassValues()` (SP-Choices, Fallback distinct aus den Daten)
> und `helpers.klassMatches(firm, value)` (exakter Vergleich) sind die **einzige** Quelle.
> Genutzt von `views.firms()`, allen Kontakt-/Batch-Pickern und der (toten) `planning()`-View.
> **Kein `startsWith`/`includes` mehr auf `klassifizierung` im Code.** Nicht wieder einführen.

**Pflege-Prädikate** — Quelle: `helpers.pflegePredicate(kind)` + `helpers.pflegeMeta`
(bewusst überlappend — jeder Chip ist eine Frage, keine Kategorie):
| Chip | Definition |
|---|---|
| `aktiv` Aktiv gepflegt | Aktivität in **24 Mt.** ODER offene Aufgabe **mit** Termin |
| `pflege` Braucht Pflege | offene Aufgabe, die **überfällig** ist |
| `offen` Beobachten | offene Aufgabe **ohne** Datum/Termin |
| `ohne` Ohne Aktivität | keine Aktivität in 24 Mt. UND keine Aufgabe mit Datum in 24 Mt. UND **keine offene Aufgabe** |

> Die 24-Mt.-Grenze bei `aktiv` und der `!openTask`-Ausschluss bei `ohne` sind **nötig**:
> ohne sie wäre eine Firma mit Besuch von 2023 gleichzeitig „aktiv gepflegt" UND „ohne
> Aktivität", und eine Firma mit unterminierter Aufgabe „beobachten" UND „ohne Aktivität".

**Keine farbige Zeilenhinterlegung** (`bbz-row-alert/cold/ok` entfernt) — der Signal-Dot
reicht, ganze Zeilen einzufärben war Rauschen. **Nicht wieder einbauen.**

**Status/Aktivität** ist klickbar (`helpers.statusAktivitaetHtml`, eine Quelle für Desktop
und Mobile): Aufgabe → `edit-task`, Aktivität → `open-history-detail`. Precedence Task > Aktivität.

**Firma erfassen/bearbeiten:** `Kategorie` ist **Pflichtfeld** im Formular UND in
`handleFirmModalSubmit` (`fields.Kategorie`). Beides nötig — das Formularfeld allein speichert
stumm nicht, weil der Submit eine explizite Feldliste baut.

> **Vokabel-Kollision aufgelöst:** Firmen-Screen und Aktivitäten-Cockpit nutzen jetzt
> **denselben Helper** (`helpers.pflegePredicate` / `helpers.pflegeMeta`). Ein Vokabular,
> eine Definition, eine Codestelle. **Nicht wieder lokal nachbauen.**


## Pflege-Status: eine Quelle für zwei Screens

`helpers.pflegeMeta` (Label/Farbe/Erklärung) und `helpers.pflegePredicate(kind)` sind die
**einzige** Definition. Genutzt von:
- `views.firms()` → Chip-Zeile „Pflege · Kunden" (`firms-pflege`, Toggle)
- `views.aktivitaeten()` → Firmencockpit-Signalfilter (`akt-sig`, exklusiv, Default `aktiv`)

Zustände: `aktiv` · `pflege` · `offen` · `ohne` · `kein` (nur Segment „alle").

> **`firmSignal` wird nicht mehr aufgerufen.** Die **Dots** in der Firmen-Tabelle laufen jetzt
> über `helpers.pflegeDot(firm)` — also über dieselben Prädikate wie Chips und Legende.
> Vorher widersprachen sich Punkt (12 Mt.) und Chip (24 Mt.) bei gleicher Beschriftung.
> `firmSignal` steht nur noch als ungenutzte Funktion im Code (nicht als toter Code löschen,
> ohne zu prüfen, ob ein künftiges Dashboard sie braucht).

**`helpers.pflegeDot(firm)`** — die Zustände überlappen, ein Punkt kann nur einen zeigen.
Feste Rangfolge **dringend vor unauffällig**: `pflege` ▸ `offen` ▸ `ohne` ▸ `aktiv`;
Nicht-Kunden und Randfälle → `null` (kein Punkt). Die **Legende wird aus `pflegeMeta`
generiert** — sie kann also nicht mehr veralten. Nicht durch fixen Text ersetzen.


## Kontakte-Screen — KPI-Zeile

Die Kacheln **„Offene Tasks"** und **„Firmen-Cockpit"** wurden **entfernt** (Navigation gehört
in die Nav, nicht in KPI-Kacheln). An ihrer Stelle steht der **Geburtstagskalender**:

- `.bbz-kpi-wide` (`grid-column: span 2`) belegt den Platz der zwei entfernten Kacheln.
- `.bbz-kpi-static` = Container ohne Hover-Lift; **die Zeilen darin sind klickbar**
  (`open-contact`), nicht die Karte.
- Inhalt: Anzahl Geburtstage in 30 Tagen (+ „N heute"), die nächsten 4 als Liste,
  „alle anzeigen →" führt auf die Route `birthdays`.

> **Nichts davon ist neu gebaut.** Es nutzt die im Handover vorgehaltenen Helper
> `helpers.upcomingBirthdays(days, contacts)` und `helpers.birthdayLabel(daysUntil, nextBirthday)`.
> Diese sind damit **nicht mehr ungenutzt** — der Hinweis „vorgehalten fürs Dashboard" gilt
> für sie nicht mehr, wohl aber weiterhin für die KPI-Aggregations-Helper.

> **Die Route `birthdays` (`views.birthdayView()`) war bis dahin über die UI unerreichbar** —
> sie stand in `knownRoutes` und im Dispatch, aber kein Link führte hin. Der „alle anzeigen →"-
> Link ist jetzt der einzige Einstieg. Wer ihn entfernt, macht die View wieder unerreichbar.

`allOpenTasks`/`overdueTasks` in `views.contacts()` sind mit den Kacheln entfallen
(die gleichnamigen Locals in `views.firms()` sind davon unberührt).


## Aufgaben OHNE Termin

Sie fielen durch **alle** Fälligkeits-Buckets (`Überfällig`/`Diesen Monat`/`Später` prüfen alle
auf ein Datum) und waren in der Agenda **unsichtbar** — die Panel-Chips summierten dann auch
nicht mehr auf „offen".

- Zustand = **„Beobachten"** (`helpers.pflegePredicate("offen")`) — greift in Firmen-Screen,
  Firmencockpit und Dots über den gemeinsamen Helper.
- **Agenda:** eigener Bucket `akt-c-undated` „Beobachten · ohne Termin", direkt nach
  „Überfällig" und **default offen** — sie brauchen eine Handlung (Termin setzen).
- **Panel:** Chip „Beobachten" (`F.faelligkeit === "undated"`), nur sichtbar wenn > 0.
  Damit gilt wieder: Überfällig + Beobachten + Diesen Monat + Später = **offen**.
- **Cockpit:** „Nächste Aufgabe" zeigt bei fehlendem Termin „ohne Termin" statt eines leeren
  Datums.

> `helpers.isOverdue("")` liefert korrekt `false` — eine unterminierte Aufgabe ist **nicht**
> überfällig, sondern unterminiert. Das ist der Grund, warum sie in `offen` landet und nicht
> in `pflege`. Nicht "reparieren".


## Dashboard (`views.dashboard()`) — Startseite

**`CONFIG.defaults.route = "dashboard"`.** State: `{ sel, per, foldOpen }`.

**Drei Zonen nach ÄNDERUNGSFREQUENZ, nicht nach Sachgebiet** — eine Startseite muss zuerst
„was braucht mich jetzt?" beantworten:
1. **Handeln** (täglich) — überfällig / diese Woche / ohne Termin. **Einzige Zone mit Rot.**
   Eine Null ist grau und **nicht klickbar**: nichts zu tun soll auch so aussehen.
   Darunter die Geburtstagskarte (Aufschlüsselung + **Erfassungsgrad**).
2. **Steuern** (monatlich) — Zeitreihe + Abdeckungs-Matrix.
3. **Pflegen** (selten, `dash-fold` einklappbar) — Stammdaten, Datenqualität, stille Wächter.

**Ein Mechanismus:** jede Zahl trägt `data-action="dash-select"` → `F.sel` steuert die
Drill-Down-Liste unten. Metrik-Registry `M` = **eine Quelle** für Zähler und Liste
(`set()` liefert die Menge, `kind` die Spalten). Neue Kennzahl = ein Eintrag in `M`.

> **⚠ Raster NUR über Klassen** (`.bbz-dash-g2/g3/g4`, `.bbz-dash-act`, `.bbz-dash-bd`).
> Gegen inline `grid-template-columns` kommt **keine Media-Query** an (ausser mit
> `!important`-Hacks) — genau daran scheiterte die erste Fassung. Breakpoints: 1000 / 780 /
> 620 / 560 / 480px.

**Zeitreihe** (`dash-per`: 30 / 12 / all) als **Fläche, nicht als Balken**. Die Linie zeichnet
sich links→rechts — **die Animation IST die Zeitachse**.
> **30 Tage = gleitende 7-Tage-Summe**, nicht Tageswerte. Bei ~0,7 Aktivitäten/Tag ist die
> Tageskurve Rauschen und ein „Tagesrekord" Unsinn. Die Glättung macht Momentum sichtbar und
> den Bestwert erst sinnvoll. Messlatte wächst mit der Auflösung: **Beste Woche · Stärkster
> Monat · Stärkstes Jahr** (Genus in `P.sup` mitführen — „Stärkster Jahr" wäre falsch).
> **Kein erfundenes Zielwert** — die eigene Historie ist die einzige ehrliche Messlatte.

**Abdeckungs-Matrix** statt Donut: „30%" beantwortet nicht, **welche** 30%. Zeilen =
Klassifizierung (aus `helpers.klassValues()`) + „ohne Klassifizierung" (⚠, da für sie die
Priorisierung blind ist) + Gesamt.
> **Nur der abgedeckte Anteil wird gefüllt, „ohne" ist die leere Spur.** Vorher war „ohne"
> sattes Rot = 70% jeder Zeile: lauteste Farbe für die nutzloseste Aussage. Nicht zurückbauen.

**Donut** nur bei Stammdaten (echtes Teil-vom-Ganzen). **Weisse Trennlücken** (`GAP`), weil
`--amber`/`--red`/`--green` in der Helligkeit zu nah liegen und sonst verschmelzen. Das Loch
ist ein **Anzeigeplatz**: Hover tauscht die Zahl darin — deshalb überhaupt ein Donut.
**Datenqualität = Ring-Gauges, kein Donut** — die Quoten summieren nicht auf 100%.

**Stille Wächter** (`int-firms`/`int-contacts`/`int-orphan`): nur sichtbar, wenn > 0. Die
Formulare erzwingen Firma (`firmaLookupId required`) bzw. Kontakt (`kontaktLookupId required`)
— diese Fälle entstehen nur über SharePoint direkt, IO-Import oder eine frisch angelegte Firma.

> **`controller.afterRender()`** läuft nach jedem Render und ist auf `route === "dashboard"`
> gegated. Dort gehört alles hin, was **gemessene Geometrie** braucht (`getTotalLength` für die
> Linien-Animation) oder **Hover ohne Re-Render** (Chart-Tooltip, Donut-Loch). Die Views
> liefern nur HTML-Strings.

> **`normalizer.firm()` mappt `spCreated`/`spCreatedBy`** — war als einziges Entity nicht
> gemappt, obwohl `createdDateTime` für alle Listen geholt wird. Nicht entfernen.
> Vorbehalt: es ist das **SharePoint-Anlagedatum**, nicht der fachliche Beziehungsbeginn.

> **CSS-Klammerbilanz prüfen** (`{` == `}`), nicht nur `node --check`. Eine einzige
> überzählige Klammer verschluckt den Rest des Stylesheets — und `node --check` sieht kein CSS.
