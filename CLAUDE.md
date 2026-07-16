# CLAUDE.md — bbz CRM

Projektgedächtnis für Claude Code. **Lies das vor jeder Änderung.**

> **Struktur:** Grundlagen → Konventionen → Deploy → Routen → Querschnitts-Helper →
> **Fallen** → Tote Zonen. Der Abschnitt **Fallen** ist der wichtigste: dort steht, was uns
> schon Zeit gekostet hat.
>
> **Keine Zeilennummern in dieser Doku.** Sie driften bei jedem Patch und werden zur Lüge.
> Verweise gehen auf **Namen** (`views.firms()`, `helpers.pflegeDot`) — die sind grepbar.

---

## Was das ist

Framework-freie Vanilla-JS-PWA. CRM für bbz Beratung: Firmen, Kontakte, Aktivitäten, Events.
Backend ist SharePoint (MS Graph v1.0, MSAL-Auth). Gehostet auf GitHub Pages:
https://markusbaechler.github.io/crm-spa/

**Kein Build-Step, kein Paketmanager, kein Bundler.** Die Dateien werden 1:1 ausgeliefert.

## Dateikarte

| Datei | Inhalt |
|---|---|
| `index.html` | App-Shell + **gesamtes CSS** (`:root`-Tokens oben) + Desktop-Nav + Bottom-Nav |
| `app.js` | Gesamte Logik (~6000 Z.): CONFIG, SCHEMA, state, helpers, dataModel, views, controller |
| `io.js` | Import/Export via SheetJS (`window.bbzIO`) |
| `service-worker.js` | Network-first für Navigationen, cacht nur Offline-URL |
| `manifest.json` | PWA-Manifest, scope `/crm-spa/` |

## Architektur (app.js)

- `CONFIG` — Graph tenantId/clientId, SharePoint-Site, Listennamen, `defaults.route`.
- `SCHEMA` — Mapping App-Feld → SharePoint-Feldname pro Liste.
- `state` — auth, meta, data (roh), enriched (angereichert), filters, selection, modal.
- `helpers` — Escaping, Datums-Utils, **Querschnitts-Logik** (siehe unten), `debounce`.
- `dataModel.enrich()` — join firms/contacts/history/tasks. Leitet ab:
  - firms → `contactsCount, contacts[], tasks[], history[], openTasksCount, nextDeadline, latestActivity`
  - contacts → `fullName, firm, firmId, firmTitle, openTasksCount`
  - history → `contactId, contactName, firmId, firmTitle, projektbezugBool`
  - tasks → `isOpen, isOverdue`
- `views` — reine String-Builder pro Route. `views.renderRoute()` dispatcht per `state.filters.route`.
- `controller.render()` — `ui.renderShell()` + `ui.renderView(views.renderRoute())` → `innerHTML`,
  danach `controller.afterRender()`.

> **Die Verkettung ist Aktivität → Kontakt → Firma.** `enrich()` löst History über
> `kontaktLookupId` zum Kontakt und über `contact.firmaLookupId` zur Firma auf. Eine Firma
> **ohne Kontakte** hat zwangsläufig `latestActivity = ""` und kann durch keine Aktivität je
> „gepflegt" werden. Wer Abdeckungszahlen interpretiert, muss das wissen.

## Datenmodell (SharePoint-Listen)

- **CRMFirms** — Title, Adresse, PLZ, Ort, Land, Hauptnummer, **Klassifizierung** (Choice, u.a.
  A-Kunde/B-Kunde/C-Kunde/**Akquisition** — Werte NIE hardcoden, s. Fallen), VIP,
  **Kategorie** (Choice: Kunde/Lieferant/Übrige)
- **CRMContacts** — Title(=Nachname), Vorname, Anrede, Firma(**Lookup, `required`**), Funktion,
  Email1/2, Direktwahl, Mobile, Rolle, Leadbbz0, SGF, Geburtstag, Kommentar, Event,
  Eventhistory, Archiviert
- **CRMHistory** — Title, Nachname(**Lookup, `required`**), Datum, Kontaktart(=`typ`), Notizen,
  Projektbezug, Leadbbz
- **CRMTasks** — Title, Name(Lookup), Deadline (**darf leer sein**), Status, Leadbbz

Alle Entitäten tragen `spCreated` / `spCreatedBy` (aus `createdDateTime`). Die Fetch-Schicht
holt das für **alle** Listen; `normalizer.firm()` mappt es seit dem Dashboard ebenfalls —
**nicht entfernen**, sonst bricht die Firmen-Entwicklung.

> `spCreated` ist das **SharePoint-Anlagedatum**, nicht der fachliche Beziehungsbeginn.
> Bei Migrations-Importen tragen hunderte Datensätze dasselbe Datum.

> **Lead-Semantik:** „Lead bbz" = **Record-Lead** (`history.leadbbz` / `task.leadbbz`),
> **kein** Fallback auf `contact.leadbbz0`.

## Konventionen — strikt einhalten

1. **Kein Framework/Build einführen.** Bleibt Vanilla. Keine neuen Dependencies ohne Auftrag.
2. **Rendering:** Views geben HTML-Strings zurück. Interaktion ausschliesslich über
   `data-action="..."` + zentrale Event-Delegation.
3. **Escaping:** Jeder Nutzer-/SP-Wert im HTML MUSS durch `helpers.escapeHtml()`.
4. **Styling:** Nur bestehende CSS-Variablen und `bbz-*`-Klassen. Keine Ad-hoc-Farben.
5. **State:** Filter/Selektion in `state.filters.<route>` bzw. `state.selection`.
   Nach Mutation `controller.render()` — **ausser** wo das Formulardaten zerstören würde (s. Fallen).
6. **SharePoint-Write:** POST speichert zuverlässig nur `Title` + Lookups. Restfelder per
   separatem PATCH auf die neue Item-ID nachschreiben. Muster nicht umgehen.
7. **Deutsch (CH):** UI-Texte deutsch, „ß" immer als „ss".
8. **Doku im selben Commit:** Jede Verhaltensänderung aktualisiert diese Datei — im **selben** Commit.
9. **Eine Quelle pro Begriff.** Wo dieselbe Frage an zwei Stellen beantwortet wird, gehört sie
   in einen Helper. Zwei Vokabulare mit denselben Wörtern haben uns schon zweimal eingeholt.

## Deploy (autonom)

Trigger: Push auf `main`. Workflow `.github/workflows/deploy.yml`:
1. `node --check` auf app.js/io.js/service-worker.js — bricht bei Syntaxfehler ab.
2. Stampt kurze Commit-SHA als Cache-Buster in Script-Tags + SW-CACHE-Namen.
3. Publiziert nach GitHub Pages.

**Regeln:** Nie mit rotem `node --check` pushen. Ein Commit = eine logische Änderung.
**Genau EIN Pages-Workflow** (`deploy.yml`) — ein zweiter kollidiert beim Artefakt-Upload.
**„Deployment failed / try again later"** ist meist transient (githubstatus.com), **kein**
Code-Fix: `gh workflow run deploy.yml --ref main` (nicht `gh run rerun --failed` — dupliziert
Artefakte). Deploy-Status ohne `$(...)` ermitteln: `gh run list` → ID lesen → `gh run watch <ID>`.

**Lokal:** `python -m http.server 8080`. Auth-Gotcha: MSAL `redirectUri` ist auf die Prod-URL
fixiert; Login lokal nur, wenn `http://localhost:8080/` in der Azure-AD-App registriert ist.

---

# Routen

Nav (Desktop + Bottom): **Dashboard · Firmen · Kontakte · Aktivitäten · Events**.
`CONFIG.defaults.route = "dashboard"`.
Nicht in der Nav: `birthdays` (hängt an **einem** Link in der Geburtstagskarte — wer ihn
entfernt, macht die View unerreichbar), `admin`.
Redirect: `planning` und `history` → `aktivitaeten` (alte Bookmarks überleben).
Die alten Views sind **gelöscht** — es gibt nur noch den Redirect und die Einträge in
`knownRoutes`. **Beide müssen bleiben**, sonst landen alte Links auf der Default-Route.

## `dashboard` — Startseite (`views.dashboard()`)

State: `{ sel, per, foldOpen }`.

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

**Zeitreihe** (`dash-per`: 30 / 12 / all) als **Fläche, nicht als Balken**. Die Linie zeichnet
sich links→rechts — **die Animation IST die Zeitachse**.
- **30 Tage = gleitende 7-Tage-Summe**, nicht Tageswerte. Bei ~0,7 Aktivitäten/Tag ist die
  Tageskurve Rauschen und ein „Tagesrekord" Unsinn. Die Glättung macht Momentum sichtbar und
  den Bestwert erst sinnvoll.
- Messlatte wächst mit der Auflösung: **Beste Woche · Stärkster Monat · Stärkstes Jahr**
  (Genus in `P.sup` mitführen — „Stärkster Jahr" wäre falsch).
- **Kein erfundener Zielwert.** Die eigene Historie ist die einzige ehrliche Messlatte.

**Abdeckungs-Matrix** statt Donut: „30%" beantwortet nicht, **welche** 30%. Zeilen =
Klassifizierung (aus `helpers.klassValues()`) + „ohne Klassifizierung" (⚠) + Gesamt.
> **Das ⚠ hier ist strittig** — in `aktivitaeten` ist `Ohne` eine gleichwertige Kategorie
> (passiv betreute Kunden), kein Mangel. Solange das Dashboard davor warnt und die
> Aktivitäten-Route nicht, sind es zwei Aussagen zum selben Zustand. Offen, bewusst nicht
> im selben Commit geändert.
- **Nur der abgedeckte Anteil wird gefüllt, „ohne" ist die leere Spur.** Vorher war „ohne"
  sattes Rot = 70% jeder Zeile: lauteste Farbe für die nutzloseste Aussage.
- **Jedes Element ist ein eigenes Klickziel** (`covSets(key, lab, base)` legt fünf Mengen an):
  Label/n → alle · grün → `__m6` · amber → `__m12` · **leere Spur** → `__none` (Arbeitsliste) ·
  **Prozentzahl** → `__cov` (abgedeckt). Man klickt, was man sieht.
- Labels müssen sagen, was sie zeigen: „Alle B-Kunde" vs. „B-Kunde · abgedeckt". Das frühere
  „`<k>` · Abdeckung" war zweideutig — Zähler korrekt, Beschriftung falsch.

**Donut** nur bei Stammdaten (echtes Teil-vom-Ganzen), mit **weissen Trennlücken** (`GAP`).
Sein Loch ist ein **Anzeigeplatz**: Hover tauscht die Zahl darin — deshalb überhaupt ein Donut.
**Datenqualität = Ring-Gauges, kein Donut** — die Quoten summieren nicht auf 100%.

**Stille Wächter** (`int-firms`/`int-contacts`/`int-orphan`): nur sichtbar, wenn > 0. Die
Formulare erzwingen Firma bzw. Kontakt — diese Fälle entstehen nur über SharePoint direkt,
IO-Import oder eine frisch angelegte Firma.

**Drill-Down-Liste ist kontextbewusst:** konstante Spalten fliegen raus (sie wiederholen nur
den Filter), Firmen tragen eine **Zustands-Spalte** (Abdeckungsband), sortiert **längster
Kontaktabstand zuerst**. Braucht ein Mobile-Pendant (s. Fallen).

## `firms` — Firmenboard (`views.firms()`)

**Header** = Titel + Zähler-Badge + **Suche** + „+ Firma" in EINER Zeile. Untertitel zeigt die
**aktive Filterkette** (`activeFilterLabel`). Clear-Button `firms-search-clear`.

**Drei Filter-Ebenen, drei visuelle Gewichte:**
1. **Kategorie** (`bbz-chip-lg`, 32px) — `Alle` (**Default**, `filters.kategorie === ""`) /
   Kunden / Lieferanten / Übrige.
2. **+3. im `.bbz-subfilter`-Panel** — **nur wenn `kategorie === "Kunde"`**: eingerückt,
   getönt, blaue Linkskante = „hängt an Kunden". Chips `bbz-chip-md` (kleiner = untergeordnet).
   Innerhalb durch `.bbz-subfilter-sep` getrennt, weil es **zwei Ebenen** sind:
   - **Klassifizierung + VIP** („Stammdaten") — **eckige** Chips (`.bbz-chip-sq`) = Etikett.
     VIP ist ein **additiver** Toggle, unabhängig von A/B/C.
   - **Pflege-Status** („errechnet") — Pillen mit Farbpunkt, im Block `.bbz-subfilter-state`.
   Unterschieden wird über die **Form**, nicht über mehr Farbe.

> Stufe 2+3 sind **doppelt abgesichert**: nur gerendert *und* nur angewendet (`!isKunde || ...`),
> plus Reset beim Kategoriewechsel — sonst wirkt ein unsichtbarer **Geisterfilter** weiter.
> Die Zähler in Stufe 2+3 zählen ebenfalls nur Kunden.

**Tabelle:** Dot · Firma · Ort · Klassifizierung · Kontakte · Status/Aktivität.
**Keine farbige Zeilenhinterlegung** — der Dot reicht, ganze Zeilen einzufärben war Rauschen.
**Status/Aktivität ist klickbar** (`helpers.statusAktivitaetHtml`, eine Quelle für Desktop und
Mobile): Aufgabe → `edit-task`, Aktivität → `open-history-detail`. Precedence Task > Aktivität.

**Dot** = `helpers.pflegeDot(firm)`. **Legende wird aus `helpers.pflegeMeta` generiert** —
sie kann nicht veralten. Nicht durch fixen Text ersetzen.

**Firma erfassen/bearbeiten:** `Kategorie` ist **Pflichtfeld** im Formular UND in
`handleFirmModalSubmit` (`fields.Kategorie`). Beides nötig — das Formularfeld allein speichert
stumm nicht, weil der Submit eine explizite Feldliste baut.

## `contacts` — Kontakte (`views.contacts()`)

KPI-Zeile: Kontakte (mit Modus-Chips) · Angezeigt · **Geburtstagskalender**
(`.bbz-kpi-wide` = span 2, `.bbz-kpi-static` = Container ohne Hover-Lift; die **Zeilen** darin
sind klickbar). „alle anzeigen →" führt auf `birthdays`.
Die früheren Kacheln „Offene Tasks" und „Firmen-Cockpit" sind **entfernt** — das war
Navigation als Kachel getarnt.

## `aktivitaeten` — drei Zonen (`views.aktivitaeten()`)

Nachfolger der gelöschten Routen `planning` + `history` (Redirect s.o.).
State: `{ segment, axis, search, lead, klass, cut, faelligkeit, sig, monat, expandedFirms,
bucketOpen, moreOpen, legendeOffen }`.

**Drei Zonen nach ARBEITSSCHRITT**, mit dem Zonenkopf-Muster des Dashboards
(`bbz-zone`: Nummer · Titel · Leitfrage · Linie · Meta):

1. **Fokus setzen** — „Wo schaue ich hin?" Klassifizierungs-Karten + Lead-Chips.
2. **Markt bearbeiten** — „Was lief, was steht an?" Agenda oder Firmencockpit.
3. **Auswerten & Steuern** — „Wohin entwickelt sich die Betreuung?" 12 Monate + Kanaltabelle.

**Zone 1 filtert Zone 2. Zone 3 misst, was Zone 2 zeigt** — Segment, Lead und Fokus wirken
durch alle drei. Das ist die einzige Richtung; es gibt keinen Rückkanal.

### Zone 1 — Fokus setzen

Karten aus **`helpers.klassValues()`** erzeugt, plus `Alle` (Default) und `Ohne`.
Kommt ein neuer Choice-Wert in SharePoint dazu, erscheint die Karte von selbst.

> **`Ohne` ist eine Kategorie, kein Mangel.** Rund die Hälfte der Kunden wird bewusst passiv
> betreut — leer heisst „nicht klassifiziert", nicht „vergessen". Kein ⚠, keine Warnfarbe.
> `F.klass`: `""` = alle, `"__none"` = ohne Klassifizierung, sonst exakter Choice-Wert
> (Vergleich über `helpers.klassMatches` — **nie** `startsWith`, s. Klassifizierung).

Jede Karte zeigt **Aktivitäten der letzten 30 Tage** und **überfällig · offen**. Beim
Segmentwechsel wird `klass` zurückgesetzt (sonst Geisterfilter, gleiche Falle wie bei `firms`).

**Lead-Chips zählen Aktivitäten und Aufgaben GETRENNT** (`AD 22 · 12`). Ein gemeinsamer
Zähler mischte zwei Objekttypen, die die View sonst konsequent durch die Form trennt — und
liess die Chip-Summe (87) über der Zahl der Aktivitäten (64) liegen.

### Zone 2 — Markt bearbeiten

**Visuelle Grammatik (Kern gegen Verwechslung):** Aktivität und Aufgabe haben
**unterschiedliche FORMEN**, nicht nur Farben.
- **Aktivität = Timeline** (`.bbz-akt-tl`, **kein Rahmen**) → „lesen". Punktfarbe = **Kanal**,
  identisch mit der Kanaltabelle in Zone 3. Klick = Detail-Modal, ✎ bei Hover.
- **Aufgabe = Karte** (Rahmen, Schatten, linker Akzent) mit **Checkbox** → „handeln".
- Firma ist in beiden Zeilen **prominent** (13px fett), dahinter der Klassifizierungs-Badge.

Spaltenüberschriften sind die Frage, die die Spalte beantwortet: **„Was lief wann bei wem"**
links, **„Was steht wann wo an"** rechts. Die Trennung ist **Objekttyp**, nicht Zeit.

**Zwei Achsen** (`akt-axis`, im Zonenkopf), **Default `chrono`**:
- `chrono` „Agenda" = Hauptansicht. Zweispaltig (`.bbz-akt-split`). Links `akt-p-week`/
  `akt-p-month` offen, `akt-p-old` zu; rechts `akt-c-over`, `akt-c-undated`, `akt-c-month`
  offen, `akt-c-later`, `akt-c-done` zu.
- `firm` „Firmencockpit": **Signal-Filter statt Rubriken** (`akt-sig`, Default `aktiv`) —
  **immer genau EINE Kategorie sichtbar**. Gliederung nach letztem Kontakt, **gleiche Richtung
  wie die Agenda (neu→alt)**. `firmRows` umfasst **alle** Firmen im Fokus, auch nie kontaktierte.

> **`Nächste 30 Tage`, NICHT `Diesen Monat`.** Die Grenze ist `today + 30`, nicht das
> Monatsende. Am 16.07.2026 standen unter „Diesen Monat" drei **August**termine (04./06./14.08.)
> — das Label log über die Mechanik. Wer auf Kalendermonat umstellen will, ändert die Logik,
> nicht nur den Text.

### Zone 3 — Auswerten & Steuern

**Rollierende 12 Monate**, einfarbige Balken: **Höhe = Menge**. Bestwert grün.
Die Balken sind ein **Filter** (`akt-monat`, Toggle) — wirkt **nur** auf die Agenda-Aktivitäten,
nicht auf Aufgaben oder Cockpit (dort würde er „letzter Touch" verfälschen).

> **Der Ø läuft ab dem ersten Monat MIT Erfassung**, nicht über volle 12 Monate. Die App ging
> **Februar 2026** produktiv; ein 12-Monats-Schnitt mittelt gegen eine Zeit, in der es sie nicht
> gab, und ist strukturell unerreichbar. Aus `Ø 5,3/Mt.` wurde so `Ø 9,5/Mt. seit Feb`.

**Kanaltabelle: der Kanal ist die ZEILE, der Schnitt wechselt die Spalten** (`akt-cut`):

| Schnitt | Spalten | Frage |
|---|---|---|
| `verlauf` | n · Sparkline 12 Mt. · % | Verschiebt sich der Kanalmix? |
| `klass` | aus `helpers.klassValues()` + `ohne` | Landen Besuche bei A-Kunden? |
| `lead` | Lead-bbz-Werte + `ohne` | Wer arbeitet wie? |

**Raster der Kanaltabelle**: Klassen `.bbz-akt-kan` (Verlauf) und `.bbz-akt-kanc` (Schnitte).
Variabel ist **nur die Spaltenzahl**, als Custom Property `--kan-n` inline. Das Raster selbst
gehört in die Klasse — sonst greift die Media-Query bei 620px nicht.

> **Menge und Mix bleiben getrennt.** Gestapelte Balken (Kanal als Segment im Monatsbalken)
> waren bei ~9 Aktivitäten/Monat unlesbar — fünf Farben auf zwölf schmalen Balken. Deshalb:
> Balken zeigen **wie viel**, Tabelle zeigt **womit**.
> **Kein Kanalmix-Prozentbalken im Kopf** — „Telefon 38%" beantwortet keine Handlungsfrage.
> Erst der Lead-Schnitt macht daraus eine Aussage (AD: 0 Online in 9 Monaten).

Unter 10 Aktivitäten im gewählten Monat erscheint ein Hinweis, dass Anteile Einzelfälle sind.

**Löschen NUR im Bearbeiten-Modus** (gegen versehentliches Löschen): kein ✕ in Zeilen, kein
Löschen im read-only Detail-Modal. **Nicht wieder ✕ in Zeilen bauen.**

**Aktivitäts-Detail-Modal** (`history-detail`, `views.renderHistoryDetail`,
`controller.openHistoryDetail`): read-only, **ungekürzte Notizen**, Footer Schliessen /
Bearbeiten.


---

# Querschnitts-Helper — je EINE Quelle

## Pflege-Status — `helpers.pflegeMeta` + `helpers.pflegePredicate(kind)`

Genutzt von `views.firms()` (Chips), `views.aktivitaeten()` (Cockpit-Filter) **und**
`helpers.pflegeDot` (Tabellen-Dots). **Nicht lokal nachbauen.**

| Zustand | Definition |
|---|---|
| `aktiv` „Aktiv gepflegt" | Aktivität in **24 Mt.** ODER offene Aufgabe **mit** Termin |
| `pflege` „Braucht Pflege" | offene Aufgabe, die **überfällig** ist |
| `offen` „Beobachten" | offene Aufgabe **ohne** Datum/Termin |
| `ohne` „Ohne Aktivität" | keine Aktivität in 24 Mt. UND keine Aufgabe mit Datum in 24 Mt. UND **keine offene Aufgabe** |
| `kein` | Nicht-Kunde |

Bewusst **überlappend** — jeder Chip ist eine Frage, keine Kategorie. Eine Firma darf
gleichzeitig `aktiv` und `pflege` sein (frischer Kontakt + überfällige Aufgabe).

> Die **24-Mt.-Grenze** bei `aktiv` und der **`!openTask`-Ausschluss** bei `ohne` sind nötig:
> ohne sie wäre eine Firma mit Besuch von vor 3 Jahren gleichzeitig „aktiv gepflegt" UND „ohne
> Aktivität", und eine Firma mit unterminierter Aufgabe „beobachten" UND „ohne Aktivität".

**`helpers.pflegeDot(firm)`** — Zustände überlappen, ein Punkt kann nur einen zeigen. Feste
Rangfolge **dringend vor unauffällig**: `pflege` ▸ `offen` ▸ `ohne` ▸ `aktiv`; sonst `null`.

> **`helpers.firmSignal` ist gelöscht** (hatte 0 Aufrufe). Wer es neu baut, baut das alte
> Doppel-Vokabular wieder auf. `pflegeDot` ist die einzige Quelle für den Punkt.

## Klassifizierung — `helpers.klassValues()` + `helpers.klassMatches(firm, value)`

Werte aus `state.meta.choices[CRMFirms].Klassifizierung`, Fallback: distinct aus dem
Datenbestand. Vergleich **exakt**.

> **⚠ NIE `["A","B","C"]` hardcoden und NIE `startsWith`/`includes` auf `klassifizierung`.**
> `"Akquisition".startsWith("A") === true` — Akquisitions-Firmen zählten und filterten
> **stillschweigend als A**: falsche Zähler UND falsche Mengen. Der Bug steckte an **sechs**
> Stellen (Firmen-Filter, `detailBandClass`, drei Kontakt-Picker, der alten `planning`-View).
> Im Code gibt es **kein** `startsWith`/`includes` auf `klassifizierung` mehr.
> `ui.detailBandClass` prüft **Akquisition zuerst**, sonst greift `includes("A")`.

## Kontakt-Auswahl — `helpers.contactOptionsHtml()` + `helpers.contactFirmFilterHtml()`

Einzige Quelle für `renderHistoryForm` **und** `renderTaskForm`. Nach Firma gruppiert
(`<optgroup>`), alphabetisch, „ohne Firma" ans Ende, plus **Firmen-Vorfilter**.

> **Warum das wichtig ist:** Vorher standen ~500 Kontakte **unsortiert in SharePoint-
> Reihenfolge** im Dropdown — man konnte einen Namen nicht finden, nur suchen. Bei **1,2
> erfassten Aktivitäten pro Woche im ganzen Team** ist nicht die Pflege das Problem, sondern
> das Erfassen. **Nicht auf eine flache, unsortierte Liste zurückbauen.**

`prefillFirmId` wird **aus dem vorgewählten Kontakt abgeleitet** — stimmt so auch im Edit-Modus.

## Geburtstage — `helpers.upcomingBirthdays(days, contacts)` + `helpers.birthdayLabel()`

Genutzt von `contacts`, `dashboard`, `firmDetail`, `birthdayView`. Behandelt Jahreswechsel und
fehlende Daten korrekt. **Nicht als toten Code entfernen.**

## Status/Aktivität — `helpers.statusAktivitaetHtml(firm)`

Eine Quelle für Desktop-Tabelle und Mobile-Karte. Precedence Task > Aktivität.

---

# ⚠ Fallen — hier haben wir schon Zeit verloren

## Tokens & CSS

- **`--red` MUSS `#a4161a` bleiben.** Ein versehentliches Teal (`#0d6e6a`) liess sämtliche
  Danger-Elemente grünlich rendern — Logik korrekt, nur der Token falsch.
- **CSS-Klammerbilanz prüfen** (`{` == `}` im `<style>`), **nicht nur `node --check`** — das
  sieht kein CSS. **Eine** überzählige Klammer verschluckt den Rest des Stylesheets: schwarze
  Gauges *und* zerlaufene Listen aus einer einzigen Ursache.
- **Raster NUR über Klassen** (`.bbz-dash-g2/g3/g4`, `.bbz-akt-split`, `.bbz-akt-kan(c)`,
  `.bbz-modal-filters`, `.bbz-ev-stats`, `.bbz-subfilter`).
  Gegen inline `grid-template-columns` kommt **keine Media-Query** an (ausser mit
  `!important`-Hacks). Daran scheiterte die erste Dashboard-Fassung — **und die erste Fassung
  der Zonen-Umstellung**: die Kanaltabelle in Zone 3 hatte bis zu 6 inline-Spalten, die
  Filterzeilen in `renderBatchEventForm` und `renderEventMatrixModal` je 4. Auf dem Handy
  überlief die Kanaltabelle die Seite; ein überlaufendes Element macht das ganze Dokument
  breiter und verschiebt damit auch `position:fixed`-Modals.
  **Braucht ein Raster eine variable Spaltenzahl, gehört sie als Custom Property inline
  (`--kan-n`), nicht das Raster.**
  Prüfen mit: kein `style="…grid-template-columns:…"` ausser `1fr` / `auto-fit` / `auto-fill`.
- `--amber` (#8a5c00), `--red` (#a4161a), `--green` (#186935) liegen in der **Helligkeit zu nah**
  beieinander. In Donuts **weisse Trennlücken** statt neuer Farben.
- **NICHT `.bbz-history-split` für neue Layouts** — die blendet Spalte 2 mobil aus (altes
  Tab-Bar-Konzept).

## Mobile

- **Die App blendet JEDE `.bbz-table-wrap` bei ≤640px aus.** Jede Liste braucht ein Pendant
  `<div class="bbz-card-list bbz-mobile-only">` mit `.bbz-list-card` — sonst ist sie auf dem
  Handy **unsichtbar**. Im `dq-*`-Kontext zeigt die Karte die **fehlende** Angabe rot als
  „fehlt" — sie ist dort die Nachricht.
- Service-Worker cached hart: nach Deploy **immer mit `?v=neu`** prüfen (Inkognito reicht nicht).

## Event-Delegation

- **Handler-Reihenfolge:** `edit-history` MUSS **vor** `open-history-detail` geprüft werden.
  Der ✎-Button liegt *innerhalb* der klickbaren Zeile; sonst gewinnt `closest()` den äusseren
  Handler und der Stift öffnet das falsche Modal.
- Verschachtelte `data-action`s sind unproblematisch, **wenn** das innerste gewinnen soll —
  `closest()` läuft von innen nach aussen (so in der Abdeckungs-Matrix genutzt).
- **Keine command substitution `$(...)`** in CC-Befehlen (löst Bestätigungs-Prompt aus).

## Rendern

- **`views.renderRoute()` ist ein Wrapper mit `try/catch`** um `views.renderRouteInner()`.
  Ohne ihn führt **ein** Fehler in **einer** View zur **komplett weissen App**. Genau das wäre
  bei `ui.afterRender()` statt `this.afterRender()` passiert: syntaktisch tadellos, zur Laufzeit
  `TypeError`. **`node --check` findet so etwas nicht.** Nicht entfernen, nicht umgehen.
- **`controller.afterRender()`** läuft nach jedem Render, gegated auf `route === "dashboard"`.
  Dorthin gehört alles, was **gemessene Geometrie** braucht (`getTotalLength`) oder **Hover
  ohne Re-Render** (Chart-Tooltip, Donut-Loch). Views liefern nur Strings.
- **Der Firmen-Vorfilter im Formular (`form-contact-firm`) darf KEIN `controller.render()`
  auslösen** — ein Re-Render baut das Modal neu und **verwirft getippte Notizen**. Er ersetzt
  nur das `innerHTML` des Kontakt-Selects (gleiches Muster wie der `firmSelect`-Handler im
  Kontakt-Formular). `data-keep` trägt die ID eines evtl. archivierten Kontakts.

## Fachlogik

- **Erledigte Aufgaben gehören NIE in den „Verlauf".** Sie werden nach *Deadline* sortiert, die
  in der Zukunft liegen kann → „in 2 Tagen" im Verlauf wäre Unsinn. Sie stehen im Bucket
  `akt-c-done`.
- **Aufgaben ohne Termin** fallen durch alle Fälligkeits-Buckets. Zustand = „Beobachten"
  (`pflegePredicate("offen")`), in der Agenda eigener Bucket `akt-c-undated`.
  `helpers.isOverdue("")` liefert korrekt `false` — eine unterminierte Aufgabe ist **nicht**
  überfällig, sondern unterminiert. Nicht „reparieren".
- **Abdeckungs-Bänder sind überschneidungsfrei** (≤6 / 6–12 / ohne), Summe = alle Kunden.
  Nicht kumulativ machen, sonst summiert der Balken nicht auf 100%.

---

# Tote Zonen / Backlog

- **Erledigt:** `views.planning()`, `views.historyView()` und ihr Anhang sind **gelöscht**
  (725 Zeilen, 6638 → 5913). Mit weg: `filters.planning`, `filters.history`,
  `CONFIG.defaults.planningShowOnlyOpen`, `helpers.firmSignal`, `firmMatchesLens`,
  `sparklineHtml`, `momentumHeatmapHtml`, `periodKey`, `periodLabel`, `ui.miniItem` und die
  Handler `navigate-planning(-filtered)`, `history-firma-filter`, `history-view-mode`,
  `filter-lens/-periode/-granularitaet`, `toggle-expand`, `task-status-change`.
  Geblieben sind **Redirect und `knownRoutes`** — bewusst.
- **Totes CSS in `index.html`** (nur die gelöschten Views nutzten es, noch nicht entfernt):
  `.bbz-actbar`, `.bbz-actbar-fill`, `.bbz-on`, `.bbz-planning-filters`, `.bbz-filters-3/-4`,
  `.bbz-history-group*`, `.bbz-history-split`, `.bbz-mini-list`, `.bbz-timeline-clamp`,
  `.bbz-timeline-item`/`.bbz-expanded`. Bewusst separat: das Stylesheet wird beim nächsten
  Schritt (Design-Angleichung Aktivitäten ↔ Dashboard) ohnehin angefasst.
- **Handler ohne Renderer** (Altlast, nicht aus dieser Löschung): `akt-legende` — der
  Legenden-Toggle der Aktivitäten-View wird nirgends gerendert,
  `filters.aktivitaeten.legendeOffen` ist damit wirkungslos.
- **`deploy.yml`** — Node-20-Deprecation der Actions (`actions/*@v4` behebt es).
- **Route `birthdays`** hängt an einem einzigen Link — Nav-Eintrag wäre ehrlicher.
- **`admin.userStats`** (Erfassungen pro Person) ist gebaut, aber nirgends im Dashboard —
  bewusst offen gelassen (sichtbare Leistungsmessung pro Person braucht eine Entscheidung).
- **Der Zustand `offen` / „Beobachten" feuert aktuell bei NULL Firmen.** Er ist als „offene
  Aufgabe ohne Datum" definiert (`pflegePredicate("offen")`), und in CRMTasks hat **jede**
  der 28 Aufgaben eine Deadline. Chip, Punkt, Signal-Filter und Bucket `akt-c-undated`
  existieren für einen Fall, den es im Datenbestand nicht gibt. Nicht löschen ohne
  Entscheidung — er ist billig und fängt den Fall ab, sobald er auftritt.
- **Totes CSS in `index.html`** (nur die gelöschten planning/history-Views nutzten es):
  `.bbz-actbar*`, `.bbz-on`, `.bbz-planning-filters`, `.bbz-filters-3/-4`,
  `.bbz-history-group*`, `.bbz-history-split`, `.bbz-mini-*`, `.bbz-timeline-clamp`.
  Sie hängen in Selektorlisten und Media-Query-Blöcken — Entfernen berührt fünf Stellen im
  `<style>`-Block. Eigener Commit, sonst Klammerbilanz-Risiko ohne Funktionsgewinn.
- **KPI-Aggregations-Helper** sind weiter vorgehalten.

## Datenbefunde (Stand letzte Sichtung)

Keine Code-Themen — aber sie erklären, warum Kennzahlen so aussehen, wie sie aussehen:

- **87 von 125 Banken (70%) ohne Aktivität in 12 Monaten.** Kein Datenartefakt: es gibt
  **keine** Firmen ohne Kontakte. Echte Pflegelücke.
- **Die App ging Februar 2026 produktiv.** Vorher liegen fast keine Datensätze (Okt 25: 1,
  Jan 26: 2). Jede Kennzahl über „12 Monate" mittelt gegen diese Leere und ist zu tief.
  Stand Juli 2026: **68 Aktivitäten**, davon 61 ab Februar → **9,5/Mt.**, nicht 5,3.
  Der alte Befund „1,2 pro Woche" beruhte auf dem 12-Monats-Schnitt und ist damit erledigt;
  aktuell sind es rund **2,3/Woche**. Ob das Erfassen der Engpass bleibt, misst Zone 3.
- **Rund 41% der Kunden ohne Klassifizierung — das ist ABSICHT**, nicht Datenmangel: sie
  werden passiv betreut. In Zone 1 ist das die Karte `Ohne`, gleichwertig neben A/B/C.
  Kein ⚠ dafür. Wer das als Lücke darstellt, warnt vor einer Entscheidung des Users.
- **32% der Kontakte ohne Geburtstag** — der Kalender ist strukturell unvollständig.
