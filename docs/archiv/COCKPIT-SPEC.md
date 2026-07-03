# Sprint-Spec: Aktivitäten -> Cockpit

Umbau der `history`-Route von retrospektivem Logbuch zu handlungsorientiertem Cockpit.
Ziel: "Was tue ich jetzt, wie stehe ich da, schnell handeln." Timeline und Pflege-Radar
bleiben erhalten und werden ergänzt.

Betroffen: `views.historyView()` (app.js ab ~Z. 3852), `state.filters.history`,
Event-Delegation (`data-action`-Handler), evtl. neue Helper. Keine neuen Dependencies,
kein Framework, bestehende `bbz-*`-Klassen + CSS-Tokens verwenden.

## State-Erweiterung
`state.filters.history` ergänzen:
- `lens: ""` — aktiver Lead-BBZ-Wert (leer = alle).
- `wochenziel: 30` — Aktivitäten-Wochenziel (später konfigurierbar).

## 1. Linse (persönlicher Filter)
Leiste oben: Dropdown "Lead BBZ" (Werte aus `[...new Set(history.map(h=>h.leadbbz))]`)
+ Zeitraum-Dropdown (bestehende `zeitfenster`-Logik wiederverwenden).
Wenn `lens` gesetzt: alle Cockpit-Kennzahlen, Next-Best-Actions und Timeline
zusätzlich auf `leadbbz === lens` filtern. `data-filter="history-lens"`.

## 2. Instrumenten-Band (ersetzt aktuelles KPI-Grid oben)
Fünf Kacheln im `bbz-kpi`-Stil:
1. **Heute fällig** — offene Tasks mit `deadline <= heute` (accent).
2. **Überfällig** — offene Tasks mit `isOverdue` (rot).
3. **Aktivitäten Woche** — Count letzte 7 Tage + `+N vs. Vorwoche` + Mini-Sparkline (inline SVG polyline).
4. **A/B on track** — `% der A/B-Firmen mit firmSignal==="ok"`.
5. **Wochenziel** — `<aktuell>/<ziel>` + 4px-Fortschrittsbalken.
Alle Werte durch `lens` gefiltert, wenn gesetzt.

## 3. Handlungszentrum (Next Best Actions) — KERN
Eine priorisierte Liste aus vorhandenen Ableitungen. Reihenfolge:
1. Überfällige offene Tasks (rot) — Aktion: "Erledigt" + "Öffnen"
2. A-Kunden mit `firmSignal==="never"` (rot) — Aktion: "Log Aktivität"
3. Eingeschlafen `firmSignal==="cold"` (amber) — Aktion: "Log Aktivität"
4. Heute fällige offene Tasks (accent) — Aktion: "Erledigt"
Pro Zeile: Signal-Punkt, Kontakt/Firma, Grund-Text, 1-Klick-Aktion.
Bestehende Radar-Berechnungen (`radarNever`, `radarCold`, `radarOverdue`) und Task-Felder
wiederverwenden — nur neu zu einer Queue zusammenführen. Auf `lens` filtern.
Aktionen über bestehende `data-action`s: `open-history-form` (data-firm-id),
`open-task-form`, plus neuer `complete-task` (data-id) -> PATCH Status erledigt + render.

## 4. Schnellerfassung
Über der Timeline eine Leiste: Buttons Anruf / Mail / Meeting / Notiz.
Klick öffnet bestehendes History-Modal (`open-history-form`) mit vorbelegtem
`typ`/Kontaktart via `data-typ`.

## 5. Timeline + Radar (bleiben)
- Timeline: unverändert (`renderCard`, Datums-/Firmen-Gruppierung), nur unter die Schnellerfassung schieben.
- Pflege-Radar: bleibt, zusätzlich Momentum-Heatmap (12 Wochen, je Woche ein Kästchen,
  Intensität aus Aktivitäts-Count). Bestehende `radarZoneHtml`-Zonen behalten.

## Akzeptanzkriterien
- [ ] Lens-Filter wirkt auf Band, Handlungszentrum und Timeline.
- [ ] Handlungszentrum zeigt priorisierte, klickbare Aktionen; "Erledigt" schreibt Task-Status per PATCH und rendert neu.
- [ ] Instrumenten-Band zeigt Heute-fällig / Überfällig / Woche+Trend / A-B-Score / Wochenziel.
- [ ] Timeline und Pflege-Radar funktionieren unverändert weiter.
- [ ] Alle Werte `escapeHtml`-gesichert, nur bestehende CSS-Tokens/Klassen.
- [ ] Mobil nutzbar (bestehende `bbz-history-tab-bar`-Umschaltung beibehalten/anpassen).
- [ ] `node --check app.js` grün; lokal geladen und `history`-Route klick-getestet.

## Vorgehen
1. `state.filters.history` erweitern.
2. `historyView()` refaktorieren: Band -> Handlungszentrum -> Schnellerfassung+Timeline / Radar.
3. `complete-task`-Handler + ggf. Sparkline/Heatmap-Helper ergänzen.
4. Lokal testen, ein Commit `feat: Aktivitäten-Cockpit`, push -> Auto-Deploy.
