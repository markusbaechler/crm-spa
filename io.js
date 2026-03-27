// io.js — Import/Export Modul
// Greift auf window._bbzApp (app.js) zu — startet keine eigene App-Instanz

(() => {
  "use strict";

  // Warten bis app.js bereit ist
  function waitForApp(cb, attempts = 0) {
    if (window._bbzApp) {
      cb(window._bbzApp);
    } else if (attempts < 50) {
      setTimeout(() => waitForApp(cb, attempts + 1), 100);
    } else {
      console.warn("bbzIO: _bbzApp nicht gefunden nach 5s");
    }
  }

  const io = {
    // Export: Kontakte als CSV
    openExportModal(type) {
      waitForApp(({ state, helpers }) => {
        const contacts = state.enriched.contacts;
        if (!contacts?.length) {
          alert("Keine Kontakte zum Exportieren vorhanden.");
          return;
        }

        let rows, filename, headers;

        if (type === "event") {
          // Event-Export: Kontakte mit Event-Feld
          const eventContacts = contacts.filter(c => c.event?.length > 0);
          headers = ["Nachname", "Vorname", "Firma", "Event", "Segment", "Email", "Mobile"];
          rows = eventContacts.map(c => [
            c.nachname || "",
            c.vorname || "",
            c.firmTitle || "",
            (c.event || []).join("; "),
            c.klassifizierung || "",
            c.email1 || "",
            c.mobile || ""
          ]);
          filename = "bbz-events.csv";
        } else {
          // Kontakt-Export
          headers = ["Nachname", "Vorname", "Anrede", "Firma", "Funktion", "Rolle",
                     "Email1", "Email2", "Direktwahl", "Mobile", "Leadbbz", "SGF",
                     "Geburtstag", "Kommentar", "Archiviert"];
          rows = contacts.map(c => [
            c.nachname || "",
            c.vorname || "",
            c.anrede || "",
            c.firmTitle || "",
            c.funktion || "",
            c.rolle || "",
            c.email1 || "",
            c.email2 || "",
            c.direktwahl || "",
            c.mobile || "",
            c.leadbbz0 || "",
            (c.sgf || []).join("; "),
            c.geburtstag ? helpers.toDateInput(c.geburtstag) : "",
            c.kommentar || "",
            c.archiviert ? "Ja" : "Nein"
          ]);
          filename = "bbz-kontakte.csv";
        }

        // CSV erstellen
        const escape = v => `"${String(v).replace(/"/g, '""')}"`;
        const csv = [headers, ...rows].map(r => r.map(escape).join(",")).join("\n");
        const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
      });
    },

    // Import: CSV/XLSX Kontakte (Stub — zeigt Hinweis)
    openImportModal() {
      waitForApp(() => {
        alert("Import-Funktion: Bitte CSV-Datei mit Spalten Nachname, Vorname, Firma etc. vorbereiten.\n\nFull Import-UI wird in v1.4.0 implementiert.");
      });
    }
  };

  window.bbzIO = io;

})();
