const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM"
};

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: "https://login.microsoftonline.com/" + config.tenantId,
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    },
    cache: { cacheLocation: "localStorage" }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Verarbeitet die Rückkehr vom Login
msalInstance.handleRedirectPromise().then(response => {
    if (response) { loadFirms(); }
});

function handleAuth() {
    msalInstance.loginRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.Write.All"] });
}

// HAUPTFUNKTION: FIRMEN ANZEIGEN
async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = '<div class="text-center p-10"><button onclick="handleAuth()" class="bg-blue-600 text-white px-6 py-2 rounded shadow-lg font-bold">Zuerst Login bestätigen</button></div>';
        return;
    }

    content.innerHTML = '<p class="p-6 text-center animate-pulse">Lade Firmen...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.Write.All"],
            account: accounts[0]
        }).catch(() => msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.Write.All"] }));

        if (!tokenRes) return;

        // Site-ID finden
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });
        const siteData = await siteRes.json();
        const realSiteId = siteData.id;

        // Liste finden
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${realSiteId}/lists`, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");

        // Items laden
        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${realSiteId}/lists/${targetList.id}/items?expand=fields`, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });
        const itemsData = await itemsRes.json();

        // UI: Header mit "Neu"-Button
        let html = `
            <div class="flex justify-between items-center mb-6">
                <h2 class="text-2xl font-bold text-slate-800 tracking-tight">🏢 Firmenliste</h2>
                <button onclick="showAddForm('${realSiteId}', '${targetList.id}')" class="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg font-bold shadow transition">
                    + Neue Firma
                </button>
            </div>
            <div id="form-container" class="hidden mb-8 p-6 bg-slate-50 border-2 border-dashed border-slate-200 rounded-xl"></div>
            <div class="grid gap-3">`;
        
        itemsData.value.forEach(item => {
            html += `<div class="p-4 bg-white border rounded-xl shadow-sm flex justify-between items-center">
                        <span class="font-bold text-slate-700">${item.fields.Title || 'Unbenannt'}</span>
                        <span class="bg-blue-100 text-blue-700 px-3 py-1 rounded-full text-xs font-black">${item.fields.Klassifizierung || '-'}</span>
                     </div>`;
        });

        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `<p class="p-6 text-red-600 font-bold italic">Fehler: ${err.message}</p>`;
    }
}

// FUNKTION: FORMULAR ANZEIGEN
function showAddForm(siteId, listId) {
    const container = document.getElementById('form-container');
    container.classList.toggle('hidden');
    container.innerHTML = `
        <h3 class="font-bold mb-4">Neue Firma erfassen</h3>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <input type="text" id="newFirmName" placeholder="Firmenname" class="p-2 border rounded">
            <select id="newFirmClass" class="p-2 border rounded">
                <option value="">Klassifizierung wählen</option>
                <option value="A">A (Top)</option>
                <option value="B">B (Normal)</option>
                <option value="C">C (Passiv)</option>
            </select>
        </div>
        <div class="mt-4 flex space-x-2">
            <button onclick="saveFirm('${siteId}', '${listId}')" class="bg-blue-600 text-white px-4 py-2 rounded font-bold">Speichern</button>
            <button onclick="this.parentElement.parentElement.classList.add('hidden')" class="bg-gray-300 px-4 py-2 rounded">Abbrechen</button>
        </div>
    `;
}

// FUNKTION: DATEN AN SHAREPOINT SENDEN
async function saveFirm(siteId, listId) {
    const name = document.getElementById('newFirmName').value;
    const klassifizierung = document.getElementById('newFirmClass').value;

    if (!name) { alert("Bitte Namen eingeben!"); return; }

    try {
        const accounts = msalInstance.getAllAccounts();
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.Write.All"],
            account: accounts[0]
        });

        const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
        
        const payload = {
            fields: {
                Title: name,
                Klassifizierung: klassifizierung
            }
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + tokenRes.accessToken,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            alert("Firma erfolgreich gespeichert!");
            loadFirms(); // Liste aktualisieren
        } else {
            const errData = await response.json();
            alert("Fehler: " + errData.error.message);
        }

    } catch (err) {
        alert("Fehler beim Speichern: " + err.message);
    }
}

function showView(v) { if(v === 'dashboard') location.reload(); }
