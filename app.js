
// 1. Deine Zugangsdaten (Das "Navi" für die App)
const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteId: "bbzsg.sharepoint.com,14cdc8c1-0ddd-44e5-c-adad-fba775e44771,f504693c-96ae-4773-a6b9-ae3f91a6db00",
    lists: {
        firms: "c763d04e-1d8d-484a-ad67-6f77f7bc9d92",
        contacts: "58e7e24d-2ea2-4313-a792-18c724ce7924"
    }
};

// 2. Microsoft Login initialisieren
const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`
    }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);

// 3. Funktion: Einloggen
async function handleAuth() {
    try {
        await msalInstance.loginPopup({ scopes: ["Sites.Read.All"] });
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
        console.log("Login erfolgreich");
    } catch (err) {
        alert("Fehler beim Login: " + err.message);
    }
}

// 4. Funktion: Firmen aus SharePoint laden
async function loadFirms() {
    const content = document.getElementById('app-content');
    content.innerHTML = '<p class="p-4">Lade Firmenliste...</p>';

    try {
        const authResult = await msalInstance.acquireTokenSilent({ scopes: ["Sites.Read.All"] });
        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields(select=Title,Klassifizierung)`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${authResult.accessToken}` }
        });
        const result = await response.json();
        const firms = result.value || [];

        // Tabelle anzeigen
        let html = `<h2 class="text-xl font-bold mb-4">Firmen im System</h2>
                    <table class="w-full text-left border">
                        <tr class="bg-gray-200">
                            <th class="p-2">Name</th>
                            <th class="p-2">Klassierung</th>
                        </tr>`;
        
        firms.forEach(f => {
            html += `<tr>
                <td class="p-2 border-t">${f.fields.Title || 'Kein Name'}</td>
                <td class="p-2 border-t">${f.fields.Klassifizierung || '-'}</td>
            </tr>`;
        });
        html += `</table>`;
        content.innerHTML = html;

    } catch (err) {
        content.innerHTML = `<p class="text-red-500">Bitte erst einloggen!</p>`;
    }
}

// Hilfsfunktion für die Navigation
function showView(view) {
    const content = document.getElementById('app-content');
    if(view === 'dashboard') content.innerHTML = '<h2 class="text-2xl">Dashboard</h2><p>Willkommen zurück!</p>';
    if(view === 'contacts') content.innerHTML = '<h2 class="text-2xl">Kontakte</h2><p>Funktion folgt in Etappe D.2.</p>';
}
