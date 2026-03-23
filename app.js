const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteId: "bbzsg.sharepoint.com,14cdc8c1-0ddd-44e5-c-adad-fba775e44771,f504693c-96ae-4773-a6b9-ae3f91a6db00",
    lists: {
        firms: "c763d04e-1d8d-484a-ad67-6f77f7bc9d92",
        contacts: "58e7e24d-2ea2-4313-a792-18c724ce7924"
    }
};

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    },
    cache: {
        cacheLocation: "sessionStorage" // Speichert den Login im Browser-Tab
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function handleAuth() {
    try {
        await msalInstance.loginPopup({ scopes: ["Sites.Read.All"] });
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
        location.reload(); // Seite neu laden, um Login-Status zu fixieren
    } catch (err) {
        console.error("Login Fehler:", err);
    }
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    content.innerHTML = '<p class="p-4 text-blue-600 animate-pulse">Verbindung zu SharePoint wird aufgebaut...</p>';

    try {
        // Versuche erst, den Account zu finden
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) throw new Error("Kein Account gefunden");

        const authResult = await msalInstance.acquireTokenSilent({
            scopes: ["Sites.Read.All"],
            account: accounts[0]
        });

        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields(select=Title,Klassifizierung)`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${authResult.accessToken}` }
        });
        
        const result = await response.json();
        const firms = result.value || [];

        let html = `<div class="flex justify-between items-center mb-6">
                        <h2 class="text-2xl font-bold text-slate-800">🏢 Firmenübersicht</h2>
                        <span class="bg-blue-100 text-blue-800 text-xs font-semibold px-2.5 py-0.5 rounded">${firms.length} Einträge</span>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left">
                            <thead>
                                <tr class="border-b-2 border-slate-100 text-slate-400 text-sm uppercase tracking-wider">
                                    <th class="p-3">Firma</th>
                                    <th class="p-3">Klassierung</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-50">`;
        
        firms.forEach(f => {
            const klasse = f.fields.Klassifizierung || '-';
            const farbe = klasse === 'A' ? 'text-green-600 font-bold' : (klasse === 'B' ? 'text-orange-500' : 'text-slate-400');
            
            html += `<tr class="hover:bg-slate-50 transition">
                <td class="p-3 font-medium text-slate-700">${f.fields.Title || 'Unbenannt'}</td>
                <td class="p-3 ${farbe}">${klasse}</td>
            </tr>`;
        });
        html += `</tbody></table></div>`;
        content.innerHTML = html;

    } catch (err) {
        console.error(err);
        content.innerHTML = `
            <div class="text-center p-8 bg-red-50 rounded-lg">
                <p class="text-red-600 font-semibold mb-2">Zugriff fehlgeschlagen</p>
                <p class="text-sm text-red-500 mb-4">${err.message}</p>
                <button onclick="handleAuth()" class="bg-red-600 text-white px-4 py-2 rounded shadow">Erneut anmelden</button>
            </div>`;
    }
}

// Hilfsfunktion für Navigation
function showView(view) {
    const content = document.getElementById('app-content');
    if(view === 'dashboard') content.innerHTML = '<h2 class="text-2xl font-bold">Dashboard</h2><p class="mt-4">Willkommen im CRM bbz Light!</p>';
    if(view === 'contacts') content.innerHTML = '<h2 class="text-2xl font-bold text-blue-600">Kontakte</h2><p class="mt-4 italic text-slate-400 font-light">Diese Funktion wird in der nächsten Etappe freigeschaltet.</p>';
}
