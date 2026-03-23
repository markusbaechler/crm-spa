// --- 1. KONFIGURATION ---
const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    // Stabiler Pfad-Weg statt fehleranfälliger langer Site-ID
    sitePath: "bbzsg.sharepoint.com:/sites/CRM", 
    lists: {
        firms: "c763d04e-1d8d-484a-ad67-6f77f7bc9d92",
        contacts: "58e7e24d-2ea2-4313-a792-18c724ce7924"
    }
};

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: "https://login.microsoftonline.com/" + config.tenantId,
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    },
    cache: { 
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true 
    }
};

// --- 2. INITIALISIERUNG ---
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Verarbeitet die Rückkehr von der Microsoft-Login-Seite
msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        console.log("Login erfolgreich zurückgekehrt");
        updateUI();
        loadFirms(); // Lädt direkt die Firmen nach dem Login
    }
}).catch(err => {
    console.error("Redirect Fehler:", err);
});

// --- 3. AUTHENTIFIZIERUNG ---
function handleAuth() {
    const loginRequest = {
        scopes: ["https://graph.microsoft.com/Sites.Read.All"]
    };
    // Nutzt Redirect (stabiler als Popup für GitHub/SharePoint)
    msalInstance.loginRedirect(loginRequest);
}

function updateUI() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Eingeloggt ✅";
        authBtn.classList.replace('bg-blue-600', 'bg-green-600');
    }
}

// --- 4. DATEN LADEN (CRUD - READ) ---
async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = `
            <div class="text-center p-10">
                <h3 class="text-lg font-bold mb-4 text-slate-700">Bitte melden Sie sich an</h3>
                <button onclick="handleAuth()" class="bg-blue-600 text-white px-8 py-3 rounded-lg shadow-lg font-bold hover:bg-blue-700 transition">
                    Jetzt Login mit Microsoft
                </button>
            </div>`;
        return;
    }

    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-500 font-bold text-lg">Verbindung zu SharePoint wird hergestellt...</p>';

    try {
        // Token im Hintergrund holen
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(() => {
            // Falls Token abgelaufen, neu einloggen via Redirect
            return msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] });
        });

        if (!tokenRes) return;

        // Abfrage-URL mit Site-Pfad
        const url = `https://graph.microsoft.com/v1.0/sites/${config.sitePath}/lists/${config.lists.firms}/items?expand=fields`;
        
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();

        // Fehlerbehandlung für die API-Antwort
        if (data.error) {
            content.innerHTML = `
                <div class="p-6 bg-orange-50 border border-orange-200 rounded-lg">
                    <h3 class="text-orange-800 font-bold mb-2">Microsoft Graph meldet ein Problem:</h3>
                    <pre class="text-xs bg-white p-4 border overflow-auto">${JSON.stringify(data.error, null, 2)}</pre>
                    <button onclick="location.reload()" class="mt-4 bg-orange-600 text-white px-4 py-2 rounded text-sm">Seite neu laden</button>
                </div>`;
            return;
        }

        // Firmenliste anzeigen
        const firms = data.value || [];
        if (firms.length === 0) {
            content.innerHTML = '<p class="p-10 text-center text-gray-500 italic">Die Liste CRMFirms enthält momentan keine Einträge.</p>';
            return;
        }

        let html = `
            <div class="flex justify-between items-center mb-6">
                <h2 class="text-2xl font-black text-slate-800">🏢 Firmenverzeichnis</h2>
                <span class="text-sm text-slate-400">${firms.length} Einträge</span>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">`;
        
        firms.forEach(item => {
            const f = item.fields || {};
            const klasse = f.Klassifizierung || '-';
            const klasseFarbe = klasse === 'A' ? 'text-green-600' : (klasse === 'B' ? 'text-orange-500' : 'text-slate-400');
            
            html += `
                <div class="p-5 bg-white border border-slate-100 rounded-xl shadow-sm hover:shadow-md hover:border-blue-200 transition group">
                    <div class="font-bold text-lg text-slate-800 group-hover:text-blue-600 transition">${f.Title || 'Unbenannt'}</div>
                    <div class="mt-2 flex items-center text-sm">
                        <span class="text-slate-400 mr-2">Klassifizierung:</span>
                        <span class="font-black ${klasseFarbe}">${klasse}</span>
                    </div>
                </div>`;
        });
        
        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `
            <div class="p-6 bg-red-50 border border-red-200 rounded-lg text-red-700">
                <p class="font-bold">Verbindungsfehler:</p>
                <p class="text-sm">${err.message}</p>
            </div>`;
    }
}

// --- 5. NAVIGATION & START ---
window.onload = () => {
    updateUI();
};

function showView(v) { 
    if(v === 'dashboard') {
        location.reload();
    } else if(v === 'contacts') {
        document.getElementById('app-content').innerHTML = `
            <h2 class="text-2xl font-bold mb-4">Kontakte</h2>
            <p class="text-slate-500 italic">Diese Funktion wird in Etappe D.2 freigeschaltet, sobald die Firmenliste stabil läuft.</p>`;
    }
}
