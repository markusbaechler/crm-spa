// --- 1. KONFIGURATION ---
const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    // Geänderter Pfad: Ohne Doppelpunkt, das ist bei v1.0 oft die Fehlerquelle
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

// Verarbeitet die Rückkehr von Microsoft nach dem Login
msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        console.log("Login erfolgreich beendet");
        updateUI();
        loadFirms(); 
    }
}).catch(err => {
    console.error("Fehler beim Seiten-Rücksprung:", err);
});

// --- 3. LOGIN FUNKTION ---
function handleAuth() {
    const loginRequest = {
        scopes: ["https://graph.microsoft.com/Sites.Read.All"]
    };
    // Wir verlassen die Seite kurz für den Microsoft-Login
    msalInstance.loginRedirect(loginRequest);
}

function updateUI() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Eingeloggt ✅";
        authBtn.classList.remove('bg-blue-600');
        authBtn.classList.add('bg-green-600');
    }
}

// --- 4. HAUPTFUNKTION: FIRMEN LADEN ---
async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    // Falls nicht eingeloggt
    if (accounts.length === 0) {
        content.innerHTML = `
            <div class="text-center p-10">
                <h3 class="text-xl font-bold mb-4 text-slate-800">Willkommen beim bbz CRM</h3>
                <p class="mb-6 text-slate-500">Bitte loggen Sie sich ein, um auf die SharePoint-Daten zuzugreifen.</p>
                <button onclick="handleAuth()" class="bg-blue-600 text-white px-8 py-3 rounded-lg shadow-lg font-bold hover:bg-blue-700 transition">
                    Mit Microsoft-Konto anmelden
                </button>
            </div>`;
        return;
    }

    content.innerHTML = '<div class="p-10 text-center"><div class="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600 mb-4"></div><p class="text-blue-600 font-bold">Lade Firmenliste aus SharePoint...</p></div>';

    try {
        // Token für den Zugriff holen
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(() => {
            return msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] });
        });

        if (!tokenRes) return;

        // Microsoft Graph Abfrage
        // Wir probieren hier die sicherste URL-Variante
        const url = `https://graph.microsoft.com/v1.0/sites/${config.sitePath}/lists/${config.lists.firms}/items?expand=fields`;
        
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();

        // Fehler von Microsoft anzeigen (wie die orange Box)
        if (data.error) {
            content.innerHTML = `
                <div class="p-6 bg-orange-50 border border-orange-200 rounded-lg">
                    <h3 class="text-orange-800 font-bold mb-2">Hinweis von Microsoft:</h3>
                    <p class="text-sm mb-4">${data.error.message}</p>
                    <pre class="text-xs bg-white p-4 border overflow-auto max-h-40">${JSON.stringify(data.error, null, 2)}</pre>
                    <button onclick="location.reload()" class="mt-4 bg-slate-800 text-white px-4 py-2 rounded text-sm font-bold">Erneut versuchen</button>
                </div>`;
            return;
        }

        const firms = data.value || [];
        if (firms.length === 0) {
            content.innerHTML = '<div class="p-10 text-center text-slate-500 italic border-2 border-dashed rounded-xl">Die Liste "CRMFirms" ist aktuell leer.</div>';
            return;
        }

        // HTML Tabelle/Liste bauen
        let html = `
            <div class="flex justify-between items-center mb-6">
                <h2 class="text-2xl font-black text-slate-800 tracking-tight">🏢 Firmenverzeichnis</h2>
                <span class="bg-slate-100 text-slate-600 text-xs font-bold px-3 py-1 rounded-full uppercase italic">${firms.length} Firmen</span>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">`;
        
        firms.forEach(item => {
            const f = item.fields || {};
            const klasse = f.Klassifizierung || '-';
            
            // Farbe basierend auf Klassifizierung
            let klasseStyle = "bg-slate-100 text-slate-500";
            if(klasse === 'A') klasseStyle = "bg-green-100 text-green-700";
            if(klasse === 'B') klasseStyle = "bg-blue-100 text-blue-700";
            if(klasse === 'C') klasseStyle = "bg-orange-100 text-orange-700";

            html += `
                <div class="p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-blue-300 transition-all duration-200 group cursor-default">
                    <div class="font-bold text-lg text-slate-800 group-hover:text-blue-600 transition-colors">${f.Title || 'Unbenannt'}</div>
                    <div class="mt-4 flex items-center justify-between">
                        <span class="text-xs text-slate-400 font-semibold uppercase tracking-widest">Klassifizierung</span>
                        <span class="px-3 py-1 rounded-lg text-xs font-black ${klasseStyle}">${klasse}</span>
                    </div>
                </div>`;
        });
        
        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `
            <div class="p-6 bg-red-50 border border-red-200 rounded-xl text-red-700 shadow-inner">
                <p class="font-bold mb-1 underline">Technischer Fehler:</p>
                <p class="text-sm font-mono">${err.message}</p>
            </div>`;
    }
}

// --- 5. INITIALER CHECK ---
window.onload = () => {
    updateUI();
};

function showView(v) { 
    if(v === 'dashboard') location.reload();
    if(v === 'contacts') {
        document.getElementById('app-content').innerHTML = '<div class="p-10 text-center"><h2 class="text-2xl font-bold text-slate-800">Kontakte</h2><p class="text-slate-500 mt-2 italic underline decoration-blue-500">Funktion folgt nach stabilen Firmen-Tests.</p></div>';
    }
}
