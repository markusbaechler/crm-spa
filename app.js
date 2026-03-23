const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteId: "bbzsg.sharepoint.com,14cdc8c1-0ddd-44e5-c-adad-fba775e44771,f504693c-96ae-4773-a6b9-ae3f91a6db00",
    lists: {
        firms: "c763d04e-1d8d-484a-ad67-6f77f7bc9d92"
    }
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

// WICHTIG: Diese Funktion verarbeitet die Rückkehr von Microsoft
msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        console.log("Login via Redirect erfolgreich");
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
        loadFirms(); // Lädt die Firmen direkt nach Rückkehr
    }
}).catch(err => {
    console.error("Redirect Fehler:", err);
});

async function handleAuth() {
    // Wir nutzen jetzt loginRedirect statt loginPopup
    const loginRequest = {
        scopes: ["https://graph.microsoft.com/Sites.Read.All"]
    };
    msalInstance.loginRedirect(loginRequest);
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = '<div class="text-center p-10"><button onclick="handleAuth()" class="bg-blue-600 text-white px-6 py-2 rounded shadow-lg">Login erforderlich</button></div>';
        return;
    }

    content.innerHTML = '<p class="p-6 text-center animate-pulse">Lade Firmenliste...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(err => {
            // Falls Silent fehlschlägt, wieder Redirect nutzen
            return msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] });
        });

        if (!tokenRes) return; // Warten auf Redirect

        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields(select=Title,Klassifizierung)`;
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();
        
        let html = '<h2 class="text-xl font-bold mb-4">🏢 Firmenliste</h2><div class="space-y-2">';
        data.value.forEach(item => {
            html += `<div class="p-3 bg-white border rounded shadow-sm flex justify-between hover:bg-blue-50 transition cursor-default">
                        <span class="font-semibold text-slate-800">${item.fields.Title || 'Unbekannt'}</span>
                        <span class="bg-blue-100 text-blue-700 px-2 py-1 rounded text-xs font-bold">${item.fields.Klassifizierung || '-'}</span>
                     </div>`;
        });
        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-orange-50 text-orange-700 border rounded">Fehler: ${err.message}</div>`;
    }
}

// Navigation
function showView(v) { 
    if(v === 'dashboard') location.reload();
}

// Button Status prüfen
if(msalInstance.getAllAccounts().length > 0) {
    document.getElementById('authBtn').innerText = "Eingeloggt ✅";
}
