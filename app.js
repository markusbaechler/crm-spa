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

// Verarbeitet die Rückkehr vom Login
msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        console.log("Login erfolgreich");
        loadFirms();
    }
});

function handleAuth() {
    msalInstance.loginRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] });
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = '<button onclick="handleAuth()" class="bg-blue-600 text-white p-4 rounded">Bitte einloggen</button>';
        return;
    }

    content.innerHTML = '<p class="p-6 animate-pulse">Rufe Daten ab...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(() => msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] }));

        if (!tokenRes) return;

        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields`;
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();

        // --- DIAGNOSE-BLOCK START ---
        // Wenn wir hier landen und "value" fehlt, zeigen wir die Rohdaten an
        if (!data.value) {
            content.innerHTML = `
                <div class="p-4 bg-orange-100 border-l-4 border-orange-500 text-orange-700 text-xs">
                    <p class="font-bold mb-2 text-sm">Warnung: Microsoft hat keine Liste gesendet.</p>
                    <p>Antwort von Microsoft:</p>
                    <pre class="bg-white p-2 mt-2 overflow-auto max-h-60 border">${JSON.stringify(data, null, 2)}</pre>
                </div>`;
            return;
        }
        // --- DIAGNOSE-BLOCK ENDE ---

        let html = '<h2 class="text-xl font-bold mb-4">🏢 Firmen</h2>';
        data.value.forEach(item => {
            const f = item.fields || {};
            html += `<div class="p-2 border-b font-medium text-slate-700">${f.Title || 'Kein Name'}</div>`;
        });
        content.innerHTML = html;

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-red-100 text-red-700 font-bold">Fehler: ${err.message}</div>`;
    }
}

// Button-Status oben rechts
window.onload = () => {
    if(msalInstance.getAllAccounts().length > 0) {
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
    }
};

function showView(v) { if(v === 'dashboard') location.reload(); }
