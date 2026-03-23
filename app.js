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

// Redirect-Ergebnis verarbeiten
msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
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
        content.innerHTML = '<button onclick="handleAuth()" class="bg-blue-600 text-white p-4 rounded shadow-lg">Bitte Login klicken</button>';
        return;
    }

    content.innerHTML = '<p class="p-6 text-center animate-pulse">SharePoint-Daten werden gelesen...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(() => msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] }));

        if (!tokenRes) return;

        // Wir holen die Rohdaten
        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields`;
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();
        
        // SICHERHEITS-CHECK: Sind Daten da?
        const firms = data.value || []; 

        if (firms.length === 0) {
            content.innerHTML = '<div class="p-6 text-orange-600 bg-orange-50 border rounded text-center">Verbindung erfolgreich, aber die SharePoint-Liste "CRMFirms" ist momentan leer.</div>';
            return;
        }

        let html = '<h2 class="text-2xl font-bold mb-6 text-slate-800 border-b pb-2">🏢 Firmenverzeichnis</h2><div class="grid gap-3">';
        
        firms.forEach(item => {
            const f = item.fields || {};
            html += `<div class="p-4 bg-white border rounded shadow-sm flex justify-between items-center hover:bg-slate-50">
                        <span class="font-bold text-slate-700">${f.Title || 'Unbenannt'}</span>
                        <span class="bg-blue-100 text-blue-800 px-3 py-1 rounded-full text-xs font-bold">${f.Klassifizierung || '-'}</span>
                     </div>`;
        });
        
        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-red-50 text-red-700 border rounded">Technischer Fehler: ${err.message}</div>`;
    }
}

// Button-Status
window.onload = () => {
    if(msalInstance.getAllAccounts().length > 0) {
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
    }
};

function showView(v) { if(v === 'dashboard') location.reload(); }
