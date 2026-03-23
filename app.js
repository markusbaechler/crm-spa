const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    // Wir lassen die App die Site-ID selbst suchen:
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

msalInstance.handleRedirectPromise().then(response => {
    if (response) { loadFirms(); }
});

function handleAuth() {
    msalInstance.loginRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] });
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = '<button onclick="handleAuth()" class="bg-blue-600 text-white p-4 rounded">Login starten</button>';
        return;
    }

    content.innerHTML = '<p class="p-6 animate-pulse text-blue-500 font-bold">🔍 Starte System-Diagnose...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(() => msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] }));

        if (!tokenRes) return;

        // SCHRITT 1: Wir fragen Microsoft, wie die Site-ID wirklich lautet
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });
        const siteData = await siteRes.json();

        if (siteData.error) {
            content.innerHTML = `<div class="p-4 bg-red-100 text-red-700"><strong>Site nicht gefunden:</strong> ${siteData.error.message}</div>`;
            return;
        }

        const realSiteId = siteData.id;

        // SCHRITT 2: Wir listen ALLE verfügbaren Listen auf, um die richtige ID zu finden
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${realSiteId}/lists`, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });
        const listsData = await listsRes.json();

        let html = `<div class="bg-slate-900 text-green-400 p-4 rounded font-mono text-xs mb-6">
                        <p>> Site-ID gefunden: ${realSiteId}</p>
                        <p>> Suche CRMFirms in ${listsData.value.length} Listen...</p>
                    </div>`;

        // SCHRITT 3: Wir suchen CRMFirms und laden die Items
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");

        if (!targetList) {
            html += `<p class="text-orange-500">❌ Liste "CRMFirms" nicht gefunden. Verfügbar sind: ${listsData.value.map(l => l.displayName).join(", ")}</p>`;
            content.innerHTML = html;
            return;
        }

        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${realSiteId}/lists/${targetList.id}/items?expand=fields`, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });
        const itemsData = await itemsRes.json();

        html += `<h2 class="text-xl font-bold text-slate-800 mb-4">🏢 Firmenliste (${itemsData.value.length})</h2><div class="space-y-2">`;
        
        itemsData.value.forEach(item => {
            html += `<div class="p-3 bg-white border rounded shadow-sm flex justify-between">
                        <span class="font-bold">${item.fields.Title || 'Unbenannt'}</span>
                        <span class="text-blue-500">${item.fields.Klassifizierung || '-'}</span>
                     </div>`;
        });

        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-red-50 text-red-700">Fehler: ${err.message}</div>`;
    }
}

function showView(v) { if(v === 'dashboard') location.reload(); }
