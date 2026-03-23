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
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function handleAuth() {
    const loginRequest = { scopes: ["Sites.Read.All"] };
    try {
        await msalInstance.loginPopup(loginRequest);
        location.reload();
    } catch (err) {
        alert("Login fehlgeschlagen: " + err.message);
    }
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    content.innerHTML = '<p class="p-4">Verbindung wird geprüft...</p>';

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        content.innerHTML = '<button onclick="handleAuth()" class="bg-blue-600 text-white p-3 rounded">Bitte erst einloggen</button>';
        return;
    }

    try {
        // Token holen
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["Sites.Read.All"],
            account: accounts[0]
        });

        // Abfrage (wir nehmen erst mal nur Title, um Fehlerquellen zu minimieren)
        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields(select=Title,Klassifizierung)`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` }
        });

        if (!response.ok) {
            const errorInfo = await response.json();
            throw new Error(errorInfo.error.message);
        }

        const data = await response.json();
        renderList(data.value);

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-orange-100 text-orange-800 border-l-4 border-orange-500">
            <strong>Fehler:</strong> ${err.message}<br>
            <small>Tipp: Wurde die "Administratorzustimmung" in Azure erteilt?</small>
        </div>`;
    }
}

function renderList(items) {
    const content = document.getElementById('app-content');
    if (!items || items.length === 0) {
        content.innerHTML = "Keine Daten gefunden (Liste ist leer).";
        return;
    }

    let html = `<h2 class="text-xl font-bold mb-4 border-b pb-2">🏢 Firmenliste</h2>`;
    items.forEach(item => {
        const name = item.fields.Title || "Kein Name";
        const klasse = item.fields.Klassifizierung || "-";
        html += `<div class="p-2 border-b flex justify-between hover:bg-gray-50">
                    <span>${name}</span>
                    <span class="font-mono font-bold">${klasse}</span>
                 </div>`;
    });
    content.innerHTML = html;
}

// Navigation
function showView(v) { if(v === 'dashboard') location.reload(); }
