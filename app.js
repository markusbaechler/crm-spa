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

// 1. Sobald die Seite lädt: Prüfen, ob wir gerade vom Login zurückkommen
msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        console.log("Login erfolgreich zurückgekehrt");
        document.getElementById('authBtn').innerText = "Eingeloggt ✅";
        loadFirms(); // Direkt Firmen laden
    }
}).catch(err => {
    console.error("Fehler beim Zurückkehren:", err);
});

// 2. Die Login-Funktion (NUR NOCH REDIRECT!)
function handleAuth() {
    const loginRequest = {
        scopes: ["https://graph.microsoft.com/Sites.Read.All"]
    };
    // Keinerlei Popups mehr -> Seite springt zu Microsoft
    msalInstance.loginRedirect(loginRequest);
}

// 3. Firmen laden
async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = '<button onclick="handleAuth()" class="bg-blue-600 text-white p-4 rounded">Bitte einloggen</button>';
        return;
    }

    content.innerHTML = "Lade Firmen...";

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(err => {
            // Wenn der Token abgelaufen ist -> wieder Redirect
            return msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] });
        });

        if (!tokenRes) return;

        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists.firms}/items?expand=fields(select=Title,Klassifizierung)`;
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();
        let html = '<h2 class="text-xl font-bold mb-4 italic">🏢 Deine Firmen</h2>';
        data.value.forEach(item => {
            html += `<div class="p-2 border-b uppercase">${item.fields.Title || 'Unbenannt'}</div>`;
        });
        content.innerHTML = html;

    } catch (err) {
        content.innerHTML = "Fehler: " + err.message;
    }
}
