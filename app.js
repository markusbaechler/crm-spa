const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM"
};

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    },
    cache: { 
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Initialisierung beim Laden
window.onload = async () => {
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            console.log("Login erfolgreich");
        }
        checkAuthState();
    } catch (err) {
        console.error("Auth Fehler:", err);
    }
};

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    const content = document.getElementById('app-content');

    if (accounts.length > 0) {
        authBtn.innerText = "Logout";
        authBtn.onclick = handleLogout;
        authBtn.classList.replace('bg-blue-600', 'bg-red-600');
        // Falls wir auf der Firmen-Seite sind, laden wir sie jetzt sauber
        if (content.innerHTML.includes('Lade Firmen')) loadFirms();
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = handleAuth;
        authBtn.classList.contains('bg-red-600') ? authBtn.classList.replace('bg-red-600', 'bg-blue-600') : null;
        content.innerHTML = '<div class="text-center p-10 font-bold text-slate-400">Bitte einloggen, um Daten zu sehen.</div>';
        localStorage.clear(); // Putzt alle Reste weg
    }
}

async function handleAuth() {
    await msalInstance.loginRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.Write.All"] });
}

async function handleLogout() {
    const logoutRequest = { account: msalInstance.getAllAccounts()[0] };
    await msalInstance.logoutRedirect(logoutRequest);
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        checkAuthState();
        return;
    }

    content.innerHTML = '<p class="p-6 text-center animate-pulse">Verifiziere Zugriff...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.Write.All"],
            account: accounts[0]
        });

        // 1. Site suchen
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const siteData = await siteRes.json();
        
        // 2. Liste suchen
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");

        // 3. Daten holen
        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/${targetList.id}/items?expand=fields`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const itemsData = await itemsRes.json();

        renderFirms(itemsData.value);

    } catch (err) {
        console.error(err);
        content.innerHTML = `<div class="p-4 bg-red-100 text-red-700">Sitzung abgelaufen. Bitte neu einloggen.</div>`;
    }
}

function renderFirms(firms) {
    let html = `<h2 class="text-xl font-bold mb-4">Firmenliste (${firms.length})</h2><div class="grid gap-2">`;
    firms.forEach(item => {
        html += `<div class="p-4 bg-white border rounded shadow-sm flex justify-between">
            <span class="font-bold">${item.fields.Title || 'Unbenannt'}</span>
            <span class="text-xs text-slate-400">${item.fields.Klassifizierung || '-'}</span>
        </div>`;
    });
    document.getElementById('app-content').innerHTML = html + '</div>';
}

function showView(v) {
    if (v === 'dashboard') location.reload();
    if (v === 'firms') loadFirms();
}
