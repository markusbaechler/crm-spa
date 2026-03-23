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

// WICHTIG: Wir nutzen jetzt exakt die Berechtigung, die du in Azure hast
const loginRequest = {
    scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"]
};

window.onload = async () => {
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            console.log("Login erfolgreich");
        }
        checkAuthState();
    } catch (err) {
        console.error("Auth Fehler:", err);
        // Falls der Fehler kommt: Lokalen Speicher putzen
        localStorage.clear();
    }
};

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    const content = document.getElementById('app-content');

    if (accounts.length > 0) {
        authBtn.innerText = "Logout";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        authBtn.classList.replace('bg-blue-600', 'bg-red-600');
        loadFirms();
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
        authBtn.classList.contains('bg-red-600') ? authBtn.classList.replace('bg-red-600', 'bg-blue-600') : null;
    }
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        }).catch(() => msalInstance.acquireTokenRedirect(loginRequest));

        if (!tokenRes) return;

        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const siteData = await siteRes.json();
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");

        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/${targetList.id}/items?expand=fields`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const itemsData = await itemsRes.json();

        let html = `<h2 class="text-xl font-bold mb-4 italic">🏢 Firmenliste (${itemsData.value.length})</h2><div class="grid gap-2">`;
        itemsData.value.forEach(item => {
            html += `<div class="p-3 bg-white border rounded shadow-sm flex justify-between">
                <span class="font-bold text-slate-700">${item.fields.Title || 'Unbenannt'}</span>
                <span class="text-xs text-blue-500 font-black">${item.fields.Klassifizierung || '-'}</span>
            </div>`;
        });
        content.innerHTML = html + '</div>';

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-orange-50 text-orange-700">Bitte erneut einloggen (Sitzung abgelaufen).</div>`;
    }
}

function showView(v) { if (v === 'dashboard') location.reload(); }
