// --- 1. KONFIGURATION ---
const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    // Der exakte Host und Pfad deiner Site
    siteHostname: "bbzsg.sharepoint.com",
    siteRelativePath: "/sites/CRM",
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

msalInstance.handleRedirectPromise().then(response => {
    if (response) {
        updateUI();
        loadFirms(); 
    }
}).catch(err => {
    console.error("Redirect Fehler:", err);
});

// --- 3. LOGIN FUNKTION ---
function handleAuth() {
    msalInstance.loginRedirect({
        scopes: ["https://graph.microsoft.com/Sites.Read.All"]
    });
}

function updateUI() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0 && authBtn) {
        authBtn.innerText = "Eingeloggt ✅";
        authBtn.classList.remove('bg-blue-600');
        authBtn.classList.add('bg-green-600');
    }
}

// --- 4. FIRMEN LADEN ---
async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = `
            <div class="text-center p-10">
                <button onclick="handleAuth()" class="bg-blue-600 text-white px-8 py-3 rounded-lg font-bold shadow-lg">
                    Login mit Microsoft
                </button>
            </div>`;
        return;
    }

    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600">Lade Firmen aus SharePoint...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({
            scopes: ["https://graph.microsoft.com/Sites.Read.All"],
            account: accounts[0]
        }).catch(() => msalInstance.acquireTokenRedirect({ scopes: ["https://graph.microsoft.com/Sites.Read.All"] }));

        if (!tokenRes) return;

        // KORREKTE URL-STRUKTUR: hostname:/sites/pfad
        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteHostname}:${config.siteRelativePath}/lists/${config.lists.firms}/items?expand=fields`;
        
        const response = await fetch(url, {
            headers: { 'Authorization': 'Bearer ' + tokenRes.accessToken }
        });

        const data = await response.json();

        if (data.error) {
            content.innerHTML = `
                <div class="p-6 bg-orange-50 border border-orange-200 rounded-lg">
                    <h3 class="text-orange-800 font-bold mb-2">Fehler beim Zugriff:</h3>
                    <p class="text-sm mb-4">${data.error.message}</p>
                    <pre class="text-xs bg-white p-4 border overflow-auto max-h-40">${JSON.stringify(data.error, null, 2)}</pre>
                </div>`;
            return;
        }

        const firms = data.value || [];
        if (firms.length === 0) {
            content.innerHTML = '<p class="p-10 text-center text-slate-500 italic font-light">Keine Firmen in der Liste "CRMFirms" gefunden.</p>';
            return;
        }

        let html = `
            <div class="flex justify-between items-center mb-6">
                <h2 class="text-2xl font-black text-slate-800 tracking-tight text-xl uppercase">🏢 Firmen</h2>
                <span class="bg-slate-200 text-slate-600 text-xs font-bold px-2 py-1 rounded">${firms.length}</span>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">`;
        
        firms.forEach(item => {
            const f = item.fields || {};
            html += `
                <div class="p-5 bg-white border border-slate-200 rounded-xl shadow-sm hover:shadow-md transition">
                    <div class="font-bold text-slate-800">${f.Title || 'Unbenannt'}</div>
                    <div class="mt-2 text-xs text-blue-600 font-bold uppercase tracking-wider">${f.Klassifizierung || '-'}</div>
                </div>`;
        });
        
        content.innerHTML = html + "</div>";

    } catch (err) {
        content.innerHTML = `<div class="p-6 bg-red-50 text-red-700 rounded-xl">Fehler: ${err.message}</div>`;
    }
}

window.onload = () => { updateUI(); };

function showView(v) { 
    if(v === 'dashboard') location.reload();
    if(v === 'contacts') document.getElementById('app-content').innerHTML = '<h2 class="p-10 text-center font-bold">Kontakte folgen...</h2>';
}
