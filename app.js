/**
 * CRM bbz (light) - Full Application Logic V3.1
 * Fehlerbehebung: Site-ID Pfad-Auflösung
 */

// 1. KONFIGURATION
const msalConfig = {
    auth: {
        clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
        authority: "https://login.microsoftonline.com/3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    }
};

// Pfad-basierte ID für Microsoft Graph (robuster als die interne ID-Kette)
const sitePath = "bbzsg.sharepoint.com:/sites/CRM"; 
const listFirms = "CRMFirms";
const listContacts = "CRMContacts";

const msalInstance = new msal.PublicClientApplication(msalConfig);

// 2. AUTHENTIFIZIERUNG
async function login() {
    try {
        await msalInstance.loginPopup({ scopes: ["Sites.ReadWrite.All"] });
        location.reload();
    } catch (err) {
        console.error("Login fehlgeschlagen:", err);
    }
}

async function getGraphToken() {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) return null;
    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes: ["Sites.ReadWrite.All"],
            account: account
        });
        return response.accessToken;
    } catch (err) {
        console.warn("Silent token acquisition failed, requiring popup", err);
        return null;
    }
}

// 3. GRAPH API WRAPPER
async function graphRequest(endpoint, method = "GET", body = null) {
    const token = await getGraphToken();
    if (!token) return Promise.reject("Nicht eingeloggt oder Token abgelaufen.");

    const options = {
        method: method,
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        }
    };
    if (body) options.body = JSON.stringify({ fields: body });

    // Auflösung über den Pfad: /sites/{host}:/{path}
    const url = `https://graph.microsoft.com/v1.0/sites/${sitePath}/lists/${endpoint}`;
    
    const response = await fetch(url, options);
    if (!response.ok) {
        const errorDetails = await response.text();
        throw new Error(errorDetails);
    }
    return await response.json();
}

// 4. FIRMEN-LOGIK
async function loadFirms() {
    const container = document.getElementById('main-content');
    container.innerHTML = '<div class="text-center p-10"><p class="animate-pulse">Lade Firmen aus SharePoint...</p></div>';
    
    try {
        const data = await graphRequest(`${listFirms}/items?expand=fields`);
        const firms = data.value.map(item => item.fields);
        renderFirms(firms);
    } catch (err) {
        container.innerHTML = `<div class="p-4 bg-red-50 text-red-700 rounded-lg"><strong>Fehler:</strong> ${err.message}</div>`;
    }
}

function renderFirms(firms) {
    const container = document.getElementById('main-content');
    container.innerHTML = `
        <div class="flex justify-between items-center mb-6">
            <h1 class="text-2xl font-bold text-slate-800">Firmenverzeichnis</h1>
        </div>
        <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse">
                <thead>
                    <tr class="bg-slate-100 border-b">
                        <th class="p-3 text-sm font-semibold text-slate-600">Firma</th>
                        <th class="p-3 text-sm font-semibold text-slate-600">Klassifizierung</th>
                        <th class="p-3 text-sm font-semibold text-slate-600">Ort</th>
                        <th class="p-3 text-sm font-semibold text-slate-600">Aktion</th>
                    </tr>
                </thead>
                <tbody class="divide-y bg-white">
                    ${firms.map(f => `
                        <tr class="hover:bg-blue-50 transition">
                            <td class="p-3 font-medium">${f.Title || 'Unbekannt'}</td>
                            <td class="p-3"><span class="px-2 py-1 rounded text-xs font-bold ${getClassBadge(f.Classification)}">${f.Classification || '-'}</span></td>
                            <td class="p-3 text-slate-600">${f.City || '-'}</td>
                            <td class="p-3">
                                <button onclick="openFirmDetails(${f.id})" class="text-blue-600 font-semibold hover:underline">Details</button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
}

// 5. KONTAKT-LOGIK
async function loadAllContacts() {
    const container = document.getElementById('main-content');
    container.innerHTML = '<p class="text-center p-10 italic text-slate-500">Lade alle Kontakte...</p>';
    
    try {
        const data = await graphRequest(`${listContacts}/items?expand=fields`);
        const contacts = data.value.map(item => item.fields);
        
        container.innerHTML = `
            <h1 class="text-2xl font-bold mb-6 text-slate-800">Alle Ansprechpartner</h1>
            <table class="w-full text-left border-collapse bg-white rounded-lg shadow-sm">
                <thead>
                    <tr class="bg-slate-100 border-b">
                        <th class="p-3 text-sm font-semibold text-slate-600">Name</th>
                        <th class="p-3 text-sm font-semibold text-slate-600">E-Mail</th>
                        <th class="p-3 text-sm font-semibold text-slate-600">FirmID</th>
                    </tr>
                </thead>
                <tbody class="divide-y">
                    ${contacts.map(c => `
                        <tr class="hover:bg-blue-50">
                            <td class="p-3 font-medium">${c.FirstName || ''} ${c.Title}</td>
                            <td class="p-3">${c.Email || '-'}</td>
                            <td class="p-3 text-xs text-slate-400 font-mono">${c.FirmID || 'Privat'}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;
    } catch (err) {
        container.innerHTML = `<p class="text-red-500">Fehler beim Laden der Kontakte: ${err.message}</p>`;
    }
}

async function openFirmDetails(firmId) {
    const modal = document.getElementById('detail-modal');
    const content = document.getElementById('modal-body-content');
    modal.classList.remove('hidden');
    content.innerHTML = '<div class="p-10 text-center animate-pulse">Lade Firmendetails...</div>';

    try {
        // Parallel laden: Die Firma selbst und alle Kontakte, die diese FirmID haben
        const [firmData, contactData] = await Promise.all([
            graphRequest(`${listFirms}/items/${firmId}?expand=fields`),
            graphRequest(`${listContacts}/items?expand=fields&$filter=fields/FirmID eq '${firmId}'`)
        ]);

        const firm = firmData.fields;
        const contacts = contactData.value.map(item => item.fields);

        content.innerHTML = `
            <h2 class="text-2xl font-bold text-slate-800 mb-1">${firm.Title}</h2>
            <p class="text-slate-500 mb-6">${firm.Address || ''}, ${firm.ZIP || ''} ${firm.City || ''}</p>
            
            <div class="mb-6">
                <h3 class="text-sm font-bold uppercase tracking-wider text-slate-400 mb-2 border-b pb-1">Ansprechpartner</h3>
                <div class="mt-2 divide-y border rounded-lg bg-slate-50">
                    ${contacts.length > 0 ? contacts.map(c => `
                        <div class="p-3 flex justify-between items-center">
                            <div>
                                <p class="font-semibold text-slate-700">${c.FirstName || ''} ${c.Title}</p>
                                <p class="text-xs text-slate-500">${c.Email || 'Keine E-Mail hinterlegt'}</p>
                            </div>
                        </div>
                    `).join('') : '<p class="p-4 text-slate-400 italic text-sm text-center">Keine Kontakte zu dieser Firma gefunden.</p>'}
                </div>
            </div>

            <div class="flex gap-3 mt-8">
                <button onclick="addContact(${firm.id})" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-semibold hover:bg-blue-700 transition">
                    + Neuer Kontakt
                </button>
                <button onclick="closeModal()" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-lg text-sm font-semibold hover:bg-slate-300 transition">
                    Schließen
                </button>
            </div>
        `;
    } catch (err) {
        content.innerHTML = `<p class="text-red-500">Fehler beim Laden der Details: ${err.message}</p>`;
    }
}

// 6. HELPER
function getClassBadge(cls) {
    const colors = {
        'A': 'bg-amber-100 text-amber-700 border border-amber-200',
        'B': 'bg-slate-200 text-slate-700 border border-slate-300',
        'C': 'bg-orange-100 text-orange-700 border border-orange-200'
    };
    return colors[cls] || 'bg-gray-100 text-gray-500 border border-gray-200';
}

function closeModal() {
    document.getElementById('detail-modal').classList.add('hidden');
}

async function addContact(firmId) {
    const lastName = prompt("Nachname des Kontakts:");
    if (!lastName) return;
    const firstName = prompt("Vorname (optional):");
    const email = prompt("E-Mail (optional):");

    const newContact = {
        Title: lastName,
        FirstName: firstName,
        Email: email,
        FirmID: String(firmId)
    };

    try {
        await graphRequest(`${listContacts}/items`, "POST", newContact);
        openFirmDetails(firmId); // Refresh des Modals
    } catch (err) {
        alert("Speichern fehlgeschlagen: " + err.message);
    }
}

// Initialer Check: Wenn User eingeloggt, Firmen laden
window.onload = async () => {
    const account = msalInstance.getAllAccounts()[0];
    if (account) {
        document.getElementById('authBtn').innerText = 'Logout';
        document.getElementById('authBtn').onclick = () => { msalInstance.logoutPopup(); };
        loadFirms();
    }
};
