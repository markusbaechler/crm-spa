/**
 * CRM bbz (light) - Full Application Logic
 * Site: bbzsg.sharepoint.com/sites/CRM
 */

// 1. KONFIGURATION
const msalConfig = {
    auth: {
        clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
        authority: "https://login.microsoftonline.com/3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    }
};

const siteId = "bbzsg.sharepoint.com,3643e7ab-d166-4e27-bd5f-c5bbfcd282d7,bd5f-c5bbfcd282d7"; // Vereinfachte Site-ID für Graph
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
    const response = await msalInstance.acquireTokenSilent({
        scopes: ["Sites.ReadWrite.All"],
        account: account
    });
    return response.accessToken;
}

// 3. GRAPH API WRAPPER
async function graphRequest(endpoint, method = "GET", body = null) {
    const token = await getGraphToken();
    if (!token) return alert("Bitte erst einloggen!");

    const options = {
        method: method,
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        }
    };
    if (body) options.body = JSON.stringify({ fields: body });

    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${endpoint}`, options);
    return response.ok ? await response.json() : Promise.reject(await response.text());
}

// 4. FIRMEN-LOGIK
async function loadFirms() {
    const container = document.getElementById('main-content');
    container.innerHTML = '<p class="text-center p-10">Lade Firmen...</p>';
    
    try {
        const data = await graphRequest(`${listFirms}/items?expand=fields`);
        const firms = data.value.map(item => item.fields);
        renderFirms(firms);
    } catch (err) {
        container.innerHTML = `<p class="text-red-500">Fehler: ${err}</p>`;
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
                <tbody class="divide-y">
                    ${firms.map(f => `
                        <tr class="hover:bg-blue-50 transition">
                            <td class="p-3 font-medium">${f.Title}</td>
                            <td class="p-3"><span class="px-2 py-1 rounded text-xs font-bold ${getClassBadge(f.Classification)}">${f.Classification || '-'}</span></td>
                            <td class="p-3 text-slate-600">${f.City || '-'}</td>
                            <td class="p-3">
                                <button onclick="openFirmDetails(${f.id})" class="text-blue-600 font-semibold hover:text-blue-800">Details</button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
}

// 5. KONTAKT-LOGIK (Verknüpfung via FirmID)
async function loadAllContacts() {
    const container = document.getElementById('main-content');
    container.innerHTML = '<p class="text-center p-10">Lade alle Kontakte...</p>';
    
    const data = await graphRequest(`${listContacts}/items?expand=fields`);
    const contacts = data.value.map(item => item.fields);
    
    container.innerHTML = `
        <h1 class="text-2xl font-bold mb-6 text-slate-800">Alle Ansprechpartner</h1>
        <table class="w-full text-left border-collapse">
            <thead>
                <tr class="bg-slate-100 border-b">
                    <th class="p-3 text-sm font-semibold text-slate-600">Name</th>
                    <th class="p-3 text-sm font-semibold text-slate-600">E-Mail</th>
                    <th class="p-3 text-sm font-semibold text-slate-600">Zugehörige FirmID</th>
                </tr>
            </thead>
            <tbody class="divide-y">
                ${contacts.map(c => `
                    <tr class="hover:bg-blue-50">
                        <td class="p-3 font-medium">${c.Title} ${c.FirstName || ''}</td>
                        <td class="p-3">${c.Email || '-'}</td>
                        <td class="p-3 text-xs text-slate-400">${c.FirmID || 'Privat'}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
}

async function openFirmDetails(firmId) {
    const modal = document.getElementById('detail-modal');
    const content = document.getElementById('modal-body-content');
    modal.classList.remove('hidden');
    content.innerHTML = '<p>Lade Details...</p>';

    // Parallel laden: Firmendaten & zugehörige Kontakte
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
            <h3 class="text-sm font-bold uppercase tracking-wider text-slate-400 mb-2">Ansprechpartner</h3>
            <div class="bg-slate-50 rounded-lg border border-slate-200 divide-y">
                ${contacts.length > 0 ? contacts.map(c => `
                    <div class="p-3 flex justify-between items-center">
                        <div>
                            <p class="font-semibold text-slate-700">${c.FirstName || ''} ${c.Title}</p>
                            <p class="text-xs text-slate-500">${c.Email || 'Keine E-Mail'}</p>
                        </div>
                    </div>
                `).join('') : '<p class="p-4 text-slate-400 italic text-sm">Keine Kontakte hinterlegt.</p>'}
            </div>
        </div>

        <div class="flex gap-3">
            <button onclick="addContact(${firm.id})" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-sm hover:bg-blue-700 shadow-sm transition">
                + Kontakt hinzufügen
            </button>
            <button onclick="closeModal()" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-lg text-sm hover:bg-slate-300 transition">
                Schließen
            </button>
        </div>
    `;
}

// 6. HELPER FUNKTIONEN
function getClassBadge(cls) {
    if (cls === 'A') return 'bg-amber-100 text-amber-700';
    if (cls === 'B') return 'bg-slate-200 text-slate-700';
    if (cls === 'C') return 'bg-orange-100 text-orange-700';
    return 'bg-gray-100 text-gray-500';
}

function closeModal() {
    document.getElementById('detail-modal').classList.add('hidden');
}

async function addContact(firmId) {
    const lastName = prompt("Nachname des Kontakts:");
    if (!lastName) return;
    const firstName = prompt("Vorname:");
    const email = prompt("E-Mail:");

    const newContact = {
        Title: lastName,
        FirstName: firstName,
        Email: email,
        FirmID: String(firmId) // Hier passiert die Magie der Verknüpfung
    };

    try {
        await graphRequest(`${listContacts}/items`, "POST", newContact);
        openFirmDetails(firmId); // Refresh
    } catch (err) {
        alert("Fehler beim Speichern: " + err);
    }
}

// Initialer Check beim Laden
(async () => {
    const account = msalInstance.getAllAccounts()[0];
    if (account) {
        document.getElementById('auth-status').innerHTML = `<span>Angemeldet als ${account.username}</span>`;
        loadFirms();
    }
})();
