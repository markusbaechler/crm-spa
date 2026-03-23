/**
 * CRM bbz (light) - Core Logic V3.0
 * Fokus: Verknüpfung Firmen & Kontakte
 */

// Konfiguration (Anpassung an deine Umgebung)
const siteId = "bbzsg.sharepoint.com:/sites/CRM";
const listFirms = "CRMFirms";
const listContacts = "CRMContacts";

let currentFirmId = null; // Speichert die ID der aktuell im Modal geöffneten Firma

// --- 1. API WRAPPER (Graph API) ---
async function graphRequest(endpoint, method = "GET", body = null) {
    const account = msalInstance.getAllAccounts()[0];
    const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Sites.ReadWrite.All"],
        account: account
    });

    const options = {
        method: method,
        headers: {
            "Authorization": `Bearer ${tokenResponse.accessToken}`,
            "Content-Type": "application/json"
        }
    };
    if (body) options.body = JSON.stringify({ fields: body });

    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${endpoint}`, options);
    return response.ok ? await response.json() : Promise.reject(response.statusText);
}

// --- 2. FIRMEN LOGIK (Bestehend) ---
async function loadFirms() {
    const data = await graphRequest(`${listFirms}/items?expand=fields`);
    renderFirms(data.value.map(item => item.fields));
}

function renderFirms(firms) {
    const container = document.getElementById('main-content');
    container.innerHTML = `
        <h1 class="text-2xl font-bold mb-4">Firmenübersicht</h1>
        <table class="min-w-full bg-white border">
            <thead>
                <tr class="bg-gray-100">
                    <th class="p-2 border">Firma</th>
                    <th class="p-2 border">Ort</th>
                    <th class="p-2 border">Aktion</th>
                </tr>
            </thead>
            <tbody>
                ${firms.map(f => `
                    <tr>
                        <td class="p-2 border">${f.Title}</td>
                        <td class="p-2 border">${f.City || '-'}</td>
                        <td class="p-2 border">
                            <button onclick="openFirmDetails(${f.id})" class="text-blue-600 hover:underline">Details</button>
                        </td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
}

// --- 3. KONTAKT LOGIK (Neu) ---

// A. Globale Kontaktliste laden
async function loadAllContacts() {
    const data = await graphRequest(`${listContacts}/items?expand=fields`);
    const contacts = data.value.map(item => item.fields);
    
    const container = document.getElementById('main-content');
    container.innerHTML = `
        <h1 class="text-2xl font-bold mb-4">Alle Kontakte</h1>
        <table class="min-w-full bg-white border text-left">
            <thead>
                <tr class="bg-gray-100">
                    <th class="p-2 border">Name</th>
                    <th class="p-2 border">E-Mail</th>
                    <th class="p-2 border">Firma (ID)</th>
                </tr>
            </thead>
            <tbody>
                ${contacts.map(c => `
                    <tr>
                        <td class="p-2 border font-bold">${c.Title} ${c.FirstName || ''}</td>
                        <td class="p-2 border">${c.Email || '-'}</td>
                        <td class="p-2 border text-sm text-gray-500">${c.FirmID || 'Privat'}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
}

// B. Firmendetails öffnen & Kontakte filtern
async function openFirmDetails(firmId) {
    currentFirmId = firmId;
    // Lade Firmendaten
    const firmData = await graphRequest(`${listFirms}/items/${firmId}?expand=fields`);
    const firm = firmData.fields;

    // Lade zugehörige Kontakte (gefiltert nach FirmID)
    const contactData = await graphRequest(`${listContacts}/items?expand=fields&$filter=fields/FirmID eq '${firmId}'`);
    const contacts = contactData.value.map(item => item.fields);

    showModal(firm, contacts);
}

// C. UI: Modal für Details und Kontakt-Subliste
function showModal(firm, contacts) {
    const modal = document.getElementById('detail-modal');
    modal.classList.remove('hidden');
    
    modal.innerHTML = `
        <div class="bg-white p-6 rounded-lg max-w-2xl mx-auto mt-20 shadow-xl border">
            <div class="flex justify-between items-center mb-4">
                <h2 class="text-xl font-bold">${firm.Title}</h2>
                <button onclick="closeModal()" class="text-gray-500 text-2xl">&times;</button>
            </div>
            
            <div class="grid grid-cols-2 gap-4 mb-6">
                <div><strong>Adresse:</strong> ${firm.Address || '-'}</div>
                <div><strong>Klassifizierung:</strong> ${firm.Classification || '-'}</div>
            </div>

            <hr class="mb-4">
            <h3 class="font-bold mb-2 text-gray-700">Ansprechpartner bei dieser Firma:</h3>
            
            <ul class="bg-gray-50 rounded p-2 mb-4 border">
                ${contacts.length > 0 ? contacts.map(c => `
                    <li class="py-1 border-b last:border-0 flex justify-between">
                        <span>${c.Title} ${c.FirstName || ''}</span>
                        <span class="text-sm text-gray-500">${c.Email || ''}</span>
                    </li>
                `).join('') : '<li class="text-gray-400 italic">Keine Kontakte zugeordnet.</li>'}
            </ul>

            <div class="flex gap-2">
                <button onclick="createNewContactPrompt(${firm.id})" class="bg-green-600 text-white px-3 py-1 rounded text-sm hover:bg-green-700">
                    + Kontakt hinzufügen
                </button>
                <button onclick="closeModal()" class="bg-gray-200 px-3 py-1 rounded text-sm hover:bg-gray-300">Schließen</button>
            </div>
        </div>
    `;
}

// D. Neuen Kontakt anlegen (Direkt-Zuordnung)
async function createNewContactPrompt(firmId) {
    const name = prompt("Nachname des neuen Kontakts:");
    const email = prompt("E-Mail Adresse:");
    
    if (name) {
        const newContact = {
            Title: name,      // 'Title' ist das Standard-Feld für Nachname
            FirmID: String(firmId), // WICHTIG: Die Verknüpfung
            Email: email
        };

        try {
            await graphRequest(`${listContacts}/items`, "POST", newContact);
            alert("Kontakt erfolgreich gespeichert!");
            openFirmDetails(firmId); // Refresh des Modals
        } catch (error) {
            alert("Fehler beim Speichern!");
        }
    }
}

function closeModal() {
    document.getElementById('detail-modal').classList.add('hidden');
}
