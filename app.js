// --- CONFIG & VERSION ---
const appVersion = "0.46";
const appName = "CRM bbz";

const config = { 
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a", 
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7", 
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM" 
};

const msalConfig = { auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, redirectUri: "https://markusbaechler.github.io/crm-spa/" }, cache: { cacheLocation: "localStorage" } };
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"] };

let allFirms = [], allContacts = [], classOptions = [], currentSiteId = "", currentListId = "", contactListId = "";

window.onload = async () => { 
    updateFooter();
    await msalInstance.handleRedirectPromise(); 
    checkAuthState(); 
};

function updateFooter() {
    const ft = document.getElementById('footer-text');
    if (ft) ft.innerHTML = `© 2026 ${appName} | <b>Build ${appVersion}</b>`;
}

function checkAuthState() {
    const acc = msalInstance.getAllAccounts();
    const btn = document.getElementById('authBtn');
    if (acc.length > 0) {
        btn.innerText = "Logout"; btn.onclick = () => msalInstance.logoutRedirect({ account: acc[0] });
        loadData();
    } else {
        btn.innerText = "Login"; btn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('main-content');
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse">Lade System...</div>`;
    try {
        const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
        const h = { 'Authorization': `Bearer ${token}` };
        const s = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers: h })).json();
        currentSiteId = s.id;
        const l = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers: h })).json();
        currentListId = l.value.find(x => x.displayName === "CRMFirms").id;
        contactListId = l.value.find(x => x.displayName === "CRMContacts").id;

        const [cO, fD, cD] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers: h }).then(r => r.json())
        ]);
        classOptions = cO.choice?.choices || []; 
        allFirms = fD.value; 
        allContacts = cD.value;
        renderFirms();
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500 font-bold italic">FEHLER: ${e.message}</div>`; }
}

// --- NAVIGATION & HEADER FIX ---
function initHeader() {
    // Falls deine index.html IDs für die Nav-Links hat, hier verknüpfen
    const logo = document.querySelector('.text-blue-600.font-black'); // CRM bbz Logo
    if(logo) logo.style.cursor = "pointer";
    if(logo) logo.onclick = () => renderFirms();
}

// --- FIRMENÜBERSICHT ---
function renderFirms() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto">
            <div class="flex justify-between items-center mb-6 border-b pb-4">
                <div class="flex gap-8 items-end">
                    <h2 onclick="renderFirms()" class="text-2xl font-black text-slate-800 uppercase tracking-tighter cursor-pointer">Firmen</h2>
                    <nav class="flex gap-4 text-xs font-bold uppercase tracking-widest pb-1">
                        <button onclick="renderFirms()" class="text-blue-600 border-b-2 border-blue-600">Übersicht</button>
                        <button onclick="renderAllContacts()" class="text-slate-400 hover:text-slate-600">Alle Kontakte</button>
                    </nav>
                </div>
                <div class="flex gap-3">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Firma suchen..." class="px-4 py-2 border rounded-lg text-sm w-64 shadow-sm outline-none">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-xs font-bold uppercase tracking-widest">+ Firma</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 border rounded-2xl shadow-inner">
                <h3 class="text-[10px] font-black text-slate-400 uppercase mb-4 tracking-widest">Neue Organisation</h3>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                    <input type="text" id="new_fName" placeholder="Name der Firma *" class="p-2 border rounded text-sm">
                    <select id="new_fClass" class="p-2 border rounded text-sm">
                        <option value="">Klassierung</option>
                        ${classOptions.map(o => `<option value="${o}">${o}</option>`).join('')}
                    </select>
                    <input type="text" id="new_fPhone" placeholder="Hauptnummer" class="p-2 border rounded text-sm">
                </div>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <input type="text" id="new_fStreet" placeholder="Strasse / Nr." class="p-2 border rounded text-sm">
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-2 border rounded text-sm">
                    <input type="text" id="new_fCountry" value="CH" class="p-2 border rounded text-sm font-bold">
                </div>
                <div class="mt-6 flex gap-3">
                    <button onclick="saveNewFirm()" class="bg-green-600 text-white font-bold text-xs px-6 py-2 rounded uppercase tracking-widest shadow-md">Speichern</button>
                    <button onclick="toggleAddForm()" class="text-slate-400 font-bold text-xs uppercase px-4">Abbrechen</button>
                </div>
            </div>

            <div class="bg-white border rounded-xl overflow-hidden shadow-sm">
                <table class="w-full text-sm text-left">
                    <thead class="bg-slate-50 border-b text-[10px] uppercase font-black text-slate-500 tracking-widest">
                        <tr><th class="p-4">Firma</th><th class="p-4">Ort</th><th class="p-4 text-center">Klasse</th><th class="p-4 text-center">Kontakte</th></tr>
                    </thead>
                    <tbody id="firmTableBody" class="divide-y divide-slate-100 italic">
                        ${allFirms.map(f => `
                            <tr onclick="renderDetail('${f.id}')" class="hover:bg-blue-50 cursor-pointer transition-colors">
                                <td class="p-4 font-bold text-slate-700">${f.fields.Title} ${f.fields.VIP ? '⭐' : ''}</td>
                                <td class="p-4 text-slate-500 font-medium">${f.fields.Ort || '-'}</td>
                                <td class="p-4 text-center"><span class="px-2 py-1 bg-slate-100 rounded text-[10px] font-bold">${f.fields.Klassifizierung || '-'}</span></td>
                                <td class="p-4 text-center text-blue-600 font-bold">${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

// --- GLOBALE KONTAKTÜBERSICHT MIT VERLINKUNG ---
function renderAllContacts() {
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto">
             <div class="flex justify-between items-center mb-6 border-b pb-4">
                <div class="flex gap-8 items-end">
                    <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter cursor-pointer">Kontakte</h2>
                    <nav class="flex gap-4 text-xs font-bold uppercase tracking-widest pb-1">
                        <button onclick="renderFirms()" class="text-slate-400 hover:text-slate-600">Firmen</button>
                        <button onclick="renderAllContacts()" class="text-blue-600 border-b-2 border-blue-600">Alle Kontakte</button>
                    </nav>
                </div>
                <input type="text" onkeyup="filterContacts(this.value)" placeholder="Person suchen..." class="px-4 py-2 border rounded-lg text-sm w-80 shadow-sm outline-none">
            </div>
            <div class="bg-white border rounded-xl overflow-hidden shadow-sm italic">
                <table class="w-full text-[11px] text-left">
                    <thead class="bg-slate-50 border-b text-[9px] uppercase font-black text-slate-500 tracking-widest">
                        <tr><th class="p-3">Name</th><th class="p-3">Firma</th><th class="p-3">Rolle / Funktion</th><th class="p-3">SGF</th><th class="p-3">Kontaktinfos</th></tr>
                    </thead>
                    <tbody id="contactTableBody" class="divide-y divide-slate-100">
                        ${allContacts.map(c => {
                            const firm = allFirms.find(x => String(x.id) === String(c.fields.FirmaLookupId));
                            return `
                            <tr class="hover:bg-slate-50">
                                <td class="p-3 font-bold text-slate-800 text-sm">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                                <td class="p-3 font-bold text-blue-600 hover:underline cursor-pointer" onclick="renderDetail('${c.fields.FirmaLookupId}')">${firm ? firm.fields.Title : '-'}</td>
                                <td class="p-3">
                                    <div class="font-bold text-blue-600 uppercase text-[9px]">${c.fields.Rolle || ''}</div>
                                    <div class="text-slate-500">${c.fields.Funktion || ''}</div>
                                </td>
                                <td class="p-3 font-bold">${c.fields.SGF || '-'}</td>
                                <td class="p-3 space-y-0.5 text-slate-500 font-medium">
                                    <div>📧 ${c.fields.Email1 || '-'}</div>
                                    <div>📞 ${c.fields.Direktwahl || '-'} | 📱 ${c.fields.Mobile || '-'}</div>
                                </td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

// --- FIRMEN DETAILSEITE MIT KONTAKT-ERFASSUNG IM FLOW ---
function renderDetail(id) {
    const firm = allFirms.find(x => String(x.id) === String(id)), f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(id));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right duration-300">
            <div class="bg-white border rounded-xl p-6 mb-6 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms()" class="bg-slate-100 p-2 rounded hover:bg-slate-200">←</button>
                    <div>
                        <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter">${f.Title}</h2>
                        <div class="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">📍 ${f.Ort} | Klasse ${f.Klassifizierung || 'LEER'}</div>
                    </div>
                </div>
                <div class="flex gap-3">
                    <button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')" class="bg-slate-800 text-white px-5 py-2 rounded text-[10px] font-bold uppercase tracking-widest">Stammdaten</button>
                </div>
            </div>

            <div class="grid grid-cols-12 gap-6">
                <div class="col-span-12">
                    <div class="bg-white border rounded-xl shadow-sm overflow-hidden italic">
                        <div class="p-4 bg-slate-50 border-b flex justify-between items-center">
                            <h3 class="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="toggleContactForm()" class="bg-blue-600 text-white px-3 py-1.5 rounded text-[9px] font-bold uppercase tracking-widest shadow-sm">+ Neuer Kontakt</button>
                        </div>

                        <div id="addContactForm" class="hidden p-6 bg-blue-50/50 border-b animate-in fade-in">
                            <h4 class="text-[10px] font-black uppercase text-blue-600 mb-4 tracking-widest">Kontakt hinzufügen</h4>
                            <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
                                <input type="text" id="c_vn" placeholder="Vorname" class="p-2 border rounded text-sm outline-none">
                                <input type="text" id="c_nn" placeholder="Nachname *" class="p-2 border rounded text-sm outline-none font-bold">
                                <input type="text" id="c_email" placeholder="E-Mail" class="p-2 border rounded text-sm outline-none">
                                <input type="text" id="c_mobil" placeholder="Mobile" class="p-2 border rounded text-sm outline-none">
                            </div>
                            <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                                <input type="text" id="c_rolle" placeholder="Rolle" class="p-2 border rounded text-sm outline-none">
                                <input type="text" id="c_sgf" placeholder="SGF" class="p-2 border rounded text-sm outline-none uppercase">
                                <button onclick="saveContact('${id}')" class="bg-blue-600 text-white font-bold text-[10px] uppercase rounded tracking-widest shadow-md">Kontakt Speichern</button>
                                <button onclick="toggleContactForm()" class="text-slate-400 font-bold text-[10px] uppercase">Abbrechen</button>
                            </div>
                        </div>

                        <table class="w-full text-left text-[11px]">
                            <thead class="text-[9px] uppercase font-black text-slate-400 border-b bg-white">
                                <tr><th class="p-4">Name</th><th class="p-4">Rolle / Funktion</th><th class="p-4">SGF</th><th class="p-4">Kontaktinfo</th><th class="p-4 text-right">Aktion</th></tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                ${contacts.map(c => `
                                    <tr class="hover:bg-slate-50 transition-colors">
                                        <td class="p-4 font-bold text-slate-800 text-sm italic underline cursor-pointer" onclick="alert('Kontaktdetail kommt in V0.47')">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                                        <td class="p-4">
                                            <div class="font-bold text-blue-600 uppercase text-[9px]">${c.fields.Rolle || '-'}</div>
                                            <div class="text-slate-400 font-medium tracking-tight">${c.fields.Funktion || '-'}</div>
                                        </td>
                                        <td class="p-4 font-black">${c.fields.SGF || '-'}</td>
                                        <td class="p-4 space-y-1 font-medium">
                                            <div>📧 ${c.fields.Email1 || '-'}</div>
                                            <div>📱 ${c.fields.Mobile || '-'}</div>
                                        </td>
                                        <td class="p-4 text-right">
                                            <button onclick="deleteContact('${c.id}','${id}')" class="text-red-300 hover:text-red-600 font-bold uppercase text-[9px]">Löschen</button>
                                        </td>
                                    </tr>`).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div id="editSide" class="fixed inset-y-0 right-0 w-[400px] bg-white shadow-2xl z-50 transform translate-x-full transition-transform p-8 border-l italic">
                 <div class="flex justify-between items-center mb-8 border-b pb-4"><h3 class="font-black uppercase text-sm tracking-tighter italic">Stammdaten</h3><button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')">✕</button></div>
                 <div class="space-y-4">
                    <label class="block text-[10px] font-black text-slate-400 uppercase italic">Firma</label>
                    <input type="text" id="e_t" value="${f.Title}" class="w-full p-2 border rounded-lg text-sm font-black italic">
                    <label class="block text-[10px] font-black text-slate-400 uppercase italic">Klassierung</label>
                    <select id="e_k" class="w-full p-2 border rounded-lg text-sm font-bold">${classOptions.map(o => `<option value="${o}" ${f.Klassifizierung===o?'selected':''}>${o}</option>`).join('')}</select>
                    <label class="block text-[10px] font-black text-slate-400 uppercase italic">Hauptnummer</label>
                    <input type="text" id="e_p" value="${f.Hauptnummer||''}" class="w-full p-2 border rounded-lg text-sm">
                    <label class="block text-[10px] font-black text-slate-400 uppercase italic">Adresse</label>
                    <input type="text" id="e_s" value="${f.Adresse||''}" class="w-full p-2 border rounded-lg text-sm mb-2">
                    <div class="flex gap-2"><input type="text" id="e_c" value="${f.Ort}" class="flex-1 p-2 border rounded-lg text-sm font-bold italic"><input type="text" id="e_l" value="${f.Land||'CH'}" class="w-16 p-2 bg-slate-100 border rounded-lg text-xs text-center font-black"></div>
                    <button onclick="updateFirmData('${id}')" class="w-full bg-slate-900 text-white py-3 rounded-lg font-black text-xs uppercase tracking-widest mt-8 shadow-lg">Speichern</button>
                 </div>
            </div>
        </div>`;
}

// --- ACTIONS & LOGIK V0.46 ---

async function saveContact(firmId) {
    const nn = document.getElementById('c_nn').value; if(!nn) return alert("Nachname fehlt!");
    const fields = { 
        Title: nn, 
        Vorname: document.getElementById('c_vn').value, 
        Email1: document.getElementById('c_email').value,
        Mobile: document.getElementById('c_mobil').value,
        Rolle: document.getElementById('c_rolle').value,
        SGF: document.getElementById('c_sgf').value,
        FirmaLookupId: firmId 
    };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

async function updateFirmData(id) {
    const fields = { Title: document.getElementById('e_t').value, Klassifizierung: document.getElementById('e_k').value, Adresse: document.getElementById('e_s').value, Ort: document.getElementById('e_c').value, Land: document.getElementById('e_l').value, Hauptnummer: document.getElementById('e_p').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}/fields`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fields) });
    loadData();
}

async function saveNewFirm() {
    const n = document.getElementById('new_fName').value; if(!n) return;
    const fields = { Title: n, Klassifizierung: document.getElementById('new_fClass').value, Ort: document.getElementById('new_fCity').value, Hauptnummer: document.getElementById('new_fPhone').value, Adresse: document.getElementById('new_fStreet').value, Land: document.getElementById('new_fCountry').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

function filterFirms(q) { const ql = q.toLowerCase(); const f = allFirms.filter(x => x.fields.Title?.toLowerCase().includes(ql) || x.fields.Ort?.toLowerCase().includes(ql)); renderFirmsList(f); }
function renderFirmsList(firms) { document.getElementById('firmTableBody').innerHTML = firms.map(f => ` <tr onclick="renderDetail('${f.id}')" class="hover:bg-blue-50 cursor-pointer transition-colors"><td class="p-4 font-bold text-slate-700">${f.fields.Title} ${f.fields.VIP ? '⭐' : ''}</td><td class="p-4 text-slate-500 font-medium">${f.fields.Ort || '-'}</td><td class="p-4 text-center font-bold text-[10px] italic">${f.fields.Klassifizierung || '-'}</td><td class="p-4 text-center text-blue-600 font-bold">${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</td></tr>`).join(''); }

function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function toggleContactForm() { document.getElementById('addContactForm').classList.toggle('hidden'); }
async function deleteContact(cid, fid) { if(!confirm("Löschen?")) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cid}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
