// --- CONFIG & VERSION ---
const appVersion = "0.45";
const appName = "CRM bbz";
console.log(`${appName} ${appVersion} - High Density Data Build`);

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
    document.title = appName;
    const ft = document.getElementById('footer-text');
    if (ft) ft.innerHTML = `© 2026 ${appName} | <b>Build ${appVersion}</b>`;
    await msalInstance.handleRedirectPromise(); 
    checkAuthState(); 
};

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
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse">Lade ${appName} Daten...</div>`;
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
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500 font-bold">FEHLER: ${e.message}</div>`; }
}

// --- NAVIGATION ---
function showPage(page) {
    if (page === 'firms') renderFirms();
    if (page === 'contacts') renderAllContacts();
}

// --- FIRMENÜBERSICHT (LISTE) ---
function renderFirms() {
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-6 border-b pb-4">
                <div class="flex gap-8 items-end">
                    <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter">${appName} Firmen</h2>
                    <nav class="flex gap-4 text-xs font-bold uppercase tracking-widest pb-1">
                        <button onclick="showPage('firms')" class="text-blue-600 border-b-2 border-blue-600">Firmen</button>
                        <button onclick="showPage('contacts')" class="text-slate-400 hover:text-slate-600">Kontakte</button>
                    </nav>
                </div>
                <div class="flex gap-3">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-4 py-2 border rounded-lg text-sm w-64 shadow-sm outline-none focus:border-blue-500">
                    <button onclick="toggleAdd()" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-xs font-bold uppercase">+ FIRMA</button>
                </div>
            </div>
            <div id="addForm" class="hidden mb-6 p-4 bg-slate-50 border rounded-xl grid grid-cols-4 gap-4">
                <input type="text" id="new_fName" placeholder="Name *" class="p-2 border rounded text-sm"><input type="text" id="new_fCity" placeholder="Ort" class="p-2 border rounded text-sm">
                <button onclick="saveNewFirm()" class="bg-green-600 text-white font-bold text-xs rounded uppercase">Speichern</button>
            </div>
            <div class="bg-white border rounded-xl overflow-hidden shadow-sm">
                <table class="w-full text-sm text-left">
                    <thead class="bg-slate-50 border-b text-[10px] uppercase font-black text-slate-500 tracking-widest">
                        <tr><th class="p-4">Firma</th><th class="p-4">Ort</th><th class="p-4 text-center">Klasse</th><th class="p-4 text-center">Kontakte</th></tr>
                    </thead>
                    <tbody id="firmTableBody" class="divide-y divide-slate-100">
                        ${allFirms.map(f => `
                            <tr onclick="renderDetail('${f.id}')" class="hover:bg-blue-50 cursor-pointer transition-colors">
                                <td class="p-4 font-bold text-slate-700">${f.fields.Title} ${f.fields.VIP ? '⭐' : ''}</td>
                                <td class="p-4 text-slate-500">${f.fields.Ort || '-'}</td>
                                <td class="p-4 text-center"><span class="px-2 py-1 bg-slate-100 rounded text-[10px] font-bold">${f.fields.Klassifizierung || '-'}</span></td>
                                <td class="p-4 text-center text-blue-600 font-bold">${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

// --- GLOBALE KONTAKTÜBERSICHT ---
function renderAllContacts() {
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
             <div class="flex justify-between items-center mb-6 border-b pb-4">
                <div class="flex gap-8 items-end">
                    <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter">${appName} Kontakte</h2>
                    <nav class="flex gap-4 text-xs font-bold uppercase tracking-widest pb-1">
                        <button onclick="showPage('firms')" class="text-slate-400 hover:text-slate-600">Firmen</button>
                        <button onclick="showPage('contacts')" class="text-blue-600 border-b-2 border-blue-600">Kontakte</button>
                    </nav>
                </div>
                <input type="text" onkeyup="filterContacts(this.value)" placeholder="Person suchen..." class="px-4 py-2 border rounded-lg text-sm w-80 shadow-sm outline-none focus:border-blue-500">
            </div>
            <div class="bg-white border rounded-xl overflow-hidden shadow-sm">
                <table class="w-full text-[11px] text-left">
                    <thead class="bg-slate-50 border-b text-[9px] uppercase font-black text-slate-500 tracking-widest">
                        <tr><th class="p-3">Name</th><th class="p-3">Firma</th><th class="p-3">Rolle / Funktion</th><th class="p-3">SGF</th><th class="p-3">Lead bbz</th><th class="p-3">Kontaktinfos</th></tr>
                    </thead>
                    <tbody id="contactTableBody" class="divide-y divide-slate-100">
                        ${allContacts.map(c => contactRow(c, true)).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function contactRow(c, showFirmName) {
    const f = c.fields;
    const firmName = showFirmName ? (allFirms.find(x => String(x.id) === String(f.FirmaLookupId))?.fields.Title || 'n.a.') : '';
    return `
        <tr class="hover:bg-slate-50 transition-colors group">
            <td class="p-3 font-bold text-slate-800">${f.Vorname || ''} ${f.Title}</td>
            <td class="p-3 text-slate-400 font-bold">${firmName}</td>
            <td class="p-3">
                <div class="font-bold text-blue-600 uppercase text-[9px]">${f.Rolle || ''}</div>
                <div class="text-slate-500">${f.Funktion || ''}</div>
            </td>
            <td class="p-3"><span class="px-1.5 py-0.5 bg-emerald-50 text-emerald-600 rounded border border-emerald-100 font-black">${f.SGF || '-'}</span></td>
            <td class="p-3 text-slate-400 font-bold uppercase">${f.Leadbbz0 || '-'}</td>
            <td class="p-3 space-y-0.5 text-slate-500">
                <div class="flex items-center gap-2">📧 ${f.Email1 || '-'}</div>
                <div class="flex items-center gap-2">📞 ${f.Direktwahl || '-'} | 📱 ${f.Mobile || '-'}</div>
            </td>
        </tr>`;
}

// --- FIRMEN DETAILSEITE ---
function renderDetail(id) {
    const firm = allFirms.find(x => String(x.id) === String(id)), f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(id));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right duration-200">
            <div class="bg-white border rounded-xl p-6 mb-6 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms()" class="bg-slate-100 p-2 rounded hover:bg-slate-200">←</button>
                    <div>
                        <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter">${f.Title}</h2>
                        <div class="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">📍 ${f.Ort} | 📞 ${f.Hauptnummer || '-'} | Klasse ${f.Klassifizierung}</div>
                    </div>
                </div>
                <button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')" class="bg-slate-800 text-white px-5 py-2 rounded text-[10px] font-bold uppercase">Stammdaten</button>
            </div>

            <div class="grid grid-cols-12 gap-6">
                <div class="col-span-12">
                    <div class="bg-white border rounded-xl shadow-sm overflow-hidden">
                        <div class="p-4 bg-slate-50 border-b flex justify-between items-center">
                            <h3 class="text-[10px] font-black text-slate-500 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${id}')" class="bg-blue-600 text-white px-3 py-1.5 rounded text-[9px] font-bold uppercase tracking-widest">+ Kontakt</button>
                        </div>
                        <table class="w-full text-left text-[11px]">
                            <thead class="text-[9px] uppercase font-black text-slate-400 border-b bg-white">
                                <tr><th class="p-4">Name</th><th class="p-4">Rolle / Funktion</th><th class="p-4">SGF</th><th class="p-4">Lead bbz</th><th class="p-4">Kontaktinfo</th><th class="p-4 text-right">Aktion</th></tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                ${contacts.map(c => `
                                    <tr class="hover:bg-slate-50">
                                        <td class="p-4 font-bold text-slate-800 text-sm">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                                        <td class="p-4">
                                            <div class="font-bold text-blue-600 uppercase text-[9px]">${c.fields.Rolle || '-'}</div>
                                            <div class="text-slate-400">${c.fields.Funktion || '-'}</div>
                                        </td>
                                        <td class="p-4"><span class="px-1.5 py-0.5 bg-slate-100 rounded border font-bold">${c.fields.SGF || '-'}</span></td>
                                        <td class="p-4 text-slate-400 font-bold uppercase">${c.fields.Leadbbz0 || '-'}</td>
                                        <td class="p-4 space-y-1">
                                            <div>📧 ${c.fields.Email1 || '-'}</div>
                                            <div>📞 ${c.fields.Direktwahl || '-'} | 📱 ${c.fields.Mobile || '-'}</div>
                                        </td>
                                        <td class="p-4 text-right">
                                            <button onclick="deleteContact('${c.id}','${id}')" class="text-red-300 hover:text-red-600 font-bold uppercase text-[9px]">Löschen</button>
                                        </td>
                                    </tr>`).join('')}
                                ${contacts.length === 0 ? '<tr><td colspan="6" class="p-10 text-center text-slate-300 font-bold uppercase tracking-widest">Keine Kontakte hinterlegt</td></tr>' : ''}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div id="editSide" class="fixed inset-y-0 right-0 w-[400px] bg-white shadow-2xl z-50 transform translate-x-full transition-transform p-8 border-l overflow-y-auto">
                <div class="flex justify-between items-center mb-8 border-b pb-4"><h3 class="font-black uppercase text-sm">Stammdaten bearbeiten</h3><button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')">✕</button></div>
                <div class="space-y-4">
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Firmenname</label><input type="text" id="e_t" value="${f.Title}" class="w-full p-2 border rounded-lg text-sm font-bold mb-4">
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Klassifizierung</label><select id="e_k" class="w-full p-2 border rounded-lg text-sm">${classOptions.map(o => `<option value="${o}" ${f.Klassifizierung===o?'selected':''}>${o}</option>`).join('')}</select>
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Zentrale</label><input type="text" id="e_p" value="${f.Hauptnummer||''}" class="w-full p-2 border rounded-lg text-sm">
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Standort</label><input type="text" id="e_s" value="${f.Adresse||''}" class="w-full p-2 border rounded-lg text-sm mb-2"><div class="flex gap-2"><input type="text" id="e_c" value="${f.Ort}" class="flex-1 p-2 border rounded-lg text-sm font-bold"><input type="text" id="e_l" value="${f.Land||'CH'}" class="w-16 p-2 bg-slate-100 border rounded-lg text-xs text-center font-bold"></div>
                    <button onclick="update('${id}')" class="w-full bg-blue-600 text-white py-3 rounded-lg font-bold text-xs uppercase tracking-widest mt-8 shadow-lg hover:bg-blue-700 transition-all">Speichern</button>
                    <button onclick="deleteFirm('${id}','${f.Title}')" class="w-full text-red-400 text-[9px] font-bold uppercase mt-6 opacity-50 hover:opacity-100">Firma aus SharePoint löschen</button>
                </div>
            </div>
        </div>`;
}

// --- LOGIK & FILTER ---
function filterFirms(q) {
    const ql = q.toLowerCase();
    const f = allFirms.filter(x => x.fields.Title?.toLowerCase().includes(ql) || x.fields.Ort?.toLowerCase().includes(ql));
    document.getElementById('firmTableBody').innerHTML = f.map(item => `
        <tr onclick="renderDetail('${item.id}')" class="hover:bg-blue-50 cursor-pointer transition-colors">
            <td class="p-4 font-bold text-slate-700">${item.fields.Title} ${item.fields.VIP ? '⭐' : ''}</td>
            <td class="p-4 text-slate-500">${item.fields.Ort || '-'}</td>
            <td class="p-4 text-center"><span class="px-2 py-1 bg-slate-100 rounded text-[10px] font-bold">${item.fields.Klassifizierung || '-'}</span></td>
            <td class="p-4 text-center text-blue-600 font-bold">${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length}</td>
        </tr>`).join('');
}

function filterContacts(q) {
    const ql = q.toLowerCase();
    const f = allContacts.filter(x => x.fields.Title?.toLowerCase().includes(ql) || x.fields.Vorname?.toLowerCase().includes(ql));
    document.getElementById('contactTableBody').innerHTML = f.map(c => contactRow(c, true)).join('');
}

async function update(id) {
    const f = { Title: document.getElementById('e_t').value, Klassifizierung: document.getElementById('e_k').value, Adresse: document.getElementById('e_s').value, Ort: document.getElementById('e_c').value, Land: document.getElementById('e_l').value, Hauptnummer: document.getElementById('e_p').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}/fields`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify(f) });
    loadData();
}

async function saveNewFirm() {
    const n = document.getElementById('new_fName').value; if(!n) return;
    const fields = { Title: n, Ort: document.getElementById('new_fCity').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    toggleAdd(); loadData();
}

function toggleAdd() { document.getElementById('addForm').classList.toggle('hidden'); }
async function addContact(id) { 
    const ln = prompt("Nachname (Pflicht):"); if (!ln) return;
    const vn = prompt("Vorname:");
    const rol = prompt("Rolle (z.B. GL, HR, Marketing):");
    const sgf = prompt("SGF (z.B. bst, bsc, tbs):");
    const fields = { Title: ln, Vorname: vn, Rolle: rol, SGF: sgf, FirmaLookupId: id };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

async function deleteContact(cid, fid) { if(!confirm("Kontakt löschen?")) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cid}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
async function deleteFirm(id, n) { if(!confirm(`Firma ${n} löschen?`)) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
