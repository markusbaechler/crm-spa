// --- CONFIG & VERSION ---
const appVersion = "0.47";
const appName = "CRM bbz";

const config = { 
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a", 
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7", 
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM" 
};

let allFirms = [], allContacts = [], currentSiteId = "", currentListId = "", contactListId = "";
let meta = { klassen: [], anreden: [], rollen: [], leads: [], sgf: [], events: [] };

const msalConfig = { auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, redirectUri: "https://markusbaechler.github.io/crm-spa/" }, cache: { cacheLocation: "localStorage" } };
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"] };

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
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse">Lade Systemintelligenz...</div>`;
    try {
        const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
        const h = { 'Authorization': `Bearer ${token}` };
        
        const s = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers: h })).json();
        currentSiteId = s.id;
        const l = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers: h })).json();
        currentListId = l.value.find(x => x.displayName === "CRMFirms").id;
        contactListId = l.value.find(x => x.displayName === "CRMContacts").id;

        // Metadaten & Daten parallel laden
        const [cKlass, cAnrede, cRolle, cLead, cSGF, cEvent, fData, cData] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Anrede`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Rolle`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Leadbbz0`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/SGF`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Event`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers: h }).then(r => r.json())
        ]);

        meta = {
            klassen: cKlass.choice?.choices || [],
            anreden: cAnrede.choice?.choices || [],
            rollen: cRolle.choice?.choices || [],
            leads: cLead.choice?.choices || [],
            sgf: cSGF.choice?.choices || [],
            events: cEvent.choice?.choices || []
        };

        allFirms = fData.value; 
        allContacts = cData.value;
        renderFirms();
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500 font-bold">CRITICAL ERROR: ${e.message}</div>`; }
}

// --- INTELLIGENTE NAVIGATION ---
function renderFirms() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-6 border-b pb-4">
                <div class="flex gap-8 items-end">
                    <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter cursor-pointer" onclick="renderFirms()">Firmen</h2>
                    <nav class="flex gap-4 text-xs font-bold uppercase tracking-widest pb-1">
                        <button onclick="renderFirms()" class="text-blue-600 border-b-2 border-blue-600 px-2">Übersicht</button>
                        <button onclick="renderAllContacts()" class="text-slate-400 hover:text-slate-600 px-2">Alle Kontakte</button>
                    </nav>
                </div>
                <div class="flex gap-3">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Firma suchen..." class="px-4 py-2 border rounded-lg text-sm w-64 shadow-sm">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-xs font-bold uppercase">+ Firma</button>
                </div>
            </div>
            
            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 border rounded-2xl shadow-inner grid grid-cols-1 md:grid-cols-3 gap-4">
                <input type="text" id="new_fName" placeholder="Firmenname *" class="p-2 border rounded text-sm">
                <select id="new_fClass" class="p-2 border rounded text-sm"><option value="">Klasse</option>${meta.klassen.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                <input type="text" id="new_fCity" placeholder="Ort" class="p-2 border rounded text-sm">
                <button onclick="saveNewFirm()" class="bg-green-600 text-white font-bold text-xs p-2 rounded uppercase">Speichern</button>
            </div>

            <div class="bg-white border rounded-xl overflow-hidden shadow-sm">
                <table class="w-full text-sm text-left">
                    <thead class="bg-slate-50 border-b text-[10px] uppercase font-black text-slate-500 tracking-widest">
                        <tr><th class="p-4">Firma</th><th class="p-4">Ort</th><th class="p-4 text-center">Klasse</th><th class="p-4 text-center">Kontakte</th></tr>
                    </thead>
                    <tbody id="firmTableBody" class="divide-y divide-slate-100">
                        ${allFirms.map(f => `<tr onclick="renderDetail('${f.id}')" class="hover:bg-blue-50 cursor-pointer transition-colors">
                            <td class="p-4 font-bold text-slate-700">${f.fields.Title} ${f.fields.VIP ? '⭐' : ''}</td>
                            <td class="p-4 text-slate-500">${f.fields.Ort || '-'}</td>
                            <td class="p-4 text-center font-bold text-[10px]">${f.fields.Klassifizierung || '-'}</td>
                            <td class="p-4 text-center text-blue-600 font-bold">${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</td>
                        </tr>`).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function renderAllContacts() {
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
             <div class="flex justify-between items-center mb-6 border-b pb-4">
                <div class="flex gap-8 items-end">
                    <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter">Kontakte</h2>
                    <nav class="flex gap-4 text-xs font-bold uppercase tracking-widest pb-1">
                        <button onclick="renderFirms()" class="text-slate-400 hover:text-slate-600 px-2">Firmen</button>
                        <button onclick="renderAllContacts()" class="text-blue-600 border-b-2 border-blue-600 px-2">Alle Kontakte</button>
                    </nav>
                </div>
                <input type="text" onkeyup="filterContacts(this.value)" placeholder="Person suchen..." class="px-4 py-2 border rounded-lg text-sm w-80">
            </div>
            <div class="bg-white border rounded-xl overflow-hidden shadow-sm">
                <table class="w-full text-[11px] text-left">
                    <thead class="bg-slate-50 border-b text-[9px] uppercase font-black text-slate-500 tracking-widest">
                        <tr><th class="p-3">Name</th><th class="p-3">Firma</th><th class="p-3">Rolle</th><th class="p-3">SGF</th><th class="p-3">Kontaktinfo</th></tr>
                    </thead>
                    <tbody id="contactTableBody" class="divide-y divide-slate-100">
                        ${allContacts.map(c => {
                            const firm = allFirms.find(x => String(x.id) === String(c.fields.FirmaLookupId));
                            return `<tr class="hover:bg-slate-50">
                                <td class="p-3 font-bold text-slate-800 text-sm">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                                <td class="p-3 font-bold text-blue-600 hover:underline cursor-pointer" onclick="renderDetail('${c.fields.FirmaLookupId}')">${firm ? firm.fields.Title : '-'}</td>
                                <td class="p-3 font-bold text-slate-500">${c.fields.Rolle || '-'}</td>
                                <td class="p-3 font-black text-emerald-600 uppercase">${c.fields.SGF || '-'}</td>
                                <td class="p-3 text-slate-400 font-medium">📧 ${c.fields.Email1 || '-'} | 📱 ${c.fields.Mobile || '-'}</td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function renderDetail(id) {
    const firm = allFirms.find(x => String(x.id) === String(id)), f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(id));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right">
            <div class="bg-white border rounded-xl p-6 mb-6 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms()" class="bg-slate-100 p-2 rounded hover:bg-slate-200 text-xl">←</button>
                    <div>
                        <h2 class="text-2xl font-black text-slate-800 uppercase tracking-tighter">${f.Title}</h2>
                        <div class="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">📍 ${f.Ort} | Klasse ${f.Klassifizierung || 'LEER'}</div>
                    </div>
                </div>
                <button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')" class="bg-slate-800 text-white px-5 py-2 rounded text-[10px] font-bold uppercase shadow-lg">Stammdaten</button>
            </div>

            <div class="bg-white border rounded-xl shadow-sm overflow-hidden">
                <div class="p-4 bg-slate-50 border-b flex justify-between items-center">
                    <h3 class="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">Ansprechpartner (${contacts.length})</h3>
                    <button onclick="toggleContactForm()" class="bg-blue-600 text-white px-3 py-1.5 rounded text-[9px] font-bold uppercase tracking-widest">+ Kontakt</button>
                </div>

                <div id="addContactForm" class="hidden p-8 bg-blue-50/50 border-b animate-in fade-in">
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">Anrede</label>
                            <select id="c_anrede" class="w-full p-2 border rounded bg-white text-sm">${meta.anreden.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                        </div>
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">Vorname</label><input type="text" id="c_vn" class="w-full p-2 border rounded text-sm"></div>
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">Nachname *</label><input type="text" id="c_nn" class="w-full p-2 border rounded text-sm font-bold"></div>
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">Rolle</label>
                            <select id="c_rolle" class="w-full p-2 border rounded bg-white text-sm">${meta.rollen.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                        </div>
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">Funktion</label><input type="text" id="c_fun" class="w-full p-2 border rounded text-sm"></div>
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">E-Mail</label><input type="text" id="c_email" class="w-full p-2 border rounded text-sm"></div>
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">SGF</label>
                            <select id="c_sgf" class="w-full p-2 border rounded bg-white text-sm uppercase"><option value="">-</option>${meta.sgf.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                        </div>
                        <div><label class="text-[9px] font-bold uppercase text-slate-400">Lead bbz</label>
                            <select id="c_lead" class="w-full p-2 border rounded bg-white text-sm uppercase"><option value="">-</option>${meta.leads.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                        </div>
                    </div>
                    <div class="flex gap-3">
                        <button onclick="saveContact('${id}')" class="bg-blue-600 text-white font-bold text-xs px-8 py-3 rounded uppercase tracking-widest shadow-md">In SharePoint speichern</button>
                        <button onclick="toggleContactForm()" class="text-slate-400 font-bold text-xs px-4">Abbrechen</button>
                    </div>
                </div>

                <table class="w-full text-left text-[11px]">
                    <thead class="text-[9px] uppercase font-black text-slate-400 border-b bg-white italic">
                        <tr><th class="p-4">Name</th><th class="p-4">Rolle / Funktion</th><th class="p-4">SGF</th><th class="p-4">Kontaktinfo</th><th class="p-4 text-right">Aktion</th></tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        ${contacts.map(c => `<tr class="hover:bg-slate-50 transition-colors font-medium">
                            <td class="p-4 font-bold text-slate-800 text-sm">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                            <td class="p-4 font-bold text-blue-600 uppercase text-[9px]">${c.fields.Rolle || '-'}</td>
                            <td class="p-4 font-black uppercase text-emerald-600">${c.fields.SGF || '-'}</td>
                            <td class="p-4 text-slate-400">📧 ${c.fields.Email1 || '-'} | 📱 ${c.fields.Mobile || '-'}</td>
                            <td class="p-4 text-right"><button onclick="deleteContact('${c.id}','${id}')" class="text-red-300 hover:text-red-600 font-bold uppercase text-[9px]">Löschen</button></td>
                        </tr>`).join('')}
                    </tbody>
                </table>
            </div>

            <div id="editSide" class="fixed inset-y-0 right-0 w-[400px] bg-white shadow-2xl z-50 transform translate-x-full transition-transform p-8 border-l italic">
                 <div class="flex justify-between items-center mb-8 border-b pb-4"><h3 class="font-black uppercase text-sm">Profil</h3><button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')">✕</button></div>
                 <div class="space-y-4">
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Firma</label><input type="text" id="e_t" value="${f.Title}" class="w-full p-2 border rounded-lg text-sm font-bold">
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Klassifizierung</label><select id="e_k" class="w-full p-2 border rounded-lg text-sm font-bold">${meta.klassen.map(o => `<option value="${o}" ${f.Klassifizierung===o?'selected':''}>${o}</option>`).join('')}</select>
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Hauptnummer</label><input type="text" id="e_p" value="${f.Hauptnummer||''}" class="w-full p-2 border rounded-lg text-sm">
                    <label class="block text-[10px] font-bold text-slate-400 uppercase">Adresse</label><input type="text" id="e_s" value="${f.Adresse||''}" class="w-full p-2 border rounded-lg text-sm mb-2"><div class="flex gap-2"><input type="text" id="e_c" value="${f.Ort}" class="flex-1 p-2 border rounded-lg text-sm font-bold"><input type="text" id="e_l" value="${f.Land||'CH'}" class="w-16 p-2 bg-slate-100 border rounded-lg text-xs text-center font-black"></div>
                    <button onclick="updateFirmData('${id}')" class="w-full bg-slate-900 text-white py-3 rounded-lg font-black text-xs uppercase tracking-widest mt-8 shadow-xl">Speichern</button>
                 </div>
            </div>
        </div>`;
}

// --- LOGIK-ACTIONS (Persistente Intelligenz) ---
async function saveContact(firmId) {
    const nn = document.getElementById('c_nn').value; if(!nn) return alert("Nachname fehlt!");
    const fields = { Title: nn, Vorname: document.getElementById('c_vn').value, Anrede: document.getElementById('c_anrede').value, Rolle: document.getElementById('c_rolle').value, Funktion: document.getElementById('c_fun').value, Email1: document.getElementById('c_email').value, SGF: document.getElementById('c_sgf').value, Leadbbz0: document.getElementById('c_lead').value, FirmaLookupId: firmId };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

function filterFirms(q) { const ql = q.toLowerCase(); const f = allFirms.filter(x => x.fields.Title?.toLowerCase().includes(ql) || x.fields.Ort?.toLowerCase().includes(ql)); document.getElementById('firmTableBody').innerHTML = f.map(f => `<tr onclick="renderDetail('${f.id}')" class="hover:bg-blue-50 cursor-pointer transition-colors"><td class="p-4 font-bold text-slate-700">${f.fields.Title}</td><td class="p-4 text-slate-500">${f.fields.Ort || '-'}</td><td class="p-4 text-center font-bold text-[10px]">${f.fields.Klassifizierung || '-'}</td><td class="p-4 text-center text-blue-600 font-bold">${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</td></tr>`).join(''); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function toggleContactForm() { document.getElementById('addContactForm').classList.toggle('hidden'); }
async function deleteContact(cid, fid) { if(!confirm("Löschen?")) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cid}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
