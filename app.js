// --- CONFIG & VERSION ---
const appVersion = "0.51";
const appName = "CRM bbz";

const config = { 
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a", 
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7", 
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM" 
};

// State Management
let allFirms = [], allContacts = [], currentSiteId = "", currentListId = "", contactListId = "";
let meta = { klassen: [], anreden: [], rollen: [], leads: [], sgf: [], events: [] };

const msalConfig = { auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, redirectUri: "https://markusbaechler.github.io/crm-spa/" }, cache: { cacheLocation: "localStorage" } };
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"] };

window.onload = async () => { 
    if(document.getElementById('footer-text')) document.getElementById('footer-text').innerHTML = `© 2026 ${appName} | <b>Build ${appVersion}</b>`;
    await msalInstance.handleRedirectPromise(); 
    checkAuthState(); 
};

function checkAuthState() {
    const acc = msalInstance.getAllAccounts();
    if (acc.length > 0) {
        document.getElementById('authBtn').innerText = "Logout";
        document.getElementById('authBtn').onclick = () => msalInstance.logoutRedirect({ account: acc[0] });
        loadData();
    } else {
        document.getElementById('authBtn').innerText = "Login";
        document.getElementById('authBtn').onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('main-content');
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse italic">Lade CRM bbz Datenkern...</div>`;
    try {
        const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
        const h = { 'Authorization': `Bearer ${token}` };
        
        const s = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers: h })).json();
        currentSiteId = s.id;
        const l = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers: h })).json();
        currentListId = l.value.find(x => x.displayName === "CRMFirms").id;
        contactListId = l.value.find(x => x.displayName === "CRMContacts").id;

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
            klassen: cKlass.choice?.choices || [], anreden: cAnrede.choice?.choices || [],
            rollen: cRolle.choice?.choices || [], leads: cLead.choice?.choices || [],
            sgf: cSGF.choice?.choices || [], events: cEvent.choice?.choices || []
        };
        allFirms = fData.value; allContacts = cData.value;
        renderAllContacts(); // Starten direkt in der Kontakt-Masterliste
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500 font-bold italic underline uppercase">Sync-Fehler: ${e.message}</div>`; }
}

// --- MASTER-VIEW: ALLE KONTAKTE (2-ZEILEN-DESIGN) ---
function renderAllContacts() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">Kontakte</h2>
                    <nav class="flex gap-6 text-xs font-black uppercase tracking-widest pb-1">
                        <button onclick="renderFirms()" class="text-slate-300 hover:text-slate-600 transition">Firmen</button>
                        <button class="text-blue-600 border-b-2 border-blue-600 px-1">Alle Kontakte</button>
                    </nav>
                </div>
                <div class="relative">
                    <input type="text" onkeyup="filterContacts(this.value)" placeholder="Person, Firma oder SGF suchen..." class="pl-10 pr-4 py-3 bg-slate-50 border-none rounded-2xl text-sm w-96 shadow-inner outline-none focus:ring-2 focus:ring-blue-500/20 font-bold italic">
                    <span class="absolute left-3 top-3.5 opacity-30">🔍</span>
                </div>
            </div>

            <div class="bg-white border border-slate-200 rounded-[2.5rem] overflow-hidden shadow-sm">
                <table class="w-full text-left">
                    <thead class="bg-slate-50 border-b text-[10px] uppercase font-black text-slate-400 tracking-[0.2em] italic">
                        <tr>
                            <th class="p-6">Name / Firma</th>
                            <th class="p-6">Rolle / SGF / Lead</th>
                            <th class="p-6">Event-Historie</th>
                            <th class="p-6">Kontaktinfo</th>
                        </tr>
                    </thead>
                    <tbody id="contactTableBody" class="divide-y divide-slate-50 italic">
                        ${allContacts.map(c => contactRow(c)).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function contactRow(c) {
    const f = c.fields;
    const firm = allFirms.find(x => String(x.id) === String(f.FirmaLookupId));
    
    // Event-Badges bauen
    const events = f.Event ? (Array.isArray(f.Event) ? f.Event : [f.Event]) : [];
    const eventBadges = events.map(e => `<span class="px-2 py-0.5 bg-blue-50 text-blue-600 border border-blue-100 rounded text-[9px] font-black uppercase tracking-tighter">${e}</span>`).join(' ');

    return `
        <tr class="hover:bg-blue-50/30 transition-all group">
            <td class="p-6">
                <div class="text-base font-black text-slate-800 uppercase tracking-tighter group-hover:text-blue-600 cursor-pointer underline" onclick="renderContactDetail('${c.id}')">
                    ${f.Vorname || ''} ${f.Title}
                </div>
                <div class="text-[11px] font-bold text-slate-400 mt-1 cursor-pointer hover:text-slate-600" onclick="renderDetail('${f.FirmaLookupId}')">
                    🏢 ${firm ? firm.fields.Title : 'Keine Firma zugeordnet'}
                </div>
            </td>
            <td class="p-6">
                <div class="flex flex-col gap-1.5">
                    <div class="text-[10px] font-black text-blue-500 uppercase tracking-widest">${f.Rolle || '-'}</div>
                    <div class="flex gap-2">
                        <span class="text-[9px] font-black bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded border border-emerald-100 uppercase italic tracking-tighter">${f.SGF || 'SGF ?'}</span>
                        <span class="text-[9px] font-black bg-slate-100 text-slate-500 px-2 py-0.5 rounded border border-slate-200 uppercase italic tracking-tighter">Lead: ${f.Leadbbz0 || '-'}</span>
                    </div>
                </div>
            </td>
            <td class="p-6">
                <div class="flex flex-wrap gap-1.5 max-w-xs">
                    ${eventBadges || '<span class="text-slate-200 text-[10px] font-bold uppercase">Keine Events</span>'}
                </div>
            </td>
            <td class="p-6">
                <div class="space-y-1 text-[11px] font-bold text-slate-500">
                    <div class="flex items-center gap-2">📧 <span class="text-slate-400 font-medium lowercase tracking-tight italic">${f.Email1 || '-'}</span></div>
                    <div class="flex items-center gap-2">📱 <span class="text-slate-800">${f.Mobile || '-'}</span></div>
                </div>
            </td>
        </tr>`;
}

// --- FILTER ---
function filterContacts(q) {
    const ql = q.toLowerCase();
    const filtered = allContacts.filter(c => {
        const firm = allFirms.find(x => String(x.id) === String(c.fields.FirmaLookupId));
        return (c.fields.Title?.toLowerCase().includes(ql) || 
                c.fields.Vorname?.toLowerCase().includes(ql) || 
                firm?.fields.Title?.toLowerCase().includes(ql) ||
                c.fields.SGF?.toLowerCase().includes(ql));
    });
    document.getElementById('contactTableBody').innerHTML = filtered.map(c => contactRow(c)).join('');
}

// --- FIRMENÜBERSICHT (Wiederhergestellt & Stabil) ---
function renderFirms() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter cursor-pointer" onclick="renderFirms()">Firmen</h2>
                    <nav class="flex gap-6 text-xs font-black uppercase tracking-widest pb-1">
                        <button class="text-blue-600 border-b-2 border-blue-600 px-1">Übersicht</button>
                        <button onclick="renderAllContacts()" class="text-slate-300 hover:text-slate-600 transition">Alle Kontakte</button>
                    </nav>
                </div>
                <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-blue-100 hover:scale-105 transition-all">+ Firma</button>
            </div>
            <div id="firmList" class="grid grid-cols-1 md:grid-cols-3 gap-8">
                ${allFirms.map(f => `
                    <div onclick="renderDetail('${f.id}')" class="bg-white border border-slate-200 p-8 rounded-[2.5rem] hover:border-blue-400 hover:shadow-2xl transition-all cursor-pointer flex flex-col h-full border-t-[6px] ${f.fields.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                        <h3 class="font-black text-slate-800 text-xl uppercase italic tracking-tighter mb-4 underline">${f.fields.Title}</h3>
                        <div class="text-[11px] font-bold text-slate-400 uppercase tracking-[0.2em] italic mb-8">📍 ${f.fields.Ort || '-'}</div>
                        <div class="mt-auto pt-6 flex justify-between items-center border-t border-slate-50">
                            <span class="text-[10px] font-black text-slate-400 uppercase tracking-widest bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100 italic tracking-widest">${f.fields.Klassifizierung || '-'}</span>
                            <span class="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-xl flex items-center gap-2 italic uppercase">👥 ${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</span>
                        </div>
                    </div>`).join('')}
            </div>
        </div>`;
}

// (Hier folgen die restlichen Funktionen: renderDetail, renderContactDetail, saveContact, etc.)
// ... aus Build 0.50 übernehmen ...
