// --- CONFIG & VERSION ---
const appVersion = "0.36";
console.log(`CRM App ${appVersion} - DetailPage Fix & Fields`);

// ... (msalConfig & msalInstance identisch zu V0.35) ...

async function loadData() {
    const content = document.getElementById('main-content'); 
    content.innerHTML = `<div class="p-20 text-center text-slate-400 text-xs uppercase tracking-widest animate-pulse">Daten-Abgleich...</div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        
        currentListId = listsData.value.find(l => l.displayName === "CRMFirms").id;
        contactListId = listsData.value.find(l => l.displayName === "CRMContacts").id;

        const [colRes, firmsRes, contactsRes] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
        ]);
        
        classOptions = (await colRes.json()).choice?.choices || ["-"];
        allFirms = (await firmsRes.json()).value;
        allContacts = (await contactsRes.json()).value;

        renderFirms(allFirms);
    } catch (err) { content.innerHTML = `<div class="p-6 text-red-500 font-bold">Fehler: ${err.message}</div>`; }
}

// --- DETAILSEITE V0.36 ---
function renderFirmDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    
    // VERKNÜPFUNG FIX: Wir prüfen sowohl 'FirmID' als auch das Lookup-Feld 'FirmaLookupId'
    const contacts = allContacts.filter(c => 
        String(c.fields.FirmID) === String(itemId) || 
        String(c.fields.FirmaLookupId) === String(itemId)
    );
    
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in slide-in-from-right duration-300">
            <div class="flex items-center gap-4 mb-8">
                <button onclick="renderFirms(allFirms)" class="p-2 hover:bg-slate-100 rounded-full transition text-slate-400 text-xl">←</button>
                <h2 class="text-2xl font-semibold text-slate-800">${f.Title}</h2>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div class="lg:col-span-1 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                        <h3 class="text-[10px] font-bold text-slate-300 uppercase tracking-widest mb-6 border-b pb-2">Unternehmensprofil</h3>
                        <div class="space-y-4">
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Firma / Name</label>
                                <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm focus:border-blue-400 outline-none">
                            </div>
                            
                            <div class="grid grid-cols-2 gap-3">
                                <div>
                                    <label class="text-[9px] font-bold text-slate-400 uppercase">Klassierung</label>
                                    <select id="edit_Klass" class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                                        ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="flex items-end pb-1.5 pl-2">
                                    <label class="flex items-center gap-2 cursor-pointer">
                                        <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="rounded text-blue-600">
                                        <span class="text-[10px] font-bold text-slate-500 uppercase">VIP</span>
                                    </label>
                                </div>
                            </div>

                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Hauptnummer</label>
                                <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" placeholder="+41..." class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                            </div>

                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Adresse & Standort</label>
                                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse / Nr" class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm mb-2">
                                <div class="flex gap-2">
                                    <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                                    <input type="text" id="edit_Country" value="${f.Land || 'CH'}" placeholder="Land" class="w-16 p-2 bg-slate-50 border border-slate-100 rounded text-sm text-center font-medium">
                                </div>
                            </div>
                            
                            <button onclick="updateFirm('${itemId}')" class="w-full bg-blue-600 text-white py-2.5 rounded-lg text-xs font-bold uppercase tracking-widest hover:bg-blue-700 transition shadow-sm mt-4">Profil aktualisieren</button>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-2 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                        <div class="flex justify-between items-center mb-6">
                            <h3 class="text-[10px] font-bold text-slate-300 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:underline">+ Kontakt hinzufügen</button>
                        </div>
                        
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="p-4 bg-slate-50 border border-slate-100 rounded-lg flex justify-between items-center group">
                                    <div>
                                        <div class="text-sm font-semibold text-slate-700">${c.fields.FirstName || ''} ${c.fields.Title}</div>
                                        <div class="text-[10px] text-slate-400 font-medium">${c.fields.Email1 || c.fields.Email || '-'}</div>
                                        <div class="text-[9px] text-blue-500 mt-1 uppercase font-bold tracking-tighter">${c.fields.Rolle || ''}</div>
                                    </div>
                                    <button onclick="deleteContact('${c.id}', '${itemId}')" class="text-slate-200 hover:text-red-400 opacity-0 group-hover:opacity-100 transition-opacity">✕</button>
                                </div>
                            `).join('') : '<div class="col-span-2 py-10 text-center text-slate-300 text-[10px] uppercase font-bold tracking-widest border border-dashed rounded-lg">Keine Kontakte verknüpft</div>'}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
}

// REST Logik - Felder um 'Land' und 'Hauptnummer' ergänzt
async function updateFirm(itemId) {
    const fields = {
        Title: document.getElementById('edit_Title').value,
        Klassifizierung: document.getElementById('edit_Klass').value,
        Adresse: document.getElementById('edit_Street').value,
        Ort: document.getElementById('edit_City').value,
        Land: document.getElementById('edit_Country').value,
        Hauptnummer: document.getElementById('edit_Phone').value,
        VIP: document.getElementById('edit_VIP').checked
    };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}/fields`, {
        method: 'PATCH', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(fields)
    });
    await loadDataSilent();
    renderFirmDetailPage(itemId);
}
