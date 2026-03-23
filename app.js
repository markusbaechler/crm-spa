const msalConfig = {
    auth: {
        clientId: "YOUR_CLIENT_ID",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

let state = {
    firms: [],
    contacts: [],
    tasks: [],
    account: null
};

async function login() {
    const response = await msalInstance.loginPopup({
        scopes: ["User.Read", "Sites.ReadWrite.All"]
    });

    state.account = response.account;
    document.getElementById("login").classList.add("hidden");
    document.getElementById("app").classList.remove("hidden");

    loadData();
}

async function getToken() {
    const response = await msalInstance.acquireTokenSilent({
        scopes: ["Sites.ReadWrite.All"],
        account: state.account
    });
    return response.accessToken;
}

async function graphGet(url) {
    const token = await getToken();
    const res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` }
    });
    return res.json();
}

async function graphPost(url, body) {
    const token = await getToken();
    await fetch(url, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });
}

async function graphPatch(url, body) {
    const token = await getToken();
    await fetch(url, {
        method: "PATCH",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });
}

// ⚠️ HIER DEINE SITE + LISTEN ANPASSEN
const SITE = "YOUR_SITE_ID";
const LIST_FIRMS = "CRMFirms";
const LIST_CONTACTS = "CRMContacts";
const LIST_TASKS = "Tasks";

async function loadData() {
    try {
        state.firms = (await graphGet(`/sites/${SITE}/lists/${LIST_FIRMS}/items?expand=fields`)).value;
        state.contacts = (await graphGet(`/sites/${SITE}/lists/${LIST_CONTACTS}/items?expand=fields`)).value;
        state.tasks = (await graphGet(`/sites/${SITE}/lists/${LIST_TASKS}/items?expand=fields`)).value;

        navigate("dashboard");
    } catch (e) {
        alert("Fehler beim Laden der Daten");
        console.error(e);
    }
}

function navigate(view) {
    if (view === "dashboard") renderDashboard();
    if (view === "firmen") renderFirms();
    if (view === "kontakte") renderContacts();
    if (view === "planung") renderPlanning();
}

function renderDashboard() {
    document.getElementById("content").innerHTML = `
        <h1 class="text-2xl font-bold mb-4">Dashboard</h1>

        <div class="grid grid-cols-3 gap-4">
            <div class="bg-white p-4 rounded shadow">
                <div class="text-gray-500">Firmen</div>
                <div class="text-2xl">${state.firms.length}</div>
            </div>
            <div class="bg-white p-4 rounded shadow">
                <div class="text-gray-500">Kontakte</div>
                <div class="text-2xl">${state.contacts.length}</div>
            </div>
            <div class="bg-white p-4 rounded shadow">
                <div class="text-gray-500">Tasks</div>
                <div class="text-2xl">${state.tasks.length}</div>
            </div>
        </div>
    `;
}

function renderFirms() {
    document.getElementById("content").innerHTML = `
        <h1 class="text-xl font-bold mb-4">Firmen</h1>
        <button onclick="openFirmForm()" class="bg-blue-600 text-white px-4 py-2 mb-4">Neue Firma</button>

        ${state.firms.map(f => `
            <div class="bg-white p-3 mb-2 rounded shadow">
                ${f.fields.Title || "-"}
            </div>
        `).join("")}
    `;
}

function renderContacts() {
    document.getElementById("content").innerHTML = `
        <h1 class="text-xl font-bold mb-4">Kontakte</h1>

        ${state.contacts.map(c => `
            <div class="bg-white p-3 mb-2 rounded shadow">
                ${c.fields.Title || "-"}
            </div>
        `).join("")}
    `;
}

function renderPlanning() {
    document.getElementById("content").innerHTML = `
        <h1 class="text-xl font-bold mb-4">Planung</h1>
        <p>Kommt in Sprint 5</p>
    `;
}

// MODAL
function openModal(html) {
    document.getElementById("modal").classList.remove("hidden");
    document.getElementById("modalContent").innerHTML = html;
}

function closeModal() {
    document.getElementById("modal").classList.add("hidden");
}

// FORM
function openFirmForm() {
    openModal(`
        <h2 class="text-lg font-bold mb-2">Neue Firma</h2>
        <input id="firmName" class="border p-2 w-full" placeholder="Name">
        <button onclick="saveFirm()" class="bg-blue-600 text-white px-4 py-2 mt-3">Speichern</button>
    `);
}

async function saveFirm() {
    const name = document.getElementById("firmName").value;

    await graphPost(`/sites/${SITE}/lists/${LIST_FIRMS}/items`, {
        fields: { Title: name }
    });

    closeModal();
    loadData();
}
