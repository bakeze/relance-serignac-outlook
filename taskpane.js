const STORAGE_KEY = "relanceIA.clients.v1";
const STATUS = {
  A_RELANCER: "À relancer",
  RELANCE: "Relancé",
  EN_ATTENTE: "En attente",
  CLOTURE: "Clôturé"
};

const TEMPLATES = {
  douce: ({ nom, days }) => ({
    subject: `Relance douce - ${nom}`,
    htmlBody: `<p>Bonjour ${nom},</p>
      <p>Je me permets de revenir vers vous suite à mon précédent message envoyé il y a ${days} jours.</p>
      <p>Je reste à votre disposition si vous souhaitez échanger.</p>
      <p>Bien cordialement,</p>`
  }),
  standard: ({ nom, days }) => ({
    subject: `Relance - suivi de dossier ${nom}`,
    htmlBody: `<p>Bonjour ${nom},</p>
      <p>Sans retour de votre part depuis ${days} jours, je me permets une relance concernant notre échange.</p>
      <p>Pouvez-vous me confirmer votre positionnement ?</p>
      <p>Merci d'avance,</p>`
  }),
  ferme: ({ nom, days }) => ({
    subject: `Relance prioritaire - action requise (${nom})`,
    htmlBody: `<p>Bonjour ${nom},</p>
      <p>Je reviens vers vous après ${days} jours sans réponse.</p>
      <p>Sans retour de votre part, je classerai ce dossier d'ici 48h.</p>
      <p>Cordialement,</p>`
  })
};

let clients = [];
let queue = [];

Office.onReady(() => {
  hydrate();
  bindEvents();
  render();
});

function bindEvents() {
  document.getElementById("excelInput").addEventListener("change", importExcel);
  document.getElementById("manualRelanceBtn").addEventListener("click", manualRelance);
  document.getElementById("autoRelanceBtn").addEventListener("click", buildAutoQueue);
  document.getElementById("launchQueueBtn").addEventListener("click", launchAutoRelance);
}

function hydrate() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    clients = JSON.parse(saved);
    return;
  }

  clients = [
    sampleClient("Société Atlas", "contact@atlas.fr", 4, 0, STATUS.A_RELANCER),
    sampleClient("Groupe Nova", "achats@nova.fr", 8, 1, STATUS.EN_ATTENTE),
    sampleClient("Cabinet Horizon", "admin@horizon.fr", 15, 2, STATUS.A_RELANCER)
  ];
  persist();
}

function sampleClient(nom, email, daysAgo, nbRelances, statut) {
  const dt = new Date();
  dt.setDate(dt.getDate() - daysAgo);
  return {
    id: crypto.randomUUID(),
    nom,
    email,
    sujet: `Suivi commercial - ${nom}`,
    dernierContact: dt.toISOString().slice(0, 10),
    nbRelances,
    statut
  };
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(clients));
}

function importExcel(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    clients = rows.map((r) => ({
      id: crypto.randomUUID(),
      nom: r.nom || r.Nom || "Client sans nom",
      email: r.email || r.Email || "",
      sujet: r.sujet || r.Sujet || "Relance commerciale",
      dernierContact: normalizeDate(r.dernierContact || r.DernierContact),
      nbRelances: Number(r.nbRelances || r.NbRelances || 0),
      statut: normalizeStatus(r.statut || r.Statut)
    }));

    persist();
    render();
    setStatus(`${clients.length} clients importés depuis Excel.`);
  };

  reader.readAsArrayBuffer(file);
}

function normalizeDate(value) {
  if (!value) return new Date().toISOString().slice(0, 10);
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return new Date().toISOString().slice(0, 10);
  return parsed.toISOString().slice(0, 10);
}

function normalizeStatus(statusRaw) {
  const v = String(statusRaw).trim().toLowerCase();
  if (v.includes("relancé")) return STATUS.RELANCE;
  if (v.includes("attente")) return STATUS.EN_ATTENTE;
  if (v.includes("clôturé") || v.includes("cloture")) return STATUS.CLOTURE;
  return STATUS.A_RELANCER;
}

function getSelectedClient() {
  const selected = document.querySelector("input[name='clientRadio']:checked");
  if (!selected) return null;
  return clients.find((c) => c.id === selected.value) || null;
}

function manualRelance() {
  const client = getSelectedClient();
  if (!client) {
    setStatus("Sélectionnez un client avant de relancer.");
    return;
  }

  openDraftForClient(client);
  markAsRelanced(client.id);
  render();
  setStatus(`Brouillon ouvert pour ${client.nom}.`);
}

function buildAutoQueue() {
  queue = clients.filter((client) => shouldRelance(client).needed);
  setStatus(`File de relance créée: ${queue.length} client(s) à relancer.`);
  render();
}

async function launchAutoRelance() {
  if (!queue.length) {
    buildAutoQueue();
    if (!queue.length) {
      setStatus("Aucun client éligible à la relance actuellement.");
      return;
    }
  }

  // Simulation intelligente : ouverture séquentielle de brouillons prêts à valider.
  for (const client of queue) {
    openDraftForClient(client);
    markAsRelanced(client.id);
    await sleep(500);
  }

  queue = [];
  persist();
  render();
  setStatus("Auto-relance assistée terminée : tous les brouillons ont été préparés.");
}

function openDraftForClient(client) {
  const decision = shouldRelance(client);
  const template = TEMPLATES[decision.level]({ nom: client.nom, days: decision.daysSinceLast });

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: [client.email],
    subject: template.subject || client.sujet,
    htmlBody: `${template.htmlBody}<hr/><p><em>Statut IA: ${decision.reason}</em></p>`
  });
}

function shouldRelance(client) {
  const daysSinceLast = diffDays(client.dernierContact);
  const nbRelances = Number(client.nbRelances || 0);

  if (client.statut === STATUS.CLOTURE) {
    return { needed: false, level: "douce", daysSinceLast, reason: "Dossier clôturé" };
  }

  if (daysSinceLast >= 14 || nbRelances >= 2) {
    return { needed: true, level: "ferme", daysSinceLast, reason: "Relance ferme J+14" };
  }
  if (daysSinceLast >= 7) {
    return { needed: true, level: "standard", daysSinceLast, reason: "Relance standard J+7" };
  }
  if (daysSinceLast >= 3) {
    return { needed: true, level: "douce", daysSinceLast, reason: "Relance douce J+3" };
  }

  return { needed: false, level: "douce", daysSinceLast, reason: "Pas encore nécessaire" };
}

function diffDays(isoDate) {
  const last = new Date(isoDate);
  const now = new Date();
  return Math.floor((now - last) / (1000 * 60 * 60 * 24));
}

function markAsRelanced(id) {
  clients = clients.map((c) => {
    if (c.id !== id) return c;
    return {
      ...c,
      statut: STATUS.RELANCE,
      nbRelances: Number(c.nbRelances || 0) + 1,
      dernierContact: new Date().toISOString().slice(0, 10)
    };
  });
  persist();
}

function render() {
  const tbody = document.getElementById("clientsTbody");
  tbody.innerHTML = "";

  clients.forEach((client, idx) => {
    const decision = shouldRelance(client);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input type="radio" name="clientRadio" value="${client.id}" ${idx === 0 ? "checked" : ""}></td>
      <td>${escapeHtml(client.nom)}</td>
      <td>${escapeHtml(client.email)}</td>
      <td>${escapeHtml(client.dernierContact)}</td>
      <td>${Number(client.nbRelances || 0)}</td>
      <td>${renderStatus(client.statut)}</td>
      <td class="${decision.needed ? "note" : ""}">${decision.reason}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderStatus(statut) {
  const cls = statut === STATUS.A_RELANCER ? "status-relance"
    : statut === STATUS.RELANCE ? "status-fait"
    : statut === STATUS.EN_ATTENTE ? "status-attente"
    : "status-cloture";

  return `<span class="status-pill ${cls}">${escapeHtml(statut)}</span>`;
}

function setStatus(message) {
  document.getElementById("statusLine").textContent = message;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function escapeHtml(v) {
  return String(v)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

/**
 * BONUS - Envoi automatique réel via Microsoft Graph API
 * ------------------------------------------------------
 * 1) Créer un backend (Node/Azure Function) avec auth Azure AD (client credentials).
 * 2) Le taskpane envoie une requête sécurisée au backend avec les relances à déclencher.
 * 3) Le backend appelle Graph: POST /users/{id}/sendMail
 *    payload: { message: { subject, body, toRecipients }, saveToSentItems: true }
 * 4) Stocker les logs d'envoi et gérer retries / throttling.
 *
 * Important : l'envoi 100% auto n'est pas autorisé directement depuis Office.js côté add-in.
 * Il faut un service backend + consentement Microsoft Graph.
 */
