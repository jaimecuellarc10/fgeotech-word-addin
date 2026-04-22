/* global Office, Word */

// Maps sidebar field IDs to Word Content Control tag names and API response paths.
// Update API_FIELD_MAP values if the Total Synergy API uses different field names.
// Verify actual field names via: developers.totalsynergy.com/swagger/ui/index
const FIELD_MAP = [
  { inputId: "f_project_number",      tag: "synergy_project_number",      apiPath: "projectNumber" },
  { inputId: "f_project_name",        tag: "synergy_project_name",        apiPath: "name" },
  { inputId: "f_project_status",      tag: "synergy_project_status",      apiPath: "status" },
  { inputId: "f_client_name",         tag: "synergy_client_name",         apiPath: "primaryContact" },
  { inputId: "f_client_contact",      tag: "synergy_client_contact",      apiPath: "clientReferenceNumber" },
  { inputId: "f_project_manager",     tag: "synergy_project_manager",     apiPath: "manager" },
  { inputId: "f_project_address",     tag: "synergy_project_address",     apiPath: "address.address1" },
  { inputId: "f_project_suburb",      tag: "synergy_project_suburb",      apiPath: "address.town" },
  { inputId: "f_project_state",       tag: "synergy_project_state",       apiPath: "address.state" },
  { inputId: "f_project_postcode",    tag: "synergy_project_postcode",    apiPath: "address.zipCode" },
  { inputId: "f_project_office",      tag: "synergy_project_office",      apiPath: "office" },
];

const STORAGE_KEY_API = "synergy_api_key";
const STORAGE_KEY_SETTINGS_OPEN = "synergy_settings_open";
const ORG_SLUG = "actgeotechnicalengineers";

Office.onReady(() => {
  loadStoredSettings();
});

// --- Settings ---

function toggleSettings() {
  const panel = document.getElementById("settings-panel");
  const arrow = document.getElementById("settings-arrow");
  const isOpen = panel.classList.toggle("open");
  arrow.textContent = isOpen ? "▲" : "▼";
  localStorage.setItem(STORAGE_KEY_SETTINGS_OPEN, isOpen ? "1" : "0");
}

function saveSettings() {
  const key = document.getElementById("apiKey").value.trim();
  if (!key) {
    setStatus("Enter an API key before saving.", "error");
    return;
  }
  localStorage.setItem(STORAGE_KEY_API, key);
  setStatus("Settings saved.", "success");
}

function clearSettings() {
  localStorage.removeItem(STORAGE_KEY_API);
  document.getElementById("apiKey").value = "";
  setStatus("Settings cleared.", "info");
}

function loadStoredSettings() {
  const saved = localStorage.getItem(STORAGE_KEY_API);
  if (saved) document.getElementById("apiKey").value = saved;
  if (localStorage.getItem(STORAGE_KEY_SETTINGS_OPEN) === "1") {
    document.getElementById("settings-panel").classList.add("open");
    document.getElementById("settings-arrow").textContent = "▲";
  }
}

// --- Project loading ---

async function loadProject() {
  const number = document.getElementById("projectNumber").value.trim();
  const apiKey = localStorage.getItem(STORAGE_KEY_API);

  if (!number) {
    setStatus("Enter a project number.", "error");
    return;
  }
  if (!apiKey) {
    setStatus("API key not set. Open Settings and save your key.", "error");
    return;
  }

  setStatus('<span class="spinner"></span>Loading project...', "info");
  setLoadBtn(true);

  try {
    const res = await fetch(
      `https://api.totalsynergy.com/api/v2/Organisation/${ORG_SLUG}/Projects?criteria.projectNumber=${encodeURIComponent(number)}`,
      { headers: { "access-token": apiKey, Accept: "application/json" } }
    );

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server returned ${res.status}`);
    }

    const data = await res.json();
    const project = data.items ? data.items[0] : (Array.isArray(data) ? data[0] : data);

    if (!project) {
      throw new Error("Project not found. Note: only active projects can be loaded. If this project is completed or on hold, it cannot be retrieved via the API.");
    }

    populateFields(project);
    document.getElementById("fields-section").style.display = "block";
    setStatus(`Project loaded: ${getNestedValue(project, "name") || number}`, "success");
  } catch (err) {
    setStatus(err.message, "error");
  } finally {
    setLoadBtn(false);
  }
}

function populateFields(project) {
  for (const field of FIELD_MAP) {
    const value = getNestedValue(project, field.apiPath) || "";
    document.getElementById(field.inputId).value = String(value);
  }
}

// Reads a dot-notation path from an object, e.g. "client.name" → obj.client.name
function getNestedValue(obj, path) {
  return path.split(".").reduce((acc, key) => (acc != null ? acc[key] : undefined), obj);
}

// --- Apply to document ---

async function applyToDocument() {
  setStatus('<span class="spinner"></span>Applying to document...', "info");
  document.getElementById("applyBtn").disabled = true;

  try {
    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag,items/text");
      await context.sync();

      const updated = [];
      const notFound = [];

      for (const field of FIELD_MAP) {
        const value = document.getElementById(field.inputId).value;
        const matches = controls.items.filter((cc) => cc.tag === field.tag);

        if (matches.length === 0) {
          if (value) notFound.push(field.tag);
          continue;
        }
        for (const cc of matches) {
          cc.insertText(value, "Replace");
          updated.push(field.tag);
        }
      }

      await context.sync();

      let msg = `Applied ${updated.length} field(s).`;
      if (notFound.length > 0) {
        msg += ` No controls found for: ${notFound.join(", ")}.`;
      }
      setStatus(msg, updated.length > 0 ? "success" : "info");
    });
  } catch (err) {
    setStatus("Error updating document: " + err.message, "error");
  } finally {
    document.getElementById("applyBtn").disabled = false;
  }
}

// --- Helpers ---

function setStatus(html, type) {
  const el = document.getElementById("status");
  el.innerHTML = html;
  el.className = type || "";
}

function setLoadBtn(loading) {
  const btn = document.getElementById("loadBtn");
  btn.disabled = loading;
  btn.textContent = loading ? "..." : "Load";
}
