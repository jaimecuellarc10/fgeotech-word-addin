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
  { inputId: "f_client_email",        tag: "synergy_client_email",        apiPath: null },
  { inputId: "f_report_writer",       tag: "synergy_report_writer",       apiPath: null },
  { inputId: "f_report_reviewer",     tag: "synergy_report_reviewer",     apiPath: null },
  { inputId: "f_investigation_type",  tag: "synergy_investigation_type",  apiPath: null },
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

    // Fetch client email from the Contacts endpoint using the project's primaryContactId
    document.getElementById("f_client_email").value = "";
    if (project.primaryContactId) {
      const contactRes = await fetch(
        `https://api.totalsynergy.com/api/v2/Organisation/${ORG_SLUG}/Contacts/${project.primaryContactId}`,
        { headers: { "access-token": apiKey, Accept: "application/json" } }
      ).catch(() => null);
      if (contactRes && contactRes.ok) {
        const contact = await contactRes.json();
        const email = contact.email || contact.emailAddress
          || (Array.isArray(contact.emails) ? contact.emails[0] : "")
          || "";
        document.getElementById("f_client_email").value = email;
      }
    }

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
    if (!field.apiPath) continue;
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
      const ooxmlResult = context.document.body.getOoxml();
      await context.sync();

      // contentControls API returns 0 for table-cell-level SDTs in this doc structure,
      // so we manipulate the OOXML directly instead.
      let xml = ooxmlResult.value;

      // Strip placeholder-text flags so filled values display instead of placeholders
      xml = xml.replace(/<w:showingPlcHdr\s*\/>/gi, "");
      xml = xml.replace(/<w:rStyle\s+w:val="PlaceholderText"\s*\/>/gi, "");

      const updated = [];
      const notFound = [];

      for (const field of FIELD_MAP) {
        const value = document.getElementById(field.inputId).value;
        const result = updateSdtByTag(xml, field.tag, value);
        xml = result.xml;
        if (result.updated) {
          updated.push(field.tag);
        } else if (value) {
          notFound.push(field.tag);
        }
      }

      context.document.body.insertOoxml(xml, "Replace");
      await context.sync();

      let msg = `Applied ${updated.length} field(s).`;
      if (notFound.length > 0) {
        const docTags = [...new Set(
          [...xml.matchAll(/w:val="(synergy_[^"]+)"/gi)].map(m => m[1].toLowerCase())
        )];
        msg += ` Missing: ${notFound.join(", ")}. Doc has: ${docTags.join(", ")}.`;
      }
      setStatus(msg, updated.length > 0 ? "success" : "info");
    });
  } catch (err) {
    setStatus("Error updating document: " + err.message, "error");
  } finally {
    document.getElementById("applyBtn").disabled = false;
  }
}

// Finds the SDT with the given tag name in raw OOXML and replaces its text content.
function updateSdtByTag(xml, tagName, value) {
  const safeValue = value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  // Case-insensitive: tag names in Word Properties dialog may differ in capitalisation
  const escapedTag = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const tagMatch = new RegExp(`w:val="${escapedTag}"`, "i").exec(xml);
  if (!tagMatch) return { xml, updated: false };
  const tagIdx = tagMatch.index;

  const sdtContentOpen = "<w:sdtContent>";
  const sdtPos = xml.indexOf(sdtContentOpen, tagIdx);
  if (sdtPos === -1) return { xml, updated: false };

  const contentStart = sdtPos + sdtContentOpen.length;
  const contentEnd = xml.indexOf("</w:sdtContent>", contentStart);
  if (contentEnd === -1) return { xml, updated: false };

  let content = xml.slice(contentStart, contentEnd);

  let firstReplaced = false;
  content = content.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/gi, () => {
    if (!firstReplaced) {
      firstReplaced = true;
      return `<w:t xml:space="preserve">${safeValue}</w:t>`;
    }
    return "<w:t/>";
  });

  if (!firstReplaced) return { xml, updated: false };

  return {
    xml: xml.slice(0, contentStart) + content + xml.slice(contentEnd),
    updated: true,
  };
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
