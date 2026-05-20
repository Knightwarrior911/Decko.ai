/* Decko Desktop — consumer polish (SP5).
   Frontend talks to the Python `Api` exposed via pywebview. All backend
   method names are unchanged from SP1/SP2 except for two narrow additions:
     - api.set_window_title(title)
     - api.pick_pptx_path()
*/
const $ = (id) => document.getElementById(id);
let api;
let currentSession = null;
let deckDirty = false;
let currentDeckName = "";

/* ---------- curated model dropdown ----------------------------------- */
const MODELS_BY_PROVIDER = {
  anthropic: [
    { id: "claude-opus-4-7",   label: "Claude Opus 4.7 (recommended)",  recommended: true },
    { id: "claude-sonnet-4-6", label: "Claude Sonnet 4.6" },
    { id: "claude-haiku-4-5",  label: "Claude Haiku 4.5" },
  ],
  openai: [
    { id: "gpt-4o",      label: "GPT-4o (recommended)", recommended: true },
    { id: "gpt-4o-mini", label: "GPT-4o mini" },
  ],
  generic: [],
};

function populateModelDropdown(selectEl, provider, freeEl, baseEl) {
  selectEl.innerHTML = "";
  const list = MODELS_BY_PROVIDER[provider] || [];
  if (list.length === 0) {
    selectEl.hidden = true;
    if (freeEl) freeEl.parentElement.parentElement.open = true;
  } else {
    selectEl.hidden = false;
    list.forEach((m) => selectEl.add(new Option(m.label, m.id)));
    const rec = list.find((m) => m.recommended);
    if (rec) selectEl.value = rec.id;
  }
}

/* ---------- friendly error translation -------------------------------- */
function friendlyError(raw) {
  if (!raw) return "Something went wrong.";
  const s = String(raw);
  if (/NoOpenDeckError|No deck open/i.test(s))
    return "No deck is open in PowerPoint. Open a deck, or use Open file instead.";
  if (/NoPowerPointError|PowerPoint.*(required|not found|not installed)/i.test(s))
    return "Microsoft PowerPoint isn't installed on this PC. Install it and try again.";
  if (/EmptyDeckError|Empty deck/i.test(s))
    return "The deck has no slides. Add at least one and try again.";
  if (/401/.test(s) && /LLM|API/i.test(s))
    return "Your AI key was rejected. Open Settings and check it.";
  if (/429/.test(s) && /LLM|API/i.test(s))
    return "The AI is rate-limited right now. Wait a minute and try again.";
  if (/Spec is not valid JSON|JSON parse/i.test(s))
    return "That isn't valid JSON. Check the brackets and quotes.";
  if (/Start a session first/i.test(s))
    return "Connect to a deck first, then try that again.";
  return "Something went wrong: " + s.slice(0, 200);
}

/* ---------- toasts ---------------------------------------------------- */
function showToast(text, kind) {
  const wrap = $("toast");
  const t = document.createElement("div");
  t.className = "toastItem " + (kind || "");
  t.textContent = text;
  wrap.appendChild(t);
  setTimeout(() => { t.style.opacity = "0"; }, 2600);
  setTimeout(() => { t.remove(); }, 3000);
}

/* ---------- clipboard + bubble helpers (preserved from SP1) ----------- */
function copyText(t) {
  try { navigator.clipboard.writeText(t); }
  catch (e) {
    const ta = document.createElement("textarea");
    ta.value = t; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); } catch (_) {}
    ta.remove();
  }
}

function bubble(kind, text, withCopy) {
  const d = document.createElement("div");
  d.className = "bubble " + kind;
  d.textContent = text;
  if (withCopy) {
    const b = document.createElement("button");
    b.className = "copybtn"; b.textContent = "Copy";
    b.onclick = (ev) => { ev.stopPropagation(); copyText(text);
      b.textContent = "Copied";
      setTimeout(() => (b.textContent = "Copy"), 1200); };
    d.appendChild(b);
  }
  $("thread").appendChild(d);
  $("thread").scrollTop = $("thread").scrollHeight;
  return d;
}

function clearThread() { $("thread").innerHTML = ""; }

function renderTurn(t) {
  bubble("user", t.request, true);
  const w = t.warnings ? ` (${t.warnings} warnings)` : "";
  bubble("app", (t.result_summary || "") + w, true);
}

function renderSessions(sessions, activeId) {
  const ul = $("sessions"); ul.innerHTML = "";
  (sessions || []).forEach((s) => {
    const li = document.createElement("li");
    if (s.id === activeId) li.className = "active";
    const cnt = document.createElement("span");
    cnt.className = "cnt"; cnt.textContent = s.turn_count;
    li.textContent = s.title;
    li.appendChild(cnt);
    li.onclick = async () => {
      const r = await api.load_session(s.id);
      clearThread();
      $("threadTitle").textContent =
        s.id === currentSession ? s.title
          : s.title + "  (past chat — start a new chat to edit again)";
      (r.turns || []).forEach(renderTurn);
      renderSessions(sessions, s.id);
    };
    ul.appendChild(li);
  });
}

async function refreshSessions(activeId) {
  const r = await api.list_sessions();
  renderSessions(r.sessions, activeId);
}

/* ---------- hero / composer visibility -------------------------------- */
function showHero() {
  $("hero").classList.remove("hidden");
  $("composerWrap").classList.add("hidden");
  $("thread").innerHTML = "";
}
function hideHero() {
  $("hero").classList.add("hidden");
  $("composerWrap").classList.remove("hidden");
}

/* ---------- dirty / title tracking ------------------------------------ */
function setDirty(v) {
  deckDirty = !!v;
  const btn = $("savePptBtn");
  if (v) btn.classList.add("dirty"); else btn.classList.remove("dirty");
  pushWindowTitle();
}
function pushWindowTitle() {
  const dot = deckDirty ? " ●" : "";
  const name = currentDeckName ? ` — ${currentDeckName}` : "";
  try { api && api.set_window_title && api.set_window_title("Decko" + name + dot); }
  catch (_) {}
}

/* ---------- theme ----------------------------------------------------- */
function applyTheme(t) {
  document.documentElement.setAttribute("data-theme", t);
  $("themeDark").classList.toggle("active", t === "dark");
  $("themeLight").classList.toggle("active", t === "light");
}
function loadTheme() {
  applyTheme(localStorage.getItem("decko_theme") || "dark");
}

/* ---------- boot ------------------------------------------------------ */
let bootSnapshot = null;

window.addEventListener("pywebviewready", async () => {
  api = window.pywebview.api;
  loadTheme();
  bootSnapshot = await api.boot();
  // Populate Settings dialog defaults from persisted settings.
  $("setProvider").value = bootSnapshot.settings.provider;
  $("setBaseUrl").value = bootSnapshot.settings.base_url || "";
  $("setModelFree").value = bootSnapshot.settings.model || "";
  populateModelDropdown($("modelPick"), bootSnapshot.settings.provider,
    $("setModelFree"), $("setBaseUrl"));
  if (MODELS_BY_PROVIDER[bootSnapshot.settings.provider] &&
      MODELS_BY_PROVIDER[bootSnapshot.settings.provider]
        .some((m) => m.id === bootSnapshot.settings.model)) {
    $("modelPick").value = bootSnapshot.settings.model;
  }
  refreshKeyStatus(bootSnapshot.has_key);
  renderSessions(bootSnapshot.sessions, null);
  showHero();

  // First-run: no key AND wizard never finished -> open wizard.
  if (!bootSnapshot.has_key &&
      localStorage.getItem("decko_wizard_done") !== "1") {
    openWizard();
  }
});

/* ---------- header / gear -------------------------------------------- */
$("gearBtn").onclick = () => openSettings();

/* ---------- wizard ---------------------------------------------------- */
function setWizStep(n) {
  ["wizardStep1", "wizardStep2", "wizardStep3"].forEach((id, i) => {
    $(id).classList.toggle("hidden", i + 1 !== n);
  });
  ["wizDot1", "wizDot2", "wizDot3"].forEach((id, i) => {
    $(id).classList.toggle("active", i + 1 === n);
  });
}
function openWizard() {
  $("wizard").classList.remove("hidden");
  setWizStep(1);
  // Seed wizard provider/model from persisted settings.
  if (bootSnapshot) {
    $("wizProvider").value = bootSnapshot.settings.provider;
    populateModelDropdown($("wizModel"), bootSnapshot.settings.provider,
      $("wizModelFree"), $("wizBaseUrl"));
    $("wizModelFree").value = bootSnapshot.settings.model || "";
    $("wizBaseUrl").value = bootSnapshot.settings.base_url || "";
  }
}
function closeWizard() { $("wizard").classList.add("hidden"); }

$("wizStart").onclick = () => setWizStep(2);
$("wizBack2").onclick = () => setWizStep(1);
$("wizBack3").onclick = () => setWizStep(2);
$("wizSkip").onclick = () => {
  localStorage.setItem("decko_wizard_done", "1");
  closeWizard();
};
$("wizProvider").onchange = (e) => {
  populateModelDropdown($("wizModel"), e.target.value,
    $("wizModelFree"), $("wizBaseUrl"));
};
$("wizSaveKey").onclick = async () => {
  const provider = $("wizProvider").value;
  const model = (provider === "generic")
    ? ($("wizModelFree").value || "")
    : ($("wizModel").value || "");
  const baseUrl = $("wizBaseUrl").value;
  const key = $("wizKey").value;
  if (provider !== "generic" && !model) {
    showToast("Pick a model.", "error"); return;
  }
  if (!key) { showToast("Enter your API key.", "error"); return; }
  const r = await api.save_settings(provider, model, baseUrl, key);
  if (!r.ok) { showToast(friendlyError(r.error), "error"); return; }
  $("wizKey").value = "";
  showToast("Settings saved.", "success");
  // Refresh boot snapshot for downstream calls.
  bootSnapshot = await api.boot();
  refreshKeyStatus(true);
  setWizStep(3);
};
$("wizConnect").onclick = async () => {
  localStorage.setItem("decko_wizard_done", "1");
  closeWizard();
  await heroConnect();
};
$("wizOpenFile").onclick = async () => {
  localStorage.setItem("decko_wizard_done", "1");
  closeWizard();
  await heroOpenFile();
};

/* ---------- settings dialog ------------------------------------------ */
function refreshKeyStatus(hasKey) {
  if (hasKey) {
    $("keyMask").textContent = "••••• Connected ✓";
    $("updateKeyLink").style.display = "";
    $("setKey").hidden = true;
    $("setKey").value = "";
  } else {
    $("keyMask").textContent = "No key saved";
    $("updateKeyLink").style.display = "none";
    $("setKey").hidden = false;
  }
}
function openSettings() {
  $("settingsDialog").classList.remove("hidden");
  if (bootSnapshot) {
    $("setProvider").value = bootSnapshot.settings.provider;
    populateModelDropdown($("modelPick"), bootSnapshot.settings.provider,
      $("setModelFree"), $("setBaseUrl"));
    if (MODELS_BY_PROVIDER[bootSnapshot.settings.provider] &&
        MODELS_BY_PROVIDER[bootSnapshot.settings.provider]
          .some((m) => m.id === bootSnapshot.settings.model)) {
      $("modelPick").value = bootSnapshot.settings.model;
    }
    $("setModelFree").value = bootSnapshot.settings.model || "";
    $("setBaseUrl").value = bootSnapshot.settings.base_url || "";
    refreshKeyStatus(bootSnapshot.has_key);
  }
}
function closeSettings() { $("settingsDialog").classList.add("hidden"); }

$("setCancel").onclick = closeSettings;
$("setProvider").onchange = (e) => {
  populateModelDropdown($("modelPick"), e.target.value,
    $("setModelFree"), $("setBaseUrl"));
};
$("updateKeyLink").onclick = () => {
  $("setKey").hidden = false;
  $("setKey").focus();
  $("setKey").placeholder = "Enter new key";
};
$("themeDark").onclick = () => {
  localStorage.setItem("decko_theme", "dark"); applyTheme("dark"); };
$("themeLight").onclick = () => {
  localStorage.setItem("decko_theme", "light"); applyTheme("light"); };

$("setSave").onclick = async () => {
  const provider = $("setProvider").value;
  const model = (provider === "generic")
    ? ($("setModelFree").value || "")
    : ($("modelPick").value || "");
  const baseUrl = $("setBaseUrl").value;
  const key = $("setKey").value;
  if (provider !== "generic" && !model) {
    showToast("Pick a model.", "error"); return;
  }
  const r = await api.save_settings(provider, model, baseUrl, key);
  if (!r.ok) { showToast(friendlyError(r.error), "error"); return; }
  bootSnapshot = await api.boot();
  refreshKeyStatus(bootSnapshot.has_key);
  showToast("Settings saved.", "success");
  closeSettings();
};

/* ---------- hero handlers -------------------------------------------- */
async function heroConnect() {
  if (!bootSnapshot || !bootSnapshot.has_key) {
    openWizard(); setWizStep(2); return;
  }
  $("heroConnect").disabled = true;
  const r = await api.open_session("attach", "");
  $("heroConnect").disabled = false;
  if (r.error) { showToast(friendlyError(r.error), "error");
    bubble("fail", friendlyError(r.error), true); return; }
  onSessionStarted(r);
}
async function heroOpenFile() {
  let path = "";
  try {
    if (api && api.pick_pptx_path) path = await api.pick_pptx_path();
  } catch (_) {}
  if (!path) {
    path = prompt("Full path to a .pptx file:");
    if (!path) return;
  }
  $("heroOpenFile").disabled = true;
  const r = await api.open_session("file", path);
  $("heroOpenFile").disabled = false;
  if (r.error) { showToast(friendlyError(r.error), "error");
    bubble("fail", friendlyError(r.error), true); return; }
  onSessionStarted(r);
}

function onSessionStarted(r) {
  currentSession = r.session_id;
  currentDeckName = (r.title || "").split(" — ")[0] || "Deck";
  $("deckStatus").innerHTML = "<b>" + currentDeckName + "</b>";
  $("threadTitle").textContent = r.title;
  hideHero();
  bubble("app", "Connected. Describe a change to get started.");
  showToast("Connected to " + currentDeckName, "success");
  setDirty(false);
  pushWindowTitle();
  refreshSessions(currentSession);
}

$("heroConnect").onclick = heroConnect;
$("heroOpenFile").onclick = heroOpenFile;
$("switchDeckBtn").onclick = async () => {
  await api.new_session();
  currentSession = null;
  currentDeckName = "";
  $("deckStatus").textContent = "No deck connected.";
  showHero();
  setDirty(false);
  pushWindowTitle();
  refreshSessions(null);
};

/* ---------- new chat -------------------------------------------------- */
$("newBtn").onclick = async () => {
  await api.new_session();
  currentSession = null;
  currentDeckName = "";
  $("deckStatus").textContent = "No deck connected.";
  $("threadTitle").textContent = "";
  showHero();
  setDirty(false);
  pushWindowTitle();
  refreshSessions(null);
};

/* ---------- save deck ------------------------------------------------- */
$("savePptBtn").onclick = async () => {
  const r = await api.save_powerpoint();
  if (r.ok) {
    showToast("Deck saved.", "success");
    setDirty(false);
  } else {
    showToast(friendlyError(r.error), "error");
    bubble("fail", friendlyError(r.error), true);
  }
};

/* ---------- composer + chips ----------------------------------------- */
document.querySelectorAll(".chip").forEach((el) => {
  el.onclick = () => {
    $("msg").value = el.dataset.prefill || el.textContent;
    $("msg").focus();
  };
});

async function doSend() {
  const t = $("msg").value.trim();
  if (!t) return;
  if (!currentSession) {
    showToast("Connect to a deck first.", "error");
    bubble("fail", "Connect to a deck first.", true);
    return;
  }
  bubble("user", t, true);
  $("msg").value = "";
  $("sendBtn").disabled = true;
  const working = bubble("working", "Editing your deck…");
  let r;
  try { r = await api.send(t); }
  finally { working.remove(); $("sendBtn").disabled = false; }
  if (r.error) {
    bubble("fail", friendlyError(r.error), true);
    showToast(friendlyError(r.error), "error");
    return;
  }
  const w = r.warnings ? ` (${r.warnings} warnings)` : "";
  bubble("app", (r.summary || "") + w, true);
  setDirty(true);
  refreshSessions(currentSession);
}

$("sendBtn").onclick = doSend;
$("msg").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
    e.preventDefault(); doSend();
  }
});

/* ---------- templates panel (SP2 wiring preserved) ------------------- */
let BUILTINS = [];
async function tplInit() {
  BUILTINS = (await api.list_builtin_templates()).templates;
  const pick = $("tplPick"), vt = $("varTpl");
  pick.innerHTML = ""; vt.innerHTML = "";
  BUILTINS.forEach((t) => {
    pick.add(new Option(t.name, t.name));
    vt.add(new Option(t.name, t.name));
  });
  renderSlots();
  await refreshCaptured();
}
function curTpl() { return BUILTINS.find((t) => t.name === $("tplPick").value); }
function renderSlots() {
  const t = curTpl(); const box = $("tplSlots"); box.innerHTML = "";
  if (!t) return;
  t.slots.forEach((s) => {
    const lbl = document.createElement("label");
    lbl.className = "muted"; lbl.textContent = s;
    const i = document.createElement("input");
    i.id = "slot_" + s; i.placeholder = s;
    i.title = "Content for the " + s + " slot.";
    box.appendChild(lbl); box.appendChild(i);
  });
}
function collectContent() {
  const t = curTpl(); const c = {};
  t.slots.forEach((s) => {
    const v = ($("slot_" + s) || {}).value || "";
    if (s === "bullets") c[s] = v ? v.split("\n") : ["Point one"];
    else if (s === "tiles") c[s] = [{ stat: "00", label: v || "Metric" }];
    else c[s] = v || s;
  });
  return c;
}
async function refreshCaptured() {
  const r = await api.list_captured_templates();
  const ul = $("capList"); ul.innerHTML = "";
  (r.templates || []).forEach((t) => {
    const li = document.createElement("li");
    const sp = document.createElement("span"); sp.textContent = t.name;
    const ap = document.createElement("button");
    ap.textContent = "Apply"; ap.title = "Apply this layout.";
    ap.onclick = async () => {
      const res = await api.apply_template(t.name, {}, tgt());
      if (res.ok) { showToast(res.summary || "Applied.", "success");
        setDirty(true); }
      else { showToast(friendlyError(res.error), "error");
        bubble("fail", friendlyError(res.error), true); }
      refreshSessions(currentSession);
    };
    const dl = document.createElement("button");
    dl.textContent = "Delete"; dl.title = "Delete this saved layout.";
    dl.onclick = async () => {
      await api.delete_template(t.name); refreshCaptured();
      showToast("Layout deleted.", "success");
    };
    li.appendChild(sp); li.appendChild(ap); li.appendChild(dl);
    ul.appendChild(li);
  });
}
function tgt() {
  return $("tplTarget").value === "replace"
    ? { mode: "replace", slide: parseInt($("tplSlideNo").value || "1") }
    : { mode: "append" };
}
$("tplBtn").onclick = () => {
  $("tplPanel").classList.toggle("tpl-hidden");
  if (!$("tplPanel").classList.contains("tpl-hidden") && api) tplInit();
};
$("tplClose").onclick = () => $("tplPanel").classList.add("tpl-hidden");
$("tplPick").onchange = renderSlots;
$("tplTarget").onchange = (e) =>
  ($("tplSlideNo").hidden = e.target.value !== "replace");
$("tplApply").onclick = async () => {
  const r = await api.apply_template($("tplPick").value, collectContent(), tgt());
  if (r.ok) { showToast(r.summary || "Layout applied.", "success");
    setDirty(true); }
  else { showToast(friendlyError(r.error), "error");
    bubble("fail", friendlyError(r.error), true); }
  refreshSessions(currentSession);
};
$("tplFill").onclick = async () => {
  const r = await api.fill_with_ai($("tplPick").value,
    $("tplBrief").value || "professional placeholder content");
  if (r.error) { showToast(friendlyError(r.error), "error"); return; }
  Object.entries(r.content).forEach(([k, v]) => {
    const el = $("slot_" + k);
    if (el) el.value = Array.isArray(v)
      ? v.map((x) => (typeof x === "object" ? JSON.stringify(x) : x)).join("\n")
      : v;
  });
  showToast("AI filled the slots — review and apply.", "success");
};
$("capBtn").onclick = async () => {
  const r = await api.capture_template($("capName").value);
  if (r.ok) {
    $("capName").value = "";
    refreshCaptured(); refreshSessions(currentSession);
    showToast("Layout saved.", "success");
  } else { showToast(friendlyError(r.error), "error"); }
};
$("varBtn").onclick = async () => {
  const r = await api.generate_variants({
    template: $("varTpl").value,
    n: parseInt($("varN").value || "3"),
    content: {} });
  if (r.ok) { showToast(r.summary || "Variants generated.", "success");
    setDirty(true); }
  else { showToast(friendlyError(r.error), "error"); }
  refreshSessions(currentSession);
};
$("specExtract").onclick = async () => {
  const r = await api.extract_spec();
  if (r.error) { showToast(friendlyError(r.error), "error"); return; }
  $("specBox").value = r.spec;
  showToast("Deck JSON ready below.", "success");
};
$("specBuild").onclick = async () => {
  const r = await api.build_deck_from_spec($("specBox").value,
    $("specClear").checked);
  if (r.ok) { showToast(r.summary || "Deck built.", "success");
    setDirty(true); }
  else { showToast(friendlyError(r.error), "error");
    bubble("fail", friendlyError(r.error), true); }
  refreshSessions(currentSession);
};
