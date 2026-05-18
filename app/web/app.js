const $ = (id) => document.getElementById(id);
let api;
let currentSession = null;   // live session id (from Start session)

function copyText(t) {
  try {
    navigator.clipboard.writeText(t);
  } catch (e) {
    const ta = document.createElement("textarea");
    ta.value = t; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); } catch (_) {}
    ta.remove();
  }
}

function bubble(kind, text, withCopy) {
  const d = document.createElement("div");
  d.className = "bubble " + kind;
  d.textContent = text;                       // textContent => no HTML injection
  if (withCopy) {
    const b = document.createElement("button");
    b.className = "copybtn"; b.textContent = "Copy";
    b.onclick = (ev) => { ev.stopPropagation(); copyText(text);
                          b.textContent = "Copied"; setTimeout(
                          () => (b.textContent = "Copy"), 1200); };
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
                                : s.title + "  (past session — Start a session to make changes)";
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

window.addEventListener("pywebviewready", async () => {
  api = window.pywebview.api;
  const s = await api.boot();
  $("provider").value = s.settings.provider;
  $("model").value = s.settings.model;
  $("baseUrl").value = s.settings.base_url || "";
  $("baseUrl").hidden = s.settings.provider !== "generic";
  $("mode").value = s.last_mode || "attach";
  $("file").value = s.last_deck_path || "";
  $("file").hidden = ($("mode").value !== "file");
  $("apiKey").placeholder = s.has_key
    ? "API key saved — leave blank to keep" : "API key";
  $("activeCfg").textContent =
    "Active: " + s.settings.provider + " / " + s.settings.model;
  renderSessions(s.sessions, null);
  clearThread();
  if (!s.has_key)
    bubble("app", "Set your API key in the side panel, then Save settings.");
});

$("provider").onchange = (e) =>
  ($("baseUrl").hidden = e.target.value !== "generic");
$("mode").onchange = (e) =>
  ($("file").hidden = e.target.value !== "file");

$("saveBtn").onclick = async () => {
  const r = await api.save_settings($("provider").value, $("model").value,
    $("baseUrl").value, $("apiKey").value);
  if (r.ok) {
    $("apiKey").value = "";
    $("apiKey").placeholder = "API key saved — leave blank to keep";
    $("activeCfg").textContent =
      "Active: " + $("provider").value + " / " + $("model").value;
    bubble("app", "Settings saved.");
  } else bubble("fail", r.error);
};

$("startBtn").onclick = async () => {
  $("startBtn").disabled = true;
  const r = await api.open_session($("mode").value, $("file").value);
  $("startBtn").disabled = false;
  if (r.error) { bubble("fail", r.error); return; }
  currentSession = r.session_id;
  clearThread();
  $("threadTitle").textContent = r.title;
  bubble("app", "Session started.");
  refreshSessions(currentSession);
};

$("newBtn").onclick = async () => {
  await api.new_session();
  currentSession = null;
  clearThread();
  $("threadTitle").textContent = "";
  bubble("app", "New session. Pick a deck and click Start session.");
  refreshSessions(null);
};

$("savePptBtn").onclick = async () => {
  const r = await api.save_powerpoint();
  bubble(r.ok ? "app" : "fail",
         r.ok ? "PowerPoint saved." : r.error);
};

async function doSend() {
  const t = $("msg").value.trim();
  if (!t) return;
  bubble("user", t, true);
  $("msg").value = "";
  $("sendBtn").disabled = true;
  const working = bubble("working", "working… (PowerPoint + LLM)");
  let r;
  try { r = await api.send(t); }
  finally { working.remove(); $("sendBtn").disabled = false; }
  if (r.error) { bubble("fail", r.error, true); return; }
  const w = r.warnings ? ` (${r.warnings} warnings)` : "";
  bubble("app", (r.summary || "") + w, true);
  refreshSessions(currentSession);
}

$("sendBtn").onclick = doSend;
$("msg").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
    e.preventDefault(); doSend();
  }
});

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
function curTpl() {
  return BUILTINS.find((t) => t.name === $("tplPick").value);
}
function renderSlots() {
  const t = curTpl(); const box = $("tplSlots"); box.innerHTML = "";
  if (!t) return;
  t.slots.forEach((s) => {
    const i = document.createElement("input");
    i.id = "slot_" + s; i.placeholder = s;
    box.appendChild(i);
  });
}
function collectContent() {
  const t = curTpl(); const c = {};
  t.slots.forEach((s) => {
    const v = ($("slot_" + s) || {}).value || "";
    if (s === "bullets") c[s] = v ? v.split("\n") : ["Point one"];
    else if (s === "tiles")
      c[s] = [{ stat: "00", label: v || "Metric" }];
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
    const ap = document.createElement("button"); ap.textContent = "Apply";
    ap.onclick = async () => {
      const res = await api.apply_template(t.name, {}, tgt());
      bubble(res.ok ? "app" : "fail", res.ok ? res.summary : res.error);
      refreshSessions(currentSession);
    };
    const dl = document.createElement("button"); dl.textContent = "Del";
    dl.onclick = async () => {
      await api.delete_template(t.name); refreshCaptured();
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
$("tplClose").onclick = () =>
  $("tplPanel").classList.add("tpl-hidden");
$("tplPick").onchange = renderSlots;
$("tplTarget").onchange = (e) =>
  ($("tplSlideNo").hidden = e.target.value !== "replace");
$("tplApply").onclick = async () => {
  const r = await api.apply_template($("tplPick").value,
    collectContent(), tgt());
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error, true);
  refreshSessions(currentSession);
};
$("tplFill").onclick = async () => {
  const r = await api.fill_with_ai($("tplPick").value,
    $("tplBrief").value || "professional placeholder content");
  if (r.error) { bubble("fail", r.error); return; }
  Object.entries(r.content).forEach(([k, v]) => {
    const el = $("slot_" + k);
    if (el) el.value = Array.isArray(v)
      ? v.map((x) => (typeof x === "object" ? JSON.stringify(x) : x)).join("\n")
      : v;
  });
  bubble("app", "AI filled the slots — review then Apply.");
};
$("capBtn").onclick = async () => {
  const r = await api.capture_template($("capName").value);
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error);
  if (r.ok) { $("capName").value = ""; refreshCaptured();
              refreshSessions(currentSession); }
};
$("varBtn").onclick = async () => {
  const r = await api.generate_variants({
    template: $("varTpl").value,
    n: parseInt($("varN").value || "3"),
    content: {} });
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error, true);
  refreshSessions(currentSession);
};
$("specExtract").onclick = async () => {
  const r = await api.extract_spec();
  if (r.error) { bubble("fail", r.error); return; }
  $("specBox").value = r.spec;
};
$("specBuild").onclick = async () => {
  const r = await api.build_deck_from_spec($("specBox").value,
    $("specClear").checked);
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error, true);
  refreshSessions(currentSession);
};
