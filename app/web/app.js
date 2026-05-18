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
