const $ = (id) => document.getElementById(id);
let api;

function bubble(kind, html) {
  const d = document.createElement("div");
  d.className = "bubble " + kind;
  d.innerHTML = html;
  $("thread").appendChild(d);
  $("thread").scrollTop = $("thread").scrollHeight;
}

window.addEventListener("pywebviewready", async () => {
  api = window.pywebview.api;
  const s = await api.boot();
  s.history.forEach((h) =>
    bubble("app", `<b>${h.request}</b><br>${h.result_summary}`));
  if (!s.has_key) bubble("app", "Set your API key in the side panel.");
});

$("provider").onchange = (e) =>
  ($("baseUrl").hidden = e.target.value !== "generic");
$("mode").onchange = (e) =>
  ($("file").hidden = e.target.value !== "file");

$("saveBtn").onclick = async () => {
  const r = await api.save_settings($("provider").value, $("model").value,
    $("baseUrl").value, $("apiKey").value);
  bubble("app", r.ok ? "Settings saved." : "Error: " + r.error);
};

$("startBtn").onclick = async () => {
  const r = await api.open_session($("mode").value, $("file").value);
  bubble("app", r.ok ? "Session started." : "fail: " + r.error);
};

$("sendBtn").onclick = async () => {
  const t = $("msg").value.trim();
  if (!t) return;
  bubble("user", t);
  $("msg").value = "";
  const r = await api.send(t);
  if (r.error) { bubble("app fail", r.error); return; }
  const w = r.warnings ? ` <span class="warn">(${r.warnings} warnings)</span>`
                       : "";
  bubble("app", r.summary + w);
};
