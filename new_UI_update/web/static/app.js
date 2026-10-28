const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
let pin = sessionStorage.getItem("pin") || "";
let lastEvent = 0;
let running = false;
const portrait = Receptionist.mount($("#portrait"));

const STATE_TEXT = {
  idle:      ["Off duty", "Start reception mode to greet visitors by voice."],
  listening: ["Listening", "Waiting for a visitor to speak."],
  thinking:  ["Thinking", "Finding the best answer."],
  speaking:  ["Speaking", "Replying to the visitor."],
};

async function api(path, { method = "GET", body } = {}) {
  const res = await fetch("/api" + path, {
    method,
    headers: { "Content-Type": "application/json", "X-Pin": pin },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    let msg = res.statusText;
    try { msg = (await res.json()).detail || msg; } catch {}
    throw new Error(msg);
  }
  return res.json();
}

function toast(msg) {
  const t = $("#toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast.t);
  toast.t = setTimeout(() => t.classList.remove("show"), 2600);
}

/* ---------- tabs ---------- */
$$("nav button").forEach(b => b.onclick = () => {
  $$("nav button").forEach(x => x.classList.toggle("active", x === b));
  $$(".tab").forEach(t => t.classList.toggle("active", t.id === "tab-" + b.dataset.tab));
  if (b.dataset.tab === "settings") loadSettings();
});

/* ---------- status + feed polling ---------- */
async function refreshStatus() {
  try {
    const s = await api("/status");
    running = s.running;
    $("#brandOffice").textContent = s.office;
    $("#brandAsst").textContent = s.assistant + " · virtual receptionist";
    $("#portrait").dataset.state = s.state;
    portrait.setState(s.state);
    portrait.setName(s.assistant);
    const [label, hint] = STATE_TEXT[s.state] || STATE_TEXT.idle;
    $("#stateLabel").textContent = label;
    $("#stateHint").textContent = hint;
    const btn = $("#toggleBtn");
    btn.textContent = running ? "Stop reception" : "Start reception";
    btn.classList.toggle("primary", !running);

    const o = $("#pillOllama");
    o.className = "pill " + (s.ollama.online ? "ok" : "bad");
    o.title = s.ollama.message;
    $("span", o).textContent = s.ollama.online ? `AI ready (${s.ollama.model})` : s.ollama.message;
    const v = $("#pillVoice");
    v.className = "pill " + (s.tts ? "ok" : "bad");
    $("span", v).textContent = s.tts ? "Voice ready" : "Voice missing: install pywin32";
    $("#qaCount").textContent = `(${s.counts.qa})`;
  } catch { /* server restarting */ }
}

function addMessage(e) {
  const feed = $("#feed");
  $(".empty", feed)?.remove();
  const div = document.createElement("div");
  div.className = "msg " + e.kind;
  if (e.source === "greeting") portrait.ding();
  const text = document.createElement("span");
  text.textContent = e.text;
  div.append(text);
  if (e.kind !== "system") {
    const meta = document.createElement("small");
    const who = e.kind === "visitor" ? (e.typed ? "Visitor (typed)" : "Visitor") : "Assistant";
    meta.append(`${who} · ${e.time} `);
    if (e.source && ["trained", "ollama", "fallback"].includes(e.source)) {
      const tag = document.createElement("span");
      tag.className = "tag " + e.source;
      tag.textContent = e.source === "trained" ? "trained answer" : e.source === "ollama" ? "local AI" : "fallback";
      meta.append(tag);
    }
    div.append(meta);
  }
  const atBottom = feed.scrollHeight - feed.scrollTop - feed.clientHeight < 80;
  feed.append(div);
  if (atBottom) feed.scrollTop = feed.scrollHeight;
}

async function pollEvents() {
  try {
    const r = await api("/events?after=" + lastEvent);
    r.events.forEach(addMessage);
    lastEvent = r.last;
  } catch {}
}

$("#toggleBtn").onclick = async () => {
  await api(running ? "/reception/stop" : "/reception/start", { method: "POST" });
  refreshStatus();
};

async function ask() {
  const input = $("#askInput");
  const text = input.value.trim();
  if (!text) return;
  input.value = "";
  $("#askBtn").disabled = true;
  try {
    await api("/ask", { method: "POST", body: { text, speak: $("#askSpeak").checked } });
    pollEvents();
  } catch (e) { toast(e.message); }
  $("#askBtn").disabled = false;
  input.focus();
}
$("#askBtn").onclick = ask;
$("#askInput").onkeydown = e => { if (e.key === "Enter") ask(); };

/* ---------- training ---------- */
function unlocked(on) {
  $("#lock").hidden = on;
  $("#trainer").hidden = !on;
  if (on) loadTraining();
}

async function tryUnlock(value) {
  try {
    await api("/auth", { method: "POST", body: { pin: value } });
    pin = value;
    sessionStorage.setItem("pin", pin);
    $("#pinErr").textContent = "";
    unlocked(true);
  } catch (e) {
    pin = "";
    sessionStorage.removeItem("pin");
    $("#pinErr").textContent = e.message;
    unlocked(false);
  }
}
$("#pinBtn").onclick = () => tryUnlock($("#pinInput").value);
$("#pinInput").onkeydown = e => { if (e.key === "Enter") tryUnlock(e.target.value); };

async function loadTraining() {
  const [greetings, qa, convo] = await Promise.all([api("/greetings"), api("/qa"), api("/conversations?limit=40")]);

  const gl = $("#greetList");
  gl.replaceChildren();
  greetings.forEach((g, i) => {
    const li = document.createElement("li");
    if (i === 0) li.className = "current";
    const t = document.createElement("span");
    t.textContent = (i === 0 ? "In use: " : "") + g.text;
    const del = document.createElement("button");
    del.textContent = "Delete";
    del.onclick = async () => { await api("/greetings/" + g.id, { method: "DELETE" }); loadTraining(); };
    li.append(t, del);
    gl.append(li);
  });

  window._qa = qa;
  renderQA();

  const ai = $("#aiLog");
  ai.replaceChildren();
  const fromAI = convo.filter(c => c.source !== "trained");
  if (!fromAI.length) {
    const li = document.createElement("li");
    li.textContent = "Nothing yet. Questions the AI answers will be listed here.";
    ai.append(li);
  }
  fromAI.forEach(c => {
    const li = document.createElement("li");
    const t = document.createElement("span");
    t.textContent = `“${c.question}” → ${c.answer}`;
    const b = document.createElement("button");
    b.textContent = "Train this";
    b.onclick = () => { startEdit(null, c.question, c.answer); scrollTo({ top: 0, behavior: "smooth" }); };
    li.append(t, b);
    ai.append(li);
  });
}

function renderQA() {
  const q = $("#qaSearch").value.toLowerCase();
  const body = $("#qaBody");
  body.replaceChildren();
  (window._qa || []).filter(r => !q || (r.question + r.answer).toLowerCase().includes(q)).forEach(r => {
    const tr = document.createElement("tr");
    [r.question, r.answer].forEach(v => { const td = document.createElement("td"); td.textContent = v; tr.append(td); });
    const hits = document.createElement("td"); hits.className = "num"; hits.textContent = r.hits; tr.append(hits);
    const act = document.createElement("td"); act.className = "act";
    const edit = document.createElement("button"); edit.textContent = "Edit"; edit.onclick = () => startEdit(r.id, r.question, r.answer);
    const del = document.createElement("button"); del.textContent = "Delete";
    del.onclick = async () => { if (confirm("Delete this answer?")) { await api("/qa/" + r.id, { method: "DELETE" }); loadTraining(); } };
    act.append(edit, " ", del);
    tr.append(act);
    body.append(tr);
  });
  if (!body.children.length) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.colSpan = 4; td.className = "muted"; td.textContent = "No answers yet. Add your first one above.";
    tr.append(td); body.append(tr);
  }
}
$("#qaSearch").oninput = renderQA;

function startEdit(id, q, a) {
  $("#qaId").value = id ?? "";
  $("#qaQ").value = q; $("#qaA").value = a;
  $("#qaTitle").textContent = id ? "Edit answer" : "Add a question and answer";
  $("#qaCancel").hidden = !id;
  $("#qaSave").textContent = id ? "Update answer" : "Save answer";
}
$("#qaCancel").onclick = () => startEdit(null, "", "");

$("#qaSave").onclick = async () => {
  const id = $("#qaId").value;
  const body = { question: $("#qaQ").value, answer: $("#qaA").value };
  try {
    await api(id ? "/qa/" + id : "/qa", { method: id ? "PUT" : "POST", body });
    toast(id ? "Answer updated" : "Answer saved");
    startEdit(null, "", "");
    loadTraining();
  } catch (e) { toast(e.message); }
};

$("#greetSave").onclick = async () => {
  try {
    await api("/greetings", { method: "POST", body: { text: $("#greetText").value } });
    $("#greetText").value = "";
    toast("Greeting saved");
    loadTraining();
  } catch (e) { toast(e.message); }
};

$$(".mic").forEach(btn => btn.onclick = async () => {
  const target = $("#" + btn.dataset.target);
  const label = btn.textContent;
  btn.classList.add("live"); btn.textContent = "Listening…"; btn.disabled = true;
  try {
    const r = await api("/listen", { method: "POST" });
    if (r.text) target.value = r.text; else toast("Didn't catch that. Try again.");
  } catch (e) { toast(e.message); }
  btn.classList.remove("live"); btn.textContent = label; btn.disabled = false;
});

/* ---------- settings ---------- */
async function loadSettings() {
  const s = await api("/settings");
  $$("#settingsForm [name]").forEach(el => { if (el.name !== "trainer_pin") el.value = s[el.name] ?? ""; });
}
$("#settingsForm").onsubmit = async e => {
  e.preventDefault();
  const body = {};
  $$("#settingsForm [name]").forEach(el => body[el.name] = el.value);
  try {
    await api("/settings", { method: "PUT", body });
    if (body.trainer_pin) { pin = body.trainer_pin; sessionStorage.setItem("pin", pin); }
    $("#settingsForm [name=trainer_pin]").value = "";
    toast("Settings saved");
    refreshStatus();
  } catch (err) { toast(err.message.includes("PIN") ? "Unlock the Training tab with your PIN first" : err.message); }
};

/* ---------- boot ---------- */
if (pin) tryUnlock(pin);
refreshStatus(); pollEvents();
setInterval(refreshStatus, 1500);
setInterval(pollEvents, 800);
