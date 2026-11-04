/* Mail Notes – Outlook add-in
 * הערות נשמרות כ-Custom Properties על המייל עצמו (בתיבת הדואר שלך).
 * בנוסף נשמר אינדקס קטן ב-roamingSettings כדי להציג את "כל ההערות".
 */
(function () {
  "use strict";

  const PROP_KEY = "mailNotes";        // custom property on the item
  const INDEX_KEY = "mailNotesIndex";  // roaming setting
  const INDEX_MAX = 150;               // roamingSettings is limited to ~32KB
  // Each note color = an Outlook category shown in the message list.
  // אפשר לשנות כאן את השמות (name). preset = צבע הקטגוריה ב-Outlook.
  const LABELS = [
    { hex: "#e0b94a", name: "📝 Note",           preset: "Preset3" }, // צהוב
    { hex: "#d9534f", name: "🔴 Urgent",         preset: "Preset0" }, // אדום
    { hex: "#4a9be0", name: "🔵 To Review",      preset: "Preset7" }, // כחול
    { hex: "#a066d3", name: "🟣 Awaiting Reply", preset: "Preset8" }, // סגול
    { hex: "#5cb85c", name: "🟢 Done",           preset: "Preset4" }, // ירוק
  ];
  const LEGACY = { "#e07a4a": "#d9534f" }; // colors from version 1.0
  // Earlier Hebrew category names – removed from a message when it is opened
  const OLD_NAMES = ["📝 הערה", "🔴 דחוף", "🔵 לבדיקה", "🟣 ממתין לתשובה", "🟢 טופל"];
  const COLORS = LABELS.map((l) => l.hex);
  const labelFor = (hex) =>
    LABELS.find((l) => l.hex === (LEGACY[hex] || hex)) || LABELS[0];

  let props = null;       // CustomProperties of the current item
  let notes = [];         // [{id, text, color, created, updated}]
  let editingId = null;
  let color = COLORS[0];

  const $ = (id) => document.getElementById(id);

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;
    buildColors();
    wireUi();
    loadItem();
    // When the pane is pinned, follow the selected message
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadItem);
    }
  });

  // ---------- current item ----------
  function loadItem() {
    const item = Office.context.mailbox.item;
    resetComposer();
    if (!item) {
      $("subject").textContent = "לא נבחר מייל";
      notes = []; props = null; renderNotes();
      return;
    }
    $("subject").textContent = item.subject || "(ללא נושא)";
    item.loadCustomPropertiesAsync((res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        toast("שגיאה בטעינת ההערות");
        return;
      }
      props = res.value;
      try { notes = JSON.parse(props.get(PROP_KEY) || "[]"); }
      catch (e) { notes = []; }
      renderNotes();
      syncCategory();
    });
  }

  // ---------- category marker in the message list ----------
  function categoriesSupported() {
    const item = Office.context.mailbox.item;
    return item && item.categories &&
      Office.context.requirements.isSetSupported("Mailbox", "1.8");
  }

  // Make sure the needed categories exist in the user's category list
  function ensureMasterCategories(names, done) {
    const mc = Office.context.mailbox.masterCategories;
    if (!mc || !names.length) return done();
    mc.getAsync((r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return done();
      const existing = new Set((r.value || []).map((c) => c.displayName));
      const missing = LABELS
        .filter((l) => names.includes(l.name) && !existing.has(l.name))
        .map((l) => ({ displayName: l.name, color: Office.MailboxEnums.CategoryColor[l.preset] }));
      if (!missing.length) return done();
      mc.addAsync(missing, () => done());
    });
  }

  // The message gets one category per note color in use; unused ones are removed
  function syncCategory() {
    if (!categoriesSupported()) return;
    const item = Office.context.mailbox.item;
    const wanted = [...new Set(notes.map((n) => labelFor(n.color).name))];
    const ours = new Set(LABELS.map((l) => l.name).concat(OLD_NAMES));
    item.categories.getAsync((r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      const current = (r.value || []).map((c) => c.displayName).filter((n) => ours.has(n));
      const toAdd = wanted.filter((n) => !current.includes(n));
      const toRemove = current.filter((n) => !wanted.includes(n));
      if (toRemove.length) item.categories.removeAsync(toRemove, () => {});
      if (toAdd.length) {
        ensureMasterCategories(toAdd, () => item.categories.addAsync(toAdd, () => {}));
      }
    });
  }

  function persist(msg) {
    if (!props) return;
    if (notes.length) props.set(PROP_KEY, JSON.stringify(notes));
    else props.remove(PROP_KEY);
    props.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        updateIndex();
        syncCategory();
        toast(msg);
      } else {
        toast("השמירה נכשלה: " + (res.error && res.error.message));
      }
    });
  }

  // ---------- notes CRUD ----------
  function saveNote() {
    const text = $("input").value.trim();
    if (!text) return;
    const now = Date.now();
    if (editingId) {
      const n = notes.find((x) => x.id === editingId);
      if (n) { n.text = text; n.color = color; n.updated = now; }
      persist("ההערה עודכנה");
    } else {
      notes.push({ id: now.toString(36), text, color, created: now, updated: now });
      persist("ההערה נשמרה");
    }
    resetComposer();
    renderNotes();
  }

  function editNote(id) {
    const n = notes.find((x) => x.id === id);
    if (!n) return;
    editingId = id;
    $("input").value = n.text;
    setColor(n.color || COLORS[0]);
    $("save").textContent = "עדכן";
    $("save").disabled = false;
    $("cancel").hidden = false;
    $("input").focus();
  }

  function deleteNote(id) {
    if (!window.confirm("למחוק את ההערה?")) return;
    notes = notes.filter((x) => x.id !== id);
    if (editingId === id) resetComposer();
    persist("ההערה נמחקה");
    renderNotes();
  }

  function resetComposer() {
    editingId = null;
    $("input").value = "";
    $("save").textContent = "הוסף";
    $("save").disabled = true;
    $("cancel").hidden = true;
  }

  // ---------- index of all notes (roaming settings) ----------
  function getIndex() {
    try { return Office.context.roamingSettings.get(INDEX_KEY) || {}; }
    catch (e) { return {}; }
  }

  function updateIndex() {
    const item = Office.context.mailbox.item;
    if (!item || !item.itemId) return;
    const idx = getIndex();
    if (notes.length) {
      const last = notes[notes.length - 1];
      idx[item.itemId] = {
        s: (item.subject || "(ללא נושא)").slice(0, 80),
        p: last.text.replace(/\s+/g, " ").slice(0, 80),
        c: notes.length,
        u: Date.now(),
      };
    } else {
      delete idx[item.itemId];
    }
    // keep only the most recent entries to stay under the size limit
    const keys = Object.keys(idx).sort((a, b) => idx[b].u - idx[a].u);
    keys.slice(INDEX_MAX).forEach((k) => delete idx[k]);
    Office.context.roamingSettings.set(INDEX_KEY, idx);
    Office.context.roamingSettings.saveAsync();
  }

  function renderIndex() {
    const q = $("search").value.trim().toLowerCase();
    const idx = getIndex();
    const list = $("index");
    list.innerHTML = "";
    const entries = Object.entries(idx)
      .filter(([, v]) => !q || (v.s + " " + v.p).toLowerCase().includes(q))
      .sort((a, b) => b[1].u - a[1].u);
    if (!entries.length) {
      list.innerHTML = '<li class="empty">' + (q ? "אין תוצאות" : "עדיין אין מיילים עם הערות") + "</li>";
      return;
    }
    entries.forEach(([itemId, v]) => {
      const li = document.createElement("li");
      const b = document.createElement("button");
      b.innerHTML =
        '<div class="idx-subj"></div><div class="idx-prev"></div>';
      b.querySelector(".idx-subj").textContent = v.s + (v.c > 1 ? "  (" + v.c + ")" : "");
      b.querySelector(".idx-prev").textContent = v.p + " · " + fmtDate(v.u);
      b.title = "פתח את המייל";
      b.onclick = () => {
        try { Office.context.mailbox.displayMessageForm(itemId); }
        catch (e) { toast("לא ניתן לפתוח – ייתכן שהמייל הועבר או נמחק"); }
      };
      li.appendChild(b);
      list.appendChild(li);
    });
  }

  // ---------- rendering ----------
  function renderNotes() {
    $("count").textContent = notes.length;
    const ul = $("notes");
    ul.innerHTML = "";
    if (!notes.length) {
      ul.innerHTML = '<li class="empty">אין הערות למייל הזה</li>';
      return;
    }
    notes.slice().reverse().forEach((n) => {
      const li = document.createElement("li");
      li.className = "note";
      const lbl = labelFor(n.color);
      li.style.setProperty("--c", lbl.hex);
      const edited = n.updated !== n.created ? " (נערך)" : "";
      li.innerHTML =
        '<div class="note-text"></div>' +
        '<div class="note-meta"><span></span>' +
        '<span class="note-actions"><button class="edit">עריכה</button>' +
        '<button class="del">מחיקה</button></span></div>';
      li.querySelector(".note-text").textContent = n.text;
      li.querySelector(".note-meta span").textContent =
        lbl.name + " · " + fmtDate(n.updated) + edited;
      li.querySelector(".edit").onclick = () => editNote(n.id);
      li.querySelector(".del").onclick = () => deleteNote(n.id);
      ul.appendChild(li);
    });
  }

  function buildColors() {
    const wrap = $("colors");
    COLORS.forEach((c) => {
      const b = document.createElement("button");
      b.className = "swatch";
      b.style.background = c;
      b.dataset.color = c;
      b.title = labelFor(c).name;
      b.onclick = () => setColor(c);
      wrap.appendChild(b);
    });
    const name = document.createElement("span");
    name.id = "labelName";
    name.className = "label-name";
    wrap.appendChild(name);
    setColor(COLORS[0]);
  }

  function setColor(c) {
    color = labelFor(c).hex;
    document.querySelectorAll(".swatch").forEach((s) =>
      s.setAttribute("aria-pressed", s.dataset.color === color ? "true" : "false"));
    $("labelName").textContent = labelFor(color).name;
  }

  function wireUi() {
    $("input").addEventListener("input", () => {
      $("save").disabled = !$("input").value.trim();
    });
    $("input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveNote(); }
      if (e.key === "Escape") resetComposer();
    });
    $("save").onclick = saveNote;
    $("cancel").onclick = resetComposer;
    $("tab-item").onclick = () => switchTab("item");
    $("tab-all").onclick = () => switchTab("all");
    $("search").addEventListener("input", renderIndex);
  }

  function switchTab(which) {
    $("tab-item").setAttribute("aria-selected", which === "item");
    $("tab-all").setAttribute("aria-selected", which === "all");
    $("view-item").hidden = which !== "item";
    $("view-all").hidden = which !== "all";
    if (which === "all") renderIndex();
  }

  function fmtDate(ts) {
    return new Date(ts).toLocaleString("he-IL", {
      day: "2-digit", month: "2-digit", year: "2-digit", hour: "2-digit", minute: "2-digit",
    });
  }

  let toastTimer;
  function toast(msg) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 1800);
  }
})();
