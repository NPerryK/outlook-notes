/* My Notes – Outlook add-in
 * Notes are stored as custom properties on the email itself (in your mailbox).
 * A compact index in roamingSettings powers "All notes", the summary and Excel copy.
 */
(function () {
  "use strict";

  // ---------- configuration ----------
  const PROP_KEY = "mailNotes";          // custom property on the item
  const INDEX_KEY = "mailNotesIndex";    // roaming setting: index of all notes
  const PREFS_KEY = "mailNotesSummary";  // roaming setting: summary choices
  const BANNER_KEY = "myNotesDue";       // info bar on the email
  const INDEX_MAX_BYTES = 26000;         // roamingSettings is limited to 32KB in total
  const NOTE_TEXT_MAX = 400;             // note text kept in the index
  const BODY_MAX = 30000;                // new-message body limit is ~32KB
  const DAY = 864e5;

  // Each note color = an Outlook category shown in the message list.
  // name = category name, preset = Outlook category color.
  const LABELS = [
    { hex: "#e0b94a", name: "📝 Note",           preset: "Preset3" }, // yellow
    { hex: "#d9534f", name: "🔴 Urgent",         preset: "Preset0" }, // red
    { hex: "#4a9be0", name: "🔵 To Review",      preset: "Preset7" }, // blue
    { hex: "#a066d3", name: "🟣 Awaiting Reply", preset: "Preset8" }, // purple
    { hex: "#5cb85c", name: "🟢 Done",           preset: "Preset4" }, // green
  ];
  const DONE_HEX = "#5cb85c";
  // Extra category while an email has an open note with a due date
  const FOLLOW = { name: "⏰ Follow-up", preset: "Preset1" };        // orange
  const ALL_CATS = LABELS.concat([FOLLOW]);
  const LEGACY = { "#e07a4a": "#d9534f" };   // colors from version 1.0
  const OLD_NAMES = ["📝 הערה", "🔴 דחוף", "🔵 לבדיקה", "🟣 ממתין לתשובה", "🟢 טופל"];
  const COLORS = LABELS.map((l) => l.hex);
  const labelFor = (hex) => LABELS.find((l) => l.hex === (LEGACY[hex] || hex)) || LABELS[0];
  const isDone = (n) => labelFor(n.color).hex === DONE_HEX;

  // ---------- state ----------
  let props = null;        // CustomProperties of the current item
  let notes = [];          // [{id, text, color, due, prev, created, updated}]
  let editingId = null;
  let color = COLORS[0];
  let compose = false;     // true when the current item is being written
  let subject = "";
  let composeId = null;    // draft id (compose mode, after the first save)

  const $ = (id) => document.getElementById(id);
  const ok = (r) => r && r.status === Office.AsyncResultStatus.Succeeded;

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;
    buildColors();
    wireUi();
    buildSummaryUi();
    loadItem();
    // When the pane is pinned, follow the selected message
    if (Office.context.mailbox.addHandlerAsync) {
      try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadItem); }
      catch (e) { /* not available in compose */ }
    }
  });

  // ---------- dates ----------
  function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d.getTime(); }
  function dueTs(s) { const p = String(s).split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]).getTime(); }
  // overdue | today | soon (7 days) | later | null (no due date or done)
  function dueState(n) {
    if (!n.due || isDone(n)) return null;
    const t = dueTs(n.due), today = startOfToday();
    if (t < today) return "overdue";
    if (t === today) return "today";
    return t < today + 7 * DAY ? "soon" : "later";
  }
  const pad = (x) => String(x).padStart(2, "0");
  function fmtDue(s) { const p = String(s).split("-"); return p[2] + "/" + p[1] + "/" + p[0].slice(2); }
  function fmtDueLong(s) { const p = String(s).split("-"); return p[2] + "/" + p[1] + "/" + p[0]; }
  function dueText(n) {
    const st = dueState(n);
    if (!st) return "";
    if (st === "overdue") {
      const days = Math.round((startOfToday() - dueTs(n.due)) / DAY);
      return "⏰ Overdue " + days + "d (" + fmtDue(n.due) + ")";
    }
    if (st === "today") return "⏰ Due today";
    return "⏰ Due " + fmtDue(n.due);
  }
  function fmtDate(ts) {
    const d = new Date(ts);
    return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + String(d.getFullYear()).slice(2) +
      " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
  }
  function fmtDay(ts) {
    const d = new Date(ts);
    return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear();
  }

  // ---------- current item ----------
  function loadItem() {
    const item = Office.context.mailbox.item;
    resetComposer();
    composeId = null;
    if (!item) {
      $("subject").textContent = "No email selected";
      notes = []; props = null; renderNotes();
      return;
    }
    compose = !!(item.subject && typeof item.subject.getAsync === "function");
    if (compose) {
      item.subject.getAsync((r) => {
        subject = ok(r) ? r.value || "" : "";
        $("subject").textContent = "✉️ " + (subject || "(new email)");
      });
    } else {
      subject = item.subject || "";
      $("subject").textContent = subject || "(no subject)";
    }
    item.loadCustomPropertiesAsync((res) => {
      if (!ok(res)) { toast("Could not load notes"); return; }
      props = res.value;
      try { notes = JSON.parse(props.get(PROP_KEY) || "[]"); }
      catch (e) { notes = []; }
      renderNotes();
      syncCategory();
      if (!compose) { updateIndex(); updateBanner(); }
      refreshDueBadge();
    });
  }

  function persist(msg) {
    if (!props) return;
    if (notes.length) props.set(PROP_KEY, JSON.stringify(notes));
    else props.remove(PROP_KEY);
    props.saveAsync((res) => {
      if (!ok(res)) { toast("Save failed: " + (res.error && res.error.message)); return; }
      syncCategory();
      if (compose) {
        saveDraft(() => { updateIndex(); refreshDueBadge(); });
      } else {
        updateIndex();
        updateBanner();
        refreshDueBadge();
      }
      toast(msg);
    });
  }

  // In compose mode the draft is saved so the note stays with it and gets an id
  function saveDraft(done) {
    const item = Office.context.mailbox.item;
    const finish = () => {
      if (!item.saveAsync) return done();
      item.saveAsync((r) => { if (ok(r) && r.value) composeId = r.value; done(); });
    };
    if (item.subject && item.subject.getAsync) {
      item.subject.getAsync((r) => { if (ok(r)) subject = r.value || ""; finish(); });
    } else finish();
  }

  // ---------- categories in the message list ----------
  function categoriesSupported() {
    const item = Office.context.mailbox.item;
    return item && item.categories &&
      Office.context.requirements.isSetSupported("Mailbox", "1.8");
  }

  function ensureMasterCategories(names, done) {
    const mc = Office.context.mailbox.masterCategories;
    if (!mc || !names.length) return done();
    mc.getAsync((r) => {
      if (!ok(r)) return done();
      const existing = new Set((r.value || []).map((c) => c.displayName));
      const missing = ALL_CATS
        .filter((l) => names.includes(l.name) && !existing.has(l.name))
        .map((l) => ({ displayName: l.name, color: Office.MailboxEnums.CategoryColor[l.preset] }));
      if (!missing.length) return done();
      mc.addAsync(missing, () => done());
    });
  }

  // One category per note color in use, plus Follow-up while a due date is open
  function syncCategory() {
    if (!categoriesSupported()) return;
    const item = Office.context.mailbox.item;
    const wanted = [...new Set(notes.map((n) => labelFor(n.color).name))];
    if (notes.some((n) => n.due && !isDone(n))) wanted.push(FOLLOW.name);
    const ours = new Set(ALL_CATS.map((l) => l.name).concat(OLD_NAMES));
    item.categories.getAsync((r) => {
      if (!ok(r)) return;
      const current = (r.value || []).map((c) => c.displayName).filter((n) => ours.has(n));
      const toAdd = wanted.filter((n) => !current.includes(n));
      const toRemove = current.filter((n) => !wanted.includes(n));
      if (toRemove.length) item.categories.removeAsync(toRemove, () => {});
      if (toAdd.length) ensureMasterCategories(toAdd, () => item.categories.addAsync(toAdd, () => {}));
    });
  }

  // Yellow info bar on the email when a note is overdue or due today
  function updateBanner() {
    const item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) return;
    const late = notes.filter((n) => ["overdue", "today"].includes(dueState(n)));
    try {
      if (!late.length) { item.notificationMessages.removeAsync(BANNER_KEY, () => {}); return; }
      const over = late.filter((n) => dueState(n) === "overdue").length;
      let msg = "My Notes: " + (over ? over + " overdue" : "") +
        (over && late.length > over ? ", " : "") +
        (late.length > over ? (late.length - over) + " due today" : "") +
        " – " + late[0].text.replace(/\s+/g, " ");
      if (msg.length > 150) msg = msg.slice(0, 147) + "…";
      item.notificationMessages.replaceAsync(BANNER_KEY, {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: msg, icon: "Icon.16x16", persistent: false,
      }, () => {});
    } catch (e) { /* not supported here */ }
  }

  // ---------- notes CRUD ----------
  function saveNote() {
    const text = $("input").value.trim();
    if (!text) return;
    const now = Date.now();
    const due = $("due").value || "";
    if (editingId) {
      const n = notes.find((x) => x.id === editingId);
      if (n) { n.text = text; n.color = color; n.due = due; n.updated = now; }
      persist("Note updated");
    } else {
      notes.push({ id: now.toString(36) + Math.random().toString(36).slice(2, 5),
        text, color, due, created: now, updated: now });
      persist("Note saved");
    }
    resetComposer();
    renderNotes();
  }

  function editNote(id) {
    const n = notes.find((x) => x.id === id);
    if (!n) return;
    editingId = id;
    $("input").value = n.text;
    $("due").value = n.due || "";
    setColor(n.color || COLORS[0]);
    $("save").textContent = "Update";
    $("save").disabled = false;
    $("cancel").hidden = false;
    $("input").focus();
  }

  function toggleDone(id) {
    const n = notes.find((x) => x.id === id);
    if (!n) return;
    if (isDone(n)) {
      n.color = n.prev && n.prev !== DONE_HEX ? n.prev : COLORS[0];
      delete n.prev;
    } else {
      n.prev = labelFor(n.color).hex;
      n.color = DONE_HEX;
    }
    n.updated = Date.now();
    persist(isDone(n) ? "Marked as done" : "Reopened");
    renderNotes();
  }

  function deleteNote(id) {
    if (!window.confirm("Delete this note?")) return;
    notes = notes.filter((x) => x.id !== id);
    if (editingId === id) resetComposer();
    persist("Note deleted");
    renderNotes();
  }

  function resetComposer() {
    editingId = null;
    $("input").value = "";
    $("due").value = "";
    $("save").textContent = "Add";
    $("save").disabled = true;
    $("cancel").hidden = true;
  }

  // ---------- index of all notes (roaming settings) ----------
  function getIndex() {
    try { return Office.context.roamingSettings.get(INDEX_KEY) || {}; }
    catch (e) { return {}; }
  }
  const byteLen = (s) => new TextEncoder().encode(s).length;

  // Entry per email: {id, s: subject, u: last update, c: count, p: preview, draft,
  //                   n: [{id, text, color, updated, due}]}
  // Keyed by internetMessageId (stays the same when an email is moved).
  function updateIndex() {
    const item = Office.context.mailbox.item;
    if (!item) return;
    const id = compose ? composeId : item.itemId;
    if (!id) return;
    const key = compose ? "draft:" + id : (item.internetMessageId || item.itemId);
    const idx = getIndex();
    const before = JSON.stringify(idx);
    const mine = new Set(notes.map((n) => n.id).filter(Boolean));
    // drop older entries of the same email (old key, moved email, or the draft it was sent from)
    Object.keys(idx).forEach((k) => {
      if (k === key) return;
      const v = idx[k];
      if (k === id || v.id === id || (v.n || []).some((n) => n.id && mine.has(n.id))) delete idx[k];
    });
    if (notes.length) {
      const last = notes[notes.length - 1];
      idx[key] = {
        id,
        s: (subject || "(no subject)").slice(0, 100),
        p: last.text.replace(/\s+/g, " ").slice(0, 80),
        c: notes.length,
        u: Math.max.apply(null, notes.map((n) => n.updated || 0)),
        n: notes.map((n) => ({
          id: n.id, text: n.text.slice(0, NOTE_TEXT_MAX), color: labelFor(n.color).hex,
          updated: n.updated, due: n.due || "",
        })),
      };
      if (compose) idx[key].draft = 1;
    } else {
      delete idx[key];
    }
    // stay under the size limit: drop the oldest emails first
    const keys = Object.keys(idx).sort((a, b) => idx[a].u - idx[b].u);
    while (keys.length && byteLen(JSON.stringify(idx)) > INDEX_MAX_BYTES) delete idx[keys.shift()];
    if (JSON.stringify(idx) === before) return;
    Office.context.roamingSettings.set(INDEX_KEY, idx);
    Office.context.roamingSettings.saveAsync();
  }

  // Normalised notes of an index entry (also reads older index formats)
  function entryNotes(v) {
    if (!v.n) return [{ text: v.p || "", color: LABELS[0].hex, updated: v.u, due: "" }];
    return v.n.map((n) => ({
      id: n.id, text: n.text != null ? n.text : n.t || "",
      color: n.color || n.c || LABELS[0].hex,
      updated: n.updated || n.d || v.u, due: n.due || "",
    }));
  }

  // All index entries as [{key, id, v, notes}]
  function allEntries() {
    const idx = getIndex();
    return Object.keys(idx).map((k) => ({ key: k, id: idx[k].id || k, v: idx[k], notes: entryNotes(idx[k]) }));
  }

  function refreshDueBadge() {
    const late = allEntries().reduce((s, e) =>
      s + e.notes.filter((n) => ["overdue", "today"].includes(dueState(n))).length, 0);
    $("dueCount").hidden = !late;
    $("dueCount").textContent = late;
    $("dueCount").title = late + " overdue / due today";
  }

  // earliest open due note of an entry
  function firstDue(ns) {
    return ns.filter((n) => dueState(n)).sort((a, b) => dueTs(a.due) - dueTs(b.due))[0];
  }

  function renderIndex() {
    const q = $("search").value.trim().toLowerCase();
    const f = $("idxFilter").value;
    const list = $("index");
    list.innerHTML = "";
    const rank = { overdue: 0, today: 1, soon: 2, later: 3 };
    const entries = allEntries()
      .filter((e) => !q || (e.v.s + " " + e.notes.map((n) => n.text).join(" ")).toLowerCase().includes(q))
      .filter((e) => {
        if (f === "open") return e.notes.some((n) => !isDone(n));
        if (f === "done") return e.notes.some(isDone);
        if (f === "due") return e.notes.some((n) => dueState(n));
        if (f === "late") return e.notes.some((n) => ["overdue", "today"].includes(dueState(n)));
        return true;
      })
      .sort((a, b) => {
        if (f === "due" || f === "late") {
          const da = firstDue(a.notes), db = firstDue(b.notes);
          return dueTs(da.due) - dueTs(db.due);
        }
        return b.v.u - a.v.u;
      });
    if (!entries.length) {
      list.innerHTML = '<li class="empty">' + (q || f !== "all" ? "No matching notes" : "No emails with notes yet") + "</li>";
      return;
    }
    entries.forEach((e) => {
      const li = document.createElement("li");
      const b = document.createElement("button");
      b.className = "rowbtn";
      b.innerHTML = '<div class="idx-top"><div class="idx-subj"></div><span class="chip"></span></div>' +
        '<div class="idx-prev"><span class="cnt"></span><bdi class="txt"></bdi></div>';
      const open = e.notes.filter((n) => !isDone(n)).length;
      b.querySelector(".idx-subj").textContent = (e.v.draft ? "✉️ " : "") + e.v.s;
      const fd = firstDue(e.notes);
      const chip = b.querySelector(".chip");
      if (fd) { chip.textContent = dueText(fd); chip.classList.add(dueState(fd)); }
      else chip.remove();
      const latest = e.notes.slice().sort((x, y) => y.updated - x.updated)[0];
      b.querySelector(".cnt").textContent =
        e.notes.length + (e.notes.length > 1 ? " notes" : " note") +
        (open < e.notes.length ? " (" + (e.notes.length - open) + " done)" : "") + " · ";
      b.querySelector(".txt").textContent = latest.text.replace(/\s+/g, " ").slice(0, 70);
      b.title = "Open email";
      b.onclick = () => {
        try { Office.context.mailbox.displayMessageForm(e.id); }
        catch (err) { toast("Can't open – the email may have been moved or deleted"); }
      };
      li.appendChild(b);
      list.appendChild(li);
    });
  }

  // ---------- document numbers (CDF etc.) ----------
  const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  function parsePrefixes(s) {
    return String(s || "").split(/[,;\s]+/).map((x) => x.trim()).filter(Boolean);
  }
  function refsOf(text, prefixes) {
    const out = new Set();
    prefixes.forEach((p) => {
      const re = new RegExp("(^|[^A-Za-z0-9])(" + escRe(p) + ")[\\s\\-_#:]*(\\d[A-Za-z0-9.\\-/]*)", "gi");
      let m;
      while ((m = re.exec(text))) {
        const num = m[3].replace(/[.\-/]+$/, "").toUpperCase();
        out.add(p.toUpperCase() + (p.length > 1 ? " " : "") + num);
      }
    });
    return [...out];
  }

  // ---------- summary & Excel ----------
  function getPrefs() {
    let p = null;
    try { p = Office.context.roamingSettings.get(PREFS_KEY); } catch (e) { /* ignore */ }
    return Object.assign({
      labels: LABELS.filter((l) => l.hex !== DONE_HEX).map((l) => l.hex),
      range: "all",
      group: "label",
      prefixes: "CDF",
      dueTop: true,
      to: (Office.context.mailbox.userProfile || {}).emailAddress || "",
    }, p || {});
  }

  function readPrefs() {
    return {
      labels: [...document.querySelectorAll("#sumLabels input:checked")].map((i) => i.value),
      range: $("sumRange").value,
      group: $("sumGroup").value,
      prefixes: $("sumPrefixes").value.trim(),
      dueTop: $("sumDueTop").checked,
      to: $("sumTo").value.trim(),
    };
  }

  function savePrefs() {
    const prefs = readPrefs();
    Office.context.roamingSettings.set(PREFS_KEY, prefs);
    Office.context.roamingSettings.saveAsync();
    return prefs;
  }

  function buildSummaryUi() {
    const prefs = getPrefs();
    const wrap = $("sumLabels");
    LABELS.forEach((l) => {
      const lab = document.createElement("label");
      lab.className = "check";
      lab.innerHTML = '<input type="checkbox"><span class="dot"></span><span class="nm"></span>';
      const cb = lab.querySelector("input");
      cb.value = l.hex;
      cb.checked = prefs.labels.includes(l.hex);
      cb.onchange = renderSummaryInfo;
      lab.querySelector(".dot").style.background = l.hex;
      lab.querySelector(".nm").textContent = l.name;
      wrap.appendChild(lab);
    });
    $("sumRange").value = prefs.range;
    $("sumGroup").value = prefs.group;
    $("sumPrefixes").value = prefs.prefixes;
    $("sumDueTop").checked = prefs.dueTop;
    $("sumTo").value = prefs.to;
    ["sumRange", "sumGroup", "sumDueTop"].forEach((id) => { $(id).onchange = renderSummaryInfo; });
    $("sumPrefixes").oninput = renderSummaryInfo;
    $("sumGo").onclick = createSummary;
    $("sumCopy").onclick = copyForExcel;
  }

  function rangeStart(range) {
    if (range === "all") return 0;
    const t = startOfToday();
    return range === "today" ? t : t - (parseInt(range, 10) - 1) * DAY;
  }

  // Notes matching the chosen categories and period: [{e, n}]
  function selectedNotes(prefs) {
    const from = rangeStart(prefs.range);
    const out = [];
    allEntries().sort((a, b) => b.v.u - a.v.u).forEach((e) => {
      e.notes.forEach((n) => {
        if (prefs.labels.includes(labelFor(n.color).hex) && (n.updated || 0) >= from) out.push({ e, n });
      });
    });
    return out;
  }

  // Groups: [{title, color, mails: [{e, notes}]}]
  function buildGroups(prefs) {
    const rows = selectedNotes(prefs);
    const groups = [];
    const addTo = (map, title, color, r) => {
      if (!map.has(title)) { const g = { title, color, mails: new Map() }; map.set(title, g); }
      const g = map.get(title);
      if (!g.mails.has(r.e.key)) g.mails.set(r.e.key, { e: r.e, notes: [] });
      g.mails.get(r.e.key).notes.push(r.n);
    };
    if (prefs.dueTop) {
      const late = new Map();
      allEntries().forEach((e) => e.notes.forEach((n) => {
        if (["overdue", "today"].includes(dueState(n))) addTo(late, "⏰ Overdue & due today", "#e07a00", { e, n });
      }));
      late.forEach((g) => groups.push(g));
    }
    const map = new Map();
    if (prefs.group === "ref") {
      const prefixes = parsePrefixes(prefs.prefixes);
      rows.forEach((r) => {
        const refs = refsOf(r.e.v.s + " " + r.n.text, prefixes);
        (refs.length ? refs : ["No document number"]).forEach((ref) => addTo(map, ref, "#8a8378", r));
      });
      [...map.keys()].sort((a, b) => (a === "No document number") - (b === "No document number") ||
        a.localeCompare(b, undefined, { numeric: true })).forEach((k) => groups.push(map.get(k)));
    } else {
      rows.forEach((r) => { const l = labelFor(r.n.color); addTo(map, l.name, l.hex, r); });
      LABELS.forEach((l) => { if (map.has(l.name)) groups.push(map.get(l.name)); });
    }
    // Map → arrays, notes sorted (open due first, then newest)
    return groups.map((g) => ({
      title: g.title, color: g.color,
      mails: [...g.mails.values()].map((m) => ({
        e: m.e,
        notes: m.notes.sort((a, b) => (a.due && !isDone(a) ? dueTs(a.due) : 9e15) -
          (b.due && !isDone(b) ? dueTs(b.due) : 9e15) || b.updated - a.updated),
      })),
    }));
  }

  function countRows(rows) {
    return { notes: rows.length, mails: new Set(rows.map((r) => r.e.key)).size };
  }

  function renderSummaryInfo() {
    const prefs = readPrefs();
    const c = countRows(selectedNotes(prefs));
    const late = allEntries().reduce((s, e) =>
      s + e.notes.filter((n) => ["overdue", "today"].includes(dueState(n))).length, 0);
    let txt = c.notes ? c.notes + " notes from " + c.mails + " emails" : "No notes match your selection";
    if (prefs.dueTop && late) txt += " · " + late + " overdue / due today";
    $("sumInfo").textContent = txt;
    $("sumGo").disabled = !c.notes && !(prefs.dueTop && late);
    $("sumCopy").disabled = !c.notes;
  }

  const esc = (s) => String(s).replace(/[&<>"]/g,
    (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
  // rtl for text that starts in Hebrew/Arabic, ltr otherwise
  function dirOf(s) {
    const m = String(s).match(/[A-Za-z֐-׿؀-ۿ]/);
    return m && /[֐-ۿ]/.test(m[0]) ? "rtl" : "ltr";
  }

  function messageLink(itemId) {
    try {
      const rest = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
      return "https://outlook.office365.com/owa/?ItemID=" + encodeURIComponent(rest) +
        "&exvsurl=1&viewmodel=ReadMessageItem";
    } catch (e) { return ""; }
  }

  const DUE_STYLE = {
    overdue: "background:#fde2e1;color:#b42318;font-weight:600",
    today: "background:#fdecd0;color:#9a5800;font-weight:600",
    soon: "background:#e4eefb;color:#1f5fa8", later: "background:#e4eefb;color:#1f5fa8",
  };
  const tag = (text, style) =>
    '<span style="font-size:11px;padding:0 5px;border-radius:6px;' + style + '">' + esc(text) + "</span>";

  function summaryHtml(groups, prefs, maxText) {
    const rows = selectedNotes(prefs);
    const c = countRows(rows);
    const rangeTxt = $("sumRange").selectedOptions[0].textContent;
    let html =
      '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222">' +
      '<div style="font-size:18px;font-weight:600">My Notes – summary ' + esc(fmtDay(Date.now())) + "</div>" +
      '<div style="color:#777;font-size:12px;margin-bottom:6px">' + c.notes + " notes in " + c.mails +
      " emails · " + esc(rangeTxt) + "</div>";
    let skipped = 0;
    groups.forEach((g) => {
      const head = '<div style="margin:16px 0 6px;padding:4px 8px;background:#f4f4f4;border-left:5px solid ' +
        g.color + ';font-weight:600">' + esc(g.title) + " (" + g.mails.length + ")</div>";
      if (html.length + head.length > BODY_MAX) { skipped += g.mails.length; return; }
      html += head;
      g.mails.forEach((m) => {
        const link = messageLink(m.e.id);
        const subj = esc((m.e.v.draft ? "✉️ " : "") + m.e.v.s);
        let block = '<div style="margin:0 0 10px"><div dir="' + dirOf(m.e.v.s) + '"><b>' +
          (link ? '<a href="' + esc(link) + '">' + subj + "</a>" : subj) + "</b></div>" +
          "";
        m.notes.forEach((n) => {
          const t = n.text.length > maxText ? n.text.slice(0, maxText) + "…" : n.text;
          const l = labelFor(n.color);
          const st = dueState(n);
          let meta = esc(fmtDate(n.updated));
          if (prefs.group === "ref" || g.title.startsWith("⏰")) meta = esc(l.name) + " · " + meta;
          if (st) meta += " " + tag(dueText(n), DUE_STYLE[st]);
          else if (isDone(n)) meta += " " + tag("✓ Done", "background:#e3f3e3;color:#2e7d32");
          block += '<div style="margin:4px 0 6px 2px;padding-left:9px;border-left:3px solid ' + l.hex + '">' +
            '<div dir="' + dirOf(t) + '" style="text-align:left">' + esc(t).replace(/\n/g, "<br>") + "</div>" +
            '<div style="color:#999;font-size:11px">' + meta + "</div></div>";
        });
        block += "</div>";
        if (html.length + block.length > BODY_MAX) { skipped++; return; }
        html += block;
      });
    });
    if (skipped) {
      html += '<div style="color:#b42318;font-size:12px">Summary shortened – ' + skipped +
        " emails did not fit (size limit). Choose fewer categories or a shorter period.</div>";
    }
    return { html: html + "</div>", skipped };
  }

  function createSummary() {
    const prefs = savePrefs();
    const groups = buildGroups(prefs);
    if (!groups.length) return toast("No notes to summarise");
    let out = summaryHtml(groups, prefs, NOTE_TEXT_MAX);
    if (out.skipped) out = summaryHtml(groups, prefs, 120); // shorten texts before dropping emails
    const opts = { subject: "My Notes – summary " + fmtDay(Date.now()), htmlBody: out.html };
    if (prefs.to) opts.toRecipients = prefs.to.split(/[;,\s]+/).filter(Boolean);
    try { Office.context.mailbox.displayNewMessageForm(opts); }
    catch (e) { toast("Could not open a new email: " + e.message); }
  }

  function copyForExcel() {
    const prefs = savePrefs();
    const rows = selectedNotes(prefs);
    if (!rows.length) return toast("No notes to copy");
    const prefixes = parsePrefixes(prefs.prefixes);
    const head = ["Document no.", "Subject", "Category", "Note", "Updated", "Due", "Status", "Link"];
    const data = rows.map((r) => {
      const st = dueState(r.n);
      const status = isDone(r.n) ? "Done" : st === "overdue" ? "Overdue" : st === "today" ? "Due today" : "Open";
      return [
        refsOf(r.e.v.s + " " + r.n.text, prefixes).join(", "),
        r.e.v.s, labelFor(r.n.color).name.replace(/^\S+\s/, ""), r.n.text,
        fmtDate(r.n.updated), r.n.due ? fmtDueLong(r.n.due) : "", status,
        messageLink(r.e.id),
      ];
    });
    const tsv = [head].concat(data).map((row) =>
      row.map((c) => /[\t\n"]/.test(c) ? '"' + String(c).replace(/"/g, '""') + '"' : c).join("\t")).join("\r\n");
    const html = "<table><tr>" + head.map((h) => "<th>" + esc(h) + "</th>").join("") + "</tr>" +
      data.map((row) => "<tr>" + row.map((c, i) => "<td>" +
        (i === 7 && c ? '<a href="' + esc(c) + '">Open</a>' : esc(c).replace(/\n/g, "<br>")) +
        "</td>").join("") + "</tr>").join("") + "</table>";
    copyRich(html, tsv).then((done) => {
      if (done) { $("copyBox").hidden = true; toast(rows.length + " rows copied – paste into Excel (Ctrl+V)"); }
      else {
        const box = $("copyBox");
        box.hidden = false; box.value = tsv; box.focus(); box.select();
        toast("Press Ctrl+C to copy, then paste into Excel");
      }
    });
  }

  async function copyRich(html, text) {
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        })]);
        return true;
      }
    } catch (e) { /* fall through */ }
    try {
      const d = document.createElement("div");
      d.contentEditable = "true";
      d.style.cssText = "position:fixed;left:-9999px;top:0";
      d.innerHTML = html;
      document.body.appendChild(d);
      const r = document.createRange();
      r.selectNodeContents(d);
      const sel = window.getSelection();
      sel.removeAllRanges(); sel.addRange(r);
      const done = document.execCommand("copy");
      sel.removeAllRanges(); d.remove();
      return done;
    } catch (e) { return false; }
  }

  // ---------- rendering ----------
  function renderNotes() {
    $("count").textContent = notes.length;
    const ul = $("notes");
    ul.innerHTML = "";
    if (!notes.length) {
      ul.innerHTML = '<li class="empty">No notes on this email</li>';
      return;
    }
    // open notes first (earliest due first), done notes at the bottom
    const sorted = notes.slice().sort((a, b) =>
      isDone(a) - isDone(b) ||
      (a.due && !isDone(a) ? dueTs(a.due) : 9e15) - (b.due && !isDone(b) ? dueTs(b.due) : 9e15) ||
      b.updated - a.updated);
    sorted.forEach((n) => {
      const li = document.createElement("li");
      const lbl = labelFor(n.color);
      const done = isDone(n);
      li.className = "note" + (done ? " done" : "");
      li.style.setProperty("--c", lbl.hex);
      const edited = n.updated !== n.created ? " (edited)" : "";
      li.innerHTML =
        '<div class="note-text" dir="auto"></div>' +
        '<div class="note-meta"><span><span class="lbl"></span> · <span class="dt"></span> <span class="chip"></span></span>' +
        '<span class="note-actions"><button class="tick"></button><button class="edit">Edit</button>' +
        '<button class="del">Delete</button></span></div>';
      li.querySelector(".note-text").textContent = n.text;
      li.querySelector(".lbl").textContent = lbl.name;
      li.querySelector(".dt").textContent = fmtDate(n.updated) + edited;
      const chip = li.querySelector(".chip");
      const st = dueState(n);
      if (st) { chip.textContent = dueText(n); chip.classList.add(st); } else chip.remove();
      const tick = li.querySelector(".tick");
      tick.textContent = done ? "↺ Reopen" : "✓ Done";
      tick.title = done ? "Reopen this note" : "Mark as done";
      tick.onclick = () => toggleDone(n.id);
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
    $("input").addEventListener("input", () => { $("save").disabled = !$("input").value.trim(); });
    $("input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveNote(); }
      if (e.key === "Escape") resetComposer();
    });
    $("save").onclick = saveNote;
    $("cancel").onclick = resetComposer;
    $("dueClear").onclick = () => { $("due").value = ""; };
    $("tab-item").onclick = () => switchTab("item");
    $("tab-all").onclick = () => switchTab("all");
    $("tab-sum").onclick = () => switchTab("sum");
    $("search").addEventListener("input", renderIndex);
    $("idxFilter").onchange = renderIndex;
  }

  function switchTab(which) {
    ["item", "all", "sum"].forEach((t) => {
      $("tab-" + t).setAttribute("aria-selected", which === t);
      $("view-" + t).hidden = which !== t;
    });
    if (which === "all") renderIndex();
    if (which === "sum") renderSummaryInfo();
  }

  let toastTimer;
  function toast(msg) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 2200);
  }
})();
