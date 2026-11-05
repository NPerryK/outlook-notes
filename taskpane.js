/* My Notes – Outlook add-in. Includes lz-string 1.5.0 (c) 2013 Pieroxy, MIT license. */
var LZString=function(){var r=String.fromCharCode,o="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",n="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+-$",e={};function t(r,o){if(!e[r]){e[r]={};for(var n=0;n<r.length;n++)e[r][r.charAt(n)]=n}return e[r][o]}var i={compressToBase64:function(r){if(null==r)return"";var n=i._compress(r,6,function(r){return o.charAt(r)});switch(n.length%4){default:case 0:return n;case 1:return n+"===";case 2:return n+"==";case 3:return n+"="}},decompressFromBase64:function(r){return null==r?"":""==r?null:i._decompress(r.length,32,function(n){return t(o,r.charAt(n))})},compressToUTF16:function(o){return null==o?"":i._compress(o,15,function(o){return r(o+32)})+" "},decompressFromUTF16:function(r){return null==r?"":""==r?null:i._decompress(r.length,16384,function(o){return r.charCodeAt(o)-32})},compressToUint8Array:function(r){for(var o=i.compress(r),n=new Uint8Array(2*o.length),e=0,t=o.length;e<t;e++){var s=o.charCodeAt(e);n[2*e]=s>>>8,n[2*e+1]=s%256}return n},decompressFromUint8Array:function(o){if(null==o)return i.decompress(o);for(var n=new Array(o.length/2),e=0,t=n.length;e<t;e++)n[e]=256*o[2*e]+o[2*e+1];var s=[];return n.forEach(function(o){s.push(r(o))}),i.decompress(s.join(""))},compressToEncodedURIComponent:function(r){return null==r?"":i._compress(r,6,function(r){return n.charAt(r)})},decompressFromEncodedURIComponent:function(r){return null==r?"":""==r?null:(r=r.replace(/ /g,"+"),i._decompress(r.length,32,function(o){return t(n,r.charAt(o))}))},compress:function(o){return i._compress(o,16,function(o){return r(o)})},_compress:function(r,o,n){if(null==r)return"";var e,t,i,s={},u={},a="",p="",c="",l=2,f=3,h=2,d=[],m=0,v=0;for(i=0;i<r.length;i+=1)if(a=r.charAt(i),Object.prototype.hasOwnProperty.call(s,a)||(s[a]=f++,u[a]=!0),p=c+a,Object.prototype.hasOwnProperty.call(s,p))c=p;else{if(Object.prototype.hasOwnProperty.call(u,c)){if(c.charCodeAt(0)<256){for(e=0;e<h;e++)m<<=1,v==o-1?(v=0,d.push(n(m)),m=0):v++;for(t=c.charCodeAt(0),e=0;e<8;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}else{for(t=1,e=0;e<h;e++)m=m<<1|t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t=0;for(t=c.charCodeAt(0),e=0;e<16;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}0==--l&&(l=Math.pow(2,h),h++),delete u[c]}else for(t=s[c],e=0;e<h;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1;0==--l&&(l=Math.pow(2,h),h++),s[p]=f++,c=String(a)}if(""!==c){if(Object.prototype.hasOwnProperty.call(u,c)){if(c.charCodeAt(0)<256){for(e=0;e<h;e++)m<<=1,v==o-1?(v=0,d.push(n(m)),m=0):v++;for(t=c.charCodeAt(0),e=0;e<8;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}else{for(t=1,e=0;e<h;e++)m=m<<1|t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t=0;for(t=c.charCodeAt(0),e=0;e<16;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}0==--l&&(l=Math.pow(2,h),h++),delete u[c]}else for(t=s[c],e=0;e<h;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1;0==--l&&(l=Math.pow(2,h),h++)}for(t=2,e=0;e<h;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1;for(;;){if(m<<=1,v==o-1){d.push(n(m));break}v++}return d.join("")},decompress:function(r){return null==r?"":""==r?null:i._decompress(r.length,32768,function(o){return r.charCodeAt(o)})},_decompress:function(o,n,e){var t,i,s,u,a,p,c,l=[],f=4,h=4,d=3,m="",v=[],g={val:e(0),position:n,index:1};for(t=0;t<3;t+=1)l[t]=t;for(s=0,a=Math.pow(2,2),p=1;p!=a;)u=g.val&g.position,g.position>>=1,0==g.position&&(g.position=n,g.val=e(g.index++)),s|=(u>0?1:0)*p,p<<=1;switch(s){case 0:for(s=0,a=Math.pow(2,8),p=1;p!=a;)u=g.val&g.position,g.position>>=1,0==g.position&&(g.position=n,g.val=e(g.index++)),s|=(u>0?1:0)*p,p<<=1;c=r(s);break;case 1:for(s=0,a=Math.pow(2,16),p=1;p!=a;)u=g.val&g.position,g.position>>=1,0==g.position&&(g.position=n,g.val=e(g.index++)),s|=(u>0?1:0)*p,p<<=1;c=r(s);break;case 2:return""}for(l[3]=c,i=c,v.push(c);;){if(g.index>o)return"";for(s=0,a=Math.pow(2,d),p=1;p!=a;)u=g.val&g.position,g.position>>=1,0==g.position&&(g.position=n,g.val=e(g.index++)),s|=(u>0?1:0)*p,p<<=1;switch(c=s){case 0:for(s=0,a=Math.pow(2,8),p=1;p!=a;)u=g.val&g.position,g.position>>=1,0==g.position&&(g.position=n,g.val=e(g.index++)),s|=(u>0?1:0)*p,p<<=1;l[h++]=r(s),c=h-1,f--;break;case 1:for(s=0,a=Math.pow(2,16),p=1;p!=a;)u=g.val&g.position,g.position>>=1,0==g.position&&(g.position=n,g.val=e(g.index++)),s|=(u>0?1:0)*p,p<<=1;l[h++]=r(s),c=h-1,f--;break;case 2:return v.join("")}if(0==f&&(f=Math.pow(2,d),d++),l[c])m=l[c];else{if(c!==h)return null;m=i+i.charAt(0)}v.push(m),l[h++]=i+m.charAt(0),i=m,0==--f&&(f=Math.pow(2,d),d++)}}};return i}();"function"==typeof define&&define.amd?define(function(){return LZString}):"undefined"!=typeof module&&null!=module?module.exports=LZString:"undefined"!=typeof angular&&null!=angular&&angular.module("LZString",[]).factory("LZString",function(){return LZString});
/* My Notes 2.0 – Outlook add-in
 * Notes live on each email (custom properties, ~2,500 characters per email).
 * A compressed overview in roamingSettings powers All notes, Labels, Summary and Excel.
 */
(function () {
  "use strict";

  // ===================================================================
  // Configuration
  // ===================================================================
  const VERSION = "3.0.0";
  const PROP_KEY = "mailNotes";          // custom property on the item
  const LBL_KEY = "mailLabels";          // labels on the email itself (no note needed)
  const DKEY_PROP = "mailNotesKey";      // key of a draft, so its notes can be reattached to the sent copy
  const DRAFTS_KEY = "myNotesDrafts";    // roaming: notes written while composing, waiting for the sent copy
  const HDR = "X-MyNotes-Key";           // hidden header that travels with the sent email
  const CP_LIMIT = 2400;                 // Outlook limit: 2,500 characters per add-in per item (measured in bytes to be safe)
  const IDX_KEY = "myNotesIdx2";         // compressed overview (roaming)
  const OLD_IDX_KEY = "mailNotesIndex";  // overview of versions 1.x
  const PREFS_KEY = "mailNotesSummary";
  const TAGS_KEY = "myNotesTags";
  const TPL_KEY = "myNotesTemplates";
  const PEND_KEY = "myNotesPending";     // changes for emails that are not open (e.g. "mark done")
  const DISM_KEY = "myNotesDismissed";   // dismissed "reply received" hints
  const BANNER_KEY = "myNotesDue";
  const IDX_BUDGET = 24000;              // bytes for the overview (roaming total is 32KB)
  const NOTE_TEXT_MAX = 300;             // note text kept in the overview
  const BODY_MAX = 30000;                // new-message body limit is ~32KB
  const TAG_PREFIX = "🏷 ";
  const TAG_MAX = 40;
  const WEEKEND = [5, 6];                // Friday, Saturday (Israeli work week)
  const DAY = 864e5;
  const BACKUP_SUBJECT = "My Notes backup";

  // Status = note color = Outlook category in the message list
  const STATUS = [
    { hex: "#e0b94a", name: "📝 Note",           short: "Note",     preset: "Preset3",  tip: "Something to do or follow" },
    { hex: "#d9534f", name: "🔴 Urgent",         short: "Urgent",   preset: "Preset0",  tip: "Needs attention first" },
    { hex: "#4a9be0", name: "🔵 To Review",      short: "Review",   preset: "Preset7",  tip: "Check a document or an answer" },
    { hex: "#a066d3", name: "🟣 Awaiting Reply", short: "Waiting",  preset: "Preset8",  tip: "Waiting for someone to answer" },
    { hex: "#5cb85c", name: "🟢 Done",           short: "Done",     preset: "Preset4",  tip: "Finished" },
    { hex: "#8a94a0", name: "ℹ️ Info",           short: "Info",     preset: "Preset12", tip: "Just information, no action needed" },
  ];
  const NOTE_HEX = STATUS[0].hex, URGENT_HEX = STATUS[1].hex, REVIEW_HEX = STATUS[2].hex;
  const WAIT_HEX = STATUS[3].hex, DONE_HEX = STATUS[4].hex, INFO_HEX = STATUS[5].hex;
  // Stages a label (e.g. a CDF) can be in
  const STAGES = [
    { id: "open", name: "Open", color: "#e0b94a" }, { id: "supplier", name: "With supplier", color: "#a066d3" },
    { id: "stress", name: "With Stress", color: "#4a9be0" }, { id: "der", name: "DER review", color: "#d9822b" },
    { id: "closed", name: "Closed", color: "#5cb85c" },
  ];
  const FOLLOW = { name: "⏰ Follow-up", preset: "Preset1" };
  const ALL_CATS = STATUS.concat([FOLLOW]);
  const LEGACY = { "#e07a4a": URGENT_HEX };
  const OLD_NAMES = ["📝 הערה", "🔴 דחוף", "🔵 לבדיקה", "🟣 ממתין לתשובה", "🟢 טופל"];
  const statusOf = (hex) => STATUS.find((s) => s.hex === (LEGACY[hex] || hex)) || STATUS[0];
  const statusIdx = (hex) => STATUS.indexOf(statusOf(hex));
  const isDone = (n) => statusOf(n.color).hex === DONE_HEX;
  const isInfo = (n) => statusOf(n.color).hex === INFO_HEX;
  const isWait = (n) => statusOf(n.color).hex === WAIT_HEX;
  const isOpen = (n) => !isDone(n) && !isInfo(n);
  const TAG_PALETTE = [
    ["Preset5", "#2a9d8f"], ["Preset9", "#b5306a"], ["Preset10", "#5b7fa6"], ["Preset6", "#8a8f2e"],
    ["Preset16", "#c8641c"], ["Preset19", "#2e7d32"], ["Preset22", "#2f4f9e"], ["Preset23", "#6b3fa0"],
    ["Preset20", "#1f6f6a"], ["Preset24", "#8e2453"],
  ];
  const DEFAULT_TEMPLATES = [
    { id: "t1", text: "Sent to Stress for review", color: REVIEW_HEX, days: 5, tags: [] },
    { id: "t2", text: "Waiting for {from}", color: WAIT_HEX, days: 3, tags: [] },
    { id: "t3", text: "Needs DER approval", color: URGENT_HEX, days: 2, tags: [] },
    { id: "t4", text: "Check drawing / CDF update", color: NOTE_HEX, days: null, tags: [] },
  ];
  const EXCEL_COLS = [
    ["ref", "Document no."], ["labels", "Labels"], ["subject", "Subject"], ["status", "Category"],
    ["note", "Note"], ["file", "Attachment"], ["created", "Created"], ["updated", "Updated"],
    ["due", "Due"], ["state", "Status"], ["link", "Link"],
  ];

  // ===================================================================
  // State
  // ===================================================================
  let ctx = null;          // {item, compose, meeting, conv, attachments}
  let props = null;        // CustomProperties of the current item
  let notes = [];          // [{id, text, color, due, tags, att, prev, created, updated, doneAt, hist}]
  let emailTags = [];      // labels on the email itself
  let lblAt = 0;           // when the email labels last changed
  let subject = "";
  let composeId = null;
  let editingId = null;
  let color = NOTE_HEX;
  let draftTags = [];
  let showDone = false;
  let threadOpen = true;
  let replyCands = [];
  let backupFound = null;
  let openTagName = null;
  let tab = "item";
  let draftKey = null;     // key of the open draft (compose mode)
  let person = null;       // {name, first} – sender (read) or first recipient (compose)
  const openHist = new Set();

  const MB = () => Office.context.mailbox;
  const RS = () => Office.context.roamingSettings;
  const $ = (id) => document.getElementById(id);
  const ok = (r) => r && r.status === Office.AsyncResultStatus.Succeeded;
  const lc = (s) => String(s).toLowerCase();
  const byteLen = (s) => new TextEncoder().encode(s).length;
  const pl = (n, w) => n + " " + w + (n === 1 ? "" : "s");
  const newId = () => (Date.now() % 1e10).toString(36) + Math.random().toString(36).slice(2, 5);

  // ===================================================================
  // Small DOM helpers and icons
  // ===================================================================
  function h(tag, attrs) {
    const e = document.createElement(tag);
    if (attrs) {
      for (const k in attrs) {
        const v = attrs[k];
        if (v == null || v === false) continue;
        if (k === "class") e.className = v;
        else if (k === "text") e.textContent = v;
        else if (k === "html") e.innerHTML = v;
        else if (k === "style") e.style.cssText = v;
        else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
        else e.setAttribute(k, v === true ? "" : v);
      }
    }
    for (let i = 2; i < arguments.length; i++) {
      [].concat(arguments[i]).forEach((c) => {
        if (c == null || c === false) return;
        e.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
      });
    }
    return e;
  }
  const ICONS = {
    note: '<path d="M3 2.5h7l3 3v8H3z"/><path d="M10 2.5v3h3"/><path d="M5.5 8h5M5.5 10.5h3.5"/>',
    check: '<path d="M3 8.5l3 3 7-7"/>',
    undo: '<path d="M4 5.5h6a3 3 0 010 6H7"/><path d="M6 3.5l-2 2 2 2"/>',
    edit: '<path d="M10.5 2.5l3 3L6 13H3v-3z"/>',
    trash: '<path d="M2.5 4.5h11"/><path d="M6 4.5V3h4v1.5"/><path d="M4 4.5l.7 9h6.6l.7-9"/>',
    cal: '<rect x="2.5" y="3.5" width="11" height="10" rx="1.5"/><path d="M2.5 6.5h11M5.5 2v3M10.5 2v3"/>',
    clock: '<circle cx="8" cy="8" r="5.5"/><path d="M8 5v3l2 1.5"/>',
    tag: '<path d="M2.5 2.5h5l6 6-5 5-6-6z"/><circle cx="5.3" cy="5.3" r=".8"/>',
    bolt: '<path d="M9 1.5L3.5 9H8l-1 5.5L12.5 7H8z"/>',
    thread: '<path d="M2.5 3h8v5.5H6L3.5 10.5V8.5h-1z"/><path d="M10.5 5.5h3V11h-1v2l-2.5-2H7"/>',
    reply: '<path d="M6.5 4L3 7.5 6.5 11"/><path d="M3 7.5h6a4 4 0 014 4v1"/>',
    replyAll: '<path d="M8.5 4L5 7.5 8.5 11"/><path d="M5.5 4L2 7.5 5.5 11"/><path d="M5 7.5h4.5a4 4 0 014 4v1"/>',
    mail: '<rect x="2" y="3.5" width="12" height="9" rx="1.5"/><path d="M2.5 4.5L8 9l5.5-4.5"/>',
    compose: '<path d="M13.5 8.5V13a.5.5 0 01-.5.5H3a.5.5 0 01-.5-.5V3a.5.5 0 01.5-.5h4.5"/><path d="M11 2l3 3-5.5 5.5H5.5v-3z"/>',
    plus: '<path d="M8 3v10M3 8h10"/>',
    chevR: '<path d="M6 3.5L10.5 8 6 12.5"/>',
    chevD: '<path d="M3.5 6L8 10.5 12.5 6"/>',
    back: '<path d="M9.5 3.5L5 8l4.5 4.5"/>',
    copy: '<rect x="5" y="5" width="8.5" height="8.5" rx="1.5"/><path d="M11 5V3.5a1 1 0 00-1-1H3.5a1 1 0 00-1 1V10a1 1 0 001 1H5"/>',
    chart: '<path d="M2.5 13.5h11"/><path d="M4.5 11V8M8 11V4.5M11.5 11V6.5"/>',
    archive: '<rect x="2" y="3" width="12" height="3" rx="1"/><path d="M3 6v6.5a1 1 0 001 1h8a1 1 0 001-1V6"/><path d="M6.5 9h3"/>',
    clip: '<path d="M12 7.5l-4.6 4.6a2.6 2.6 0 01-3.7-3.7L8.6 3.5a1.7 1.7 0 012.4 2.4L6.3 10.6a.8.8 0 01-1.1-1.1L9.5 5.2"/>',
    history: '<path d="M2.7 8a5.3 5.3 0 101.4-3.6"/><path d="M3.5 2.5v2.5H6"/><path d="M8 5.5V8l2 1.2"/>',
    insert: '<path d="M8 2.5v7M5 6.5l3 3 3-3"/><path d="M3 12.5h10"/>',
    x: '<path d="M4 4l8 8M12 4l-8 8"/>',
    alert: '<path d="M8 2.2l6 10.8H2z"/><path d="M8 6.5v3"/><path d="M8 11.4v.1"/>',
    send: '<path d="M2.5 7.5l11-5-4.5 11-2-4z"/><path d="M7 9.5l6.5-7"/>',
    list: '<path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><path d="M2.5 4h.1M2.5 8h.1M2.5 12h.1"/>',
    search: '<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5l3 3"/>',
    snooze: '<circle cx="6.5" cy="9.5" r="4.5"/><path d="M6.5 7v2.5l1.6 1.2"/><path d="M10 2.5h3.5L10 6h3.5"/>',
    user: '<circle cx="8" cy="5.5" r="2.8"/><path d="M2.8 13.5a5.2 5.2 0 0110.4 0"/>',
    kbd: '<rect x="2" y="4" width="12" height="8.5" rx="1.5"/><path d="M4.5 7h1M7.5 7h1M10.5 7h1M4.5 9.5h7"/>',
    stage: '<path d="M2.5 3.5h11"/><path d="M2.5 8h7"/><path d="M2.5 12.5h4"/>',
    text: '<path d="M3 4h10M3 7h10M3 10h6"/>',
  };
  function svg(name, size) {
    const s = size || 16;
    return '<svg viewBox="0 0 16 16" width="' + s + '" height="' + s + '" fill="none" stroke="currentColor" ' +
      'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ICONS[name] + "</svg>";
  }
  const ic = (name, size) => h("span", { class: "ic", html: svg(name, size) });
  function iconBtn(name, title, onClick, cls) {
    return h("button", { class: "ibtn" + (cls ? " " + cls : ""), title, "aria-label": title, html: svg(name), onclick: onClick });
  }

  // ===================================================================
  // Dates (Israeli work week: Sunday–Thursday)
  // ===================================================================
  const pad = (x) => String(x).padStart(2, "0");
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d.getTime(); }
  function iso(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function dueTs(s) { const p = String(s).split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]).getTime(); }
  function addWorkdays(from, n) {
    const d = new Date(from); d.setHours(0, 0, 0, 0);
    while (n > 0) { d.setDate(d.getDate() + 1); if (!WEEKEND.includes(d.getDay())) n--; }
    return d;
  }
  function workdaysBetween(fromTs, toTs) {
    let n = 0; const d = new Date(fromTs); d.setHours(0, 0, 0, 0);
    while (d.getTime() < toTs) { d.setDate(d.getDate() + 1); if (!WEEKEND.includes(d.getDay())) n++; }
    return n;
  }
  function nextSunday() { const d = new Date(startOfToday()); d.setDate(d.getDate() + ((7 - d.getDay()) % 7 || 7)); return d; }
  function weekStart() { const d = new Date(startOfToday()); d.setDate(d.getDate() - d.getDay()); return d.getTime(); }
  function waitDays(n) { return workdaysBetween(n.since || n.created || n.updated || Date.now(), startOfToday()); }
  function dueState(n) {
    if (!n.due || isDone(n) || isInfo(n)) return null;
    const t = dueTs(n.due), today = startOfToday();
    if (t < today) return "overdue";
    if (t === today) return "today";
    return t < today + 7 * DAY ? "soon" : "later";
  }
  const isLate = (n) => { const s = dueState(n); return s === "overdue" || s === "today"; };
  function fmtShort(ts) {
    const d = new Date(ts), y = new Date().getFullYear();
    return d.getDate() + " " + MONTHS[d.getMonth()] + (d.getFullYear() !== y ? " " + d.getFullYear() : "");
  }
  function fmtDate(ts) { const d = new Date(ts); return fmtShort(ts) + ", " + pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function fmtDay(ts) { const d = new Date(ts); return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear(); }
  function fmtXl(ts) { const d = new Date(ts); return fmtDay(ts) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function fmtDueLong(s) { const p = String(s).split("-"); return p[2] + "/" + p[1] + "/" + p[0]; }
  function dueText(n) {
    const st = dueState(n);
    if (!st) return "";
    if (st === "overdue") return "Overdue " + Math.round((startOfToday() - dueTs(n.due)) / DAY) + "d";
    if (st === "today") return "Due today";
    return "Due " + fmtShort(dueTs(n.due));
  }

  // ===================================================================
  // Per-email storage (compact + compressed, under the 2,500-char limit)
  // ===================================================================
  const t36 = (ms) => Math.round((ms || 0) / 1000).toString(36);
  const f36 = (s) => parseInt(s, 36) * 1000;
  function toCompact(n) {
    const o = { i: n.id, t: n.text, c: statusIdx(n.color), cr: t36(n.created) };
    if (n.due) o.d = n.due;
    if (n.tags && n.tags.length) o.g = n.tags;
    if (n.att) o.a = n.att;
    if (n.prev && isDone(n)) o.p = statusIdx(n.prev);
    if (t36(n.updated) !== o.cr) o.u = t36(n.updated);
    if (n.doneAt) o.dn = t36(n.doneAt);
    if (n.since) o.w = t36(n.since);
    if (n.hist && n.hist.length) o.h = n.hist.map((x) => [x.t, t36(x.u)]);
    return o;
  }
  function fromCompact(o) {
    const cr = f36(o.cr || "0");
    return {
      id: o.i || newId(), text: o.t || "", color: (STATUS[o.c] || STATUS[0]).hex,
      due: o.d || "", tags: o.g || [], att: o.a || "",
      prev: o.p != null && STATUS[o.p] ? STATUS[o.p].hex : undefined,
      created: cr, updated: o.u ? f36(o.u) : cr, doneAt: o.dn ? f36(o.dn) : undefined, since: o.w ? f36(o.w) : undefined,
      hist: (o.h || []).map((x) => ({ t: x[0], u: f36(x[1]) })),
    };
  }
  // Size as Outlook may count it: UTF-8 bytes (never less than the character count)
  const cpSize = (v) => byteLen(JSON.stringify({ [PROP_KEY]: v }));
  function encodeNotes(list) {
    const json = JSON.stringify(list.map(toCompact));
    const options = ["v2" + json, "b2" + LZString.compressToBase64(json), "z2" + LZString.compressToUTF16(json)];
    return options.reduce((best, x) => (cpSize(x) < cpSize(best) ? x : best));
  }
  function decodeNotes(raw) {
    if (!raw || typeof raw !== "string") return [];
    try {
      if (raw[0] === "[") { // versions 1.x
        return JSON.parse(raw).map((n) => Object.assign({ tags: [], due: "", hist: [] }, n,
          { color: statusOf(n.color).hex, created: n.created || n.updated || Date.now(), updated: n.updated || n.created || Date.now() }));
      }
      const body = raw.slice(2);
      const json = raw.startsWith("z2") ? LZString.decompressFromUTF16(body)
        : raw.startsWith("b2") ? LZString.decompressFromBase64(body) : body;
      return JSON.parse(json || "[]").map(fromCompact);
    } catch (e) { return []; }
  }
  const lblValue = () => JSON.stringify({ t: emailTags, u: lblAt });
  const labelsCost = () => (emailTags.length ? byteLen(JSON.stringify({ [LBL_KEY]: lblValue() })) : 0);
  // Encodes the notes; drops old edit history if needed. null = does not fit.
  function fitNotes(list) {
    const limit = CP_LIMIT - labelsCost();
    let enc = encodeNotes(list);
    if (cpSize(enc) <= limit) return enc;
    const copy = list.map((n) => Object.assign({}, n, { hist: (n.hist || []).slice() }));
    for (;;) {
      const withHist = copy.filter((n) => n.hist.length);
      if (!withHist.length) break;
      withHist.forEach((n) => n.hist.pop());
      enc = encodeNotes(copy);
      if (cpSize(enc) <= limit) {
        list.forEach((n, i) => { n.hist = copy[i].hist; });
        return enc;
      }
    }
    return null;
  }
  const spaceUsed = () => ((notes.length ? cpSize(encodeNotes(notes)) : 0) + labelsCost()) / CP_LIMIT;

  // ===================================================================
  // Roaming storage (overview, labels, templates, settings)
  // ===================================================================
  function rsGet(k) { try { return RS().get(k); } catch (e) { return undefined; } }
  let saveTimer = null;
  function rsSet(k, v) {
    RS().set(k, v);
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { try { RS().saveAsync(); } catch (e) { /* ignore */ } }, 0);
  }
  function rsRemove(k) { try { RS().remove(k); } catch (e) { /* ignore */ } }

  let idxCache = null, idxBytes = 0;
  function getIndex() {
    if (idxCache) return idxCache;
    const raw = rsGet(IDX_KEY);
    let idx = {};
    if (typeof raw === "string" && raw[0] === "z") {
      try { idx = JSON.parse(LZString.decompressFromUTF16(raw.slice(1)) || "{}"); } catch (e) { idx = {}; }
      idxBytes = byteLen(raw);
    } else {
      const old = rsGet(OLD_IDX_KEY);
      if (old && typeof old === "object") { idx = old; idxCache = idx; setIndex(idx); } // migrate 1.x
    }
    let stale = false;
    Object.keys(idx).forEach((k) => { if (idx[k].draft && Date.now() - (idx[k].u || 0) > 21 * DAY) { delete idx[k]; stale = true; } });
    idxCache = idx;
    if (stale) setIndex(idx);
    return idx;
  }
  const compressIdx = (idx) => "z" + LZString.compressToUTF16(JSON.stringify(idx));
  // Keeps the overview under its budget: removes done notes (oldest first), then the oldest emails.
  function trimIndex(idx) {
    let z = compressIdx(idx);
    if (byteLen(z) <= IDX_BUDGET) return z;
    const done = [];
    Object.keys(idx).forEach((k) => (idx[k].n || []).forEach((n) => {
      if (statusOf(n.color).hex === DONE_HEX) done.push({ k, id: n.id, t: n.doneAt || n.updated || 0 });
    }));
    done.sort((a, b) => a.t - b.t);
    let i = 0;
    while (i < done.length) {
      for (let j = 0; j < 5 && i < done.length; j++, i++) {
        const d = done[i], v = idx[d.k];
        if (!v) continue;
        v.n = (v.n || []).filter((n) => n.id !== d.id);
        v.c = v.n.length;
        if (!v.n.length && !(v.lb && v.lb.length)) delete idx[d.k];
      }
      z = compressIdx(idx);
      if (byteLen(z) <= IDX_BUDGET) return z;
    }
    const keys = Object.keys(idx).sort((a, b) => (idx[a].u || 0) - (idx[b].u || 0));
    while (keys.length && byteLen(z) > IDX_BUDGET) {
      for (let j = 0; j < 3 && keys.length; j++) delete idx[keys.shift()];
      z = compressIdx(idx);
    }
    return z;
  }
  function setIndex(idx) {
    const z = trimIndex(idx);
    rsSet(IDX_KEY, z);
    if (rsGet(OLD_IDX_KEY) !== undefined && rsGet(OLD_IDX_KEY) !== null) rsRemove(OLD_IDX_KEY);
    idxCache = idx;
    idxBytes = byteLen(z);
  }
  function indexUsage() {
    getIndex();
    return Math.min(1, idxBytes / IDX_BUDGET);
  }

  // Normalised notes of an overview entry (also older formats)
  function entryNotes(v, st) {
    if (!v.n) return [{ id: undefined, text: v.p || "", color: NOTE_HEX, updated: v.u, created: v.u, due: "", tags: [], att: "" }];
    st = st || getTagStore();
    return v.n.map((n) => ({
      id: n.id, text: n.text != null ? n.text : n.t || "",
      color: statusOf(n.color || n.c || NOTE_HEX).hex,
      updated: n.updated || n.d || v.u, created: n.created || n.updated || n.d || v.u,
      doneAt: n.doneAt, since: n.since, due: n.due || "", tags: normTags(n.tags || [], st), att: n.att || "",
    }));
  }
  function allEntries() {
    const idx = getIndex(), st = getTagStore();
    return Object.keys(idx).map((k) => {
      const labels = normTags(idx[k].lb || [], st);
      const notesN = entryNotes(idx[k], st);
      return { key: k, id: idx[k].id || k, v: idx[k], notes: notesN, labels, tags: unionTags(labels, notesN.flatMap((n) => n.tags)) };
    });
  }

  function currentKey() {
    if (!ctx) return null;
    if (ctx.compose) return composeId ? "draft:" + composeId : null;
    return ctx.item.internetMessageId || ctx.item.itemId;
  }
  function updateIndex() {
    if (!ctx) return;
    const id = ctx.compose ? composeId : ctx.item.itemId;
    const key = currentKey();
    if (!id || !key) return;
    const idx = getIndex();
    const before = JSON.stringify(idx);
    const mine = new Set(notes.map((n) => n.id).filter(Boolean));
    // remove older entries of the same email (old key, moved email, or the draft it was sent from)
    Object.keys(idx).forEach((k) => {
      if (k === key) return;
      const v = idx[k];
      if (k === id || v.id === id || (v.n || []).some((n) => n.id && mine.has(n.id))) delete idx[k];
    });
    if (notes.length || emailTags.length) {
      const last = notes.slice().sort((a, b) => b.updated - a.updated)[0];
      const e = {
        id, s: (subject || (ctx.meeting ? "(no title)" : "(no subject)")).slice(0, 100),
        p: last ? last.text.replace(/\s+/g, " ").slice(0, 80) : "", c: notes.length,
        u: Math.max.apply(null, notes.map((n) => n.updated || 0).concat([lblAt || 0])),
        n: notes.map((n) => {
          const o = { id: n.id, text: n.text.slice(0, NOTE_TEXT_MAX), color: statusOf(n.color).hex,
            updated: n.updated, created: n.created };
          if (n.due) o.due = n.due;
          if (n.tags && n.tags.length) o.tags = n.tags;
          if (n.att) o.att = n.att;
          if (n.doneAt) o.doneAt = n.doneAt;
          if (n.since) o.since = n.since;
          return o;
        }),
      };
      if (emailTags.length) e.lb = emailTags.slice();
      if (ctx.compose) e.draft = 1;
      if (ctx.meeting) e.mt = 1;
      if (ctx.conv) e.cv = ctx.conv;
      idx[key] = e;
    } else {
      delete idx[key];
    }
    if (JSON.stringify(idx) === before) return;
    setIndex(idx);
  }

  // ===================================================================
  // Current item: load, save
  // ===================================================================
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;
    wireUi();
    buildStatusPicker();
    buildSummaryUi();
    getIndex(); // migrates the 1.x overview if needed
    loadItem();
    try { MB().addHandlerAsync(Office.EventType.ItemChanged, loadItem); } catch (e) { /* not in compose */ }
  });

  function loadItem() {
    const item = MB().item;
    editingId = null; composeId = null; showDone = false; replyCands = []; backupFound = null;
    openHist.clear();
    if (!item) {
      ctx = null; notes = []; emailTags = []; props = null; subject = "";
      resetComposer(); renderAll();
      return;
    }
    const meeting = item.itemType === "appointment" ||
      (Office.MailboxEnums && Office.MailboxEnums.ItemType && item.itemType === Office.MailboxEnums.ItemType.Appointment);
    ctx = {
      item, meeting,
      compose: !!(item.subject && typeof item.subject.getAsync === "function"),
      conv: item.conversationId || "",
      attachments: Array.isArray(item.attachments)
        ? item.attachments.filter((a) => !a.isInline && a.name).map((a) => a.name) : [],
    };
    subject = ctx.compose ? "" : item.subject || "";
    person = null;
    if (ctx.compose) {
      item.subject.getAsync((r) => { subject = ok(r) ? r.value || "" : ""; renderHeader(); renderTagComposer(); });
      try { item.to.getAsync((r) => { if (ok(r) && r.value && r.value.length) { person = personOf(r.value[0]); renderTagComposer(); } }); } catch (e) { /* not a message */ }
    } else if (!meeting) {
      const me = lc((MB().userProfile || {}).emailAddress || "");
      const from = item.from || item.sender;
      if (from && lc(from.emailAddress || "") !== me) person = personOf(from);
    }
    item.loadCustomPropertiesAsync((res) => {
      if (!ok(res)) { toast("Couldn't load notes for this item"); return; }
      props = res.value;
      notes = decodeNotes(props.get(PROP_KEY));
      draftKey = ctx.compose ? props.get(DKEY_PROP) || null : null;
      emailTags = []; lblAt = 0;
      try {
        const lb = JSON.parse(props.get(LBL_KEY) || "null");
        if (lb && Array.isArray(lb.t)) { emailTags = lb.t.filter((x) => typeof x === "string"); lblAt = lb.u || 0; }
      } catch (e) { /* ignore */ }
      const changed = applyPending() | normaliseNoteTags();
      resetComposer();
      renderAll();
      if (changed && !ctx.compose) persist(null);
      else {
        syncCategory();
        if (!ctx.compose) { updateIndex(); updateBanner(); }
      }
      detectReply();
      detectBackup();
      refreshBadges();
      if (!ctx.compose && !ctx.meeting && !notes.length && !emailTags.length) recoverDraftNotes();
    });
  }

  // ----- notes written while composing: kept aside until the sent copy is opened -----
  function getDrafts() { const d = rsGet(DRAFTS_KEY); return d && typeof d === "object" ? d : {}; }
  function saveDraftStore() {
    if (!ctx || !ctx.compose || ctx.meeting || !draftKey) return;
    const d = getDrafts();
    Object.keys(d).forEach((k) => { if (Date.now() - (d[k].at || 0) > 21 * DAY) delete d[k]; });
    if (notes.length || emailTags.length) d[draftKey] = { s: subject, at: Date.now(), n: notes.map(toCompact), lb: emailTags.slice() };
    else delete d[draftKey];
    rsSet(DRAFTS_KEY, d);
  }
  function stampDraft() {
    const ih = ctx.item.internetHeaders;
    if (!ih || !ih.setAsync) return;
    try { ih.setAsync({ [HDR]: draftKey }, () => {}); } catch (e) { /* not supported here */ }
  }
  const normSubject = (x) => lc(String(x || "")).replace(/^(\s*(re|fw|fwd|aw|wg|השב|הועבר)\s*:\s*)+/i, "").replace(/\s+/g, " ").trim();
  function recoverDraftNotes() {
    const d = getDrafts();
    const keys = Object.keys(d);
    if (!keys.length) return;
    const attach = (key, how) => {
      const e = d[key];
      notes = (e.n || []).map(fromCompact);
      emailTags = (e.lb || []).slice();
      lblAt = Date.now();
      delete d[key];
      rsSet(DRAFTS_KEY, d);
      renderAll();
      persist(how === "header" ? "Notes from your draft were attached" : "Notes from a draft with the same subject were attached");
    };
    const bySubject = () => {
      const created = ctx.item.dateTimeCreated ? new Date(ctx.item.dateTimeCreated).getTime() : Date.now();
      const me = lc((MB().userProfile || {}).emailAddress || "");
      const from = ctx.item.from || ctx.item.sender;
      if (!from || lc(from.emailAddress || "") !== me) return;
      const k = keys.find((x) => normSubject(d[x].s) === normSubject(subject) && Math.abs(created - (d[x].at || 0)) < 3 * DAY);
      if (k) attach(k, "subject");
    };
    const item = ctx.item;
    if (item.getAllInternetHeadersAsync) {
      try {
        item.getAllInternetHeadersAsync((r) => {
          const m = ok(r) && String(r.value || "").match(new RegExp("^" + HDR + ":\\s*(\\S+)", "im"));
          if (m && d[m[1]]) attach(m[1], "header"); else bySubject();
        });
        return;
      } catch (e) { /* fall through */ }
    }
    bySubject();
  }
  function personOf(x) {
    let dn = String((x && x.displayName) || "").replace(/\(.*?\)/g, "").replace(/<.*?>/g, "").trim();
    if (!dn && x && x.emailAddress) dn = String(x.emailAddress).split("@")[0].replace(/[._-]+/g, " ").replace(/\d+/g, "").trim().toUpperCase();
    if (dn.includes(",")) { const p = dn.split(","); dn = (p[1] + " " + p[0]).replace(/\s+/g, " ").trim(); }
    if (dn && dn === dn.toUpperCase() && /[A-Z]/.test(dn)) dn = dn.toLowerCase().replace(/(^|[\s-])\S/g, (c) => c.toUpperCase());
    dn = dn.slice(0, TAG_MAX).trim();
    return dn ? { name: dn, first: dn.split(" ")[0] } : null;
  }

  // Saves the notes; on "no space" the change is rolled back.
  function persist(msg, rollback) {
    if (!props) return false;
    const enc = notes.length ? fitNotes(notes) : "";
    if (enc === null) {
      if (rollback) rollback();
      renderNotes();
      dialog({ title: "No space left on this email", alert: true, ok: "OK",
        text: "Outlook keeps about 2,500 characters of notes per email. Shorten a note or delete done notes, then try again." });
      return false;
    }
    if (notes.length) props.set(PROP_KEY, enc); else props.remove(PROP_KEY);
    if (emailTags.length) props.set(LBL_KEY, lblValue()); else props.remove(LBL_KEY);
    const stamp = ctx.compose && !ctx.meeting && (notes.length || emailTags.length) && !draftKey;
    if (stamp) { draftKey = newId(); props.set(DKEY_PROP, draftKey); }
    props.saveAsync((res) => {
      if (!ok(res)) { toast("Couldn't save: " + ((res.error && res.error.message) || "unknown error")); return; }
      syncCategory();
      const after = () => { updateIndex(); refreshBadges(); renderThread(); if (tab === "today") renderToday(); };
      if (ctx.compose) { if (stamp) stampDraft(); saveDraftStore(); saveDraft(after); }
      else { after(); updateBanner(); }
      if (msg) toast(msg);
    });
    renderHeader();
    return true;
  }
  function mutate(fn, msg) {
    const snap = JSON.stringify(notes);
    fn();
    const done = persist(msg, () => { notes = JSON.parse(snap); });
    renderNotes();
    return done;
  }
  function saveDraft(done) {
    const item = ctx.item;
    const finish = () => {
      if (!item.saveAsync) return done();
      item.saveAsync((r) => { if (ok(r) && r.value) composeId = r.value; done(); });
    };
    if (item.subject && item.subject.getAsync) item.subject.getAsync((r) => { if (ok(r)) subject = r.value || ""; finish(); });
    else finish();
  }

  // Changes made from another email (e.g. "mark done" after a reply arrived)
  function applyPending() {
    const pend = rsGet(PEND_KEY) || {};
    let notesChanged = false, pendChanged = false;
    notes.forEach((n) => {
      const p = pend[n.id];
      if (!p) return;
      if (p.due !== undefined && !isDone(n) && n.due !== p.due) { n.due = p.due; n.updated = p.at; notesChanged = true; }
      if (p.done && !isDone(n)) {
        n.prev = statusOf(n.color).hex; n.color = DONE_HEX; n.doneAt = p.at; n.updated = p.at; notesChanged = true;
      }
      delete pend[n.id];
      pendChanged = true;
    });
    Object.keys(pend).forEach((k) => { if (Date.now() - (pend[k].at || 0) > 90 * DAY) { delete pend[k]; pendChanged = true; } });
    if (pendChanged) rsSet(PEND_KEY, pend);
    return notesChanged;
  }

  // ===================================================================
  // Outlook categories and info bar
  // ===================================================================
  function categoriesSupported() {
    return ctx && ctx.item.categories && Office.context.requirements.isSetSupported("Mailbox", "1.8");
  }
  function ensureMasterCategories(names, done) {
    const mc = MB().masterCategories;
    if (!mc || !names.length) return done();
    mc.getAsync((r) => {
      if (!ok(r)) return done();
      const existing = new Set((r.value || []).map((c) => c.displayName));
      const missing = names.filter((nm) => !existing.has(nm)).map((nm) => {
        const def = ALL_CATS.find((l) => l.name === nm);
        const preset = def ? def.preset : tagPalette(nm.slice(TAG_PREFIX.length))[0];
        return { displayName: nm, color: Office.MailboxEnums.CategoryColor[preset] };
      });
      if (!missing.length) return done();
      mc.addAsync(missing, () => done());
    });
  }
  // Open notes mark the email; "Done" only when every note is done; labels optional
  function syncCategory() {
    if (!categoriesSupported()) return;
    const item = ctx.item;
    const open = notes.filter((n) => !isDone(n));
    const wanted = open.length ? [...new Set(open.map((n) => statusOf(n.color).name))]
      : notes.length ? [statusOf(DONE_HEX).name] : [];
    if (open.some((n) => n.due)) wanted.push(FOLLOW.name);
    if (getTagStore().cats) unionTags(emailTags, notes.flatMap((n) => n.tags || [])).forEach((t) => wanted.push(TAG_PREFIX + t));
    const ours = new Set(ALL_CATS.map((l) => l.name).concat(OLD_NAMES));
    const isOurs = (n) => ours.has(n) || n.startsWith(TAG_PREFIX);
    item.categories.getAsync((r) => {
      if (!ok(r)) return;
      const current = (r.value || []).map((c) => c.displayName).filter(isOurs);
      const toAdd = wanted.filter((n) => !current.includes(n));
      const toRemove = current.filter((n) => !wanted.includes(n));
      if (toRemove.length) item.categories.removeAsync(toRemove, () => {});
      if (toAdd.length) ensureMasterCategories(toAdd, () => item.categories.addAsync(toAdd, () => {}));
    });
  }
  function updateBanner() {
    const item = ctx && ctx.item;
    if (!item || !item.notificationMessages) return;
    const late = notes.filter(isLate);
    try {
      if (!late.length) { item.notificationMessages.removeAsync(BANNER_KEY, () => {}); return; }
      const over = late.filter((n) => dueState(n) === "overdue").length;
      const parts = [];
      if (over) parts.push(over + " overdue");
      if (late.length > over) parts.push((late.length - over) + " due today");
      let msg = "My Notes: " + parts.join(", ") + " – " + late[0].text.replace(/\s+/g, " ");
      if (msg.length > 150) msg = msg.slice(0, 147) + "…";
      item.notificationMessages.replaceAsync(BANNER_KEY, {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: msg, icon: "Icon.16x16", persistent: false,
      }, () => {});
    } catch (e) { /* not supported here */ }
  }

  // ===================================================================
  // Notes: create, edit, done, delete, history
  // ===================================================================
  function saveNote() {
    const text = $("input").value.trim();
    if (!text) return;
    addTagFromInput();
    const now = Date.now();
    const due = statusOf(color).hex === INFO_HEX ? "" : $("due").value || "";
    const tags = draftTags.slice();
    const att = $("attSel") && !$("attRow").hidden ? $("attSel").value : "";
    registerTags(tags);
    const wasEditing = editingId;
    const saved = mutate(() => {
      if (editingId) {
        const n = notes.find((x) => x.id === editingId);
        if (!n) return;
        if (n.text !== text) n.hist = [{ t: n.text, u: n.updated }].concat(n.hist || []).slice(0, 3);
        const wasDone = isDone(n);
        n.text = text; n.due = due; n.tags = tags; n.att = att; n.updated = now;
        if (statusOf(color).hex !== statusOf(n.color).hex) {
          if (statusOf(color).hex === DONE_HEX && !wasDone) { n.prev = statusOf(n.color).hex; n.doneAt = now; }
          if (statusOf(color).hex !== DONE_HEX) { delete n.prev; delete n.doneAt; }
          n.since = statusOf(color).hex === WAIT_HEX ? now : undefined;
          n.color = color;
        }
      } else {
        const n = { id: newId(), text, color, due, tags, att, created: now, updated: now, hist: [] };
        if (statusOf(color).hex === DONE_HEX) n.doneAt = now;
        if (statusOf(color).hex === WAIT_HEX) n.since = now;
        notes.push(n);
      }
    }, wasEditing ? "Note updated" : "Note added");
    if (saved) { resetComposer(); if (document.activeElement && document.activeElement.blur) document.activeElement.blur(); }
  }
  function editNote(id) {
    const n = notes.find((x) => x.id === id);
    if (!n) return;
    editingId = id;
    $("input").value = n.text;
    $("due").value = n.due || "";
    draftTags = (n.tags || []).slice();
    setColor(n.color);
    renderTagComposer();
    renderAttachments(n.att);
    $("save").textContent = "Save changes";
    $("save").disabled = false;
    $("cancel").hidden = false;
    $("composer").classList.add("editing");
    $("input").focus();
    updateComposerState();
  }
  function toggleDone(id) {
    mutate(() => {
      const n = notes.find((x) => x.id === id);
      if (!n) return;
      if (isDone(n)) { n.color = n.prev && n.prev !== DONE_HEX ? n.prev : NOTE_HEX; delete n.prev; delete n.doneAt; if (isWait(n)) n.since = Date.now(); }
      else { n.prev = statusOf(n.color).hex; n.color = DONE_HEX; n.doneAt = Date.now(); }
      n.updated = Date.now();
    }, notes.find((x) => x.id === id && isDone(x)) ? "Note reopened" : "Marked as done");
  }
  async function deleteNote(id) {
    const yes = await confirmDialog("Delete this note?", "The note will be deleted permanently. This can't be undone.", "Delete");
    if (!yes) return;
    if (editingId === id) resetComposer();
    mutate(() => { notes = notes.filter((x) => x.id !== id); }, "Note deleted");
  }
  async function deleteDone() {
    const done = notes.filter(isDone);
    if (!done.length) return;
    const all = done.length === notes.length;
    const many = (k) => k + " note" + (k > 1 ? "s" : "");
    const yes = all
      ? await confirmDialog("Delete all notes?", "Are you sure you want to delete all " + many(notes.length) +
        " on this email? This can't be undone.", "Delete all")
      : await confirmDialog("Delete done notes?", "Are you sure you want to delete " + many(done.length) +
        " marked as done? Open notes are kept. This can't be undone.", "Delete");
    if (!yes) return;
    showDone = false;
    mutate(() => { notes = notes.filter((n) => !isDone(n)); }, all ? "All notes deleted" : "Done notes deleted");
  }
  function resetComposer() {
    editingId = null;
    $("input").value = "";
    $("due").value = "";
    draftTags = defaultTags();
    setColor(NOTE_HEX);
    renderTagComposer();
    renderAttachments("");
    $("save").textContent = "Add note";
    $("cancel").hidden = true;
    $("composer").classList.remove("editing");
    updateComposerState();
  }
  // Inserts "05/11/2026 - " at the cursor (Alt+D)
  function insertToday() {
    const ta = $("input");
    const d = new Date();
    let ins = pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear() + " - ";
    const start = ta.selectionStart != null ? ta.selectionStart : ta.value.length;
    const end = ta.selectionEnd != null ? ta.selectionEnd : start;
    if (start > 0 && !/\s$/.test(ta.value.slice(0, start))) ins = " " + ins;
    ta.setRangeText(ins, start, end, "end");
    ta.focus();
    updateComposerState();
  }
  function updateComposerState() {
    const has = !!$("input").value.trim();
    $("save").disabled = !has || !ctx;
    $("tplSave").hidden = !has;
    $("composer").classList.toggle("open", has || !!editingId);
    document.querySelectorAll("#dueChips button").forEach((b) => b.setAttribute("aria-pressed", "false"));
  }

  // ----- status picker -----
  function buildStatusPicker() {
    const wrap = $("statusPick");
    STATUS.forEach((s) => {
      wrap.appendChild(h("button", { class: "st", "data-hex": s.hex, title: s.name.replace(/^\S+\s/, "") + " – " + s.tip + " (Alt+" + (STATUS.indexOf(s) + 1) + ")",
        style: "--c:" + s.hex, "aria-pressed": "false", onclick: () => setColor(s.hex) }, h("i", { style: "background:" + s.hex }), h("span", { text: s.short })));
    });
  }
  function setColor(c) {
    color = statusOf(c).hex;
    document.querySelectorAll("#statusPick .st").forEach((b) => b.setAttribute("aria-pressed", b.dataset.hex === color ? "true" : "false"));
    $("dueRow").classList.toggle("off", color === INFO_HEX);
    $("dueRow").title = color === INFO_HEX ? "Info notes have no due date" : "";
  }

  // ----- due date shortcuts -----
  function setDue(kind, btn) {
    let d;
    if (kind === "today") d = new Date(startOfToday());
    else if (kind === "sun") d = nextSunday();
    else d = addWorkdays(startOfToday(), parseInt(kind, 10));
    $("due").value = iso(d);
    document.querySelectorAll("#dueChips button").forEach((b) => b.setAttribute("aria-pressed", b === btn ? "true" : "false"));
  }

  // ----- attachments -----
  function renderAttachments(selected) {
    const list = ctx ? ctx.attachments : [];
    $("attRow").hidden = !list.length;
    const sel = $("attSel");
    sel.innerHTML = "";
    sel.appendChild(h("option", { value: "", text: "Not about a specific file" }));
    list.forEach((nm) => sel.appendChild(h("option", { value: nm, text: nm })));
    if (selected && !list.includes(selected)) sel.appendChild(h("option", { value: selected, text: selected }));
    sel.value = selected || "";
  }

  // ===================================================================
  // Quick notes (templates)
  // ===================================================================
  function getTemplates() { const t = rsGet(TPL_KEY); return Array.isArray(t) ? t : DEFAULT_TEMPLATES.map((x) => Object.assign({}, x)); }
  function setTemplates(t) { rsSet(TPL_KEY, t); }
  const fillFrom = (text) => String(text).replace(/\{from\}/gi, person ? person.first : "…");
  function applyTemplate(t) {
    if (editingId) resetComposer();
    $("input").value = fillFrom(t.text);
    setColor(t.color);
    $("due").value = t.days != null ? iso(addWorkdays(startOfToday(), t.days)) : "";
    (t.tags || []).forEach((x) => { if (!draftTags.some((y) => lc(y) === lc(x))) draftTags.push(x); });
    renderTagComposer();
    updateComposerState();
    $("input").focus();
    $("input").setSelectionRange(t.text.length, t.text.length);
  }
  function saveAsTemplate() {
    const text = $("input").value.trim();
    if (!text) return;
    addTagFromInput();
    const due = $("due").value;
    const t = { id: newId(), text: text.slice(0, 120), color, tags: draftTags.slice(),
      days: due ? workdaysBetween(startOfToday(), dueTs(due)) : null };
    const list = getTemplates().filter((x) => lc(x.text) !== lc(t.text));
    list.push(t);
    setTemplates(list.slice(-12));
    renderQuick();
    toast("Saved as quick note");
  }
  function renderQuick() {
    const list = getTemplates();
    const row = $("quickRow");
    row.innerHTML = "";
    list.forEach((t) => {
      const st = statusOf(t.color);
      row.appendChild(h("button", { class: "quick", style: "--c:" + st.hex, title: t.text + (t.days != null ? " (due in " + t.days + " workdays)" : ""),
        onclick: () => applyTemplate(t) }, h("bdi", { text: fillFrom(t.text) })));
    });
    $("quickWrap").hidden = !ctx || !list.length;
    requestAnimationFrame(() => row.classList.toggle("fade", row.scrollWidth > row.clientWidth + 2));
    const panel = $("tplList");
    panel.innerHTML = "";
    if (!list.length) panel.appendChild(h("li", { class: "muted small", text: "No quick notes yet." }));
    list.forEach((t) => {
      const st = statusOf(t.color);
      panel.appendChild(h("li", { class: "tpl" },
        h("i", { class: "dot", style: "background:" + st.hex }),
        h("bdi", { class: "tpl-text", text: t.text }),
        h("span", { class: "muted small", text: t.days != null ? "+" + t.days + " wd" : "" }),
        iconBtn("trash", "Delete quick note", () => { setTemplates(getTemplates().filter((x) => x.id !== t.id)); renderQuick(); })));
    });
  }

  // ===================================================================
  // Thread: notes on other emails of the conversation, "reply received"
  // ===================================================================
  function threadEntries() {
    if (!ctx || !ctx.conv) return [];
    const key = currentKey();
    return allEntries().filter((e) => e.notes.length && e.v.cv === ctx.conv && e.key !== key && e.id !== (ctx.item.itemId || ""));
  }
  function detectReply() {
    replyCands = [];
    if (!ctx || ctx.compose || ctx.meeting || !ctx.conv) { renderReply(); return; }
    const me = lc((MB().userProfile || {}).emailAddress || "");
    const from = ctx.item.from || ctx.item.sender;
    if (from && me && lc(from.emailAddress || "") === me) { renderReply(); return; }
    const created = ctx.item.dateTimeCreated ? new Date(ctx.item.dateTimeCreated).getTime() : Date.now();
    const dism = rsGet(DISM_KEY) || [];
    threadEntries().forEach((e) => e.notes.forEach((n) => {
      if (statusOf(n.color).hex === WAIT_HEX && n.id && (n.updated || 0) < created && !dism.includes(n.id)) {
        replyCands.push({ e, n });
      }
    }));
    renderReply();
  }
  function markRemoteDone(c) {
    replyCands = replyCands.filter((x) => x !== c);
    remoteUpdate(c.e, c.n.id, { done: 1 }, "Marked as done");
  }
  // Change a note on any email. The open email is changed directly; others get the change
  // in the overview now and on the email itself the next time it is opened.
  function remoteUpdate(e, noteId, patch, msg) {
    const now = Date.now();
    const apply = (n) => {
      if (patch.due !== undefined) n.due = patch.due;
      if (patch.done) { n.prev = statusOf(n.color).hex; n.color = DONE_HEX; n.doneAt = now; }
      n.updated = now;
    };
    if (ctx && e.key === currentKey()) {
      mutate(() => { const n = notes.find((x) => x.id === noteId); if (n) apply(n); }, msg);
    } else {
      const idx = getIndex();
      const v = idx[e.key];
      if (v && v.n) { v.n.forEach((n) => { if (n.id === noteId) apply(n); }); v.u = now; setIndex(idx); }
      const pend = rsGet(PEND_KEY) || {};
      pend[noteId] = Object.assign({}, pend[noteId] || {}, patch, { at: now });
      rsSet(PEND_KEY, pend);
      toast(msg);
    }
    renderReply(); renderThread(); refreshBadges();
    if (tab === "today") renderToday();
    if (tab === "all") renderIndex();
    if (tab === "tag") renderTags();
  }

  // ----- snooze: a small menu of new due dates -----
  function snoozeMenu(anchor, current, onPick) {
    const t = startOfToday();
    const items = [
      ["Today", iso(new Date(t))], ["Next workday", iso(addWorkdays(t, 1))], ["In 3 workdays", iso(addWorkdays(t, 3))],
      ["In 5 workdays", iso(addWorkdays(t, 5))], ["Next Sunday", iso(nextSunday())],
    ];
    popover(anchor, items.map(([label, v]) => ({ label, hint: fmtShort(dueTs(v)), active: v === current, onClick: () => onPick(v) }))
      .concat(current ? [{ label: "Remove due date", danger: true, onClick: () => onPick("") }] : []));
  }
  let popEl = null;
  function closePopover() { if (popEl) { popEl.remove(); popEl = null; document.removeEventListener("mousedown", onPopDown, true); } }
  function onPopDown(e) { if (popEl && !popEl.contains(e.target)) closePopover(); }
  function popover(anchor, items) {
    closePopover();
    const pop = h("div", { class: "pop", role: "menu" });
    items.forEach((it) => pop.appendChild(h("button", { class: "pop-item" + (it.active ? " active" : "") + (it.danger ? " danger" : ""), role: "menuitem",
      onclick: () => { closePopover(); it.onClick(); } }, h("span", { text: it.label }), it.hint ? h("span", { class: "pop-hint", text: it.hint }) : null)));
    document.body.appendChild(pop);
    const r = anchor.getBoundingClientRect();
    const w = pop.offsetWidth, hgt = pop.offsetHeight;
    let left = Math.min(Math.max(8, r.left), window.innerWidth - w - 8);
    let top = r.bottom + 4;
    if (top + hgt > window.innerHeight - 8) top = Math.max(8, r.top - hgt - 4);
    pop.style.left = left + window.scrollX + "px";
    pop.style.top = top + window.scrollY + "px";
    popEl = pop;
    setTimeout(() => document.addEventListener("mousedown", onPopDown, true), 0);
    const first = pop.querySelector("button");
    if (first) first.focus();
  }
  function dismissReply() {
    const dism = (rsGet(DISM_KEY) || []).concat(replyCands.map((c) => c.n.id)).slice(-200);
    rsSet(DISM_KEY, dism);
    replyCands = [];
    renderReply();
  }
  function openEmail(id) {
    try { MB().displayMessageForm(id); }
    catch (e) { toast("Couldn't open it. The email may have been moved or deleted."); }
  }

  // ===================================================================
  // Calendar reminder, reply/insert with notes
  // ===================================================================
  function addToCalendar(n) {
    if (!n.due) return;
    const start = new Date(dueTs(n.due)); start.setHours(9, 0, 0, 0);
    const end = new Date(start.getTime() + 15 * 60000);
    const link = ctx && !ctx.compose ? messageLink(ctx.item.itemId) : "";
    try {
      MB().displayNewAppointmentForm({
        requiredAttendees: [], optionalAttendees: [], location: "", resources: [],
        start, end, subject: "⏰ " + n.text.replace(/\s+/g, " ").slice(0, 90),
        body: n.text + "\n\n" + (ctx.meeting ? "Meeting: " : "Email: ") + subject + (link ? "\n" + link : ""),
      });
    } catch (e) { toast("Calendar isn't available here. Try Outlook on the desktop or the web."); }
  }
  function notesHtml(list) {
    return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px">' +
      list.map((n) => '<p dir="' + dirOf(n.text) + '" style="margin:0 0 6px">• ' + esc(n.text).replace(/\n/g, "<br>") + "</p>").join("") +
      "</div><p></p>";
  }
  function replyWithNotes(all) {
    const open = notes.filter((n) => !isDone(n));
    if (!open.length) return toast("There are no open notes to add");
    const opts = { htmlBody: notesHtml(open) };
    try { all ? ctx.item.displayReplyAllForm(opts) : ctx.item.displayReplyForm(opts); }
    catch (e) { toast("Couldn't open a reply here"); }
  }
  function insertHtml(html, msg) {
    try {
      ctx.item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (r) => {
        toast(ok(r) ? msg : "Couldn't insert. Click inside the body first.");
      });
    } catch (e) { toast("Couldn't insert here"); }
  }

  // ===================================================================
  // Labels
  // ===================================================================
  function getTagStore() {
    const st = rsGet(TAGS_KEY);
    return Object.assign({ names: [], renamed: {}, removed: [], cats: true, stages: {}, byStage: false }, st || {});
  }
  const stageOf = (name, st) => { const k = Object.keys((st || getTagStore()).stages).find((x) => lc(x) === lc(name)); return k ? STAGES.find((g) => g.id === (st || getTagStore()).stages[k]) || null : null; };
  function setStage(name, id) {
    const st = getTagStore();
    Object.keys(st.stages).forEach((k) => { if (lc(k) === lc(name)) delete st.stages[k]; });
    if (id) st.stages[name] = id;
    saveTagStore(st);
  }
  function saveTagStore(st) { rsSet(TAGS_KEY, st); }
  function cleanTag(s) { return String(s || "").replace(/[,;]/g, " ").replace(/\s+/g, " ").trim().slice(0, TAG_MAX); }
  function resolveTag(name, st) {
    let cur = name;
    for (let i = 0; i < 10; i++) {
      const k = Object.keys(st.renamed).find((x) => lc(x) === lc(cur));
      if (!k) break;
      cur = st.renamed[k];
    }
    if (st.removed.some((r) => lc(r) === lc(cur))) return null;
    return st.names.find((x) => lc(x) === lc(cur)) || cur;
  }
  function normTags(arr, st) {
    const out = [];
    (arr || []).forEach((t) => { const r = resolveTag(t, st); if (r && !out.some((x) => lc(x) === lc(r))) out.push(r); });
    return out;
  }
  function unionTags(a, b) {
    const out = [];
    [].concat(a || [], b || []).forEach((t) => { if (!out.some((x) => lc(x) === lc(t))) out.push(t); });
    return out;
  }
  function normaliseNoteTags() {
    const st = getTagStore();
    let changed = false;
    if (emailTags.length) {
      const t = normTags(emailTags, st);
      if (JSON.stringify(t) !== JSON.stringify(emailTags)) { emailTags = t; changed = true; }
    }
    notes.forEach((n) => {
      if (!n.tags || !n.tags.length) return;
      const t = normTags(n.tags, st);
      if (JSON.stringify(t) !== JSON.stringify(n.tags)) { n.tags = t; changed = true; }
    });
    return changed;
  }
  function registerTags(tags) {
    if (!tags.length) return;
    const st = getTagStore();
    let changed = false;
    tags.forEach((t) => {
      if (!st.names.some((x) => lc(x) === lc(t))) { st.names.push(t); changed = true; }
      Object.keys(st.renamed).forEach((k) => { if (lc(k) === lc(t)) { delete st.renamed[k]; changed = true; } });
      const before = st.removed.length;
      st.removed = st.removed.filter((r) => lc(r) !== lc(t));
      if (st.removed.length !== before) changed = true;
    });
    if (changed) saveTagStore(st);
  }
  function allTagNames() {
    const st = getTagStore();
    const out = normTags(st.names, st);
    allEntries().forEach((e) => e.tags.forEach((t) => { if (!out.some((x) => lc(x) === lc(t))) out.push(t); }));
    notes.forEach((n) => normTags(n.tags || [], st).forEach((t) => { if (!out.some((x) => lc(x) === lc(t))) out.push(t); }));
    normTags(emailTags, st).forEach((t) => { if (!out.some((x) => lc(x) === lc(t))) out.push(t); });
    return out.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }
  function tagPalette(name) {
    let x = 0;
    for (const ch of lc(name)) x = (x * 31 + ch.codePointAt(0)) >>> 0;
    return TAG_PALETTE[x % TAG_PALETTE.length];
  }
  function tagChip(name, onClick, onRemove) {
    const el = h(onClick ? "button" : "span", { class: "tag-chip", style: "--t:" + tagPalette(name)[1],
      title: onClick ? "Show all notes with this label" : null, onclick: onClick || null },
      h("span", { class: "tg", html: svg("tag", 11) }), h("bdi", { text: name }));
    if (onRemove) {
      el.appendChild(h("button", { class: "tag-x", title: "Remove label", "aria-label": "Remove label " + name, text: "×",
        onclick: (e) => { e.stopPropagation(); onRemove(); } }));
    }
    return el;
  }
  function defaultTags() {
    const last = notes.slice().sort((a, b) => b.updated - a.updated)[0];
    return last && last.tags ? last.tags.slice() : [];
  }
  function addTag(raw) {
    const t0 = cleanTag(raw);
    if (!t0) return;
    const t = allTagNames().find((x) => lc(x) === lc(t0)) || t0;
    if (!draftTags.some((x) => lc(x) === lc(t))) draftTags.push(t);
    renderTagComposer();
  }
  function addTagFromInput() { const v = $("tagInput").value; $("tagInput").value = ""; if (v.trim()) addTag(v); }
  function renderTagComposer() {
    const wrap = $("tagChips");
    wrap.innerHTML = "";
    draftTags.forEach((t) => wrap.appendChild(tagChip(t, null, () => { draftTags = draftTags.filter((x) => x !== t); renderTagComposer(); })));
    const dl = $("tagOptions");
    dl.innerHTML = "";
    allTagNames().filter((t) => !draftTags.some((x) => lc(x) === lc(t))).forEach((t) => dl.appendChild(h("option", { value: t })));
    renderEmailTags();
  }

  // ----- labels on the email itself -----
  function renderEmailTags() {
    const row = $("elRow");
    row.hidden = !ctx;
    if (!ctx) return;
    const wrap = $("elChips");
    wrap.innerHTML = "";
    emailTags.forEach((t) => wrap.appendChild(tagChip(t, () => { switchTab("tag"); openTag(t); }, () => removeEmailTag(t))));
    $("elInput").placeholder = emailTags.length ? "Add label" : "Label this " + (ctx.meeting ? "meeting" : "email");
    const dl = $("elOptions");
    dl.innerHTML = "";
    allTagNames().filter((t) => !emailTags.some((x) => lc(x) === lc(t))).forEach((t) => dl.appendChild(h("option", { value: t })));
    // suggestions from document numbers in the subject and file names
    const sug = $("elSuggest");
    sug.innerHTML = "";
    const src = (subject || "") + " " + ctx.attachments.join(" ");
    const live = $("sumPrefixes") && $("sumPrefixes").value.trim();
    const refs = refsOf(src, parsePrefixes(live || (prefsCache || getPrefs()).prefixes)).filter((r) => !emailTags.some((x) => lc(x) === lc(r)));
    refs.slice(0, 4).forEach((r) => sug.appendChild(h("button", { class: "tag-sug", title: "Label this email " + r,
      onclick: () => addEmailTag(r) }, h("span", { html: svg("plus", 10) }), h("bdi", { text: r }))));
    if (person && !emailTags.some((x) => lc(x) === lc(person.name))) {
      sug.appendChild(h("button", { class: "tag-sug person", title: "Label this email with the " + (ctx.compose ? "recipient" : "sender") + "'s name, to see everything open with " + person.first,
        onclick: () => addEmailTag(person.name) }, h("span", { html: svg("user", 10) }), h("bdi", { text: person.name })));
    }
    sug.hidden = !sug.childElementCount;
  }
  function changeEmailTags(fn, msg) {
    if (!props) return;
    const snap = emailTags.slice(), snapAt = lblAt;
    fn();
    lblAt = Date.now();
    persist(msg, () => { emailTags = snap; lblAt = snapAt; });
    renderEmailTags();
    renderTagComposer();
  }
  function addEmailTag(raw) {
    const t0 = cleanTag(raw);
    if (!t0) return;
    const t = allTagNames().find((x) => lc(x) === lc(t0)) || t0;
    if (emailTags.some((x) => lc(x) === lc(t))) return;
    registerTags([t]);
    changeEmailTags(() => { emailTags.push(t); }, "Label added");
  }
  function removeEmailTag(t) { changeEmailTags(() => { emailTags = emailTags.filter((x) => x !== t); }, "Label removed"); }
  function addEmailTagFromInput() { const v = $("elInput").value; $("elInput").value = ""; if (v.trim()) addEmailTag(v); }

  // ----- Labels tab -----
  function tagStats() {
    const map = new Map();
    allTagNames().forEach((t) => map.set(lc(t), { name: t, notes: 0, open: 0, late: 0, mails: new Set(), u: 0 }));
    allEntries().forEach((e) => e.tags.forEach((t) => {
      const s = map.get(lc(t));
      if (!s) return;
      s.mails.add(e.key); s.u = Math.max(s.u, e.v.u || 0);
      notesWithTag(e, t).forEach((n) => {
        s.notes++;
        if (!isDone(n)) s.open++;
        if (isLate(n)) s.late++;
      });
    }));
    return [...map.values()];
  }
  function renderTags() {
    const detail = !!openTagName;
    $("tagListView").hidden = detail;
    $("tagDetail").hidden = !detail;
    if (detail) return renderTagDetail();
    const st = getTagStore();
    $("tagCats").checked = st.cats;
    $("tagByStage").checked = st.byStage;
    const q = lc($("tagSearch").value.trim());
    const list = $("tagList");
    list.innerHTML = "";
    const stats = tagStats().filter((s) => !q || lc(s.name).includes(q))
      .sort((a, b) => b.late - a.late || b.open - a.open || b.u - a.u || a.name.localeCompare(b.name));
    stats.forEach((s) => { s.stage = stageOf(s.name, st); });
    renderStageBar(tagStats().map((s) => stageOf(s.name, st)));
    if (!stats.length) {
      list.appendChild(emptyState(q ? "No labels match your search." : "No labels yet",
        q ? "" : "Labels link notes from different emails. Add one in the label field when you write a note, or create one here."));
      return;
    }
    const rowEl = (s) => h("li", null, h("button", { class: "row", onclick: () => openTag(s.name) },
      h("div", { class: "row-top" }, tagChip(s.name),
        s.late ? h("span", { class: "pill overdue", text: s.late + " late" }) : (!st.byStage && s.stage) ? h("span", { class: "stage-tag", style: "--c:" + s.stage.color, text: s.stage.name }) : null),
      h("div", { class: "row-sub" }, s.notes
        ? pl(s.notes, "note") + " on " + pl(s.mails.size, "email") + ", " + s.open + " open"
        : "No notes yet")));
    if (!st.byStage) { stats.forEach((s) => list.appendChild(rowEl(s))); return; }
    const groups = STAGES.map((g) => ({ g, items: stats.filter((s) => s.stage && s.stage.id === g.id) }))
      .concat([{ g: { id: "", name: "No stage", color: "#8a94a0" }, items: stats.filter((s) => !s.stage) }]);
    groups.forEach(({ g, items }) => {
      if (!items.length) return;
      const closed = g.id === "closed";
      const head = h("li", { class: "stage-head" + (closed ? " closed" : ""), style: "--c:" + g.color },
        h("button", { class: "sec-toggle", "aria-expanded": !closed, onclick: (ev) => {
          const open = ev.currentTarget.getAttribute("aria-expanded") !== "true";
          ev.currentTarget.setAttribute("aria-expanded", open);
          let el = ev.currentTarget.parentElement.nextElementSibling;
          while (el && !el.classList.contains("stage-head")) { el.hidden = !open; el = el.nextElementSibling; }
        }, html: svg(closed ? "chevR" : "chevD", 13) + "<i class=\"dot\" style=\"background:" + g.color + "\"></i><span>" + esc(g.name) + "</span><span class=\"today-n\">" + items.length + "</span>" }));
      list.appendChild(head);
      items.forEach((s) => { const li = rowEl(s); li.hidden = closed; list.appendChild(li); });
    });
  }
  function renderStageBar(stageList) {
    const bar = $("stageBar");
    bar.innerHTML = "";
    const counts = STAGES.map((g) => ({ g, n: stageList.filter((x) => x && x.id === g.id).length }));
    const total = counts.reduce((a, c) => a + c.n, 0);
    bar.hidden = !total;
    if (!total) return;
    counts.filter((c) => c.n).forEach((c) => bar.appendChild(h("div", { class: "stage-seg", style: "--c:" + c.g.color + ";flex:" + c.n, title: c.n + " × " + c.g.name },
      h("b", { text: c.n }), h("span", { text: c.g.name }))));
  }
  function openTag(name) { openTagName = name; renderTags(); }
  // notes that belong to a label: all notes if the email has the label, otherwise notes with it
  function notesWithTag(e, name) {
    return e.labels.some((t) => lc(t) === lc(name)) ? e.notes : e.notes.filter((n) => n.tags.some((t) => lc(t) === lc(name)));
  }
  function tagRows(name) {
    const rows = [];
    allEntries().forEach((e) => {
      if (!e.tags.some((t) => lc(t) === lc(name))) return;
      rows.push({ e, notes: notesWithTag(e, name) });
    });
    rows.sort((a, b) => (b.notes.some((n) => !isDone(n)) - a.notes.some((n) => !isDone(n))) || b.e.v.u - a.e.v.u);
    return rows;
  }
  function renderTagDetail() {
    const name = openTagName;
    const t = $("tagTitle");
    t.innerHTML = "";
    t.appendChild(tagChip(name));
    const rows = tagRows(name);
    const total = rows.reduce((s, r) => s + r.notes.length, 0);
    const open = rows.reduce((s, r) => s + r.notes.filter((n) => !isDone(n)).length, 0);
    $("tagStats").textContent = rows.length
      ? pl(rows.length, "email") + ", " + pl(total, "note") + ", " + open + " open"
      : "Nothing has this label yet.";
    const cur = stageOf(name);
    const sp = $("stagePick");
    sp.innerHTML = "";
    STAGES.forEach((g) => sp.appendChild(h("button", { class: "st", style: "--c:" + g.color, "aria-pressed": cur && cur.id === g.id ? "true" : "false",
      title: cur && cur.id === g.id ? "Click to clear the stage" : "Set stage: " + g.name,
      onclick: () => { setStage(name, cur && cur.id === g.id ? "" : g.id); renderTags(); toast(cur && cur.id === g.id ? "Stage cleared" : "Stage: " + g.name); } },
      h("i", { style: "background:" + g.color }), h("span", { text: g.name }))));
    $("tagAddHere").hidden = !ctx;
    $("tagLabelHere").hidden = !ctx || emailTags.some((x) => lc(x) === lc(name));
    $("tagLabelHere").textContent = "Label this " + (ctx && ctx.meeting ? "meeting" : "email");
    $("tagInsert").hidden = !(ctx && ctx.compose) || !open;
    $("tagMail").hidden = !rows.length;
    const box = $("tagMails");
    box.innerHTML = "";
    rows.forEach((r) => box.appendChild(mailBlock(r.e, r.notes)));
  }
  // A compact block: email subject (opens it) + its notes
  function mailBlock(e, list) {
    const block = h("div", { class: "mblock" },
      h("button", { class: "mblock-subj", title: "Open", onclick: () => openEmail(e.id) },
        ic(e.v.mt ? "cal" : e.v.draft ? "compose" : "mail", 14), h("bdi", { text: e.v.s })));
    if (!list.length) block.appendChild(h("div", { class: "mnote-empty muted small", text: "Labeled, no notes" }));
    list.slice().sort((a, b) => isDone(a) - isDone(b) || b.updated - a.updated).forEach((n) => {
      const st = statusOf(n.color), ds = dueState(n);
      block.appendChild(h("div", { class: "mnote" + (isDone(n) ? " done" : ""), style: "--c:" + st.hex },
        h("div", { class: "note-text", dir: "auto", text: n.text }),
        h("div", { class: "mnote-meta" }, h("span", { class: "st-mini" }, h("i", { style: "background:" + st.hex }), st.short),
          h("span", { class: "muted", text: fmtShort(n.updated) }), ds ? h("span", { class: "pill " + ds, text: dueText(n) }) : null, waitChip(n))));
    });
    return block;
  }
  async function newTag() {
    const v = await promptDialog("New label", "Name it after a CDF number, a topic or a supplier.", "", "Create label");
    const t = cleanTag(v);
    if (!t) return;
    registerTags([t]);
    $("tagSearch").value = "";
    openTag(allTagNames().find((x) => lc(x) === lc(t)) || t);
  }
  async function renameTag() {
    const old = openTagName;
    const v = await promptDialog("Rename label", "Notes on all emails will use the new name.", old, "Rename");
    const t = cleanTag(v);
    if (!t || t === old) return;
    const st = getTagStore();
    Object.keys(st.renamed).forEach((k) => { if (lc(k) === lc(t)) delete st.renamed[k]; });
    st.renamed[old] = t;
    st.removed = st.removed.filter((r) => lc(r) !== lc(t));
    st.names = st.names.filter((x) => lc(x) !== lc(old) && lc(x) !== lc(t)).concat([t]);
    const sg = st.stages[old]; if (sg) { delete st.stages[old]; st.stages[t] = sg; }
    saveTagStore(st);
    applyTagChangeHere("Label renamed");
    openTag(t);
  }
  async function deleteTag() {
    const name = openTagName;
    const yes = await confirmDialog("Delete label?", "Are you sure you want to delete the label “" + name +
      "”? It will be removed from all notes. The notes themselves are kept.", "Delete label");
    if (!yes) return;
    const st = getTagStore();
    st.removed.push(name);
    delete st.stages[name];
    st.names = st.names.filter((x) => lc(x) !== lc(name));
    Object.keys(st.renamed).forEach((k) => { if (lc(k) === lc(name)) delete st.renamed[k]; });
    saveTagStore(st);
    applyTagChangeHere("Label deleted");
    openTagName = null;
    renderTags();
  }
  function applyTagChangeHere(msg) {
    if (normaliseNoteTags()) persist(msg); else toast(msg);
    draftTags = normTags(draftTags, getTagStore());
    renderTagComposer();
    renderNotes();
  }
  function labelGroups(name) {
    return [{ title: TAG_PREFIX + name, color: tagPalette(name)[1],
      mails: tagRows(name).map((r) => ({ e: r.e, notes: r.notes.filter((n) => !isDone(n)).concat(r.notes.filter(isDone)) })) }];
  }
  function emailLabel() {
    const name = openTagName;
    const sg = stageOf(name);
    const out = renderMailHtml({ title: "Notes on " + name, subtitle: (sg ? "Stage: " + sg.name + " · " : "") + tagStatsLine(name), groups: labelGroups(name), maxText: NOTE_TEXT_MAX, showStatus: true });
    try { MB().displayNewMessageForm({ subject: "Notes – " + name, htmlBody: out.html }); }
    catch (e) { toast("Couldn't open a new email"); }
  }
  function insertLabel() {
    const name = openTagName;
    const groups = labelGroups(name).map((g) => Object.assign({}, g, { mails: g.mails.map((m) => ({ e: m.e, notes: m.notes.filter((n) => !isDone(n)) })).filter((m) => m.notes.length) }));
    insertHtml(renderMailHtml({ title: name, subtitle: "Open items", groups, maxText: NOTE_TEXT_MAX, showStatus: true, compact: true }).html,
      "Open items inserted");
  }
  function tagStatsLine(name) {
    const rows = tagRows(name);
    const total = rows.reduce((s, r) => s + r.notes.length, 0);
    const open = rows.reduce((s, r) => s + r.notes.filter((n) => !isDone(n)).length, 0);
    return pl(rows.length, "email") + ", " + pl(total, "note") + ", " + open + " open";
  }

  // ===================================================================
  // Document numbers (CDF etc.)
  // ===================================================================
  const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  function parsePrefixes(s) { return String(s || "").split(/[,;\s]+/).map((x) => x.trim()).filter(Boolean); }
  function refsOf(text, prefixes) {
    const out = new Set();
    prefixes.forEach((p) => {
      const re = new RegExp("(^|[^A-Za-z0-9])(" + escRe(p) + ")[\\s\\-_#:]*(\\d[A-Za-z0-9.\\-/]*)", "gi");
      let m;
      while ((m = re.exec(text))) {
        const num = m[3].replace(/\.(pdf|docx?|xlsx?|msg|zip|png|jpe?g)$/i, "").replace(/[.\-/]+$/, "").toUpperCase();
        out.add(p.toUpperCase() + (p.length > 1 ? " " : "") + num);
      }
    });
    return [...out];
  }

  // ===================================================================
  // Summary, weekly review, Excel, backup
  // ===================================================================
  function getPrefs() {
    const p = rsGet(PREFS_KEY);
    const d = {
      labels: STATUS.filter((s) => s.hex !== DONE_HEX).map((s) => s.hex), range: "all", group: "label",
      prefixes: "CDF", dueTop: true, to: (MB().userProfile || {}).emailAddress || "",
      cols: EXCEL_COLS.map((c) => ({ k: c[0], on: c[0] !== "created" })), header: true, onlyChanged: false, lastCopy: 0, waitAlert: 3,
    };
    const out = Object.assign(d, p || {});
    // keep column list complete if new columns were added
    const have = new Set(out.cols.map((c) => c.k));
    EXCEL_COLS.forEach((c) => { if (!have.has(c[0])) out.cols.push({ k: c[0], on: false }); });
    out.cols = out.cols.filter((c) => EXCEL_COLS.some((x) => x[0] === c.k));
    return out;
  }
  let prefsCache = null;
  function readPrefs() {
    const base = prefsCache || getPrefs();
    return Object.assign({}, base, {
      labels: [...document.querySelectorAll("#sumLabels input:checked")].map((i) => i.value),
      range: $("sumRange").value, group: $("sumGroup").value, prefixes: $("sumPrefixes").value.trim(),
      dueTop: $("sumDueTop").checked, to: $("sumTo").value.trim(),
      header: $("xlHeader").checked, onlyChanged: $("xlChanged").checked,
      waitAlert: (() => { const n = parseInt($("waitAlert").value, 10); return isNaN(n) ? 3 : Math.min(30, Math.max(1, n)); })(),
    });
  }
  function savePrefs(p) { prefsCache = p || readPrefs(); rsSet(PREFS_KEY, prefsCache); return prefsCache; }

  function buildSummaryUi() {
    const prefs = prefsCache = getPrefs();
    const wrap = $("sumLabels");
    STATUS.forEach((s) => {
      const cb = h("input", { type: "checkbox", value: s.hex, onchange: renderSummaryInfo });
      cb.checked = prefs.labels.includes(s.hex);
      wrap.appendChild(h("label", { class: "check" }, cb, h("i", { class: "dot", style: "background:" + s.hex }), s.short === "Review" ? "To review" : s.short === "Waiting" ? "Awaiting reply" : s.short));
    });
    $("sumRange").value = prefs.range;
    $("sumGroup").value = prefs.group;
    $("sumPrefixes").value = prefs.prefixes;
    $("sumDueTop").checked = prefs.dueTop;
    $("sumTo").value = prefs.to;
    $("xlHeader").checked = prefs.header;
    $("xlChanged").checked = prefs.onlyChanged;
    ["sumRange", "sumGroup", "sumDueTop", "xlChanged"].forEach((id) => $(id).addEventListener("change", renderSummaryInfo));
    $("sumPrefixes").addEventListener("input", () => { renderSummaryInfo(); renderTagComposer(); });
    $("sumPrefixes").addEventListener("change", () => { savePrefs(); renderTagComposer(); });
    $("sumTo").addEventListener("change", () => savePrefs());
    $("xlHeader").addEventListener("change", () => savePrefs());
    $("sumGo").onclick = createSummary;
    $("sumText").onclick = copySummaryText;
    $("sumCopy").onclick = copyForExcel;
    $("waitAlert").value = prefs.waitAlert;
    $("waitAlert").addEventListener("change", () => { savePrefs(); renderAll(); });
    $("weekGo").onclick = createWeekly;
    $("backupGo").onclick = createBackup;
    renderColumns();
  }
  function renderColumns() {
    const prefs = prefsCache || getPrefs();
    const ul = $("xlCols");
    ul.innerHTML = "";
    prefs.cols.forEach((c, i) => {
      const cb = h("input", { type: "checkbox", onchange: () => { c.on = cb.checked; savePrefs(Object.assign(readPrefs(), { cols: prefs.cols })); } });
      cb.checked = c.on;
      const move = (d) => {
        const j = i + d;
        if (j < 0 || j >= prefs.cols.length) return;
        const tmp = prefs.cols[i]; prefs.cols[i] = prefs.cols[j]; prefs.cols[j] = tmp;
        savePrefs(Object.assign(readPrefs(), { cols: prefs.cols }));
        renderColumns();
      };
      ul.appendChild(h("li", { class: "col" },
        h("label", { class: "check" }, cb, (EXCEL_COLS.find((x) => x[0] === c.k) || [0, c.k])[1]),
        h("span", { class: "col-move" },
          h("button", { class: "ibtn sm", title: "Move up", "aria-label": "Move up", text: "↑", disabled: i === 0 ? true : null, onclick: () => move(-1) }),
          h("button", { class: "ibtn sm", title: "Move down", "aria-label": "Move down", text: "↓", disabled: i === prefs.cols.length - 1 ? true : null, onclick: () => move(1) }))));
    });
  }
  function rangeStart(range) {
    if (range === "all") return 0;
    const t = startOfToday();
    return range === "today" ? t : t - (parseInt(range, 10) - 1) * DAY;
  }
  function selectedNotes(prefs) {
    const from = rangeStart(prefs.range);
    const out = [];
    allEntries().sort((a, b) => b.v.u - a.v.u).forEach((e) => e.notes.forEach((n) => {
      if (prefs.labels.includes(statusOf(n.color).hex) && (n.updated || 0) >= from) out.push({ e, n });
    }));
    return out;
  }
  function groupRows(rows, keyFn, colorFn, order) {
    const map = new Map();
    rows.forEach((r) => [].concat(keyFn(r)).forEach((k) => {
      if (!map.has(k)) map.set(k, { title: k, color: colorFn(k, r), mails: new Map() });
      const g = map.get(k);
      if (!g.mails.has(r.e.key)) g.mails.set(r.e.key, { e: r.e, notes: [] });
      g.mails.get(r.e.key).notes.push(r.n);
    }));
    const keys = order ? order.filter((k) => map.has(k)) : [...map.keys()];
    return keys.map((k) => { const g = map.get(k); return { title: g.title, color: g.color, mails: [...g.mails.values()] }; });
  }
  const sortNotes = (list) => list.sort((a, b) => isDone(a) - isDone(b) ||
    (a.due && !isDone(a) ? dueTs(a.due) : 9e15) - (b.due && !isDone(b) ? dueTs(b.due) : 9e15) || b.updated - a.updated);
  function buildGroups(prefs) {
    const rows = selectedNotes(prefs);
    const groups = [];
    if (prefs.dueTop) {
      const late = [];
      allEntries().forEach((e) => e.notes.forEach((n) => { if (isLate(n)) late.push({ e, n }); }));
      groups.push(...groupRows(late, () => "⏰ Overdue and due today", () => "#d9822b"));
    }
    if (prefs.group === "ref") {
      const pre = parsePrefixes(prefs.prefixes);
      const g = groupRows(rows, (r) => { const x = refsOf(r.e.v.s + " " + r.n.text + " " + (r.n.att || ""), pre); return x.length ? x : ["No document number"]; }, () => "#8a8378");
      g.sort((a, b) => (a.title === "No document number") - (b.title === "No document number") || a.title.localeCompare(b.title, undefined, { numeric: true }));
      groups.push(...g);
    } else if (prefs.group === "tag") {
      const g = groupRows(rows, (r) => { const tg = unionTags(r.e.labels, r.n.tags); return tg.length ? tg.map((t) => TAG_PREFIX + t) : ["No label"]; },
        (k) => (k === "No label" ? "#8a8378" : tagPalette(k.slice(TAG_PREFIX.length))[1]));
      g.sort((a, b) => (a.title === "No label") - (b.title === "No label") || a.title.localeCompare(b.title, undefined, { numeric: true }));
      groups.push(...g);
    } else {
      groups.push(...groupRows(rows, (r) => statusOf(r.n.color).name, (k) => STATUS.find((s) => s.name === k).hex, STATUS.map((s) => s.name)));
    }
    groups.forEach((g) => g.mails.forEach((m) => sortNotes(m.notes)));
    return groups;
  }
  const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
  function dirOf(s) {
    const m = String(s).match(/[A-Za-z֐-׿؀-ۿ]/);
    return m && /[֐-ۿ]/.test(m[0]) ? "rtl" : "ltr";
  }
  function messageLink(itemId) {
    if (!itemId) return "";
    try {
      const rest = MB().convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
      return "https://outlook.office365.com/owa/?ItemID=" + encodeURIComponent(rest) + "&exvsurl=1&viewmodel=ReadMessageItem";
    } catch (e) { return ""; }
  }
  const PILL = {
    overdue: "background:#fde2e1;color:#a8241b;font-weight:600", today: "background:#fdecd0;color:#8a5000;font-weight:600",
    soon: "background:#e4eefb;color:#1f5fa8", later: "background:#e4eefb;color:#1f5fa8",
    done: "background:#e3f3e3;color:#2e7d32",
  };
  const pillHtml = (text, style) => '<span style="font-size:11px;padding:1px 6px;border-radius:9px;white-space:nowrap;' + style + '">' + esc(text) + "</span>";
  // Email HTML for summaries (inline styles only – Outlook renders email with Word)
  function renderMailHtml(o) {
    let html = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1d2731">';
    if (!o.compact) {
      html += '<div style="font-size:18px;font-weight:600;margin-bottom:2px">' + esc(o.title) + "</div>" +
        '<div style="color:#66737d;font-size:12px;margin-bottom:8px">' + esc(o.subtitle || "") + "</div>";
    } else {
      html += '<div style="font-weight:600;margin:4px 0">' + esc(o.title) + (o.subtitle ? ' <span style="color:#66737d;font-weight:400">– ' + esc(o.subtitle) + "</span>" : "") + "</div>";
    }
    let skipped = 0;
    o.groups.forEach((g) => {
      if (!g.mails.length) return;
      const head = '<div style="margin:16px 0 6px;padding:5px 9px;background:#f1f3f4;border-left:4px solid ' + g.color +
        ';font-weight:600">' + esc(g.title) + ' <span style="color:#66737d;font-weight:400">(' + g.mails.length + ")</span></div>";
      if (html.length + head.length > BODY_MAX) { skipped += g.mails.length; return; }
      html += head;
      g.mails.forEach((m) => {
        const link = messageLink(m.e.id);
        const subj = esc((m.e.v.mt ? "📅 " : m.e.v.draft ? "✉️ " : "") + m.e.v.s);
        let block = '<div style="margin:0 0 12px"><div dir="' + dirOf(m.e.v.s) + '" style="text-align:left"><b>' +
          (link ? '<a href="' + esc(link) + '" style="color:#1f5fad">' + subj + "</a>" : subj) + "</b></div>";
        m.notes.forEach((n) => {
          const t = n.text.length > o.maxText ? n.text.slice(0, o.maxText) + "…" : n.text;
          const st = statusOf(n.color), ds = dueState(n);
          const meta = [];
          if (o.showStatus || g.title.startsWith("⏰")) meta.push(esc(st.name));
          meta.push(esc(fmtDate(n.updated)));
          if (n.tags && n.tags.length && !g.title.startsWith(TAG_PREFIX)) meta.push(esc(n.tags.map((x) => TAG_PREFIX + x).join(", ")));
          if (n.att) meta.push("📎 " + esc(n.att));
          let pills = "";
          if (ds) pills = " " + pillHtml(dueText(n), PILL[ds]);
          else if (isDone(n)) pills = " " + pillHtml("✓ Done", PILL.done);
          block += '<div style="margin:4px 0 6px 2px;padding-left:9px;border-left:3px solid ' + st.hex + '">' +
            '<div dir="' + dirOf(t) + '" style="text-align:left">' + esc(t).replace(/\n/g, "<br>") + "</div>" +
            '<div style="color:#8b959d;font-size:11px">' + meta.join("&nbsp;&nbsp;") + pills + "</div></div>";
        });
        block += "</div>";
        if (html.length + block.length > BODY_MAX) { skipped++; return; }
        html += block;
      });
    });
    if (skipped) html += '<div style="color:#a8241b;font-size:12px">Shortened: ' + skipped + " emails didn't fit (size limit). Choose fewer categories or a shorter period.</div>";
    return { html: html + "</div>", skipped };
  }
  function countRows(rows) { return { notes: rows.length, mails: new Set(rows.map((r) => r.e.key)).size }; }
  function renderSummaryInfo() {
    const prefs = readPrefs();
    const c = countRows(selectedNotes(prefs));
    const late = allEntries().reduce((s, e) => s + e.notes.filter(isLate).length, 0);
    let txt = c.notes ? pl(c.notes, "note") + " from " + pl(c.mails, "email") : "No notes match these settings";
    if (prefs.dueTop && late) txt += ", plus " + late + " overdue or due today";
    $("sumInfo").textContent = txt;
    $("sumGo").disabled = !c.notes && !(prefs.dueTop && late);
    const xr = excelRows(prefs);
    $("xlInfo").textContent = xr.length ? pl(xr.length, "row") + (prefs.onlyChanged ? " changed since " + (prefs.lastCopy ? fmtDate(prefs.lastCopy) : "the start") : "") : "No rows to copy";
    $("sumCopy").disabled = !xr.length;
    renderStats();
  }
  function renderStats() {
    const all = allEntries();
    let open = 0, late = 0, wait = 0, doneWeek = 0;
    const ws = weekStart();
    all.forEach((e) => e.notes.forEach((n) => {
      if (isOpen(n)) open++;
      if (isLate(n)) late++;
      if (isWait(n) && !isDone(n)) wait++;
      if (isDone(n) && (n.doneAt || 0) >= ws) doneWeek++;
    }));
    $("stOpen").textContent = open; $("stLate").textContent = late; $("stWait").textContent = wait; $("stDone").textContent = doneWeek;
    $("stLateCell").classList.toggle("alert", late > 0);
  }
  function createSummary() {
    const prefs = savePrefs();
    const groups = buildGroups(prefs);
    if (!groups.length) return toast("Nothing to summarise");
    const c = countRows(selectedNotes(prefs));
    const o = { title: "My Notes summary – " + fmtDay(Date.now()), subtitle: pl(c.notes, "note") + " on " + pl(c.mails, "email") + ", " + $("sumRange").selectedOptions[0].textContent.toLowerCase(),
      groups, maxText: NOTE_TEXT_MAX, showStatus: prefs.group !== "label" };
    let out = renderMailHtml(o);
    if (out.skipped) out = renderMailHtml(Object.assign(o, { maxText: 120 }));
    sendNew("My Notes summary " + fmtDay(Date.now()), out.html, prefs.to);
  }
  // Plain text for Teams / chat
  function summaryText(groups, title, subtitle) {
    const lines = [title, subtitle, ""];
    groups.forEach((g) => {
      if (!g.mails.length) return;
      lines.push(g.title + " (" + g.mails.length + ")");
      g.mails.forEach((m) => {
        lines.push("• " + m.e.v.s);
        m.notes.forEach((n) => {
          const st = statusOf(n.color), ds = dueState(n);
          const extra = [st.short].concat(ds ? [dueText(n)] : [], isDone(n) ? ["done"] : [], isWait(n) && !isDone(n) ? ["waiting " + waitDays(n) + " wd"] : []);
          lines.push("   – " + n.text.replace(/\s+/g, " ") + "  [" + extra.join(", ") + "]");
        });
      });
      lines.push("");
    });
    return lines.join("\n").trim();
  }
  function copySummaryText() {
    const prefs = savePrefs();
    const groups = buildGroups(prefs);
    if (!groups.length) return toast("Nothing to copy");
    const c = countRows(selectedNotes(prefs));
    const text = summaryText(groups, "My Notes summary – " + fmtDay(Date.now()), pl(c.notes, "note") + " on " + pl(c.mails, "email") + ", " + $("sumRange").selectedOptions[0].textContent.toLowerCase());
    copyPlain(text).then((done) => {
      if (done) { $("copyBox").hidden = true; toast("Summary copied as text. Paste it into Teams."); }
      else { const box = $("copyBox"); box.hidden = false; box.value = text; box.focus(); box.select(); toast("Press Ctrl+C to copy"); }
    });
  }
  async function copyPlain(text) {
    try { if (navigator.clipboard && navigator.clipboard.writeText) { await navigator.clipboard.writeText(text); return true; } } catch (e) { /* fall through */ }
    try {
      const ta = h("textarea", { style: "position:fixed;left:-9999px;top:0" }); ta.value = text;
      document.body.appendChild(ta); ta.select(); const done = document.execCommand("copy"); ta.remove(); return done;
    } catch (e) { return false; }
  }
  function sendNew(subj, html, to) {
    const opts = { subject: subj, htmlBody: html };
    if (to) opts.toRecipients = to.split(/[;,\s]+/).filter(Boolean);
    try { MB().displayNewMessageForm(opts); } catch (e) { toast("Couldn't open a new email: " + e.message); }
  }
  function weeklyGroups() {
    const ws = weekStart(), we = ws + 7 * DAY, today = startOfToday();
    const rows = [];
    allEntries().forEach((e) => e.notes.forEach((n) => rows.push({ e, n })));
    const sec = (title, color, test) => groupRows(rows.filter((r) => test(r.n)), () => title, () => color)[0] || { title, color, mails: [] };
    return [
      sec("Overdue", "#d9534f", (n) => !isDone(n) && n.due && dueTs(n.due) < today),
      sec("Due this week", "#d9822b", (n) => !isDone(n) && n.due && dueTs(n.due) >= today && dueTs(n.due) < we),
      sec("Awaiting reply", WAIT_HEX, (n) => !isDone(n) && statusOf(n.color).hex === WAIT_HEX),
      sec("New this week", REVIEW_HEX, (n) => (n.created || n.updated) >= ws),
      sec("Closed this week", DONE_HEX, (n) => isDone(n) && (n.doneAt || 0) >= ws),
    ];
  }
  function createWeekly() {
    const groups = weeklyGroups();
    groups.forEach((g) => g.mails.forEach((m) => sortNotes(m.notes)));
    const count = (i) => groups[i].mails.reduce((s, m) => s + m.notes.length, 0);
    const sub = count(0) + " overdue, " + count(1) + " due this week, " + count(2) + " awaiting reply, " + count(3) + " new, " + count(4) + " closed";
    const out = renderMailHtml({ title: "Weekly review – week of " + fmtDay(weekStart()), subtitle: sub, groups, maxText: 200, showStatus: true });
    sendNew("My Notes weekly review " + fmtDay(weekStart()), out.html, (prefsCache || getPrefs()).to);
  }
  function excelRows(prefs) {
    let rows = selectedNotes(prefs);
    if (prefs.onlyChanged && prefs.lastCopy) rows = rows.filter((r) => (r.n.updated || 0) > prefs.lastCopy);
    return rows;
  }
  function copyForExcel() {
    const prefs = savePrefs();
    const rows = excelRows(prefs);
    if (!rows.length) return toast("No rows to copy");
    const pre = parsePrefixes(prefs.prefixes);
    const cols = prefs.cols.filter((c) => c.on).map((c) => c.k);
    const val = (k, r) => {
      const n = r.n, st = dueState(n);
      switch (k) {
        case "ref": return refsOf(r.e.v.s + " " + n.text + " " + (n.att || ""), pre).join(", ");
        case "labels": return unionTags(r.e.labels, n.tags).join(", ");
        case "subject": return r.e.v.s;
        case "status": return statusOf(n.color).name.replace(/^\S+\s/, "");
        case "note": return n.text;
        case "file": return n.att || "";
        case "created": return fmtXl(n.created || n.updated);
        case "updated": return fmtXl(n.updated);
        case "due": return n.due ? fmtDueLong(n.due) : "";
        case "state": return isDone(n) ? "Done" : isInfo(n) ? "Info" : st === "overdue" ? "Overdue" : st === "today" ? "Due today" : "Open";
        case "link": return messageLink(r.e.id);
        default: return "";
      }
    };
    const head = cols.map((k) => EXCEL_COLS.find((x) => x[0] === k)[1]);
    const data = rows.map((r) => cols.map((k) => val(k, r)));
    const all = prefs.header ? [head].concat(data) : data;
    const tsv = all.map((row) => row.map((c) => (/[\t\n"]/.test(c) ? '"' + String(c).replace(/"/g, '""') + '"' : c)).join("\t")).join("\r\n");
    const html = "<table>" + (prefs.header ? "<tr>" + head.map((x) => "<th>" + esc(x) + "</th>").join("") + "</tr>" : "") +
      data.map((row) => "<tr>" + row.map((c, i) => "<td>" + (cols[i] === "link" && c ? '<a href="' + esc(c) + '">Open</a>' : esc(c).replace(/\n/g, "<br>")) + "</td>").join("") + "</tr>").join("") + "</table>";
    copyRich(html, tsv).then((done) => {
      if (done) {
        $("copyBox").hidden = true;
        savePrefs(Object.assign(readPrefs(), { lastCopy: Date.now(), cols: prefs.cols }));
        renderSummaryInfo();
        toast(pl(rows.length, "row") + " copied. Paste into Excel.");
      } else {
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
          "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([text], { type: "text/plain" }) })]);
        return true;
      }
    } catch (e) { /* fall through */ }
    try {
      const d = h("div", { contenteditable: "true", style: "position:fixed;left:-9999px;top:0", html });
      document.body.appendChild(d);
      const r = document.createRange(); r.selectNodeContents(d);
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
      const done = document.execCommand("copy");
      sel.removeAllRanges(); d.remove();
      return done;
    } catch (e) { return false; }
  }

  // ----- backup and restore -----
  function backupPayload(dropDone) {
    const idx = JSON.parse(JSON.stringify(getIndex()));
    if (dropDone) Object.keys(idx).forEach((k) => {
      idx[k].n = (idx[k].n || []).filter((n) => statusOf(n.color).hex !== DONE_HEX);
      if (!idx[k].n.length) delete idx[k];
    });
    return { v: 2, at: Date.now(), idx, tags: getTagStore(), tpl: getTemplates(), prefs: prefsCache || getPrefs() };
  }
  function createBackup() {
    let p = backupPayload(false);
    let data = LZString.compressToBase64(JSON.stringify(p));
    let partial = false;
    if (data.length > BODY_MAX - 3000) { p = backupPayload(true); data = LZString.compressToBase64(JSON.stringify(p)); partial = true; }
    if (data.length > BODY_MAX - 3000) return toast("The backup is too big for one email");
    const n = Object.values(p.idx).reduce((s, v) => s + (v.n || []).length, 0);
    const lines = data.match(/.{1,76}/g).join("\n");
    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1d2731">' +
      '<div style="font-size:18px;font-weight:600">My Notes backup – ' + fmtDay(p.at) + "</div>" +
      "<p>" + pl(n, "note") + " on " + pl(Object.keys(p.idx).length, "email") + ", " + pl(p.tags.names.length, "label") + ", " + pl(p.tpl.length, "quick note") + "." +
      (partial ? " Done notes were left out to fit in one email." : "") + "</p>" +
      "<p>To restore: open this email, then open My Notes and click Restore. The notes themselves stay on your emails; this backup restores the overview, labels, quick notes and settings.</p>" +
      '<pre style="font-family:Consolas,monospace;font-size:9px;color:#8b959d;white-space:pre-wrap">MYNOTES-BACKUP-BEGIN\n' + lines + "\nMYNOTES-BACKUP-END</pre></div>";
    sendNew(BACKUP_SUBJECT + " " + fmtDay(p.at), html, (MB().userProfile || {}).emailAddress || "");
  }
  function detectBackup() {
    backupFound = null;
    if (!ctx || ctx.compose || !String(subject).startsWith(BACKUP_SUBJECT) || !ctx.item.body) { renderBackup(); return; }
    ctx.item.body.getAsync(Office.CoercionType.Text, (r) => {
      if (!ok(r)) return;
      const m = String(r.value).match(/MYNOTES-BACKUP-BEGIN([\s\S]*?)MYNOTES-BACKUP-END/);
      if (!m) return;
      try {
        const p = JSON.parse(LZString.decompressFromBase64(m[1].replace(/\s+/g, "")));
        if (p && p.v === 2 && p.idx) { backupFound = p; renderBackup(); }
      } catch (e) { /* not a valid backup */ }
    });
  }
  async function restoreBackup() {
    const p = backupFound;
    if (!p) return;
    const yes = await dialog({ title: "Restore this backup?", ok: "Restore",
      text: "Emails, labels and quick notes from " + fmtDay(p.at) + " are added to what you have now. Newer notes are kept." });
    if (!yes) return;
    const idx = getIndex();
    Object.keys(p.idx).forEach((k) => { if (!idx[k] || (idx[k].u || 0) < (p.idx[k].u || 0)) idx[k] = p.idx[k]; });
    setIndex(idx);
    const st = getTagStore();
    (p.tags.names || []).forEach((t) => { if (!st.names.some((x) => lc(x) === lc(t))) st.names.push(t); });
    saveTagStore(st);
    const tpl = getTemplates();
    (p.tpl || []).forEach((t) => { if (!tpl.some((x) => lc(x.text) === lc(t.text))) tpl.push(t); });
    setTemplates(tpl.slice(-12));
    backupFound = null;
    renderAll();
    toast("Backup restored");
  }

  // ===================================================================
  // Rendering
  // ===================================================================
  function renderAll() {
    renderHeader();
    renderQuick();
    renderNotes();
    renderThread();
    renderReply();
    renderBackup();
    renderActions();
    refreshBadges();
    if (tab === "today") renderToday();
    if (tab === "all") renderIndex();
    if (tab === "tag") renderTags();
    if (tab === "sum") renderSummaryInfo();
  }
  function renderHeader() {
    const kind = !ctx ? "note" : ctx.meeting ? "cal" : ctx.compose ? "compose" : "mail";
    $("tbIcon").innerHTML = svg(kind, 15);
    $("subject").textContent = !ctx ? "Select an email to see its notes" : subject || (ctx.meeting ? "(no title)" : ctx.compose ? "(new email)" : "(no subject)");
    $("subject").setAttribute("dir", "auto");
    const open = notes.filter(isOpen).length, info = notes.filter(isInfo).length, done = notes.filter(isDone).length;
    const parts = [];
    if (open) parts.push(open + " open");
    if (info) parts.push(info + " info");
    if (done) parts.push(done + " done");
    $("tbNotes").textContent = !notes.length ? "None" : parts.join(", ");
    $("tabItemName").textContent = !ctx ? "Email" : ctx.meeting ? "Meeting" : ctx.compose ? "Draft" : "Email";
    // next due on this item
    const nd = notes.filter((n) => dueState(n)).sort((a, b) => dueTs(a.due) - dueTs(b.due))[0];
    const dv = $("tbDue");
    dv.className = "tb-v" + (nd ? " " + dueState(nd) : "");
    dv.textContent = !nd ? "–" : dueState(nd) === "today" ? "Today" : dueState(nd) === "overdue" ? dueText(nd) : fmtShort(dueTs(nd.due));
    const th = threadEntries();
    const thN = th.reduce((s, e) => s + e.notes.length, 0);
    $("tbThread").textContent = !ctx || !ctx.conv ? "–" : thN ? pl(thN, "note") : "None";
    $("tbThreadCell").title = thN ? pl(thN, "note") + " on " + pl(th.length, "other email") + " in this conversation. Click to show."
      : "Notes on other emails in this conversation";
    $("tbThreadCell").classList.toggle("link", !!thN);
    const used = spaceUsed();
    const pct = Math.round(used * 100);
    $("tbSpace").style.width = Math.min(100, pct) + "%";
    $("tbSpaceCell").classList.toggle("warn", used > 0.7 && used <= 0.9);
    $("tbSpaceCell").classList.toggle("full", used > 0.9);
    $("tbSpacePct").textContent = pct + "%";
    $("tbSpaceCell").title = "Note space on this email: " + pct + "% of about 2,500 characters.";
    $("spaceNote").hidden = used <= 0.7;
    $("spaceNote").textContent = pct + "% of note space used";
  }
  function renderNotes() {
    const open = notes.filter((n) => !isDone(n));
    const done = notes.filter(isDone);
    $("count").textContent = open.length;
    $("count").hidden = !open.length;
    const ul = $("notes");
    ul.innerHTML = "";
    if (ctx && !notes.length) ul.appendChild(emptyState("No notes on this email yet", "Write one above, or pick a quick note."));
    else if (ctx && !open.length) ul.appendChild(emptyState("Everything here is done", "Done notes are below. Delete them when you no longer need them.", "check"));
    open.sort((a, b) => isInfo(a) - isInfo(b) || (a.due ? dueTs(a.due) : 9e15) - (b.due ? dueTs(b.due) : 9e15) || b.updated - a.updated)
      .forEach((n) => ul.appendChild(noteEl(n)));
    $("doneWrap").hidden = !done.length;
    const dl = $("doneNotes");
    dl.innerHTML = "";
    if (done.length) {
      $("doneToggle").innerHTML = svg(showDone ? "chevD" : "chevR", 14) + "<span>Done (" + done.length + ")</span>";
      $("doneToggle").setAttribute("aria-expanded", showDone);
      $("doneClear").innerHTML = svg("trash", 14) + "<span>" + (open.length ? "Delete done notes" : "Delete all notes") + "</span>";
      dl.hidden = !showDone;
      done.sort((a, b) => b.updated - a.updated).forEach((n) => dl.appendChild(noteEl(n)));
    }
    renderHeader();
    renderActions();
  }
  function noteEl(n) {
    const st = statusOf(n.color), done = isDone(n), ds = dueState(n);
    const edited = n.updated - n.created > 1500;
    const li = h("li", { class: "note" + (done ? " done" : "") + (editingId === n.id ? " is-editing" : ""), style: "--c:" + st.hex });
    li.appendChild(h("div", { class: "note-head" },
      h("span", { class: "st-pill" }, h("i", { style: "background:" + st.hex }), st.short),
      ds ? h("span", { class: "due " + ds, html: svg("clock", 11) + "<span>" + esc(dueText(n)) + "</span>" }) : null,
      waitChip(n)));
    li.appendChild(h("div", { class: "note-text", dir: "auto", text: n.text }));
    if (n.tags && n.tags.length) li.appendChild(h("div", { class: "note-tags" }, n.tags.map((t) => tagChip(t, () => { switchTab("tag"); openTag(t); }))));
    const actions = h("div", { class: "note-actions" },
      isInfo(n) ? null : iconBtn(done ? "undo" : "check", done ? "Reopen" : "Mark as done", () => toggleDone(n.id), done ? "" : "ok"),
      !done && !isInfo(n) ? iconBtn("snooze", n.due ? "Move the due date" : "Set a due date", (ev) => snoozeMenu(ev.currentTarget, n.due,
        (v) => mutate(() => { const x = notes.find((y) => y.id === n.id); if (x) { x.due = v; x.updated = Date.now(); } }, v ? "Due " + fmtShort(dueTs(v)) : "Due date removed"))) : null,
      n.due && !done ? iconBtn("cal", "Add a reminder to my calendar", () => addToCalendar(n)) : null,
      iconBtn("edit", "Edit", () => editNote(n.id)),
      iconBtn("trash", "Delete", () => deleteNote(n.id), "danger"));
    const foot = h("div", { class: "note-foot" },
      h("span", { class: "muted", text: (edited ? "Edited " : "") + fmtDate(n.updated) }),
      n.att ? h("span", { class: "att", title: n.att, html: svg("clip", 11) + "<bdi>" + esc(n.att) + "</bdi>" }) : null,
      n.hist && n.hist.length ? h("button", { class: "linkbtn small", "aria-expanded": openHist.has(n.id),
        title: "Earlier versions of this note", html: svg("history", 12) + "<span>History (" + n.hist.length + ")</span>",
        onclick: () => { openHist.has(n.id) ? openHist.delete(n.id) : openHist.add(n.id); renderNotes(); } }) : null,
      actions);
    li.appendChild(foot);
    if (n.hist && n.hist.length && openHist.has(n.id)) {
      li.appendChild(h("ol", { class: "hist" }, n.hist.map((x) => h("li", null,
        h("div", { class: "note-text", dir: "auto", text: x.t }), h("div", { class: "muted small", text: "Written " + fmtDate(x.u) })))));
    }
    return li;
  }
  function waitChip(n) {
    if (!isWait(n) || isDone(n)) return null;
    const d = waitDays(n), lim = waitAlertDays();
    return h("span", { class: "wait" + (d >= lim ? " long" : ""), title: d + " workdays since this was marked as awaiting reply" + (d >= lim ? ". Longer than your " + lim + "-workday limit." : ""),
      html: svg("reply", 11) + "<span>" + (d ? "Waiting " + d + " wd" : "Waiting since today") + "</span>" });
  }
  function waitAlertDays() { const v = parseInt((prefsCache || getPrefs()).waitAlert, 10); return isNaN(v) ? 3 : v; }
  function emptyState(title, text, iconName) {
    return h("li", { class: "empty" },
      h("div", { class: "empty-art", html: iconName === "check" ? svg("check", 22) : svg("note", 22) }),
      h("div", { class: "empty-title", text: title }), text ? h("div", { class: "empty-text", text }) : null);
  }
  function renderThread() {
    const wrap = $("threadWrap");
    const th = threadEntries();
    const thN = th.reduce((s, e) => s + e.notes.length, 0);
    wrap.hidden = !thN;
    if (!thN) return;
    $("threadToggle").innerHTML = svg(threadOpen ? "chevD" : "chevR", 14) + svg("thread", 14) +
      "<span>In this thread: " + thN + (thN > 1 ? " notes" : " note") + " on " + th.length + (th.length > 1 ? " other emails" : " other email") + "</span>";
    $("threadToggle").setAttribute("aria-expanded", threadOpen);
    const body = $("threadBody");
    body.hidden = !threadOpen;
    body.innerHTML = "";
    th.sort((a, b) => b.v.u - a.v.u).forEach((e) => body.appendChild(mailBlock(e, e.notes)));
    renderHeader();
  }
  function renderReply() {
    const box = $("replyBox");
    box.hidden = !replyCands.length;
    if (!replyCands.length) return;
    const list = $("replyList");
    list.innerHTML = "";
    replyCands.slice(0, 4).forEach((c) => list.appendChild(h("li", null,
      h("div", { class: "note-text", dir: "auto", text: c.n.text }),
      h("div", { class: "reply-meta" }, h("bdi", { class: "muted small", text: c.e.v.s }),
        h("span", { class: "reply-btns" },
          h("button", { class: "btn sm", onclick: () => openEmail(c.e.id), text: "Open" }),
          h("button", { class: "btn sm primary", onclick: () => markRemoteDone(c), html: svg("check", 13) + "<span>Mark done</span>" }))))));
  }
  function renderBackup() {
    const box = $("restoreBox");
    box.hidden = !backupFound;
    if (!backupFound) return;
    const p = backupFound;
    const n = Object.values(p.idx).reduce((s, v) => s + (v.n || []).length, 0);
    $("restoreText").textContent = "From " + fmtDay(p.at) + ": " + pl(n, "note") + " on " + pl(Object.keys(p.idx).length, "email") + ", " +
      pl((p.tags.names || []).length, "label") + ", " + pl((p.tpl || []).length, "quick note") + ".";
  }
  function renderActions() {
    const open = notes.filter((n) => !isDone(n)).length;
    const read = ctx && !ctx.compose;
    $("actions").hidden = !ctx || !open;
    $("actReply").hidden = !read;
    $("actReplyAll").hidden = !read;
    $("actInsert").hidden = !(ctx && ctx.compose);
  }
  function attention() {
    const lim = waitAlertDays();
    const out = { late: 0, long: 0 };
    allEntries().forEach((e) => e.notes.forEach((n) => {
      if (isLate(n)) out.late++;
      else if (isWait(n) && waitDays(n) >= lim) out.long++;
    }));
    return out;
  }
  function refreshBadges() {
    const a = attention(), total = a.late + a.long;
    $("dueCount").hidden = !total;
    $("dueCount").textContent = total;
    $("dueCount").title = a.late + " overdue or due today" + (a.long ? ", " + a.long + " waiting too long" : "");
  }

  // ----- Today tab -----
  function renderToday() {
    const body = $("todayBody");
    body.innerHTML = "";
    const lim = waitAlertDays();
    const rows = [];
    allEntries().forEach((e) => e.notes.forEach((n) => { if (n.id && isOpen(n)) rows.push({ e, n }); }));
    const sec = [
      { title: "Overdue", color: "#d9534f", rows: rows.filter((r) => dueState(r.n) === "overdue").sort((a, b) => dueTs(a.n.due) - dueTs(b.n.due)) },
      { title: "Due today", color: "#d9822b", rows: rows.filter((r) => dueState(r.n) === "today") },
      { title: "Waiting too long", color: WAIT_HEX, rows: rows.filter((r) => isWait(r.n) && !isLate(r.n) && waitDays(r.n) >= lim).sort((a, b) => waitDays(b.n) - waitDays(a.n)),
        hint: "Awaiting reply for " + lim + " workdays or more" },
      { title: "Due this week", color: "#1f5fad", rows: rows.filter((r) => dueState(r.n) === "soon").sort((a, b) => dueTs(a.n.due) - dueTs(b.n.due)) },
      { title: "Still waiting", color: "#8a94a0", rows: rows.filter((r) => isWait(r.n) && !isLate(r.n) && dueState(r.n) !== "soon" && waitDays(r.n) < lim).sort((a, b) => waitDays(b.n) - waitDays(a.n)) },
    ];
    const urgent = sec[0].rows.length + sec[1].rows.length + sec[2].rows.length;
    const doneWeek = allEntries().reduce((s, e) => s + e.notes.filter((n) => isDone(n) && (n.doneAt || 0) >= weekStart()).length, 0);
    body.appendChild(h("div", { class: "today-lead" },
      h("div", { class: "today-date", text: new Date().toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" }) }),
      h("div", { class: "today-sum", text: urgent ? pl(urgent, "item") + " need attention" : "Nothing needs attention right now" })));
    if (!sec.some((x) => x.rows.length)) {
      body.appendChild(h("ul", { class: "notes" }, emptyState("All clear", "Overdue, due today, due this week and waiting notes show up here.", "check")));
    }
    sec.forEach((x) => {
      if (!x.rows.length) return;
      const block = h("section", { class: "today-sec", style: "--c:" + x.color },
        h("h3", { class: "today-h" }, h("span", { text: x.title }), h("span", { class: "today-n", text: x.rows.length })),
        x.hint ? h("div", { class: "hint", style: "margin:-2px 0 6px", text: x.hint }) : null);
      x.rows.forEach((r) => block.appendChild(todayRow(r.e, r.n)));
      body.appendChild(block);
    });
    body.appendChild(h("div", { class: "today-foot muted", text: doneWeek ? pl(doneWeek, "note") + " closed this week" : "Nothing closed this week yet" }));
  }
  function todayRow(e, n) {
    const st = statusOf(n.color), ds = dueState(n);
    const row = h("div", { class: "trow", style: "--c:" + st.hex });
    row.appendChild(h("div", { class: "note-text", dir: "auto", text: n.text }));
    row.appendChild(h("div", { class: "trow-meta" },
      h("button", { class: "trow-subj", title: "Open " + (e.v.mt ? "meeting" : "email"), onclick: () => openEmail(e.id) },
        ic(e.v.mt ? "cal" : "mail", 12), h("bdi", { text: e.v.s })),
      ds ? h("span", { class: "due " + ds, text: dueText(n) }) : null, waitChip(n),
      h("span", { class: "trow-actions" },
        iconBtn("check", "Mark as done", () => remoteUpdate(e, n.id, { done: 1 }, "Marked as done"), "ok"),
        iconBtn("snooze", n.due ? "Move the due date" : "Set a due date", (ev) => snoozeMenu(ev.currentTarget, n.due,
          (v) => remoteUpdate(e, n.id, { due: v }, v ? "Due " + fmtShort(dueTs(v)) : "Due date removed"))))));
    return row;
  }

  // ----- All notes tab -----
  function firstDue(ns) { return ns.filter((n) => dueState(n)).sort((a, b) => dueTs(a.due) - dueTs(b.due))[0]; }
  function renderIndex() {
    const q = lc($("search").value.trim());
    const f = $("idxFilter").value;
    const list = $("index");
    list.innerHTML = "";
    const entries = allEntries()
      .filter((e) => !q || lc(e.v.s + " " + e.tags.join(" ") + " " + e.notes.map((n) => n.text + " " + n.att).join(" ")).includes(q))
      .filter((e) => {
        if (f === "open") return e.notes.some(isOpen);
        if (f === "done") return e.notes.some(isDone);
        if (f === "due") return e.notes.some((n) => dueState(n));
        if (f === "late") return e.notes.some(isLate);
        if (f === "wait") return e.notes.some((n) => isWait(n) && !isDone(n));
        if (f === "meet") return !!e.v.mt;
        if (f === "tagged") return e.tags.length > 0;
        return true;
      })
      .sort((a, b) => {
        if (f === "due" || f === "late") return dueTs(firstDue(a.notes).due) - dueTs(firstDue(b.notes).due);
        return b.v.u - a.v.u;
      });
    if (!entries.length) {
      list.appendChild(emptyState(q || f !== "all" ? "Nothing matches" : "No notes yet",
        q || f !== "all" ? "Try another word or filter." : "Notes you write on emails show up here."));
    }
    entries.forEach((e) => {
      const open = e.notes.filter(isOpen);
      const fd = firstDue(e.notes);
      const latest = (open.length ? open : e.notes).slice().sort((x, y) => y.updated - x.updated)[0];
      const tags = e.tags.slice(0, 3);
      const dots = [...new Set(e.notes.map((n) => statusOf(n.color).hex))];
      list.appendChild(h("li", null, h("button", { class: "row", title: "Open", onclick: () => openEmail(e.id) },
        h("div", { class: "row-top" },
          h("span", { class: "row-subj" }, ic(e.v.mt ? "cal" : e.v.draft ? "compose" : "mail", 13), h("bdi", { text: e.v.s })),
          fd ? h("span", { class: "pill " + dueState(fd), text: dueText(fd) }) : null),
        h("div", { class: "row-sub" },
          h("span", { class: "dots" }, dots.map((c) => h("i", { style: "background:" + c }))),
          h("span", { class: "muted", text: !e.notes.length ? "Labeled, no notes" : [open.length + " open", e.notes.filter(isInfo).length ? e.notes.filter(isInfo).length + " info" : "", e.notes.filter(isDone).length ? e.notes.filter(isDone).length + " done" : ""].filter(Boolean).join(", ") })),
        latest ? h("div", { class: "row-prev", dir: "auto", text: latest.text.replace(/\s+/g, " ").slice(0, 90) }) : null,
        tags.length ? h("div", { class: "row-tags" }, tags.map((t) => tagChip(t))) : null)));
    });
    const used = indexUsage();
    $("storeBar").style.width = Math.round(used * 100) + "%";
    $("storeText").textContent = "Overview storage " + Math.round(used * 100) + "% used. When it fills up, the oldest done notes leave the overview first. They stay on their emails.";
    $("storeWrap").classList.toggle("warn", used > 0.85);
  }

  // ===================================================================
  // Dialogs, tabs, toast, wiring
  // ===================================================================
  function dialog(o) {
    return new Promise((resolve) => {
      const m = $("modal"), inp = $("modalInput"), okb = $("modalOk"), cancel = $("modalCancel");
      $("modalTitle").textContent = o.title;
      $("modalText").textContent = o.text || "";
      $("modalText").hidden = !o.text;
      const ml = $("modalList");
      ml.innerHTML = "";
      ml.hidden = !o.list;
      (o.list || []).forEach(([k, v]) => ml.appendChild(h("li", null, h("kbd", { text: k }), h("span", { text: v }))));
      $("modalIcon").innerHTML = o.danger ? svg("alert", 18) : o.input ? svg("tag", 18) : o.list ? svg("kbd", 18) : svg("note", 18);
      $("modalIcon").className = "modal-ic" + (o.danger ? " danger" : "");
      okb.textContent = o.ok;
      okb.className = "btn " + (o.danger ? "danger" : "primary");
      cancel.hidden = !!o.alert;
      inp.hidden = !o.input;
      inp.value = o.value || "";
      m.hidden = false;
      const close = (yes) => {
        m.hidden = true;
        okb.onclick = cancel.onclick = m.onclick = null;
        document.removeEventListener("keydown", onKey, true);
        resolve(yes ? (o.input ? inp.value : true) : (o.input ? null : false));
      };
      const onKey = (e) => {
        if (e.key === "Escape") { e.preventDefault(); close(false); }
        if (e.key === "Enter" && o.input && document.activeElement === inp) { e.preventDefault(); close(true); }
        if (e.key === "Tab") { // keep focus inside the dialog
          const f = [inp, cancel, okb].filter((x) => !x.hidden);
          const i = f.indexOf(document.activeElement);
          e.preventDefault();
          f[(i + (e.shiftKey ? -1 : 1) + f.length) % f.length].focus();
        }
      };
      okb.onclick = () => close(true);
      cancel.onclick = () => close(false);
      m.onclick = (e) => { if (e.target === m) close(false); };
      document.addEventListener("keydown", onKey, true);
      if (o.input) { inp.focus(); inp.select(); } else if (o.alert) okb.focus(); else cancel.focus();
    });
  }
  const confirmDialog = (title, text, okLabel) => dialog({ title, text, ok: okLabel, danger: true });
  const promptDialog = (title, text, value, okLabel) => dialog({ title, text, value, ok: okLabel, input: true });

  const SHORTCUTS = [
    ["Alt+N", "New note (focus the note field)"], ["Ctrl+Enter", "Save the note"], ["Esc", "Cancel editing / close menus"],
    ["Alt+D", "Insert today's date"], ["Alt+T", "Due today"], ["Alt+L", "Add a label to the note"],
    ["Alt+1 … Alt+6", "Status: Note, Urgent, Review, Waiting, Done, Info"], ["Alt+E", "Label the email"], ["Alt+/", "This list"],
  ];
  function showShortcuts() {
    dialog({ title: "Keyboard shortcuts", ok: "Close", alert: true, list: SHORTCUTS });
  }
  function onGlobalKey(e) {
    if (!$("modal").hidden) return;
    if (e.key === "Escape") { if (popEl) { e.preventDefault(); closePopover(); } return; }
    if (!e.altKey || e.ctrlKey || e.metaKey) return;
    const openComposer = () => { switchTab("item"); $("composer").classList.add("open"); };
    if (e.code === "KeyN") { e.preventDefault(); openComposer(); if (!editingId) resetComposer(); $("input").focus(); }
    else if (e.code === "KeyL") { e.preventDefault(); openComposer(); $("tagInput").focus(); }
    else if (e.code === "KeyE") { e.preventDefault(); switchTab("item"); $("elInput").focus(); }
    else if (e.code === "KeyT") { e.preventDefault(); openComposer(); setDue("today", document.querySelector('#dueChips button[data-due="today"]')); $("input").focus(); }
    else if (e.code === "KeyD") { e.preventDefault(); openComposer(); insertToday(); }
    else if (e.code === "Slash") { e.preventDefault(); showShortcuts(); }
    else if (/^Digit[1-6]$/.test(e.code)) { e.preventDefault(); openComposer(); setColor(STATUS[parseInt(e.code.slice(5), 10) - 1].hex); }
  }
  function switchTab(which) {
    tab = which;
    closePopover();
    ["item", "today", "all", "tag", "sum"].forEach((t) => {
      $("tab-" + t).setAttribute("aria-selected", which === t);
      $("view-" + t).hidden = which !== t;
    });
    if (which === "today") renderToday();
    if (which === "all") renderIndex();
    if (which === "tag") renderTags();
    if (which === "sum") renderSummaryInfo();
  }
  let toastTimer;
  function toast(msg) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 2400);
  }

  function wireUi() {
    $("input").addEventListener("input", updateComposerState);
    $("input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveNote(); }
      if (e.key === "Escape" && editingId) resetComposer();
    });
    $("save").onclick = saveNote;
    $("cancel").onclick = resetComposer;
    $("tplSave").onclick = saveAsTemplate;
    $("insDate").onclick = insertToday;
    $("tplEdit").onclick = () => { const p = $("tplPanel"); p.hidden = !p.hidden; $("tplEdit").setAttribute("aria-expanded", !p.hidden); };
    $("dueClear").onclick = () => { $("due").value = ""; document.querySelectorAll("#dueChips button").forEach((b) => b.setAttribute("aria-pressed", "false")); };
    document.querySelectorAll("#dueChips button").forEach((b) => { b.onclick = () => setDue(b.dataset.due, b); });
    $("doneToggle").onclick = () => { showDone = !showDone; renderNotes(); };
    $("doneClear").onclick = deleteDone;
    $("threadToggle").onclick = () => { threadOpen = !threadOpen; renderThread(); };
    $("tbThreadCell").onclick = () => { if (threadEntries().length) { switchTab("item"); threadOpen = true; renderThread(); $("threadWrap").scrollIntoView({ block: "start" }); } };
    $("replyDismiss").onclick = dismissReply;
    $("restoreGo").onclick = restoreBackup;
    $("actReply").onclick = () => replyWithNotes(false);
    $("actReplyAll").onclick = () => replyWithNotes(true);
    $("actInsert").onclick = () => insertHtml(notesHtml(notes.filter((n) => !isDone(n))), "Notes inserted");
    ["item", "today", "all", "tag", "sum"].forEach((t) => { $("tab-" + t).onclick = () => { if (t === "tag") openTagName = null; switchTab(t); }; });
    $("search").addEventListener("input", renderIndex);
    $("idxFilter").onchange = renderIndex;
    // labels
    const inp = $("tagInput");
    inp.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === "," || e.key === ";" || (e.key === "Tab" && inp.value.trim())) { e.preventDefault(); addTagFromInput(); }
      else if (e.key === "Backspace" && !inp.value && draftTags.length) { draftTags.pop(); renderTagComposer(); }
    });
    inp.addEventListener("change", () => { if (allTagNames().some((t) => lc(t) === lc(inp.value.trim()))) addTagFromInput(); });
    $("tagSearch").addEventListener("input", renderTags);
    $("tagNew").onclick = newTag;
    $("tagBack").onclick = () => { openTagName = null; renderTags(); };
    $("tagRename").onclick = renameTag;
    $("tagDelete").onclick = deleteTag;
    $("tagMail").onclick = emailLabel;
    $("tagInsert").onclick = insertLabel;
    $("tagLabelHere").onclick = () => { addEmailTag(openTagName); renderTags(); };
    const el = $("elInput");
    el.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === "," || e.key === ";" || (e.key === "Tab" && el.value.trim())) { e.preventDefault(); addEmailTagFromInput(); }
      else if (e.key === "Backspace" && !el.value && emailTags.length) removeEmailTag(emailTags[emailTags.length - 1]);
    });
    el.addEventListener("change", () => { if (allTagNames().some((t) => lc(t) === lc(el.value.trim()))) addEmailTagFromInput(); });
    $("tagAddHere").onclick = () => { const t = openTagName; switchTab("item"); resetComposer(); draftTags = [t]; renderTagComposer(); $("input").focus(); };
    $("tagByStage").onchange = () => { const st = getTagStore(); st.byStage = $("tagByStage").checked; saveTagStore(st); renderTags(); };
    $("tagCats").onchange = () => {
      const st = getTagStore(); st.cats = $("tagCats").checked; saveTagStore(st); syncCategory();
      toast(st.cats ? "Labels now show as categories" : "Label categories turned off");
    };
    $("ver").textContent = "v" + VERSION;
    $("kbdHelp").onclick = showShortcuts;
    document.addEventListener("keydown", onGlobalKey, true);
    window.addEventListener("resize", () => { const r = $("quickRow"); r.classList.toggle("fade", r.scrollWidth > r.clientWidth + 2); });
  }
})();
