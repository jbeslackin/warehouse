/* Warehouse Commander: Ranked Play — web app
 * Sign-in: Microsoft Entra ID (MSAL.js). Data: one JSON file in a SharePoint
 * document library, read and written through Microsoft Graph as the signed-in
 * user. Writes use the file's eTag (If-Match) so two people saving at once
 * never overwrite each other; the loser re-reads and re-applies automatically.
 * XP, levels, streaks, badges and seasons are all derived from completed
 * quests, so they can never drift out of sync with the board.
 */
(() => {
  const C = window.DQB_CONFIG || {};
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite"];

  const TIERS = {
    urgent: { label: "URGENT", xp: 100, order: 0, color: "var(--urgent)" },
    main:   { label: "MAIN QUEST", xp: 60, order: 1, color: "var(--main)" },
    side:   { label: "SIDE QUEST", xp: 30, order: 2, color: "var(--side)" },
    lore:   { label: "LORE", xp: 10, order: 3, color: "var(--lore)" },
  };
  const ON_TIME_BONUS = 0.25;
  const BOUNTY_MULT = 2, ASSIST_XP = 15, ASSIST_GIVER_XP = 5, MAX_ASSISTS = 4, BOSS_STEP_XP = 15;
  const SHIP_ST = {
    exception: { label: "EXCEPTION", order: 0 }, claim: { label: "CLAIM OPEN", order: 1 }, resolved: { label: "EXCEPTION RESOLVED", order: 2 },
    transit: { label: "IN TRANSIT", order: 3 }, label: { label: "LABEL CREATED", order: 4 }, delivered: { label: "DELIVERED", order: 5 },
  };
  const UPS_TABS = [["alerts","NEEDS ATTENTION"],["moving","ON THE WAY"],["delivered","DELIVERED"],["all","ALL"]];
  let upsFilter = "alerts", upsQuery = "", upsOpen = new Set();
  const RANKS = ["DOCK ROOKIE", "PALLET JOCKEY", "LABEL WIZARD", "FREIGHT CAPTAIN", "SHIPPING LEGEND", "LOGISTICS OVERLORD"];
  const COLORS = ["#ff8a1f", "#4fc3d9", "#5fd38a", "#ffb547", "#a99bf0", "#ff6fa8", "#7db7ff", "#e8ecf1"];
  const PER_LEVEL = 150;
  const COMPACT_AFTER_DAYS = 45;
  const TABS = [["all","ALL"],["mine","MINE"],["open","UNCLAIMED"],["urgent","URGENT"],["main","MAIN"],["side","SIDE"],["lore","LORE"],["intel","INTEL"],["warehouse","WAREHOUSE"],["personal","PERSONAL"],["done","CLEARED"]];
  const BADGES = [
    { key: "first",   label: "FIRST MISSION",   hint: "Clear your first quest",            test: s => s.clears >= 1 },
    { key: "ontime",  label: "ON THE CLOCK",    hint: "Clear 5 quests on time",            test: s => s.onTime >= 5 },
    { key: "iron",    label: "IRON DOCK",       hint: "5-workday streak",                  test: s => s.bestStreak >= 5 },
    { key: "boss",    label: "BOSS SLAYER",     hint: "Clear 10 urgent quests",            test: s => s.urgent >= 10 },
    { key: "paper",   label: "PAPER TRAIL",     hint: "Log proof on 10 quests",            test: s => s.proof >= 10 },
    { key: "squad",   label: "SQUAD UP",        hint: "Clear a squad quest",               test: s => s.squad >= 1 },
    { key: "customs", label: "CUSTOMS BREAKER", hint: "Clear a customs or broker hold",    test: s => s.customs >= 1 },
    { key: "century", label: "CENTURY CLUB",    hint: "Earn 1,000 XP",                     test: s => s.xp >= 1000 },
    { key: "assist",  label: "TEAM PLAYER",     hint: "Get credited with 5 assists",       test: s => s.assists >= 5 },
    { key: "bounty",  label: "BOUNTY HUNTER",   hint: "Clear 3 daily bounties",            test: s => s.bounties >= 3 },
    { key: "raid",    label: "RAID PARTY",      hint: "Help take down a boss fight",       test: s => s.raids >= 1 },
  ];

  let msalApp = null, account = null, siteId = null, itemId = null;
  let data = emptyDoc(), pdata = { version: 2, quests: [] }, loaded = false, busy = false, pendingRender = false, spawnChecked = false;
  let myId = null, myName = "", myEmail = "", isAdmin = false;
  let view = "board", filter = "all", lbMode = "season", dbMode = "week";
  let openSet = new Set(), completing = null, pickColor = COLORS[0], editing = false;
  let confirmQ = null, confirmP = null, confirmR = null, lastLevel = null;
  let replying = null, assisting = null, goalEdit = false;
  const drafts = {};
  let ledger = [], statsCache = {};
  const $ = (id) => document.getElementById(id);
  try { filter = localStorage.getItem("dqb-filter") || "all"; lbMode = localStorage.getItem("dqb-lb") || "season"; } catch (e) {}
  if (location.hash === "#tv") view = "tv";

  /* ---------- helpers ---------- */
  function emptyDoc() { return { version: 2, crew: [], quests: [], recurring: [] }; }
  function el(tag, attrs = {}, ...kids) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") n.className = v;
      else if (k === "style") n.setAttribute("style", v);
      else if (k === "value") n.value = v;
      else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
      else if (v !== null && v !== undefined && v !== false) n.setAttribute(k, v === true ? "" : v);
    }
    for (const k of kids.flat(Infinity)) if (k !== null && k !== undefined && k !== false) n.append(k.nodeType ? k : document.createTextNode(String(k)));
    return n;
  }
  const uid = (p = "q") => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const nowIso = () => new Date().toISOString();
  const pad = (n) => String(n).padStart(2, "0");
  const localDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const todayIso = () => localDate(new Date());
  const isoToLocalDay = (iso) => iso ? localDate(new Date(iso)) : null;
  const monthKey = (iso) => iso ? isoToLocalDay(iso).slice(0, 7) : null;
  const monthName = (key) => { const [y, m] = key.split("-").map(Number); return new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "short", year: "numeric" }); };
  const parseDay = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
  const isWorkday = (d) => d.getDay() >= 1 && d.getDay() <= 5;
  function prevWorkday(d) { const x = new Date(d); do { x.setDate(x.getDate() - 1); } while (!isWorkday(x)); return x; }
  function daysLeft(iso) { if (!iso) return null; const t = parseDay(todayIso()); return Math.round((parseDay(iso) - t) / 86400000); }
  const fmtDate = (iso) => parseDay(iso).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" });
  const fmtWhen = (iso) => new Date(iso).toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  const tierOf = (q) => TIERS[q.tier] || TIERS.lore;
  const member = (id) => data.crew.find(c => c.id === id) || null;
  const nameOf = (id) => id === "claude" ? "Claude" : ((member(id) && member(id).name) || "Former crew");
  const colorOf = (id) => (member(id) && member(id).color) || "#93a0b0";
  const joined = () => !!(myId && member(myId));
  const initials = (n) => (n || "?").split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();
  const levelOf = (xp) => Math.floor(xp / PER_LEVEL) + 1;
  const rankOf = (lvl) => RANKS[Math.min(lvl - 1, RANKS.length - 1)];
  const isPersonal = (q) => q.scope === "personal";
  const allQ = () => isAdmin ? data.quests.concat(pdata.quests) : data.quests;
  const isMine = (q) => isPersonal(q) ? isAdmin : q.assignees.includes(myId);
  const canEditQuest = (q) => isAdmin || isMine(q);
  const tabsFor = () => TABS.filter(([k]) => isAdmin || (k !== "personal" && k !== "warehouse"));

  let toastTimer;
  function toast(msg, kind = "") {
    const t = $("toast"); t.textContent = msg; t.className = "toast show" + (kind ? " " + kind : "");
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.className = "toast"; }, 2800);
  }
  function fatal(title, body) { $("list").replaceChildren(el("div", { class: "empty" }, el("h2", {}, title), el("p", {}, body))); }

  /* ---------- Data shape ---------- */
  function normalize(doc, kind = "shared") {
    doc = doc && typeof doc === "object" ? doc : {};
    doc.version = 2;
    doc.crew = Array.isArray(doc.crew) ? doc.crew : [];
    doc.recurring = Array.isArray(doc.recurring) ? doc.recurring : [];
    doc.shipments = (Array.isArray(doc.shipments) ? doc.shipments : []).filter(s => s && s.tracking).map(s => {
      s.events = Array.isArray(s.events) ? s.events : [];
      s.status = SHIP_ST[s.status] ? s.status : "transit";
      s.ack = !!s.ack;
      return s;
    });
    doc.quests = (Array.isArray(doc.quests) ? doc.quests : []).map(q => {
      q.assignees = Array.isArray(q.assignees) ? q.assignees : (q.assignee ? [q.assignee] : []);
      delete q.assignee;
      q.credit = Array.isArray(q.credit) ? q.credit : (q.doneBy ? [q.doneBy] : []);
      delete q.doneBy;
      q.notes = Array.isArray(q.notes) ? q.notes : [];
      q.history = Array.isArray(q.history) ? q.history : [];
      q.checklist = Array.isArray(q.checklist) ? q.checklist : [];
      q.status = q.status === "done" ? "done" : "open";
      q.squad = !!q.squad; q.requiresProof = !!q.requiresProof; q.boss = !!q.boss;
      q.assists = Array.isArray(q.assists) ? q.assists : [];
      q.fromInbox = !!q.fromInbox; q.reviewed = q.fromInbox ? !!q.reviewed : true;
      q.intel = q.intel && q.intel.text ? q.intel : null;
      q.scope = kind === "personal" ? "personal" : "warehouse";
      return q;
    });
    return doc;
  }
  /* Keep the file small: long-finished quests keep only what XP and history need. */
  function compact(doc) {
    const shipCut = Date.now() - 30 * 86400000;
    if (doc.shipments) doc.shipments = doc.shipments.filter(s => s.status !== "delivered" || !s.lastAt || Date.parse(s.lastAt) >= shipCut);
    const cutoff = Date.now() - COMPACT_AFTER_DAYS * 86400000;
    doc.quests.forEach(q => {
      if (q.status === "done" && q.doneAt && Date.parse(q.doneAt) < cutoff && !q.compacted) {
        q.notes = []; q.history = []; q.details = ""; q.intel = null; q.compacted = true;
        q.checklist = q.boss ? q.checklist.filter(c => c.done).map(c => ({ id: c.id, text: "", done: true, by: c.by, at: c.at })) : [];
      }
    });
  }
  function log(q, action, detail) {
    q.history.push({ at: nowIso(), by: myId, action, detail: detail || "" });
    if (q.history.length > 40) q.history = q.history.slice(-40);
  }

  /* ---------- XP ledger, streaks, badges ---------- */
  /* Every XP grant is a ledger entry. kind: clear | boss-step | assist | assist-given */
  function awards(q) {
    const out = [], mk = (id, xp, kind, at, extra = {}) => ({ id, xp, kind, at, day: isoToLocalDay(at), month: monthKey(at), q, onTime: false, ...extra });
    if (q.boss) q.checklist.forEach(c => { if (c.done && c.by) out.push(mk(c.by, BOSS_STEP_XP, "boss-step", c.at || q.doneAt || q.createdAt)); });
    (q.assists || []).forEach(a => { out.push(mk(a.to, ASSIST_XP, "assist", a.at)); if (a.by) out.push(mk(a.by, ASSIST_GIVER_XP, "assist-given", a.at)); });
    if (q.status === "done" && q.credit.length && q.doneAt) {
      const onTime = !!(q.due && isoToLocalDay(q.doneAt) <= q.due);
      const bounty = !!(q.bounty && isoToLocalDay(q.doneAt) === q.bounty);
      const total = Math.round(tierOf(q).xp * (onTime ? 1 + ON_TIME_BONUS : 1) * (bounty ? BOUNTY_MULT : 1));
      const each = Math.round(total / q.credit.length);
      q.credit.forEach(id => out.push(mk(q.doneAt ? id : id, each, "clear", q.doneAt, { onTime, bounty })));
    }
    return out;
  }
  function buildLedger() {
    ledger = data.quests.flatMap(awards);
    statsCache = {};
  }
  const clearsOf = (list) => list.filter(a => a.kind === "clear");
  function xpFor(id, month) { return ledger.filter(a => a.id === id && (!month || a.month === month)).reduce((s, a) => s + a.xp, 0); }
  function streakInfo(id) {
    const days = new Set(ledger.filter(a => a.id === id && (a.kind === "clear" || a.kind === "boss-step")).map(a => a.day));
    let d = parseDay(todayIso());
    if (!isWorkday(d)) d = prevWorkday(d);
    else if (!days.has(localDate(d))) d = prevWorkday(d); // today isn't over: the streak is still alive
    let cur = 0;
    while (days.has(localDate(d))) { cur++; d = prevWorkday(d); }
    let best = 0, run = 0, prev = null;
    [...days].filter(x => isWorkday(parseDay(x))).sort().forEach(x => {
      run = prev && localDate(prevWorkday(parseDay(x))) === prev ? run + 1 : 1;
      best = Math.max(best, run); prev = x;
    });
    return { cur, best: Math.max(best, cur) };
  }
  function stats(id) {
    if (statsCache[id]) return statsCache[id];
    const all = ledger.filter(a => a.id === id), mine = clearsOf(all);
    const st = streakInfo(id);
    const s = {
      clears: mine.length, xp: all.reduce((t, a) => t + a.xp, 0), onTime: mine.filter(a => a.onTime).length,
      assists: all.filter(a => a.kind === "assist").length, bounties: mine.filter(a => a.bounty).length, raids: mine.filter(a => a.q.boss).length,
      urgent: mine.filter(a => a.q.tier === "urgent").length, proof: mine.filter(a => a.q.proof && (a.q.proof.tracking || a.q.proof.note)).length,
      squad: mine.filter(a => a.q.squad).length, customs: mine.filter(a => /customs|broker/i.test((a.q.title || "") + " " + (a.q.details || ""))).length,
      streak: st.cur, bestStreak: st.best,
    };
    s.badges = BADGES.filter(b => b.test(s)).map(b => b.key);
    return (statsCache[id] = s);
  }
  function previewXp(q) {
    const mult = isBountyToday(q) ? BOUNTY_MULT : 1;
    const base = tierOf(q).xp * mult, n = Math.max(1, q.squad ? q.assignees.length || 1 : 1);
    return { base: Math.round(base / n), bonus: Math.round(base * (1 + ON_TIME_BONUS) / n) };
  }

  /* ---------- Daily bounty ---------- */
  const isBountyToday = (q) => !!(q && q.bounty === todayIso() && q.status !== "done" && !isPersonal(q));
  const bountyQuest = () => data.bounty && data.bounty.date === todayIso() ? data.quests.find(q => q.id === data.bounty.questId && q.status !== "done") || null : null;
  function pickBounty(doc) {
    const t = todayIso();
    if (doc.bounty && doc.bounty.date === t) return false;
    const pool = doc.quests.filter(q => q.status !== "done" && !q.assignees.length && !q.boss && !q.intel && q.tier !== "lore" && q.createdAt && isoToLocalDay(q.createdAt) < t)
      .sort((a, b) => a.createdAt.localeCompare(b.createdAt));
    const q = pool[0] || null;
    if (q) { q.bounty = t; q.history.push({ at: nowIso(), by: "claude", action: "posted a bounty on it", detail: "2× XP if cleared today" }); }
    doc.bounty = { date: t, questId: q ? q.id : null };
  }
  async function maybeBounty() {
    if (!account || !isWorkday(parseDay(todayIso())) || (data.bounty && data.bounty.date === todayIso())) return;
    await mutate(doc => pickBounty(doc));
  }

  /* ---------- Weekly crew goal ---------- */
  function weekStart(d = parseDay(todayIso())) { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; }
  function goalStatus(offsetWeeks = 0) {
    const g = data.goal; if (!g || !g.active || !g.target) return null;
    const start = weekStart(); start.setDate(start.getDate() - 7 * offsetWeeks);
    const end = new Date(start); end.setDate(end.getDate() + 7);
    const s0 = localDate(start), e0 = localDate(end);
    const cleared = data.quests.filter(q => q.status === "done" && q.doneAt && isoToLocalDay(q.doneAt) >= s0 && isoToLocalDay(q.doneAt) < e0).length;
    const overdueUrgent = offsetWeeks ? 0 : data.quests.filter(q => q.status !== "done" && q.tier === "urgent" && q.due && daysLeft(q.due) < 0).length;
    const daysLeftInWeek = Math.max(0, Math.round((end - parseDay(todayIso())) / 86400000) - 2); // workdays-ish: ignore the weekend
    return { g, cleared, target: g.target, pct: Math.min(100, Math.round(cleared / g.target * 100)), overdueUrgent, met: cleared >= g.target && !overdueUrgent, start, daysLeftInWeek };
  }
  function goalBanner(tv = false) {
    const st = goalStatus();
    if (!st) {
      if (!isAdmin || tv || !loaded) return null;
      return el("div", { class: "goal off" }, el("div", { class: "goal-top" }, el("b", {}, "CREW GOAL"), el("span", { class: "hint" }, "Set a weekly team target and a real reward. Everyone fills one bar together."),
        el("button", { class: "btn small", onclick: () => { goalEdit = true; render(); } }, "SET A GOAL")), goalEdit ? goalForm() : null);
    }
    const last = goalStatus(1);
    return el("div", { class: "goal" + (st.met ? " met" : ""), role: "group", "aria-label": "Weekly crew goal" },
      el("div", { class: "goal-top" },
        el("b", {}, st.met ? "CREW GOAL MET 🎉" : "CREW GOAL · THIS WEEK"),
        el("span", { class: "goal-num" }, `${st.cleared} / ${st.target}`, el("small", {}, " QUESTS")),
        st.g.reward ? el("span", { class: "goal-reward" }, "REWARD: ", el("b", {}, st.g.reward)) : null,
        isAdmin && !tv ? el("button", { class: "btn ghost small", onclick: () => { goalEdit = !goalEdit; render(); } }, goalEdit ? "CLOSE" : "EDIT") : null),
      el("div", { class: "goal-bar", role: "progressbar", "aria-valuemin": "0", "aria-valuemax": String(st.target), "aria-valuenow": String(st.cleared) }, el("i", { style: `width:${st.pct}%` })),
      el("div", { class: "goal-sub" },
        el("span", { class: st.overdueUrgent ? "bad" : "ok" }, st.overdueUrgent ? `✗ ${st.overdueUrgent} OVERDUE URGENT: CLEAR IT TO QUALIFY` : "✓ NO OVERDUE URGENT QUESTS"),
        el("span", {}, st.met ? "Reward unlocked. Nice work, crew." : `${Math.max(0, st.target - st.cleared)} to go`),
        last ? el("span", {}, `LAST WEEK: ${last.cleared}/${last.target}${last.cleared >= last.target ? " ✓" : ""}`) : null),
      goalEdit && isAdmin && !tv ? goalForm() : null);
  }
  function goalForm() {
    const g = data.goal || {};
    if (drafts["g-target"] === undefined) drafts["g-target"] = String(g.target || 25);
    if (drafts["g-reward"] === undefined) drafts["g-reward"] = g.reward || "";
    return el("form", { class: "goal-form", onsubmit: (e) => { e.preventDefault(); saveGoal(true); } },
      el("div", { class: "field" }, el("label", { for: "g-target" }, "QUESTS TO CLEAR PER WEEK"), inputDraft("g-target", { type: "number", min: 1, max: 500 })),
      el("div", { class: "field" }, el("label", { for: "g-reward" }, "REWARD IF THE CREW HITS IT"), inputDraft("g-reward", { maxlength: 80, placeholder: "Pizza Friday" })),
      el("div", { class: "q-btns" }, data.goal && data.goal.active ? el("button", { type: "button", class: "btn ghost small", onclick: () => saveGoal(false) }, "TURN OFF") : null,
        el("button", { type: "submit", class: "btn solid small" }, "SAVE GOAL")));
  }
  async function saveGoal(on) {
    const target = Math.max(1, Math.min(500, parseInt(drafts["g-target"], 10) || 25)), reward = (drafts["g-reward"] || "").trim().slice(0, 80);
    if (await mutate(doc => { doc.goal = on ? { target, reward, active: true, updatedAt: nowIso() } : { ...(doc.goal || {}), active: false }; })) {
      goalEdit = false; clearDrafts("g-target", "g-reward"); render(); toast(on ? "CREW GOAL SAVED" : "CREW GOAL TURNED OFF");
    }
  }

  /* ---------- Assists + shout-outs ---------- */
  function shoutouts(n = 6) {
    return data.quests.flatMap(q => (q.assists || []).map(a => ({ ...a, q }))).sort((a, b) => (b.at || "").localeCompare(a.at || "")).slice(0, n);
  }
  function canAssist(q) {
    if (isPersonal(q) || !joined() || q.compacted) return false;
    return isAdmin || q.assignees.includes(myId) || q.credit.includes(myId);
  }
  function assistBox(q) {
    const taken = new Set((q.assists || []).map(a => a.to));
    const options = data.crew.filter(c => c.id !== myId && !taken.has(c.id));
    if (drafts["as-who-" + q.id] === undefined) drafts["as-who-" + q.id] = options[0] ? options[0].id : "";
    const sel = el("select", { id: "as-who-" + q.id, "aria-label": "Who helped", onchange: (e) => { drafts["as-who-" + q.id] = e.target.value; } },
      options.map(c => el("option", { value: c.id, selected: drafts["as-who-" + q.id] === c.id || null }, c.name)));
    return el("form", { class: "assistbox", onsubmit: (e) => { e.preventDefault(); giveAssist(q.id); } },
      el("h4", {}, "WHO HELPED YOU?"),
      options.length ? el("div", { class: "fields" },
        el("div", { class: "field" }, el("label", { for: "as-who-" + q.id }, "TEAMMATE"), sel),
        el("div", { class: "field" }, el("label", { for: "as-note-" + q.id }, "SHOUT-OUT (OPTIONAL)"), inputDraft("as-note-" + q.id, { maxlength: 120, placeholder: "Saved the day with the forklift" })))
        : el("p", { class: "hint" }, "Everyone on the crew already has an assist on this one."),
      el("span", { class: "hint" }, `They get +${ASSIST_XP} XP and you get +${ASSIST_GIVER_XP} for giving credit. Shout-outs show on the board and the warehouse TV.`),
      el("div", { class: "q-btns" }, el("button", { type: "button", class: "btn ghost small", onclick: () => { assisting = null; render(); } }, "CANCEL"),
        options.length ? el("button", { type: "submit", class: "btn solid small", disabled: busy || null }, "GIVE ASSIST ▸") : null));
  }
  async function giveAssist(id) {
    const to = drafts["as-who-" + id], note = (drafts["as-note-" + id] || "").trim().slice(0, 120);
    if (!to) return;
    if (await withQuest(id, q => {
      q.assists = q.assists || [];
      if (to === myId || q.assists.some(a => a.to === to) || q.assists.length >= MAX_ASSISTS || !canAssist(q)) { toast("That assist can't be added."); return false; }
      q.assists.push({ id: uid("a"), by: myId, to, at: nowIso(), note });
      log(q, "gave an assist", "to " + nameOf(to) + (note ? ": " + note : ""));
    })) { assisting = null; clearDrafts("as-who-" + id, "as-note-" + id); render(); toast(`ASSIST · +${ASSIST_XP} XP TO ${nameOf(to).toUpperCase()}`); }
  }
  function shoutFeed(tv = false) {
    const list = shoutouts(tv ? 4 : 5);
    if (!list.length) return tv ? [] : [el("li", { class: "lb-empty" }, "When someone helps on your quest, give them an assist. Shout-outs show up here.")];
    return list.map(a => el("li", {}, el("b", { style: `color:${colorOf(a.to)}` }, nameOf(a.to)), " helped ", el("b", {}, nameOf(a.by)),
      tv ? null : [" on ", el("span", { class: "muted" }, a.q.title)], a.note ? [" · “", a.note, "”"] : null));
  }

  /* ---------- Boss fights ---------- */
  function bossHp(q) { const total = q.checklist.length, left = q.checklist.filter(c => !c.done).length; return { total, left, pct: total ? Math.round(left / total * 100) : 0 }; }
  function bossBar(q, big = false) {
    const h = bossHp(q), fighters = [...new Set(q.checklist.filter(c => c.done && c.by).map(c => c.by))];
    return el("div", { class: "boss" + (big ? " big" : "") },
      el("div", { class: "boss-top" }, el("b", {}, h.left ? "BOSS HP" : "BOSS DEFEATED"), el("span", {}, `${h.left} / ${h.total}`),
        fighters.length ? el("span", { class: "boss-crew" }, fighters.map(id => el("span", { class: "av", title: nameOf(id), style: `--c:${colorOf(id)}` }, initials(nameOf(id))))) : null),
      el("div", { class: "boss-bar", role: "progressbar", "aria-label": "Boss health", "aria-valuemin": "0", "aria-valuemax": String(h.total), "aria-valuenow": String(h.left) }, el("i", { style: `width:${h.pct}%` })),
      big ? null : el("span", { class: "hint" }, h.left ? `Each step anyone checks off: +${BOSS_STEP_XP} XP. Victory XP is split across everyone who landed a hit.` : "All steps done. Clear it to claim the victory XP."));
  }

  /* ---------- Auth + Graph ---------- */
  async function token() {
    const req = { scopes: SCOPES, account };
    try { return (await msalApp.acquireTokenSilent(req)).accessToken; }
    catch (e) { await msalApp.acquireTokenRedirect(req); throw e; }
  }
  async function graph(path, opts = {}) {
    const t = await token();
    return fetch(path.startsWith("http") ? path : GRAPH + path, { ...opts, headers: { Authorization: "Bearer " + t, ...(opts.headers || {}) } });
  }
  /* Two files: the shared warehouse board in SharePoint (everyone), and the admin's
   * personal list in the admin's own OneDrive (private to them). */
  const store = { shared: { base: null }, personal: { base: null } };
  const enc = (path) => path.split("/").map(encodeURIComponent).join("/");
  async function locate() {
    if (C.DEMO) { siteId = "demo"; itemId = "demo"; store.shared.base = "demo"; if (isAdmin) store.personal.base = "demo"; return; }
    let r = await graph(`/sites/${C.SITE_HOSTNAME}:${C.SITE_PATH}?$select=id`);
    if (!r.ok) throw new Error("site:" + r.status);
    siteId = (await r.json()).id;
    r = await graph(`/sites/${siteId}/drive/root:/${enc(C.DATA_FILE)}?$select=id`);
    if (!r.ok) throw new Error("file:" + r.status);
    itemId = (await r.json()).id;
    store.shared.base = `/sites/${siteId}/drive/items/${itemId}`;
    if (isAdmin && C.PERSONAL_FILE) {
      try {
        const path = enc(C.PERSONAL_FILE);
        r = await graph(`/me/drive/root:/${path}?$select=id`);
        if (r.status === 404) r = await graph(`/me/drive/root:/${path}:/content`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ version: 2, quests: [] }, null, 2) });
        if (r.ok) store.personal.base = `/me/drive/items/${(await r.json()).id}`;
      } catch (e) { /* personal list unavailable; the shared board still works */ }
    }
  }
  async function readFile(kind = "shared") {
    if (C.DEMO) { const s = window.DQB_DEMO.load(kind); return { doc: normalize(JSON.parse(JSON.stringify(s.doc)), kind), etag: s.etag }; }
    const r = await graph(`${store[kind].base}?$select=id,eTag,@microsoft.graph.downloadUrl`);
    if (!r.ok) throw new Error("meta:" + r.status);
    const meta = await r.json();
    const c = await fetch(meta["@microsoft.graph.downloadUrl"], { cache: "no-store" });
    if (!c.ok) throw new Error("download:" + c.status);
    const text = await c.text();
    return { doc: normalize(text.trim() ? JSON.parse(text) : {}, kind), etag: meta.eTag };
  }
  async function writeFile(doc, etag, kind = "shared") {
    if (kind === "personal") doc = { version: 2, updatedAt: doc.updatedAt, quests: doc.quests };
    if (C.DEMO) return window.DQB_DEMO.save(kind, doc, etag);
    const r = await graph(`${store[kind].base}/content`, {
      method: "PUT", headers: { "Content-Type": "application/json", "If-Match": etag }, body: JSON.stringify(doc, null, 2),
    });
    if (r.status === 412 || r.status === 409) return false;
    if (r.status === 403) throw Object.assign(new Error("forbidden"), { code: "forbidden" });
    if (!r.ok) throw new Error("save:" + r.status);
    return true;
  }
  /* Apply a change against the newest copy; retry on conflict. fn returns false to cancel. */
  async function mutate(fn, kind = "shared") {
    if (!store[kind].base) { if (kind === "personal") toast("Your personal list isn't available right now."); return false; }
    if (busy) { toast("Still saving the last change…"); return false; }
    busy = true; render();
    const setDoc = (doc) => { if (kind === "personal") pdata = doc; else data = doc; };
    try {
      for (let attempt = 0; attempt < 5; attempt++) {
        const { doc, etag } = await readFile(kind);
        if (fn(doc) === false) { setDoc(doc); return false; }
        compact(doc);
        doc.updatedAt = nowIso();
        if (await writeFile(doc, etag, kind)) { setDoc(normalize(doc, kind)); return true; }
        await new Promise(r => setTimeout(r, 250 + Math.random() * 500));
      }
      toast("The board is busy. Try again in a moment."); return false;
    } catch (e) {
      toast(e.code === "forbidden" ? "You don't have permission to change the board. Ask your manager." : "Couldn't save. Check your connection and try again.");
      return false;
    } finally { busy = false; render(); }
  }
  function typing() {
    const a = document.activeElement;
    return a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) && $("app").contains(a);
  }
  async function refresh() {
    if (!store.shared.base || busy) return;
    try {
      const { doc } = await readFile("shared"); data = doc; loaded = true;
      if (store.personal.base) { try { pdata = (await readFile("personal")).doc; } catch (e) { /* keep last copy */ } }
      if (typing()) pendingRender = true; else render();
      if (!spawnChecked) { spawnChecked = true; maybeSpawn().then(maybeBounty); }
      else if (data.bounty && data.bounty.date !== todayIso()) maybeBounty();
      else if (!data.bounty) maybeBounty();
    } catch (e) { /* keep last view */ }
  }
  document.addEventListener("focusout", () => setTimeout(() => { if (pendingRender && !typing()) { pendingRender = false; render(); } }, 50));

  /* ---------- Recurring missions ---------- */
  function dueToday(r, d) {
    if (!r.active) return false;
    if (r.cadence === "daily") return isWorkday(d);
    if (r.cadence === "weekly") return d.getDay() === Number(r.weekday);
    if (r.cadence === "monthly") return d.getDate() === Number(r.day);
    return false;
  }
  function spawnInto(doc) {
    const d = parseDay(todayIso()), t = todayIso(); let n = 0;
    doc.recurring.forEach(r => {
      if (dueToday(r, d) && r.lastSpawned !== t) {
        doc.quests.push({ id: uid(), title: r.title, giver: r.giver || "Recurring mission", tier: r.tier, due: t, details: r.details || "",
          source: "Recurring mission", status: "open", assignees: [], credit: [], doneAt: null, createdAt: nowIso(), createdBy: "recurring",
          checklist: (r.checklist || []).map(text => ({ id: uid("c"), text, done: false })), notes: [], history: [{ at: nowIso(), by: "recurring", action: "posted", detail: "Recurring mission" }],
          squad: !!r.squad, boss: !!r.boss, requiresProof: !!r.requiresProof, fromInbox: false, reviewed: true, intel: null, recurringId: r.id });
        r.lastSpawned = t; n++;
      }
    });
    return n;
  }
  async function maybeSpawn() {
    const d = parseDay(todayIso()), t = todayIso();
    if (!data.recurring.some(r => dueToday(r, d) && r.lastSpawned !== t)) return;
    let n = 0;
    if (await mutate(doc => { n = spawnInto(doc); if (!n) return false; })) toast(`${n} RECURRING MISSION${n > 1 ? "S" : ""} POSTED`);
  }

  /* ---------- Render: HUD + leaderboard ---------- */
  function renderSwatches() {
    $("swatches").replaceChildren(...COLORS.map(c => el("button", { type: "button", class: "sw", style: `--c:${c}`, "aria-label": "Color " + c,
      "aria-pressed": String(c === pickColor), onclick: () => { pickColor = c; renderSwatches(); } })));
  }
  function renderHud() {
    const hud = $("hud"), me = member(myId);
    const showJoin = !!account && loaded && (!me || editing);
    $("joinForm").hidden = !showJoin;
    $("hudRow").hidden = !(me && !editing);
    $("badgeRow").hidden = !(me && !editing);
    $("xpwrap").hidden = !me;
    $("adminPanel").hidden = !isAdmin || !loaded;
    $("v-debrief").hidden = !isAdmin; $("v-ops").hidden = !isAdmin;
    $("eyebrow").replaceChildren(...["NOW PLAYING", isAdmin ? el("span", { class: "admin-tag" }, "ADMIN") : null].filter(Boolean));
    $("who").textContent = account ? (myEmail || myName) : "";
    $("signout").hidden = !account || !!C.DEMO;
    if (showJoin) renderSwatches();
    if (!account) { $("pname").textContent = "Sign in to play"; $("prole").textContent = "Use your Wyanoke Microsoft account."; $("lvl").textContent = "–"; return; }
    if (!me) { hud.style.removeProperty("--pc"); $("lvl").textContent = "1"; $("pname").textContent = myName || "New player"; $("prole").textContent = "Join the crew to claim quests and earn XP."; return; }
    hud.style.setProperty("--pc", me.color || COLORS[0]);
    const s = stats(myId), lvl = levelOf(s.xp), into = s.xp % PER_LEVEL;
    $("lvl").textContent = lvl; $("pname").textContent = me.name;
    $("prole").textContent = editing ? "Update your role and color below." : (me.role || "Crew");
    $("xpfill").style.width = (into / PER_LEVEL * 100) + "%";
    $("xpbar").setAttribute("aria-valuenow", Math.round(into / PER_LEVEL * 100));
    $("xptext").textContent = `${into} / ${PER_LEVEL} XP · ${s.xp} total · ${xpFor(myId, monthKey(nowIso()))} this month`;
    $("rank").textContent = "RANK: " + rankOf(lvl);
    $("streak").textContent = s.streak ? `🔥 ${s.streak}-DAY STREAK · BEST ${s.bestStreak}` : "NO STREAK YET · CLEAR A QUEST TODAY";
    $("badges").replaceChildren(...BADGES.map(b => el("span", { class: "badge" + (s.badges.includes(b.key) ? "" : " locked"), title: b.hint }, b.label)));
    if (lastLevel !== null && lvl > lastLevel) toast(`LEVEL UP! LVL ${lvl} · ${rankOf(lvl)}`, "level");
    lastLevel = lvl;
  }
  function ranking(month) {
    return data.crew.map(c => ({ c, xp: xpFor(c.id, month), s: stats(c.id) })).sort((a, b) => b.xp - a.xp || a.c.name.localeCompare(b.c.name));
  }
  function hallOfFame() {
    const cur = monthKey(nowIso());
    const months = [...new Set(ledger.map(a => a.month))].filter(m => m && m < cur).sort().reverse().slice(0, 3);
    return months.map(m => {
      const totals = {}; ledger.filter(a => a.month === m).forEach(a => totals[a.id] = (totals[a.id] || 0) + a.xp);
      const [id, xp] = Object.entries(totals).sort((a, b) => b[1] - a[1])[0] || [];
      return id ? { m, id, xp } : null;
    }).filter(Boolean);
  }
  const shipAlerts = () => (data.shipments || []).filter(s => (s.status === "exception" || s.status === "claim") && !s.ack);
  function renderBoardSide() {
    const alerts = shipAlerts().length;
    $("s-ups").textContent = alerts;
    $("ups-badge").textContent = alerts; $("ups-badge").hidden = !alerts;
    const open = data.quests.filter(q => q.status !== "done");
    $("s-urgent").textContent = open.filter(q => q.tier === "urgent").length;
    $("s-open").textContent = open.length;
    $("s-done").textContent = data.quests.length - open.length;
    $("s-intel").textContent = allQ().filter(q => q.status !== "done" && (q.intel || (q.fromInbox && !q.reviewed))).length;
    $("lb-season").setAttribute("aria-pressed", String(lbMode === "season"));
    $("lb-all").setAttribute("aria-pressed", String(lbMode === "all"));
    const month = lbMode === "season" ? monthKey(nowIso()) : null;
    $("lbTitle").textContent = month ? "SEASON · " + monthName(month).toUpperCase() : "ALL-TIME LEADERBOARD";
    const lb = $("lb"); lb.replaceChildren();
    if (!loaded) lb.append(el("li", { class: "lb-empty" }, account ? "Loading crew…" : "Sign in to see the crew."));
    else if (!data.crew.length) lb.append(el("li", { class: "lb-empty" }, "No crew yet. The first person to join takes the top spot."));
    else ranking(month).forEach(({ c, xp, s }, i) => {
      const remove = isAdmin && c.id !== myId
        ? (confirmP === c.id
            ? el("span", { style: "display:flex;gap:4px" },
                el("button", { class: "btn danger small", onclick: () => removePlayer(c.id) }, "REMOVE"),
                el("button", { class: "btn ghost small", onclick: () => { confirmP = null; render(); } }, "KEEP"))
            : el("button", { class: "btn ghost small", "aria-label": "Remove " + c.name, onclick: () => { confirmP = c.id; render(); } }, "✕"))
        : el("span");
      lb.append(el("li", { class: c.id === myId ? "me" : "" },
        el("span", { class: "pos" }, i + 1), el("span", { class: "av", style: `--c:${c.color}` }, initials(c.name)),
        el("span", { class: "nm", title: c.role || "" }, c.name, el("span", { style: "color:var(--muted)" }, ` · L${levelOf(s.xp)}${s.streak ? " · 🔥" + s.streak : ""}`)),
        el("span", { class: "pts" }, xp + " XP"), remove));
    });
    $("shouts").replaceChildren(...shoutFeed());
    const hof = hallOfFame();
    $("hof").replaceChildren(...(hof.length ? ["HALL OF FAME: ", ...hof.map((h, i) => [i ? " · " : "", monthName(h.m) + " MVP ", el("b", {}, nameOf(h.id)), ` (${h.xp} XP)`])]
      : ["Monthly seasons: the leaderboard resets on the 1st, and each month's MVP joins the hall of fame."]));
  }

  /* ---------- Render: quests ---------- */
  function filtered() {
    const Q = allQ(), open = Q.filter(q => q.status !== "done");
    const pick = {
      done: () => Q.filter(q => q.status === "done").sort((a, b) => (b.doneAt || "").localeCompare(a.doneAt || "")),
      personal: () => open.filter(isPersonal), warehouse: () => open.filter(q => !isPersonal(q)),
      all: () => open, mine: () => open.filter(isMine), open: () => open.filter(q => !isPersonal(q) && !q.assignees.length),
      intel: () => open.filter(q => q.intel || (q.fromInbox && !q.reviewed)),
    }[filter] || (() => open.filter(q => q.tier === filter));
    const list = pick();
    return filter === "done" ? list : list.sort(sortQ);
  }
  function sortQ(a, b) {
    const ta = tierOf(a).order, tb = tierOf(b).order;
    if (ta !== tb) return ta - tb;
    const da = a.due || "9999", dbb = b.due || "9999";
    return da < dbb ? -1 : da > dbb ? 1 : (a.title || "").localeCompare(b.title || "");
  }
  function renderTabs() {
    const Q = allQ(), open = Q.filter(q => q.status !== "done");
    if (!tabsFor().some(([k]) => k === filter)) filter = "all";
    const count = { all: open.length, mine: open.filter(isMine).length, open: open.filter(q => !isPersonal(q) && !q.assignees.length).length,
      intel: open.filter(q => q.intel || (q.fromInbox && !q.reviewed)).length, done: Q.length - open.length,
      personal: open.filter(isPersonal).length, warehouse: open.filter(q => !isPersonal(q)).length };
    for (const t of ["urgent","main","side","lore"]) count[t] = open.filter(q => q.tier === t).length;
    $("tabs").replaceChildren(...tabsFor().map(([key, label]) => el("button", {
      class: "tab" + (key === "intel" ? " intel-tab" : ""), role: "tab", "aria-selected": String(filter === key),
      onclick: () => { filter = key; try { localStorage.setItem("dqb-filter", key); } catch (e) {} render(); },
    }, label, el("span", { class: "n" }, count[key]))));
  }
  function dueChip(q) {
    if (q.status === "done") {
      const a = awards(q)[0];
      return a && a.onTime ? el("span", { class: "chip ok" }, "ON TIME") : null;
    }
    const d = daysLeft(q.due);
    if (d === null) return el("span", { class: "chip" }, "NO DEADLINE");
    if (d < 0) return el("span", { class: "chip overdue" }, `OVERDUE ${-d}D`);
    if (d === 0) return el("span", { class: "chip today" }, "DUE TODAY");
    return el("span", { class: "chip" + (d <= 2 ? " soon" : "") }, `${d}D LEFT · ${fmtDate(q.due).toUpperCase()}`);
  }
  function whoChips(q) {
    if (isPersonal(q)) return [q.status === "done" ? el("span", { class: "chip ok" }, "CLEARED") : null, el("span", { class: "chip new", style: "background:transparent;color:var(--intel)" }, "PERSONAL · ONLY YOU")];
    if (q.status === "done") return q.credit.length
      ? q.credit.map(id => el("span", { class: "chip who", style: `--c:${colorOf(id)}` }, "CLEARED · " + nameOf(id).toUpperCase()))
      : [el("span", { class: "chip ok" }, "CLEARED")];
    if (!q.assignees.length) return [el("span", { class: "chip" }, q.squad ? "SQUAD OPEN" : "UNCLAIMED")];
    return q.assignees.map(id => el("span", { class: "chip who", style: `--c:${colorOf(id)}` }, (id === myId ? "YOU · " : "") + nameOf(id).toUpperCase()));
  }
  function inputDraft(id, attrs = {}) {
    const n = el(attrs.multiline ? "textarea" : "input", { id, ...attrs, multiline: null, value: drafts[id] || "", oninput: (e) => { drafts[id] = e.target.value; } });
    return n;
  }
  function clearDrafts(...ids) { ids.forEach(i => delete drafts[i]); }

  /* ---------- Draft replies ---------- */
  const firstWord = (s) => (s || "").trim().split(/[\s\/(,]+/)[0] || "";
  function greetingFor(q) {
    const f = firstWord(q.giver);
    if (!f || /^(ups|fedex|usps|finance|dhl|the)$/i.test(f) || f === f.toUpperCase()) return "Hello,";
    return `Hi ${f},`;
  }
  function replySubject(q) {
    if (q.replyDraft && q.replyDraft.subject) return q.replyDraft.subject;
    const src = (q.source || "").replace(/^(warehouse inbox|calendar)\s*:\s*/i, "").trim();
    const base = src && !/^posted on the board/i.test(src) && !/^ups tracker$/i.test(src) ? src : q.title;
    return /^re:/i.test(base) ? base : "RE: " + base;
  }
  function replyBody(q) {
    if (q.replyDraft && q.replyDraft.body && q.status !== "done") return q.replyDraft.body;
    const ships = (data.shipments || []).filter(s => s.questId === q.id);
    const lines = [greetingFor(q), ""];
    const by = q.due ? fmtDate(q.due) : null;
    if (q.status === "done") {
      lines.push("Quick update: this is taken care of.");
      if (q.proof && q.proof.tracking) lines.push(`Tracking / confirmation: ${q.proof.tracking}`);
    } else if (q.intel && q.intel.text) {
      lines.push(`Quick update: ${q.intel.text}`);
    } else if (q.assignees.length && !isPersonal(q)) {
      lines.push(`Quick update: we're working on this${by ? ` and are on track for ${by}` : ""}.`);
    } else {
      lines.push(`Thanks for sending this. We've got it${by ? ` and will have it done by ${by}` : ""}.`);
    }
    if (ships.length && q.intel) lines.push(`UPS tracking: ${ships.map(s => s.tracking).join(", ")}`);
    else ships.forEach(s => lines.push(`UPS ${s.tracking}: ${s.lastText || SHIP_ST[s.status].label.toLowerCase()}`));
    lines.push("", "Let me know if you need anything else.", "", "Thanks,", firstWord(myName) || "", "Wyanoke Campus Logistics");
    return lines.join("\n");
  }
  function openReply(q) {
    replying = replying === q.id ? null : q.id;
    if (replying) {
      const k = (f) => `rp-${f}-${q.id}`;
      if (drafts[k("to")] === undefined) drafts[k("to")] = (q.replyDraft && q.replyDraft.to) || q.replyTo || "";
      if (drafts[k("su")] === undefined) drafts[k("su")] = replySubject(q);
      if (drafts[k("bo")] === undefined) drafts[k("bo")] = replyBody(q);
    }
    render();
  }
  async function copyReply(id) {
    const t = drafts["rp-bo-" + id] || "";
    try { await navigator.clipboard.writeText(t); toast("REPLY COPIED. PASTE IT INTO OUTLOOK"); } catch (e) { toast("SELECT THE TEXT AND COPY IT"); }
  }
  function outlookLink(id) {
    const enc = encodeURIComponent;
    return "https://outlook.office.com/mail/deeplink/compose?to=" + enc(drafts["rp-to-" + id] || "") + "&subject=" + enc(drafts["rp-su-" + id] || "") + "&body=" + enc(drafts["rp-bo-" + id] || "");
  }
  /* mailto: opens the computer's default mail app (desktop Outlook when it's the default).
   * Windows caps these links at about 2,000 characters, so a long message goes via the clipboard. */
  const MAILTO_MAX = 1900;
  async function openDesktop(id) {
    const enc = (t) => encodeURIComponent(t).replace(/%0A/g, "%0D%0A");
    const to = (drafts["rp-to-" + id] || "").split(/[;,]\s*/).filter(Boolean).join(";"), su = drafts["rp-su-" + id] || "", bo = drafts["rp-bo-" + id] || "";
    const head = "mailto:" + encodeURIComponent(to).replace(/%3B/g, ";").replace(/%40/g, "@") + "?subject=" + enc(su);
    let url = head + "&body=" + enc(bo);
    if (url.length > MAILTO_MAX) {
      url = head;
      try { await navigator.clipboard.writeText(bo); toast("MESSAGE COPIED. PASTE IT INTO THE OUTLOOK WINDOW"); } catch (e) { toast("MESSAGE TOO LONG FOR A LINK. COPY IT FIRST, THEN PASTE"); }
    }
    window.location.href = url;
  }
  function replyBox(q) {
    const k = (f) => `rp-${f}-${q.id}`;
    return el("div", { class: "replybox" },
      el("h4", {}, "DRAFT REPLY", q.replyDraft && q.replyDraft.body && q.status !== "done" ? el("span", { class: "chip new", style: "margin-left:8px" }, "WRITTEN BY CLAUDE FROM THE THREAD") : null),
      el("div", { class: "fields" },
        el("div", { class: "field" }, el("label", { for: k("to") }, "TO"), inputDraft(k("to"), { maxlength: 300, placeholder: "name@company.com (optional)" })),
        el("div", { class: "field" }, el("label", { for: k("su") }, "SUBJECT"), inputDraft(k("su"), { maxlength: 200 }))),
      el("div", { class: "field" }, el("label", { for: k("bo") }, "MESSAGE"), inputDraft(k("bo"), { multiline: true, rows: 9, maxlength: 4000 })),
      el("span", { class: "hint" }, "Nothing is sent from here. Open in Outlook starts a new message in the Outlook app on this computer. To keep it in the same thread, open the original email, click Reply and paste."),
      el("div", { class: "q-btns" },
        el("button", { type: "button", class: "btn ghost small", onclick: () => { [k("to"), k("su"), k("bo")].forEach(i => delete drafts[i]); replying = null; render(); } }, "CLOSE"),
        el("button", { type: "button", class: "btn ghost small", title: "Start over from the current quest status", onclick: () => { [k("to"), k("su"), k("bo")].forEach(i => delete drafts[i]); replying = null; openReply(q); } }, "RESET"),
        q.sourceLink ? el("a", { class: "btn ghost small", href: q.sourceLink, target: "_blank", rel: "noopener" }, "OPEN ORIGINAL EMAIL ↗") : null,
        el("button", { type: "button", class: "btn small", onclick: () => copyReply(q.id) }, "COPY REPLY"),
        el("a", { class: "btn ghost small", href: "#", target: "_blank", rel: "noopener", title: "Outlook in the browser", onclick: (e) => { e.currentTarget.href = outlookLink(q.id); } }, "OUTLOOK WEB ↗"),
        el("button", { type: "button", class: "btn solid small", onclick: () => openDesktop(q.id) }, "OPEN IN OUTLOOK ▸")));
  }

  function card(q) {
    const t = tierOf(q), done = q.status === "done", isOpen = openSet.has(q.id);
    const lock = !joined() || busy, mine = isMine(q);
    const takenByOther = !q.squad && q.assignees.length && !mine;
    const clDone = q.checklist.filter(c => c.done).length;
    const btns = [];
    if (done) { if (isAdmin || q.credit.includes(myId)) btns.push(el("button", { class: "btn ghost", disabled: lock || null, onclick: () => act(q.id, "reopen") }, "REOPEN")); }
    else if (completing !== q.id && isPersonal(q)) {
      btns.push(el("button", { class: "btn", disabled: busy || null, onclick: () => { completing = q.id; render(); setTimeout(() => { const f = $("pf-t-" + q.id); f && f.focus(); }, 30); } }, "COMPLETE ▸"));
    }
    else if (completing !== q.id && q.boss) {
      const h = bossHp(q);
      if (!h.left || isAdmin) btns.push(el("button", { class: "btn" + (h.left ? "" : " solid"), disabled: lock || null, onclick: () => { completing = q.id; render(); } }, h.left ? "FORCE CLEAR ▸" : "CLAIM VICTORY ▸"));
      if (h.left) btns.push(el("button", { class: "btn ghost", onclick: () => { openSet.add(q.id); render(); setTimeout(() => { const d = $("d-" + q.id); d && d.scrollIntoView({ block: "nearest", behavior: "smooth" }); }, 30); } }, "⚔ ATTACK"));
    }
    else if (completing !== q.id) {
      if (q.squad) btns.push(mine ? el("button", { class: "btn ghost", disabled: lock || null, onclick: () => act(q.id, "drop") }, "LEAVE SQUAD")
                                  : el("button", { class: "btn ghost", disabled: lock || null, onclick: () => act(q.id, "claim") }, "JOIN SQUAD"));
      else if (!q.assignees.length) btns.push(el("button", { class: "btn ghost", disabled: lock || null, onclick: () => act(q.id, "claim") }, "CLAIM"));
      else if (mine) btns.push(el("button", { class: "btn ghost", disabled: lock || null, onclick: () => act(q.id, "drop") }, "DROP"));
      if (!takenByOther || isAdmin) btns.push(el("button", { class: "btn", disabled: lock || null, onclick: () => { completing = q.id; render(); setTimeout(() => { const f = $("pf-t-" + q.id); f && f.focus(); }, 30); } }, "COMPLETE ▸"));
    }
    if (canAssist(q)) btns.push(el("button", { class: "btn ghost" + (assisting === q.id ? " on" : ""), "aria-expanded": String(assisting === q.id), onclick: () => { assisting = assisting === q.id ? null : q.id; render(); } }, "🤝 ASSIST"));
    btns.push(el("button", { class: "btn ghost" + (replying === q.id ? " on" : ""), "aria-expanded": String(replying === q.id), onclick: () => openReply(q) }, "✉ DRAFT REPLY"));
    const px = previewXp(q);
    const reward = isPersonal(q) ? el("span", { class: "reward", style: "color:var(--intel)" }, "PRIVATE", el("small", { style: "color:var(--muted)" }, "NOT ON LEADERBOARD"))
      : done
      ? el("span", { class: "reward" }, `+${clearsOf(awards(q)).reduce((s, a) => s + a.xp, 0)} XP`, q.credit.length > 1 ? el("small", {}, `SPLIT ${q.credit.length} WAYS`) : null)
      : q.boss ? el("span", { class: "reward" }, `+${BOSS_STEP_XP} / STEP`, el("small", {}, `+${tierOf(q).xp} VICTORY, SPLIT`))
      : el("span", { class: "reward" + (isBountyToday(q) ? " bounty" : "") }, `+${px.base} XP`, isBountyToday(q) ? el("small", {}, "2× BOUNTY TODAY") : null, q.due ? el("small", {}, `+${px.bonus} IF ON TIME`) : null, q.squad && q.assignees.length > 1 ? el("small", {}, `EACH · ${q.assignees.length} IN SQUAD`) : null);

    const meta = el("div", { class: "q-meta" },
      el("span", { class: "chip tier" }, t.label), dueChip(q), whoChips(q),
      q.fromInbox && !q.reviewed ? el("span", { class: "chip new" }, "NEW FROM INBOX") : null,
      q.squad ? el("span", { class: "chip squad" }, "SQUAD") : null,
      q.boss ? el("span", { class: "chip bosschip" }, "BOSS FIGHT") : null,
      isBountyToday(q) ? el("span", { class: "chip bountychip" }, "💰 BOUNTY · 2× XP TODAY") : null,
      ...(q.assists || []).map(a => el("span", { class: "chip who", style: `--c:${colorOf(a.to)}`, title: a.note || "" }, "🤝 " + nameOf(a.to).toUpperCase())),
      q.checklist.length ? el("span", { class: "chip" + (clDone === q.checklist.length ? " ok" : "") }, `☑ ${clDone}/${q.checklist.length}`) : null,
      q.requiresProof && !done ? el("span", { class: "chip" }, "NEEDS PROOF") : null,
      q.notes.length ? el("span", { class: "chip" }, `✎ ${q.notes.length}`) : null,
      ...(data.shipments || []).filter(s => s.questId === q.id).slice(0, 2).map(s => el("span", { class: "chip st", style: `--st:${s.status === "delivered" ? "var(--done)" : s.status === "exception" || s.status === "claim" ? "var(--urgent)" : "var(--side)"}` }, "UPS · " + SHIP_ST[s.status].label)));

    const intel = !done && q.intel ? el("div", { class: "intel-bar" },
      el("p", {}, el("b", {}, "INTEL"), q.intel.text, q.intel.link ? [" ", el("a", { href: q.intel.link, target: "_blank", rel: "noopener" }, "Open email ↗")] : null),
      el("div", { class: "q-btns" },
        joined() ? el("button", { class: "btn intel small", disabled: busy || null, onclick: () => confirmIntel(q.id) }, "CONFIRM & CLEAR") : null,
        isAdmin ? el("button", { class: "btn ghost small", disabled: busy || null, onclick: () => dismissIntel(q.id) }, "DISMISS") : null)) : null;

    const proofForm = !done && completing === q.id ? el("form", { class: "proofbox", onsubmit: (e) => { e.preventDefault(); complete(q.id); } },
      el("h4", {}, "MISSION DEBRIEF"),
      el("div", { class: "fields" },
        el("div", { class: "field" }, el("label", { for: "pf-t-" + q.id }, q.requiresProof ? "TRACKING # / CONFIRMATION (REQUIRED)" : "TRACKING # / CONFIRMATION (OPTIONAL)"), inputDraft("pf-t-" + q.id, { maxlength: 80, placeholder: "1Z18X416…" })),
        el("div", { class: "field" }, el("label", { for: "pf-n-" + q.id }, "NOTE"), inputDraft("pf-n-" + q.id, { maxlength: 300, placeholder: "Left on dock 2:10 PM, Maura notified" }))),
      q.checklist.length && clDone < q.checklist.length ? el("span", { class: "hint" }, `${q.checklist.length - clDone} checklist step(s) still open. You can clear anyway.`) : null,
      el("div", { class: "form-row" }, el("span", { class: "err", id: "pf-e-" + q.id }),
        el("div", { class: "q-btns" }, el("button", { type: "button", class: "btn ghost", onclick: () => { completing = null; render(); } }, "CANCEL"), el("button", { type: "submit", class: "btn solid", disabled: busy || null }, "CONFIRM CLEAR ▸")))) : null;

    let drawer = null;
    if (isOpen) {
      const editCl = canEditQuest(q) && !done && joined();
      drawer = el("div", { class: "drawer", id: "d-" + q.id },
        el("div", { class: "dsec" }, el("span", { class: "flabel" }, "BRIEFING"), el("p", {}, q.details || "No briefing."),
          el("div", { class: "q-source" }, q.source ? "SOURCE: " + q.source : null, q.sourceLink ? el("a", { href: q.sourceLink, target: "_blank", rel: "noopener" }, "OPEN EMAIL ↗") : null)),
        q.proof ? el("div", { class: "dsec" }, el("span", { class: "flabel" }, "PROOF"),
          el("div", { class: "proof" }, q.proof.tracking ? [el("b", {}, "TRACKING"), q.proof.tracking, " "] : null, q.proof.note ? [el("b", {}, "NOTE"), q.proof.note] : null,
            el("div", { class: "hint" }, `Logged by ${nameOf(q.proof.by)} · ${fmtWhen(q.proof.at)}`))) : null,
        (q.checklist.length || editCl) ? el("div", { class: "dsec" }, el("span", { class: "flabel" }, "CHECKLIST"),
          el("ul", { class: "cl" }, q.checklist.map(c => el("li", { class: c.done ? "done" : "" },
            el("label", {}, el("input", { type: "checkbox", checked: c.done || null, disabled: !joined() || done || busy || (q.boss && c.done && c.by !== myId && !isAdmin) || null, onchange: () => toggleStep(q.id, c.id) }), el("span", {}, c.text)),
            el("span", { class: "x" }, c.done && c.by ? nameOf(c.by) : "", editCl && isAdmin ? el("button", { class: "btn ghost small", style: "margin-left:6px", "aria-label": "Remove step", onclick: () => removeStep(q.id, c.id) }, "✕") : null)))),
          editCl ? el("form", { class: "inline", onsubmit: (e) => { e.preventDefault(); addStep(q.id); } }, inputDraft("cl-" + q.id, { maxlength: 120, placeholder: "Add a step" }), el("button", { class: "btn ghost small", type: "submit" }, "ADD")) : null) : null,
        el("div", { class: "dsec" }, el("span", { class: "flabel" }, `NOTES (${q.notes.length})`),
          el("div", { class: "notes" }, q.notes.map(n => el("div", { class: "note" }, el("div", { class: "meta" }, `${nameOf(n.by).toUpperCase()} · ${fmtWhen(n.at)}`), el("p", {}, n.text)))),
          joined() && !q.compacted ? el("form", { class: "inline", onsubmit: (e) => { e.preventDefault(); addNote(q.id); } }, inputDraft("note-" + q.id, { maxlength: 400, placeholder: "Add an update (e.g. waiting on Tara for flyers)" }), el("button", { class: "btn ghost small", type: "submit" }, "POST")) : null),
        q.history.length ? el("div", { class: "dsec" }, el("span", { class: "flabel" }, "ACTIVITY"),
          el("ul", { class: "hist" }, q.history.slice(-15).reverse().map(h => el("li", {}, el("time", {}, fmtWhen(h.at)), `${nameOf(h.by)} ${h.action}${h.detail ? " · " + h.detail : ""}`)))) : null);
    }

    let admin = null;
    if (isAdmin) {
      const sel = el("select", { "aria-label": "Assign quest", onchange: (e) => { if (e.target.value !== "-") assign(q.id, e.target.value); } },
        el("option", { value: "-" }, q.squad ? "Add to squad…" : "Assign to…"), el("option", { value: "" }, "Unassigned (clear)"),
        data.crew.map(c => el("option", { value: c.id }, c.name)));
      admin = el("div", { class: "admin-row" }, el("span", { class: "lbl" }, "ADMIN"), !done && !isPersonal(q) ? sel : null,
        el("button", { class: "btn ghost small", disabled: busy || null, onclick: () => moveQuest(q.id, isPersonal(q) ? "shared" : "personal") }, isPersonal(q) ? "MOVE TO WAREHOUSE BOARD" : "MOVE TO MY PERSONAL LIST"),
        q.fromInbox && !q.reviewed ? el("button", { class: "btn intel small", onclick: () => markReviewed(q.id) }, "MARK REVIEWED") : null,
        confirmQ === q.id
          ? [el("span", { class: "err" }, "Delete this quest for good?"), el("button", { class: "btn danger small", onclick: () => deleteQuest(q.id) }, "DELETE"), el("button", { class: "btn ghost small", onclick: () => { confirmQ = null; render(); } }, "KEEP")]
          : el("button", { class: "btn ghost small", onclick: () => { confirmQ = q.id; render(); } }, "DELETE QUEST"));
    }
    return el("article", { class: "quest" + (done ? " is-done" : "") + (q.boss && !done ? " is-boss" : "") + (isBountyToday(q) ? " is-bounty" : ""), "data-tier": q.tier, id: "q-" + q.id },
      el("div", { class: "q-main" }, meta, el("h3", { class: "q-title" }, q.title), el("div", { class: "q-giver" }, "Quest giver: ", el("b", {}, q.giver || "—")), q.boss && !done ? bossBar(q) : null),
      el("div", { class: "q-side" }, reward, el("div", { class: "q-btns" }, btns)),
      intel, proofForm, assisting === q.id ? assistBox(q) : null, replying === q.id ? replyBox(q) : null,
      el("button", { class: "toggle", "aria-expanded": String(isOpen), "aria-controls": "d-" + q.id, onclick: () => { isOpen ? openSet.delete(q.id) : openSet.add(q.id); render(); } }, isOpen ? "▾ HIDE BRIEFING" : q.boss && !done ? "▸ BOSS CHECKLIST · BRIEFING · NOTES" : "▸ MISSION BRIEFING · NOTES · CHECKLIST"),
      drawer, admin);
  }
  function renderList() {
    const list = $("list"); list.replaceChildren();
    if (!account) { list.append(el("div", { class: "empty" }, el("h2", {}, "Sign in to open the board"), el("p", {}, "Use the same Microsoft account you use for Outlook."), el("button", { class: "btn solid", onclick: signIn }, "SIGN IN WITH MICROSOFT ▸"))); return; }
    if (!loaded) { list.append(el("div", { class: "empty" }, el("h2", {}, "Loading the quest board…"), el("p", {}, "Fetching quests from SharePoint."))); return; }
    if (!joined()) list.append(el("div", { class: "empty" }, el("h2", {}, "Join the crew to play"), el("p", {}, "Pick your role and color in the panel above, then claim your first quest.")));
    if (!allQ().length) { list.append(el("div", { class: "empty" }, el("h2", {}, "No quests on the board"), el("p", {}, isAdmin ? "Post the first quest with the admin form above." : "New quests appear here when your manager posts them."))); return; }
    const banners = [goalBanner()];
    const bq = bountyQuest();
    if (bq) banners.push(el("div", { class: "bountybar" }, el("b", {}, "💰 TODAY'S BOUNTY"), el("span", {}, bq.title), el("span", { class: "hint" }, `2× XP (+${tierOf(bq).xp * BOUNTY_MULT}) if it's cleared today`),
      el("button", { class: "btn small", onclick: () => { filter = "all"; openSet.add(bq.id); render(); setTimeout(() => { const n = $("q-" + bq.id); n && n.scrollIntoView({ block: "center", behavior: "smooth" }); }, 40); } }, "SHOW ME")));
    banners.filter(Boolean).forEach(b => list.append(b));
    const shown = filtered();
    if (!shown.length) {
      const msg = { mine: ["No claimed quests", "Claim a quest and it shows up here."], done: ["Nothing cleared yet", "Complete a quest and it shows up here with its XP."],
        intel: ["No new intel", "Quests Claude adds from the inbox, and delivery confirmations it spots, show up here."],
        personal: ["Personal list clear", "Quests from your own inbox show up here. Only you can see them."] }[filter] || ["Board clear", "No open quests here. Nice work."];
      list.append(el("div", { class: "empty" }, el("h2", {}, msg[0]), el("p", {}, msg[1]))); return;
    }
    if (filter === "all") {
      let cur = null;
      for (const q of shown) { if (q.tier !== cur) { cur = q.tier; list.append(el("div", { class: "section-label" }, tierOf(q).label)); } list.append(card(q)); }
    } else shown.forEach(q => list.append(card(q)));
  }

  /* ---------- Render: debrief ---------- */
  function period() {
    if (dbMode === "month") { const m = monthKey(nowIso()); return { label: monthName(m), inP: (iso) => iso && monthKey(iso) === m }; }
    const since = Date.now() - 7 * 86400000;
    return { label: "the last 7 days", inP: (iso) => iso && Date.parse(iso) >= since };
  }
  function renderDebrief() {
    $("db-week").setAttribute("aria-pressed", String(dbMode === "week"));
    $("db-month").setAttribute("aria-pressed", String(dbMode === "month"));
    const P = period();
    const cleared = data.quests.filter(q => q.status === "done" && P.inP(q.doneAt));
    const aw = ledger.filter(a => P.inP(a.at));
    const xp = aw.reduce((s, a) => s + a.xp, 0);
    const withDue = cleared.filter(q => q.due);
    const onTimeRate = withDue.length ? Math.round(withDue.filter(q => isoToLocalDay(q.doneAt) <= q.due).length / withDue.length * 100) : null;
    const hrs = cleared.filter(q => q.createdAt).map(q => (Date.parse(q.doneAt) - Date.parse(q.createdAt)) / 3600000).filter(h => h >= 0);
    const avg = hrs.length ? hrs.reduce((a, b) => a + b, 0) / hrs.length : null;
    const avgTxt = avg === null ? "—" : avg < 24 ? `${Math.round(avg)} h` : `${(avg / 24).toFixed(1)} d`;
    const open = data.quests.filter(q => q.status !== "done");
    const overdue = open.filter(q => q.due && daysLeft(q.due) < 0);
    const fromInbox = data.quests.filter(q => q.fromInbox && P.inP(q.createdAt)).length;
    const upsExc = (data.shipments || []).filter(s => s.events.some(e => e.type === "exception" && P.inP(e.at))).length;
    const kp = [["QUESTS CLEARED", cleared.length], ["XP EARNED", xp], ["ON-TIME RATE", onTimeRate === null ? "—" : onTimeRate + "%"], ["AVG TIME TO CLEAR", avgTxt], ["OPEN NOW", open.length], ["OVERDUE NOW", overdue.length], ["ADDED FROM INBOX", fromInbox], ["UPS EXCEPTIONS", upsExc]];
    $("kpis").replaceChildren(...kp.map(([l, v]) => el("div", { class: "kpi" }, el("b", { style: l === "OVERDUE NOW" && v ? "color:var(--urgent)" : "" }, v), el("span", {}, l))));
    const rows = data.crew.map(c => {
      const mine = aw.filter(a => a.id === c.id);
      const cl = clearsOf(mine);
      return { c, n: cl.length, as: mine.filter(a => a.kind === "assist").length, xp: mine.reduce((s, a) => s + a.xp, 0), on: cl.length ? Math.round(cl.filter(a => a.onTime).length / cl.length * 100) : null, s: stats(c.id),
        open: open.filter(q => q.assignees.includes(c.id)).length };
    }).sort((a, b) => b.xp - a.xp);
    $("perf").replaceChildren(rows.length ? el("table", {}, el("thead", {}, el("tr", {}, ["PLAYER", "ROLE", "CLEARED", "ASSISTS", "XP", "ON TIME", "STREAK", "CLAIMED NOW", "LEVEL"].map((h, i) => el("th", { class: i >= 2 ? "num" : "" }, h)))),
      el("tbody", {}, rows.map(r => el("tr", {}, el("td", {}, r.c.name), el("td", {}, r.c.role || "—"), el("td", { class: "num" }, r.n), el("td", { class: "num" }, r.as), el("td", { class: "num" }, r.xp),
        el("td", { class: "num" }, r.on === null ? "—" : r.on + "%"), el("td", { class: "num" }, r.s.streak), el("td", { class: "num" }, r.open), el("td", { class: "num" }, levelOf(r.s.xp))))))
      : el("p", { class: "hint" }, "No crew yet."));
    const openSorted = [...open].sort((a, b) => (a.due || "9999").localeCompare(b.due || "9999"));
    $("openTable").replaceChildren(openSorted.length ? el("table", {}, el("thead", {}, el("tr", {}, ["QUEST", "TIER", "DUE", "STATUS", "WHO"].map(h => el("th", {}, h)))),
      el("tbody", {}, openSorted.map(q => { const d = daysLeft(q.due); return el("tr", {}, el("td", { style: "white-space:normal;min-width:220px" }, q.title), el("td", {}, tierOf(q).label),
        el("td", {}, q.due ? fmtDate(q.due) : "—"), el("td", { style: d !== null && d < 0 ? "color:var(--urgent)" : "" }, d === null ? "No deadline" : d < 0 ? `Overdue ${-d}d` : d === 0 ? "Due today" : `${d}d left`),
        el("td", {}, q.assignees.length ? q.assignees.map(nameOf).join(", ") : "Unclaimed")); })))
      : el("p", { class: "hint" }, "Nothing open."));
    const top = rows.filter(r => r.n).slice(0, 3).map(r => `${r.c.name} (${r.n} cleared, ${r.xp} XP)`);
    const lines = [
      `Warehouse quest board, ${P.label}:`,
      `- ${cleared.length} quests cleared, ${xp} XP earned${onTimeRate === null ? "" : `, ${onTimeRate}% on time`}.`,
      avg === null ? null : `- Average time from posting to clearing: ${avgTxt}.`,
      `- ${open.length} still open${overdue.length ? `, ${overdue.length} overdue: ${overdue.map(q => q.title).join("; ")}` : ", none overdue"}.`,
      fromInbox ? `- ${fromInbox} quests were added automatically from the inbox.` : null,
      upsExc ? `- ${upsExc} UPS shipment${upsExc > 1 ? "s" : ""} had an exception; ${shipAlerts().length} still need${shipAlerts().length === 1 ? "s" : ""} attention.` : null,
      goalStatus() ? `- Crew goal this week: ${goalStatus().cleared} of ${goalStatus().target} quests${goalStatus().met ? " (met!)" : ""}${data.goal.reward ? `, reward: ${data.goal.reward}` : ""}.` : null,
      (() => { const b = aw.filter(a => a.kind === "clear" && a.bounty).map(a => a.q.id); return b.length ? `- Daily bounties cleared: ${new Set(b).size}.` : null; })(),
      (() => { const n = aw.filter(a => a.kind === "assist").length; return n ? `- Assists given: ${n}.` : null; })(),
      (() => { const n = new Set(aw.filter(a => a.kind === "clear" && a.q.boss).map(a => a.q.id)).size; return n ? `- Boss fights defeated: ${n}.` : null; })(),
      top.length ? `- Top performers: ${top.join(", ")}.` : null,
    ].filter(Boolean);
    $("summary").textContent = lines.join("\n");
  }
  $("copySum").onclick = async () => {
    const text = $("summary").textContent;
    try { await navigator.clipboard.writeText(text); toast("SUMMARY COPIED"); }
    catch (e) { const r = document.createRange(); r.selectNodeContents($("summary")); const s = getSelection(); s.removeAllRanges(); s.addRange(r); toast("Press Ctrl+C to copy"); }
  };
  $("db-week").onclick = () => { dbMode = "week"; render(); };
  $("db-month").onclick = () => { dbMode = "month"; render(); };

  /* ---------- Render: recurring ---------- */
  const DOW = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  function cadenceText(r) { return r.cadence === "daily" ? "Every workday" : r.cadence === "weekly" ? "Every " + DOW[r.weekday] : `Monthly on day ${r.day}`; }
  function renderOps() {
    $("rList").replaceChildren(...(data.recurring.length ? data.recurring.map(r => el("div", { class: "trow" },
      el("span", {}, el("b", {}, r.title), el("span", { class: "hint" }, ` · ${cadenceText(r)} · ${tierOf(r).label}${r.lastSpawned ? " · last posted " + fmtDate(r.lastSpawned) : ""}${r.active ? "" : " · PAUSED"}`)),
      el("div", { class: "q-btns" },
        el("button", { class: "btn ghost small", disabled: busy || null, onclick: () => toggleRecurring(r.id) }, r.active ? "PAUSE" : "RESUME"),
        confirmR === r.id ? [el("button", { class: "btn danger small", onclick: () => deleteRecurring(r.id) }, "DELETE"), el("button", { class: "btn ghost small", onclick: () => { confirmR = null; render(); } }, "KEEP")]
          : el("button", { class: "btn ghost small", onclick: () => { confirmR = r.id; render(); } }, "DELETE"))))
      : [el("p", { class: "hint" }, "No recurring missions yet. Add a daily dock sweep, weekly supply count or monthly maintenance check above.")]));
  }
  const syncCadence = () => { const c = $("r-cadence").value; $("r-weekday-f").hidden = c !== "weekly"; $("r-day-f").hidden = c !== "monthly"; };
  $("r-cadence").onchange = syncCadence; syncCadence();
  function syncRepeat() {
    const personal = $("q-board").value === "personal";
    $("q-repeat-f").hidden = personal; $("q-squad-l").hidden = personal; $("q-boss-l").hidden = personal; if (personal) $("q-repeat").value = "";
    const c = $("q-repeat").value; $("q-weekday-f").hidden = c !== "weekly"; $("q-day-f").hidden = c !== "monthly"; $("q-repeat-hint").hidden = !c; }
  $("q-repeat").onchange = syncRepeat; $("q-board").onchange = syncRepeat; syncRepeat();

  /* ---------- Render: UPS tracker ---------- */
  const upsLink = (t) => "https://www.ups.com/track?loc=en_US&tracknum=" + encodeURIComponent(t);
  function shipGroup(s) { return (s.status === "exception" || s.status === "claim") && !s.ack ? "alerts" : s.status === "delivered" ? "delivered" : "moving"; }
  function renderUps() {
    const all = [...(data.shipments || [])];
    const today = todayIso(), week = Date.now() - 7 * 86400000;
    const kp = [["NEEDS ATTENTION", shipAlerts().length, "var(--urgent)"], ["ON THE WAY", all.filter(s => shipGroup(s) === "moving").length, ""],
      ["DELIVERED TODAY", all.filter(s => s.status === "delivered" && isoToLocalDay(s.lastAt) === today).length, "var(--done)"],
      ["DELIVERED · 7 DAYS", all.filter(s => s.status === "delivered" && Date.parse(s.lastAt) >= week).length, ""]];
    $("ups-kpis").replaceChildren(...kp.map(([l, v, c]) => el("div", { class: "kpi" }, el("b", { style: c && v ? `color:${c}` : "" }, v), el("span", {}, l))));
    const counts = { alerts: 0, moving: 0, delivered: 0, all: all.length };
    all.forEach(s => counts[shipGroup(s)]++);
    $("ups-tabs").replaceChildren(...UPS_TABS.map(([k, l]) => el("button", { class: "tab", role: "tab", "aria-selected": String(upsFilter === k), onclick: () => { upsFilter = k; render(); } }, l, el("span", { class: "n" }, counts[k]))));
    const q = upsQuery.trim().toLowerCase();
    const shown = all.filter(s => upsFilter === "all" || shipGroup(s) === upsFilter)
      .filter(s => !q || [s.tracking, s.ref, s.shipTo, s.lastText, s.service].join(" ").toLowerCase().includes(q))
      .sort((x, y) => (SHIP_ST[x.status].order - SHIP_ST[y.status].order) || (y.lastAt || "").localeCompare(x.lastAt || ""));
    const list = $("ups-list"); list.replaceChildren();
    if (!loaded) { list.append(el("div", { class: "empty" }, el("h2", {}, "Loading shipments…"))); return; }
    if (!shown.length) {
      const msg = upsFilter === "alerts" ? ["No UPS problems", "Exceptions, delivery problems and open claims show up here as soon as Claude spots them in the Warehouse mailbox."]
        : q ? ["No matches", "Try a different tracking number, job number or destination."] : ["Nothing here yet", "UPS notifications from the Warehouse mailbox show up here, grouped by tracking number."];
      list.append(el("div", { class: "empty" }, el("h2", {}, msg[0]), el("p", {}, msg[1]))); return;
    }
    shown.forEach(s => list.append(shipCard(s)));
  }
  function shipCard(s) {
    const st = SHIP_ST[s.status], open = upsOpen.has(s.tracking), quest = s.questId && allQ().find(x => x.id === s.questId);
    const last = s.events[s.events.length - 1];
    const actions = [
      el("a", { class: "btn ghost small", href: upsLink(s.tracking), target: "_blank", rel: "noopener" }, "TRACK ON UPS ↗"),
      last && last.link ? el("a", { class: "btn ghost small", href: last.link, target: "_blank", rel: "noopener" }, "OPEN EMAIL ↗") : null,
    ];
    if ((s.status === "exception" || s.status === "claim") && joined()) actions.push(el("button", { class: "btn small" + (s.ack ? " ghost" : ""), disabled: busy || null, onclick: () => ackShip(s.tracking, !s.ack) }, s.ack ? "UNDO ACK" : "ACKNOWLEDGE"));
    if (quest) actions.push(el("button", { class: "btn intel small", onclick: () => { view = "board"; filter = "all"; openSet.add(quest.id); render(); } }, "VIEW QUEST"));
    else if (isAdmin && s.status !== "delivered") actions.push(el("button", { class: "btn small", disabled: busy || null, onclick: () => questFromShip(s.tracking) }, "CREATE QUEST"));
    return el("article", { class: "ship" + (s.ack ? " acked" : ""), "data-st": s.status },
      el("div", { class: "ship-main" },
        el("div", { class: "ship-top" },
          el("span", { class: "chip st" }, st.label),
          el("button", { class: "trk", title: "Copy tracking number", onclick: () => copyText(s.tracking) }, s.tracking),
          s.ref ? el("span", { class: "chip" }, "REF · " + s.ref) : null,
          s.direction ? el("span", { class: "chip" }, s.direction.toUpperCase()) : null,
          s.service ? el("span", { class: "chip" }, s.service.toUpperCase()) : null,
          s.ack ? el("span", { class: "chip ok" }, "ACKNOWLEDGED" + (s.ackBy ? " · " + nameOf(s.ackBy).toUpperCase() : "")) : null),
        el("div", { class: "ship-to" }, s.shipTo || "Destination not listed"),
        el("div", { class: "ship-last" }, s.lastAt ? el("time", {}, fmtWhen(s.lastAt)) : null, s.lastText || "")),
      el("div", { class: "ship-side" }, el("div", { class: "q-btns" }, actions)),
      el("button", { class: "toggle", "aria-expanded": String(open), onclick: () => { open ? upsOpen.delete(s.tracking) : upsOpen.add(s.tracking); render(); } }, open ? "▾ HIDE HISTORY" : `▸ HISTORY (${s.events.length})`),
      open ? el("ol", { class: "tl" }, [...s.events].reverse().map(e => el("li", { class: e.type || "" }, el("time", {}, fmtWhen(e.at)),
        el("span", {}, e.text || "", e.link ? [" ", el("a", { href: e.link, target: "_blank", rel: "noopener" }, "email ↗")] : null)))) : null);
  }
  async function copyText(t) {
    try { await navigator.clipboard.writeText(t); toast("TRACKING # COPIED"); } catch (e) { toast(t); }
  }
  async function ackShip(tracking, on) {
    if (await mutate(doc => { const s = doc.shipments.find(x => x.tracking === tracking); if (!s) return false; s.ack = on; s.ackBy = on ? myId : null; s.ackAt = on ? nowIso() : null; }))
      toast(on ? "ACKNOWLEDGED" : "ACKNOWLEDGEMENT REMOVED");
  }
  async function questFromShip(tracking) {
    const s = (data.shipments || []).find(x => x.tracking === tracking); if (!s) return;
    const id = uid(), urgent = s.status === "exception" || s.status === "claim";
    const q = { id, title: (urgent ? "Resolve UPS " + SHIP_ST[s.status].label.toLowerCase() : "Watch UPS shipment") + ": " + (s.ref || s.tracking) + (s.shipTo ? " → " + s.shipTo.split(",")[0] : ""),
      giver: "UPS tracker", tier: urgent ? "urgent" : "side", due: todayIso(), details: `Tracking ${s.tracking}. Latest: ${s.lastText || "—"}`,
      source: "UPS tracker", sourceLink: (s.events[s.events.length - 1] || {}).link || null, status: "open", assignees: [], credit: [], doneAt: null, createdAt: nowIso(), createdBy: myId,
      checklist: [], notes: [], history: [{ at: nowIso(), by: myId, action: "posted it", detail: "from the UPS tracker" }], squad: false, requiresProof: false, fromInbox: false, reviewed: true, intel: null };
    if (await mutate(doc => { const sh = doc.shipments.find(x => x.tracking === tracking); if (!sh) return false; doc.quests.push(q); sh.questId = id; })) toast("QUEST CREATED");
  }
  $("ups-search").addEventListener("input", (e) => { upsQuery = e.target.value; renderUps(); });

  /* ---------- Render: TV ---------- */
  function renderTv() {
    const open = data.quests.filter(q => q.status !== "done").sort((a, b) => {
      const da = daysLeft(a.due), dbb = daysLeft(b.due);
      const ua = (a.tier === "urgent" || (da !== null && da <= 0)) ? 0 : 1, ub = (b.tier === "urgent" || (dbb !== null && dbb <= 0)) ? 0 : 1;
      return ua - ub || sortQ(a, b);
    }).slice(0, 8);
    const open2 = open.filter(q => !q.boss);
    $("tvq").replaceChildren(...(open2.length ? open2.slice(0, 7).map(q => el("div", { class: "tvq", style: `--tier:${tierOf(q).color}` },
      el("div", { class: "q-meta" }, el("span", { class: "chip tier", style: `--tier:${tierOf(q).color}` }, tierOf(q).label), dueChip(q), whoChips(q)),
      el("h3", {}, q.title))) : [el("div", { class: "empty" }, el("h2", {}, "Board clear"), el("p", {}, "No open missions. Nice work, crew."))]));
    const bosses = data.quests.filter(q => q.boss && q.status !== "done");
    $("tvboss").replaceChildren(...bosses.slice(0, 2).map(q => el("div", { class: "tvq tvbossq" }, el("h3", {}, "⚔ " + q.title), bossBar(q, true))));
    const bq = bountyQuest();
    $("tvbounty").replaceChildren(...(bq ? [el("div", { class: "bountybar tvb" }, el("b", {}, "💰 TODAY'S BOUNTY · 2× XP"), el("span", {}, bq.title))] : []));
    $("tvgoal").replaceChildren(...[goalBanner(true)].filter(Boolean));
    const sh = shoutFeed(true);
    $("tvshouts").replaceChildren(...(sh.length ? [el("div", { class: "tvh" }, el("span", {}, "SHOUT-OUTS"))].concat(el("ul", { class: "shouts tv" }, sh)) : []));
    const m = monthKey(nowIso());
    $("tvlbTitle").textContent = "SEASON · " + monthName(m).toUpperCase();
    $("tvlb").replaceChildren(...ranking(m).slice(0, 8).map(({ c, xp, s }, i) => el("li", {},
      el("span", { class: "pos" }, i + 1), el("span", {}, c.name, el("span", { style: "color:var(--muted);font-size:16px" }, ` L${levelOf(s.xp)}${s.streak ? " · 🔥" + s.streak : ""}`)), el("span", { class: "pts" }, xp))));
    const hof = hallOfFame();
    $("tvhof").textContent = hof.length ? "HALL OF FAME · " + hof.map(h => `${monthName(h.m)}: ${nameOf(h.id)}`).join(" · ") : "";
    $("tvclock").textContent = new Date().toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  }

  /* ---------- Render: frame ---------- */
  function render() {
    buildLedger();
    if (view !== "board" && view !== "tv" && view !== "ups" && !isAdmin) view = "board";
    document.querySelectorAll(".view").forEach(b => b.setAttribute("aria-current", b.dataset.view === view ? "page" : "false"));
    ["board", "debrief", "ops", "ups", "tv"].forEach(v => { $("view-" + v).hidden = view !== v; });
    document.body.classList.toggle("tv", view === "tv");
    document.querySelector(".views").hidden = view === "tv";
    $("foot").hidden = view === "tv";
    renderBoardSide();
    if (view === "board") { renderHud(); renderTabs(); renderList(); }
    if (view === "debrief") renderDebrief();
    if (view === "ops") renderOps();
    if (view === "ups") renderUps();
    if (view === "tv") renderTv();
  }
  document.querySelectorAll(".view").forEach(b => b.onclick = () => { view = b.dataset.view; history.replaceState(null, "", view === "tv" ? "#tv" : location.pathname); render(); });
  $("tvExit").onclick = () => { view = "board"; history.replaceState(null, "", location.pathname); render(); };
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && view === "tv") $("tvExit").click(); });
  $("lb-season").onclick = () => { lbMode = "season"; try { localStorage.setItem("dqb-lb", lbMode); } catch (e) {} render(); };
  $("lb-all").onclick = () => { lbMode = "all"; try { localStorage.setItem("dqb-lb", lbMode); } catch (e) {} render(); };

  /* ---------- Actions ---------- */
  const kindOf = (id) => pdata.quests.some(x => x.id === id) ? "personal" : "shared";
  function withQuest(id, fn) { return mutate(doc => { const q = doc.quests.find(x => x.id === id); if (!q) { toast("That quest was removed."); return false; } return fn(q, doc); }, kindOf(id)); }
  async function moveQuest(id, to) {
    const from = kindOf(id), src = (from === "personal" ? pdata : data).quests.find(x => x.id === id);
    if (!src || from === to) return;
    const copy = JSON.parse(JSON.stringify(src));
    delete copy.scope; copy.assignees = []; copy.credit = copy.status === "done" ? [myId] : [];
    copy.history.push({ at: nowIso(), by: myId, action: to === "personal" ? "moved it to their personal list" : "moved it to the warehouse board", detail: "" });
    if (!(await mutate(doc => { if (doc.quests.some(x => x.id === id)) return false; doc.quests.push(copy); }, to))) return;
    if (await mutate(doc => { doc.quests = doc.quests.filter(x => x.id !== id); }, from)) toast(to === "personal" ? "MOVED TO YOUR PERSONAL LIST" : "MOVED TO THE WAREHOUSE BOARD");
  }
  async function act(id, what) {
    if (!joined()) return;
    const ok = await withQuest(id, (q) => {
      if (what === "claim") {
        if (q.status === "done") return false;
        if (!q.squad && q.assignees.length && !q.assignees.includes(myId)) { toast("Someone else just claimed it."); return false; }
        if (!q.assignees.includes(myId)) q.assignees.push(myId);
        log(q, q.squad ? "joined the squad" : "claimed it");
      }
      if (what === "drop") { if (!q.assignees.includes(myId)) return false; q.assignees = q.assignees.filter(x => x !== myId); log(q, q.squad ? "left the squad" : "dropped it"); }
      if (what === "reopen") { if (!isAdmin && !q.credit.includes(myId)) return false; q.status = "open"; q.credit = []; q.doneAt = null; q.proof = null; log(q, "reopened it"); }
    });
    if (ok) toast({ claim: "QUEST CLAIMED", drop: "QUEST DROPPED", reopen: "QUEST REOPENED" }[what]);
  }
  function creditFor(q) {
    if (isPersonal(q)) return [myId];
    if (q.boss) { const f = [...new Set(q.checklist.filter(c => c.done && c.by).map(c => c.by))]; if (!f.includes(myId) && !isAdmin) f.push(myId); return f.length ? f : [myId]; }
    if (q.squad) { const s = [...q.assignees]; if (!isAdmin && !s.includes(myId)) s.push(myId); return s.length ? s : [myId]; }
    if (isAdmin && q.assignees.length) return [...q.assignees];
    return [myId];
  }
  async function complete(id, intelProof) {
    const tracking = intelProof ? "" : (drafts["pf-t-" + id] || "").trim();
    const note = intelProof ? intelProof : (drafts["pf-n-" + id] || "").trim();
    const local = allQ().find(x => x.id === id);
    if (local && local.requiresProof && !tracking && !note) { const e = $("pf-e-" + id); if (e) e.textContent = "Add a tracking number or a note to clear this one."; return; }
    let credited = [], gained = 0;
    const ok = await withQuest(id, (q) => {
      if (q.status === "done") { toast("Already cleared."); return false; }
      if (!isPersonal(q) && !q.squad && !q.boss && q.assignees.length && !q.assignees.includes(myId) && !isAdmin) { toast("That quest is claimed by someone else."); return false; }
      credited = creditFor(q);
      q.status = "done"; q.credit = credited; q.doneAt = nowIso();
      if (!q.assignees.length) q.assignees = [...credited];
      q.proof = (tracking || note) ? { tracking, note, by: myId, at: nowIso() } : null;
      if (intelProof) q.intel = null;
      log(q, "cleared it", tracking ? "tracking " + tracking : (intelProof ? "confirmed Claude's intel" : ""));
      gained = awards(q).filter(a => a.id === myId && a.kind === "clear").reduce((t, a) => t + a.xp, 0);
    });
    if (!ok) return;
    completing = null; clearDrafts("pf-t-" + id, "pf-n-" + id); render();
    toast(local && isPersonal(local) ? "PERSONAL QUEST CLEARED" : credited.includes(myId) ? `QUEST COMPLETE · +${gained} XP` : `CLEARED · XP TO ${credited.map(nameOf).join(", ").toUpperCase()}`);
    const a = $("avatarBox"); a.classList.remove("pop"); void a.offsetWidth; a.classList.add("pop");
  }
  async function confirmIntel(id) {
    const q = allQ().find(x => x.id === id); if (!q || !q.intel) return;
    await complete(id, "Confirmed from inbox: " + q.intel.text);
  }
  async function dismissIntel(id) { if (await withQuest(id, q => { q.intel = null; log(q, "dismissed Claude's intel"); })) toast("INTEL DISMISSED"); }
  async function markReviewed(id) { if (await withQuest(id, q => { q.reviewed = true; log(q, "reviewed the inbox quest"); })) toast("MARKED REVIEWED"); }
  async function assign(id, who) {
    if (await withQuest(id, (q) => {
      if (!who) { q.assignees = []; log(q, "cleared the assignment"); return; }
      if (q.squad) { if (!q.assignees.includes(who)) q.assignees.push(who); } else q.assignees = [who];
      log(q, "assigned it", "to " + nameOf(who));
    })) toast(who ? "ASSIGNED TO " + nameOf(who).toUpperCase() : "UNASSIGNED");
  }
  async function deleteQuest(id) { confirmQ = null; if (await mutate(doc => { const n = doc.quests.length; doc.quests = doc.quests.filter(x => x.id !== id); if (doc.quests.length === n) return false; })) toast("QUEST DELETED"); }
  async function removePlayer(id) {
    confirmP = null;
    if (await mutate(doc => { doc.crew = doc.crew.filter(c => c.id !== id); doc.quests.forEach(q => { if (q.status !== "done") q.assignees = q.assignees.filter(x => x !== id); }); })) toast("PLAYER REMOVED");
  }
  async function addNote(id) {
    const text = (drafts["note-" + id] || "").trim(); if (!text) return;
    if (await withQuest(id, q => { q.notes.push({ id: uid("n"), by: myId, at: nowIso(), text: text.slice(0, 400) }); if (q.notes.length > 50) q.notes = q.notes.slice(-50); })) { clearDrafts("note-" + id); render(); toast("NOTE POSTED"); }
  }
  async function addStep(id) {
    const text = (drafts["cl-" + id] || "").trim(); if (!text) return;
    if (await withQuest(id, q => { if (!canEditQuest(q)) return false; q.checklist.push({ id: uid("c"), text: text.slice(0, 120), done: false }); log(q, "added a step", text); })) { clearDrafts("cl-" + id); render(); }
  }
  async function toggleStep(id, stepId) {
    let hit = false;
    await withQuest(id, q => {
      const c = q.checklist.find(x => x.id === stepId); if (!c) return false;
      if (q.boss && c.done && c.by && c.by !== myId && !isAdmin) { toast("Only the person who landed that hit can undo it."); return false; }
      c.done = !c.done; c.by = c.done ? myId : null; c.at = c.done ? nowIso() : null; hit = q.boss && c.done;
      log(q, c.done ? (q.boss ? "hit the boss" : "checked off") : "unchecked", c.text);
    });
    if (hit) toast(`HIT! +${BOSS_STEP_XP} XP`);
  }
  async function removeStep(id, stepId) { await withQuest(id, q => { const n = q.checklist.length; q.checklist = q.checklist.filter(x => x.id !== stepId); if (q.checklist.length === n) return false; }); }
  async function toggleRecurring(rid) { await mutate(doc => { const r = doc.recurring.find(x => x.id === rid); if (!r) return false; r.active = !r.active; }); }
  async function deleteRecurring(rid) { confirmR = null; if (await mutate(doc => { doc.recurring = doc.recurring.filter(x => x.id !== rid); })) toast("RECURRING MISSION DELETED"); }

  $("joinForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const role = $("j-role").value.trim().slice(0, 50), color = pickColor, first = !member(myId);
    const ok = await mutate(doc => {
      const me = doc.crew.find(c => c.id === myId);
      if (me) { me.role = role; me.color = color; me.name = myName || me.name; me.email = myEmail; }
      else doc.crew.push({ id: myId, name: myName || myEmail || "Crew member", email: myEmail, role, color, joinedAt: nowIso() });
    });
    if (ok) { editing = false; render(); toast(first ? "WELCOME TO THE CREW!" : "PLAYER UPDATED"); }
  });
  $("editMe").onclick = () => { const m = member(myId); editing = true; $("j-role").value = (m && m.role) || ""; pickColor = (m && m.color) || COLORS[0]; render(); $("j-role").focus(); };

  $("qForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const title = $("q-title").value.trim();
    if (!title) { $("q-msg").className = "err"; $("q-msg").textContent = "Give the quest a name."; return; }
    const steps = $("q-checklist").value.split("\n").map(s => s.trim()).filter(Boolean).slice(0, 25);
    const q = { id: uid(), title, giver: $("q-giver").value.trim(), tier: $("q-tier").value, due: $("q-due").value || null, details: $("q-details").value.trim(),
      source: "Posted on the board by " + (myName || myEmail), status: "open", assignees: [], credit: [], doneAt: null, createdAt: nowIso(), createdBy: myId,
      checklist: steps.map(text => ({ id: uid("c"), text, done: false })), notes: [], history: [{ at: nowIso(), by: myId, action: "posted it", detail: "" }],
      squad: $("q-squad").checked, boss: $("q-boss").checked, requiresProof: $("q-proof").checked, fromInbox: false, reviewed: true, intel: null };
    const board = $("q-board").value;
    if (q.boss) {
      if (board === "personal") q.boss = false;
      else if (steps.length < 2) { $("q-msg").className = "err"; $("q-msg").textContent = "A boss fight needs at least 2 checklist steps (one per line). Each step is a hit."; return; }
      else q.squad = false;
    }
    if (board === "personal") q.squad = false;
    const repeat = board === "personal" ? "" : $("q-repeat").value;
    let tmpl = null;
    if (repeat) {
      tmpl = { id: uid("r"), title, giver: q.giver, tier: q.tier, cadence: repeat, weekday: Number($("q-weekday").value), day: Math.min(28, Math.max(1, Number($("q-day").value) || 1)),
        details: q.details, checklist: steps, squad: q.squad, boss: q.boss, requiresProof: q.requiresProof, active: true, lastSpawned: todayIso(), createdAt: nowIso() };
      q.recurringId = tmpl.id; q.source = "Recurring mission (" + cadenceText(tmpl).replace(/^E/, "e").replace(/^M/, "m") + ")";
      if (!q.due) q.due = todayIso();
    }
    if (await mutate(doc => { doc.quests.push(q); if (tmpl) doc.recurring.push(tmpl); }, board)) {
      ["q-title","q-giver","q-due","q-details","q-checklist"].forEach(id => $(id).value = ""); $("q-proof").checked = false; $("q-squad").checked = false; $("q-boss").checked = false;
      $("q-repeat").value = ""; syncRepeat();
      $("q-msg").className = "hint"; $("q-msg").textContent = tmpl ? `Quest posted. It will repeat ${cadenceText(tmpl).replace(/^E/, "e").replace(/^M/, "m")}.` : "Quest posted.";
      toast(board === "personal" ? "ADDED TO YOUR PERSONAL LIST" : tmpl ? "RECURRING QUEST POSTED" : "QUEST POSTED");
    }
  });
  $("rForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const title = $("r-title").value.trim();
    if (!title) { $("r-msg").className = "err"; $("r-msg").textContent = "Give the mission a name."; return; }
    const r = { id: uid("r"), title, tier: $("r-tier").value, cadence: $("r-cadence").value, weekday: Number($("r-weekday").value), day: Math.min(28, Math.max(1, Number($("r-day").value) || 1)),
      details: $("r-details").value.trim(), checklist: $("r-checklist").value.split("\n").map(s => s.trim()).filter(Boolean).slice(0, 25), active: true, lastSpawned: null, createdAt: nowIso() };
    let posted = 0;
    if (await mutate(doc => { doc.recurring.push(r); posted = spawnInto(doc); })) {
      ["r-title","r-details","r-checklist"].forEach(id => $(id).value = "");
      $("r-msg").className = "hint"; $("r-msg").textContent = posted ? "Added, and today's mission is already on the board." : "Added. It will post itself on its next day.";
      toast("RECURRING MISSION ADDED");
    }
  });

  /* ---------- Boot ---------- */
  function signIn() { msalApp.loginRedirect({ scopes: SCOPES }); }
  $("signout").onclick = () => msalApp.logoutRedirect({ account });
  setInterval(() => { if (view === "tv") renderTv(); }, 30000);

  /* Preview mode: same app, data kept in this browser instead of SharePoint, no Microsoft sign-in. */
  async function bootDemo() {
    const acct = C.DEMO.account;
    account = { name: acct.name, username: acct.email };
    myId = acct.id; myName = acct.name; myEmail = acct.email.toLowerCase();
    isAdmin = (C.ADMIN_EMAILS || []).map(s => s.toLowerCase()).includes(myEmail);
    $("signout").hidden = true;
    await locate(); await refresh();
    setInterval(() => { if (document.visibilityState === "visible") refresh(); }, 5000);
  }
  async function boot() {
    render();
    if (C.DEMO) return bootDemo();
    if (!C.CLIENT_ID || C.CLIENT_ID.startsWith("PASTE")) { fatal("Setup not finished", "Fill in config.js with the IDs from the app registration (see README)."); return; }
    if (!window.msal) { fatal("Sign-in library didn't load", "Check your connection, then reload the page."); return; }
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: C.CLIENT_ID, authority: "https://login.microsoftonline.com/" + C.TENANT_ID, redirectUri: window.location.origin + window.location.pathname },
      cache: { cacheLocation: "localStorage" },
    });
    try {
      const res = await msalApp.handleRedirectPromise();
      account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
    } catch (e) { fatal("Sign-in failed", "Reload the page and try again. If it keeps happening, ask IT to check the app registration."); return; }
    if (!account) { render(); return; }
    const claims = account.idTokenClaims || {};
    myId = claims.oid || account.localAccountId;
    myName = account.name || "";
    myEmail = (account.username || "").toLowerCase();
    isAdmin = (C.ADMIN_EMAILS || []).map(s => s.toLowerCase()).includes(myEmail);
    render();
    try { await locate(); }
    catch (e) {
      const m = String(e.message);
      fatal("Can't reach the board's data", m.startsWith("site") ? "The SharePoint site in config.js wasn't found, or you don't have access to it. Ask your manager to add you."
        : m.startsWith("file") ? "The board's data file wasn't found in SharePoint. Check DATA_FILE in config.js." : "Reload the page to try again.");
      return;
    }
    await refresh();
    setInterval(() => { if (document.visibilityState === "visible") refresh(); }, (C.REFRESH_SECONDS || 20) * 1000);
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") refresh(); });
  }
  boot();
})();
