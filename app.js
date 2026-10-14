(() => {
  var __defProp = Object.defineProperty;
  var __defProps = Object.defineProperties;
  var __getOwnPropDescs = Object.getOwnPropertyDescriptors;
  var __getOwnPropSymbols = Object.getOwnPropertySymbols;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __propIsEnum = Object.prototype.propertyIsEnumerable;
  var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
  var __spreadValues = (a, b) => {
    for (var prop in b || (b = {}))
      if (__hasOwnProp.call(b, prop))
        __defNormalProp(a, prop, b[prop]);
    if (__getOwnPropSymbols)
      for (var prop of __getOwnPropSymbols(b)) {
        if (__propIsEnum.call(b, prop))
          __defNormalProp(a, prop, b[prop]);
      }
    return a;
  };
  var __spreadProps = (a, b) => __defProps(a, __getOwnPropDescs(b));

  // shim/msal.js
  var createNestablePublicClientApplication = (c) => window.msal.createNestablePublicClientApplication(c);

  // src/slots.js
  var DEFAULT_SETTINGS = {
    durationMin: 60,
    // meeting length
    count: 3,
    // how many options to offer
    bufferMin: 15,
    // gap kept before/after existing meetings
    workStart: "08:30",
    // earliest start (local time)
    workEnd: "17:00",
    // latest end (local time)
    workDays: [1, 2, 3, 4, 5],
    // 0=Sun ... 6=Sat
    startOffsetDays: 1,
    // 0 = may suggest later today, 1 = from tomorrow
    searchDays: 21,
    // how far ahead to look (calendar days)
    stepMin: 30,
    // candidate start times every N minutes
    onePerDay: true,
    // spread the options over different days
    tentativeIsBusy: true,
    minNoticeHours: 2,
    // never suggest something sooner than this
    intro: ""
    // optional line inserted above the options
  };
  var BUSY = /* @__PURE__ */ new Set(["busy", "oof", "unknown"]);
  function isBusyStatus(status2, tentativeIsBusy) {
    if (status2 === "tentative") return tentativeIsBusy;
    return BUSY.has(status2);
  }
  function hm(str) {
    const [h, m] = String(str).split(":").map(Number);
    return { h: h || 0, m: m || 0 };
  }
  function findSlots(busy, settings, now = /* @__PURE__ */ new Date()) {
    const s = __spreadValues(__spreadValues({}, DEFAULT_SETTINGS), settings);
    const dur = s.durationMin * 6e4;
    const buf = s.bufferMin * 6e4;
    const earliest = now.getTime() + s.minNoticeHours * 36e5;
    const blocks = busy.map((b) => ({ start: b.start - buf, end: b.end + buf })).sort((a, b) => a.start - b.start);
    const ws = hm(s.workStart);
    const we = hm(s.workEnd);
    const out = [];
    for (let d = s.startOffsetDays; d < s.startOffsetDays + s.searchDays && out.length < s.count; d++) {
      const day = new Date(now.getFullYear(), now.getMonth(), now.getDate() + d);
      if (!s.workDays.includes(day.getDay())) continue;
      const dayStart = new Date(day.getFullYear(), day.getMonth(), day.getDate(), ws.h, ws.m).getTime();
      const dayEnd = new Date(day.getFullYear(), day.getMonth(), day.getDate(), we.h, we.m).getTime();
      for (let t = dayStart; t + dur <= dayEnd; t += s.stepMin * 6e4) {
        if (t < earliest) continue;
        const end = t + dur;
        const clash = blocks.some((b) => b.start < end && b.end > t);
        const overlapsChosen = out.some((o) => o.start.getTime() < end && o.end.getTime() > t);
        if (clash || overlapsChosen) continue;
        out.push({ start: new Date(t), end: new Date(end) });
        if (s.onePerDay || out.length >= s.count) break;
      }
    }
    return out;
  }
  function formatSlot(slot, locale = "en-ZA") {
    var _a;
    const date = slot.start.toLocaleDateString(locale, { weekday: "long", day: "numeric", month: "long" });
    const time = (d) => d.toLocaleTimeString(locale, { hour: "2-digit", minute: "2-digit", hour12: false });
    let tz = "";
    try {
      tz = ((_a = new Intl.DateTimeFormat(locale, { timeZoneName: "short" }).formatToParts(slot.start).find((p) => p.type === "timeZoneName")) == null ? void 0 : _a.value) || "";
    } catch (_) {
    }
    return date + ", " + time(slot.start) + " \u2013 " + time(slot.end) + (tz ? " (" + tz + ")" : "");
  }
  function esc(t) {
    return t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function slotsToHtml(slots, intro = "") {
    const lines = slots.map((s) => esc(formatSlot(s)));
    const head = intro ? esc(intro) + "<br>" : "";
    return "<div>" + head + lines.join("<br>") + "</div><br>";
  }

  // src/core.js
  var SCOPES = ["User.Read", "Calendars.ReadBasic"];
  var SETTINGS_KEY = "threeSlotsSettings";
  var pcaPromise = null;
  function config() {
    const c = window.ADDIN_CONFIG || {};
    if (!c.clientId || c.clientId.startsWith("PASTE")) {
      throw new Error("Add-in not configured yet: put your Application (client) ID in config.js.");
    }
    return c;
  }
  function getPca() {
    if (!pcaPromise) {
      const c = config();
      pcaPromise = createNestablePublicClientApplication({
        auth: { clientId: c.clientId, authority: "https://login.microsoftonline.com/" + (c.tenant || "organizations") }
      });
    }
    return pcaPromise;
  }
  async function getToken(interactive) {
    const pca = await getPca();
    const req = { scopes: SCOPES };
    try {
      return (await pca.acquireTokenSilent(req)).accessToken;
    } catch (silentErr) {
      if (!interactive) {
        const e = new Error("Please open 'Availability settings' once and click Sign in.");
        e.cause = silentErr;
        throw e;
      }
      return (await pca.acquireTokenPopup(req)).accessToken;
    }
  }
  function loadSettings() {
    let saved = {};
    try {
      saved = Office.context.roamingSettings.get(SETTINGS_KEY) || {};
    } catch (_) {
    }
    return __spreadValues(__spreadValues({}, DEFAULT_SETTINGS), saved);
  }
  function saveSettings(settings) {
    return new Promise((resolve, reject) => {
      Office.context.roamingSettings.set(SETTINGS_KEY, settings);
      Office.context.roamingSettings.saveAsync((r) => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error));
    });
  }
  function graphDate(d) {
    return d.toISOString().replace("Z", "");
  }
  function parseGraphUtc(s) {
    return Date.parse(s.slice(0, 19) + "Z");
  }
  async function fetchBusy(token, settings, now = /* @__PURE__ */ new Date()) {
    const email = Office.context.mailbox.userProfile.emailAddress;
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const end = new Date(start.getTime() + (settings.startOffsetDays + settings.searchDays + 1) * 864e5);
    const res = await fetch("https://graph.microsoft.com/v1.0/me/calendar/getSchedule", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
        Prefer: 'outlook.timezone="UTC"'
      },
      body: JSON.stringify({
        schedules: [email],
        startTime: { dateTime: graphDate(start), timeZone: "UTC" },
        endTime: { dateTime: graphDate(end), timeZone: "UTC" },
        availabilityViewInterval: 15
      })
    });
    if (!res.ok) throw new Error("Calendar lookup failed (" + res.status + "): " + (await res.text()).slice(0, 200));
    const data = await res.json();
    const sched = data.value && data.value[0] || {};
    if (sched.error) throw new Error("Calendar lookup failed: " + sched.error.message);
    return (sched.scheduleItems || []).filter((i) => isBusyStatus(i.status, settings.tentativeIsBusy)).map((i) => ({ start: parseGraphUtc(i.start.dateTime), end: parseGraphUtc(i.end.dateTime) }));
  }
  async function computeSlots(interactive) {
    const settings = loadSettings();
    const token = await getToken(interactive);
    const busy = await fetchBusy(token, settings);
    return { slots: findSlots(busy, settings), settings };
  }
  function insertHtml(html) {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setSelectedDataAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (r) => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
      );
    });
  }
  async function insertSlots(interactive) {
    const { slots, settings } = await computeSlots(interactive);
    if (!slots.length) throw new Error("No free slots found in your search window. Try widening it in Availability settings.");
    await insertHtml(slotsToHtml(slots, settings.intro));
    return slots;
  }

  // src/commands.js
  function notify(type, message) {
    const item = Office.context.mailbox.item;
    const msg = { type, message: message.slice(0, 150) };
    if (type === Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
      msg.icon = "Icon.16x16";
      msg.persistent = false;
    }
    item.notificationMessages.replaceAsync("threeSlots", msg);
  }
  async function insertThreeSlots(event) {
    try {
      const slots = await insertSlots(false);
      notify(
        Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        "Inserted " + slots.length + " available time" + (slots.length === 1 ? "" : "s") + " from your calendar."
      );
    } catch (err) {
      notify(Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, err.message || String(err));
    } finally {
      event.completed();
    }
  }
  window.insertThreeSlots = insertThreeSlots;
  Office.onReady(() => {
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("insertThreeSlots", insertThreeSlots);
    }
  });

  // src/taskpane.js
  var $ = (id) => document.getElementById(id);
  var DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var lastSlots = [];
  function status(text, isError) {
    const el = $("status");
    el.textContent = text;
    el.className = isError ? "status error" : "status";
  }
  function fillForm(s) {
    $("durationMin").value = s.durationMin;
    $("count").value = s.count;
    $("bufferMin").value = s.bufferMin;
    $("workStart").value = s.workStart;
    $("workEnd").value = s.workEnd;
    $("startOffsetDays").value = String(s.startOffsetDays);
    $("searchDays").value = s.searchDays;
    $("onePerDay").checked = s.onePerDay;
    $("tentativeIsBusy").checked = s.tentativeIsBusy;
    $("intro").value = s.intro;
    const box = $("workDays");
    box.innerHTML = "";
    [1, 2, 3, 4, 5, 6, 0].forEach((d) => {
      const label = document.createElement("label");
      label.className = "day";
      label.innerHTML = '<input type="checkbox" value="' + d + '" ' + (s.workDays.includes(d) ? "checked" : "") + "><span>" + DAYS[d] + "</span>";
      box.appendChild(label);
    });
  }
  function readForm() {
    const num = (id, min, max, dflt) => {
      const v = parseInt($(id).value, 10);
      return Number.isFinite(v) ? Math.min(max, Math.max(min, v)) : dflt;
    };
    return __spreadProps(__spreadValues({}, loadSettings()), {
      durationMin: num("durationMin", 15, 480, 60),
      count: num("count", 1, 10, 3),
      bufferMin: num("bufferMin", 0, 120, 15),
      workStart: $("workStart").value || "08:30",
      workEnd: $("workEnd").value || "17:00",
      startOffsetDays: num("startOffsetDays", 0, 14, 1),
      searchDays: num("searchDays", 1, 60, 21),
      onePerDay: $("onePerDay").checked,
      tentativeIsBusy: $("tentativeIsBusy").checked,
      intro: $("intro").value.trim(),
      workDays: [...document.querySelectorAll("#workDays input:checked")].map((i) => Number(i.value))
    });
  }
  async function onSave() {
    try {
      await saveSettings(readForm());
      status("Saved. The 'Insert 3 times' button will use these settings.");
    } catch (e) {
      status("Couldn't save settings: " + (e.message || e), true);
    }
  }
  async function onSignIn() {
    try {
      status("Signing in\u2026");
      await getToken(true);
      status("Signed in. You can now use the 'Insert 3 times' button directly.");
      $("signin").textContent = "Signed in \u2713";
    } catch (e) {
      status("Sign-in failed: " + (e.message || e), true);
    }
  }
  async function onPreview() {
    try {
      await saveSettings(readForm());
      status("Checking your calendar\u2026");
      const { slots } = await computeSlots(true);
      lastSlots = slots;
      const list = $("preview");
      list.innerHTML = "";
      slots.forEach((s) => {
        const li = document.createElement("li");
        li.textContent = formatSlot(s);
        list.appendChild(li);
      });
      $("insert").disabled = !slots.length;
      status(slots.length ? "Here's what would be inserted." : "No free slots in that window. Try widening it.", !slots.length);
    } catch (e) {
      status(e.message || String(e), true);
    }
  }
  async function onInsert() {
    try {
      if (!lastSlots.length) await onPreview();
      if (!lastSlots.length) return;
      await insertHtml(slotsToHtml(lastSlots, readForm().intro));
      status("Inserted into your email.");
    } catch (e) {
      status("Couldn't insert: " + (e.message || e), true);
    }
  }
  Office.onReady(() => {
    if (!document.getElementById("durationMin")) return;
    fillForm(loadSettings());
    $("save").onclick = onSave;
    $("signin").onclick = onSignIn;
    $("previewBtn").onclick = onPreview;
    $("insert").onclick = onInsert;
  });
})();
