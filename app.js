document.addEventListener("DOMContentLoaded", () => {

  const GOOGLE_SHEET_WEBAPP_URL =
    "https://script.google.com/macros/s/AKfycby6r0mKxNzW2UmNKzJeWhkE4sxdABmaHBOCM7XvAvxxabFqmqNOA2at2gXGhshvG-9b/exec";

  /* ================= TIEMPO ================= */
  function isoNowSeconds() {
    const d = new Date();
    d.setMilliseconds(0);
    return d.toISOString();
  }

  function formatDateTimeAR(iso) {
    try {
      return new Date(iso).toLocaleString("es-AR", { timeZone: "America/Argentina/Buenos_Aires" });
    } catch {
      return "";
    }
  }

  function dayKeyAR() {
    // YYYY-MM-DD siempre, en horario AR, sin depender del locale
    const parts = new Intl.DateTimeFormat("en-GB", {
      timeZone: "America/Argentina/Buenos_Aires",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).formatToParts(new Date());
  
    const y = parts.find(p => p.type === "year")?.value || "0000";
    const m = parts.find(p => p.type === "month")?.value || "00";
    const d = parts.find(p => p.type === "day")?.value || "00";
  
    return `${y}-${m}-${d}`;
  }

  function nowMinutesAR() {
    const parts = new Intl.DateTimeFormat("es-AR", {
      timeZone: "America/Argentina/Buenos_Aires",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    }).formatToParts(new Date());

    const hh = Number(parts.find(p => p.type === "hour")?.value || 0);
    const mm = Number(parts.find(p => p.type === "minute")?.value || 0);
    return hh * 60 + mm;
  }
  // ✅ Matriz 501: permite decimales en Cajón (C)
  function isMatrix501(state) {
    return String(state?.lastMatrix?.texto || "").trim() === "501";
  }
  
  // ✅ Normaliza: punto -> coma (12.5 => 12,5)
  function normalizeToComma(value) {
    return String(value || "").trim().replace(/\./g, ",");
  }


  /* ================= KEYS (Cervantes) ================= */
  const APP_TAG = "_Cervantes";
  const VERSION = "_v1";
  const MAX_DAY_HISTORY = 700; // historial del día por legajo (se borra al cambiar de día)

  const MIGRATION_FLAG = `prod_migrated${APP_TAG}${VERSION}`;
  const LS_PREFIX      = `prod_state${APP_TAG}${VERSION}`;
  const LS_QUEUE       = `prod_queue${APP_TAG}${VERSION}`;
  const LS_FAILED      = `prod_failed${APP_TAG}${VERSION}`;
  const DAY_GUARD_KEY  = `prod_day_guard${APP_TAG}${VERSION}`;


  /* ================= RESET DIARIO ================= */
  function clearAllCervantesData() {
    const statePrefix = `${LS_PREFIX}::`;
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) keys.push(localStorage.key(i));

    keys.forEach(k => {
      if (!k) return;
      if (k.startsWith(statePrefix)) localStorage.removeItem(k);
    });
  }

  const today = dayKeyAR();
  const lastDay = localStorage.getItem(DAY_GUARD_KEY);
  if (lastDay && lastDay !== today) {
    clearAllCervantesData();
  }
  localStorage.setItem(DAY_GUARD_KEY, today);

  /* ================= LIMPIEZA (1 vez) ================= */
  if (!localStorage.getItem(MIGRATION_FLAG)) {
    [
      "prod_day_state_ls_v1",
      "prod_send_queue_ls_v1",
      "legajo_history_v1",
      "prod_day_state_v7",
      "prod_state_ls_v1",
      "prod_queue_v1",
      // OJO: NO borrar `prod_state${APP_TAG}${VERSION}`
      // OJO: NO borrar `prod_queue${APP_TAG}${VERSION}`
    ].forEach(k => localStorage.removeItem(k));

    localStorage.setItem(MIGRATION_FLAG, "1");
  }

  /* ================= UUID ================= */
  function uuidv4() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    const bytes = new Uint8Array(16);
    (window.crypto || window.msCrypto).getRandomValues(bytes);
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = [...bytes].map(b => b.toString(16).padStart(2, "0")).join("");
    return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
  }

  /* ================= ELEMENTOS ================= */
  const $ = (id) => document.getElementById(id);

  const legajoScreen  = $("legajoScreen");
  const optionsScreen = $("optionsScreen");
  const legajoInput   = $("legajoInput");

  const btnContinuar = $("btnContinuar");
  const btnBackTop   = $("btnBackTop");
  const btnBackLabel = $("btnBackLabel");

  const row1 = $("row1");
  const row2 = $("row2");
  const row3 = $("row3");
  const row4 = $("row4");


  const selectedArea = $("selectedArea");
  const selectedBox  = $("selectedBox");
  const selectedDesc = $("selectedDesc");
  const inputArea    = $("inputArea");
  const inputLabel   = $("inputLabel");
  const textInput    = $("textInput");
  const btnResetSelection = $("btnResetSelection");
  const btnEnviar    = $("btnEnviar");
  const error        = $("error");

  const daySummary = $("daySummary");
  const matrizInfo = $("matrizInfo");

  const required = {
    legajoScreen, optionsScreen, legajoInput,
    btnContinuar, btnBackTop, btnBackLabel,
    row1, row2, row3,row4,
    selectedArea, selectedBox, selectedDesc, inputArea, inputLabel, textInput,
    btnResetSelection, btnEnviar, error,
    daySummary, matrizInfo
  };
  const missing = Object.entries(required).filter(([,v]) => !v).map(([k]) => k);
  if (missing.length) {
    console.error("FALTAN ELEMENTOS EN EL HTML (ids):", missing);
    alert("Error: faltan elementos en el HTML. Mirá consola (F12).");
    return;
  }

  /* ================= OPCIONES ================= */
  const OPTIONS = [
    {code:"E",desc:"Empecé Matriz",row:1,input:{show:true,label:"Ingresar número",placeholder:"Ejemplo: 110",validate:/^[0-9]+$/}},
    // ✅ C solo números (ya estaba OK)
    {code:"C",desc:"Cajón",row:1,input:{show:true,label:"Ingresar número",placeholder:"Ejemplo: 1500",validate:/^[0-9]+$/}},
    {code:"PB",desc:"Paré Baño",row:2,input:{show:false}},
    {code:"BC",desc:"Busqué Cajón",row:2,input:{show:false}},
    {code:"MOV",desc:"Movimiento",row:2,input:{show:false}},
    {code:"LIMP",desc:"Limpieza",row:2,input:{show:false}},
    {code:"Perm",desc:"Permiso",row:2,input:{show:false}},
    {code:"AL",desc:"Ayuda Logística",row:3,input:{show:false}},
    {code:"PR",desc:"Paré Carga Rollo",row:3,input:{show:false}},
    {code:"CM",desc:"Cambiar Matriz",row:4,input:{show:false}},
    {code:"PM",desc:"Paré Matriz",row:4,input:{show:false}},
    {code:"RM",desc:"Rotura Matriz",row:4,input:{show:false}},
    {code:"REM",desc:"Reparando Matriz",row:4,input:{show:false}},
    {code:"PC",desc:"Paré Comida",row:3,input:{show:false}},
    {code:"RD",desc:"Rollo Fleje Doblado",row:3,input:{show:false}}
  ];

  const NON_DOWNTIME_CODES = new Set(["E","C","RM","PM","RD","LT"]);
  const isDowntime = (op) => !NON_DOWNTIME_CODES.has(op);

  const sameDowntime = (a,b) =>
    a && b &&
    String(a.opcion) === String(b.opcion) &&
    String(a.texto || "") === String(b.texto || "");

  let selected = null;

  /* ================= STORAGE POR LEGAJO ================= */
  function legajoKey() {
    return String(legajoInput.value || "").trim();
  }

  function stateKeyFor(legajo) {
    return `${LS_PREFIX}::${dayKeyAR()}::${String(legajo).trim()}`;
  }
  
  function freshState() {
    return {
      lastMatrix:null,
      lastCajon:null,
      lastDowntime:null,
      last2:[],
      lateArrivalSent:false,
      lateArrivalDiscarded:false,
      // ✅ NUEVO: no permite nuevo E hasta que haya al menos 1 C
      matrixNeedsC:false,
      pcDone: false // ✅
    };
  }

 function readStateForLegajo(legajo) {
    try {
      const raw = localStorage.getItem(stateKeyFor(legajo));
      if (!raw) return freshState();
  
      const s = JSON.parse(raw);
      if (!s || typeof s !== "object") return freshState();
  
      s.last2 = Array.isArray(s.last2) ? s.last2 : [];
      // ✅ Normaliza items viejos (compatibilidad con versiones anteriores)
      s.last2 = s.last2.map(it => {
        if (!it || typeof it !== "object") return it;
        return {
          id: it.id || "", // antes no existía
          opcion: it.opcion || "",
          descripcion: it.descripcion || "",
          texto: it.texto || "",
          ts: it.ts || it.tsEvent || "",
          status: it.status || "sent", // lo viejo asumimos enviado
          tries: Number(it.tries || 0),
          lastError: it.lastError || "",
          sentAt: it.sentAt || "",
          failedAt: it.failedAt || ""
        };
      });
      s.lastMatrix = s.lastMatrix || null;
      s.lastCajon = s.lastCajon || null;
      s.lastDowntime = s.lastDowntime || null;
      s.lateArrivalSent = !!s.lateArrivalSent;
      s.lateArrivalDiscarded = !!s.lateArrivalDiscarded;
      s.matrixNeedsC = !!s.matrixNeedsC;
      s.pcDone = !!s.pcDone;
  
      return s;
    } catch {
      return freshState();
    }
  }
  
  function writeStateForLegajo(legajo, state) {
    localStorage.setItem(stateKeyFor(legajo), JSON.stringify(state));
  }
  function updateHistoryItem(legajo, eventId, patch) {
    if (!legajo || !eventId) return;
  
    const s = readStateForLegajo(legajo);
    if (!Array.isArray(s.last2) || !s.last2.length) return;
  
    const idx = s.last2.findIndex(x => x && x.id === eventId);
    if (idx === -1) return;
  
    s.last2[idx] = { ...s.last2[idx], ...patch };
    writeStateForLegajo(legajo, s);
  }

  /* ================= COLA PENDIENTES ================= */
  function readQueue() {
    try {
      const raw = localStorage.getItem(LS_QUEUE);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }
  function writeQueue(arr) {
    localStorage.setItem(LS_QUEUE, JSON.stringify(arr)); // sin slice
  }
  function enqueue(payload) {
    const q = readQueue();
    q.push({ ...payload, __tries: 0 });
    try {
      writeQueue(q);
      // ✅ Registrar en historial del día como "queued" (pendiente)
      const leg = String(payload.legajo || "").trim();
      if (leg) {
        const s = readStateForLegajo(leg);
        pushLast2(s, payload, "queued", { tries: 0 });
        writeStateForLegajo(leg, s);
      }
    } catch (e) {
      alert("⚠️ Sin espacio local para guardar la cola. Avisar a Sistemas.");
      console.error("QUEUE WRITE FAILED (QuotaExceeded):", e);
      // opcional: revertir push si querés
    }
  }
  function queueLength() {
    return readQueue().length;
  }
    /* ================= COLA FALLIDOS (no bloquear la cola) ================= */
  function readFailed() {
    try {
      return JSON.parse(localStorage.getItem(LS_FAILED) || "[]");
    } catch {
      return [];
    }
  }

  function writeFailed(arr) {
    localStorage.setItem(LS_FAILED, JSON.stringify(arr.slice(-200)));
  }

  function moveToFailed(item, reason) {
    const f = readFailed();
    f.push({
      ...item,
      __failedAt: isoNowSeconds(),
      __reason: String(reason || "")
    });
    writeFailed(f);
  }


  /* ================= UI: RESUMEN ================= */
  function renderSummary() {
    const leg = legajoKey();

    if (!leg) {
      daySummary.className = "history-empty";
      daySummary.innerText = "Ingresá tu legajo para ver el resumen";
      return;
    }

    const s = readStateForLegajo(leg);
    const qLen = queueLength();

    const renderItem = (title, item) => {
      if (!item) return `<div class="day-item"><div class="t1">${title}</div><div class="t2">—</div></div>`;
      return `
        <div class="day-item">
          <div class="t1">${title}</div>
          <div class="t2">
            ${item.opcion} — ${item.descripcion}<br>
            ${item.texto ? `Dato: <b>${item.texto}</b><br>` : ""}
            ${item.ts ? `Fecha: ${formatDateTimeAR(item.ts)}` : ""}
          </div>
        </div>`;
    };

    const renderHistory = (arr) => {
    if (!arr || !arr.length) {
      return `<div class="day-item"><div class="t1">Historial del día</div><div class="t2">—</div></div>`;
    }
  
    const badge = (st) => {
      const s = String(st || "").toLowerCase();
      if (s === "sent")   return `<span style="padding:2px 8px;border-radius:999px;background:#e8fff0;color:#0b6b2c;font-weight:800;font-size:12px;">ENVIADO</span>`;
      if (s === "queued") return `<span style="padding:2px 8px;border-radius:999px;background:#fff7e6;color:#8a5a00;font-weight:800;font-size:12px;">PENDIENTE</span>`;
      if (s === "failed") return `<span style="padding:2px 8px;border-radius:999px;background:#ffecec;color:#9b1c1c;font-weight:800;font-size:12px;">ERROR</span>`;
      if (s === "dead")   return `<span style="padding:2px 8px;border-radius:999px;background:#eee;color:#444;font-weight:800;font-size:12px;">NO ENVIADO</span>`;
      return `<span style="padding:2px 8px;border-radius:999px;background:#eee;color:#444;font-weight:800;font-size:12px;">${s.toUpperCase()}</span>`;
    };
  
    return `
      <div class="day-item">
        <div class="t1">Historial del día (${arr.length})</div>
        <div class="t2" style="max-height:360px; overflow:auto; padding-right:6px;">
          ${arr.map(it => `
            <div style="margin-top:10px; padding-bottom:10px; border-bottom:1px solid rgba(0,0,0,.08);">
              <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
                <span style="font-weight:900; font-size:34px;">
                  ${it.opcion}${it.texto ? `: ${it.texto}` : ""}
                </span>
                ${badge(it.status)}
                ${it.tries ? `<span style="font-size:12px; color:#666;">intentos: ${it.tries}</span>` : ""}
              </div>
  
              ${it.descripcion ? `<div>— ${it.descripcion}</div>` : ""}
              ${it.ts ? `<div style="color:#555;">Evento: ${formatDateTimeAR(it.ts)}</div>` : ""}
  
              ${it.sentAt ? `<div style="color:#0b6b2c;">Enviado: ${formatDateTimeAR(it.sentAt)}</div>` : ""}
              ${it.failedAt ? `<div style="color:#9b1c1c;">Último error: ${formatDateTimeAR(it.failedAt)}</div>` : ""}
              ${it.lastError ? `<div style="color:#9b1c1c; font-size:12px;">${it.lastError}</div>` : ""}
            </div>
          `).join("")}
        </div>
      </div>`;
  };


    daySummary.className = "";
    daySummary.innerHTML = [
      qLen ? `<div class="day-item"><div class="t1">Pendientes de envío</div><div class="t2"><b>${qLen}</b></div></div>` : "",
      renderHistory(s.last2)
    ].join("");

  }

  function renderMatrizInfoForCajon() {
    const leg = legajoKey();
    if (!leg || !selected || selected.code !== "C") {
      matrizInfo.classList.add("hidden");
      matrizInfo.innerHTML = "";
      return;
    }

    const s = readStateForLegajo(leg);
    const lm = s.lastMatrix;

    matrizInfo.classList.remove("hidden");
    if (!lm || !lm.texto) {
      matrizInfo.innerHTML = `⚠️ No hay matriz registrada hoy.<br><small>Enviá primero "E (Empecé Matriz)"</small>`;
      return;
    }

    matrizInfo.innerHTML =
      `Matriz en uso: <span style="font-size:22px;">${lm.texto}</span>
       <small>Última matriz: ${lm.ts ? formatDateTimeAR(lm.ts) : ""}</small>`;
  }

  /* ================= BLOQUEO UI POR TM PENDIENTE ================= */
  function getPendingDowntime() {
    const leg = legajoKey();
    if (!leg) return null;
    const s = readStateForLegajo(leg);
    return s.lastDowntime || null;
  }

  function isAllowedWhenPending(optCode, pending) {
    if (!pending) return true;
    return String(optCode) === String(pending.opcion);
  }



  // ✅ NUEVO: regla de matriz -> bloquea E hasta que exista al menos 1 C después del último E
  function isAllowedByMatrixRule(optCode, state) {
    if (optCode !== "E") return true;
    if (!state) return true;
    return !state.matrixNeedsC;
  }

  function applyDisabledStyle(el, disabled) {
    if (!disabled) {
      el.style.opacity = "";
      el.style.pointerEvents = "";
      el.style.filter = "";
      return;
    }
    el.style.opacity = "0.35";
    el.style.pointerEvents = "none";
    el.style.filter = "grayscale(100%)";
  }

  function renderOptions() {
    row1.innerHTML=""; row2.innerHTML=""; row3.innerHTML=""; row4.innerHTML="";
    const pending = getPendingDowntime();

    const leg = legajoKey();
    const state = leg ? readStateForLegajo(leg) : null;

    OPTIONS.forEach(o=>{
      const d=document.createElement("div");
      d.className="box";
      d.dataset.code = o.code;
      d.innerHTML=`<div class="box-title">${o.code}</div><div class="box-desc">${o.desc}</div>`;

      const allowedPending = isAllowedWhenPending(o.code, pending);
      const allowedMatrix  = isAllowedByMatrixRule(o.code, state);
      const allowed = allowedPending && allowedMatrix;

      if (!allowed) {
        applyDisabledStyle(d, true);
      } else {
        d.addEventListener("click",()=>selectOption(o, d));
      }

      const target =
        o.row === 1 ? row1 :
        o.row === 2 ? row2 :
        o.row === 3 ? row3 :
        row4; // row 4
      
      target.appendChild(d);

    });

    // Mensaje por regla de matriz (solo si no hay TM pendiente)
    if (!pending && state && state.matrixNeedsC) {
      error.style.color = "#b26a00";
      error.innerText = `⚠️ Para iniciar una nueva matriz (E), primero tenes que terminar la cantidad que hiciste en la matriz en curso.`;
    }

    if (pending) {
      const opt = OPTIONS.find(x => x.code === pending.opcion);
      if (opt) {
        const el = document.querySelector(`.box[data-code="${opt.code}"]`);
        selectOption(opt, el);
        btnResetSelection.style.opacity = "0.4";
        btnResetSelection.style.pointerEvents = "none";
        error.style.color = "#b26a00";
        error.innerText =
          `⚠️ Hay un Tiempo Muerto pendiente (${pending.opcion}). ` +
          `Solo podés reenviar el MISMO para cerrarlo.`;
      }
    } else {
      btnResetSelection.style.opacity = "";
      btnResetSelection.style.pointerEvents = "";
      if (!selected) {
        // no pisar el mensaje de matriz
        if (!(state && state.matrixNeedsC)) {
          error.style.color = "";
          error.innerText = "";
        }
      }
    }
  }

  /* ================= NAVEGACIÓN ================= */
  function goToOptions() {
    if (!legajoKey()) { alert("Ingresá el número de legajo"); return; }
    legajoScreen.classList.add("hidden");
    optionsScreen.classList.remove("hidden");

    renderOptions();
    renderMatrizInfoForCajon();
  }

  function backToLegajo() {
    optionsScreen.classList.add("hidden");
    legajoScreen.classList.remove("hidden");
    renderSummary();
  }

  /* ================= SELECCIÓN ================= */
  function selectOption(opt, elBox) {
    selected = opt;

    document.querySelectorAll(".box.selected").forEach(x => x.classList.remove("selected"));
    if (elBox) elBox.classList.add("selected");
    else {
      const found = document.querySelector(`.box[data-code="${opt.code}"]`);
      if (found) found.classList.add("selected");
    }

    selectedArea.classList.remove("hidden");
    selectedBox.innerText = opt.code;
    selectedDesc.innerText = opt.desc;

    const pending = getPendingDowntime();
    const leg = legajoKey();
    const state = leg ? readStateForLegajo(leg) : null;

    // Mensaje por regla de matriz (solo si no hay TM pendiente)
    if (!pending && state && state.matrixNeedsC) {
      error.style.color = "#b26a00";
      error.innerText = `⚠️ Para iniciar una nueva matriz (E), primero tenes que terminar la cantidad que hiciste en la matriz en curso.`;
    } else if (!pending) {
      error.style.color = "";
      error.innerText = "";
    }

    textInput.value = "";

    if (opt.input.show) {
      inputArea.classList.remove("hidden");
      inputLabel.innerText = opt.input.label;
      textInput.placeholder = opt.input.placeholder;
      if (pending && pending.opcion === opt.code && pending.texto) {
        textInput.value = String(pending.texto || "");
      }
    } else {
      inputArea.classList.add("hidden");
      textInput.placeholder = "";
    }

    renderMatrizInfoForCajon();
  }

  function resetSelection() {
    const pending = getPendingDowntime();
    if (pending) return;
    selected = null;
    selectedArea.classList.add("hidden");
    error.innerText = "";
    textInput.value = "";
    matrizInfo.classList.add("hidden");
    matrizInfo.innerHTML = "";
    document.querySelectorAll(".box.selected").forEach(x => x.classList.remove("selected"));
  }

  /* ================= REGLAS Hs Inicio ================= */
  function computeHsInicioForC(state) {
    if (state.lastCajon && state.lastCajon.ts) return state.lastCajon.ts;
    if (state.lastMatrix && state.lastMatrix.ts) return state.lastMatrix.ts;
    return "";
  }

  /* ================= VALIDACIÓN TM ================= */
  function validateBeforeSend(legajo, payload) {
    const s = readStateForLegajo(legajo);
    const ld = s.lastDowntime;

    if (!ld) return { ok:true };

    

    if (!sameDowntime(ld, payload)) {
      return {
        ok:false,
        msg:`Hay un "Tiempo Muerto" pendiente (${ld.opcion}${ld.texto ? " " + ld.texto : ""}).\n` +
            `Solo podés enviar el MISMO tiempo muerto para cerrarlo, o enviar RM / RD.`
      };
    }

    return { ok:true, isSecondSameDowntime:true, downtimeTs: ld.ts || "" };
  }

  /* ================= ACTUALIZAR ESTADO ================= */
  function pushLast2(s, payload, status = "queued", extra = {}) {
    const item = {
      id: payload.id,
      opcion: payload.opcion,
      descripcion: payload.descripcion,
      texto: payload.texto || "",
      ts: payload.tsEvent,      // momento del evento
      status,                  // queued | sent | failed | dead
      tries: extra.tries ?? 0,
      lastError: extra.lastError ?? "",
      sentAt: extra.sentAt ?? "",
      failedAt: extra.failedAt ?? ""
    };
  
    s.last2.unshift(item);
    s.last2 = s.last2.slice(0, MAX_DAY_HISTORY);
  }

  function updateStateAfterSend(legajo, payload) {
    const s = readStateForLegajo(legajo);

    if (payload.opcion === "LT") {
      
      writeStateForLegajo(legajo, s);
      return;
    }

    

    if (payload.opcion === "E") {
      if (s.lastMatrix && String(s.lastMatrix.texto||"") !== String(payload.texto||"")) {
        s.lastCajon = null;
      }
      s.lastMatrix = { opcion:payload.opcion, descripcion:payload.descripcion, texto:payload.texto||"", ts:payload.tsEvent };
      s.lastDowntime = null;

      // ✅ NUEVO: al iniciar matriz, exige al menos un C antes de permitir otro E
      s.matrixNeedsC = true;

      writeStateForLegajo(legajo, s);
      return;
    }

    if (payload.opcion === "C") {
      s.lastCajon = { opcion:payload.opcion, descripcion:payload.descripcion, texto:payload.texto||"", ts:payload.tsEvent };
      s.lastDowntime = null;

      // ✅ NUEVO: con el primer C luego de E, ya habilita nuevamente E
      s.matrixNeedsC = false;

      writeStateForLegajo(legajo, s);
      return;
    }

    if (payload.opcion === "RM" || payload.opcion === "PM" || payload.opcion === "RD") {
      s.lastDowntime = null;
      writeStateForLegajo(legajo, s);
      return;
    }

    if (isDowntime(payload.opcion)) {
      const item = { opcion:payload.opcion, descripcion:payload.descripcion, texto:payload.texto||"", ts:payload.tsEvent };
      if (!s.lastDowntime) s.lastDowntime = item;
      else if (sameDowntime(s.lastDowntime, payload)) s.lastDowntime = null;
      else s.lastDowntime = item;
      writeStateForLegajo(legajo, s);
      return;
    }

    writeStateForLegajo(legajo, s);
  }

  /* ================= ENVÍO (CON CONFIRMACIÓN REAL) ================= */
  function sleep(ms){ return new Promise(r=>setTimeout(r,ms)); }

  async function postToSheet(payload) {
    const res = await fetch(GOOGLE_SHEET_WEBAPP_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      mode: "cors",
      cache: "no-store"
    });
  
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
  
    const data = await res.json().catch(() => null);
  
    // ✅ si Apps Script devuelve {ok:false,error:"..."} lo vas a conservar como error
    if (!data || data.ok !== true) {
      throw new Error(data?.error || "Respuesta inválida del WebApp");
    }
  
    return data;
  }
  

  let isFlushing = false;

  async function flushQueueOnce() {
    if (isFlushing) return;
    if (!navigator.onLine) return;
    isFlushing = true;

    try {
      let q = readQueue();
      if (!q.length) return;

      const batchMax = 15;

      for (let processed=0; processed<batchMax; processed++) {
        q = readQueue();
        if (!q.length) break;

        const item = q[0];
        const tries = Number(item.__tries || 0);
        try {
          await postToSheet(item); // ✅ ahora espera OK real
          q.shift();
          writeQueue(q);
          await sleep(140);
          // ✅ Marcar como enviado
          updateHistoryItem(item.legajo, item.id, {
            status: "sent",
            sentAt: isoNowSeconds(),
            tries: Number(item.__tries || 0),
            lastError: "",
            failedAt: ""
          });
        } catch (err) {
          console.warn("postToSheet failed:", err);
        
          item.__tries = tries + 1;
          q[0] = item;
          writeQueue(q);
          // ✅ Marcar como error (sigue pendiente en cola)
          updateHistoryItem(item.legajo, item.id, {
            status: "failed",
            failedAt: isoNowSeconds(),
            tries: Number(item.__tries || 0),
            lastError: String(err?.message || err)
          });
        
          // backoff progresivo (máximo 60s)
          const backoff = Math.min(1000 * item.__tries, 60000);
          await sleep(backoff);
          break; // corta el batch y reintenta en el próximo ciclo
        }
      }
    } finally {
      isFlushing = false;
      renderSummary();
      console.log("QUEUE LEN:", readQueue().length, "FAILED LEN:", readFailed().length);
    }
  }

  /* ================= LLEGADA TARDE ================= */
  function maybeSendLateArrival(legajo) {
    const s = readStateForLegajo(legajo);

    const isFirstMessage = (!s.last2 || s.last2.length === 0)
      && !s.lastMatrix && !s.lastCajon && !s.lastDowntime;

   if (!isFirstMessage) return false;
    
  // ✅ si ya fue enviado o descartado, no hacer nada
  if (s.lateArrivalSent || s.lateArrivalDiscarded) return false;
  
  const nowMin = nowMinutesAR();
  const limitMin = 8 * 60 + 30;
  
  // ✅ Si el primer mensaje fue ANTES o A LAS 08:30, se descarta para todo el día
  if (nowMin <= limitMin) {
    s.lateArrivalDiscarded = true;
    writeStateForLegajo(legajo, s);
    return false;
  }


    const day = dayKeyAR();
    const hsInicioISO = `${day}T08:30:00-03:00`;

    const tsEvent = isoNowSeconds();
    const latePayload = {
      id: uuidv4(),
      legajo,
      opcion: "LT",
      descripcion: "LLegada Tarde",
      texto: "",
      tsEvent,
      "Hs Inicio": hsInicioISO,
      matriz: ""
    };

    s.lateArrivalSent = true;
    writeStateForLegajo(legajo, s);

    updateStateAfterSend(legajo, latePayload);
    enqueue(latePayload);

    return true;
  }

  async function sendFast() {
    if (!selected) return;

    const legajo = legajoKey();
    if (!legajo) { alert("Ingresá el número de legajo"); return; }

    maybeSendLateArrival(legajo);

    const texto = String(textInput.value || "").trim();

   // ✅ Validación dinámica según matriz (solo afecta a C)
    if (selected.input.show) {
      const stateTmp = readStateForLegajo(legajo);
    
      let ok = true;
      if (selected.code === "C" && isMatrix501(stateTmp)) {
        // Matriz 501: permite 123,5 o 123.5
        ok = /^\d+(?:[.,]\d+)?$/.test(texto);
      } else {
        // Resto (incluye E y C no-501): solo enteros
        ok = /^[0-9]+$/.test(texto);
      }
    
      if (!ok) {
        error.style.color = "red";
        error.innerText = (selected.code === "C" && isMatrix501(stateTmp))
          ? "Para matriz 501: usar coma o punto (ej: 12,5 o 12.5)"
          : "Solo se permiten números enteros";
        return;
      }
    }


    const tsEvent = isoNowSeconds();
    const stateBefore = readStateForLegajo(legajo);

    // ✅ NUEVO: no permitir nuevo E si falta al menos un C luego del E anterior
    if (selected.code === "E" && stateBefore.matrixNeedsC) {
      alert('Antes de iniciar una nueva matriz (E), tenés que enviar al menos 1 Cajón (C) para cerrar la matriz anterior.');
      return;
    }

    // ✅ Si es Cajón (C) y matriz 501, normaliza punto->coma antes de guardar/enviar
    let textoToSend = texto;
    if (selected.code === "C" && isMatrix501(stateBefore)) {
      textoToSend = normalizeToComma(texto);
    }
    
    const payload = {
      id: uuidv4(),
      legajo,
      opcion: selected.code,
      descripcion: selected.desc,
      texto: textoToSend,
      tsEvent,
      "Hs Inicio": "",
      matriz: ""
    };


    if (payload.opcion === "C" || payload.opcion === "RM" || payload.opcion === "PM" || payload.opcion === "RD") {
      if (!stateBefore.lastMatrix || !stateBefore.lastMatrix.ts || !stateBefore.lastMatrix.texto) {
        alert('Primero tenés que enviar "E (Empecé Matriz)" para registrar una matriz.');
        return;
      }
      payload.matriz = String(stateBefore.lastMatrix.texto || "").trim();
    }

    if (payload.opcion === "C") {
      payload["Hs Inicio"] = computeHsInicioForC(stateBefore);
    }

    if (payload.opcion === "RM" || payload.opcion === "PM" || payload.opcion === "RD") {
      payload["Hs Inicio"] = tsEvent;
    }


    const v = validateBeforeSend(legajo, payload);
    if (!v.ok) { alert(v.msg); return; }

    if (v.isSecondSameDowntime) {
      payload["Hs Inicio"] = v.downtimeTs || "";
    }

    btnEnviar.disabled = true;
    const prev = btnEnviar.innerText;
    btnEnviar.innerText = "Enviando...";

    updateStateAfterSend(legajo, payload);
    renderSummary();

    selected = null;
    selectedArea.classList.add("hidden");
    optionsScreen.classList.add("hidden");
    legajoScreen.classList.remove("hidden");
    matrizInfo.classList.add("hidden");
    matrizInfo.innerHTML = "";
    // ✅ no pisar mensaje cuando volvés (el resumen ya lo muestra)
    error.innerText = "";
    document.querySelectorAll(".box.selected").forEach(x => x.classList.remove("selected"));

    enqueue(payload);
    await flushQueueOnce();

    btnEnviar.disabled = false;
    btnEnviar.innerText = prev;
  }

  /* ================= EVENTOS ================= */
  btnContinuar.addEventListener("click", goToOptions);
  btnBackTop.addEventListener("click", backToLegajo);
  btnBackLabel.addEventListener("click", backToLegajo);
  btnResetSelection.addEventListener("click", resetSelection);

  btnEnviar.addEventListener("click", sendFast);
  legajoInput.addEventListener("keydown", (e)=>{ if(e.key==="Enter") goToOptions(); });

  let legajoTimer = null;
  legajoInput.addEventListener("input", () => {
    clearTimeout(legajoTimer);
    legajoTimer = setTimeout(renderSummary, 120);
  });

  window.addEventListener("focus", () => flushQueueOnce());
  window.addEventListener("online", () => flushQueueOnce());
  setInterval(() => flushQueueOnce(), 5000);


  /* ================= INIT ================= */
  renderOptions();
  renderSummary();

  console.log("app.js OK ✅ (bloqueo E hasta C + confirmación real + reset diario + keys _Cervantes_v1)");

});
