/* Multi-IA para Excel — GPT + Claude + Gemini con tus propias API keys */

const DEFAULT_MODELS = {
  openai: ["gpt-6.1-sol", "gpt-6-luna"],
  anthropic: ["claude-sonnet-5-5", "claude-opus-5-5", "claude-haiku-5-5"],
  gemini: ["gemini-3.8-flash", "gemini-3.5-flash-lite"],
};
const PROVIDER_LABEL = { openai: "OpenAI", anthropic: "Claude", gemini: "Gemini" };
const MAX_CTX_ROWS = 200;
const MAX_CTX_COLS = 30;
const REQUEST_TIMEOUT_MS = 60000;

const $ = (id) => document.getElementById(id);
let stopFlag = false;

/* ---------- almacenamiento local ---------- */
function store(key, value) {
  try { localStorage.setItem("multiia." + key, value); } catch (e) {}
}
function load(key, fallback = "") {
  try { return localStorage.getItem("multiia." + key) ?? fallback; } catch (e) { return fallback; }
}
function getModels(provider) {
  const raw = load("models." + provider, "");
  const list = raw.split("\n").map((s) => s.trim()).filter(Boolean);
  return list.length ? list : DEFAULT_MODELS[provider];
}

/* ---------- UI ---------- */
function showMsg(el, text, isError = false) {
  el.textContent = text;
  el.classList.toggle("error", isError);
}

function buildModelSelect() {
  const sel = $("modelSelect");
  const saved = load("selectedModel", "");
  sel.innerHTML = "";
  for (const provider of Object.keys(DEFAULT_MODELS)) {
    const group = document.createElement("optgroup");
    group.label = PROVIDER_LABEL[provider];
    for (const m of getModels(provider)) {
      const opt = document.createElement("option");
      opt.value = provider + "|" + m;
      opt.textContent = m;
      group.appendChild(opt);
    }
    sel.appendChild(group);
  }
  if ([...sel.options].some((o) => o.value === saved)) {
    sel.value = saved;
  } else {
    // Sin selección guardada: elegir el primer modelo de un proveedor que ya tenga key.
    const withKey = [...sel.options].find((o) => load("key." + o.value.split("|")[0]).trim());
    if (withKey) sel.value = withKey.value;
  }
  updateAgentMode();
}

function updateAgentMode() {
  const badge = $("agentMode");
  if (!badge) return;
  const isGemini = $("modelSelect").value.startsWith("gemini|");
  badge.textContent = isGemini ? "Gemini · herramientas nativas · reversión automática" : "Modo JSON · reversión automática";
  badge.classList.toggle("gemini", isGemini);
}

function loadSettingsUI() {
  $("keyOpenai").value = load("key.openai");
  $("keyAnthropic").value = load("key.anthropic");
  $("keyGemini").value = load("key.gemini");
  $("modelsOpenai").value = getModels("openai").join("\n");
  $("modelsAnthropic").value = getModels("anthropic").join("\n");
  $("modelsGemini").value = getModels("gemini").join("\n");
}

function setupTabs() {
  document.querySelectorAll(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      $("tab-" + btn.dataset.tab).classList.add("active");
    });
  });
}

/* ---------- llamadas a los proveedores ---------- */
async function readError(res) {
  let detail = "";
  try { detail = await res.text(); } catch (e) {}
  return `Error ${res.status}: ${detail.slice(0, 400)}`;
}

// Peticiones en curso, para que los botones "Detener" puedan abortarlas.
const activeRequests = new Set();

function cancelRequests() {
  for (const ctrl of activeRequests) {
    ctrl.cancelledByUser = true;
    ctrl.abort();
  }
}

async function fetchWithTimeout(url, options, ms = REQUEST_TIMEOUT_MS) {
  const ctrl = new AbortController();
  activeRequests.add(ctrl);
  const timer = setTimeout(() => ctrl.abort(), ms);
  try {
    return await fetch(url, { ...options, signal: ctrl.signal });
  } catch (e) {
    if (e.name === "AbortError") {
      throw new Error(ctrl.cancelledByUser ? "Cancelado por el usuario." : `El proveedor no respondió en ${ms / 1000} s.`);
    }
    throw e;
  } finally {
    clearTimeout(timer);
    activeRequests.delete(ctrl);
  }
}

function getActiveModel() {
  const [provider, model] = $("modelSelect").value.split("|");
  const key = load("key." + provider).trim();
  if (!key) throw new Error(`Falta la API key de ${PROVIDER_LABEL[provider] || "este proveedor"}. Agrégala en Ajustes.`);
  return { provider, model, key };
}

async function callModel(system, user) {
  const { provider, model, key } = getActiveModel();

  if (provider === "openai") {
    const res = await fetchWithTimeout("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + key },
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: system },
          { role: "user", content: user },
        ],
      }),
    });
    if (!res.ok) throw new Error(await readError(res));
    const data = await res.json();
    return data.choices?.[0]?.message?.content ?? "";
  }

  if (provider === "anthropic") {
    const res = await fetchWithTimeout("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model,
        max_tokens: 8192,
        system,
        messages: [{ role: "user", content: user }],
      }),
    });
    if (!res.ok) throw new Error(await readError(res));
    const data = await res.json();
    return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
  }

  if (provider === "gemini") {
    // geminiGenerate / parseGeminiResponse viven en gemini-agent.js.
    const data = await geminiGenerate(model, key, {
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: "user", parts: [{ text: user }] }],
    });
    return parseGeminiResponse(data).text;
  }

  throw new Error("Proveedor desconocido");
}

function extractJSON(text) {
  const cleaned = text.replace(/```json|```/gi, "");
  const start = cleaned.search(/[\[{]/);
  if (start === -1) throw new Error("La respuesta no contiene JSON.");
  const open = cleaned[start];
  const close = open === "{" ? "}" : "]";
  const end = cleaned.lastIndexOf(close);
  if (end <= start) throw new Error("JSON incompleto en la respuesta.");
  return JSON.parse(cleaned.slice(start, end + 1));
}

/* ---------- utilidades de Excel ---------- */
const LARGE_SELECTION_ROWS = 500;

// Aviso pre-vuelo: la IA solo recibe MAX_CTX_ROWS filas como contexto.
function largeSelectionWarning(rows) {
  return rows > LARGE_SELECTION_ROWS
    ? `⚠ La selección tiene ${rows} filas con datos; la IA solo ve las primeras ${MAX_CTX_ROWS} como contexto. ` +
        "Para tablas grandes usa «Rellenar columna», que procesa por lotes."
    : "";
}

async function getSelectionContext() {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const ws = range.worksheet;
    const used = range.getUsedRangeOrNullObject(true);
    range.load(["address", "rowCount", "columnCount"]);
    ws.load("name");
    used.load("rowCount");
    await ctx.sync();

    const rows = Math.min(range.rowCount, MAX_CTX_ROWS);
    const cols = Math.min(range.columnCount, MAX_CTX_COLS);
    const limited = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    limited.load(["address", "formulas"]);
    await ctx.sync();

    return {
      sheet: ws.name,
      address: limited.address,
      truncated: rows < range.rowCount || cols < range.columnCount,
      dataRows: used.isNullObject ? 0 : used.rowCount,
      formulas: limited.formulas,
    };
  });
}

/* ---------- Agente ---------- */
const AGENT_SYSTEM = `Eres un asistente experto en Excel integrado en un panel lateral.
Recibes el contenido de las celdas seleccionadas (fórmulas o valores) y una petición.
Responde SOLO con un objeto JSON válido, sin texto adicional ni bloques de código, con esta forma:
{"message": "explicación breve en español", "writes": [{"start": "D2", "values": [["=B2+C2"],["=B3+C3"]], "overwrite": false}]}
Reglas:
- "start" es la celda superior izquierda (sin nombre de hoja) donde se escribe la matriz "values".
- "values" es una matriz 2D de filas; las fórmulas empiezan con "=" y usan sintaxis en inglés de Excel con separador de coma. Usa "_" para dejar una celda sin cambios.
- "overwrite": true solo si la petición pide modificar datos existentes (limpiar, traducir, normalizar, reemplazar); con false se conservan las celdas que ya tienen datos.
- Si solo hay que responder o explicar algo, devuelve "writes": [].`;

async function runAgent() {
  const prompt = $("agentPrompt").value.trim();
  if (!prompt) return showMsg($("agentOut"), "Escribe una petición primero.", true);

  $("agentRun").disabled = true;
  $("agentStop").disabled = false;
  $("agentApply").disabled = true;
  clearGeminiPending();
  showMsg($("agentOut"), "Pensando...");

  try {
    const { provider, model, key } = getActiveModel();
    if (provider === "gemini") {
      // Agente Gemini: function calling nativo + motor local atómico (gemini-agent.js).
      return await runGeminiAgent(prompt, $("agentUseSel").checked, model, key);
    }

    let userMsg = prompt;
    let sheetName;
    let warning = "";
    if ($("agentUseSel").checked) {
      const c = await getSelectionContext();
      sheetName = c.sheet;
      warning = largeSelectionWarning(c.dataRows);
      userMsg =
        `Hoja: ${c.sheet}\nRango: ${c.address}${c.truncated ? " (recortado al límite de contexto)" : ""}\n` +
        `Contenido (matriz 2D, fila por fila):\n${JSON.stringify(c.formulas)}\n\nPetición: ${prompt}`;
    } else {
      sheetName = await getActiveSheetName();
    }
    const raw = await callModel(AGENT_SYSTEM, userMsg);
    const parsed = extractJSON(raw);
    const writes = Array.isArray(parsed.writes) ? parsed.writes : [];

    // Las escrituras pasan por el mismo motor que Gemini (gemini-agent.js): snapshot previo,
    // reversión automática ante errores masivos y «Deshacer». Se aplican en la hoja del contexto.
    const actions = [];
    const rejected = [];
    for (const w of writes) {
      try {
        actions.push(
          buildAction({
            name: EXCEL_TOOL.name,
            args: { targetRange: w.start, operation: "set_formulas", values: w.values, overwrite: w.overwrite === true },
          })
        );
      } catch (e) {
        rejected.push(e.message);
      }
    }

    const out = [warning, parsed.message || "Listo."];
    if (actions.length) {
      out.push(`Cambios propuestos en ${sheetName}:\n` + actions.map((a, i) => describeAction(a, i + 1)).join("\n"));
    }
    if (rejected.length) out.push("⚠ Cambios descartados:\n" + rejected.map((m) => "• " + m).join("\n"));
    showMsg($("agentOut"), out.filter(Boolean).join("\n\n"));
    if (actions.length) setPendingActions(sheetName, actions);
  } catch (e) {
    showMsg($("agentOut"), e.message, true);
  } finally {
    $("agentRun").disabled = false;
    $("agentStop").disabled = true;
    setAgentStatus("");
  }
}

async function applyAgent() {
  if (hasGeminiPending()) return applyGeminiPending();
}

/* ---------- Rellenar columna ---------- */
const FILL_SYSTEM = `Eres un asistente que completa una columna nueva de una tabla de Excel.
Recibes los encabezados, un lote de filas y una instrucción.
Responde SOLO con un arreglo JSON de strings, uno por fila, en el mismo orden y con la misma longitud que el lote.
No agregues texto fuera del JSON.`;

async function runFill() {
  const header = $("fillHeader").value.trim();
  const prompt = $("fillPrompt").value.trim();
  const batch = Math.max(1, Math.min(100, parseInt($("fillBatch").value, 10) || 20));
  if (!header || !prompt) return showMsg($("fillOut"), "Completa el nombre de la columna y la instrucción.", true);

  stopFlag = false;
  discardGeminiUndo();
  $("fillRun").disabled = true;
  $("fillStop").disabled = false;

  let sel = null;
  let failure = null;
  const results = [];
  try {
    // Lectura. El rango se fija aquí: si el usuario hace clic en otra celda mientras se procesa,
    // el resultado se escribe igual en el lugar correcto.
    sel = await Excel.run(async (ctx) => {
      const selected = ctx.workbook.getSelectedRange();
      // Solo el bloque con datos dentro de la selección (evita cargar 1 millón de filas con columnas completas).
      const r = selected.getUsedRangeOrNullObject(true);
      selected.worksheet.load("name");
      r.load("isNullObject");
      await ctx.sync();
      if (r.isNullObject) return null;

      const right = r.getColumnsAfter(1);
      r.load(["rowCount", "columnCount", "values", "address"]);
      right.load(["address", "values"]);
      await ctx.sync();
      return {
        rows: r.rowCount,
        cols: r.columnCount,
        values: r.values,
        sheet: selected.worksheet.name,
        address: r.address.split("!").pop(),
        rightAddress: right.address.split("!").pop(),
        rightHasData: right.values.some((row) => row[0] !== ""),
      };
    });
    if (!sel) throw new Error("La selección está vacía.");
    if (sel.rows < 2) throw new Error("Selecciona al menos el encabezado y una fila de datos.");
    if (sel.rightHasData) {
      throw new Error(`La columna de la derecha (${sel.rightAddress}) ya tiene datos. Inserta una columna vacía o ajusta la selección.`);
    }

    // Red: los lotes se procesan en memoria; la hoja no se toca hasta el final.
    const headers = sel.values[0];
    const data = sel.values.slice(1);
    for (let i = 0; i < data.length && !stopFlag; i += batch) {
      const chunk = data.slice(i, i + batch);
      showMsg($("fillOut"), `Procesando filas ${i + 1}–${i + chunk.length} de ${data.length}...`);
      const userMsg =
        `Encabezados: ${JSON.stringify(headers)}\nNueva columna: "${header}"\nInstrucción: ${prompt}\n` +
        `Filas (${chunk.length}):\n${JSON.stringify(chunk)}`;
      const raw = await callModel(FILL_SYSTEM, userMsg);
      const arr = extractJSON(raw);
      if (!Array.isArray(arr) || arr.length !== chunk.length) {
        throw new Error(`El modelo devolvió ${Array.isArray(arr) ? arr.length : "un formato inválido"} elementos y se esperaban ${chunk.length}.`);
      }
      results.push(...arr.map((v) => [String(v ?? "")]));
    }
  } catch (e) {
    failure = e;
  }

  try {
    // Escritura única: encabezado + todas las filas obtenidas (también si se detuvo o falló a medias).
    if (sel && results.length) {
      await Excel.run(async (ctx) => {
        const r = ctx.workbook.worksheets.getItem(sel.sheet).getRange(sel.address);
        r.getCell(0, sel.cols).getResizedRange(results.length, 0).values = [[header], ...results];
        await ctx.sync();
      });
    }
    if (stopFlag) {
      showMsg($("fillOut"), `Detenido. Se llenaron ${results.length} filas.`);
    } else if (failure) {
      const partial = results.length ? `\nSe escribieron las ${results.length} filas completadas antes del error.` : "";
      showMsg($("fillOut"), failure.message + partial, true);
    } else {
      showMsg($("fillOut"), `Listo: ${results.length} filas completadas.`);
    }
  } catch (e) {
    showMsg($("fillOut"), "No se pudo escribir en la hoja: " + e.message, true);
  } finally {
    $("fillRun").disabled = false;
    $("fillStop").disabled = true;
  }
}

/* ---------- Ajustes ---------- */
function saveSettings() {
  store("key.openai", $("keyOpenai").value.trim());
  store("key.anthropic", $("keyAnthropic").value.trim());
  store("key.gemini", $("keyGemini").value.trim());
  store("models.openai", $("modelsOpenai").value);
  store("models.anthropic", $("modelsAnthropic").value);
  store("models.gemini", $("modelsGemini").value);
  buildModelSelect();
  showMsg($("settingsOut"), "Ajustes guardados.");
}

function clearKeys() {
  ["openai", "anthropic", "gemini"].forEach((p) => store("key." + p, ""));
  loadSettingsUI();
  showMsg($("settingsOut"), "API keys borradas de este equipo.");
}

/* ---------- arranque ---------- */
Office.onReady(() => {
  setupTabs();
  buildModelSelect();
  loadSettingsUI();
  $("modelSelect").addEventListener("change", () => {
    store("selectedModel", $("modelSelect").value);
    updateAgentMode();
  });
  $("agentRun").addEventListener("click", runAgent);
  $("agentStop").addEventListener("click", cancelRequests);
  $("agentApply").addEventListener("click", applyAgent);
  $("agentUndo").addEventListener("click", undoGemini);
  $("fillRun").addEventListener("click", runFill);
  $("fillStop").addEventListener("click", () => {
    stopFlag = true;
    cancelRequests();
  });
  $("saveSettings").addEventListener("click", saveSettings);
  $("clearKeys").addEventListener("click", clearKeys);
});
