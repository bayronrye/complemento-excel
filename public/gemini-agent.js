/* Agente Gemini para Multi-IA (especificación "Antigravity").
   - Function calling nativo de Gemini con la herramienta execute_excel_action.
   - Motor local atómico: 1 sync de lectura, 1 sync de escritura + verificación.
   - Snapshot previo, reversión automática ante errores masivos y botón Deshacer.
   Usa helpers globales de taskpane.js: $, showMsg, readError, extractJSON,
   fetchWithTimeout, MAX_CTX_ROWS, MAX_CTX_COLS. */

const GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta";
const EMPTY_TOKEN = "_";             // celdas vacías en el payload / "no tocar" en values
const MAX_TARGET_CELLS = 100000;     // tope de celdas por acción
const MAX_CLASSIFY_ROWS = 2000;      // tope de filas por classify_batch (costo de inferencia)
const CLASSIFY_BATCH = 25;           // filas por solicitud de clasificación
const CLASSIFY_CONCURRENCY = 3;      // solicitudes de clasificación en paralelo
const ERROR_ROLLBACK_RATIO = 0.25;   // revertir si >= 25 % de las celdas escritas terminan en error
const SAMPLE_SIZE = 3;
const OPERATIONS = ["fill_blanks_above", "set_formulas", "set_values", "classify_batch"];
const OP_LABEL = {
  fill_blanks_above: "Relleno de blancos ↓",
  set_formulas: "Fórmulas",
  set_values: "Valores",
  classify_batch: "Clasificación",
};

/* ---------- herramienta (spec §4 + parámetros de datos) ---------- */
const EXCEL_TOOL = {
  name: "execute_excel_action",
  description: "Aplica transformaciones de datos o inyección de fórmulas en la hoja activa",
  parameters: {
    type: "object",
    properties: {
      targetRange: { type: "string", description: "Coordenadas del rango destino (ej: C7:D24), sin nombre de hoja" },
      operation: { type: "string", enum: OPERATIONS },
      fillPattern: {
        type: "string",
        enum: ["relative_above", "static_repeat"],
        description: "Estrategia para rellenar celdas en blanco intermedias",
      },
      headerRowIncluded: {
        type: "boolean",
        description: "true si la primera fila de targetRange (y de sourceRange) es cabecera; no se modifica",
      },
      values: {
        type: "array",
        description:
          'set_values / set_formulas: matriz 2D fila por fila del tamaño de targetRange. Fórmulas en inglés con coma. "_" = dejar la celda sin cambios.',
        items: { type: "array", items: { type: "string" } },
      },
      overwrite: {
        type: "boolean",
        description:
          "true solo si el usuario pidió explícitamente modificar datos existentes (limpiar, normalizar, reemplazar). Por defecto se conservan las celdas con datos.",
      },
      sourceRange: { type: "string", description: "classify_batch: filas a clasificar (ej: A2:D50)" },
      categories: {
        type: "array",
        items: { type: "string" },
        description: "classify_batch: lista cerrada de categorías",
      },
      instruction: { type: "string", description: "classify_batch: criterio de clasificación" },
    },
    required: ["targetRange", "operation"],
  },
};

/* ---------- system prompt (spec §5 literal + reglas de herramienta) ---------- */
const GEMINI_AGENT_SYSTEM = `Eres el copiloto autónomo de Excel en el entorno Antigravity. Tu función principal es limpiar, normalizar, clasificar y propagar fórmulas en tablas complejas.

DIRECTRICES OPERATIVAS:
1. Análisis de Jerarquías:
   Al recibir rangos con columnas anidadas (ej. TURNO primario, RESPONSABLE secundario):
   - Respeta las cabeceras intactas.
   - Identifica celdas ancla (las que tienen valores explícitos, ej: MAÑANA, Carlos).
   - Identifica las celdas intermedias vacías dependientes.

2. Relleno Hacia Abajo (Forward Fill / Especial Celdas en Blanco):
   - Si se solicita propagar categorías hacia abajo sobre celdas vacías (ej: C7:D24), no devuelvas matrices gigantescas de texto.
   - Emula la función nativa de Excel: F5 -> Especial -> Celdas en blanco con fórmula relativa ascendente (=C[fila-1]).
   - Define el método como "fill_blanks_above" para que el cliente aplique las referencias relativas en memoria con cero overhead.

3. Integridad de Datos:
   - Jamás sobreescribas celdas que ya contengan datos válidos.
   - No extiendas las operaciones más allá de la última fila con datos de las columnas adyacentes.
   - Entrega siempre un informe conciso: Rango alterado, patrón aplicado y validación de muestra.

USO DE LA HERRAMIENTA execute_excel_action:
- Para modificar la hoja llama a execute_excel_action (puedes llamarla varias veces). Si solo hay que responder o explicar, contesta en texto sin llamarla.
- El contexto llega como "número de fila: [celdas]"; "_" significa celda vacía. targetRange y sourceRange van sin nombre de hoja.
- La "Cabecera previa" está fuera de la selección: úsala para entender las columnas, nunca la incluyas en targetRange.
- fill_blanks_above: targetRange = columnas a rellenar, empezando en la primera celda ancla o en la cabecera (en ese caso headerRowIncluded=true). fillPattern "relative_above" (por defecto, fórmula a la celda superior) o "static_repeat" (copia el valor). El cliente solo toca celdas vacías y se detiene en la última fila con datos.
- set_formulas / set_values: "values" es una matriz 2D del tamaño de targetRange. Fórmulas en inglés con separador coma (=SUM(B2,C2)). Usa "_" para dejar una celda sin cambios. overwrite=true solo si el usuario pidió explícitamente modificar datos existentes.
- classify_batch: targetRange = celda superior de la columna de salida; sourceRange = filas a clasificar; categories = lista cerrada; instruction = criterio. No envíes "values": el cliente clasifica por lotes.
- Junto con la llamada escribe en texto el informe conciso (rango, patrón, qué se espera ver).`;

const CLASSIFY_SYSTEM = `Eres un clasificador de filas de Excel. Recibes un criterio, las categorías permitidas, las cabeceras (si existen) y un lote de filas ("_" = celda vacía).
Devuelve SOLO un arreglo JSON de strings con exactamente una etiqueta por fila, en el mismo orden y con la misma longitud que el lote.`;

/* ---------- estado ---------- */
let geminiPending = null; // { sheet, actions }
let geminiUndo = null;    // { sheet, runs }

function hasGeminiPending() { return !!geminiPending; }
function clearGeminiPending() { geminiPending = null; }
// También lo usa el modo JSON (GPT / Claude) de taskpane.js para pasar por el mismo motor.
function setPendingActions(sheet, actions) {
  geminiPending = { sheet, actions };
  $("agentApply").disabled = false;
}
function discardGeminiUndo() {
  geminiUndo = null;
  const btn = $("agentUndo");
  if (btn) btn.disabled = true;
}

/* ---------- estado visual no bloqueante (spec §2.3) ---------- */
function setAgentStatus(text) {
  requestAnimationFrame(() => {
    const el = $("agentStatus");
    if (!el) return;
    el.textContent = text || "";
    el.classList.toggle("busy", !!text);
  });
}
// Cede un frame para que la UI pinte antes de trabajo pesado (con respaldo si el panel está oculto).
function nextFrame() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => resolve());
    setTimeout(resolve, 50);
  });
}

/* ---------- utilidades A1 ---------- */
function toExcelCol(index) {
  let col = "";
  while (index >= 0) {
    col = String.fromCharCode((index % 26) + 65) + col;
    index = Math.floor(index / 26) - 1;
  }
  return col;
}
function colToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
// "Hoja1!$C$7:D24" -> { row, col, rows, cols } (índices base 0) o null si es inválido.
function parseA1(address) {
  const a = String(address ?? "").split("!").pop().replace(/\$/g, "").trim().toUpperCase();
  const m = a.match(/^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/);
  if (!m) return null;
  const c1 = colToIndex(m[1]);
  const r1 = Number(m[2]) - 1;
  const c2 = m[3] ? colToIndex(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) - 1 : r1;
  const g = {
    row: Math.min(r1, r2),
    col: Math.min(c1, c2),
    rows: Math.abs(r2 - r1) + 1,
    cols: Math.abs(c2 - c1) + 1,
  };
  if (g.row < 0 || g.row + g.rows > 1048576 || g.col + g.cols > 16384) return null;
  return g;
}
function a1(row, col, rows = 1, cols = 1) {
  const first = toExcelCol(col) + (row + 1);
  return rows === 1 && cols === 1 ? first : `${first}:${toExcelCol(col + cols - 1)}${row + rows}`;
}
const a1Block = (b) => a1(b.row, b.col, b.rows, b.cols);

/* ---------- utilidades de matrices ---------- */
const compactCell = (v) => (v === "" || v === null || v === undefined ? EMPTY_TOKEN : v);
const makeMatrix = (rows, cols, fill) => Array.from({ length: rows }, () => new Array(cols).fill(fill));
const formatCell = (v) => (v === "" || v === null || v === undefined ? "(vacía)" : String(v));

// Una fila "parece cabecera" si tiene contenido y todo es texto no numérico y sin fórmulas.
function isHeaderLike(row) {
  const cells = (row || []).filter((v) => v !== "" && v !== null && v !== undefined);
  return (
    cells.length > 0 &&
    cells.every((v) => typeof v === "string" && !v.startsWith("=") && !/^[-+]?\d+([.,]\d+)?%?$/.test(v.trim()))
  );
}

/* ---------- API de Gemini ---------- */
async function geminiGenerate(model, key, body) {
  const res = await fetchWithTimeout(
    `${GEMINI_ENDPOINT}/models/${encodeURIComponent(model)}:generateContent`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": key },
      body: JSON.stringify(body),
    }
  );
  if (!res.ok) throw new Error(await readError(res));
  return res.json();
}

function parseGeminiResponse(data) {
  const cand = data.candidates?.[0];
  if (!cand) {
    const why = data.promptFeedback?.blockReason;
    throw new Error(why ? `Gemini bloqueó la solicitud (${why}).` : "Gemini no devolvió ninguna respuesta.");
  }
  const parts = cand.content?.parts || [];
  const text = parts
    .filter((p) => typeof p.text === "string" && !p.thought)
    .map((p) => p.text)
    .join("")
    .trim();
  const calls = parts.filter((p) => p.functionCall).map((p) => p.functionCall);
  if (!text && !calls.length) {
    throw new Error(`Gemini terminó sin contenido (${cand.finishReason || "motivo desconocido"}).`);
  }
  return { text, calls };
}

/* ---------- lectura del contexto (spec §2.1 / §2.2) ---------- */
async function readSelectionContext() {
  try {
    return await readSelectionFast();
  } catch (e) {
    // La selección empieza en la fila 1 (no hay fila superior), son columnas completas
    // o el bloque es demasiado grande: ruta segura con recorte previo.
    return await readSelectionSafe();
  }
}

// Ruta rápida: una sola sincronización (selección + geometría usada + ventana acotada).
// Solo se cargan como máximo MAX_CTX_ROWS × MAX_CTX_COLS celdas desde la esquina de la selección,
// así una selección de miles de filas no congela Excel. La fila previa (posible cabecera) solo se
// pide si existe, en una segunda lectura de una fila.
async function readSelectionFast() {
  return Excel.run(async (ctx) => {
    const sel = ctx.workbook.getSelectedRange();
    const ws = sel.worksheet;
    const used = sel.getUsedRangeOrNullObject(true);
    const view = sel.getCell(0, 0).getResizedRange(MAX_CTX_ROWS - 1, MAX_CTX_COLS - 1);
    sel.load(["address", "rowIndex", "columnIndex"]);
    ws.load("name");
    used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    view.load("formulas");
    await ctx.sync();

    if (used.isNullObject) return { sheet: ws.name, selection: sel.address.split("!").pop(), empty: true };
    const g = { row: used.rowIndex, col: used.columnIndex, rows: used.rowCount, cols: used.columnCount };
    const dr = g.row - sel.rowIndex;
    const dc = g.col - sel.columnIndex;
    const rows = Math.min(g.rows, MAX_CTX_ROWS);
    const cols = Math.min(g.cols, MAX_CTX_COLS);
    // Los datos empiezan lejos de la esquina de la selección: la ventana no los cubre.
    if (dr + rows > MAX_CTX_ROWS || dc + cols > MAX_CTX_COLS) throw new Error("fuera de la ventana");
    const formulas = view.formulas.slice(dr, dr + rows).map((r) => r.slice(dc, dc + cols));

    let aboveRow = null;
    if (dr === 0 && g.row > 0) {
      const above = ws.getRangeByIndexes(g.row - 1, g.col, 1, cols);
      above.load("formulas");
      await ctx.sync();
      aboveRow = above.formulas[0];
    }
    return buildContext(ws.name, sel.address, g, formulas, aboveRow);
  });
}

// Ruta segura: geometría primero, luego solo el bloque recortado.
async function readSelectionSafe() {
  return Excel.run(async (ctx) => {
    const sel = ctx.workbook.getSelectedRange();
    const ws = sel.worksheet;
    const used = sel.getUsedRangeOrNullObject(true);
    sel.load(["address", "rowIndex"]);
    ws.load("name");
    used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();

    if (used.isNullObject) return { sheet: ws.name, selection: sel.address.split("!").pop(), empty: true };
    const g = { row: used.rowIndex, col: used.columnIndex, rows: used.rowCount, cols: used.columnCount };
    const rows = Math.min(g.rows, MAX_CTX_ROWS);
    const cols = Math.min(g.cols, MAX_CTX_COLS);
    const block = ws.getRangeByIndexes(g.row, g.col, rows, cols);
    block.load("formulas");
    let above = null;
    if (g.row > 0 && g.row === sel.rowIndex) {
      above = ws.getRangeByIndexes(g.row - 1, g.col, 1, cols);
      above.load("formulas");
    }
    await ctx.sync();
    return buildContext(ws.name, sel.address, g, block.formulas, above ? above.formulas[0] : null);
  });
}

function buildContext(sheet, selection, g, formulas, aboveRow) {
  const rows = Math.min(g.rows, MAX_CTX_ROWS);
  const cols = Math.min(g.cols, MAX_CTX_COLS);
  const header =
    g.row > 0 && aboveRow && isHeaderLike(aboveRow.slice(0, cols))
      ? { row: g.row, cells: aboveRow.slice(0, cols).map(compactCell) } // row = número de fila (base 1)
      : null;
  return {
    sheet,
    selection: selection.split("!").pop(),
    block: g,
    rows,
    cols,
    truncated: rows < g.rows || cols < g.cols,
    data: formulas.slice(0, rows).map((r) => r.slice(0, cols).map(compactCell)),
    header,
  };
}

function formatContext(c) {
  const lines = [`Hoja: ${c.sheet}`, `Selección: ${c.selection}`];
  if (c.empty) {
    lines.push("La selección está vacía.");
    return lines.join("\n");
  }
  const g = c.block;
  lines.push(
    `Bloque con datos: ${a1Block(g)} (${g.rows}×${g.cols})` +
      (c.truncated ? ` — se envían solo ${c.rows}×${c.cols}` : "")
  );
  lines.push(`Columnas: ${Array.from({ length: c.cols }, (_, i) => toExcelCol(g.col + i)).join(", ")}`);
  if (c.header) {
    lines.push(`Cabecera previa (fila ${c.header.row}, fuera de la selección, solo lectura): ${JSON.stringify(c.header.cells)}`);
  }
  lines.push(`Contenido (número de fila: celdas; "${EMPTY_TOKEN}" = vacía):`);
  c.data.forEach((row, i) => lines.push(`${g.row + i + 1}: ${JSON.stringify(row)}`));
  return lines.join("\n");
}

async function getActiveSheetName() {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    await ctx.sync();
    return ws.name;
  });
}

/* ---------- validación de llamadas a la herramienta ---------- */
function normalizeMatrix(values) {
  if (!Array.isArray(values) || !values.length) return null;
  const rows = values.map((r) =>
    (Array.isArray(r) ? r : [r]).map((v) => (v === null || v === undefined ? EMPTY_TOKEN : String(v)))
  );
  const cols = Math.max(...rows.map((r) => r.length));
  if (!cols) return null;
  return rows.map((r) => {
    const copy = r.slice();
    while (copy.length < cols) copy.push(EMPTY_TOKEN);
    return copy;
  });
}

function buildAction(call) {
  if (call.name !== EXCEL_TOOL.name) throw new Error(`herramienta desconocida "${call.name}"`);
  const args = call.args || {};
  const op = args.operation;
  if (!OPERATIONS.includes(op)) throw new Error(`operación no soportada "${op}"`);
  const target = parseA1(args.targetRange);
  if (!target) throw new Error(`${op}: targetRange inválido "${args.targetRange ?? ""}"`);

  const action = {
    op,
    header: args.headerRowIncluded === true,
    overwrite: args.overwrite === true,
    requested: a1Block(target),
  };

  if (op === "fill_blanks_above") {
    action.pattern = args.fillPattern === "static_repeat" ? "static_repeat" : "relative_above";
    action.block = target;
  } else if (op === "classify_batch") {
    const source = parseA1(args.sourceRange);
    if (!source) throw new Error(`classify_batch: sourceRange inválido "${args.sourceRange ?? ""}"`);
    if (source.rows > MAX_CLASSIFY_ROWS) {
      throw new Error(`classify_batch: ${source.rows} filas superan el límite de ${MAX_CLASSIFY_ROWS}; divide la tarea.`);
    }
    if (source.rows * source.cols > MAX_TARGET_CELLS) throw new Error("classify_batch: sourceRange demasiado grande.");
    action.source = source;
    action.categories = [
      ...new Set((Array.isArray(args.categories) ? args.categories : []).map((s) => String(s).trim()).filter(Boolean)),
    ];
    action.instruction = String(args.instruction || "").trim();
    action.block = { row: target.row, col: target.col, rows: source.rows, cols: 1 };
  } else {
    const matrix = normalizeMatrix(args.values);
    if (!matrix) throw new Error(`${op}: falta "values" (matriz 2D).`);
    action.values = matrix;
    action.block = { row: target.row, col: target.col, rows: matrix.length, cols: matrix[0].length };
    action.resized =
      (target.rows > 1 || target.cols > 1) && (target.rows !== matrix.length || target.cols !== matrix[0].length);
  }

  const b = action.block;
  if (b.row + b.rows > 1048576 || b.col + b.cols > 16384) throw new Error(`${op}: el destino sale de la hoja.`);
  if (b.rows * b.cols > MAX_TARGET_CELLS) {
    throw new Error(`${op}: ${a1Block(b)} supera ${MAX_TARGET_CELLS} celdas; selecciona solo el área con datos.`);
  }
  return action;
}

/* ---------- classify_batch: clasificación por lotes con salida JSON estructurada ---------- */
async function runPool(tasks, limit, onProgress) {
  const results = new Array(tasks.length);
  let next = 0;
  let failed = false;
  const worker = async () => {
    while (!failed && next < tasks.length) {
      const i = next++;
      try {
        results[i] = await tasks[i]();
      } catch (e) {
        failed = true;
        throw e;
      }
      if (!failed && onProgress) onProgress(i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, tasks.length) }, worker));
  return results;
}

async function classifyBatch(model, key, action, headers, rows) {
  const user = [
    `Criterio: ${action.instruction || "(clasifica según las categorías)"}`,
    action.categories.length
      ? `Categorías permitidas: ${JSON.stringify(action.categories)}`
      : "Categorías: libres (etiqueta corta)",
    headers ? `Cabeceras: ${JSON.stringify(headers.map(compactCell))}` : null,
    `Filas (${rows.length}):`,
    JSON.stringify(rows),
  ]
    .filter(Boolean)
    .join("\n");
  const items = action.categories.length ? { type: "string", enum: action.categories } : { type: "string" };
  const data = await geminiGenerate(model, key, {
    systemInstruction: { parts: [{ text: CLASSIFY_SYSTEM }] },
    contents: [{ role: "user", parts: [{ text: user }] }],
    generationConfig: { responseMimeType: "application/json", responseSchema: { type: "array", items } },
  });
  const arr = extractJSON(parseGeminiResponse(data).text);
  if (!Array.isArray(arr) || arr.length !== rows.length) {
    throw new Error(
      `classify_batch: Gemini devolvió ${Array.isArray(arr) ? arr.length : "un formato inválido de"} etiquetas y se esperaban ${rows.length}.`
    );
  }
  return arr.map((v) => String(v ?? ""));
}

async function runClassification(sheetName, actions, model, key) {
  // Lectura de todos los orígenes en una sola sincronización.
  const reads = await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const items = actions.map((a) => {
      const s = a.source;
      const src = ws.getRangeByIndexes(s.row, s.col, s.rows, s.cols);
      src.load("values");
      let above = null;
      if (!a.header && s.row > 0) {
        above = ws.getRangeByIndexes(s.row - 1, s.col, 1, s.cols);
        above.load("values");
      }
      return { src, above };
    });
    await ctx.sync();
    return items.map((it) => ({ values: it.src.values, above: it.above ? it.above.values[0] : null }));
  });

  for (let k = 0; k < actions.length; k++) {
    const a = actions[k];
    const { values, above } = reads[k];
    const headers = a.header ? values[0] : above && isHeaderLike(above) ? above : null;
    const pending = [];
    for (let r = a.header ? 1 : 0; r < values.length; r++) {
      if (values[r].some((v) => v !== "" && v !== null)) pending.push(r); // filas vacías no se envían
    }
    const batches = [];
    for (let i = 0; i < pending.length; i += CLASSIFY_BATCH) batches.push(pending.slice(i, i + CLASSIFY_BATCH));

    let done = 0;
    setAgentStatus(`Clasificando 0/${pending.length} filas...`);
    const results = await runPool(
      batches.map((idx) => () => classifyBatch(model, key, a, headers, idx.map((r) => values[r].map(compactCell)))),
      CLASSIFY_CONCURRENCY,
      (b) => {
        done += batches[b].length;
        setAgentStatus(`Clasificando ${done}/${pending.length} filas...`);
      }
    );

    const out = values.map(() => [EMPTY_TOKEN]);
    batches.forEach((idx, b) => idx.forEach((r, j) => { out[r] = [results[b][j]]; }));
    a.values = out;
    a.classified = pending.length;
  }
}

/* ---------- motor de escritura atómica (spec §2.2, §2.4, §6) ---------- */
// Agrupa las celdas marcadas en tramos verticales contiguos (o el bloque completo si todo se escribe).
function maskToRuns(block, mask, data, orig, prop) {
  if (mask.every((row) => row.every(Boolean))) {
    return [{ row: block.row, col: block.col, rows: block.rows, cols: block.cols, prop, data, orig }];
  }
  const runs = [];
  for (let c = 0; c < block.cols; c++) {
    let r = 0;
    while (r < block.rows) {
      if (!mask[r][c]) { r++; continue; }
      let end = r;
      while (end + 1 < block.rows && mask[end + 1][c]) end++;
      const span = [];
      for (let k = r; k <= end; k++) span.push(k);
      runs.push({
        row: block.row + r,
        col: block.col + c,
        rows: span.length,
        cols: 1,
        prop,
        data: span.map((k) => [data[k][c]]),
        orig: span.map((k) => [orig[k][c]]),
      });
      r = end + 1;
    }
  }
  return runs;
}

function planWrites(a, read) {
  const { rows, cols } = a.block;
  const F = read.range.formulas;
  const mask = makeMatrix(rows, cols, false);
  const data = makeMatrix(rows, cols, null);
  const stats = { noAnchor: 0, kept: 0, trailing: 0 };
  let prop;

  if (a.op === "fill_blanks_above") {
    const V = read.range.values;
    prop = a.pattern === "relative_above" ? "formulas" : "values";
    // Última fila (dentro del destino) con datos en cualquier columna de esas filas.
    const lastAbs = read.band.isNullObject ? -1 : read.band.rowIndex + read.band.rowCount - 1;
    const last = Math.max(-1, Math.min(rows - 1, lastAbs - a.block.row));
    stats.trailing = rows - 1 - last;
    for (let c = 0; c < cols; c++) {
      const colLetter = toExcelCol(a.block.col + c);
      let hasAnchor = false;
      let anchor = null;
      for (let r = a.header ? 1 : 0; r <= last; r++) {
        if (F[r][c] === "") {
          if (!hasAnchor) { stats.noAnchor++; continue; }
          mask[r][c] = true;
          // Fila superior en base 1 = índice base 0 de la celda actual.
          data[r][c] = prop === "formulas" ? `=${colLetter}${a.block.row + r}` : anchor;
        } else {
          hasAnchor = true;
          anchor = V[r][c];
        }
      }
    }
  } else {
    prop = a.op === "set_formulas" ? "formulas" : "values";
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const v = a.values[r][c];
        if (v === EMPTY_TOKEN || (a.header && r === 0)) continue;
        const cur = F[r][c];
        if (cur !== "" && !a.overwrite) { stats.kept++; continue; }
        if (String(cur) === v) continue;
        mask[r][c] = true;
        data[r][c] = v;
      }
    }
  }

  const cells = [];
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) if (mask[r][c]) cells.push([r, c]);
  return { cells, stats, origTypes: read.range.valueTypes, runs: maskToRuns(a.block, mask, data, F, prop) };
}

function restoreRuns(ws, runs) {
  // En orden inverso para que, si dos acciones se solapan, quede el contenido original.
  for (let i = runs.length - 1; i >= 0; i--) {
    const r = runs[i];
    ws.getRangeByIndexes(r.row, r.col, r.rows, r.cols).formulas = r.orig;
  }
}

async function applyActions(sheetName, actions) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const rangeOf = (b) => ws.getRangeByIndexes(b.row, b.col, b.rows, b.cols);

    // 1) Lectura atómica: todas las propiedades de todos los destinos en una sola sincronización.
    const reads = actions.map((a) => {
      const range = rangeOf(a.block);
      range.load(a.op === "fill_blanks_above" ? ["formulas", "values", "valueTypes"] : ["formulas", "valueTypes"]);
      let band = null;
      if (a.op === "fill_blanks_above") {
        band = range.getEntireRow().getUsedRangeOrNullObject(true);
        band.load(["rowIndex", "rowCount"]);
      }
      return { range, band };
    });
    await ctx.sync();

    // 2) Cálculo en memoria. Cada tramo guarda su contenido original (snapshot).
    const plans = actions.map((a, i) => planWrites(a, reads[i]));
    const runs = plans.flatMap((p) => p.runs);
    if (!runs.length) return { plans, runs, written: 0, newErrors: [], rolledBack: false };

    // 3) Escritura atómica + carga de verificación en la misma sincronización.
    for (const run of runs) ws.getRangeByIndexes(run.row, run.col, run.rows, run.cols)[run.prop] = run.data;
    const checks = actions.map((a) => {
      const r = rangeOf(a.block);
      r.load(["values", "valueTypes"]);
      return r;
    });
    try {
      await ctx.sync();
    } catch (e) {
      restoreRuns(ws, runs);
      try { await ctx.sync(); } catch (_) { /* se informa el error original */ }
      throw new Error(`Excel rechazó la escritura (${e.message}). Se restauró el estado previo.`);
    }

    // 4) Malla de seguridad: errores nuevos (#REF!, #VALUE!, #N/A...) en las celdas escritas.
    const newErrors = [];
    let written = 0;
    plans.forEach((p, i) => {
      const vt = checks[i].valueTypes;
      const vals = checks[i].values;
      const b = actions[i].block;
      p.samples = [];
      for (const [r, c] of p.cells) {
        written++;
        const addr = a1(b.row + r, b.col + c);
        if (vt[r][c] === "Error" && p.origTypes[r][c] !== "Error") newErrors.push(`${addr} ${vals[r][c]}`);
        if (p.samples.length < SAMPLE_SIZE) p.samples.push(`${addr} = ${formatCell(vals[r][c])}`);
      }
    });
    const rolledBack = newErrors.length >= Math.max(1, Math.ceil(written * ERROR_ROLLBACK_RATIO));
    if (rolledBack) {
      restoreRuns(ws, runs);
      await ctx.sync();
    }
    return { plans, runs, written, newErrors, rolledBack };
  });
}

/* ---------- textos de vista previa e informe ---------- */
function previewSamples(a) {
  const out = [];
  for (let r = 0; r < a.values.length && out.length < SAMPLE_SIZE; r++) {
    for (let c = 0; c < a.values[r].length && out.length < SAMPLE_SIZE; c++) {
      const v = a.values[r][c];
      if (v !== EMPTY_TOKEN && !(a.header && r === 0)) out.push(`${a1(a.block.row + r, a.block.col + c)} → ${v}`);
    }
  }
  return out;
}

function describeAction(a, n) {
  const where = a1Block(a.block);
  let line;
  if (a.op === "fill_blanks_above") {
    line =
      `${n}. ${OP_LABEL[a.op]} en ${where} · patrón ` +
      (a.pattern === "relative_above" ? "relative_above (fórmula =celda superior)" : "static_repeat (copia el valor ancla)");
  } else if (a.op === "classify_batch") {
    line =
      `${n}. ${OP_LABEL[a.op]} en ${where} desde ${a1Block(a.source)} · ${a.classified} fila(s)` +
      (a.categories.length ? ` · categorías: ${a.categories.join(", ")}` : "");
  } else {
    line =
      `${n}. ${OP_LABEL[a.op]} en ${where} (${a.block.rows}×${a.block.cols})` +
      (a.resized ? ` · ajustado desde ${a.requested}` : "");
  }
  if (a.header) line += " · cabecera intacta";
  if (a.op !== "fill_blanks_above") line += a.overwrite ? " · SOBRESCRIBE datos existentes" : " · conserva celdas con datos";
  const samples = a.values ? previewSamples(a) : [];
  if (samples.length) line += `\n   Muestra: ${samples.join(" · ")}`;
  return line;
}

function describeApplyResult(sheet, actions, res) {
  if (!res.written) {
    return "No había celdas que modificar: las celdas destino ya tenían datos o no había celdas ancla.";
  }
  if (res.rolledBack) {
    const list = res.newErrors.slice(0, 5).join(", ") + (res.newErrors.length > 5 ? "..." : "");
    return (
      `⚠ Se detectaron ${res.newErrors.length} error(es) nuevos en ${res.written} celda(s) escritas (${list}).\n` +
      "Se restauró automáticamente el estado anterior; no quedó ningún cambio en la hoja."
    );
  }
  const lines = [`✔ Aplicado en ${sheet}: ${res.written} celda(s) modificadas.`];
  actions.forEach((a, i) => {
    const p = res.plans[i];
    let line = `${i + 1}. ${OP_LABEL[a.op]} ${a1Block(a.block)}: ${p.cells.length} celda(s)`;
    if (a.op === "fill_blanks_above") line += ` · ${a.pattern}`;
    if (p.stats.noAnchor) line += ` · ${p.stats.noAnchor} sin ancla superior (quedaron vacías)`;
    if (p.stats.trailing) line += ` · ${p.stats.trailing} fila(s) finales sin datos ignoradas`;
    if (p.stats.kept) line += ` · ${p.stats.kept} con datos conservadas`;
    lines.push(line);
    if (p.samples && p.samples.length) lines.push(`   Validación: ${p.samples.join(" · ")}`);
  });
  if (res.newErrors.length) {
    lines.push(`⚠ ${res.newErrors.length} celda(s) con error (debajo del umbral de reversión): ${res.newErrors.slice(0, 5).join(", ")}`);
  }
  lines.push("Puedes revertirlo con «Deshacer».");
  return lines.join("\n");
}

/* ---------- flujos de UI ---------- */
// Lo llama runAgent() (taskpane.js), que maneja errores, botones y limpia el estado.
async function runGeminiAgent(prompt, useSelection, model, key) {
  setAgentStatus("Analizando rango...");
  await nextFrame();

  let sheet;
  let userMsg;
  let warning = "";
  if (useSelection) {
    const c = await readSelectionContext();
    sheet = c.sheet;
    if (!c.empty) warning = largeSelectionWarning(c.block.rows);
    userMsg = `${formatContext(c)}\n\nPetición: ${prompt}`;
  } else {
    sheet = await getActiveSheetName();
    userMsg = `Hoja: ${sheet}\n\nPetición: ${prompt}`;
  }

  setAgentStatus(`Consultando ${model}...`);
  const { text, calls } = parseGeminiResponse(
    await geminiGenerate(model, key, {
      systemInstruction: { parts: [{ text: GEMINI_AGENT_SYSTEM }] },
      contents: [{ role: "user", parts: [{ text: userMsg }] }],
      tools: [{ functionDeclarations: [EXCEL_TOOL] }],
      toolConfig: { functionCallingConfig: { mode: "AUTO" } },
    })
  );

  const actions = [];
  const rejected = [];
  for (const call of calls) {
    try { actions.push(buildAction(call)); } catch (e) { rejected.push(e.message); }
  }

  const toClassify = actions.filter((a) => a.op === "classify_batch");
  if (toClassify.length) await runClassification(sheet, toClassify, model, key);

  const out = [warning, text || (actions.length ? "" : "Listo.")];
  if (actions.length) {
    out.push(`Acciones propuestas en ${sheet}:\n` + actions.map((a, i) => describeAction(a, i + 1)).join("\n"));
  }
  if (rejected.length) out.push("⚠ Acciones descartadas:\n" + rejected.map((m) => "• " + m).join("\n"));
  showMsg($("agentOut"), out.filter(Boolean).join("\n\n"));

  if (actions.length) setPendingActions(sheet, actions);
}

async function applyGeminiPending() {
  const { sheet, actions } = geminiPending;
  $("agentApply").disabled = true;
  setAgentStatus("Aplicando cambios...");
  await nextFrame();
  try {
    const res = await applyActions(sheet, actions);
    geminiPending = null;
    showMsg($("agentOut"), describeApplyResult(sheet, actions, res), res.rolledBack);
    if (res.written && !res.rolledBack) {
      geminiUndo = { sheet, runs: res.runs };
      $("agentUndo").disabled = false;
    }
  } catch (e) {
    $("agentApply").disabled = false;
    showMsg($("agentOut"), "No se pudo aplicar: " + e.message, true);
  } finally {
    setAgentStatus("");
  }
}

async function undoGemini() {
  if (!geminiUndo) return;
  const { sheet, runs } = geminiUndo;
  $("agentUndo").disabled = true;
  try {
    await Excel.run(async (ctx) => {
      restoreRuns(ctx.workbook.worksheets.getItem(sheet), runs);
      await ctx.sync();
    });
    geminiUndo = null;
    const cells = runs.reduce((n, r) => n + r.rows * r.cols, 0);
    showMsg($("agentOut"), `Cambios deshechos en ${sheet}: ${cells} celda(s) restauradas.`);
  } catch (e) {
    $("agentUndo").disabled = false;
    showMsg($("agentOut"), "No se pudo deshacer: " + e.message, true);
  }
}
