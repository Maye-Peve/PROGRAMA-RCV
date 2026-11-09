const input = document.getElementById("workbookInput");
const statusBox = document.getElementById("status");
const chartsBox = document.getElementById("basicCharts");
const results = document.getElementById("results");
const clearButton = document.getElementById("clearButton");

const numberFormat = new Intl.NumberFormat("es-CO");
const percentFormat = new Intl.NumberFormat("es-CO", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 2,
});
const decoder = new TextDecoder("utf-8");
const LABORATORY_FIELDS = [
  { hint: "var27.creatinina", label: "Creatinina", missingCodes: ["99"] },
  { hint: "hemoglobina.glicosilada", label: "Hemoglobina glicosilada (HbA1c)", missingCodes: ["98", "99"] },
  { hint: "albuminuria-24horas", label: "Albuminuria 24 horas", missingCodes: ["9999"] },
  { hint: "creatinuria-rac", label: "Creatinuria / RAC", missingCodes: ["9999"] },
  { hint: "colesterol.total", label: "Colesterol total", missingCodes: ["999"] },
  { hint: "var32.hdl", label: "HDL", missingCodes: ["999"] },
  { hint: "triglicerios", label: "Triglicéridos", missingCodes: ["999"] },
  { hint: "var33.ldl", label: "LDL", missingCodes: ["999"] },
  { hint: "no-hdl", label: "No-HDL", missingCodes: ["999", "9999"] },
  { hint: "var34.pth", label: "PTH", missingCodes: ["9988"] },
  { hint: "tasa.de.filtracion.glomerularccghc", label: "TFG Cockcroft-Gault", missingCodes: ["999"] },
  { hint: "tasa.de.filtracion.glomerularckdepi", label: "TFG CKD-EPI", missingCodes: ["999"] },
  { hint: "depuraciondecreatinina", label: "Depuración de creatinina", missingCodes: ["99"] },
  { hint: "var59.hemoglobina", label: "Hemoglobina", missingCodes: ["98", "99"] },
  { hint: "var60.albumina.serica", label: "Albúmina sérica", missingCodes: ["98", "99"] },
  { hint: "var61.fosforo", label: "Fósforo", missingCodes: ["98", "99"] },
  { hint: "glucosa", label: "Glucosa", missingCodes: ["99"], dateHint: "fechaglucosa", validateByDate: true },
  { hint: "uroanalisis", label: "Uroanálisis", missingCodes: ["99"] },
];
const MISSING_DATE_VALUES = new Set(["1800-01-01", "1845-01-01"]);
const MISSING_ERROR_VALUES = new Set(["#ref!", "#value!", "#div/0!", "#n/a", "#name?", "#num!", "#null!"]);
const LABORATORY_PERIOD_START = "2026-07-01";
const LABORATORY_PERIOD_FIELDS = [
  { hint: "var27_1.fecha.de.ultima.creatinina", label: "Creatinina" },
  { hint: "var30_1.fecha.de.la.ultima.creatinuria-rac", label: "RAC" },
  { hint: "var28_1.fecha.de.ultima.hemoglobina.glicosilada", label: "Hemoglobina glicosilada" },
  { hint: "var31_1.fecha.de.la.ultimo.colesterol.total", label: "Colesterol total" },
  { hint: "var32_1.fecha.de.la.ultimo.hdl", label: "HDL" },
  { hint: "var33_1.fecha.del.ultimo.ldl", label: "LDL" },
];

function setStatus(message, kind = "") {
  statusBox.textContent = message;
  statusBox.className = "status" + (kind ? " " + kind : "");
}

function findEndOfCentralDirectory(view) {
  const minOffset = Math.max(0, view.byteLength - 65_557);
  for (let offset = view.byteLength - 22; offset >= minOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error("El archivo no parece ser un libro Excel .xlsx válido.");
}

function zipEntries(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const endOffset = findEndOfCentralDirectory(view);
  const entryCount = view.getUint16(endOffset + 10, true);
  let directoryOffset = view.getUint32(endOffset + 16, true);
  const entries = new Map();

  for (let index = 0; index < entryCount; index += 1) {
    if (view.getUint32(directoryOffset, true) !== 0x02014b50) {
      throw new Error("No se pudo leer el índice interno del archivo Excel.");
    }
    const method = view.getUint16(directoryOffset + 10, true);
    const compressedSize = view.getUint32(directoryOffset + 20, true);
    const nameLength = view.getUint16(directoryOffset + 28, true);
    const extraLength = view.getUint16(directoryOffset + 30, true);
    const commentLength = view.getUint16(directoryOffset + 32, true);
    const localHeaderOffset = view.getUint32(directoryOffset + 42, true);
    const nameStart = directoryOffset + 46;
    const name = decoder.decode(bytes.subarray(nameStart, nameStart + nameLength));

    if (view.getUint32(localHeaderOffset, true) !== 0x04034b50) {
      throw new Error("El archivo Excel contiene una entrada comprimida inválida.");
    }
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);
    entries.set(name, { method, compressed });

    directoryOffset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

async function unzipEntry(entries, name) {
  const entry = entries.get(name);
  if (!entry) return null;
  if (entry.method === 0) return entry.compressed;
  if (entry.method !== 8) throw new Error("El libro usa un método de compresión no compatible.");
  if (typeof DecompressionStream === "undefined") {
    throw new Error("Este navegador no permite leer .xlsx. Actualiza Chrome o Edge e inténtalo de nuevo.");
  }
  try {
    const stream = new Blob([entry.compressed])
      .stream()
      .pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch (_error) {
    throw new Error("No se pudo descomprimir el libro. Verifica que el archivo no esté dañado.");
  }
}

async function unzipText(entries, name, required = true) {
  const bytes = await unzipEntry(entries, name);
  if (!bytes) {
    if (required) throw new Error("Falta un componente necesario del libro Excel: " + name);
    return null;
  }
  return decoder.decode(bytes);
}

function parseXml(text, name) {
  const document = new DOMParser().parseFromString(text, "application/xml");
  if (document.getElementsByTagName("parsererror").length) {
    throw new Error("No se pudo interpretar la estructura Excel: " + name);
  }
  return document;
}

function zipPath(base, target) {
  const raw = target.startsWith("/") ? target.slice(1) : base + "/" + target;
  const parts = [];
  for (const part of raw.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

function columnIndex(cellReference) {
  const letters = (cellReference.match(/^[A-Z]+/i) || [""])[0].toUpperCase();
  let result = 0;
  for (const letter of letters) result = result * 26 + letter.charCodeAt(0) - 64;
  return result;
}

function cellText(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") {
    return Array.from(cell.getElementsByTagName("t"))
      .map((item) => item.textContent || "")
      .join("");
  }
  const valueNode = cell.getElementsByTagName("v")[0];
  if (!valueNode) return "";
  const value = valueNode.textContent || "";
  if (type === "s") return sharedStrings[Number(value)] ?? "";
  return value;
}

function rowHasData(row) {
  return Array.from(row.getElementsByTagName("c")).some((cell) => {
    if (cell.getElementsByTagName("f").length || cell.getElementsByTagName("is").length) return true;
    const value = cell.getElementsByTagName("v")[0];
    return Boolean(value && value.textContent !== "");
  });
}

function normalizeCategory(value) {
  return value.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase("es-CO");
}

function normalizeMissingValue(value) {
  let text = String(value ?? "").trim();
  if (!text) return "";
  if (/^\d{4}-\d{2}-\d{2}t/i.test(text)) text = text.slice(0, 10);
  const numeric = Number(text.replace(",", "."));
  if (Number.isFinite(numeric) && Number.isInteger(numeric)) return String(numeric);
  return normalizeCategory(text);
}

function isMissingLabValue(value, missingCodes) {
  const token = normalizeMissingValue(value);
  return !token || missingCodes.includes(token) || MISSING_DATE_VALUES.has(token) || MISSING_ERROR_VALUES.has(token);
}

function currentDateISO() {
  const today = new Date();
  return [today.getFullYear(), String(today.getMonth() + 1).padStart(2, "0"), String(today.getDate()).padStart(2, "0")].join("-");
}

function excelDateISO(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const serial = Number(text.replace(",", "."));
  if (!Number.isFinite(serial) || serial <= 0) return null;
  const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86_400_000);
  if (Number.isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 10);
}

function displayISODate(value) {
  return new Intl.DateTimeFormat("es-CO", {
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  }).format(new Date(value + "T12:00:00Z"));
}

function addCount(counter, rawValue) {
  const value = rawValue.trim();
  const key = value ? "value:" + normalizeCategory(value) : "missing:";
  const previous = counter.get(key);
  counter.set(key, {
    label: value ? value.toLocaleUpperCase("es-CO") : "Sin dato",
    count: (previous?.count || 0) + 1,
  });
}

function makeDistribution(label, counter, total) {
  const items = Array.from(counter.values())
    .sort((left, right) => right.count - left.count || left.label.localeCompare(right.label, "es"))
    .map((item) => ({
      ...item,
      percent: total ? (item.count / total) * 100 : 0,
    }));
  return { label, total, items };
}

async function readDistributions(file) {
  const entries = zipEntries(await file.arrayBuffer());
  const workbookXml = parseXml(await unzipText(entries, "xl/workbook.xml"), "xl/workbook.xml");
  const relationshipsXml = parseXml(
    await unzipText(entries, "xl/_rels/workbook.xml.rels"),
    "xl/_rels/workbook.xml.rels",
  );
  const relationships = new Map(
    Array.from(relationshipsXml.getElementsByTagName("Relationship"))
      .map((item) => [item.getAttribute("Id"), item.getAttribute("Target")]),
  );
  const sheet = Array.from(workbookXml.getElementsByTagName("sheet"))
    .find((item) => item.getAttribute("name") === "RCV");
  if (!sheet) throw new Error("El libro no contiene una hoja llamada RCV.");

  const relationId = sheet.getAttribute("r:id") || sheet.getAttributeNS(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "id",
  );
  const sheetTarget = relationships.get(relationId);
  if (!sheetTarget) throw new Error("No se pudo localizar la hoja RCV dentro del libro.");
  const sheetPath = zipPath("xl", sheetTarget);

  const sharedText = await unzipText(entries, "xl/sharedStrings.xml", false);
  const sharedStrings = sharedText
    ? Array.from(parseXml(sharedText, "xl/sharedStrings.xml").getElementsByTagName("si"))
        .map((item) => Array.from(item.getElementsByTagName("t"))
          .map((part) => part.textContent || "")
          .join(""))
    : [];
  const worksheet = parseXml(await unzipText(entries, sheetPath), sheetPath);
  const rows = Array.from(worksheet.getElementsByTagName("row"));
  const headerRow = rows.find((row) => Number(row.getAttribute("r")) === 1);
  if (!headerRow) throw new Error("La hoja RCV no tiene encabezados en la primera fila.");

  const headers = new Map();
  for (const cell of Array.from(headerRow.getElementsByTagName("c"))) {
    headers.set(columnIndex(cell.getAttribute("r")), cellText(cell, sharedStrings));
  }
  const findHeader = (needle) => {
    for (const [index, header] of headers) {
      if (normalizeCategory(header).includes(needle)) return index;
    }
    return null;
  };
  const columns = {
    "Sexo": findHeader("sexo") || 3,
    "Régimen": findHeader("regimen") || 4,
    "Zonificación": findHeader("zonificacion") || 5,
  };

  const labColumns = LABORATORY_FIELDS.flatMap((field) => {
    const match = Array.from(headers).find(([, header]) => {
      const normalized = normalizeCategory(header);
      return normalized.includes(field.hint) && !normalized.includes("fecha");
    });
    if (!match) return [];
    const dateColumn = field.dateHint ? findHeader(field.dateHint) : null;
    return [{ ...field, column: match[0], dateColumn }];
  });
  const labPeriodColumns = LABORATORY_PERIOD_FIELDS.flatMap((field) => {
    const column = findHeader(field.hint);
    return column ? [{ ...field, column }] : [];
  });

  const counters = new Map(Object.keys(columns).map((label) => [label, new Map()]));
  const missingCounts = new Map(labColumns.map((field) => [field.label, 0]));
  const takenCounts = new Map(labPeriodColumns.map((field) => [field.label, 0]));
  const relevantColumns = new Set([
    ...Object.values(columns),
    ...labColumns.flatMap((field) => [field.column, field.dateColumn].filter(Boolean)),
    ...labPeriodColumns.map((field) => field.column),
  ]);
  const periodEnd = currentDateISO();
  let total = 0;
  for (const row of rows) {
    if (Number(row.getAttribute("r")) <= 1 || !rowHasData(row)) continue;
    total += 1;
    const values = new Map();
    for (const cell of Array.from(row.getElementsByTagName("c"))) {
      const index = columnIndex(cell.getAttribute("r"));
      if (relevantColumns.has(index)) {
        values.set(index, cellText(cell, sharedStrings));
      }
    }
    for (const [label, index] of Object.entries(columns)) {
      addCount(counters.get(label), values.get(index) || "");
    }
    for (const field of labColumns) {
      const value = values.get(field.column) || "";
      let missing = isMissingLabValue(value, field.missingCodes);
      if (
        field.validateByDate &&
        normalizeMissingValue(value) === "99" &&
        !isMissingLabValue(values.get(field.dateColumn), [])
      ) {
        missing = false;
      }
      if (missing) missingCounts.set(field.label, missingCounts.get(field.label) + 1);
    }
    for (const field of labPeriodColumns) {
      const date = excelDateISO(values.get(field.column));
      if (date && date >= LABORATORY_PERIOD_START && date <= periodEnd) {
        takenCounts.set(field.label, takenCounts.get(field.label) + 1);
      }
    }
  }
  if (!total) throw new Error("La hoja RCV no contiene filas de datos.");
  const distributions = Object.entries(columns).map(([label]) => makeDistribution(label, counters.get(label), total));
  const laboratoryMissingness = labColumns
    .map((field) => {
      const missing = missingCounts.get(field.label);
      return {
        label: field.label,
        total,
        missing,
        available: total - missing,
        percent: total ? (missing / total) * 100 : 0,
      };
    })
    .sort((left, right) => right.missing - left.missing || left.label.localeCompare(right.label, "es"));
  const laboratoryPeriodStatus = labPeriodColumns.map((field) => {
    const taken = takenCounts.get(field.label);
    const pending = total - taken;
    return {
      label: field.label,
      total,
      taken,
      pending,
      takenPercent: total ? (taken / total) * 100 : 0,
      pendingPercent: total ? (pending / total) * 100 : 0,
    };
  });
  return {
    distributions,
    laboratoryMissingness,
    laboratoryPeriodStatus,
    laboratoryPeriodStart: LABORATORY_PERIOD_START,
    laboratoryPeriodEnd: periodEnd,
    total,
  };
}

function renderCharts(distributions) {
  chartsBox.replaceChildren();
  for (const distribution of distributions) {
    const card = document.createElement("article");
    card.className = "chart-card";

    const heading = document.createElement("div");
    heading.className = "chart-card-heading";
    const title = document.createElement("h3");
    title.textContent = distribution.label;
    const total = document.createElement("span");
    total.textContent = numberFormat.format(distribution.total) + " registros";
    heading.append(title, total);

    const list = document.createElement("ul");
    list.className = "chart-list";
    const maxCount = Math.max(1, ...distribution.items.map((item) => item.count));
    for (const item of distribution.items) {
      const row = document.createElement("li");
      row.className = "chart-row";
      const labels = document.createElement("div");
      labels.className = "chart-row-labels";
      const category = document.createElement("span");
      category.className = "chart-category";
      category.textContent = item.label;
      const value = document.createElement("span");
      value.className = "chart-value";
      value.textContent = numberFormat.format(item.count) + " · " + percentFormat.format(item.percent) + "%";
      labels.append(category, value);

      const track = document.createElement("div");
      track.className = "chart-track";
      track.setAttribute("aria-hidden", "true");
      const bar = document.createElement("div");
      bar.className = "chart-bar";
      bar.style.width = ((item.count / maxCount) * 100).toFixed(1) + "%";
      track.append(bar);
      row.append(labels, track);
      list.append(row);
    }
    card.append(heading, list);
    chartsBox.append(card);
  }
}

function renderLaboratoryChart(items) {
  const list = document.getElementById("labCharts");
  list.replaceChildren();
  for (const item of items) {
    const row = document.createElement("li");
    row.className = "chart-row";

    const labels = document.createElement("div");
    labels.className = "chart-row-labels";
    const category = document.createElement("span");
    category.className = "chart-category";
    category.textContent = item.label;
    const value = document.createElement("span");
    value.className = "chart-value";
    value.textContent = numberFormat.format(item.missing) + " · " + percentFormat.format(item.percent) + "% sin resultado";
    labels.append(category, value);

    const track = document.createElement("div");
    track.className = "chart-track";
    track.setAttribute("aria-hidden", "true");
    const bar = document.createElement("div");
    bar.className = "chart-bar chart-bar-missing";
    bar.style.width = Math.min(100, Math.max(0, item.percent)).toFixed(1) + "%";
    track.append(bar);
    row.append(labels, track);
    list.append(row);
  }
}

function renderLaboratoryPeriodChart(data) {
  const list = document.getElementById("labPeriodCharts");
  list.replaceChildren();
  document.getElementById("labPeriodRange").textContent =
    "Del " + displayISODate(data.laboratoryPeriodStart) + " al " + displayISODate(data.laboratoryPeriodEnd) + ".";

  for (const item of data.laboratoryPeriodStatus) {
    const row = document.createElement("li");
    row.className = "lab-period-row";

    const labels = document.createElement("div");
    labels.className = "lab-period-labels";
    const name = document.createElement("span");
    name.className = "lab-period-name";
    name.textContent = item.label;
    const counts = document.createElement("div");
    counts.className = "lab-period-counts";
    const taken = document.createElement("span");
    taken.className = "lab-taken-count";
    taken.textContent = "Tomados " + numberFormat.format(item.taken) + " · " + percentFormat.format(item.takenPercent) + "%";
    const pending = document.createElement("span");
    pending.className = "lab-pending-count";
    pending.textContent = "Pendientes " + numberFormat.format(item.pending) + " · " + percentFormat.format(item.pendingPercent) + "%";
    counts.append(taken, pending);
    labels.append(name, counts);

    const track = document.createElement("div");
    track.className = "lab-period-track";
    track.setAttribute("aria-label", item.label + ": " + numberFormat.format(item.taken) + " tomados y " + numberFormat.format(item.pending) + " pendientes");
    const takenBar = document.createElement("div");
    takenBar.className = "lab-period-taken";
    takenBar.style.width = item.takenPercent.toFixed(2) + "%";
    const pendingBar = document.createElement("div");
    pendingBar.className = "lab-period-pending";
    pendingBar.style.width = item.pendingPercent.toFixed(2) + "%";
    track.append(takenBar, pendingBar);
    row.append(labels, track);
    list.append(row);
  }
}

function clearResults() {
  input.value = "";
  document.getElementById("fileName").textContent = "Elige un archivo .xlsx";
  document.getElementById("recordCount").textContent = "0";
  chartsBox.replaceChildren();
  document.getElementById("labCharts").replaceChildren();
  document.getElementById("labPeriodCharts").replaceChildren();
  results.hidden = true;
  clearButton.hidden = true;
  setStatus("Elige el libro desde tu equipo para ver los gráficos.");
}

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;
  document.getElementById("fileName").textContent = file.name;
  clearButton.hidden = true;
  results.hidden = true;
  chartsBox.replaceChildren();
  setStatus("Leyendo la hoja RCV en este navegador…");
  try {
    const data = await readDistributions(file);
    renderCharts(data.distributions);
    renderLaboratoryChart(data.laboratoryMissingness);
    renderLaboratoryPeriodChart(data);
    document.getElementById("recordCount").textContent = numberFormat.format(data.total);
    results.hidden = false;
    clearButton.hidden = false;
    setStatus("Gráficos generados. El libro no se ha enviado ni guardado.", "success");
  } catch (error) {
    setStatus(error instanceof Error ? error.message : "No se pudo leer el libro.", "error");
  }
});

clearButton.addEventListener("click", clearResults);
