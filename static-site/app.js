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

  const counters = new Map(Object.keys(columns).map((label) => [label, new Map()]));
  let total = 0;
  for (const row of rows) {
    if (Number(row.getAttribute("r")) <= 1 || !rowHasData(row)) continue;
    total += 1;
    const values = new Map();
    for (const cell of Array.from(row.getElementsByTagName("c"))) {
      const index = columnIndex(cell.getAttribute("r"));
      if (Object.values(columns).includes(index)) {
        values.set(index, cellText(cell, sharedStrings));
      }
    }
    for (const [label, index] of Object.entries(columns)) {
      addCount(counters.get(label), values.get(index) || "");
    }
  }
  if (!total) throw new Error("La hoja RCV no contiene filas de datos.");
  return Object.entries(columns).map(([label]) => makeDistribution(label, counters.get(label), total));
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

function clearResults() {
  input.value = "";
  document.getElementById("fileName").textContent = "Elige un archivo .xlsx";
  document.getElementById("recordCount").textContent = "0";
  chartsBox.replaceChildren();
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
    const distributions = await readDistributions(file);
    renderCharts(distributions);
    document.getElementById("recordCount").textContent = numberFormat.format(distributions[0].total);
    results.hidden = false;
    clearButton.hidden = false;
    setStatus("Gráficos generados. El libro no se ha enviado ni guardado.", "success");
  } catch (error) {
    setStatus(error instanceof Error ? error.message : "No se pudo leer el libro.", "error");
  }
});

clearButton.addEventListener("click", clearResults);
