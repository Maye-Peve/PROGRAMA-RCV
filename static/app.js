const state = {
  meta: null,
  page: 1,
  pageSize: 50,
  search: "",
  filters: {},
  columns: [],
  total: 0,
  pages: 1,
  searchTimer: null,
};

const byId = (id) => document.getElementById(id);
const formatNumber = (value) => new Intl.NumberFormat("es-CO").format(value);

function showToast(message) {
  const toast = byId("toast");
  toast.textContent = message;
  toast.classList.add("visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("visible"), 2600);
}

async function requestJson(url, options) {
  const response = await fetch(url, { cache: "no-store", ...options });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || "No se pudo leer el archivo.");
  return payload;
}

function renderMetadata() {
  const meta = state.meta;
  byId("sourceName").textContent = meta.source_name;
  byId("sourceUpdated").textContent = "Actualizado: " + meta.last_modified;
  byId("recordCount").textContent = formatNumber(meta.record_count);
  byId("fieldCount").textContent = formatNumber(meta.field_count);
  byId("formulaErrorCount").textContent = formatNumber(meta.formula_ref_errors);
  byId("sheetName").textContent = meta.sheet;
  byId("otherSheets").textContent = "Hojas en el libro: " + meta.source_sheets.join(", ");
  byId("formulaMetric").hidden = meta.formula_ref_errors === 0;
  if (meta.formula_ref_errors > 0) {
    const warning = byId("formulaWarning");
    warning.textContent =
      "El libro contiene " +
      formatNumber(meta.formula_ref_errors) +
      " resultados #REF! en campos de TFG. La aplicación conserva esos valores del archivo y no los recalcula.";
    warning.hidden = false;
  }
  renderFilters();
  renderColumnOptions();
  renderBasicCharts();
}

function renderBasicCharts() {
  const container = byId("basicCharts");
  container.replaceChildren();
  for (const distribution of state.meta.basic_distributions || []) {
    const card = document.createElement("article");
    card.className = "chart-card";

    const heading = document.createElement("div");
    heading.className = "chart-card-heading";
    const title = document.createElement("h3");
    title.textContent = distribution.label;
    const total = document.createElement("span");
    total.textContent = formatNumber(distribution.total) + " registros";
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
      value.textContent =
        formatNumber(item.count) +
        " · " +
        new Intl.NumberFormat("es-CO", { maximumFractionDigits: 1 }).format(item.percent) +
        "%";
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
    container.append(card);
  }
}

function renderFilters() {
  const container = byId("filterControls");
  container.replaceChildren();
  for (const filter of state.meta.filters) {
    const label = document.createElement("label");
    label.className = "filter-control";
    const select = document.createElement("select");
    select.setAttribute("aria-label", "Filtrar por " + filter.label);
    select.dataset.column = filter.id;

    const all = document.createElement("option");
    all.value = "__ALL__";
    all.textContent = filter.label + ": Todos";
    select.append(all);
    for (const value of filter.options) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.append(option);
    }
    select.addEventListener("change", () => {
      const value = select.value;
      if (value === "__ALL__") delete state.filters[filter.id];
      else state.filters[filter.id] = value;
      state.page = 1;
      loadRows();
    });
    label.append(select);
    container.append(label);
  }
}

function renderColumnOptions() {
  const container = byId("columnOptions");
  container.replaceChildren();
  const selected = new Set(state.columns);
  for (const column of state.meta.columns) {
    const label = document.createElement("label");
    label.className = "column-option";
    label.dataset.search = column.label.toLocaleLowerCase("es");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = column.id;
    checkbox.checked = selected.has(column.id);
    checkbox.addEventListener("change", () => {
      const next = new Set(state.columns);
      if (checkbox.checked) {
        if (next.size >= 20) {
          checkbox.checked = false;
          showToast("Puedes mostrar hasta 20 campos a la vez.");
          return;
        }
        next.add(column.id);
      } else {
        next.delete(column.id);
      }
      state.columns = Array.from(next);
      updateVisibleCount();
      state.page = 1;
      loadRows();
    });
    const text = document.createElement("span");
    text.textContent = column.label;
    label.append(checkbox, text);
    container.append(label);
  }
  updateVisibleCount();
}

function updateVisibleCount() {
  byId("visibleCount").textContent = state.columns.length + "/" + state.meta.field_count;
}

function formatCell(value) {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "number") return new Intl.NumberFormat("es-CO", { maximumFractionDigits: 3 }).format(value);
  return String(value);
}

function renderRows(result) {
  state.total = result.total;
  state.page = result.page;
  state.pages = Math.max(1, Math.ceil(result.total / result.page_size));

  const head = byId("tableHead");
  head.replaceChildren();
  const headerRow = document.createElement("tr");
  for (const id of state.columns) {
    const column = state.meta.columns.find((item) => item.id === id);
    const th = document.createElement("th");
    th.textContent = column ? column.label : id;
    th.title = column ? column.header : id;
    headerRow.append(th);
  }
  const actionHeader = document.createElement("th");
  actionHeader.textContent = "Detalle";
  headerRow.append(actionHeader);
  head.append(headerRow);

  const body = byId("tableBody");
  body.replaceChildren();
  if (!result.rows.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = Math.max(1, state.columns.length + 1);
    cell.className = "empty-cell";
    cell.textContent = "No hay registros que coincidan con la búsqueda.";
    row.append(cell);
    body.append(row);
  }

  for (const record of result.rows) {
    const row = document.createElement("tr");
    for (const columnId of state.columns) {
      const cell = document.createElement("td");
      const value = record.values[columnId];
      cell.textContent = formatCell(value);
      cell.title = formatCell(value);
      if (typeof value === "string" && value.startsWith("#")) cell.classList.add("error-value");
      row.append(cell);
    }
    const actionCell = document.createElement("td");
    const detailButton = document.createElement("button");
    detailButton.type = "button";
    detailButton.className = "row-action";
    detailButton.textContent = "Ver fila";
    detailButton.addEventListener("click", () => openRecord(record.source_row));
    actionCell.append(detailButton);
    row.append(actionCell);
    body.append(row);
  }

  const first = result.total === 0 ? 0 : (result.page - 1) * result.page_size + 1;
  const last = Math.min(result.page * result.page_size, result.total);
  byId("resultSummary").textContent =
    "Filas " + formatNumber(first) + "–" + formatNumber(last) + " de " + formatNumber(result.total);
  byId("pageLabel").textContent = "Página " + result.page + " de " + state.pages;
  byId("previousPage").disabled = result.page <= 1;
  byId("nextPage").disabled = result.page >= state.pages;
}

async function loadRows() {
  const body = byId("tableBody");
  body.replaceChildren();
  const loading = document.createElement("tr");
  const cell = document.createElement("td");
  cell.colSpan = Math.max(1, state.columns.length + 1);
  cell.className = "loading-cell";
  cell.textContent = "Consultando registros locales…";
  loading.append(cell);
  body.append(loading);
  try {
    const result = await requestJson("/api/query", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        page: state.page,
        page_size: state.pageSize,
        search: state.search,
        filters: state.filters,
        columns: state.columns,
      }),
    });
    renderRows(result);
  } catch (error) {
    body.replaceChildren();
    const row = document.createElement("tr");
    const errorCell = document.createElement("td");
    errorCell.colSpan = Math.max(1, state.columns.length + 1);
    errorCell.className = "error-cell";
    errorCell.textContent = error.message;
    row.append(errorCell);
    body.append(row);
    byId("resultSummary").textContent = "No se pudo consultar el libro.";
  }
}

async function openRecord(sourceRow) {
  try {
    const record = await requestJson("/api/record", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ source_row: sourceRow }),
    });
    byId("dialogTitle").textContent = "Fila " + record.source_row + " de RCV";
    const details = byId("recordDetails");
    details.replaceChildren();
    for (const column of state.meta.columns) {
      const value = record.values[column.id];
      const block = document.createElement("div");
      const term = document.createElement("dt");
      term.textContent = column.label;
      const description = document.createElement("dd");
      description.textContent = formatCell(value);
      if (typeof value === "string" && value.startsWith("#")) description.className = "error-value";
      block.append(term, description);
      details.append(block);
    }
    byId("recordDialog").showModal();
  } catch (error) {
    showToast(error.message);
  }
}

async function start() {
  try {
    state.meta = await requestJson("/api/meta");
    state.columns = state.meta.default_columns.slice();
    renderMetadata();
    await loadRows();
  } catch (error) {
    byId("sourceName").textContent = "No disponible";
    byId("tableBody").replaceChildren();
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "error-cell";
    cell.textContent = error.message;
    row.append(cell);
    byId("tableBody").append(row);
    byId("resultSummary").textContent = "No se pudo abrir el libro.";
  }
}

byId("searchInput").addEventListener("input", (event) => {
  window.clearTimeout(state.searchTimer);
  state.searchTimer = window.setTimeout(() => {
    state.search = event.target.value.trim();
    state.page = 1;
    loadRows();
  }, 240);
});
byId("clearSearch").addEventListener("click", () => {
  byId("searchInput").value = "";
  state.search = "";
  state.page = 1;
  loadRows();
});
byId("previousPage").addEventListener("click", () => {
  state.page = Math.max(1, state.page - 1);
  loadRows();
});
byId("nextPage").addEventListener("click", () => {
  state.page = Math.min(state.pages, state.page + 1);
  loadRows();
});
byId("resetColumns").addEventListener("click", () => {
  state.columns = state.meta.default_columns.slice();
  renderColumnOptions();
  state.page = 1;
  loadRows();
});
byId("columnSearch").addEventListener("input", (event) => {
  const query = event.target.value.toLocaleLowerCase("es").trim();
  for (const item of byId("columnOptions").children) {
    item.hidden = query !== "" && !item.dataset.search.includes(query);
  }
});
byId("closeDialog").addEventListener("click", () => byId("recordDialog").close());
byId("recordDialog").addEventListener("click", (event) => {
  if (event.target === byId("recordDialog")) byId("recordDialog").close();
});

start();
