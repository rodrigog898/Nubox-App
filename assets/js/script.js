const STORAGE_KEY = "nubox.integration.config.v1";
const NUBOX_BASE_URL = "https://app.nubox.com";
const NUBOX_PATH = "/ServiFactura/paginas/dteDocumentosRecibidos.aspx/ObtenerPorFiltro";
const NUBOX_PDF_PATH = "/ServiFactura/paginas/dteDocumentosRecibidos.aspx/VerPDFDocManualesElectronicos";
const PROXY_ENDPOINT = "http://127.0.0.1:8765/api/nubox/obtener-por-filtro";
const PROXY_PDF_ENDPOINT = "http://127.0.0.1:8765/api/nubox/ver-pdf";
const ZIP_ANALYZE_ENDPOINT = "http://127.0.0.1:8765/api/zip/analizar";

const tabButtons = Array.from(document.querySelectorAll(".tab-button"));
const panels = Array.from(document.querySelectorAll(".panel"));

const cookieInput = document.getElementById("cookieInput");
const fechaDesdeInput = document.getElementById("fechaDesdeInput");
const fechaHastaInput = document.getElementById("fechaHastaInput");

const saveConfigBtn = document.getElementById("saveConfigBtn");
const clearConfigBtn = document.getElementById("clearConfigBtn");
const runRequestBtn = document.getElementById("runRequestBtn");
const downloadZipBtn = document.getElementById("downloadZipBtn");
const clearResultBtn = document.getElementById("clearResultBtn");

const requestState = document.getElementById("requestState");
const requestHttpCode = document.getElementById("requestHttpCode");
const resultOutput = document.getElementById("resultOutput");
const cookieStatus = document.getElementById("cookieStatus");
const resultTableHead = document.getElementById("resultTableHead");
const resultTableBody = document.getElementById("resultTableBody");
const emptyTableMessage = document.getElementById("emptyTableMessage");
const paginationInfo = document.getElementById("paginationInfo");
const prevPageBtn = document.getElementById("prevPageBtn");
const nextPageBtn = document.getElementById("nextPageBtn");
const zipProgressModal = document.getElementById("zipProgressModal");
const zipProgressOverlay = document.getElementById("zipProgressOverlay");
const zipCloseBtn = document.getElementById("zipCloseBtn");
const zipProgressStatus = document.getElementById("zipProgressStatus");
const zipTotalRegistros = document.getElementById("zipTotalRegistros");
const zipDescargados = document.getElementById("zipDescargados");
const zipFallidos = document.getElementById("zipFallidos");
const zipProgressFill = document.getElementById("zipProgressFill");
const zipPercentText = document.getElementById("zipPercentText");
const TRUSTED_NUBOX_ORIGIN = "https://app.nubox.com";
const zipFileInput = document.getElementById("zipFileInput");
const analyzeZipBtn = document.getElementById("analyzeZipBtn");
const downloadAnalyzedExcelBtn = document.getElementById("downloadAnalyzedExcelBtn");
const zipAnalyzeState = document.getElementById("zipAnalyzeState");
const zipAnalysisTableHead = document.getElementById("zipAnalysisTableHead");
const zipAnalysisTableBody = document.getElementById("zipAnalysisTableBody");
const zipAnalysisEmptyMessage = document.getElementById("zipAnalysisEmptyMessage");

const TABLE_PAGE_SIZE = 12;

const defaultConfig = {
    cookieHeader: "",
};

const tableState = {
    rows: [],
    columns: [],
    page: 1,
    ids: [],
};

function safeParseJson(text, fallbackValue) {
    try {
        return JSON.parse(text);
    } catch (_error) {
        return fallbackValue;
    }
}
const zipProgressState = {
    running: false,
};

const zipAnalysisStateData = {
    rows: [],
};

function updateZipButtonState() {
    const hasIds = tableState.idFolioMap && Object.keys(tableState.idFolioMap).length > 0;
    downloadZipBtn.disabled = zipProgressState.running || !hasIds;
}

function openZipProgressModal() {
    zipProgressModal.classList.add("active");
    zipProgressModal.setAttribute("aria-hidden", "false");
}

function closeZipProgressModal() {
    if (zipProgressState.running) {
        return;
    }
    zipProgressModal.classList.remove("active");
    zipProgressModal.setAttribute("aria-hidden", "true");
}

function updateZipProgress({ total, downloaded, failed, status }) {
    const processed = downloaded + failed;
    const safeTotal = Math.max(total, 1);
    const percent = Math.round((processed / safeTotal) * 100);

    zipProgressStatus.textContent = status;
    zipTotalRegistros.textContent = `Registros: ${total}`;
    zipDescargados.textContent = `Descargados: ${downloaded}`;
    zipFallidos.textContent = `Fallidos: ${failed}`;
    zipProgressFill.style.width = `${percent}%`;
    zipPercentText.textContent = `${percent}%`;
}

function tokenFromCookieHeader(cookieHeader) {
    if (!cookieHeader) {
        return "";
    }

    const parts = cookieHeader.split(";").map((part) => part.trim()).filter(Boolean);
    for (const part of parts) {
        const [rawName, ...valueParts] = part.split("=");
        if (!rawName) {
            continue;
        }

        const name = decodeURIComponent(rawName);
        if (name === "COOKIE_SEGURA_SERVIPYME" || rawName === "COOKIE_SEGURA_SERVIPYME") {
            return decodeURIComponent(valueParts.join("="));
        }
    }

    return "";
}

function toIsoDate(dateValue) {
    const year = dateValue.getFullYear();
    const month = String(dateValue.getMonth() + 1).padStart(2, "0");
    const day = String(dateValue.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function slashDateFromIso(isoText) {
    const [year, month, day] = isoText.split("-");
    return `${day}/${month}/${year}`;
}

function isoDateFromSlash(slashText) {
    const [day, month, year] = slashText.split("/");
    if (!day || !month || !year) {
        return "";
    }
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function setDefaultDates() {
    const now = new Date();
    const from = new Date(now.getFullYear(), now.getMonth(), 1);
    const to = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    fechaDesdeInput.value = toIsoDate(from);
    fechaHastaInput.value = toIsoDate(to);
}

function buildRequestBody(token) {
    return {
        filtro: "<Terminos></Terminos>",
        fechaDesde: slashDateFromIso(fechaDesdeInput.value),
        fechaHasta: slashDateFromIso(fechaHastaInput.value),
        folioDesde: 0,
        folioHasta: 0,
        token,
        EstadoId: 4,
    };
}

function findRows(value) {
    if (Array.isArray(value)) {
        return value;
    }

    if (!value || typeof value !== "object") {
        return [];
    }

    if (typeof value.d === "string") {
        const parsed = safeParseJson(value.d, null);
        if (parsed) {
            return findRows(parsed);
        }
    }

    if (value.d && typeof value.d === "object") {
        const nestedRows = findRows(value.d);
        if (nestedRows.length) {
            return nestedRows;
        }
    }

    for (const nestedValue of Object.values(value)) {
        const nestedRows = findRows(nestedValue);
        if (nestedRows.length) {
            return nestedRows;
        }
    }

    return [];
}

function normalizeRow(row) {
    if (row && typeof row === "object" && !Array.isArray(row)) {
        return row;
    }

    return { value: row };
}

function buildColumns(rows) {
    const columnSet = new Set();
    rows.forEach((row) => {
        Object.keys(normalizeRow(row)).forEach((key) => columnSet.add(key));
    });
    return Array.from(columnSet);
}

function renderTable() {
    const totalRows = tableState.rows.length;
    const totalPages = totalRows ? Math.ceil(totalRows / TABLE_PAGE_SIZE) : 0;

    if (tableState.page > totalPages && totalPages > 0) {
        tableState.page = totalPages;
    }
    if (tableState.page < 1) {
        tableState.page = 1;
    }

    const start = (tableState.page - 1) * TABLE_PAGE_SIZE;
    const end = start + TABLE_PAGE_SIZE;
    const pageRows = tableState.rows.slice(start, end).map(normalizeRow);

    resultTableHead.innerHTML = "";
    resultTableBody.innerHTML = "";

    if (!totalRows || !tableState.columns.length) {
        emptyTableMessage.style.display = "block";
        paginationInfo.textContent = "Pagina 0 de 0";
        prevPageBtn.disabled = true;
        nextPageBtn.disabled = true;
        return;
    }

    emptyTableMessage.style.display = "none";

    const headRow = document.createElement("tr");
    tableState.columns.forEach((column) => {
        const th = document.createElement("th");
        th.textContent = column;
        headRow.appendChild(th);
    });
    const actionsTh = document.createElement("th");
    actionsTh.textContent = "PDF";
    headRow.appendChild(actionsTh);
    resultTableHead.appendChild(headRow);

    pageRows.forEach((row) => {
        const tr = document.createElement("tr");
        tableState.columns.forEach((column) => {
            const td = document.createElement("td");
            const rawValue = row[column];
            if (rawValue === null || rawValue === undefined) {
                td.textContent = "";
            } else if (typeof rawValue === "object") {
                td.textContent = JSON.stringify(rawValue);
            } else {
                td.textContent = String(rawValue);
            }
            tr.appendChild(td);
        });

        const actionTd = document.createElement("td");
        const rowId = getRowId(row);
        if (rowId) {
            const actionBtn = document.createElement("button");
            actionBtn.type = "button";
            actionBtn.className = "table-action-btn";
                actionBtn.textContent = "Ver";
            actionBtn.title = `Ver PDF ${rowId}`;
            actionBtn.addEventListener("click", () => {
                fetchAndShowPdf(rowId);
            });
            actionTd.appendChild(actionBtn);
        }
        tr.appendChild(actionTd);

        resultTableBody.appendChild(tr);
    });

    paginationInfo.textContent = `Pagina ${tableState.page} de ${totalPages}`;
    prevPageBtn.disabled = tableState.page <= 1;
    nextPageBtn.disabled = tableState.page >= totalPages;
}

function getRowId(row) {
    const keys = ["id", "Id", "ID", "dteId", "DocumentoId", "docId"];
    for (const key of keys) {
        if (row[key] !== undefined && row[key] !== null && String(row[key]).trim()) {
            return String(row[key]);
        }
    }

    return "";
}

function buildProxyHeaders(cookieHeader) {
    return {
        Cookie: cookieHeader,
        "Content-Type": "application/json;charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "*/*",
        "Origin": "https://app.nubox.com",
        "Referer": "https://app.nubox.com/ServiFactura/paginas/dteDocumentosRecibidos.aspx",
    };
}

function buildTempPdfProxyHeaders(cookieHeader) {
    return {
        Cookie: cookieHeader,
        "Accept": "application/pdf,*/*",
        "Referer": "https://app.nubox.com/ServiFactura/paginas/dteDocumentosRecibidos.aspx",
    };
}

function tryDecodeBase64Pdf(value) {
    if (typeof value !== "string" || !value.length) {
        return null;
    }

    const cleaned = value.replace(/^data:application\/pdf;base64,/, "").trim();
    if (!/^[A-Za-z0-9+/=\s]+$/.test(cleaned) || cleaned.length < 100) {
        return null;
    }

    try {
        const binary = atob(cleaned.replace(/\s/g, ""));
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i += 1) {
            bytes[i] = binary.charCodeAt(i);
        }
        return new Blob([bytes], { type: "application/pdf" });
    } catch (_error) {
        return null;
    }
}

function extractPdfBlobFromJson(value) {
    if (!value) {
        return null;
    }

    if (typeof value === "string") {
        const parsed = safeParseJson(value, null);
        if (parsed) {
            return extractPdfBlobFromJson(parsed);
        }
        return tryDecodeBase64Pdf(value);
    }

    if (Array.isArray(value)) {
        for (const item of value) {
            const blob = extractPdfBlobFromJson(item);
            if (blob) {
                return blob;
            }
        }
        return null;
    }

    if (typeof value === "object") {
        for (const nested of Object.values(value)) {
            const blob = extractPdfBlobFromJson(nested);
            if (blob) {
                return blob;
            }
        }
    }

    return null;
}

function getTempPdfPathFromResponse(value) {
    if (!value || typeof value !== "object") {
        return "";
    }

    if (typeof value.d === "string" && value.d.startsWith("/")) {
        return value.d;
    }

    if (typeof value.d === "string") {
        const parsed = safeParseJson(value.d, null);
        if (parsed) {
            return getTempPdfPathFromResponse(parsed);
        }
    }

    for (const nested of Object.values(value)) {
        if (nested && typeof nested === "object") {
            const found = getTempPdfPathFromResponse(nested);
            if (found) {
                return found;
            }
        }
    }

    return "";
}

async function fetchPdfBlobFromTempPath(tempPath, cookieHeader) {
    const rawPath = String(tempPath || "").trim();
    let normalizedPath = rawPath;
    if (!rawPath.startsWith("//") && rawPath.startsWith("/")) {
        normalizedPath = `/${rawPath}`;
    } else if (!rawPath.startsWith("/")) {
        normalizedPath = `//${rawPath}`;
    }

    const payload = {
        target: {
            baseUrl: NUBOX_BASE_URL,
            path: normalizedPath,
            method: "GET",
        },
        headers: buildTempPdfProxyHeaders(cookieHeader),
        body: {},
    };

    const response = await fetch(PROXY_ENDPOINT, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`No fue posible descargar PDF temporal (${response.status}): ${errorText}`);
    }

    return response.blob();
}

async function fetchPdfBlobForRow(rowId, cookieHeader) {
    const response = await fetch(PROXY_PDF_ENDPOINT, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify({
            id: String(rowId),
            cookieHeader,
        }),
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`No fue posible obtener el PDF (${response.status}): ${errorText}`);
    }

    return response.blob();
}

async function fetchAndShowPdf(rowId) {
    const config = activeConfigFromForm();
    const token = tokenFromCookieHeader(config.cookieHeader);
    if (!token) {
        setResultState("No se encontro COOKIE_SEGURA_SERVIPYME en el header Cookie.", true);
        return;
    }

    setResultState(`Abriendo PDF del registro ${rowId}...`, false);

    try {
        const pdfBlob = await fetchPdfBlobForRow(rowId, config.cookieHeader);
        const blobUrl = URL.createObjectURL(pdfBlob);
        window.open(blobUrl, "_blank", "noopener,noreferrer");
        setResultState("PDF abierto en nueva pestana.", false);
    } catch (error) {
        setResultState(`Error cargando PDF: ${String(error)}`, true);
    }
}

async function downloadAllPdfsAsZip() {
    if (!Array.isArray(tableState.rows) || !tableState.rows.length) {
        setResultState("Primero ejecuta una consulta para obtener registros.", true);
        return;
    }

    if (typeof window.JSZip !== "function") {
        setResultState("No se pudo cargar JSZip para crear el archivo ZIP.", true);
        return;
    }

    const config = activeConfigFromForm();
    const token = tokenFromCookieHeader(config.cookieHeader);
    if (!token) {
        setResultState("No se encontro COOKIE_SEGURA_SERVIPYME en el header Cookie.", true);
        return;
    }

    const ids = Object.keys(tableState.idFolioMap ?? {}).filter(Boolean);

    if (!ids.length) {
        setResultState("No se encontraron IDs para descargar PDFs.", true);
        return;
    }

    zipProgressState.running = true;
    updateZipButtonState();
    openZipProgressModal();
    updateZipProgress({
        total: ids.length,
        downloaded: 0,
        failed: 0,
        status: "Preparando descarga...",
    });

    const zip = new window.JSZip();
    const failedIds = [];
    const failedDetails = [];
    let successCount = 0;

    for (let index = 0; index < ids.length; index += 1) {
        const id = ids[index];
        setResultState(`Descargando PDFs ${index + 1}/${ids.length}...`, false);

        try {
            const pdfBlob = await fetchPdfBlobForRow(id, config.cookieHeader);
            const folio = tableState.idFolioMap[id] || id;
            zip.file(`folio_${folio}.pdf`, pdfBlob);
            successCount += 1;
            updateZipProgress({
                total: ids.length,
                downloaded: successCount,
                failed: failedIds.length,
                status: `Descargando PDFs ${index + 1}/${ids.length}...`,
            });
        } catch (error) {
            failedIds.push(id);
            failedDetails.push(`ID ${id}: ${error instanceof Error ? error.message : String(error)}`);
            updateZipProgress({
                total: ids.length,
                downloaded: successCount,
                failed: failedIds.length,
                status: `Descargando PDFs ${index + 1}/${ids.length}...`,
            });
        }
    }

    if (!successCount) {
        zipProgressState.running = false;
        updateZipButtonState();
        setResultState("No se pudo descargar ningun PDF para el ZIP.", true);
        updateZipProgress({
            total: ids.length,
            downloaded: 0,
            failed: failedIds.length,
            status: "No se pudo descargar ningun PDF.",
        });
        return;
    }

    setResultState("Generando ZIP...", false);
    updateZipProgress({
        total: ids.length,
        downloaded: successCount,
        failed: failedIds.length,
        status: "Generando archivo ZIP...",
    });
    const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    const zipUrl = URL.createObjectURL(zipBlob);
    const link = document.createElement("a");
    link.href = zipUrl;
    link.download = `pdf_nubox_${new Date().toISOString().slice(0, 10)}.zip`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(zipUrl);

    zipProgressState.running = false;
    updateZipButtonState();
    if (failedIds.length) {
        setResultState(`ZIP generado con ${successCount} PDFs. Fallaron ${failedIds.length}.`, true);
        resultOutput.textContent = `IDs con error: ${failedIds.join(", ")}\n\nDetalle:\n${failedDetails.join("\n")}`;
        updateZipProgress({
            total: ids.length,
            downloaded: successCount,
            failed: failedIds.length,
            status: "ZIP generado con errores.",
        });
        return;
    }

    setResultState(`ZIP generado correctamente (${successCount} PDFs).`, false);
    updateZipProgress({
        total: ids.length,
        downloaded: successCount,
        failed: 0,
        status: "ZIP generado correctamente.",
    });
}

function setTableRowsFromResponse(responseData) {
    const rows = findRows(responseData).map(normalizeRow);
    tableState.rows = rows;
    tableState.columns = buildColumns(rows);
    tableState.idFolioMap = Object.fromEntries(
        rows.map((row) => [getRowId(row), row.folio ?? row.Folio ?? row.FOLIO ?? ""])
            .filter(([id]) => id)
    );    tableState.page = 1;
    renderTable();
    updateZipButtonState();
}

function activeConfigFromForm() {
    return {
        cookieHeader: cookieInput.value.trim(),
    };
}

function setFormFromConfig(config) {
    cookieInput.value = config.cookieHeader;
    updateCookieStatus(config.cookieHeader);
}

function loadConfig() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
        return { ...defaultConfig };
    }

    const parsed = safeParseJson(raw, null);
    if (!parsed || typeof parsed !== "object") {
        return { ...defaultConfig };
    }

    return {
        ...defaultConfig,
        ...parsed,
    };
}

function saveConfig(config) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
    updateCookieStatus(config.cookieHeader);
}

function updateCookieStatus(cookieHeaderValue) {
    const cookieCount = cookieHeaderValue
        .split(";")
        .map((part) => part.trim())
        .filter(Boolean).length;

    if (!cookieHeaderValue) {
        cookieStatus.textContent = "Sin configurar";
        return;
    }

    cookieStatus.textContent = `Configuradas ${cookieCount} cookies.`;
}

function setResultState(message, isError) {
    requestState.textContent = message;
    requestState.classList.toggle("is-error", Boolean(isError));
}

function setZipAnalyzeState(message, isError) {
    zipAnalyzeState.textContent = message;
    zipAnalyzeState.classList.toggle("is-error", Boolean(isError));
}

function renderZipAnalysisTable(rows) {
    zipAnalysisTableHead.innerHTML = "";
    zipAnalysisTableBody.innerHTML = "";

    if (!Array.isArray(rows) || !rows.length) {
        zipAnalysisEmptyMessage.style.display = "block";
        downloadAnalyzedExcelBtn.disabled = true;
        return;
    }

    const columns = [
        { key: "rut_emisor", label: "RUT Emisor" },
        { key: "razon_social", label: "Razon Social" },
        { key: "folio", label: "Folio" },
        { key: "fecha", label: "Fecha" },
        { key: "orden_compra", label: "Orden Compra" },
        { key: "fecha_orden_compra", label: "Fecha Orden Compra" },
        { key: "item_numero", label: "Item" },
        { key: "item_descripcion", label: "Descripcion" },
        { key: "precio_unitario", label: "Precio Unitario" },
        { key: "total_item", label: "Total Item" },
    ];

    const headRow = document.createElement("tr");
    columns.forEach((column) => {
        const th = document.createElement("th");
        th.textContent = column.label;
        headRow.appendChild(th);
    });
    zipAnalysisTableHead.appendChild(headRow);

    rows.forEach((row) => {
        const tr = document.createElement("tr");
        columns.forEach((column) => {
            const td = document.createElement("td");
            const value = row?.[column.key];
            td.textContent = value === null || value === undefined ? "" : String(value);
            tr.appendChild(td);
        });
        zipAnalysisTableBody.appendChild(tr);
    });

    zipAnalysisEmptyMessage.style.display = "none";
    downloadAnalyzedExcelBtn.disabled = false;
}

async function analyzeUploadedZip() {
    const file = zipFileInput.files?.[0];
    if (!file) {
        setZipAnalyzeState("Debes seleccionar un archivo ZIP.", true);
        return;
    }

    if (!file.name.toLowerCase().endsWith(".zip")) {
        setZipAnalyzeState("El archivo seleccionado no es ZIP.", true);
        return;
    }

    setZipAnalyzeState("Analizando ZIP en servidor...", false);
    downloadAnalyzedExcelBtn.disabled = true;
    zipAnalysisStateData.rows = [];
    renderZipAnalysisTable([]);

    try {
        const response = await fetch(ZIP_ANALYZE_ENDPOINT, {
            method: "POST",
            headers: {
                "Content-Type": "application/zip",
            },
            body: file,
        });

        const result = await response.json();
        if (!response.ok || !result?.ok) {
            throw new Error(result?.error || `Error ${response.status}`);
        }

        const rows = Array.isArray(result.rows) ? result.rows : [];
        zipAnalysisStateData.rows = rows;
        renderZipAnalysisTable(rows);

        const failedCount = Number(result.failedCount || 0);
        if (failedCount > 0) {
            setZipAnalyzeState(`Analisis completado con ${failedCount} PDF(s) con error.`, true);
            return;
        }

        setZipAnalyzeState(`Analisis completado. Filas detectadas: ${rows.length}.`, false);
    } catch (error) {
        setZipAnalyzeState(`No fue posible analizar el ZIP: ${String(error)}`, true);
        renderZipAnalysisTable([]);
    }
}

function downloadAnalyzedExcel() {
    const rows = zipAnalysisStateData.rows;
    if (!Array.isArray(rows) || !rows.length) {
        setZipAnalyzeState("No hay resultados para exportar.", true);
        return;
    }

    if (typeof window.XLSX !== "object") {
        setZipAnalyzeState("No se pudo cargar la libreria XLSX para exportar.", true);
        return;
    }

    const worksheet = window.XLSX.utils.json_to_sheet(rows);
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "Facturas");
    const stamp = new Date().toISOString().slice(0, 10);
    window.XLSX.writeFile(workbook, `analisis_facturas_${stamp}.xlsx`);
    setZipAnalyzeState("Excel descargado correctamente.", false);
}

function showPanel(panelId) {
    panels.forEach((panel) => {
        panel.classList.toggle("active", panel.id === panelId);
    });

    tabButtons.forEach((button) => {
        const isActive = button.dataset.panel === panelId;
        button.classList.toggle("active", isActive);
        button.setAttribute("aria-selected", String(isActive));
    });
}

async function executeRequest() {
    const config = activeConfigFromForm();

    if (!fechaDesdeInput.value || !fechaHastaInput.value) {
        setResultState("Debes seleccionar fecha desde y fecha hasta.", true);
        requestHttpCode.textContent = "Error";
        return;
    }

    if (fechaDesdeInput.value > fechaHastaInput.value) {
        setResultState("La fecha desde no puede ser mayor que fecha hasta.", true);
        requestHttpCode.textContent = "Error";
        return;
    }

    const detectedToken = tokenFromCookieHeader(config.cookieHeader);
    if (!detectedToken) {
        setResultState("No se encontro COOKIE_SEGURA_SERVIPYME en el header Cookie.", true);
        requestHttpCode.textContent = "Error";
        return;
    }

    const bodyJson = buildRequestBody(detectedToken);

    const payload = {
        target: {
            baseUrl: NUBOX_BASE_URL,
            path: NUBOX_PATH,
            method: "POST",
        },
        headers: buildProxyHeaders(config.cookieHeader),
        body: bodyJson,
    };

    setResultState("Enviando solicitud...", false);
    requestHttpCode.textContent = "...";
    tableState.rows = [];
    tableState.columns = [];
    tableState.idFolioMap = {};
    tableState.page = 1;
    renderTable();
    updateZipButtonState();
    resultOutput.textContent = "Esperando respuesta...";

    try {
        const response = await fetch(PROXY_ENDPOINT, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify(payload),
        });

        requestHttpCode.textContent = String(response.status);

        const text = await response.text();
        const json = safeParseJson(text, null);
        const responsePayload = json ?? text;
        resultOutput.textContent = json ? JSON.stringify(json, null, 2) : text;
        setTableRowsFromResponse(responsePayload);

        if (!response.ok) {
            setResultState("Solicitud respondio con error.", true);
            return;
        }

        setResultState("Solicitud completada.", false);
    } catch (error) {
        requestHttpCode.textContent = "Sin conexion";
        setResultState("No fue posible conectar con el proxy.", true);
        resultOutput.textContent = String(error);
    }
}

function init() {
    const config = loadConfig();
    setFormFromConfig(config);
    setDefaultDates();

    tabButtons.forEach((button) => {
        button.addEventListener("click", () => {
            showPanel(button.dataset.panel);
        });
    });

    saveConfigBtn.addEventListener("click", () => {
        const nextConfig = activeConfigFromForm();
        saveConfig(nextConfig);
        setResultState("Configuracion guardada localmente.", false);
    });

    clearConfigBtn.addEventListener("click", () => {
        setFormFromConfig({ ...defaultConfig });
        saveConfig({ ...defaultConfig });
        setResultState("Configuracion reiniciada.", false);
    });

    runRequestBtn.addEventListener("click", () => {
        saveConfig(activeConfigFromForm());
        executeRequest();
    });

    clearResultBtn.addEventListener("click", () => {
        requestHttpCode.textContent = "-";
        setResultState("Listo para enviar", false);
        tableState.rows = [];
        tableState.columns = [];
        tableState.idFolioMap = {};
        tableState.page = 1;
        renderTable();
        updateZipButtonState();
        resultOutput.textContent = "Aqui veras la respuesta del servicio.";
    });

    prevPageBtn.addEventListener("click", () => {
        if (tableState.page > 1) {
            tableState.page -= 1;
            renderTable();
        }
    });

    nextPageBtn.addEventListener("click", () => {
        const totalPages = Math.ceil(tableState.rows.length / TABLE_PAGE_SIZE);
        if (tableState.page < totalPages) {
            tableState.page += 1;
            renderTable();
        }
    });

    cookieInput.addEventListener("input", () => {
        updateCookieStatus(cookieInput.value.trim());
    });

    downloadZipBtn.addEventListener("click", () => {
        downloadAllPdfsAsZip();
    });

    zipCloseBtn.addEventListener("click", () => {
        closeZipProgressModal();
    });

    zipProgressOverlay.addEventListener("click", () => {
        closeZipProgressModal();
    });

    analyzeZipBtn.addEventListener("click", () => {
        analyzeUploadedZip();
    });

    downloadAnalyzedExcelBtn.addEventListener("click", () => {
        downloadAnalyzedExcel();
    });

    window.addEventListener("message", (event) => {
        if (event.origin !== TRUSTED_NUBOX_ORIGIN) {
            return;
        }

        const message = event.data;
        if (!message || message.source !== "nubox-userscript") {
            return;
        }

        requestHttpCode.textContent = String(message.status ?? "-");
        setResultState(message.ok ? "Datos recibidos desde Nubox." : "Nubox devolvio error.", !message.ok);

        if (message.requestBody?.fechaDesde) {
            const isoFrom = isoDateFromSlash(message.requestBody.fechaDesde);
            if (isoFrom) {
                fechaDesdeInput.value = isoFrom;
            }
        }

        if (message.requestBody?.fechaHasta) {
            const isoTo = isoDateFromSlash(message.requestBody.fechaHasta);
            if (isoTo) {
                fechaHastaInput.value = isoTo;
            }
        }

        resultOutput.textContent = JSON.stringify(message.response, null, 2);
        setTableRowsFromResponse(message.response);
    });

    renderTable();
    updateZipButtonState();
}

init();