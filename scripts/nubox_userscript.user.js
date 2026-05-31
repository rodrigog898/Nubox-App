// ==UserScript==
// @name         Nubox Export To Local Web
// @namespace    nubox.integration.local
// @version      1.0.0
// @description  Ejecuta ObtenerPorFiltro dentro de Nubox y envia la respuesta a tu panel local sin CORS.
// @match        https://app.nubox.com/*
// @grant        none
// ==/UserScript==

(function () {
    "use strict";

    const NUBOX_ORIGIN = "https://app.nubox.com";
    const NUBOX_PATH = "/ServiFactura/paginas/dteDocumentosRecibidos.aspx/ObtenerPorFiltro";
    const TARGET_APP_URL = "http://127.0.0.1:8765/app/index.html";

    function safeParseJson(text) {
        try {
            return JSON.parse(text);
        } catch (_error) {
            return text;
        }
    }

    function pad2(number) {
        return String(number).padStart(2, "0");
    }

    function formatDateSlash(dateValue) {
        const day = pad2(dateValue.getDate());
        const month = pad2(dateValue.getMonth() + 1);
        const year = dateValue.getFullYear();
        return `${day}/${month}/${year}`;
    }

    function getCurrentMonthRange() {
        const now = new Date();
        const start = new Date(now.getFullYear(), now.getMonth(), 1);
        const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        return {
            fechaDesde: formatDateSlash(start),
            fechaHasta: formatDateSlash(end)
        };
    }

    function getCookieValue(cookieName) {
        const cookies = document.cookie.split(";");
        for (const cookiePart of cookies) {
            const [rawName, ...rawValue] = cookiePart.trim().split("=");
            if (!rawName) {
                continue;
            }

            const decodedName = decodeURIComponent(rawName);
            if (rawName === cookieName || decodedName === cookieName) {
                return decodeURIComponent(rawValue.join("="));
            }
        }

        return "";
    }

    function buildDynamicBody() {
        const range = getCurrentMonthRange();
        const fechaDesde = window.prompt("fechaDesde (dd/mm/yyyy)", range.fechaDesde);
        if (!fechaDesde) {
            return null;
        }

        const fechaHasta = window.prompt("fechaHasta (dd/mm/yyyy)", range.fechaHasta);
        if (!fechaHasta) {
            return null;
        }

        const filtro = window.prompt("filtro XML", "<Terminos></Terminos>");
        if (filtro === null) {
            return null;
        }

        const folioDesdeInput = window.prompt("folioDesde", "0");
        if (folioDesdeInput === null) {
            return null;
        }

        const folioHastaInput = window.prompt("folioHasta", "0");
        if (folioHastaInput === null) {
            return null;
        }

        const cookieToken = getCookieValue("COOKIE_SEGURA_SERVIPYME");
        const token = window.prompt("token", cookieToken || "");
        if (!token) {
            window.alert("El token es obligatorio.");
            return null;
        }

        const estadoInput = window.prompt("EstadoId", "4");
        if (estadoInput === null) {
            return null;
        }

        return {
            filtro: filtro || "<Terminos></Terminos>",
            fechaDesde,
            fechaHasta,
            folioDesde: Number.parseInt(folioDesdeInput, 10) || 0,
            folioHasta: Number.parseInt(folioHastaInput, 10) || 0,
            token,
            EstadoId: Number.parseInt(estadoInput, 10) || 4
        };
    }

    function buildDefaultBody() {
        return {
            filtro: "<Terminos></Terminos>",
            ...getCurrentMonthRange(),
            folioDesde: 0,
            folioHasta: 0,
            token: getCookieValue("COOKIE_SEGURA_SERVIPYME") || "",
            EstadoId: 4
        };
    }

    async function runQueryAndSend() {
        const defaultBody = buildDefaultBody();
        let requestBody = buildDynamicBody();
        if (!requestBody) {
            return;
        }

        // Keeps default values if any prompt leaves an empty string.
        requestBody = {
            ...defaultBody,
            ...requestBody
        };

        const endpoint = `${NUBOX_ORIGIN}${NUBOX_PATH}`;

        let response;
        let payload;

        try {
            response = await fetch(endpoint, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json;charset=UTF-8",
                    "X-Requested-With": "XMLHttpRequest"
                },
                credentials: "include",
                body: JSON.stringify(requestBody)
            });

            const text = await response.text();
            payload = safeParseJson(text);
        } catch (error) {
            window.alert(`Error consultando Nubox: ${String(error)}`);
            return;
        }

        const message = {
            source: "nubox-userscript",
            endpoint,
            status: response.status,
            ok: response.ok,
            requestBody,
            response: payload,
            fetchedAt: new Date().toISOString()
        };

        const targetOrigin = new URL(TARGET_APP_URL).origin;
        const targetWindow = window.open(TARGET_APP_URL, "_blank");

        if (!targetWindow) {
            window.alert("No se pudo abrir tu panel local. Revisa bloqueo de popups.");
            return;
        }

        let attempts = 0;
        const maxAttempts = 24;
        const timer = window.setInterval(() => {
            targetWindow.postMessage(message, targetOrigin);
            attempts += 1;

            if (attempts >= maxAttempts) {
                window.clearInterval(timer);
            }
        }, 500);

        window.alert("Consulta ejecutada. Se envio la respuesta a tu panel local.");
    }

    function createButton() {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = "Exportar a mi panel";
        button.style.position = "fixed";
        button.style.right = "18px";
        button.style.bottom = "18px";
        button.style.zIndex = "999999";
        button.style.background = "#0f7da8";
        button.style.color = "#ffffff";
        button.style.border = "none";
        button.style.padding = "10px 14px";
        button.style.borderRadius = "8px";
        button.style.cursor = "pointer";
        button.style.fontFamily = "Arial, sans-serif";
        button.style.fontSize = "13px";

        button.addEventListener("click", () => {
            runQueryAndSend();
        });

        document.body.appendChild(button);
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", createButton);
    } else {
        createButton();
    }
})();
