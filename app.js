/* app.js
   Shared helpers used by index.html, entry.html (AI Accounting Assistant),
   and dashboard.html. If your repo already has an app.js with these same
   function names, merge rather than blindly overwrite - this version adds
   nothing beyond what those pages call. */

const API_BASE = window.API_BASE || ""; // same-origin by default; set window.API_BASE before this script loads to point elsewhere

function requireLogin() {
    const name = localStorage.getItem("userName");
    if (!name) {
        window.location.href = "index.html";
        return "";
    }
    return name;
}

function logout() {
    localStorage.removeItem("userName");
    window.location.href = "index.html";
}

async function apiGet(path) {
    const res = await fetch(API_BASE + path);
    if (!res.ok) {
        let detail = res.statusText;
        try { detail = (await res.json()).detail || detail; } catch (e) {}
        throw new Error(detail);
    }
    return res.json();
}

async function apiPost(path, body) {
    const res = await fetch(API_BASE + path, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
    });
    if (!res.ok) {
        let detail = res.statusText;
        try { detail = (await res.json()).detail || detail; } catch (e) {}
        throw new Error(detail);
    }
    return res.json();
}

function fmtMoney(v) {
    v = Number(v || 0);
    const abs = Math.abs(v);
    const sign = v < 0 ? "-" : "";
    if (abs >= 10000000) return sign + "₹" + (abs / 10000000).toFixed(2) + " Cr";
    if (abs >= 100000) return sign + "₹" + (abs / 100000).toFixed(2) + " L";
    return sign + "₹" + abs.toLocaleString("en-IN", { maximumFractionDigits: 0 });
}

function growthPill(val) {
    if (val === null || val === undefined) return `<span class="pill flat">—</span>`;
    const cls = val > 0 ? "up" : val < 0 ? "down" : "flat";
    const arrow = val > 0 ? "▲" : val < 0 ? "▼" : "•";
    return `<span class="pill ${cls}">${arrow} ${Math.abs(val).toFixed(1)}%</span>`;
}

let _toastTimer = null;
function showToast(msg) {
    let el = document.getElementById("_toast");
    if (!el) {
        el = document.createElement("div");
        el.id = "_toast";
        el.style.cssText = "position:fixed;bottom:24px;left:50%;transform:translateX(-50%);" +
            "background:#111827;color:#fff;padding:10px 18px;border-radius:8px;font-size:13px;" +
            "z-index:9999;box-shadow:0 4px 16px rgba(0,0,0,.2);max-width:80vw;text-align:center;";
        document.body.appendChild(el);
    }
    el.textContent = msg;
    el.style.display = "block";
    clearTimeout(_toastTimer);
    _toastTimer = setTimeout(() => { el.style.display = "none"; }, 4000);
}
