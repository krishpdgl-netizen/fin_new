const API_BASE = ""; // same origin as the frontend (FastAPI serves both)

function requireLogin() {
    const name = localStorage.getItem("userName");
    if (!name) window.location.href = "index.html";
    return name;
}

function logout() {
    localStorage.removeItem("userName");
    window.location.href = "index.html";
}

function showToast(msg) {
    let t = document.getElementById("toast");
    if (!t) {
        t = document.createElement("div");
        t.id = "toast";
        t.className = "toast";
        document.body.appendChild(t);
    }
    t.textContent = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 2800);
}

async function apiGet(path) {
    const res = await fetch(API_BASE + path);
    if (!res.ok) throw new Error(await res.text());
    return res.json();
}

async function apiPost(path, body) {
    const res = await fetch(API_BASE + path, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
}

function fmtMoney(n) {
    if (n === null || n === undefined || isNaN(n)) return "—";
    return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function growthPill(val) {
    if (val === null || val === undefined) return `<span class="growth-pill flat">—</span>`;
    const cls = val > 0 ? "up" : val < 0 ? "down" : "flat";
    const sign = val > 0 ? "▲" : val < 0 ? "▼" : "";
    return `<span class="growth-pill ${cls}">${sign} ${Math.abs(val).toFixed(1)}%</span>`;
}
