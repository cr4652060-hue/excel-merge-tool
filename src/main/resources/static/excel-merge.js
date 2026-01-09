const TEMPLATE_API = "/api/excel/template";
const MERGE_API = "/api/excel/merge";
const EXPORT_API = "/api/excel/export";

const byId = (id) => document.getElementById(id);

function setText(id, text) {
    const el = byId(id);
    if (el) el.textContent = text || "";
}

function renderIssues(issues) {
    const box = byId("issueList");
    if (!box) return;
    if (!issues || issues.length === 0) {
        box.innerHTML = `<div class="hint">未发现问题。</div>`;
        return;
    }
    box.innerHTML = `
        <ul>
            ${issues.map((i) => {
        const rowInfo = i.rowNo ? `行 ${i.rowNo}` : "行 -";
        const colInfo = i.columnName ? `列【${escapeHtml(i.columnName)}】` : "列 -";
        return `<li><b>${escapeHtml(i.fileName || "未知文件")}</b> / ${escapeHtml(i.sheetName || "Sheet1")} / ${rowInfo} / ${colInfo}：${escapeHtml(i.message)}</li>`;
    }).join("")}
        </ul>
    `;
}

function renderPreview(headers, rows) {
    const box = byId("previewTable");
    if (!box) return;
    if (!rows || rows.length === 0) {
        box.innerHTML = `<div class="hint">暂无数据。</div>`;
        return;
    }
    box.innerHTML = `
        <table class="tb">
            <thead>
                <tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>
            </thead>
            <tbody>
                ${rows.map((r) => `
                    <tr>
                        ${r.map((v) => `<td>${escapeHtml(v || "")}</td>`).join("")}
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;
}

function escapeHtml(value) {
    return (value ?? "")
        .toString()
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}

async function uploadTemplate() {
    const input = byId("templateFile");
    if (!input || !input.files || input.files.length === 0) {
        setText("templateInfo", "请先选择模板文件。");
        return;
    }
    const fd = new FormData();
    fd.append("file", input.files[0]);
    setText("templateInfo", "正在识别模板...");
    try {
        const res = await fetch(TEMPLATE_API, {
            method: "POST",
            body: fd
        });
        if (!res.ok) {
            throw new Error(await res.text());
        }
        const data = await res.json();
        setText("templateInfo",
            `已识别 ${data.headers.length} 列，表头行：第 ${data.headerRowIndex} 行，数据起始：第 ${data.dataStartRow} 行。`);
        byId("btnExport").disabled = true;
    } catch (e) {
        setText("templateInfo", `模板识别失败：${e.message}`);
    }
}

async function mergeFiles() {
    const input = byId("dataFiles");
    if (!input || !input.files || input.files.length === 0) {
        setText("mergeStatus", "请先选择支行 Excel 文件。");
        return;
    }
    const fd = new FormData();
    for (const file of input.files) {
        fd.append("files", file);
    }
    setText("mergeStatus", "正在合并，请稍候...");
    try {
        const res = await fetch(MERGE_API, {
            method: "POST",
            body: fd
        });
        if (!res.ok) {
            throw new Error(await res.text());
        }
        const data = await res.json();
        setText("mergeStatus", `合并完成：共 ${data.totalRows} 行。`);
        renderIssues(data.issues);
        renderPreview(data.headers, data.previewRows);
        byId("btnExport").disabled = data.totalRows === 0;
    } catch (e) {
        setText("mergeStatus", `合并失败：${e.message}`);
    }
}

function exportMerged() {
    window.location.href = EXPORT_API;
}

document.addEventListener("DOMContentLoaded", () => {
    const btnTemplate = byId("btnUploadTemplate");
    if (btnTemplate) btnTemplate.addEventListener("click", uploadTemplate);

    const btnMerge = byId("btnMerge");
    if (btnMerge) btnMerge.addEventListener("click", mergeFiles);

    const btnExport = byId("btnExport");
    if (btnExport) btnExport.addEventListener("click", exportMerged);
});