const state = {
  uploadId: "",
  dates: [],
  uploadInfo: null,
};

const el = (id) => document.getElementById(id);

const uploadForm = el("uploadForm");
const generateForm = el("generateForm");
const dateSelect = el("dateSelect");
const previewBtn = el("previewBtn");
const generateBtn = el("generateBtn");
const downloadLink = el("downloadLink");
const warningBox = el("warningBox");
const previewTable = el("previewTable");
const logOutput = el("logOutput");

function log(message) {
  const time = new Date().toLocaleTimeString("zh-CN", { hour12: false });
  logOutput.textContent += `[${time}] ${message}\n`;
  logOutput.scrollTop = logOutput.scrollHeight;
}

function setBusy(button, busy, text) {
  if (!button) return;
  if (busy) {
    button.dataset.text = button.textContent;
    button.textContent = text;
    button.disabled = true;
  } else {
    button.textContent = button.dataset.text || button.textContent;
    button.disabled = false;
  }
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let payload = {};
  if (text) {
    try {
      payload = JSON.parse(text);
    } catch {
      payload = { detail: text };
    }
  }
  if (!response.ok) {
    throw new Error(payload.detail || `请求失败：${response.status}`);
  }
  return payload;
}

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  downloadLink.classList.add("hidden");
  warningBox.classList.add("hidden");
  setBusy(el("uploadBtn"), true, "解析中...");
  try {
    const formData = new FormData(uploadForm);
    const payload = await requestJson("/api/upload", { method: "POST", body: formData });
    state.uploadId = payload.upload_id;
    state.dates = payload.dates || [];
    state.uploadInfo = payload;
    fillDates(state.dates);
    el("sessionText").textContent = `upload_id：${payload.upload_id}`;
    el("recordCount").textContent = payload.record_count || 0;
    showWarnings(payload.warnings || []);
    previewBtn.disabled = state.dates.length === 0;
    generateBtn.disabled = state.dates.length === 0;
    log(`台账解析完成：${payload.record_count} 条记录，${payload.image_count} 张问题照片`);
    if (state.dates.length) {
      await refreshPreview();
    }
  } catch (error) {
    log(error.message);
    alert(error.message);
  } finally {
    setBusy(el("uploadBtn"), false);
  }
});

previewBtn.addEventListener("click", refreshPreview);
dateSelect.addEventListener("change", refreshPreview);

generateForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!state.uploadId) return;
  const types = selectedTypes();
  if (!types.length) {
    alert("请至少选择一种日报类型");
    return;
  }
  setBusy(generateBtn, true, "生成中...");
  downloadLink.classList.add("hidden");
  try {
    const payload = await requestJson("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        upload_id: state.uploadId,
        date: dateSelect.value,
        types,
        zip: el("zipToggle").checked,
      }),
    });
    downloadLink.href = payload.download_url;
    downloadLink.textContent = `下载 ${payload.filename}`;
    downloadLink.classList.remove("hidden");
    log(`生成完成：${payload.files.join("、")}`);
  } catch (error) {
    log(error.message);
    alert(error.message);
  } finally {
    setBusy(generateBtn, false);
  }
});

el("resetBtn").addEventListener("click", () => {
  uploadForm.reset();
  generateForm.reset();
  state.uploadId = "";
  state.dates = [];
  state.uploadInfo = null;
  dateSelect.innerHTML = "";
  dateSelect.disabled = true;
  previewBtn.disabled = true;
  generateBtn.disabled = true;
  downloadLink.classList.add("hidden");
  warningBox.classList.add("hidden");
  previewTable.innerHTML = "";
  el("sessionText").textContent = "等待上传台账";
  updateMetrics(0, 0, 0, 0);
  log("已重置");
});

el("clearLogBtn").addEventListener("click", () => {
  logOutput.textContent = "";
});

function fillDates(dates) {
  dateSelect.innerHTML = "";
  for (const item of dates) {
    const option = document.createElement("option");
    option.value = item.date;
    option.textContent = `${item.date}（${item.count} 条）`;
    dateSelect.appendChild(option);
  }
  dateSelect.disabled = dates.length === 0;
}

async function refreshPreview() {
  if (!state.uploadId || !dateSelect.value) return;
  try {
    const payload = await requestJson(`/api/preview?upload_id=${encodeURIComponent(state.uploadId)}&date=${encodeURIComponent(dateSelect.value)}`);
    renderPreview(payload);
    log(`预览刷新：${payload.report_date_text}，${payload.total_station_count} 个站点`);
  } catch (error) {
    log(error.message);
  }
}

function selectedTypes() {
  const values = [];
  if (el("typeTransfer").checked) values.push(el("typeTransfer").value);
  if (el("typeClean").checked) values.push(el("typeClean").value);
  return values;
}

function showWarnings(warnings) {
  if (!warnings.length) {
    warningBox.classList.add("hidden");
    warningBox.textContent = "";
    return;
  }
  warningBox.textContent = warnings.join(" ");
  warningBox.classList.remove("hidden");
}

function renderPreview(payload) {
  updateMetrics(
    state.uploadInfo?.record_count || 0,
    payload.total_station_count,
    payload.total_item_count,
    payload.total_photo_count,
  );
  el("previewSubtitle").textContent = `${payload.report_date_text} 聚合结果`;

  const rows = [];
  for (const [typeName, typeData] of Object.entries(payload.types)) {
    for (const station of typeData.stations) {
      rows.push({ typeName, station });
    }
  }
  if (!rows.length) {
    previewTable.innerHTML = `<div class="empty-state">该日期未找到匹配站点</div>`;
    return;
  }
  previewTable.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>类型</th>
          <th>站点</th>
          <th>整体情况</th>
          <th>检查项</th>
          <th>照片</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(({ typeName, station }) => `
          <tr>
            <td><span class="tag">${escapeHtml(typeName)}</span></td>
            <td>${escapeHtml(station.display_name)}</td>
            <td>${escapeHtml(station.summary)}</td>
            <td>${station.item_count}</td>
            <td>${station.photo_count}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function updateMetrics(records, stations, items, photos) {
  el("recordCount").textContent = records;
  el("stationCount").textContent = stations;
  el("itemCount").textContent = items;
  el("photoCount").textContent = photos;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

log("系统就绪");

