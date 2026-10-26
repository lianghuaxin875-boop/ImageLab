const form = document.querySelector("#detectForm");
const imagesInput = document.querySelector("#images");
const imagesFolderInput = document.querySelector("#imagesFolder");
const fileHint = document.querySelector("#fileHint");
const analysisMode = document.querySelector("#analysisMode");
const regionMode = document.querySelector("#regionMode");
const weightMode = document.querySelector("#weightMode");
const weightInputs = {
  clarity: document.querySelector('[name="weight_clarity"]'),
  exposure: document.querySelector('[name="weight_exposure"]'),
  contrast: document.querySelector('[name="weight_contrast"]'),
  noise: document.querySelector('[name="weight_noise"]'),
  artifact: document.querySelector('[name="weight_artifact"]'),
};
const statusText = document.querySelector("#statusText");
const countText = document.querySelector("#countText");
const percentText = document.querySelector("#percentText");
const progressBar = document.querySelector("#progressBar");
const messageText = document.querySelector("#messageText");
const reportPath = document.querySelector("#reportPath");
const downloadReport = document.querySelector("#downloadReport");
const downloadCsvReport = document.querySelector("#downloadCsvReport");
const outputSummary = document.querySelector("#outputSummary");
const outputDirs = document.querySelector("#outputDirs");
const resultRows = document.querySelector("#resultRows");
const resultCount = document.querySelector("#resultCount");
const imageStandardSummary = document.querySelector("#imageStandardSummary");
const logs = document.querySelector("#logs");
const submitButton = form.querySelector('button[type="submit"]');
const navItems = document.querySelectorAll(".nav-item[data-module]");
const quickModuleButtons = document.querySelectorAll("[data-go-module]");
const inputModeSelects = document.querySelectorAll(".input-mode-select");
const currentModuleTitle = document.querySelector("#currentModuleTitle");
const clearCacheButton = document.querySelector("#clearCacheButton");
const cacheStatus = document.querySelector("#cacheStatus");
const homeModule = document.querySelector("#homeModule");
const imageModule = document.querySelector("#imageModule");
const videoModule = document.querySelector("#videoModule");
const videoForm = document.querySelector("#videoForm");
const videosInput = document.querySelector("#videos");
const videosFolderInput = document.querySelector("#videosFolder");
const videoFileHint = document.querySelector("#videoFileHint");
const videoProgressBar = document.querySelector("#videoProgressBar");
const videoMessageText = document.querySelector("#videoMessageText");
const videoReportPath = document.querySelector("#videoReportPath");
const downloadVideoReport = document.querySelector("#downloadVideoReport");
const downloadVideoCsvReport = document.querySelector("#downloadVideoCsvReport");
const videoResultRows = document.querySelector("#videoResultRows");
const videoResultCount = document.querySelector("#videoResultCount");
const videoStandardSummary = document.querySelector("#videoStandardSummary");
const videoLogs = document.querySelector("#videoLogs");
const videoQualityModule = document.querySelector("#videoQualityModule");
const videoQualityForm = document.querySelector("#videoQualityForm");
const qualityVideosInput = document.querySelector("#qualityVideos");
const qualityVideosFolderInput = document.querySelector("#qualityVideosFolder");
const qualityVideoFileHint = document.querySelector("#qualityVideoFileHint");
const qualityProgressBar = document.querySelector("#qualityProgressBar");
const qualityMessageText = document.querySelector("#qualityMessageText");
const qualityReportPath = document.querySelector("#qualityReportPath");
const downloadQualityReport = document.querySelector("#downloadQualityReport");
const downloadQualityCsvReport = document.querySelector("#downloadQualityCsvReport");
const qualityResultRows = document.querySelector("#qualityResultRows");
const qualityResultCount = document.querySelector("#qualityResultCount");
const qualityStandardSummary = document.querySelector("#qualityStandardSummary");
const qualityLogs = document.querySelector("#qualityLogs");
const ocrModule = document.querySelector("#ocrModule");
const ocrForm = document.querySelector("#ocrForm");
const ocrFilesInput = document.querySelector("#ocrFiles");
const ocrFolderInput = document.querySelector("#ocrFolder");
const ocrFileHint = document.querySelector("#ocrFileHint");
const ocrProgressBar = document.querySelector("#ocrProgressBar");
const ocrMessageText = document.querySelector("#ocrMessageText");
const ocrReportPath = document.querySelector("#ocrReportPath");
const downloadOcrReport = document.querySelector("#downloadOcrReport");
const downloadOcrCsvReport = document.querySelector("#downloadOcrCsvReport");
const ocrOutputSummary = document.querySelector("#ocrOutputSummary");
const ocrOutputDirs = document.querySelector("#ocrOutputDirs");
const ocrResultRows = document.querySelector("#ocrResultRows");
const ocrResultCount = document.querySelector("#ocrResultCount");
const ocrStandardSummary = document.querySelector("#ocrStandardSummary");
const ocrLogs = document.querySelector("#ocrLogs");
const ocrPreviewImage = document.querySelector("#ocrPreviewImage");
const ocrPreviewEmpty = document.querySelector("#ocrPreviewEmpty");
const ocrTextPreview = document.querySelector("#ocrTextPreview");
const enhanceModule = document.querySelector("#enhanceModule");
const enhanceForm = document.querySelector("#enhanceForm");
const enhanceImagesInput = document.querySelector("#enhanceImages");
const enhanceImagesFolderInput = document.querySelector("#enhanceImagesFolder");
const enhanceFileHint = document.querySelector("#enhanceFileHint");
const enhanceProgressBar = document.querySelector("#enhanceProgressBar");
const enhanceMessageText = document.querySelector("#enhanceMessageText");
const enhanceReportPath = document.querySelector("#enhanceReportPath");
const downloadEnhanceReport = document.querySelector("#downloadEnhanceReport");
const downloadEnhanceCsvReport = document.querySelector("#downloadEnhanceCsvReport");
const enhanceOutputSummary = document.querySelector("#enhanceOutputSummary");
const enhanceOutputDirs = document.querySelector("#enhanceOutputDirs");
const enhanceResultRows = document.querySelector("#enhanceResultRows");
const enhanceResultCount = document.querySelector("#enhanceResultCount");
const enhanceStandardSummary = document.querySelector("#enhanceStandardSummary");
const enhanceLogs = document.querySelector("#enhanceLogs");
const enhanceComparePreview = document.querySelector("#enhanceComparePreview");
const enhanceCompareEmpty = document.querySelector("#enhanceCompareEmpty");
const enhanceStepsValue = document.querySelector("#enhanceStepsValue");
const enhanceStepOptions = document.querySelectorAll(".enhance-step-option");
const historyModule = document.querySelector("#historyModule");
const settingsModule = document.querySelector("#settingsModule");
const recentHistoryRows = document.querySelector("#recentHistoryRows");
const historyRows = document.querySelector("#historyRows");
const clearHistoryButton = document.querySelector("#clearHistoryButton");
const saveSettingsButton = document.querySelector("#saveSettingsButton");
const settingsStatus = document.querySelector("#settingsStatus");
const mascotWidget = document.querySelector(".mascot-widget");
const mascotTitle = document.querySelector("#mascotTitle");
const mascotText = document.querySelector("#mascotText");
const mascotTime = document.querySelector("#mascotTime");
const petMood = document.querySelector("#petMood");
const petTask = document.querySelector("#petTask");
const petEnergy = document.querySelector("#petEnergy");
const dashboardNodes = {
  todayImages: document.querySelector("#todayImageCount"),
  ok: document.querySelector("#okCount"),
  warning: document.querySelector("#warningCount"),
  ng: document.querySelector("#ngCount"),
  reports: document.querySelector("#reportCount"),
  image: document.querySelector("#imageDetectCount"),
  video: document.querySelector("#videoDetectCount"),
  ocr: document.querySelector("#ocrDetectCount"),
  total: document.querySelector("#totalDetectCount"),
};
const settingsNodes = {
  weightClarity: document.querySelector("#settingWeightClarity"),
  weightExposure: document.querySelector("#settingWeightExposure"),
  weightContrast: document.querySelector("#settingWeightContrast"),
  weightNoise: document.querySelector("#settingWeightNoise"),
  weightArtifact: document.querySelector("#settingWeightArtifact"),
  passScore: document.querySelector("#settingPassScore"),
  warningScore: document.querySelector("#settingWarningScore"),
  ocrQualityThreshold: document.querySelector("#settingOcrQualityThreshold"),
  enhanceGamma: document.querySelector("#settingEnhanceGamma"),
  enhanceClahe: document.querySelector("#settingEnhanceClahe"),
  enhanceDenoise: document.querySelector("#settingEnhanceDenoise"),
  enhanceSharpen: document.querySelector("#settingEnhanceSharpen"),
  ocrMinConfidence: document.querySelector("#settingOcrMinConfidence"),
  ocrMaxGarbled: document.querySelector("#settingOcrMaxGarbled"),
  ocrFrameInterval: document.querySelector("#settingOcrFrameInterval"),
  ocrLang: document.querySelector("#settingOcrLang"),
};

let pollTimer = null;
let currentJobId = "";
let videoPollTimer = null;
let currentVideoJobId = "";
let qualityPollTimer = null;
let currentQualityJobId = "";
let ocrPollTimer = null;
let currentOcrJobId = "";
let enhancePollTimer = null;
let currentEnhanceJobId = "";
let mascotDrag = null;
let mascotDragReady = false;
let dashboardState = loadDashboardState();
let countedDoneJobs = new Set(dashboardState.doneJobs || []);
let countedReports = new Set(dashboardState.reportEvents || []);
let detectionHistory = loadDetectionHistory();
let countedHistoryJobs = new Set(detectionHistory.map((record) => String(record.jobId || "").split(":")[0]).filter(Boolean));

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function restoreMascotPosition() {
  if (!mascotWidget) {
    return;
  }
  try {
    const saved = JSON.parse(localStorage.getItem("codexPetPosition") || "null");
    if (!saved) {
      return;
    }
    const rect = mascotWidget.getBoundingClientRect();
    const maxLeft = Math.max(8, window.innerWidth - rect.width - 8);
    const maxTop = Math.max(8, window.innerHeight - rect.height - 8);
    const left = clamp(Number(saved.left) || 0, 8, maxLeft);
    const top = clamp(Number(saved.top) || 0, 8, maxTop);
    mascotWidget.style.left = `${left}px`;
    mascotWidget.style.top = `${top}px`;
    mascotWidget.style.right = "auto";
    mascotWidget.style.bottom = "auto";
  } catch (_error) {
    localStorage.removeItem("codexPetPosition");
  }
}

function saveMascotPosition() {
  if (!mascotWidget) {
    return;
  }
  const rect = mascotWidget.getBoundingClientRect();
  try {
    localStorage.setItem("codexPetPosition", JSON.stringify({ left: Math.round(rect.left), top: Math.round(rect.top) }));
  } catch (_error) {
    // Dragging should still work when browser storage is unavailable.
  }
}

function initMascotDrag() {
  if (!mascotWidget || mascotDragReady) {
    return;
  }
  mascotDragReady = true;
  mascotWidget.dataset.draggableReady = "true";
  mascotWidget.setAttribute("draggable", "true");
  restoreMascotPosition();

  const moveMascot = (event) => {
    if (!mascotDrag) {
      return;
    }
    event.preventDefault();
    const point = event.touches ? event.touches[0] : event;
    if (!point) {
      return;
    }
    const rect = mascotWidget.getBoundingClientRect();
    const maxLeft = Math.max(8, window.innerWidth - rect.width - 8);
    const maxTop = Math.max(8, window.innerHeight - rect.height - 8);
    const left = clamp(point.clientX - mascotDrag.offsetX, 8, maxLeft);
    const top = clamp(point.clientY - mascotDrag.offsetY, 8, maxTop);
    mascotWidget.style.left = `${left}px`;
    mascotWidget.style.top = `${top}px`;
    mascotWidget.style.right = "auto";
    mascotWidget.style.bottom = "auto";
  };

  const startDrag = (clientX, clientY) => {
    const rect = mascotWidget.getBoundingClientRect();
    mascotDrag = {
      offsetX: clientX - rect.left,
      offsetY: clientY - rect.top,
    };
    mascotWidget.classList.add("dragging");
  };

  const endDrag = (event) => {
    if (!mascotDrag) {
      return;
    }
    mascotDrag = null;
    mascotWidget.classList.remove("dragging");
    try {
      mascotWidget.releasePointerCapture(event.pointerId);
    } catch (_error) {
      // Pointer may already be released by the browser.
    }
    saveMascotPosition();
  };

  mascotWidget.addEventListener("pointerdown", (event) => {
    if (event.button !== undefined && event.button !== 0) {
      return;
    }
    event.preventDefault();
    startDrag(event.clientX, event.clientY);
    try {
      mascotWidget.setPointerCapture(event.pointerId);
    } catch (_error) {
      // Some browser surfaces do not allow capture on every pointer target.
    }
  });
  mascotWidget.addEventListener("mousedown", (event) => {
    if (event.button !== 0) {
      return;
    }
    event.preventDefault();
    startDrag(event.clientX, event.clientY);
  });
  mascotWidget.addEventListener("touchstart", (event) => {
    const touch = event.touches[0];
    if (!touch) {
      return;
    }
    event.preventDefault();
    startDrag(touch.clientX, touch.clientY);
  }, { passive: false });
  mascotWidget.addEventListener("dragstart", (event) => {
    startDrag(event.clientX || mascotWidget.getBoundingClientRect().left, event.clientY || mascotWidget.getBoundingClientRect().top);
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", "codex-pet");
    }
  });
  window.addEventListener("dragover", (event) => {
    if (!mascotDrag) {
      return;
    }
    event.preventDefault();
    moveMascot(event);
  });
  window.addEventListener("drop", endDrag);
  window.addEventListener("dragend", endDrag);
  mascotWidget.addEventListener("pointermove", moveMascot);
  window.addEventListener("pointermove", moveMascot);
  window.addEventListener("mousemove", moveMascot);
  window.addEventListener("touchmove", moveMascot, { passive: false });
  mascotWidget.addEventListener("pointerup", endDrag);
  mascotWidget.addEventListener("pointercancel", endDrag);
  window.addEventListener("pointerup", endDrag);
  window.addEventListener("pointercancel", endDrag);
  window.addEventListener("mouseup", endDrag);
  window.addEventListener("touchend", endDrag);
  window.addEventListener("touchcancel", endDrag);
  mascotWidget.addEventListener("dblclick", () => {
    try {
      localStorage.removeItem("codexPetPosition");
    } catch (_error) {
      // Ignore storage errors while resetting the visible position.
    }
    mascotWidget.style.left = "";
    mascotWidget.style.top = "";
    mascotWidget.style.right = "22px";
    mascotWidget.style.bottom = "22px";
  });
  window.addEventListener("resize", restoreMascotPosition);
}

initMascotDrag();

function formatDuration(seconds) {
  if (seconds === null || seconds === undefined || Number.isNaN(Number(seconds))) {
    return "等待计算";
  }
  const value = Math.max(0, Math.round(Number(seconds)));
  const minutes = Math.floor(value / 60);
  const secs = value % 60;
  if (minutes <= 0) {
    return `${secs} 秒`;
  }
  return `${minutes} 分 ${String(secs).padStart(2, "0")} 秒`;
}

function petDefaults(state) {
  if (state === "running") {
    return { mood: "专注", task: "处理中", energy: 78 };
  }
  if (state === "done") {
    return { mood: "开心", task: "已完成", energy: 100 };
  }
  if (state === "error") {
    return { mood: "警觉", task: "待排查", energy: 42 };
  }
  return { mood: "待命", task: "空闲", energy: 58 };
}

function updateMascot(state, title, text, timeText = "", pet = {}) {
  mascotWidget.classList.remove("idle", "running", "done", "error");
  mascotWidget.classList.add(state);
  mascotTitle.textContent = title;
  mascotText.textContent = text;
  if (mascotTime) {
    mascotTime.textContent = timeText || "预计时间：等待任务";
  }
  const vitals = { ...petDefaults(state), ...pet };
  if (petMood) {
    petMood.textContent = `状态：${vitals.mood}`;
  }
  if (petTask) {
    petTask.textContent = `任务：${vitals.task}`;
  }
  if (petEnergy) {
    petEnergy.style.width = `${Math.max(8, Math.min(100, Number(vitals.energy) || 0))}%`;
  }
}

function normalizeModuleAreas() {
  document.querySelectorAll(".module-workspace .form-panel").forEach((form) => {
    if (form.classList.contains("module-form-grid")) {
      return;
    }
    form.classList.add("module-form-grid");
    const titles = Array.from(form.querySelectorAll(":scope > .area-title"));
    const groups = [
      { match: "文件导入", className: "workspace-card upload-workspace-card" },
      { match: "参数设置", className: "workspace-card params-workspace-card" },
      { match: "操作按钮", className: "workspace-card action-workspace-card" },
    ];

    groups.forEach((group) => {
      const title = titles.find((item) => item.textContent.includes(group.match));
      if (!title || title.parentElement?.classList.contains("workspace-card")) {
        return;
      }
      const wrapper = document.createElement("div");
      wrapper.className = group.className;
      form.insertBefore(wrapper, title);
      let node = title;
      while (node) {
        const next = node.nextElementSibling;
        wrapper.appendChild(node);
        if (next?.classList.contains("area-title")) {
          break;
        }
        node = next;
      }
    });

    const uploadCard = form.querySelector(".upload-workspace-card");
    const actionCard = form.querySelector(".action-workspace-card");
    if (uploadCard && actionCard && actionCard.parentElement !== uploadCard) {
      uploadCard.appendChild(actionCard);
    }
  });

  document.querySelectorAll(".monitor-panel").forEach((panel) => {
    const section = panel.closest(".module-workspace");
    const form = section?.querySelector(".form-panel");
    const titles = Array.from(panel.querySelectorAll(".area-title"));
    const resultTitle = titles.find((title) => title.textContent.includes("检测结果"));
    const reportTitle = titles.find((title) => title.textContent.includes("报告导出"));
    if (!resultTitle || !reportTitle) {
      return;
    }
    const resultBox = resultTitle.nextElementSibling;
    resultTitle.classList.add("module-result-title");
    resultBox?.classList.add("module-result-box");
    if (form && resultBox) {
      form.insertAdjacentElement("afterend", resultTitle);
      resultTitle.insertAdjacentElement("afterend", resultBox);
      ensureResultChart(resultBox);
      return;
    }
    if (resultBox) {
      panel.insertBefore(resultTitle, reportTitle);
      panel.insertBefore(resultBox, reportTitle);
      ensureResultChart(resultBox);
    }
  });
}

function ensureResultChart(resultBox) {
  if (!resultBox || resultBox.querySelector(".module-chart-card")) {
    return;
  }
  const chart = document.createElement("div");
  chart.className = "module-chart-card";
  chart.innerHTML = `
    <div class="chart-head">
      <strong>数据分布</strong>
      <span>检测结果统计</span>
    </div>
    <div class="chart-bars" aria-hidden="true">
      <i style="height: 24%"><b>1</b></i>
      <i style="height: 38%"><b>2</b></i>
      <i style="height: 56%"><b>4</b></i>
      <i style="height: 78%"><b>6</b></i>
      <i style="height: 42%"><b>3</b></i>
      <i style="height: 30%"><b>2</b></i>
    </div>
    <div class="chart-legend">
      <span><i></i>良好</span>
      <span><i></i>警告</span>
      <span><i></i>不合格</span>
    </div>
  `;
  const tableWrap = resultBox.querySelector(".table-wrap");
  if (tableWrap) {
    tableWrap.insertAdjacentElement("afterend", chart);
  } else {
    resultBox.appendChild(chart);
  }
}

function switchModule(moduleName) {
  document.body.dataset.module = moduleName;
  navItems.forEach((item) => item.classList.toggle("active", item.dataset.module === moduleName));
  if (currentModuleTitle) {
    currentModuleTitle.textContent = moduleLabel(moduleName);
  }
  homeModule.classList.toggle("active", moduleName === "home");
  imageModule.classList.toggle("active", moduleName === "image");
  videoModule.classList.toggle("active", moduleName === "video");
  videoQualityModule.classList.toggle("active", moduleName === "videoQuality");
  ocrModule.classList.toggle("active", moduleName === "ocr");
  enhanceModule.classList.toggle("active", moduleName === "enhance");
  historyModule.classList.toggle("active", moduleName === "history");
  settingsModule.classList.toggle("active", moduleName === "settings");
  window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  const mainContent = document.querySelector(".main");
  if (mainContent) {
    mainContent.scrollTop = 0;
  }
  renderHistoryTables();
  const isVideoQuality = moduleName === "videoQuality";
  const isVideo = moduleName === "video";
  const isOcr = moduleName === "ocr";
  const isEnhance = moduleName === "enhance";
  const isHome = moduleName === "home";
  const isHistory = moduleName === "history";
  const isSettings = moduleName === "settings";
  updateMascot(
    "idle",
    isHome ? "CodexPet 首页总览" : isHistory ? "CodexPet 历史记录" : isSettings ? "CodexPet 参数设置" : isEnhance ? "CodexPet 图像增强模式" : isOcr ? "CodexPet OCR 2.0 模式" : isVideoQuality ? "CodexPet 综合评分模式" : isVideo ? "CodexPet 视频模式" : "CodexPet 图片模式",
    isHome ? "选择一个功能卡片就能开始处理" : isHistory ? "这里会保存最近检测记录和报告路径" : isSettings ? "常用参数会保存到当前浏览器" : isEnhance ? "我会先检测质量问题，再选择增强策略并生成前后对比" : isOcr ? "我会盯住识别文字、NG 子代码和多 Sheet 报告" : isVideoQuality ? "我会盯住总分、异常截图和剩余时间" : isVideo ? "我会盯住码率、白平衡和剩余时间" : "我会盯住评分、标注图和报告状态",
    "预计时间：等待任务",
    { task: isHome ? "首页" : isHistory ? "历史" : isSettings ? "设置" : isEnhance ? "图像增强" : isOcr ? "OCR 识别" : isVideoQuality ? "综合评分" : isVideo ? "视频巡检" : "图片巡检" },
  );
}

function setProgress(percent) {
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  progressBar.style.width = `${value}%`;
  percentText.textContent = `${value}%`;
}

function setBar(bar, percent) {
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  bar.style.width = `${value}%`;
}

function syncRegionState() {
  regionMode.disabled = analysisMode.value !== "区域检测";
}

function syncWeights() {
  const mode = weightMode.value;
  const custom = mode === "自定义模式";
  const values = window.WEIGHT_MODES[mode];

  if (values) {
    weightInputs.clarity.value = values[0].toFixed(2);
    weightInputs.exposure.value = values[1].toFixed(2);
    weightInputs.contrast.value = values[2].toFixed(2);
    weightInputs.noise.value = values[3].toFixed(2);
    weightInputs.artifact.value = values[4].toFixed(2);
  }

  Object.values(weightInputs).forEach((input) => {
    input.disabled = !custom;
  });
}

function refreshInputHint(form) {
  if (!form) {
    return;
  }
  if (form.id === "detectForm") updateFilesHint();
  else if (form.id === "videoForm") updateVideoFilesHint();
  else if (form.id === "videoQualityForm") updateQualityVideoFilesHint();
  else if (form.id === "ocrForm") updateOcrFilesHint();
  else if (form.id === "enhanceForm") updateEnhanceFilesHint();
}

function syncInputMode(select) {
  const form = document.querySelector(select.dataset.form);
  if (!form) {
    return;
  }
  const mode = select.value;
  form.querySelectorAll("[data-input-mode]").forEach((element) => {
    const active = element.dataset.inputMode === mode;
    element.classList.toggle("input-mode-hidden", !active);
    element.querySelectorAll("input, select, textarea, button").forEach((control) => {
      if (control === select) {
        return;
      }
      control.disabled = !active;
      if (!active && control.type === "file") {
        control.value = "";
      }
    });
  });
  refreshInputHint(form);
}
function selectedFiles(...inputs) {
  return inputs.flatMap((input) => input ? Array.from(input.files || []) : []);
}

function filePreview(files) {
  return files.slice(0, 3).map((file) => file.webkitRelativePath || file.name).join("，");
}

function updateFilesHint() {
  const files = selectedFiles(imagesInput, imagesFolderInput);
  const count = files.length;
  if (!count) {
    fileHint.textContent = "未选择文件";
    return;
  }
  const preview = filePreview(files);
  fileHint.textContent = count > 3 ? `已选择 ${count} 个文件：${preview}...` : `已选择 ${count} 个文件：${preview}`;
}

function updateVideoFilesHint() {
  const files = selectedFiles(videosInput, videosFolderInput);
  const count = files.length;
  if (!count) {
    videoFileHint.textContent = "未选择文件";
    return;
  }
  const preview = filePreview(files);
  videoFileHint.textContent = count > 3 ? `已选择 ${count} 个视频：${preview}...` : `已选择 ${count} 个视频：${preview}`;
}

function updateQualityVideoFilesHint() {
  const files = selectedFiles(qualityVideosInput, qualityVideosFolderInput);
  const count = files.length;
  if (!count) {
    qualityVideoFileHint.textContent = "未选择文件";
    return;
  }
  const preview = filePreview(files);
  qualityVideoFileHint.textContent = count > 3 ? `已选择 ${count} 个视频：${preview}...` : `已选择 ${count} 个视频：${preview}`;
}

function updateOcrFilesHint() {
  const files = selectedFiles(ocrFilesInput, ocrFolderInput);
  const count = files.length;
  if (!count) {
    ocrFileHint.textContent = "未选择文件";
    return;
  }
  const preview = filePreview(files);
  ocrFileHint.textContent = count > 3 ? `已选择 ${count} 个 OCR 文件：${preview}...` : `已选择 ${count} 个 OCR 文件：${preview}`;
  const firstImage = files.find((file) => file.type.startsWith("image/"));
  if (firstImage) {
    ocrPreviewImage.src = URL.createObjectURL(firstImage);
    ocrPreviewImage.parentElement.classList.add("has-image");
    ocrPreviewEmpty.textContent = "";
  }
}

function updateEnhanceFilesHint() {
  const files = selectedFiles(enhanceImagesInput, enhanceImagesFolderInput);
  const count = files.length;
  if (!count) {
    enhanceFileHint.textContent = "未选择文件";
    return;
  }
  const preview = filePreview(files);
  enhanceFileHint.textContent = count > 3 ? `已选择 ${count} 张待增强图片：${preview}...` : `已选择 ${count} 张待增强图片：${preview}`;
}

function syncEnhanceSteps() {
  if (!enhanceStepsValue || !enhanceStepOptions.length) {
    return;
  }
  const selected = Array.from(enhanceStepOptions)
    .filter((option) => option.checked)
    .map((option) => option.value);
  enhanceStepsValue.value = selected.length ? selected.join(",") : "gamma,clahe,sharpen";
}
function levelClass(level) {
  return `level-${String(level || "").toLowerCase()}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function localDateKey() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function emptyDashboardState() {
  return {
    date: localDateKey(),
    todayImages: 0,
    ok: 0,
    warning: 0,
    ng: 0,
    reports: 0,
    image: 0,
    video: 0,
    ocr: 0,
    total: 0,
    doneJobs: [],
    reportEvents: [],
  };
}

function loadDashboardState() {
  try {
    const parsed = JSON.parse(localStorage.getItem("imageLabDashboardStats") || "null");
    if (parsed && parsed.date === localDateKey()) {
      return { ...emptyDashboardState(), ...parsed };
    }
  } catch (_error) {
    localStorage.removeItem("imageLabDashboardStats");
  }
  return emptyDashboardState();
}

function saveDashboardState() {
  dashboardState.doneJobs = Array.from(countedDoneJobs).slice(-200);
  dashboardState.reportEvents = Array.from(countedReports).slice(-200);
  localStorage.setItem("imageLabDashboardStats", JSON.stringify(dashboardState));
}

function renderDashboard() {
  if (!dashboardNodes.todayImages) {
    return;
  }
  dashboardNodes.todayImages.textContent = `${dashboardState.todayImages} 张`;
  if (dashboardNodes.ok) dashboardNodes.ok.textContent = `${dashboardState.ok} 张`;
  if (dashboardNodes.warning) dashboardNodes.warning.textContent = `${dashboardState.warning} 张`;
  if (dashboardNodes.ng) dashboardNodes.ng.textContent = `${dashboardState.ng} 张`;
  if (dashboardNodes.reports) dashboardNodes.reports.textContent = `${dashboardState.reports} 份`;
  if (dashboardNodes.image) dashboardNodes.image.textContent = dashboardState.image;
  if (dashboardNodes.video) dashboardNodes.video.textContent = dashboardState.video;
  if (dashboardNodes.ocr) dashboardNodes.ocr.textContent = dashboardState.ocr;
  if (dashboardNodes.total) dashboardNodes.total.textContent = dashboardState.total;
}

function loadDetectionHistory() {
  try {
    const parsed = JSON.parse(localStorage.getItem("imageLabDetectionHistory") || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    localStorage.removeItem("imageLabDetectionHistory");
    return [];
  }
}

function saveDetectionHistory() {
  detectionHistory = detectionHistory.slice(0, 200);
  localStorage.setItem("imageLabDetectionHistory", JSON.stringify(detectionHistory));
}

function moduleLabel(moduleName) {
  const labels = {
    image: "图片质量检测",
    enhance: "图像增强",
    video: "视频码率分析",
    videoQuality: "视频综合检测",
    video_quality: "视频综合检测",
    ocr: "OCR文字识别",
  };
  return labels[moduleName] || moduleName || "-";
}

function rowResultText(row, moduleName) {
  if (moduleName === "enhance") {
    return row.effect || row.after_level || "-";
  }
  return row.level || row.judgement || row.status || row.after_level || row.quality_level || "-";
}

function rowScoreText(row, moduleName) {
  const value = moduleName === "enhance" ? row.after_score : row.score ?? row.quality_score ?? row.avg_confidence ?? "";
  if (value === "" || value === undefined || value === null) {
    return "-";
  }
  const number = Number(value);
  return Number.isFinite(number) ? number.toFixed(1) : String(value);
}

function addHistoryFromJob(job, moduleName) {
  if (!job || job.status !== "done" || countedHistoryJobs.has(job.id)) {
    return;
  }
  const rows = job.result_rows || [];
  const report = job.report_path || job.csv_report_path || "";
  const created = new Date().toLocaleString("zh-CN", { hour12: false });
  const records = rows.slice(0, 100).map((row) => ({
    id: `${job.id}:${row.filename || row.file || Math.random()}`,
    jobId: job.id,
    time: created,
    module: moduleLabel(moduleName),
    filename: row.filename || "-",
    result: rowResultText(row, moduleName),
    score: rowScoreText(row, moduleName),
    report,
  }));
  if (!records.length) {
    records.push({
      id: `${job.id}:empty`,
      jobId: job.id,
      time: created,
      module: moduleLabel(moduleName),
      filename: "-",
      result: "完成",
      score: "-",
      report,
    });
  }
  detectionHistory = [...records, ...detectionHistory].slice(0, 200);
  countedHistoryJobs.add(job.id);
  saveDetectionHistory();
  renderHistoryTables();
}

function renderHistoryTable(target, rows, emptyText) {
  if (!target) {
    return;
  }
  const compact = target.id === "recentHistoryRows";
  if (!rows.length) {
    target.innerHTML = `<tr><td colspan="${compact ? 3 : 6}">${emptyText}</td></tr>`;
    return;
  }
  if (compact) {
    target.innerHTML = rows.map((record) => `
      <tr>
        <td title="${escapeHtml(record.filename)}">
          <strong>${escapeHtml(record.filename)}</strong>
          <small>${escapeHtml(record.module)} | ${escapeHtml(record.time)}</small>
        </td>
        <td><span class="level-pill ${levelClass(record.result)}">${escapeHtml(record.result)}</span></td>
        <td>${escapeHtml(record.score)}</td>
      </tr>
    `).join("");
    return;
  }
  target.innerHTML = rows.map((record) => `
    <tr>
      <td>${escapeHtml(record.time)}</td>
      <td>${escapeHtml(record.module)}</td>
      <td title="${escapeHtml(record.filename)}">${escapeHtml(record.filename)}</td>
      <td><span class="level-pill ${levelClass(record.result)}">${escapeHtml(record.result)}</span></td>
      <td>${escapeHtml(record.score)}</td>
      <td title="${escapeHtml(record.report)}">${escapeHtml(record.report || "-")}</td>
    </tr>
  `).join("");
}

function renderHistoryTables() {
  renderHistoryTable(recentHistoryRows, detectionHistory.slice(0, 5), "暂无检测记录");
  renderHistoryTable(historyRows, detectionHistory, "暂无检测历史");
}

function loadSettings() {
  try {
    return JSON.parse(localStorage.getItem("imageLabSettings") || "{}") || {};
  } catch (_error) {
    localStorage.removeItem("imageLabSettings");
    return {};
  }
}

function collectSettings() {
  return Object.fromEntries(Object.entries(settingsNodes).map(([key, node]) => [key, node ? node.value : ""]));
}

function fillSettingsPanel(settings) {
  Object.entries(settings).forEach(([key, value]) => {
    if (settingsNodes[key] && value !== undefined) {
      settingsNodes[key].value = value;
    }
  });
}

function applySettingsToForms(settings) {
  if (settings.weightClarity) weightInputs.clarity.value = settings.weightClarity;
  if (settings.weightExposure) weightInputs.exposure.value = settings.weightExposure;
  if (settings.weightContrast) weightInputs.contrast.value = settings.weightContrast;
  if (settings.weightNoise) weightInputs.noise.value = settings.weightNoise;
  if (settings.weightArtifact) weightInputs.artifact.value = settings.weightArtifact;
  if (settings.ocrQualityThreshold) document.querySelector('[name="ocr_quality_review_threshold"]').value = settings.ocrQualityThreshold;
  if (settings.ocrMinConfidence) document.querySelector('[name="ocr_min_confidence"]').value = settings.ocrMinConfidence;
  if (settings.ocrMaxGarbled) document.querySelector('[name="ocr_max_garbled"]').value = settings.ocrMaxGarbled;
  if (settings.ocrFrameInterval) document.querySelector('[name="ocr_frame_interval"]').value = settings.ocrFrameInterval;
  if (settings.ocrLang) document.querySelector('[name="ocr_lang"]').value = settings.ocrLang;
  if (settings.enhanceGamma) document.querySelector('[name="enhance_gamma"]').value = settings.enhanceGamma;
  if (settings.enhanceClahe) document.querySelector('[name="enhance_clahe_clip"]').value = settings.enhanceClahe;
  if (settings.enhanceDenoise) document.querySelector('[name="enhance_denoise_h"]').value = settings.enhanceDenoise;
  if (settings.enhanceSharpen) document.querySelector('[name="enhance_sharpen_amount"]').value = settings.enhanceSharpen;
}

function saveSettings() {
  const settings = collectSettings();
  localStorage.setItem("imageLabSettings", JSON.stringify(settings));
  applySettingsToForms(settings);
  settingsStatus.textContent = "参数已保存，并已同步到对应模块页面。";
}

function initSettings() {
  const settings = loadSettings();
  fillSettingsPanel(settings);
  applySettingsToForms(settings);
}

function addImageLevels(rows) {
  rows.forEach((row) => {
    const level = String(row.level || row.after_level || row.status || "").toUpperCase();
    if (level.includes("OK")) {
      dashboardState.ok += 1;
    } else if (level.includes("WARNING")) {
      dashboardState.warning += 1;
    } else if (level.includes("NG") || level.includes("ERROR")) {
      dashboardState.ng += 1;
    }
  });
}

function recordCompletedJob(job, moduleName) {
  if (!job || job.status !== "done" || countedDoneJobs.has(job.id)) {
    return;
  }
  countedDoneJobs.add(job.id);
  const rows = job.result_rows || [];
  const count = rows.length || Number(job.total) || 0;
  if (moduleName === "image") {
    dashboardState.image += count;
    dashboardState.todayImages += count;
    addImageLevels(rows);
  } else if (moduleName === "ocr") {
    dashboardState.ocr += count;
  } else {
    dashboardState.video += count;
  }
  dashboardState.total += count;
  saveDashboardState();
  renderDashboard();
}

function recordReportExport(jobId, reportType) {
  if (!jobId) {
    return;
  }
  const eventId = `${jobId}:${reportType}`;
  if (countedReports.has(eventId)) {
    return;
  }
  countedReports.add(eventId);
  dashboardState.reports += 1;
  saveDashboardState();
  renderDashboard();
}

function setStandardSummary(container, values) {
  if (!container) {
    return;
  }
  const cells = container.querySelectorAll("strong");
  const ordered = [
    values.score || "--",
    values.level || "--",
    values.reason || "暂无问题",
    values.advice || "无需处理",
    values.report || "等待生成",
  ];
  cells.forEach((cell, index) => {
    cell.textContent = ordered[index] || "--";
    cell.title = ordered[index] || "";
  });
}

function fileUrl(path) {
  return path ? `/api/files?path=${encodeURIComponent(path)}` : "";
}

function resultThumb(row, type = "image") {
  const path = row.compare_path || row.enhanced_path || row.output_path || row.annotated_path || row.preview_path || "";
  if (path) {
    return `<div class="result-thumb has-image"><img src="${fileUrl(path)}" alt="缩略图"></div>`;
  }
  const labels = {
    image: "IMG",
    video: "VID",
    ocr: "OCR",
    enhance: "ENH",
  };
  return `<div class="result-thumb ${escapeHtml(type)}"><span>${labels[type] || "FILE"}</span></div>`;
}

function averageNumber(rows, key) {
  const values = rows.map((row) => Number(row[key])).filter((value) => Number.isFinite(value));
  if (!values.length) {
    return "";
  }
  return (values.reduce((sum, value) => sum + value, 0) / values.length).toFixed(1);
}

function summarizeLevel(rows, key) {
  const levels = rows.map((row) => String(row[key] || "").toUpperCase());
  if (levels.some((level) => level.includes("ERROR"))) {
    return "ERROR";
  }
  if (levels.some((level) => level.includes("NG"))) {
    return "NG";
  }
  if (levels.some((level) => level.includes("WARNING"))) {
    return "WARNING";
  }
  if (levels.some((level) => level.includes("OK"))) {
    return "OK";
  }
  return rows.length ? rows[0][key] || "--" : "--";
}

function firstText(rows, keys, fallback = "暂无问题") {
  for (const row of rows) {
    for (const key of keys) {
      const value = row[key];
      if (value !== undefined && value !== null && String(value).trim()) {
        return String(value);
      }
    }
  }
  return fallback;
}

function renderResults(rows) {
  if (!rows || !rows.length) {
    resultCount.textContent = "暂无结果";
    setStandardSummary(imageStandardSummary, {
      reason: "等待检测",
      advice: "等待检测",
      report: reportPath.textContent,
    });
    resultRows.innerHTML = '<tr><td colspan="14">检测完成后在这里查看评分</td></tr>';
    return;
  }

  resultCount.textContent = `${rows.length} 张图片`;
  const imageLevel = summarizeLevel(rows, "level");
  setStandardSummary(imageStandardSummary, {
    score: averageNumber(rows, "score"),
    level: imageLevel,
    reason: firstText(rows, ["summary", "histogram", "color_cast"]),
    advice: imageLevel === "OK" ? "质量合格，可直接使用" : "建议复核问题项并重新采集或修图",
    report: reportPath.textContent,
  });
  resultRows.innerHTML = rows.map((row) => `
    <tr>
      <td>${resultThumb(row, "image")}</td>
      <td title="${escapeHtml(row.summary)}">${escapeHtml(row.filename)}</td>
      <td>${Number(row.score || 0).toFixed(1)}</td>
      <td><span class="level-pill ${levelClass(row.level)}">${escapeHtml(row.level)}</span></td>
      <td>${Number(row.clarity || 0).toFixed(1)}</td>
      <td>${Number(row.exposure || 0).toFixed(1)}</td>
      <td>${Number(row.contrast || 0).toFixed(1)}</td>
      <td>${Number(row.noise || 0).toFixed(1)}</td>
      <td>${Number(row.artifact || 0).toFixed(1)}</td>
      <td>${escapeHtml(row.color_cast || "")}</td>
      <td>${row.white_balance_score === "" ? "" : Number(row.white_balance_score || 0).toFixed(1)}</td>
      <td>${escapeHtml(row.white_balance || "")}</td>
      <td>${escapeHtml(row.histogram || "")}</td>
      <td title="${escapeHtml(row.output_path || row.output_status)}">${escapeHtml(row.output_status || "未知")}</td>
    </tr>
  `).join("");
}

function renderOutputInfo(info) {
  if (!info || !Object.keys(info).length) {
    outputSummary.textContent = "等待检测";
    outputDirs.textContent = "OK/NG 输出目录会在检测完成后显示";
    return;
  }
  outputSummary.textContent = `已保存 ${info.saved_count || 0} 张，失败 ${info.failed_count || 0} 张`;
  outputDirs.textContent = `OK：${info.ok_dir || "-"} ｜ NG：${info.ng_dir || "-"} ｜ 报告：${info.report_dir || "-"}`;
}

function renderVideoResults(rows) {
  if (!rows || !rows.length) {
    videoResultCount.textContent = "暂无结果";
    setStandardSummary(videoStandardSummary, {
      reason: "等待检测",
      advice: "等待检测",
      report: videoReportPath.textContent,
    });
    videoResultRows.innerHTML = '<tr><td colspan="12">检测完成后在这里查看视频结果</td></tr>';
    return;
  }
  videoResultCount.textContent = `${rows.length} 个视频`;
  const okCount = rows.filter((row) => String(row.judgement || "").toUpperCase().includes("OK")).length;
  const score = `${((okCount / rows.length) * 100).toFixed(1)}%`;
  const videoLevel = summarizeLevel(rows, "judgement");
  setStandardSummary(videoStandardSummary, {
    score,
    level: videoLevel,
    reason: firstText(rows, ["error", "white_balance"], videoLevel === "OK" ? "码率、帧率、分辨率达标" : "存在码率或视频参数未达标"),
    advice: videoLevel === "OK" ? "达标，可直接归档" : "建议核对导出码率、分辨率和帧率设置",
    report: videoReportPath.textContent,
  });
  videoResultRows.innerHTML = rows.map((row) => `
    <tr>
      <td>${resultThumb(row, "video")}</td>
      <td title="${escapeHtml(row.error)}">${escapeHtml(row.filename)}</td>
      <td>${escapeHtml(row.resolution)}</td>
      <td>${escapeHtml(row.video_bitrate)}</td>
      <td>${escapeHtml(row.audio_bitrate)}</td>
      <td>${escapeHtml(row.fps)}</td>
      <td>${escapeHtml(row.frames)}</td>
      <td>${escapeHtml(row.threshold)}</td>
      <td><span class="level-pill ${levelClass(row.judgement)}">${escapeHtml(row.judgement)}</span></td>
      <td>${row.white_balance_score === "" ? "" : Number(row.white_balance_score || 0).toFixed(1)}</td>
      <td>${escapeHtml(row.white_balance)}</td>
      <td>${row.white_balance_abnormal === "" ? "" : `${(Number(row.white_balance_abnormal || 0) * 100).toFixed(1)}%`}</td>
    </tr>
  `).join("");
}

function renderQualityResults(rows) {
  if (!rows || !rows.length) {
    qualityResultCount.textContent = "暂无结果";
    setStandardSummary(qualityStandardSummary, {
      reason: "等待检测",
      advice: "等待检测",
      report: qualityReportPath.textContent,
    });
    qualityResultRows.innerHTML = '<tr><td colspan="13">检测完成后在这里查看综合评分</td></tr>';
    return;
  }
  qualityResultCount.textContent = `${rows.length} 个视频`;
  const qualityLevel = summarizeLevel(rows, "level");
  setStandardSummary(qualityStandardSummary, {
    score: averageNumber(rows, "score"),
    level: qualityLevel,
    reason: firstText(rows, ["error", "bitrate", "sharpness", "exposure", "stability"]),
    advice: qualityLevel === "OK" ? "质量合格，可直接归档" : "建议复核异常帧并调整拍摄或导出参数",
    report: qualityReportPath.textContent,
  });
  qualityResultRows.innerHTML = rows.map((row) => `
    <tr>
      <td>${resultThumb(row, "video")}</td>
      <td title="${escapeHtml(row.error)}">${escapeHtml(row.filename)}</td>
      <td>${escapeHtml(row.resolution)}</td>
      <td>${Number(row.score || 0).toFixed(1)}</td>
      <td><span class="level-pill ${levelClass(row.level)}">${escapeHtml(row.level)}</span></td>
      <td>${escapeHtml(row.bitrate)}</td>
      <td>${escapeHtml(row.sharpness)}</td>
      <td>${escapeHtml(row.exposure)}</td>
      <td>${escapeHtml(row.stability)}</td>
      <td>${escapeHtml(row.drops)}</td>
      <td>${escapeHtml(row.black_white)}</td>
      <td>${escapeHtml(row.frozen)}</td>
      <td>${escapeHtml(row.audio)}</td>
    </tr>
  `).join("");
}

function renderOcrResults(rows) {
  if (!rows || !rows.length) {
    ocrResultCount.textContent = "暂无结果";
    setStandardSummary(ocrStandardSummary, {
      reason: "等待识别",
      advice: "等待识别",
      report: ocrReportPath.textContent,
    });
    ocrTextPreview.textContent = "OCR 完成后显示识别文本和置信度摘要";
    ocrResultRows.innerHTML = '<tr><td colspan="14">OCR 完成后在这里查看识别结果</td></tr>';
    return;
  }
  ocrResultCount.textContent = `${rows.length} 个文件`;
  const ocrLevel = summarizeLevel(rows, "status");
  setStandardSummary(ocrStandardSummary, {
    score: averageNumber(rows, "quality_score") || averageNumber(rows, "avg_confidence"),
    level: ocrLevel,
    reason: firstText(rows, ["reasons", "quality_level"]),
    advice: firstText(rows, ["review_advice"], ocrLevel === "OK" ? "识别可信，可导出文本" : "识别结果建议人工复核"),
    report: ocrReportPath.textContent,
  });
  const firstRow = rows[0] || {};
  ocrTextPreview.textContent = `置信度：${Number(firstRow.avg_confidence || 0).toFixed(3)}\n\n${firstRow.text || "暂无识别文本"}`;
  ocrResultRows.innerHTML = rows.map((row) => `
    <tr>
      <td>${resultThumb(row, "ocr")}</td>
      <td title="${escapeHtml(row.preprocess || "")}">${escapeHtml(row.filename)}</td>
      <td>${row.quality_score === "" ? "" : Number(row.quality_score || 0).toFixed(1)}</td>
      <td>${escapeHtml(row.quality_level || "")}</td>
      <td>${escapeHtml(row.review_advice || "")}</td>
      <td><span class="level-pill ${levelClass(row.status)}">${escapeHtml(row.status)}</span></td>
      <td>${escapeHtml(row.sub_codes)}</td>
      <td>${escapeHtml(row.ocr_count)}</td>
      <td>${Number(row.avg_confidence || 0).toFixed(3)}</td>
      <td>${Number(row.min_confidence || 0).toFixed(3)}</td>
      <td>${(Number(row.low_conf_ratio || 0) * 100).toFixed(1)}%</td>
      <td>${Number(row.hit_rate || 0).toFixed(1)}%</td>
      <td title="${escapeHtml(row.text)}">${escapeHtml(String(row.text || "").slice(0, 80))}</td>
      <td title="${escapeHtml(row.reasons)}">${escapeHtml(row.reasons)}</td>
    </tr>
  `).join("");
}

function renderEnhanceResults(rows) {
  if (!rows || !rows.length) {
    enhanceResultCount.textContent = "暂无结果";
    setStandardSummary(enhanceStandardSummary, {
      reason: "等待增强",
      advice: "等待增强",
      report: enhanceReportPath.textContent,
    });
    enhanceComparePreview.removeAttribute("src");
    enhanceComparePreview.parentElement.classList.remove("has-image");
    enhanceCompareEmpty.textContent = "增强完成后显示对比图";
    enhanceResultRows.innerHTML = '<tr><td colspan="8">增强完成后在这里查看前后对比结果</td></tr>';
    return;
  }
  enhanceResultCount.textContent = `${rows.length} 张图片`;
  const level = summarizeLevel(rows, "after_level");
  setStandardSummary(enhanceStandardSummary, {
    score: averageNumber(rows, "after_score"),
    level,
    reason: firstText(rows, ["after_issue", "before_issue"]),
    advice: firstText(rows, ["advice"], "建议结合对比图复核增强效果"),
    report: enhanceReportPath.textContent,
  });
  const firstRow = rows[0] || {};
  if (firstRow.compare_path) {
    enhanceComparePreview.src = `/api/files?path=${encodeURIComponent(firstRow.compare_path)}`;
    enhanceComparePreview.parentElement.classList.add("has-image");
    enhanceCompareEmpty.textContent = "";
  }
  enhanceResultRows.innerHTML = rows.map((row) => `
    <tr>
      <td>${resultThumb(row, "enhance")}</td>
      <td title="${escapeHtml(row.compare_path || "")}">${escapeHtml(row.filename)}</td>
      <td>${Number(row.before_score || 0).toFixed(1)}</td>
      <td>${Number(row.after_score || 0).toFixed(1)}</td>
      <td>${Number(row.delta || 0).toFixed(1)}</td>
      <td><span class="level-pill ${levelClass(row.effect)}">${escapeHtml(row.effect)}</span></td>
      <td title="${escapeHtml(row.steps)}">${escapeHtml(row.steps)}</td>
      <td title="${escapeHtml(row.advice)}">${escapeHtml(row.advice)}</td>
    </tr>
  `).join("");
}

function renderEnhanceOutputInfo(info) {
  if (!info || !Object.keys(info).length) {
    enhanceOutputSummary.textContent = "等待增强";
    enhanceOutputDirs.textContent = "增强图、对比图、报告和日志目录会在完成后显示";
    return;
  }
  enhanceOutputSummary.textContent = `已增强 ${info.saved_count || 0} 张，失败 ${info.failed_count || 0} 张`;
  enhanceOutputDirs.textContent = `增强图：${info.enhanced_dir || "-"} · 对比图：${info.compare_dir || "-"} · 报告：${info.report_dir || "-"} · 日志：${info.log_path || "-"}`;
}

function renderVideoJob(job) {
  const statusMap = { queued: "排队中", running: "运行中", done: "完成", error: "失败" };
  const etaText = formatDuration(job.eta_seconds);
  const elapsedText = formatDuration(job.elapsed_seconds || 0);
  const progressText = `${job.done || 0}/${job.total || 0}`;
  videoMessageText.textContent = `${statusMap[job.status] || job.status}：${job.message || "等待更新"}；预计还需 ${etaText}`;
  setBar(videoProgressBar, job.percent || 0);
  videoReportPath.textContent = job.report_path || (job.status === "done" ? "尚未生成，点击按钮后下载" : "检测完成后可按需生成并下载");
  downloadVideoReport.disabled = job.status !== "done";
  downloadVideoCsvReport.disabled = job.status !== "done";
  renderVideoResults(job.result_rows || []);
  videoLogs.textContent = job.logs && job.logs.length ? job.logs.join("\n") : "[ready] 等待视频检测任务...";
  videoLogs.scrollTop = videoLogs.scrollHeight;
  recordCompletedJob(job, "video");
  addHistoryFromJob(job, "video");
  if (job.status === "running") {
    updateMascot(
      "running",
      "CodexPet 分析视频中",
      `${progressText}，${job.message || "正在读取视频码率与白平衡"}`,
      `已用时 ${elapsedText} · 预计还需 ${etaText}`,
      {
        mood: "专注",
        task: `视频 ${progressText}`,
        energy: Math.max(12, Math.min(96, Number(job.percent || 0))),
      },
    );
  } else if (job.status === "done") {
    updateMascot("done", "CodexPet 已完成", "视频码率和白平衡结果已生成，可以下载报告", `总耗时 ${elapsedText}`, { task: "报告就绪" });
  } else if (job.status === "error") {
    updateMascot("error", "CodexPet 检测失败", "检查 ffprobe 路径或日志提示", `已用时 ${elapsedText}`, { task: "等待修复" });
  }
}

function renderQualityJob(job) {
  const statusMap = { queued: "排队中", running: "运行中", done: "完成", error: "失败" };
  const etaText = formatDuration(job.eta_seconds);
  const elapsedText = formatDuration(job.elapsed_seconds || 0);
  const progressText = `${job.done || 0}/${job.total || 0}`;
  qualityMessageText.textContent = `${statusMap[job.status] || job.status}：${job.message || "等待更新"}；预计还需 ${etaText}`;
  setBar(qualityProgressBar, job.percent || 0);
  qualityReportPath.textContent = job.report_path || (job.status === "done" ? "尚未生成，点击按钮后下载" : "检测完成后可按需生成并下载");
  downloadQualityReport.disabled = job.status !== "done";
  downloadQualityCsvReport.disabled = job.status !== "done";
  renderQualityResults(job.result_rows || []);
  qualityLogs.textContent = job.logs && job.logs.length ? job.logs.join("\n") : "[ready] 等待视频综合评分任务...";
  qualityLogs.scrollTop = qualityLogs.scrollHeight;
  recordCompletedJob(job, "video");
  addHistoryFromJob(job, "video_quality");
  if (job.status === "running") {
    updateMascot(
      "running",
      "CodexPet 综合评分中",
      `${progressText}，${job.message || "正在检测视频综合质量"}`,
      `已用时 ${elapsedText} · 预计还需 ${etaText}`,
      { mood: "专注", task: `综合 ${progressText}`, energy: Math.max(12, Math.min(96, Number(job.percent || 0))) },
    );
  } else if (job.status === "done") {
    updateMascot("done", "CodexPet 已完成", "视频综合评分已生成，可以下载报告", `总耗时 ${elapsedText}`, { task: "综合报告就绪" });
  } else if (job.status === "error") {
    updateMascot("error", "CodexPet 评分失败", "检查 ffprobe/ffmpeg 路径或日志提示", `已用时 ${elapsedText}`, { task: "等待修复" });
  }
}

function renderOcrJob(job) {
  const statusMap = { queued: "排队中", running: "运行中", done: "完成", error: "失败" };
  const etaText = formatDuration(job.eta_seconds);
  const elapsedText = formatDuration(job.elapsed_seconds || 0);
  const progressText = `${job.done || 0}/${job.total || 0}`;
  ocrMessageText.textContent = `${statusMap[job.status] || job.status}：${job.message || "等待更新"}；预计还需 ${etaText}`;
  setBar(ocrProgressBar, job.percent || 0);
  ocrReportPath.textContent = job.report_path || (job.status === "done" ? "报告已生成，可以下载" : "识别完成后可下载多 Sheet 报告");
  downloadOcrReport.disabled = job.status !== "done";
  downloadOcrCsvReport.disabled = job.status !== "done";
  renderOcrResults(job.result_rows || []);
  const info = job.output_info || {};
  if (Object.keys(info).length) {
    ocrOutputSummary.textContent = `OK ${info.ok_count || 0}，NG ${info.ng_count || 0}`;
    ocrOutputDirs.textContent = `文本：${info.text_dir || "-"} · 标注图：${info.annotated_dir || "-"} · NG图：${info.ng_dir || "-"} · 报告：${info.report_dir || "-"}`;
  }
  ocrLogs.textContent = job.logs && job.logs.length ? job.logs.join("\n") : "[ready] 等待 OCR 2.0 任务...";
  ocrLogs.scrollTop = ocrLogs.scrollHeight;
  recordCompletedJob(job, "ocr");
  addHistoryFromJob(job, "ocr");
  if (job.status === "running") {
    updateMascot("running", "CodexPet OCR 识别中", `${progressText}，${job.message || "正在识别文字"}`, `已用时 ${elapsedText} · 预计还需 ${etaText}`, {
      mood: "专注",
      task: `OCR ${progressText}`,
      energy: Math.max(12, Math.min(96, Number(job.percent || 0))),
    });
  } else if (job.status === "done") {
    updateMascot("done", "CodexPet 已完成", "OCR 2.0 识别完成，报告和结果表已就绪", `总耗时 ${elapsedText}`, { task: "OCR 报告就绪" });
  } else if (job.status === "error") {
    updateMascot("error", "CodexPet OCR 失败", job.message || "检查 OCR 依赖或日志提示", `已用时 ${elapsedText}`, { task: "等待修复" });
  }
}

function renderEnhanceJob(job) {
  const statusMap = { queued: "排队中", running: "运行中", done: "完成", error: "失败" };
  const etaText = formatDuration(job.eta_seconds);
  const elapsedText = formatDuration(job.elapsed_seconds || 0);
  const progressText = `${job.done || 0}/${job.total || 0}`;
  enhanceMessageText.textContent = `${statusMap[job.status] || job.status}：${job.message || "等待更新"}；预计还需 ${etaText}`;
  setBar(enhanceProgressBar, job.percent || 0);
  enhanceReportPath.textContent = job.report_path || (job.status === "done" ? "报告已生成，可以下载" : "增强完成后可下载前后对比报告");
  downloadEnhanceReport.disabled = job.status !== "done";
  downloadEnhanceCsvReport.disabled = job.status !== "done";
  renderEnhanceResults(job.result_rows || []);
  renderEnhanceOutputInfo(job.output_info);
  enhanceLogs.textContent = job.logs && job.logs.length ? job.logs.join("\n") : "[ready] 等待图像增强任务...";
  enhanceLogs.scrollTop = enhanceLogs.scrollHeight;
  recordCompletedJob(job, "image");
  addHistoryFromJob(job, "enhance");
  if (job.status === "running") {
    updateMascot("running", "CodexPet 增强图片中", `${progressText}，${job.message || "正在执行规则版自动增强"}`, `已用时 ${elapsedText} · 预计还需 ${etaText}`, {
      mood: "专注",
      task: `增强 ${progressText}`,
      energy: Math.max(12, Math.min(96, Number(job.percent || 0) || 35)),
    });
  } else if (job.status === "done") {
    updateMascot("done", "CodexPet 已完成", "图像增强完成，前后对比图和报告已经就绪", `总耗时 ${elapsedText}`, { task: "增强报告就绪" });
  } else if (job.status === "error") {
    updateMascot("error", "CodexPet 增强失败", job.message || "检查增强日志提示", `已用时 ${elapsedText}`, { task: "等待修复" });
  }
}

function renderJob(job) {
  const elapsedText = formatDuration(job.elapsed_seconds || 0);
  const statusMap = {
    queued: "排队中",
    running: "运行中",
    done: "完成",
    error: "失败",
  };
  statusText.textContent = statusMap[job.status] || job.status;
  countText.textContent = `${job.done || 0}/${job.total || 0}`;
  setProgress(job.percent || 0);
  messageText.textContent = `${job.message || "等待更新"} · 已用时 ${elapsedText}`;
  reportPath.textContent = job.report_path || (job.status === "done" ? "尚未生成，点击按钮后下载" : "检测完成后可按需生成并下载");
  downloadReport.disabled = job.status !== "done";
  downloadCsvReport.disabled = job.status !== "done";
  renderOutputInfo(job.output_info);
  renderResults(job.result_rows || []);
  logs.textContent = job.logs && job.logs.length ? job.logs.join("\n") : "[ready] 等待检测任务...";
  logs.scrollTop = logs.scrollHeight;
  recordCompletedJob(job, "image");
  addHistoryFromJob(job, "image");
  if (job.status === "running") {
    updateMascot("running", "CodexPet 检测图片中", `${job.done || 0}/${job.total || 0}，正在分析图片质量`, `已用时 ${elapsedText} · 预计还需 ${formatDuration(job.eta_seconds)}`, {
      task: `图片 ${job.done || 0}/${job.total || 0}`,
      energy: Math.max(12, Math.min(96, Number(job.percent || 0))),
    });
  } else if (job.status === "done") {
    updateMascot("done", "CodexPet 已完成", "评分已更新，可以下载 Excel 报告", `总耗时 ${elapsedText}`, { task: "报告就绪" });
  } else if (job.status === "error") {
    updateMascot("error", "CodexPet 发现问题", "看一下右侧日志，我会帮你定位", `已用时 ${elapsedText}`, { task: "等待修复" });
  }
}

async function pollJob(jobId) {
  let response;
  try {
    response = await fetch(`/api/jobs/${jobId}`);
  } catch (_error) {
    const error = new Error("进度连接中断，请检查网络或稍后重试。");
    error.retryable = true;
    throw error;
  }
  if (response.status === 404) {
    throw new Error("任务记录已丢失，服务可能已重启，请重新提交图片。");
  }
  if (!response.ok) {
    const error = new Error(`服务暂时无法返回进度（HTTP ${response.status}），请稍后重试。`);
    error.retryable = [502, 503, 504].includes(response.status);
    throw error;
  }
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "任务查询失败");
  }
  renderJob(payload.job);

  if (payload.job.status === "done" || payload.job.status === "error") {
    clearTimeout(pollTimer);
    pollTimer = null;
    submitButton.disabled = false;
  }
  return payload.job.status;
}

async function pollVideoJob(jobId) {
  const response = await fetch(`/api/jobs/${jobId}`);
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "视频任务查询失败");
  }
  renderVideoJob(payload.job);
  if (payload.job.status === "done" || payload.job.status === "error") {
    clearInterval(videoPollTimer);
    videoPollTimer = null;
    videoForm.querySelector('button[type="submit"]').disabled = false;
  }
}

async function pollQualityJob(jobId) {
  const response = await fetch(`/api/jobs/${jobId}`);
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "视频综合任务查询失败");
  }
  renderQualityJob(payload.job);
  if (payload.job.status === "done" || payload.job.status === "error") {
    clearInterval(qualityPollTimer);
    qualityPollTimer = null;
    videoQualityForm.querySelector('button[type="submit"]').disabled = false;
  }
}

async function pollOcrJob(jobId) {
  const response = await fetch(`/api/jobs/${jobId}`);
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "OCR 任务查询失败");
  }
  renderOcrJob(payload.job);
  if (payload.job.status === "done" || payload.job.status === "error") {
    clearInterval(ocrPollTimer);
    ocrPollTimer = null;
    ocrForm.querySelector('button[type="submit"]').disabled = false;
  }
}

async function pollEnhanceJob(jobId) {
  const response = await fetch(`/api/jobs/${jobId}`);
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "图像增强任务查询失败");
  }
  renderEnhanceJob(payload.job);
  if (payload.job.status === "done" || payload.job.status === "error") {
    clearInterval(enhancePollTimer);
    enhancePollTimer = null;
    enhanceForm.querySelector('button[type="submit"]').disabled = false;
  }
}

navItems.forEach((item) => {
  item.addEventListener("click", () => switchModule(item.dataset.module));
});
quickModuleButtons.forEach((item) => {
  item.addEventListener("click", () => switchModule(item.dataset.goModule));
});
inputModeSelects.forEach((select) => {
  select.addEventListener("change", () => syncInputMode(select));
  syncInputMode(select);
});

analysisMode.addEventListener("change", syncRegionState);
weightMode.addEventListener("change", syncWeights);
imagesInput.addEventListener("change", updateFilesHint);
imagesFolderInput.addEventListener("change", updateFilesHint);
videosInput.addEventListener("change", updateVideoFilesHint);
videosFolderInput.addEventListener("change", updateVideoFilesHint);
qualityVideosInput.addEventListener("change", updateQualityVideoFilesHint);
qualityVideosFolderInput.addEventListener("change", updateQualityVideoFilesHint);
ocrFilesInput.addEventListener("change", updateOcrFilesHint);
ocrFolderInput.addEventListener("change", updateOcrFilesHint);
enhanceImagesInput.addEventListener("change", updateEnhanceFilesHint);
enhanceImagesFolderInput.addEventListener("change", updateEnhanceFilesHint);
enhanceStepOptions.forEach((option) => option.addEventListener("change", syncEnhanceSteps));
syncEnhanceSteps();
saveSettingsButton.addEventListener("click", saveSettings);
clearHistoryButton.addEventListener("click", () => {
  detectionHistory = [];
  countedHistoryJobs = new Set();
  saveDetectionHistory();
  renderHistoryTables();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  submitButton.disabled = true;
  downloadReport.disabled = true;
  downloadCsvReport.disabled = true;
  currentJobId = "";
  statusText.textContent = "提交中";
  countText.textContent = "0/0";
  setProgress(0);
  messageText.textContent = "正在提交任务";
  updateMascot("running", "CodexPet 启动", "收到，我开始检查图片质量", "预计时间：正在计算", { task: "提交任务" });
  reportPath.textContent = "检测完成后可按需生成并下载";
  renderOutputInfo({});
  renderResults([]);
  logs.textContent = "[submit] 正在提交检测任务...";

  try {
    const response = await fetch("/api/detect", {
      method: "POST",
      body: new FormData(form),
    });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "提交失败");
    }

    if (pollTimer) {
      clearTimeout(pollTimer);
    }
    currentJobId = payload.job_id;
    let retryCount = 0;
    const poll = async () => {
      try {
        const status = await pollJob(payload.job_id);
        retryCount = 0;
        if (status !== "done" && status !== "error") {
          pollTimer = setTimeout(poll, 900);
        }
      } catch (error) {
        if (error.retryable && retryCount < 3) {
          const delay = 1000 * 2 ** retryCount;
          retryCount += 1;
          const message = `进度连接暂时中断，正在重连（${retryCount}/3），请稍候。`;
          statusText.textContent = "重连中";
          messageText.textContent = message;
          logs.textContent += `\n[retry] ${error.message} ${message}`;
          updateMascot("running", "CodexPet 正在重连", message, "", { task: "重连进度" });
          pollTimer = setTimeout(poll, delay);
          return;
        }
        clearTimeout(pollTimer);
        pollTimer = null;
        submitButton.disabled = false;
        const message = error.retryable ? `${error.message} 已重连 3 次，仍无法恢复进度连接。` : error.message;
        statusText.textContent = error.retryable ? "连接失败" : "失败";
        messageText.textContent = message;
        logs.textContent += `\n[error] ${message}`;
        updateMascot("error", "CodexPet 连接中断", message, "", { task: "等待重试" });
      }
    };
    await poll();
  } catch (error) {
    submitButton.disabled = false;
    statusText.textContent = "失败";
    messageText.textContent = error.message;
    logs.textContent += `\n[error] ${error.message}`;
  }
});

downloadReport.addEventListener("click", () => {
  if (!currentJobId) {
    return;
  }
  recordReportExport(currentJobId, "image-excel");
  window.location.href = `/api/jobs/${currentJobId}/report`;
  reportPath.textContent = "正在生成并下载 Excel 报告...";
});

downloadCsvReport.addEventListener("click", () => {
  if (!currentJobId) {
    return;
  }
  recordReportExport(currentJobId, "image-csv");
  window.location.href = `/api/jobs/${currentJobId}/csv`;
});

ocrForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = ocrForm.querySelector('button[type="submit"]');
  button.disabled = true;
  currentOcrJobId = "";
  setBar(ocrProgressBar, 0);
  ocrMessageText.textContent = "正在提交 OCR 任务";
  ocrReportPath.textContent = "识别完成后可下载多 Sheet 报告";
  ocrOutputSummary.textContent = "等待识别";
  ocrOutputDirs.textContent = "文本、标注图、NG 图和报告目录会在完成后显示";
  downloadOcrReport.disabled = true;
  downloadOcrCsvReport.disabled = true;
  renderOcrResults([]);
  ocrLogs.textContent = "[submit] 正在提交 OCR 2.0 任务...";
  updateMascot("running", "CodexPet 启动 OCR", "收到，我开始加载 OCR 2.0", "预计时间：正在计算", { task: "提交 OCR" });

  try {
    const response = await fetch("/api/ocr/detect", {
      method: "POST",
      body: new FormData(ocrForm),
    });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "OCR 任务提交失败");
    }
    if (ocrPollTimer) {
      clearInterval(ocrPollTimer);
    }
    currentOcrJobId = payload.job_id;
    await pollOcrJob(payload.job_id);
    ocrPollTimer = setInterval(() => pollOcrJob(payload.job_id).catch((error) => {
      clearInterval(ocrPollTimer);
      ocrPollTimer = null;
      button.disabled = false;
      ocrMessageText.textContent = error.message;
    }), 900);
  } catch (error) {
    button.disabled = false;
    ocrMessageText.textContent = error.message;
    ocrLogs.textContent += `\n[error] ${error.message}`;
  }
});

downloadOcrReport.addEventListener("click", () => {
  if (!currentOcrJobId) {
    return;
  }
  recordReportExport(currentOcrJobId, "ocr-excel");
  window.location.href = `/api/ocr/jobs/${currentOcrJobId}/report`;
  ocrReportPath.textContent = "正在下载 OCR 2.0 Excel 报告...";
});

downloadOcrCsvReport.addEventListener("click", () => {
  if (!currentOcrJobId) {
    return;
  }
  recordReportExport(currentOcrJobId, "ocr-csv");
  window.location.href = `/api/jobs/${currentOcrJobId}/csv`;
});

videoForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = videoForm.querySelector('button[type="submit"]');
  button.disabled = true;
  currentVideoJobId = "";
  setBar(videoProgressBar, 0);
  videoMessageText.textContent = "正在提交视频任务";
  updateMascot("running", "CodexPet 启动", "收到，我开始读取视频信息", "预计时间：正在计算", { task: "提交任务" });
  videoReportPath.textContent = "检测完成后可按需生成并下载";
  downloadVideoReport.disabled = true;
  downloadVideoCsvReport.disabled = true;
  renderVideoResults([]);
  videoLogs.textContent = "[submit] 正在提交视频检测任务...";

  try {
    const response = await fetch("/api/video/detect", {
      method: "POST",
      body: new FormData(videoForm),
    });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "视频任务提交失败");
    }
    if (videoPollTimer) {
      clearInterval(videoPollTimer);
    }
    currentVideoJobId = payload.job_id;
    await pollVideoJob(payload.job_id);
    videoPollTimer = setInterval(() => pollVideoJob(payload.job_id).catch((error) => {
      clearInterval(videoPollTimer);
      videoPollTimer = null;
      button.disabled = false;
      videoMessageText.textContent = error.message;
    }), 900);
  } catch (error) {
    button.disabled = false;
    videoMessageText.textContent = error.message;
    videoLogs.textContent += `\n[error] ${error.message}`;
  }
});

downloadVideoReport.addEventListener("click", () => {
  if (!currentVideoJobId) {
    return;
  }
  recordReportExport(currentVideoJobId, "video-excel");
  window.location.href = `/api/video/jobs/${currentVideoJobId}/report`;
  videoReportPath.textContent = "正在生成并下载视频 Excel 报告...";
});

downloadVideoCsvReport.addEventListener("click", () => {
  if (!currentVideoJobId) {
    return;
  }
  recordReportExport(currentVideoJobId, "video-csv");
  window.location.href = `/api/jobs/${currentVideoJobId}/csv`;
});

videoQualityForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = videoQualityForm.querySelector('button[type="submit"]');
  button.disabled = true;
  currentQualityJobId = "";
  setBar(qualityProgressBar, 0);
  qualityMessageText.textContent = "正在提交视频综合评分任务";
  updateMascot("running", "CodexPet 启动", "收到，我开始综合检测视频质量", "预计时间：正在计算", { task: "提交综合任务" });
  qualityReportPath.textContent = "检测完成后可按需生成并下载";
  downloadQualityReport.disabled = true;
  downloadQualityCsvReport.disabled = true;
  renderQualityResults([]);
  qualityLogs.textContent = "[submit] 正在提交视频综合评分任务...";

  try {
    const response = await fetch("/api/video-quality/detect", {
      method: "POST",
      body: new FormData(videoQualityForm),
    });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "视频综合任务提交失败");
    }
    if (qualityPollTimer) {
      clearInterval(qualityPollTimer);
    }
    currentQualityJobId = payload.job_id;
    await pollQualityJob(payload.job_id);
    qualityPollTimer = setInterval(() => pollQualityJob(payload.job_id).catch((error) => {
      clearInterval(qualityPollTimer);
      qualityPollTimer = null;
      button.disabled = false;
      qualityMessageText.textContent = error.message;
    }), 900);
  } catch (error) {
    button.disabled = false;
    qualityMessageText.textContent = error.message;
    qualityLogs.textContent += `\n[error] ${error.message}`;
  }
});

downloadQualityReport.addEventListener("click", () => {
  if (!currentQualityJobId) {
    return;
  }
  recordReportExport(currentQualityJobId, "quality-excel");
  window.location.href = `/api/video-quality/jobs/${currentQualityJobId}/report`;
  qualityReportPath.textContent = "正在生成并下载视频综合 Excel 报告...";
});

downloadQualityCsvReport.addEventListener("click", () => {
  if (!currentQualityJobId) {
    return;
  }
  recordReportExport(currentQualityJobId, "quality-csv");
  window.location.href = `/api/jobs/${currentQualityJobId}/csv`;
});

enhanceForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  syncEnhanceSteps();
  const button = enhanceForm.querySelector('button[type="submit"]');
  button.disabled = true;
  currentEnhanceJobId = "";
  setBar(enhanceProgressBar, 0);
  enhanceMessageText.textContent = "正在提交图像增强任务";
  enhanceReportPath.textContent = "增强完成后可下载前后对比报告";
  enhanceOutputSummary.textContent = "等待增强";
  enhanceOutputDirs.textContent = "增强图、对比图、报告和日志目录会在完成后显示";
  downloadEnhanceReport.disabled = true;
  downloadEnhanceCsvReport.disabled = true;
  renderEnhanceResults([]);
  enhanceLogs.textContent = "[submit] 正在提交图像增强任务...";
  updateMascot("running", "CodexPet 启动增强", "收到，我开始检测质量问题并选择增强策略", "预计时间：正在计算", { task: "提交增强" });

  try {
    const response = await fetch("/api/enhance/detect", {
      method: "POST",
      body: new FormData(enhanceForm),
    });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "图像增强任务提交失败");
    }
    if (enhancePollTimer) {
      clearInterval(enhancePollTimer);
    }
    currentEnhanceJobId = payload.job_id;
    await pollEnhanceJob(payload.job_id);
    enhancePollTimer = setInterval(() => pollEnhanceJob(payload.job_id).catch((error) => {
      clearInterval(enhancePollTimer);
      enhancePollTimer = null;
      button.disabled = false;
      enhanceMessageText.textContent = error.message;
    }), 900);
  } catch (error) {
    button.disabled = false;
    enhanceMessageText.textContent = error.message;
    enhanceLogs.textContent += `\n[error] ${error.message}`;
  }
});

downloadEnhanceReport.addEventListener("click", () => {
  if (!currentEnhanceJobId) {
    return;
  }
  recordReportExport(currentEnhanceJobId, "enhance-excel");
  window.location.href = `/api/enhance/jobs/${currentEnhanceJobId}/report`;
  enhanceReportPath.textContent = "正在下载图像增强 Excel 报告...";
});

downloadEnhanceCsvReport.addEventListener("click", () => {
  if (!currentEnhanceJobId) {
    return;
  }
  recordReportExport(currentEnhanceJobId, "enhance-csv");
  window.location.href = `/api/jobs/${currentEnhanceJobId}/csv`;
});

clearCacheButton.addEventListener("click", async () => {
  const confirmed = window.confirm("确定清理 web_uploads / web_results / web_video_uploads / web_video_results 吗？正在运行的网页检测缓存也会被清掉。");
  if (!confirmed) {
    return;
  }
  clearCacheButton.disabled = true;
  cacheStatus.textContent = "正在清理缓存...";
  try {
    const response = await fetch("/api/clear-uploads", { method: "POST" });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "清理失败");
    }
    cacheStatus.textContent = `已清理 ${payload.cleared.length} 个缓存目录`;
    updateMascot("done", "CodexPet 已清理", "网页上传缓存和检测缓存已清空", "", { task: "缓存已清理" });
  } catch (error) {
    cacheStatus.textContent = error.message;
    updateMascot("error", "CodexPet 清理失败", error.message, "", { task: "清理失败" });
  } finally {
    clearCacheButton.disabled = false;
  }
});

syncRegionState();
syncWeights();
normalizeModuleAreas();
renderDashboard();
renderHistoryTables();
initSettings();
initMascotDrag();
switchModule("home");
