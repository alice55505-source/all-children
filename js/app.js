(function () {
  "use strict";

  var DEFAULT_CONGREGATIONS = [
    "斗六", "古坑", "林內", "西螺", "莿桐", "斗南", "虎尾", "土庫",
    "崙背", "褒忠", "二崙", "麥寮", "北港", "口湖", "嘉義市", "中埔",
    "竹崎", "番路", "民雄", "溪口", "大林", "新港", "六腳", "朴子",
    "布袋", "鹿草", "太保", "水上"
  ];

  var CONGREGATIONS_STORAGE_KEY = "congregationList_v1";

  function loadCongregations() {
    try {
      var raw = localStorage.getItem(CONGREGATIONS_STORAGE_KEY);
      if (raw == null) return DEFAULT_CONGREGATIONS.slice();
      var parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : DEFAULT_CONGREGATIONS.slice();
    } catch (e) {
      return DEFAULT_CONGREGATIONS.slice();
    }
  }

  function saveCongregations(list) {
    try {
      localStorage.setItem(CONGREGATIONS_STORAGE_KEY, JSON.stringify(list));
    } catch (e) {
      /* storage unavailable, ignore */
    }
  }

  var CONGREGATIONS = loadCongregations();

  var TOTAL_LABEL = "合計";

  function findHeaderIndex(header, name) {
    for (var i = 0; i < header.length; i++) {
      if (header[i] === name) return i;
    }
    return -1;
  }

  function getSheet(workbook, preferredName, fallbackIndex) {
    var sheet = workbook.Sheets[preferredName];
    if (!sheet) {
      var fallbackName = workbook.SheetNames[fallbackIndex];
      sheet = fallbackName ? workbook.Sheets[fallbackName] : null;
    }
    return sheet;
  }

  function formatNum(n) {
    return Math.round(n * 10) / 10;
  }

  // ---- 兒童：週報矩陣 + 摘要 ----

  var CHILDREN_WEEKLY_SHEET_NAME = "週報矩陣";
  var CHILDREN_SUMMARY_SHEET_NAME = "摘要";
  var CHILDREN_LIFE_KPI_LABEL = "期間有召會生活";

  function parseChildrenWeeklyMatrix(sheet) {
    if (!sheet) {
      throw new Error("找不到「" + CHILDREN_WEEKLY_SHEET_NAME + "」分頁，請確認上傳的檔案格式");
    }

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!rows.length) {
      throw new Error("「" + CHILDREN_WEEKLY_SHEET_NAME + "」分頁是空的");
    }

    var header = rows[0].map(function (cell) {
      return cell == null ? "" : String(cell).trim();
    });

    var idx = {
      區排: findHeaderIndex(header, "區排"),
      主日: findHeaderIndex(header, "主日"),
      小排: findHeaderIndex(header, "小排")
    };

    if (idx.區排 === -1) {
      throw new Error("在「" + CHILDREN_WEEKLY_SHEET_NAME + "」分頁找不到欄位：區排");
    }

    var sums = { 主日: 0, 小排: 0 };
    var weeks = 0;

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row) continue;
      var label = row[idx.區排];
      if (label != null && String(label).trim() === TOTAL_LABEL) {
        weeks++;
        sums.主日 += idx.主日 === -1 ? 0 : (Number(row[idx.主日]) || 0);
        sums.小排 += idx.小排 === -1 ? 0 : (Number(row[idx.小排]) || 0);
      }
    }

    if (weeks === 0) {
      throw new Error("在「" + CHILDREN_WEEKLY_SHEET_NAME + "」分頁中找不到「" + TOTAL_LABEL + "」列");
    }

    return {
      weeks: weeks,
      avgSunday: sums.主日 / weeks,
      avgGroup: sums.小排 / weeks
    };
  }

  function parseChildrenLifeKpi(sheet) {
    if (!sheet) return 0;

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row) continue;
      var label = row[0];
      if (label != null && String(label).trim() === CHILDREN_LIFE_KPI_LABEL) {
        return Number(row[1]) || 0;
      }
    }

    return 0;
  }

  function parseChildrenWorkbook(workbook) {
    var weeklySheet = getSheet(workbook, CHILDREN_WEEKLY_SHEET_NAME, 2);
    var summarySheet = getSheet(workbook, CHILDREN_SUMMARY_SHEET_NAME, 1);

    var weekly = parseChildrenWeeklyMatrix(weeklySheet);
    var lifeCount = parseChildrenLifeKpi(summarySheet);

    return {
      weeks: weekly.weeks,
      avgSunday: weekly.avgSunday,
      avgGroup: weekly.avgGroup,
      lifeCount: lifeCount
    };
  }

  // ---- 青職：週報網格 ----

  var YOUTH_SHEET_NAME = "週報網格";
  var YOUTH_AGE_LABEL = "青職";
  var YOUTH_WEEK_PATTERN = /^\d+W\d+$/;
  var YOUTH_TARGET_CATEGORIES = ["主日", "家聚會出訪", "家聚會受訪", "小排", "生命讀經"];

  function fillForward(row, startCol) {
    var out = [];
    var last = null;
    var len = row ? row.length : 0;
    for (var c = 0; c < len; c++) {
      var v = row[c];
      if (v != null && String(v).trim() !== "") last = String(v).trim();
      out[c] = c >= startCol ? last : null;
    }
    return out;
  }

  function parseYouthWorkbook(workbook) {
    var sheet = getSheet(workbook, YOUTH_SHEET_NAME, 0);
    if (!sheet) {
      throw new Error("找不到「" + YOUTH_SHEET_NAME + "」分頁，請確認上傳的檔案格式");
    }

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!rows.length) {
      throw new Error("「" + YOUTH_SHEET_NAME + "」分頁是空的");
    }

    var headerRowIdx = -1;
    var groupColIdx = -1;
    for (var r = 0; r < rows.length && headerRowIdx === -1; r++) {
      var row = rows[r];
      if (!row) continue;
      for (var c = 0; c < row.length - 1; c++) {
        if (String(row[c] || "").trim() === "區" && String(row[c + 1] || "").trim() === "排") {
          headerRowIdx = r;
          groupColIdx = c + 1;
          break;
        }
      }
    }

    if (headerRowIdx === -1) {
      throw new Error("在「" + YOUTH_SHEET_NAME + "」分頁找不到「區」「排」表頭");
    }

    var weekRow = rows[headerRowIdx];
    var categoryRow = rows[headerRowIdx + 1];
    var ageRow = rows[headerRowIdx + 2];

    if (!categoryRow || !ageRow) {
      throw new Error("在「" + YOUTH_SHEET_NAME + "」分頁找不到分類或身份列");
    }

    var weekPerCol = fillForward(weekRow, groupColIdx + 1);
    var catPerCol = fillForward(categoryRow, groupColIdx + 1);

    var matches = [];
    var width = Math.max(weekRow.length, categoryRow.length, ageRow.length);
    for (var col = groupColIdx + 1; col < width; col++) {
      var wk = weekPerCol[col];
      var cat = catPerCol[col];
      var age = ageRow[col] != null ? String(ageRow[col]).trim() : null;
      if (wk && YOUTH_WEEK_PATTERN.test(wk) && age === YOUTH_AGE_LABEL && YOUTH_TARGET_CATEGORIES.indexOf(cat) !== -1) {
        matches.push({ week: wk, cat: cat, col: col });
      }
    }

    if (!matches.length) {
      throw new Error("在「" + YOUTH_SHEET_NAME + "」分頁找不到「" + YOUTH_AGE_LABEL + "」欄位資料");
    }

    var dataStart = headerRowIdx + 3;
    var aggRow = null;
    var dataRows = [];
    for (var i = dataStart; i < rows.length; i++) {
      var dr = rows[i];
      if (!dr) continue;
      var groupLabel = dr[groupColIdx];
      if (groupLabel == null || String(groupLabel).trim() === "") continue;
      if (String(groupLabel).trim() === TOTAL_LABEL) {
        aggRow = dr;
      } else {
        dataRows.push(dr);
      }
    }

    if (!aggRow && !dataRows.length) {
      throw new Error("在「" + YOUTH_SHEET_NAME + "」分頁找不到任何資料列");
    }

    var sums = {};
    var weeksSet = {};
    var weeksCount = 0;

    matches.forEach(function (m) {
      if (!weeksSet[m.week]) {
        weeksSet[m.week] = true;
        weeksCount++;
      }
      var value;
      if (aggRow) {
        value = Number(aggRow[m.col]) || 0;
      } else {
        value = dataRows.reduce(function (acc, dr) {
          return acc + (Number(dr[m.col]) || 0);
        }, 0);
      }
      sums[m.cat] = (sums[m.cat] || 0) + value;
    });

    var get = function (key) { return sums[key] || 0; };

    return {
      weeks: weeksCount,
      avgSunday: get("主日") / weeksCount,
      avgFamily: (get("家聚會出訪") + get("家聚會受訪")) / weeksCount,
      avgGroup: get("小排") / weeksCount,
      avgLifeReading: get("生命讀經") / weeksCount
    };
  }

  // ---- 共用模組工廠 ----

  function createStatsModule(cfg) {
    var prefix = cfg.prefix;
    var storageKey = cfg.storageKey;
    var parseFn = cfg.parseFn;
    var metrics = cfg.metrics;

    var state = loadState();
    var pickedFile = null;

    function id(name) { return document.getElementById(prefix + "-" + name); }

    var els = {
      select: id("congregation-select"),
      dropzone: id("dropzone"),
      fileInput: id("file-input"),
      filePicked: id("file-picked"),
      parseBtn: id("parse-btn"),
      resetAllBtn: id("reset-all-btn"),
      statusMsg: id("status-msg"),
      progressCount: id("progress-count"),
      progressTotal: id("progress-total"),
      progressBar: id("progress-bar"),
      chipGrid: id("chip-grid"),
      resultsTheadRow: id("results-thead-row"),
      resultsTbody: id("results-tbody"),
      resultsEmpty: id("results-empty"),
      summaryGrid: id("summary-grid"),
      summaryWarn: id("summary-warn"),
      copySummaryBtn: id("copy-summary-btn")
    };

    function loadState() {
      try {
        var raw = localStorage.getItem(storageKey);
        return raw ? JSON.parse(raw) : {};
      } catch (e) {
        return {};
      }
    }

    function saveState() {
      try {
        localStorage.setItem(storageKey, JSON.stringify(state));
      } catch (e) {
        /* storage unavailable, ignore */
      }
    }

    function initSelect() {
      els.select.innerHTML = "";
      if (!CONGREGATIONS.length) {
        var placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "尚未設定召會，請先到設定新增";
        els.select.appendChild(placeholder);
        els.parseBtn.disabled = true;
        return;
      }
      CONGREGATIONS.forEach(function (name) {
        var opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        els.select.appendChild(opt);
      });
    }

    function initTableHeader() {
      var html = "<th>召會</th><th>週數</th>";
      metrics.forEach(function (m) {
        html += "<th>" + m.label + "</th>";
      });
      html += "<th>操作</th>";
      els.resultsTheadRow.innerHTML = html;
    }

    function initSummaryGrid() {
      var html = "";
      metrics.forEach(function (m) {
        html +=
          '<div class="stat-tile">' +
          '<div class="stat-label">' + m.totalLabel + "</div>" +
          '<div class="stat-value" id="' + prefix + "-total-" + m.key + '">0</div>' +
          "</div>";
      });
      els.summaryGrid.innerHTML = html;
    }

    function showStatus(message, type) {
      els.statusMsg.textContent = message;
      els.statusMsg.className = "status-msg show " + type;
    }

    function clearStatus() {
      els.statusMsg.className = "status-msg";
    }

    function handleFileSelected(file) {
      pickedFile = file;
      els.filePicked.textContent = file ? "已選擇檔案：" + file.name : "";
    }

    function setupDropzone() {
      var dz = els.dropzone;
      ["dragenter", "dragover"].forEach(function (evt) {
        dz.addEventListener(evt, function (e) {
          e.preventDefault();
          dz.classList.add("drag-over");
        });
      });
      ["dragleave", "drop"].forEach(function (evt) {
        dz.addEventListener(evt, function (e) {
          e.preventDefault();
          dz.classList.remove("drag-over");
        });
      });
      dz.addEventListener("drop", function (e) {
        var files = e.dataTransfer.files;
        if (files && files.length) {
          handleFileSelected(files[0]);
        }
      });
      els.fileInput.addEventListener("change", function (e) {
        if (e.target.files && e.target.files.length) {
          handleFileSelected(e.target.files[0]);
        }
      });
    }

    function onParseClick() {
      clearStatus();
      var congregation = els.select.value;
      if (!congregation) {
        showStatus("請先到設定新增召會", "error");
        return;
      }
      if (!pickedFile) {
        showStatus("請先選擇一個 .xlsx 檔案", "error");
        return;
      }
      els.parseBtn.disabled = true;

      pickedFile.arrayBuffer().then(function (buf) {
        var workbook = XLSX.read(buf, { type: "array" });
        var result = parseFn(workbook);
        var isUpdate = Object.prototype.hasOwnProperty.call(state, congregation);

        if (isUpdate) {
          var ok = window.confirm("「" + congregation + "」已經上傳過資料，是否要用這個新檔案覆蓋？");
          if (!ok) {
            els.parseBtn.disabled = false;
            return;
          }
        }

        var entry = { fileName: pickedFile.name, updatedAt: new Date().toISOString() };
        metrics.forEach(function (m) { entry[m.key] = result[m.key]; });
        entry.weeks = result.weeks;
        state[congregation] = entry;
        saveState();
        renderAll();

        var parts = metrics.map(function (m) {
          return m.label + " " + formatNum(result[m.key]);
        });
        showStatus(
          "已加入「" + congregation + "」：共 " + result.weeks + " 週，" + parts.join("、"),
          "ok"
        );

        pickedFile = null;
        els.fileInput.value = "";
        els.filePicked.textContent = "";
        els.parseBtn.disabled = false;
      }).catch(function (err) {
        showStatus("解析失敗：" + err.message, "error");
        els.parseBtn.disabled = false;
      });
    }

    function onResetAllClick() {
      if (!Object.keys(state).length) return;
      var ok = window.confirm("確定要清除所有已上傳的召會資料嗎？此操作無法復原。");
      if (!ok) return;
      state = {};
      saveState();
      renderAll();
      clearStatus();
    }

    function removeCongregation(name) {
      var ok = window.confirm("確定要移除「" + name + "」的資料嗎？");
      if (!ok) return;
      delete state[name];
      saveState();
      renderAll();
    }

    function renderChecklist() {
      var uploadedCount = 0;
      els.chipGrid.innerHTML = "";
      CONGREGATIONS.forEach(function (name) {
        var chip = document.createElement("span");
        var done = Object.prototype.hasOwnProperty.call(state, name);
        if (done) uploadedCount++;
        chip.className = "chip" + (done ? " done" : "");
        chip.textContent = name;
        els.chipGrid.appendChild(chip);
      });
      els.progressCount.textContent = uploadedCount;
      els.progressTotal.textContent = CONGREGATIONS.length;
      els.progressBar.style.width = (uploadedCount / CONGREGATIONS.length * 100) + "%";
      return uploadedCount;
    }

    function renderTable() {
      els.resultsTbody.innerHTML = "";
      var names = CONGREGATIONS.filter(function (name) {
        return Object.prototype.hasOwnProperty.call(state, name);
      });

      els.resultsEmpty.style.display = names.length ? "none" : "block";

      names.forEach(function (name) {
        var d = state[name];
        var html = "<td>" + name + "</td><td>" + d.weeks + "</td>";
        metrics.forEach(function (m) {
          html += "<td>" + formatNum(d[m.key]) + "</td>";
        });
        html += "<td><button class=\"btn-danger-ghost\" data-remove=\"" + name + "\">移除</button></td>";
        var tr = document.createElement("tr");
        tr.innerHTML = html;
        els.resultsTbody.appendChild(tr);
      });

      els.resultsTbody.querySelectorAll("[data-remove]").forEach(function (btn) {
        btn.addEventListener("click", function () {
          removeCongregation(btn.getAttribute("data-remove"));
        });
      });
    }

    function computeTotals() {
      var totals = {};
      metrics.forEach(function (m) { totals[m.key] = 0; });
      Object.keys(state).forEach(function (name) {
        var d = state[name];
        metrics.forEach(function (m) { totals[m.key] += d[m.key] || 0; });
      });
      return totals;
    }

    function renderSummary() {
      var totals = computeTotals();
      metrics.forEach(function (m) {
        var el = document.getElementById(prefix + "-total-" + m.key);
        if (el) el.textContent = formatNum(totals[m.key]);
      });

      var missing = CONGREGATIONS.filter(function (name) {
        return !Object.prototype.hasOwnProperty.call(state, name);
      });

      if (missing.length) {
        els.summaryWarn.style.display = "block";
        els.summaryWarn.textContent =
          "尚有 " + missing.length + " 個召會未上傳（" + missing.join("、") + "），以上總計僅計入已上傳的召會。";
      } else {
        els.summaryWarn.style.display = "none";
      }
    }

    function onCopySummaryClick() {
      var uploadedNames = CONGREGATIONS.filter(function (name) {
        return Object.prototype.hasOwnProperty.call(state, name);
      });
      var totals = computeTotals();

      var lines = [];
      lines.push(cfg.summaryTitle + "（已上傳 " + uploadedNames.length + " / " + CONGREGATIONS.length + " 個召會）");
      metrics.forEach(function (m) {
        lines.push(m.totalLabel + "：" + formatNum(totals[m.key]));
      });

      var text = lines.join("\n");

      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
          showStatus("已複製總計文字", "ok");
        }).catch(function () {
          window.prompt("複製以下文字：", text);
        });
      } else {
        window.prompt("複製以下文字：", text);
      }
    }

    function renderAll() {
      renderChecklist();
      renderTable();
      renderSummary();
    }

    function init() {
      initSelect();
      initTableHeader();
      initSummaryGrid();
      setupDropzone();
      els.parseBtn.addEventListener("click", onParseClick);
      els.resetAllBtn.addEventListener("click", onResetAllClick);
      els.copySummaryBtn.addEventListener("click", onCopySummaryClick);
      renderAll();
    }

    init();
  }

  function initTabs() {
    var buttons = document.querySelectorAll("[data-tab-target]");
    buttons.forEach(function (btn) {
      btn.addEventListener("click", function () {
        var target = btn.getAttribute("data-tab-target");
        buttons.forEach(function (b) { b.classList.toggle("active", b === btn); });
        document.querySelectorAll("[data-tab-panel]").forEach(function (panel) {
          panel.style.display = panel.getAttribute("data-tab-panel") === target ? "" : "none";
        });
      });
    });
  }

  function parseCongregationLines(text) {
    var seen = {};
    var list = [];
    text.split("\n").forEach(function (line) {
      var name = line.trim();
      if (!name || seen[name]) return;
      seen[name] = true;
      list.push(name);
    });
    return list;
  }

  function initSettingsModal() {
    var openBtn = document.getElementById("open-settings-btn");
    var backdrop = document.getElementById("settings-backdrop");
    var textarea = document.getElementById("settings-textarea");
    var countEl = document.getElementById("settings-count");
    var clearBtn = document.getElementById("settings-clear-btn");
    var defaultBtn = document.getElementById("settings-default-btn");
    var cancelBtn = document.getElementById("settings-cancel-btn");
    var saveBtn = document.getElementById("settings-save-btn");

    function updateCount() {
      var count = parseCongregationLines(textarea.value).length;
      countEl.textContent = "目前共 " + count + " 個召會";
    }

    function open() {
      textarea.value = loadCongregations().join("\n");
      updateCount();
      backdrop.style.display = "flex";
    }

    function close() {
      backdrop.style.display = "none";
    }

    openBtn.addEventListener("click", open);
    cancelBtn.addEventListener("click", close);
    backdrop.addEventListener("click", function (e) {
      if (e.target === backdrop) close();
    });
    textarea.addEventListener("input", updateCount);

    clearBtn.addEventListener("click", function () {
      textarea.value = "";
      updateCount();
    });

    defaultBtn.addEventListener("click", function () {
      textarea.value = DEFAULT_CONGREGATIONS.join("\n");
      updateCount();
    });

    saveBtn.addEventListener("click", function () {
      var list = parseCongregationLines(textarea.value);
      saveCongregations(list);
      window.location.reload();
    });
  }

  function init() {
    initTabs();
    initSettingsModal();

    createStatsModule({
      prefix: "children",
      storageKey: "childrenStatsGenerator_v1",
      parseFn: parseChildrenWorkbook,
      summaryTitle: "全台兒童統計",
      metrics: [
        { key: "avgSunday", label: "主日（週平均）", totalLabel: "主日總計" },
        { key: "lifeCount", label: "召會生活（期間人數）", totalLabel: "召會生活總計" },
        { key: "avgGroup", label: "小排（週平均）", totalLabel: "小排總計" }
      ]
    });

    createStatsModule({
      prefix: "youth",
      storageKey: "youthStatsGenerator_v1",
      parseFn: parseYouthWorkbook,
      summaryTitle: "全台青職統計",
      metrics: [
        { key: "avgSunday", label: "主日（週平均）", totalLabel: "主日總計" },
        { key: "avgFamily", label: "家聚會（出訪+受訪，週平均）", totalLabel: "家聚會總計" },
        { key: "avgGroup", label: "小排（週平均）", totalLabel: "小排總計" },
        { key: "avgLifeReading", label: "生命讀經（週平均）", totalLabel: "生命讀經總計" }
      ]
    });

    if ("serviceWorker" in navigator) {
      window.addEventListener("load", function () {
        navigator.serviceWorker.register("sw.js").catch(function () {
          /* offline support unavailable, app still works online */
        });
      });
    }
  }

  init();
})();
