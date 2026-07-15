(function () {
  "use strict";

  var DEFAULT_GROUPS = [
    { region: "雲東區", members: ["斗六", "古坑", "林內", "西螺", "莿桐", "斗南"] },
    { region: "雲西區", members: ["虎尾", "土庫", "崙背", "褒忠", "二崙", "麥寮", "北港", "口湖"] },
    { region: "嘉義區", members: ["嘉義市（梅山）", "中埔", "竹崎", "番路"] },
    { region: "民雄區", members: ["民雄", "溪口", "大林", "新港", "(六腳)"] },
    { region: "朴子區", members: ["朴子", "布袋", "鹿草", "太保", "水上"] }
  ];
  var DEFAULT_REGION_NAME = "未分區";

  var GROUPS_STORAGE_KEY = "congregationGroups_v1";
  var LEGACY_FLAT_STORAGE_KEY = "congregationList_v1";

  function normalizeGroups(raw) {
    if (!Array.isArray(raw)) return null;
    var seen = {};
    var groups = [];
    raw.forEach(function (g) {
      if (!g || typeof g.region !== "string" || !Array.isArray(g.members)) return;
      var region = g.region.trim();
      if (!region) return;
      var members = [];
      g.members.forEach(function (name) {
        var n = typeof name === "string" ? name.trim() : "";
        if (!n || seen[n]) return;
        seen[n] = true;
        members.push(n);
      });
      groups.push({ region: region, members: members });
    });
    return groups;
  }

  function loadCongregationGroups() {
    try {
      var raw = localStorage.getItem(GROUPS_STORAGE_KEY);
      if (raw != null) {
        var parsed = normalizeGroups(JSON.parse(raw));
        if (parsed) return parsed;
      }
      var legacyRaw = localStorage.getItem(LEGACY_FLAT_STORAGE_KEY);
      if (legacyRaw != null) {
        var legacyParsed = JSON.parse(legacyRaw);
        if (Array.isArray(legacyParsed)) {
          return normalizeGroups([{ region: DEFAULT_REGION_NAME, members: legacyParsed }]) || DEFAULT_GROUPS;
        }
      }
      return DEFAULT_GROUPS.map(function (g) { return { region: g.region, members: g.members.slice() }; });
    } catch (e) {
      return DEFAULT_GROUPS.map(function (g) { return { region: g.region, members: g.members.slice() }; });
    }
  }

  function saveCongregationGroups(groups) {
    try {
      localStorage.setItem(GROUPS_STORAGE_KEY, JSON.stringify(groups));
    } catch (e) {
      /* storage unavailable, ignore */
    }
  }

  function flattenGroups(groups) {
    var flat = [];
    groups.forEach(function (g) {
      g.members.forEach(function (name) { flat.push(name); });
    });
    return flat;
  }

  var CONGREGATION_GROUPS = loadCongregationGroups();
  var CONGREGATIONS = flattenGroups(CONGREGATION_GROUPS);

  var TOTAL_LABEL = "合計";

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

  // ---- 週報網格：一次掃描，兒童／青職共用 ----

  var GRID_SHEET_NAME = "週報網格";
  var GRID_WEEK_PATTERN = /^\d+W\d+$/;

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

  function scanWeeklyGrid(workbook) {
    var sheet = getSheet(workbook, GRID_SHEET_NAME, 0);
    if (!sheet) {
      throw new Error("找不到「" + GRID_SHEET_NAME + "」分頁，請確認上傳的檔案格式");
    }

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!rows.length) {
      throw new Error("「" + GRID_SHEET_NAME + "」分頁是空的");
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
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到「區」「排」表頭");
    }

    var weekRow = rows[headerRowIdx];
    var categoryRow = rows[headerRowIdx + 1];
    var ageRow = rows[headerRowIdx + 2];

    if (!categoryRow || !ageRow) {
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到分類或身份列");
    }

    var weekPerCol = fillForward(weekRow, groupColIdx + 1);
    var catPerCol = fillForward(categoryRow, groupColIdx + 1);

    var matches = [];
    var width = Math.max(weekRow.length, categoryRow.length, ageRow.length);
    for (var col = groupColIdx + 1; col < width; col++) {
      var wk = weekPerCol[col];
      if (!wk || !GRID_WEEK_PATTERN.test(wk)) continue;
      var cat = catPerCol[col];
      if (!cat) continue;
      var age = ageRow[col] != null ? String(ageRow[col]).trim() : "";
      matches.push({ week: wk, key: cat + "|" + age, col: col });
    }

    if (!matches.length) {
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到任何週別欄位資料");
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
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到任何資料列");
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
      sums[m.key] = (sums[m.key] || 0) + value;
    });

    return { weeks: weeksCount, sums: sums };
  }

  function deriveChildrenResult(scan) {
    var get = function (cat, age) { return scan.sums[cat + "|" + age] || 0; };
    var weeks = scan.weeks;
    return {
      weeks: weeks,
      avgSunday: (get("兒童", "小計") + get("主日", "國小")) / weeks,
      avgGroup: (get("小排", "學齡前") + get("小排", "國小")) / weeks,
      avgLife: (get("召會生活", "學齡前") + get("召會生活", "國小")) / weeks
    };
  }

  function deriveYouthResult(scan) {
    var get = function (cat, age) { return scan.sums[cat + "|" + age] || 0; };
    var weeks = scan.weeks;
    return {
      weeks: weeks,
      avgSunday: get("主日", "青職") / weeks,
      avgFamily: (get("家聚會出訪", "青職") + get("家聚會受訪", "青職")) / weeks,
      avgGroup: get("小排", "青職") / weeks,
      avgLifeReading: get("生命讀經", "青職") / weeks
    };
  }

  // ---- 共用區塊（進度／表格／總計）算繪工廠 ----

  var STORAGE_KEY = "statsGeneratorData_v1";

  function createSection(cfg) {
    var prefix = cfg.prefix;
    var metrics = cfg.metrics;
    var extract = cfg.extract;

    function id(name) { return document.getElementById(prefix + "-" + name); }

    var els = {
      progressCount: id("progress-count"),
      progressTotal: id("progress-total"),
      progressBar: id("progress-bar"),
      chipGrid: id("chip-grid"),
      resultsTheadRow: id("results-thead-row"),
      resultsTbody: id("results-tbody"),
      resultsEmpty: id("results-empty"),
      downloadBtn: id("download-btn"),
      summaryGrid: id("summary-grid"),
      summaryWarn: id("summary-warn")
    };

    function initTableHeader() {
      var html = "<th>召會 / 區域</th><th>週數</th>";
      metrics.forEach(function (m) { html += "<th>" + m.label + "</th>"; });
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

    function renderChecklist(state) {
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
      els.progressBar.style.width = (CONGREGATIONS.length ? (uploadedCount / CONGREGATIONS.length * 100) : 0) + "%";
    }

    function averageRow(entries) {
      var avg = { weeks: 0 };
      metrics.forEach(function (m) { avg[m.key] = 0; });
      if (!entries.length) return avg;
      avg.weeks = entries.reduce(function (acc, d) { return acc + d.weeks; }, 0) / entries.length;
      metrics.forEach(function (m) {
        avg[m.key] = entries.reduce(function (acc, d) { return acc + (d[m.key] || 0); }, 0) / entries.length;
      });
      return avg;
    }

    function buildRows(state) {
      // returns [{type:'region', region, data}, {type:'congregation', name, data}, ...]
      var out = [];
      CONGREGATION_GROUPS.forEach(function (group) {
        var uploaded = group.members.filter(function (name) {
          return Object.prototype.hasOwnProperty.call(state, name);
        });
        if (!uploaded.length) return;
        var entries = uploaded.map(function (name) { return extract(state[name]); });
        out.push({ type: "region", region: group.region, data: averageRow(entries) });
        uploaded.forEach(function (name) {
          out.push({ type: "congregation", name: name, data: extract(state[name]) });
        });
      });
      return out;
    }

    function renderTable(state, onRemove) {
      els.resultsTbody.innerHTML = "";
      var rowsData = buildRows(state);

      els.resultsEmpty.style.display = rowsData.length ? "none" : "block";
      els.downloadBtn.disabled = !rowsData.length;

      rowsData.forEach(function (r) {
        var tr = document.createElement("tr");
        var html;
        if (r.type === "region") {
          tr.className = "region-row";
          html = "<td>" + r.region + "（區域平均）</td><td>" + formatNum(r.data.weeks) + "</td>";
          metrics.forEach(function (m) { html += "<td>" + formatNum(r.data[m.key]) + "</td>"; });
          html += "<td></td>";
        } else {
          html = "<td class=\"congregation-name\">" + r.name + "</td><td>" + r.data.weeks + "</td>";
          metrics.forEach(function (m) { html += "<td>" + formatNum(r.data[m.key]) + "</td>"; });
          html += "<td><button class=\"btn-danger-ghost\" data-remove=\"" + r.name + "\">移除</button></td>";
        }
        tr.innerHTML = html;
        els.resultsTbody.appendChild(tr);
      });

      els.resultsTbody.querySelectorAll("[data-remove]").forEach(function (btn) {
        btn.addEventListener("click", function () {
          onRemove(btn.getAttribute("data-remove"));
        });
      });
    }

    function renderSummary(state) {
      var totals = {};
      metrics.forEach(function (m) { totals[m.key] = 0; });
      Object.keys(state).forEach(function (name) {
        var d = extract(state[name]);
        metrics.forEach(function (m) { totals[m.key] += d[m.key] || 0; });
      });

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

    function onDownloadClick(state) {
      var rowsData = buildRows(state);
      if (!rowsData.length) return;

      var header = ["召會 / 區域", "週數"].concat(metrics.map(function (m) { return m.label; }));
      var aoa = [header];
      rowsData.forEach(function (r) {
        var label = r.type === "region" ? (r.region + "（區域平均）") : r.name;
        var row = [label, formatNum(r.data.weeks)];
        metrics.forEach(function (m) { row.push(formatNum(r.data[m.key])); });
        aoa.push(row);
      });

      var ws = XLSX.utils.aoa_to_sheet(aoa);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, cfg.summaryTitle);
      XLSX.writeFile(wb, cfg.fileBaseName + ".xlsx");
    }

    initTableHeader();
    initSummaryGrid();

    return {
      renderAll: function (state, onRemove) {
        renderChecklist(state);
        renderTable(state, onRemove);
        renderSummary(state);
      },
      bindDownloadButton: function (getState) {
        els.downloadBtn.addEventListener("click", function () {
          onDownloadClick(getState());
        });
      }
    };
  }

  function loadDataState() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }

  function saveDataState(state) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      /* storage unavailable, ignore */
    }
  }

  function initUploadAndSections() {
    var state = loadDataState();

    var els = {
      select: document.getElementById("congregation-select"),
      dropzone: document.getElementById("dropzone"),
      fileInput: document.getElementById("file-input"),
      filePicked: document.getElementById("file-picked"),
      parseBtn: document.getElementById("parse-btn"),
      resetAllBtn: document.getElementById("reset-all-btn"),
      statusMsg: document.getElementById("status-msg")
    };

    var pickedFile = null;

    function showStatus(message, type) {
      els.statusMsg.textContent = message;
      els.statusMsg.className = "status-msg show " + type;
    }

    function clearStatus() {
      els.statusMsg.className = "status-msg";
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
      CONGREGATION_GROUPS.forEach(function (group) {
        var optgroup = document.createElement("optgroup");
        optgroup.label = group.region;
        group.members.forEach(function (name) {
          var opt = document.createElement("option");
          opt.value = name;
          opt.textContent = name;
          optgroup.appendChild(opt);
        });
        els.select.appendChild(optgroup);
      });
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
        if (files && files.length) handleFileSelected(files[0]);
      });
      els.fileInput.addEventListener("change", function (e) {
        if (e.target.files && e.target.files.length) handleFileSelected(e.target.files[0]);
      });
    }

    var childrenSection = createSection({
      prefix: "children",
      summaryTitle: "全台兒童統計",
      fileBaseName: "children-stats",
      extract: function (entry) { return entry.children; },
      metrics: [
        { key: "avgSunday", label: "主日（週平均）", totalLabel: "主日總計" },
        { key: "avgLife", label: "召會生活（週平均）", totalLabel: "召會生活總計" },
        { key: "avgGroup", label: "小排（週平均）", totalLabel: "小排總計" }
      ]
    });

    var youthSection = createSection({
      prefix: "youth",
      summaryTitle: "全台青職統計",
      fileBaseName: "youth-stats",
      extract: function (entry) { return entry.youth; },
      metrics: [
        { key: "avgSunday", label: "主日（週平均）", totalLabel: "主日總計" },
        { key: "avgFamily", label: "家聚會（出訪+受訪，週平均）", totalLabel: "家聚會總計" },
        { key: "avgGroup", label: "小排（週平均）", totalLabel: "小排總計" },
        { key: "avgLifeReading", label: "生命讀經（週平均）", totalLabel: "生命讀經總計" }
      ]
    });

    function removeCongregation(name) {
      var ok = window.confirm("確定要移除「" + name + "」的資料嗎？（兒童、青職資料會一併移除）");
      if (!ok) return;
      delete state[name];
      saveDataState(state);
      renderAll();
    }

    function renderAll() {
      childrenSection.renderAll(state, removeCongregation);
      youthSection.renderAll(state, removeCongregation);
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
        var scan = scanWeeklyGrid(workbook);
        var childrenResult = deriveChildrenResult(scan);
        var youthResult = deriveYouthResult(scan);

        var isUpdate = Object.prototype.hasOwnProperty.call(state, congregation);
        if (isUpdate) {
          var ok = window.confirm("「" + congregation + "」已經上傳過資料（兒童、青職），是否要用這個新檔案覆蓋？");
          if (!ok) {
            els.parseBtn.disabled = false;
            return;
          }
        }

        state[congregation] = {
          weeks: scan.weeks,
          children: childrenResult,
          youth: youthResult,
          fileName: pickedFile.name,
          updatedAt: new Date().toISOString()
        };
        saveDataState(state);
        renderAll();

        showStatus(
          "已加入「" + congregation + "」：共 " + scan.weeks + " 週。" +
          "兒童 — 主日 " + formatNum(childrenResult.avgSunday) + "、召會生活 " + formatNum(childrenResult.avgLife) + "、小排 " + formatNum(childrenResult.avgGroup) + "；" +
          "青職 — 主日 " + formatNum(youthResult.avgSunday) + "、家聚會 " + formatNum(youthResult.avgFamily) + "、小排 " + formatNum(youthResult.avgGroup) + "、生命讀經 " + formatNum(youthResult.avgLifeReading),
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
      var ok = window.confirm("確定要清除所有已上傳的召會資料嗎？（兒童、青職資料會一併清除，此操作無法復原）");
      if (!ok) return;
      state = {};
      saveDataState(state);
      renderAll();
      clearStatus();
    }

    initSelect();
    setupDropzone();
    els.parseBtn.addEventListener("click", onParseClick);
    els.resetAllBtn.addEventListener("click", onResetAllClick);
    childrenSection.bindDownloadButton(function () { return state; });
    youthSection.bindDownloadButton(function () { return state; });

    renderAll();
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

  function groupsToText(groups) {
    return groups.map(function (g) {
      return g.region + "\n" + g.members.map(function (m) { return "  " + m; }).join("\n");
    }).join("\n");
  }

  function textToGroups(text) {
    var seen = {};
    var groups = [];
    var current = null;

    text.split("\n").forEach(function (line) {
      if (/^\s/.test(line)) {
        var member = line.trim();
        if (!member || seen[member]) return;
        if (!current) {
          current = { region: DEFAULT_REGION_NAME, members: [] };
          groups.push(current);
        }
        seen[member] = true;
        current.members.push(member);
      } else {
        var region = line.trim();
        if (!region) return;
        current = { region: region, members: [] };
        groups.push(current);
      }
    });

    return groups.filter(function (g) { return g.members.length > 0; });
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
      var groups = textToGroups(textarea.value);
      var total = groups.reduce(function (acc, g) { return acc + g.members.length; }, 0);
      countEl.textContent = "目前共 " + groups.length + " 個區域、" + total + " 個召會";
    }

    function open() {
      textarea.value = groupsToText(loadCongregationGroups());
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
      textarea.value = groupsToText(DEFAULT_GROUPS);
      updateCount();
    });

    saveBtn.addEventListener("click", function () {
      var groups = textToGroups(textarea.value);
      saveCongregationGroups(groups);
      window.location.reload();
    });
  }

  function init() {
    initTabs();
    initSettingsModal();
    initUploadAndSections();

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
