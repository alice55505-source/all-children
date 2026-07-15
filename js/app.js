(function () {
  "use strict";

  var CONGREGATIONS = [
    "斗六", "古坑", "林內", "西螺", "莿桐", "斗南", "虎尾", "土庫",
    "崙背", "褒忠", "二崙", "麥寮", "北港", "口湖", "嘉義市", "中埔",
    "竹崎", "番路", "民雄", "溪口", "大林", "新港", "六腳", "朴子",
    "布袋", "鹿草", "太保", "水上"
  ];

  var STORAGE_KEY = "childrenStatsGenerator_v1";
  var WEEKLY_SHEET_NAME = "週報矩陣";
  var SUMMARY_SHEET_NAME = "摘要";
  var TOTAL_LABEL = "合計";
  var LIFE_KPI_LABEL = "期間有召會生活";

  var state = loadState();
  var pickedFile = null;

  var els = {
    select: document.getElementById("congregation-select"),
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    filePicked: document.getElementById("file-picked"),
    parseBtn: document.getElementById("parse-btn"),
    resetAllBtn: document.getElementById("reset-all-btn"),
    statusMsg: document.getElementById("status-msg"),
    progressCount: document.getElementById("progress-count"),
    progressTotal: document.getElementById("progress-total"),
    progressBar: document.getElementById("progress-bar"),
    chipGrid: document.getElementById("chip-grid"),
    resultsTbody: document.getElementById("results-tbody"),
    resultsEmpty: document.getElementById("results-empty"),
    totalSunday: document.getElementById("total-sunday"),
    totalGroup: document.getElementById("total-group"),
    totalLife: document.getElementById("total-life"),
    summaryWarn: document.getElementById("summary-warn"),
    copySummaryBtn: document.getElementById("copy-summary-btn")
  };

  function loadState() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      /* storage unavailable, ignore */
    }
  }

  function initSelect() {
    els.select.innerHTML = "";
    CONGREGATIONS.forEach(function (name) {
      var opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      els.select.appendChild(opt);
    });
  }

  function showStatus(message, type) {
    els.statusMsg.textContent = message;
    els.statusMsg.className = "status-msg show " + type;
  }

  function clearStatus() {
    els.statusMsg.className = "status-msg";
  }

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

  function parseWeeklyMatrix(sheet) {
    if (!sheet) {
      throw new Error("找不到「" + WEEKLY_SHEET_NAME + "」分頁，請確認上傳的檔案格式");
    }

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!rows.length) {
      throw new Error("「" + WEEKLY_SHEET_NAME + "」分頁是空的");
    }

    var header = rows[0].map(function (cell) {
      return cell == null ? "" : String(cell).trim();
    });

    var idx = {
      區排: findHeaderIndex(header, "區排"),
      主日: findHeaderIndex(header, "主日"),
      小排: findHeaderIndex(header, "小排")
    };

    var missing = Object.keys(idx).filter(function (key) { return idx[key] === -1; });
    if (missing.length) {
      throw new Error("在「" + WEEKLY_SHEET_NAME + "」分頁找不到欄位：" + missing.join("、"));
    }

    var sums = { 主日: 0, 小排: 0 };
    var weeks = 0;

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row) continue;
      var label = row[idx.區排];
      if (label != null && String(label).trim() === TOTAL_LABEL) {
        weeks++;
        sums.主日 += Number(row[idx.主日]) || 0;
        sums.小排 += Number(row[idx.小排]) || 0;
      }
    }

    if (weeks === 0) {
      throw new Error("在「" + WEEKLY_SHEET_NAME + "」分頁中找不到「" + TOTAL_LABEL + "」列");
    }

    return {
      weeks: weeks,
      avgSunday: sums.主日 / weeks,
      avgGroup: sums.小排 / weeks
    };
  }

  function parseLifeKpi(sheet) {
    if (!sheet) {
      throw new Error("找不到「" + SUMMARY_SHEET_NAME + "」分頁，請確認上傳的檔案格式");
    }

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row) continue;
      var label = row[0];
      if (label != null && String(label).trim() === LIFE_KPI_LABEL) {
        return Number(row[1]) || 0;
      }
    }

    throw new Error("在「" + SUMMARY_SHEET_NAME + "」分頁找不到「" + LIFE_KPI_LABEL + "」");
  }

  function parseWorkbook(workbook) {
    var weeklySheet = getSheet(workbook, WEEKLY_SHEET_NAME, 2);
    var summarySheet = getSheet(workbook, SUMMARY_SHEET_NAME, 1);

    var weekly = parseWeeklyMatrix(weeklySheet);
    var lifeCount = parseLifeKpi(summarySheet);

    return {
      weeks: weekly.weeks,
      avgSunday: weekly.avgSunday,
      avgGroup: weekly.avgGroup,
      lifeCount: lifeCount
    };
  }

  function formatNum(n) {
    return Math.round(n * 10) / 10;
  }

  function getLifeCount(d) {
    return d.lifeCount != null ? d.lifeCount : d.avgLife;
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
    if (!pickedFile) {
      showStatus("請先選擇一個 .xlsx 檔案", "error");
      return;
    }
    var congregation = els.select.value;
    els.parseBtn.disabled = true;

    pickedFile.arrayBuffer().then(function (buf) {
      var workbook = XLSX.read(buf, { type: "array" });
      var result = parseWorkbook(workbook);
      var isUpdate = Object.prototype.hasOwnProperty.call(state, congregation);

      if (isUpdate) {
        var ok = window.confirm("「" + congregation + "」已經上傳過資料，是否要用這個新檔案覆蓋？");
        if (!ok) {
          els.parseBtn.disabled = false;
          return;
        }
      }

      state[congregation] = {
        weeks: result.weeks,
        avgSunday: result.avgSunday,
        avgGroup: result.avgGroup,
        lifeCount: result.lifeCount,
        fileName: pickedFile.name,
        updatedAt: new Date().toISOString()
      };
      saveState();
      renderAll();

      showStatus(
        "已加入「" + congregation + "」：共 " + result.weeks + " 週，" +
        "主日平均 " + formatNum(result.avgSunday) + "、" +
        "召會生活（期間有召會生活人數）" + formatNum(result.lifeCount) + "、" +
        "小排平均 " + formatNum(result.avgGroup),
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
      var tr = document.createElement("tr");
      tr.innerHTML =
        "<td>" + name + "</td>" +
        "<td>" + d.weeks + "</td>" +
        "<td>" + formatNum(d.avgSunday) + "</td>" +
        "<td>" + formatNum(getLifeCount(d)) + "</td>" +
        "<td>" + formatNum(d.avgGroup) + "</td>" +
        "<td><button class=\"btn-danger-ghost\" data-remove=\"" + name + "\">移除</button></td>";
      els.resultsTbody.appendChild(tr);
    });

    els.resultsTbody.querySelectorAll("[data-remove]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        removeCongregation(btn.getAttribute("data-remove"));
      });
    });
  }

  function renderSummary(uploadedCount) {
    var totals = { sunday: 0, group: 0, life: 0 };
    Object.keys(state).forEach(function (name) {
      var d = state[name];
      totals.sunday += d.avgSunday;
      totals.group += d.avgGroup;
      totals.life += getLifeCount(d);
    });

    els.totalSunday.textContent = formatNum(totals.sunday);
    els.totalGroup.textContent = formatNum(totals.group);
    els.totalLife.textContent = formatNum(totals.life);

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

    return totals;
  }

  function onCopySummaryClick() {
    var uploadedNames = CONGREGATIONS.filter(function (name) {
      return Object.prototype.hasOwnProperty.call(state, name);
    });
    var totals = { sunday: 0, group: 0, life: 0 };
    uploadedNames.forEach(function (name) {
      var d = state[name];
      totals.sunday += d.avgSunday;
      totals.group += d.avgGroup;
      totals.life += getLifeCount(d);
    });

    var lines = [];
    lines.push("全台兒童統計（已上傳 " + uploadedNames.length + " / " + CONGREGATIONS.length + " 個召會）");
    lines.push("主日總計：" + formatNum(totals.sunday));
    lines.push("召會生活總計：" + formatNum(totals.life));
    lines.push("小排總計：" + formatNum(totals.group));

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
    var uploadedCount = renderChecklist();
    renderTable();
    renderSummary(uploadedCount);
  }

  function init() {
    initSelect();
    setupDropzone();
    els.parseBtn.addEventListener("click", onParseClick);
    els.resetAllBtn.addEventListener("click", onResetAllClick);
    els.copySummaryBtn.addEventListener("click", onCopySummaryClick);
    renderAll();

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
