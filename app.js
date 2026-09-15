/* ==========================================================
   SOLAR DGR ANALYTICS — app.js
   Exact Excel data mapping + browser-side analysis engine.

   REQUIRED SOURCES
   Daily_KPI:
     B  = Date
     I  = Operating Hours
     S  = PA %
     V  = PR %
     AD = System Loss %

   PA:
     B  = Date
     W  = Issue / Fault
     Z  = Fault Start
     AC = Work Completion
     AG = Breakdown Time (minutes)
     AL = System Loss (MWh)

   Curtailment records:
     C = Date
     H = Start
     I = End
     R = Loss of Generation (MWh)

   Annual_KPI:
     H10:H21 = Target PR
     I10:I21 = Measured PR
     E10:E21 = Budgeted Energy
     F10:F21 = Measured Energy
   ========================================================== */

(() => {
  "use strict";

  const REQUIRED_SHEETS = [
    "Dashboard",
    "Annual_KPI",
    "Daily_KPI",
    "PA",
    "Curtailment records"
  ];

  const MONTHS = [
    "April", "May", "June", "July", "August", "September",
    "October", "November", "December", "January", "February", "March"
  ];

  const SHORT_MONTHS = [
    "Apr", "May", "Jun", "Jul", "Aug", "Sep",
    "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"
  ];

  const state = {
    workbook: null,
    sheets: {},
    charts: {},
    data: {
      daily: [],
      paEvents: [],
      paBreakdown: [],
      paLoss: [],
      annual: [],
      curtailmentIntervals: [],
      curtailmentDaily: []
    }
  };

  /* ==========================================================
     INITIALISATION
     ========================================================== */

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    bindUpload();
    bindNavigation();
    renderSheetBadges();
  }

  function bindUpload() {
    const uploadButton = document.getElementById("uploadBtn");
    const fileInput = document.getElementById("excelFile");

    if (!uploadButton || !fileInput) return;

    uploadButton.addEventListener("click", () => {
      fileInput.click();
    });

    fileInput.addEventListener("change", async event => {
      const file = event.target.files?.[0];

      if (!file) return;

      setText("fileName", file.name);
      setText("sidebarWorkbook", file.name);

      try {
        await loadWorkbook(file);
      } catch (error) {
        console.error("DGR workbook load error:", error);

        alert(
          "The workbook could not be read. Please upload a valid Excel workbook."
        );
      }
    });
  }

  function bindNavigation() {
    document.querySelectorAll(".nav button").forEach(button => {
      button.addEventListener("click", () => {
        const page = button.dataset.page;

        if (!page) return;

        document
          .querySelectorAll(".nav button")
          .forEach(item => item.classList.remove("active"));

        document
          .querySelectorAll(".page")
          .forEach(item => item.classList.remove("active"));

        button.classList.add("active");

        const section = document.getElementById(`page-${page}`);

        if (section) {
          section.classList.add("active");
        }

        const metadata = {
          dashboard: [
            "Dashboard",
            "Daily plant performance overview"
          ],

          pa: [
            "PA Analysis",
            "Availability, unavailability, breakdown and system loss"
          ],

          performance: [
            "Performance",
            "Daily and monthly performance analysis"
          ],

          curtailment: [
            "Curtailment",
            "Daily losses and curtailment intervals"
          ],

          energy: [
            "Energy",
            "Budgeted versus measured monthly energy"
          ]
        };

        const [title, subtitle] =
          metadata[page] || [
            "Solar DGR Analytics",
            ""
          ];

        setText("pageTitle", title);
        setText("pageSubtitle", subtitle);
      });
    });
  }

  /* ==========================================================
     WORKBOOK LOADING
     ========================================================== */

  async function loadWorkbook(file) {
    if (typeof XLSX === "undefined") {
      throw new Error(
        "SheetJS/XLSX is not available."
      );
    }

    const buffer = await file.arrayBuffer();

    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellFormula: true,
      cellNF: true,
      cellText: true
    });

    state.workbook = workbook;

    state.sheets = resolveSheets(workbook);

    renderSheetBadges();

    extractAllData();

    renderAll();
  }

  function normalizeSheetName(name) {
    return String(name || "")
      .trim()
      .toLowerCase()
      .replace(/[\s_-]+/g, "");
  }

  function resolveSheets(workbook) {
    const sheets = {};

    for (const logicalName of REQUIRED_SHEETS) {
      const wanted = normalizeSheetName(
        logicalName
      );

      const actualName =
        workbook.SheetNames.find(
          name =>
            normalizeSheetName(name) === wanted
        );

      sheets[logicalName] = actualName
        ? workbook.Sheets[actualName]
        : null;
    }

    return sheets;
  }

  function renderSheetBadges() {
    const container =
      document.getElementById("sheetStatus");

    if (!container) return;

    container.innerHTML = "";

    REQUIRED_SHEETS.forEach(logicalName => {
      const exists =
        !!state.sheets[logicalName];

      const badge =
        document.createElement("div");

      badge.className =
        `sheet-badge ${exists ? "ok" : "missing"}`;

      badge.innerHTML =
        `<span class="dot"></span>` +
        escapeHtml(logicalName) +
        (exists ? "" : " · missing");

      container.appendChild(badge);
    });
  }

  /* ==========================================================
     EXCEL CELL ACCESS
     ========================================================== */

  function getCell(sheet, column, row) {
    if (!sheet) return null;

    const address = `${column}${row}`;

    return sheet[address] || null;
  }

  function getDisplayedValue(sheet, column, row) {
    const cell =
      getCell(sheet, column, row);

    if (!cell) return null;

    /*
      Priority:
      1. Excel displayed/cached text
      2. Raw value
    */

    if (
      cell.w !== undefined &&
      cell.w !== null &&
      String(cell.w).trim() !== ""
    ) {
      return cell.w;
    }

    if (
      cell.v !== undefined &&
      cell.v !== null
    ) {
      return cell.v;
    }

    return null;
  }

  function getRawCell(sheet, column, row) {
    return getCell(
      sheet,
      column,
      row
    );
  }

  function getSheetRowCount(sheet) {
    if (
      !sheet ||
      !sheet["!ref"]
    ) {
      return 0;
    }

    const range =
      XLSX.utils.decode_range(
        sheet["!ref"]
      );

    return range.e.r + 1;
  }

  /* ==========================================================
     VALIDATION + NUMBER HELPERS
     ========================================================== */

  function isInvalid(value) {
    if (
      value === null ||
      value === undefined
    ) {
      return true;
    }

    const text =
      String(value)
        .trim()
        .toUpperCase();

    return (
      text === "" ||
      text === "#REF!" ||
      text === "#VALUE!" ||
      text === "#DIV/0!" ||
      text === "#N/A" ||
      text === "#NAME?" ||
      text === "#NUM!" ||
      text === "#NULL!"
    );
  }

  function toNumber(value) {
    if (isInvalid(value)) {
      return null;
    }

    if (typeof value === "number") {
      return Number.isFinite(value)
        ? value
        : null;
    }

    const cleaned =
      String(value)
        .replace(/,/g, "")
        .replace(/%/g, "")
        .trim();

    if (!cleaned) {
      return null;
    }

    const number =
      Number(cleaned);

    return Number.isFinite(number)
      ? number
      : null;
  }

  function toPercentage(value) {
    const number =
      toNumber(value);

    if (number === null) {
      return null;
    }

    /*
      If Excel stores:
      0.95  => 95%
      95    => 95%
    */

    return Math.abs(number) <= 1.5
      ? number * 100
      : number;
  }

  /* ==========================================================
     DATE HELPERS
     ========================================================== */

  function excelSerialToDate(serial) {
    if (!Number.isFinite(serial)) {
      return null;
    }

    const epoch =
      Date.UTC(
        1899,
        11,
        30
      );

    const timestamp =
      epoch +
      serial * 86400000;

    const date =
      new Date(timestamp);

    if (
      Number.isNaN(
        date.getTime()
      )
    ) {
      return null;
    }

    return new Date(
      date.getUTCFullYear(),
      date.getUTCMonth(),
      date.getUTCDate()
    );
  }

  function parseDate(value) {
    if (isInvalid(value)) {
      return null;
    }

    if (value instanceof Date) {
      if (
        Number.isNaN(
          value.getTime()
        )
      ) {
        return null;
      }

      return new Date(
        value.getFullYear(),
        value.getMonth(),
        value.getDate()
      );
    }

    if (typeof value === "number") {
      return excelSerialToDate(value);
    }

    const text =
      String(value).trim();

    /*
      Prevent times such as 12:00
      from becoming dates.
    */

    if (
      /^\d{1,2}:\d{2}(:\d{2})?$/.test(
        text
      )
    ) {
      return null;
    }

    /*
      yyyy-mm-dd
    */

    let match =
      text.match(
        /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/
      );

    if (match) {
      const year =
        Number(match[1]);

      const month =
        Number(match[2]);

      const day =
        Number(match[3]);

      const date =
        new Date(
          year,
          month - 1,
          day
        );

      return isValidDateParts(
        date,
        year,
        month,
        day
      )
        ? date
        : null;
    }

    /*
      dd/mm/yyyy
    */

    match =
      text.match(
        /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/
      );

    if (match) {
      const day =
        Number(match[1]);

      const month =
        Number(match[2]);

      const year =
        Number(match[3]);

      const date =
        new Date(
          year,
          month - 1,
          day
        );

      return isValidDateParts(
        date,
        year,
        month,
        day
      )
        ? date
        : null;
    }

    const parsed =
      new Date(text);

    if (
      Number.isNaN(
        parsed.getTime()
      )
    ) {
      return null;
    }

    return new Date(
      parsed.getFullYear(),
      parsed.getMonth(),
      parsed.getDate()
    );
  }

  function isValidDateParts(
    date,
    year,
    month,
    day
  ) {
    return (
      date.getFullYear() === year &&
      date.getMonth() ===
        month - 1 &&
      date.getDate() === day
    );
  }

  function dateKey(date) {
    if (
      !(date instanceof Date) ||
      Number.isNaN(
        date.getTime()
      )
    ) {
      return "";
    }

    return [
      date.getFullYear(),
      String(
        date.getMonth() + 1
      ).padStart(2, "0"),
      String(
        date.getDate()
      ).padStart(2, "0")
    ].join("-");
  }

  function addDays(date, days) {
    const result =
      new Date(date);

    result.setDate(
      result.getDate() + days
    );

    return result;
  }

  function formatDate(date) {
    if (
      !(date instanceof Date)
    ) {
      return "—";
    }

    return date.toLocaleDateString(
      "en-GB",
      {
        day: "2-digit",
        month: "short",
        year: "numeric"
      }
    );
  }

  /* ==========================================================
     TIME HELPERS
     ========================================================== */

  function parseTimeMinutes(value) {
    if (isInvalid(value)) {
      return null;
    }

    if (value instanceof Date) {
      if (
        Number.isNaN(
          value.getTime()
        )
      ) {
        return null;
      }

      return (
        value.getHours() * 60 +
        value.getMinutes() +
        value.getSeconds() / 60
      );
    }

    if (
      typeof value === "number" &&
      Number.isFinite(value)
    ) {
      /*
        Excel time:
        0.5 = 12:00 = 720 minutes
      */

      if (
        value >= 0 &&
        value < 1
      ) {
        return value * 1440;
      }

      /*
        Direct minute value.
      */

      if (
        value >= 1 &&
        value <= 1440
      ) {
        return value;
      }

      return null;
    }

    const text =
      String(value).trim();

    const match =
      text.match(
        /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
      );

    if (match) {
      const hours =
        Number(match[1]);

      const minutes =
        Number(match[2]);

      const seconds =
        Number(match[3] || 0);

      if (
        hours < 0 ||
        hours > 23 ||
        minutes < 0 ||
        minutes > 59 ||
        seconds < 0 ||
        seconds > 59
      ) {
        return null;
      }

      return (
        hours * 60 +
        minutes +
        seconds / 60
      );
    }

    return null;
  }

  function minutesToTime(minutes) {
    const normalized =
      (
        (
          Math.round(minutes) %
          1440
        ) +
        1440
      ) % 1440;

    const hours =
      Math.floor(
        normalized / 60
      );

    const mins =
      normalized % 60;

    return (
      String(hours)
        .padStart(2, "0") +
      ":" +
      String(mins)
        .padStart(2, "0")
    );
  }

  /* ==========================================================
     DAILY_KPI
     B  = DATE
     I  = OPERATING HOURS
     S  = PA %
     V  = PR %
     AD = SYSTEM LOSS %
     ========================================================== */

  function readDailyKPI() {
    const sheet =
      state.sheets["Daily_KPI"];

    if (!sheet) {
      return [];
    }

    const lastRow =
      getSheetRowCount(sheet);

    const records = [];

    let previousDate = null;

    for (
      let row = 1;
      row <= lastRow;
      row++
    ) {
      const dateCell =
        getRawCell(
          sheet,
          "B",
          row
        );

      let date =
        parseDate(
          getDisplayedValue(
            sheet,
            "B",
            row
          )
        );

      /*
        Formula-driven date handling.

        Example:
        B6 = B5+1

        If SheetJS cannot retrieve
        a valid cached date, reconstruct
        from the previous valid date.
      */

      if (
        !date &&
        dateCell &&
        dateCell.f &&
        previousDate
      ) {
        date =
          addDays(
            previousDate,
            1
          );
      }

      /*
        Protect against the workbook
        exposing the same cached displayed
        date for sequential formulas.

        Only do this when the formula
        explicitly indicates previous-row + 1.
      */

      if (
        dateCell &&
        dateCell.f &&
        previousDate &&
        /B\d+\s*\+\s*1/i.test(
          String(dateCell.f)
        )
      ) {
        const parsed =
          date;

        if (
          !parsed ||
          dateKey(parsed) ===
            dateKey(previousDate)
        ) {
          date =
            addDays(
              previousDate,
              1
            );
        }
      }

      if (date) {
        previousDate = date;
      }

      if (!date) {
        continue;
      }

      const pa =
        toPercentage(
          getDisplayedValue(
            sheet,
            "S",
            row
          )
        );

      const pr =
        toPercentage(
          getDisplayedValue(
            sheet,
            "V",
            row
          )
        );

      const systemLoss =
        toPercentage(
          getDisplayedValue(
            sheet,
            "AD",
            row
          )
        );

      const operatingHours =
        toNumber(
          getDisplayedValue(
            sheet,
            "I",
            row
          )
        );

      const validMetricCount = [
        pa,
        pr,
        systemLoss,
        operatingHours
      ].filter(
        value => value !== null
      ).length;

      /*
        Ignore rows containing
        only a date and no mapped metric.
      */

      if (
        validMetricCount === 0
      ) {
        continue;
      }

      records.push({
        date,
        key: dateKey(date),
        pa,
        pr,
        systemLoss,
        operatingHours,
        validMetricCount,
        sourceRow: row
      });
    }

    /*
      Duplicate date rule:
      retain row with greatest
      number of valid metrics.
    */

    const byDate =
      new Map();

    for (
      const record of records
    ) {
      const current =
        byDate.get(
          record.key
        );

      if (
        !current ||
        record.validMetricCount >
          current.validMetricCount
      ) {
        byDate.set(
          record.key,
          record
        );
      }
    }

    return [
      ...byDate.values()
    ].sort(
      (a, b) =>
        a.date - b.date
    );
  }

  /* ==========================================================
     PA — PLANT UNAVAILABILITY
     W  = ISSUE / FAULT
     Z  = FAULT START
     AC = WORK COMPLETION
     ========================================================== */

  function readPAEvents() {
    const sheet =
      state.sheets["PA"];

    if (!sheet) {
      return [];
    }

    const lastRow =
      getSheetRowCount(sheet);

    const events = [];

    for (
      let row = 1;
      row <= lastRow;
      row++
    ) {
      const issue =
        getDisplayedValue(
          sheet,
          "W",
          row
        );

      if (
        isInvalid(issue)
      ) {
        continue;
      }

      const start =
        parseTimeMinutes(
          getDisplayedValue(
            sheet,
            "Z",
            row
          )
        );

      const originalEnd =
        parseTimeMinutes(
          getDisplayedValue(
            sheet,
            "AC",
            row
          )
        );

      if (
        start === null ||
        originalEnd === null
      ) {
        continue;
      }

      let end =
        originalEnd;

      /*
        Overnight fault:
        end < start
        => add 24 hours.
      */

      if (end < start) {
        end += 1440;
      }

      events.push({
        issue:
          String(issue).trim(),

        start,
        end,

        duration:
          end - start,

        sourceRow: row
      });
    }

    return events;
  }

  /* ==========================================================
     PA — BREAKDOWN
     AG = BREAKDOWN TIME IN MINUTES
     ========================================================== */

  function readPABreakdown() {
    const sheet =
      state.sheets["PA"];

    if (!sheet) {
      return [];
    }

    const lastRow =
      getSheetRowCount(sheet);

    const map =
      new Map();

    for (
      let row = 1;
      row <= lastRow;
      row++
    ) {
      const date =
        parseDate(
          getDisplayedValue(
            sheet,
            "B",
            row
          )
        );

      const breakdown =
        toNumber(
          getDisplayedValue(
            sheet,
            "AG",
            row
          )
        );

      if (
        !date ||
        breakdown === null
      ) {
        continue;
      }

      const key =
        dateKey(date);

      if (!map.has(key)) {
        map.set(
          key,
          {
            date,
            value: 0
          }
        );
      }

      /*
        AG is already defined
        as duration in minutes.
        Never interpret it as
        time-of-day.
      */

      map.get(key).value +=
        breakdown;
    }

    return [
      ...map.values()
    ].sort(
      (a, b) =>
        a.date - b.date
    );
  }

  /* ==========================================================
     PA — SYSTEM LOSS MWh
     AL = SYSTEM LOSS
     ========================================================== */

  function readPASystemLoss() {
    const sheet =
      state.sheets["PA"];

    if (!sheet) {
      return [];
    }

    const lastRow =
      getSheetRowCount(sheet);

    const map =
      new Map();

    for (
      let row = 1;
      row <= lastRow;
      row++
    ) {
      const date =
        parseDate(
          getDisplayedValue(
            sheet,
            "B",
            row
          )
        );

      const loss =
        toNumber(
          getDisplayedValue(
            sheet,
            "AL",
            row
          )
        );

      if (
        !date ||
        loss === null
      ) {
        continue;
      }

      const key =
        dateKey(date);

      if (!map.has(key)) {
        map.set(
          key,
          {
            date,
            value: 0
          }
        );
      }

      /*
        AL is MWh.
        It is completely separate
        from Daily_KPI AD.
      */

      map.get(key).value +=
        loss;
    }

    return [
      ...map.values()
    ].sort(
      (a, b) =>
        a.date - b.date
    );
  }

  /* ==========================================================
     ANNUAL_KPI
     H10:H21 = TARGET PR
     I10:I21 = MEASURED PR
     E10:E21 = BUDGETED ENERGY
     F10:F21 = MEASURED ENERGY
     ========================================================== */

  function readAnnualKPI() {
    const sheet =
      state.sheets["Annual_KPI"];

    if (!sheet) {
      return [];
    }

    return MONTHS.map(
      (month, index) => {
        const row =
          10 + index;

        return {
          month,

          shortMonth:
            SHORT_MONTHS[index],

          targetPR:
            toPercentage(
              getDisplayedValue(
                sheet,
                "H",
                row
              )
            ),

          measuredPR:
            toPercentage(
              getDisplayedValue(
                sheet,
                "I",
                row
              )
            ),

          budgetedEnergy:
            toNumber(
              getDisplayedValue(
                sheet,
                "E",
                row
              )
            ),

          measuredEnergy:
            toNumber(
              getDisplayedValue(
                sheet,
                "F",
                row
              )
            ),

          sourceRow: row
        };
      }
    );
  }

  /* ==========================================================
     CURTAILMENT
     C = DATE
     H = START
     I = END
     R = LOSS MWh
     ========================================================== */

  function readCurtailment() {
    const ws = getSheet("Curtailment records");

    if (!ws) {
        return {
            daily: [],
            intervals: []
        };
    }

    const dailyMap = new Map();
    const intervals = [];

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

    for (let r = 1; r <= range.e.r; r++) {
        // -------------------------------------------------
        // EXACT CURTAILMENT MAPPING
        // C = Date
        // H = From Time
        // I = To Time
        // R = Loss of Generation MWh
        // -------------------------------------------------

        const dateCell = ws[XLSX.utils.encode_cell({ r, c: 2 })];  // C
        const fromCell = ws[XLSX.utils.encode_cell({ r, c: 7 })];  // H
        const toCell = ws[XLSX.utils.encode_cell({ r, c: 8 })];    // I
        const lossCell = ws[XLSX.utils.encode_cell({ r, c: 17 })]; // R

        const date = parseExcelDateCell(dateCell);

        if (!date) continue;

        const fromMinutes = parseExcelTimeCell(fromCell);
        const toMinutes = parseExcelTimeCell(toCell);
        const loss = toNumber(lossCell);

        // A row is useful for the daily loss analysis
        // even if timing information is unavailable.
        if (loss !== null) {
            const dateKey = formatDateKey(date);

            if (!dailyMap.has(dateKey)) {
                dailyMap.set(dateKey, {
                    date: date,
                    loss: 0,
                    intervals: 0
                });
            }

            dailyMap.get(dateKey).loss += loss;
            dailyMap.get(dateKey).intervals += 1;
        }

        // For the duration Gantt, both start and end
        // timings are required.
        if (
            fromMinutes !== null &&
            toMinutes !== null
        ) {
            let endMinutes = toMinutes;

            // Overnight interval
            if (endMinutes < fromMinutes) {
                endMinutes += 1440;
            }

            intervals.push({
                date: date,
                dateKey: formatDateKey(date),
                start: fromMinutes,
                end: endMinutes,
                loss: loss !== null ? loss : 0
            });
        }
    }

    const daily = Array.from(dailyMap.values())
        .sort((a, b) => a.date - b.date);

    intervals.sort((a, b) => {
        if (a.date - b.date !== 0) {
            return a.date - b.date;
        }

        return a.start - b.start;
    });

    return {
        daily,
        intervals
    };
}
  /* ==========================================================
     EXTRACT ALL DATA
     ========================================================== */

  function extractAllData() {
    state.data.daily =
      readDailyKPI();

    state.data.paEvents =
      readPAEvents();

    state.data.paBreakdown =
      readPABreakdown();

    state.data.paLoss =
      readPASystemLoss();

    state.data.annual =
      readAnnualKPI();

    const curtailment =
      readCurtailment();

    state.data.curtailmentIntervals =
      curtailment.intervals;

    state.data.curtailmentDaily =
      curtailment.daily;
  }

  /* ==========================================================
     CHART LIFECYCLE
     ========================================================== */

  function destroyChart(key) {
    if (
      state.charts[key]
    ) {
      state.charts[key].destroy();

      delete state.charts[key];
    }
  }

  function showNoData(
    wrapperId,
    message = "No data found"
  ) {
    const wrapper =
      document.getElementById(
        wrapperId
      );

    if (!wrapper) return;

    destroyChartByCanvas(
      wrapper
    );

    wrapper.innerHTML = "";

    wrapper.style.width =
      "100%";

    wrapper.style.height =
      "";

    const element =
      document.createElement(
        "div"
      );

    element.className =
      "no-data";

    element.textContent =
      message;

    wrapper.appendChild(
      element
    );
  }

  function destroyChartByCanvas(
    wrapper
  ) {
    const canvas =
      wrapper.querySelector(
        "canvas"
      );

    if (!canvas) return;

    Object.keys(
      state.charts
    ).forEach(key => {
      const chart =
        state.charts[key];

      if (
        chart &&
        chart.canvas === canvas
      ) {
        chart.destroy();

        delete state.charts[key];
      }
    });
  }

  function chartPixelWidth(
    count,
    pixelsPerPoint = 76
  ) {
    return Math.max(
      760,
      Math.max(1, count) *
        pixelsPerPoint
    );
  }

  function prepareChart(
    wrapperId,
    count,
    pixelsPerPoint = 76
  ) {
    const wrapper =
      document.getElementById(
        wrapperId
      );

    if (!wrapper) {
      return null;
    }

    destroyChartByCanvas(
      wrapper
    );

    const width =
      chartPixelWidth(
        count,
        pixelsPerPoint
      );

    wrapper.style.width =
      `${width}px`;

    wrapper.innerHTML = "";

    const canvas =
      document.createElement(
        "canvas"
      );

    canvas.width =
      width;

    canvas.height =
      wrapper.clientHeight ||
      315;

    wrapper.appendChild(
      canvas
    );

    return canvas;
  }

  function commonChartOptions() {
    return {
      responsive: false,

      maintainAspectRatio:
        false,

      animation: false,

      interaction: {
        mode: "index",
        intersect: false
      },

      plugins: {
        legend: {
          labels: {
            font: {
              size: 10
            },

            usePointStyle:
              true
          }
        },

        tooltip: {
          titleFont: {
            size: 11
          },

          bodyFont: {
            size: 10
          }
        }
      }
    };
  }

  /* ==========================================================
     LINE CHART
     ========================================================== */

  function createLineChart(config) {
    const {
      key,
      wrapperId,
      labels,
      values,
      datasetLabel,
      yMin,
      yMax,
      yStep,
      yTitle,
      percentage = false,
      pixelsPerPoint = 76
    } = config;

    destroyChart(key);

    if (!values.length) {
      showNoData(
        wrapperId
      );

      return;
    }

    const canvas =
      prepareChart(
        wrapperId,
        values.length,
        pixelsPerPoint
      );

    if (!canvas) return;

    try {
      state.charts[key] =
        new Chart(
          canvas,
          {
            type: "line",

            data: {
              labels,

              datasets: [{
                label:
                  datasetLabel,

                data: values,

                borderWidth: 2,

                pointRadius: 2.5,

                pointHoverRadius: 4,

                tension: 0.22,

                fill: false
              }]
            },

            options: {
              ...commonChartOptions(),

              scales: {
                x: {
                  ticks: {
                    autoSkip: false,

                    maxRotation: 45,

                    minRotation: 45,

                    font: {
                      size: 9
                    },

                    callback(
                      value,
                      index
                    ) {
                      return (
                        index % 2 === 0
                      )
                        ? this.getLabelForValue(
                            value
                          )
                        : "";
                    }
                  },

                  grid: {
                    display: false
                  }
                },

                y: {
                  min: yMin,

                  max: yMax,

                  ticks: {
                    stepSize: yStep,

                    font: {
                      size: 9
                    },

                    callback(value) {
                      return percentage
                        ? `${value}%`
                        : value;
                    }
                  },

                  title: {
                    display: true,

                    text: yTitle,

                    font: {
                      size: 10,

                      weight: "700"
                    }
                  }
                }
              }
            }
          }
        );
    } catch (error) {
      console.error(
        `Line chart failed: ${key}`,
        error
      );

      showNoData(
        wrapperId,
        "Unable to render this chart."
      );
    }
  }

  /* ==========================================================
     BAR CHART
     ========================================================== */

  function createBarChart(config) {
    const {
      key,
      wrapperId,
      labels,
      values,
      datasetLabel,
      secondValues = null,
      secondDatasetLabel = null,
      indexAxis = "x",
      xMin,
      xMax,
      yMin,
      yMax,
      xStep,
      yStep,
      xTitle = "",
      yTitle = "",
      percentage = false,
      pixelsPerPoint = 76
    } = config;

    destroyChart(key);

    const valid =
      values.some(
        value =>
          value !== null
      );

    if (
      !valid &&
      !secondValues
    ) {
      showNoData(
        wrapperId
      );

      return;
    }

    const canvas =
      prepareChart(
        wrapperId,
        labels.length,
        indexAxis === "y"
          ? 20
          : pixelsPerPoint
      );

    if (!canvas) return;

    const datasets = [{
      label:
        datasetLabel,

      data: values,

      borderWidth: 1,

      borderRadius: 3
    }];

    if (secondValues) {
      datasets.push({
        label:
          secondDatasetLabel,

        data:
          secondValues,

        borderWidth: 1,

        borderRadius: 3
      });
    }

    try {
      state.charts[key] =
        new Chart(
          canvas,
          {
            type: "bar",

            data: {
              labels,

              datasets
            },

            options: {
              ...commonChartOptions(),

              indexAxis,

              scales: {
                x: {
                  min: xMin,

                  max: xMax,

                  ticks: {
                    stepSize: xStep,

                    font: {
                      size: 9
                    },

                    callback(value) {
                      return percentage
                        ? `${value}%`
                        : value;
                    }
                  },

                  title: {
                    display:
                      !!xTitle,

                    text: xTitle,

                    font: {
                      size: 10,

                      weight: "700"
                    }
                  }
                },

                y: {
                  min: yMin,

                  max: yMax,

                  ticks: {
                    stepSize: yStep,

                    font: {
                      size: 9
                    },

                    callback(value) {
                      return percentage
                        ? `${value}%`
                        : value;
                    }
                  },

                  title: {
                    display:
                      !!yTitle,

                    text: yTitle,

                    font: {
                      size: 10,

                      weight: "700"
                    }
                  }
                }
              }
            }
          }
        );
    } catch (error) {
      console.error(
        `Bar chart failed: ${key}`,
        error
      );

      showNoData(
        wrapperId,
        "Unable to render this chart."
      );
    }
  }

  /* ==========================================================
     DASHBOARD
     ========================================================== */

  function renderDashboard() {
    const data =
      state.data.daily;

    if (
      !state.sheets["Daily_KPI"] ||
      !data.length
    ) {
      setText(
        "dashPA",
        "—"
      );

      setText(
        "dashPR",
        "—"
      );

      setText(
        "dashLoss",
        "—"
      );

      setText(
        "dashHours",
        "—"
      );

      setText(
        "dashPADate",
        "No Daily_KPI data"
      );

      setText(
        "dashPRDate",
        "No Daily_KPI data"
      );

      setText(
        "dashLossDate",
        "No Daily_KPI data"
      );

      setText(
        "dashHoursDate",
        "No Daily_KPI data"
      );

      showNoData(
        "dashPRChartWrap",
        state.sheets["Daily_KPI"]
          ? "No data found"
          : "Daily_KPI worksheet missing"
      );

      showNoData(
        "dashLossChartWrap",
        state.sheets["Daily_KPI"]
          ? "No data found"
          : "Daily_KPI worksheet missing"
      );

      setText(
        "dashRange",
        "No Daily_KPI data"
      );

      return;
    }

    /*
      Because records are sorted,
      the last record is the latest
      chronologically valid record.
    */

    const latest =
      data[data.length - 1];

    const latestDate =
      formatDate(
        latest.date
      );

    setText(
      "dashPA",
      latest.pa === null
        ? "—"
        : `${latest.pa.toFixed(2)}%`
    );

    setText(
      "dashPR",
      latest.pr === null
        ? "—"
        : `${latest.pr.toFixed(2)}%`
    );

    setText(
      "dashLoss",
      latest.systemLoss === null
        ? "—"
        : `${latest.systemLoss.toFixed(2)}%`
    );

    setText(
      "dashHours",
      latest.operatingHours === null
        ? "—"
        : latest.operatingHours.toFixed(2)
    );

    setText(
      "dashPADate",
      latestDate
    );

    setText(
      "dashPRDate",
      latestDate
    );

    setText(
      "dashLossDate",
      latestDate
    );

    setText(
      "dashHoursDate",
      latestDate
    );

    setText(
      "dashRange",
      `${formatDate(
        data[0].date
      )} → ${latestDate}`
    );

    const pr =
      data.filter(
        record =>
          record.pr !== null
      );

    createLineChart({
      key: "dashPR",

      wrapperId:
        "dashPRChartWrap",

      labels:
        pr.map(
          record =>
            formatDate(
              record.date
            )
        ),

      values:
        pr.map(
          record =>
            record.pr
        ),

      datasetLabel:
        "PR (%)",

      yMin: 0,

      yMax: 100,

      yStep: 20,

      yTitle:
        "PR (%)",

      percentage: true
    });

    const losses =
      data.filter(
        record =>
          record.systemLoss !== null
      );

    const maximumLoss =
      losses.length
        ? Math.max(
            ...losses.map(
              record =>
                record.systemLoss
            )
          )
        : 0;

    createLineChart({
      key: "dashLoss",

      wrapperId:
        "dashLossChartWrap",

      labels:
        losses.map(
          record =>
            formatDate(
              record.date
            )
        ),

      values:
        losses.map(
          record =>
            record.systemLoss
        ),

      datasetLabel:
        "System Loss (%)",

     yMin: 0,

yMax: 3,

yStep: 0.5,

      yTitle:
        "System Loss (%)",

      percentage: true
    });
  }

  /* ==========================================================
     PA ANALYSIS
     ========================================================== */

  function renderPAAnalysis() {
    renderPlantAvailability();

    renderUnavailabilityGantt();

    renderBreakdownTimeline();

    renderPASystemLoss();
  }

  function renderPlantAvailability() {
    const data =
      state.data.daily.filter(
        record =>
          record.pa !== null
      );

    if (
      !state.sheets["Daily_KPI"]
    ) {
      showNoData(
        "paAvailabilityWrap",
        "Daily_KPI worksheet missing"
      );

      return;
    }

    createLineChart({
      key:
        "paAvailability",

      wrapperId:
        "paAvailabilityWrap",

      labels:
        data.map(
          record =>
            formatDate(
              record.date
            )
        ),

      values:
        data.map(
          record =>
            record.pa
        ),

      datasetLabel:
        "Plant Availability (%)",

      yMin: 80,

      yMax: 100,

      yStep: 5,

      yTitle:
        "Plant Availability (%)",

      percentage: true
    });
  }

  /* ==========================================================
     PA UNAVAILABILITY GANTT
     ========================================================== */

  function renderUnavailabilityGantt() {
    destroyChart(
      "paGantt"
    );

    if (
      !state.sheets["PA"]
    ) {
      showNoData(
        "paGanttWrap",
        "PA worksheet missing"
      );

      return;
    }

    const events =
      state.data.paEvents;

    if (!events.length) {
      showNoData(
        "paGanttWrap"
      );

      return;
    }

    const issues = [
      ...new Set(
        events.map(
          event =>
            event.issue
        )
      )
    ];

    const wrapper =
      document.getElementById(
        "paGanttWrap"
      );

    if (!wrapper) return;

    const width =
      Math.max(
        1150,
        issues.length * 125
      );

    wrapper.style.width =
      `${width}px`;

    wrapper.style.height =
      "390px";

    wrapper.innerHTML = "";

    const canvas =
      document.createElement(
        "canvas"
      );

    canvas.width =
      width;

    canvas.height =
      390;

    wrapper.appendChild(
      canvas
    );

    /*
      One dataset per issue.
      Every fault event remains separate.
    */

    const datasets =
      issues.map(
        issue => ({
          label: issue,

          data:
            events
              .filter(
                event =>
                  event.issue ===
                  issue
              )
              .map(
                event => ({
                  x: [
                    event.start,
                    event.end
                  ],

                  y: issue,

                  issue:
                    event.issue,

                  start:
                    event.start,

                  end:
                    event.end,

                  duration:
                    event.duration,

                  sourceRow:
                    event.sourceRow
                })
              ),

          parsing: false,

          borderWidth: 8,

          borderSkipped:
            false,

          pointRadius: 0
        })
      );

    try {
      state.charts.paGantt =
        new Chart(
          canvas,
          {
            type: "bar",

            data: {
              labels: issues,

              datasets
            },

            options: {
              ...commonChartOptions(),

              indexAxis: "y",

              scales: {
                x: {
                  type:
                    "linear",

                  min: 0,

                  max: 1440,

                  ticks: {
                    stepSize: 30,

                    font: {
                      size: 9
                    },

                    callback(value) {
                      return minutesToTime(
                        value
                      );
                    }
                  },

                  title: {
                    display: true,

                    text:
                      "Time of day",

                    font: {
                      size: 10,

                      weight: "700"
                    }
                  }
                },

                y: {
                  ticks: {
                    font: {
                      size: 9
                    }
                  }
                }
              },

              plugins: {
                legend: {
                  display: false
                },

                tooltip: {
                  callbacks: {
                    title(items) {
                      return (
                        items[0]?.raw
                          ?.issue ||
                        ""
                      );
                    },

                    label(item) {
                      const raw =
                        item.raw;

                      return [
                        `Start: ${minutesToTime(
                          raw.start
                        )}`,

                        `End: ${minutesToTime(
                          raw.end
                        )}`,

                        `Duration: ${raw.duration.toFixed(
                          0
                        )} min`
                      ];
                    }
                  }
                }
              }
            }
          }
        );
    } catch (error) {
      console.error(
        "PA unavailability Gantt failed:",
        error
      );

      showNoData(
        "paGanttWrap",
        "Unable to render this chart."
      );
    }
  }

  /* ==========================================================
     BREAKDOWN TIMELINE
     ========================================================== */

  function renderBreakdownTimeline() {
    const data =
      state.data.paBreakdown;

    if (
      !state.sheets["PA"]
    ) {
      showNoData(
        "breakdownWrap",
        "PA worksheet missing"
      );

      return;
    }

    if (!data.length) {
      showNoData(
        "breakdownWrap"
      );

      return;
    }

    createBarChart({
      key:
        "breakdown",

      wrapperId:
        "breakdownWrap",

      labels:
        data.map(
          record =>
            formatDate(
              record.date
            )
        ),

      values:
        data.map(
          record =>
            record.value
        ),

      datasetLabel:
        "Breakdown Time",

      indexAxis:
        "y",

      xMin: 0,

      xMax: 13,

      xStep: 1,

      xTitle:
        "Breakdown Time (minutes)",

      pixelsPerPoint: 20
    });
  }

  /* ==========================================================
     PA SYSTEM LOSS MWh
     ========================================================== */

  function renderPASystemLoss() {
    const data =
      state.data.paLoss;

    if (
      !state.sheets["PA"]
    ) {
      showNoData(
        "paLossWrap",
        "PA worksheet missing"
      );

      return;
    }

    if (!data.length) {
      showNoData(
        "paLossWrap"
      );

      return;
    }

    const values =
      data.map(
        record =>
          record.value
      );

    const maximum =
      Math.max(...values);

    createBarChart({
      key:
        "paLoss",

      wrapperId:
        "paLossWrap",

      labels:
        data.map(
          record =>
            formatDate(
              record.date
            )
        ),

      values,

      datasetLabel:
        "System Loss (MWh)",

      yMin: 0,

      yMax: Math.max(
        1,
        Math.ceil(
          maximum * 1.15
        )
      ),

      yTitle:
        "System Loss (MWh)",

      pixelsPerPoint: 60
    });
  }

  /* ==========================================================
     PERFORMANCE
     ========================================================== */

  function renderPerformance() {
    const daily =
      state.data.daily;

    if (
      !state.sheets["Daily_KPI"]
    ) {
      showNoData(
        "performancePRWrap",
        "Daily_KPI worksheet missing"
      );

      showNoData(
        "operatingWrap",
        "Daily_KPI worksheet missing"
      );

      showNoData(
        "performanceLossWrap",
        "Daily_KPI worksheet missing"
      );
    } else {
      const pr =
        daily.filter(
          record =>
            record.pr !== null
        );

      createLineChart({
        key:
          "performancePR",

        wrapperId:
          "performancePRWrap",

        labels:
          pr.map(
            record =>
              formatDate(
                record.date
              )
          ),

        values:
          pr.map(
            record =>
              record.pr
          ),

        datasetLabel:
          "PR (%)",

        yMin: 0,

        yMax: 100,

        yStep: 20,

        yTitle:
          "PR (%)",

        percentage: true
      });

      const operating =
        daily.filter(
          record =>
            record.operatingHours !==
            null
        );

      const maxHours =
        operating.length
          ? Math.max(
              ...operating.map(
                record =>
                  record.operatingHours
              )
            )
          : 24;

      createLineChart({
        key:
          "operating",

        wrapperId:
          "operatingWrap",

        labels:
          operating.map(
            record =>
              formatDate(
                record.date
              )
          ),

        values:
          operating.map(
            record =>
              record.operatingHours
          ),

        datasetLabel:
          "Operating Hours",

        yMin: 0,

        yMax: Math.max(
          24,
          Math.ceil(
            maxHours
          )
        ),

        yStep: 2,

        yTitle:
          "Operating Hours"
      });

      const losses =
        daily.filter(
          record =>
            record.systemLoss !==
            null
        );

      const maxLoss =
        losses.length
          ? Math.max(
              ...losses.map(
                record =>
                  record.systemLoss
              )
            )
          : 0;

      createLineChart({
        key:
          "performanceLoss",

        wrapperId:
          "performanceLossWrap",

        labels:
          losses.map(
            record =>
              formatDate(
                record.date
              )
          ),

        values:
          losses.map(
            record =>
              record.systemLoss
          ),

        datasetLabel:
          "System Loss (%)",

        yMin: 0,
yMax: 3,
yStep: 0.5,

        yTitle:
          "System Loss (%)",

        percentage: true
      });
    }

    renderMonthlyPR();
  }

  /* ==========================================================
     MONTHLY PR
     ANNUAL_KPI I10:I21
     ========================================================== */

  function renderMonthlyPR() {
    if (
      !state.sheets["Annual_KPI"]
    ) {
      showNoData(
        "monthlyPRWrap",
        "Annual_KPI worksheet missing"
      );

      return;
    }

    const annual =
      state.data.annual;

    if (
      !annual.length ||
      !annual.some(
        record =>
          record.measuredPR !==
          null
      )
    ) {
      showNoData(
        "monthlyPRWrap"
      );

      return;
    }

    createBarChart({
      key:
        "monthlyPR",

      wrapperId:
        "monthlyPRWrap",

      labels:
        annual.map(
          record =>
            record.month
        ),

      values:
        annual.map(
          record =>
            record.measuredPR
        ),

      datasetLabel:
        "Measured PR (%)",

      indexAxis:
        "y",

      xMin: 0,

      xMax: 100,

      xStep: 20,

      xTitle:
        "Performance Ratio (%)",

      percentage: true,

      pixelsPerPoint: 20
    });
  }

  /* ==========================================================
     CURTAILMENT
     ========================================================== */

  function renderCurtailment() {
    renderCurtailmentTable();

    renderCurtailmentTrend();

    renderCurtailmentGantt();
  }

  /* ==========================================================
     CURTAILMENT TABLE
     ========================================================== */

  function renderCurtailmentTable() {
    const tbody =
      document.getElementById(
        "curtailmentTableBody"
      );

    if (!tbody) return;

    if (
      !state.sheets[
        "Curtailment records"
      ]
    ) {
      tbody.innerHTML =
        `<tr>
          <td colspan="3">
            Curtailment records worksheet missing
          </td>
        </tr>`;

      return;
    }

    const daily =
      state.data.curtailmentDaily;

    if (!daily.length) {
      tbody.innerHTML =
        `<tr>
          <td colspan="3">
            No data found
          </td>
        </tr>`;

      return;
    }

    tbody.innerHTML =
      daily
        .map(
          record => `
            <tr>
              <td>
                ${escapeHtml(
                  formatDate(
                    record.date
                  )
                )}
              </td>

              <td>
                ${record.loss.toFixed(2)}
              </td>

              <td>
                ${record.intervalCount}
              </td>
            </tr>
          `
        )
        .join("");
  }

  /* ==========================================================
     CURTAILMENT TREND
     ========================================================== */

  function renderCurtailmentTrend() {
    if (
      !state.sheets[
        "Curtailment records"
      ]
    ) {
      showNoData(
        "curtailmentTrendWrap",
        "Curtailment records worksheet missing"
      );

      return;
    }

    const daily =
      state.data.curtailmentDaily;

    if (!daily.length) {
      showNoData(
        "curtailmentTrendWrap"
      );

      return;
    }

    const maximum =
      Math.max(
        ...daily.map(
          record =>
            record.loss
        )
      );

    createLineChart({
      key:
        "curtailmentTrend",

      wrapperId:
        "curtailmentTrendWrap",

      labels:
        daily.map(
          record =>
            formatDate(
              record.date
            )
        ),

      values:
        daily.map(
          record =>
            record.loss
        ),

      datasetLabel:
        "Curtailment Loss (MWh)",

      yMin: 0,

      yMax: Math.max(
        1,
        Math.ceil(
          maximum * 1.15
        )
      ),

      yTitle:
        "Curtailment Loss (MWh)",

      pixelsPerPoint: 70
    });
  }

  /* ==========================================================
     CURTAILMENT DURATION GANTT
     ========================================================== */

  function renderCurtailmentGantt() {
    destroyChart(
      "curtailmentGantt"
    );

    if (
      !state.sheets[
        "Curtailment records"
      ]
    ) {
      showNoData(
        "curtailmentGanttWrap",
        "Curtailment records worksheet missing"
      );

      return;
    }

    const intervals =
      state.data.curtailmentIntervals;

    if (!intervals.length) {
      showNoData(
        "curtailmentGanttWrap"
      );

      return;
    }

    const dateKeys = [
      ...new Set(
        intervals.map(
          interval =>
            interval.key
        )
      )
    ];

    const dateByKey =
      new Map();

    intervals.forEach(
      interval => {
        dateByKey.set(
          interval.key,
          interval.date
        );
      }
    );

    const wrapper =
      document.getElementById(
        "curtailmentGanttWrap"
      );

    if (!wrapper) return;

    /*
      06:00–18:00 fixed internal
      chart period.

      The outer CSS wrapper should
      provide horizontal scrolling.
    */

    const width =
      1200;

    const height =
      Math.max(
        390,
        dateKeys.length * 25
      );

    wrapper.style.width =
      `${width}px`;

    wrapper.style.height =
      `${height}px`;

    wrapper.innerHTML = "";

    const canvas =
      document.createElement(
        "canvas"
      );

    canvas.width =
      width;

    canvas.height =
      height;

    wrapper.appendChild(
      canvas
    );

    /*
      Every curtailment interval
      remains an individual bar.
      Intervals on the same date
      are NOT merged.
    */

    const data =
      intervals.map(
        interval => ({
          x: [
            interval.start,
            interval.end
          ],

          y: interval.key,

          date:
            interval.date,

          start:
            interval.start,

          end:
            interval.end,

          duration:
            interval.duration,

          loss:
            interval.loss
        })
      );

    try {
      state.charts
        .curtailmentGantt =
        new Chart(
          canvas,
          {
            type: "bar",

            data: {
              labels:
                dateKeys.map(
                  key =>
                    formatDate(
                      dateByKey.get(
                        key
                      )
                    )
                ),

              datasets: [{
                label:
                  "Curtailment interval",

                data,

                parsing: false,

                borderWidth: 8,

                borderSkipped:
                  false,

                pointRadius: 0
              }]
            },

            options: {
              ...commonChartOptions(),

              indexAxis: "y",

              scales: {
                x: {
                  type:
                    "linear",

                  /*
                    Exact requested
                    visible range:
                    06:00–18:00
                  */

                  min: 360,

                  max: 1080,

                  ticks: {
                    stepSize: 15,

                    font: {
                      size: 9
                    },

                    callback(value) {
                      return minutesToTime(
                        value
                      );
                    }
                  },

                  title: {
                    display: true,

                    text:
                      "Time",

                    font: {
                      size: 10,

                      weight: "700"
                    }
                  }
                },

                y: {
                  type:
                    "category",

                  ticks: {
                    font: {
                      size: 9
                    }
                  }
                }
              },

              plugins: {
                legend: {
                  display: false
                },

                tooltip: {
                  callbacks: {
                    title(items) {
                      const raw =
                        items[0]?.raw;

                      return raw?.date
                        ? formatDate(
                            raw.date
                          )
                        : "";
                    },

                    label(item) {
                      const raw =
                        item.raw;

                      const loss =
                        raw.loss ===
                          null ||
                        raw.loss ===
                          undefined
                          ? "Not available"
                          : `${raw.loss.toFixed(
                              2
                            )} MWh`;

                      return [
                        `Start: ${minutesToTime(
                          raw.start
                        )}`,

                        `End: ${minutesToTime(
                          raw.end
                        )}`,

                        `Duration: ${raw.duration.toFixed(
                          0
                        )} min`,

                        `Loss of Generation: ${loss}`
                      ];
                    }
                  }
                }
              }
            }
          }
        );
    } catch (error) {
      console.error(
        "Curtailment Gantt failed:",
        error
      );

      showNoData(
        "curtailmentGanttWrap",
        "Unable to render this chart."
      );
    }
  }

  /* ==========================================================
     ENERGY
     ========================================================== */

  function renderEnergy() {
    if (
      !state.sheets["Annual_KPI"]
    ) {
      setText(
        "totalBudget",
        "—"
      );

      setText(
        "totalMeasured",
        "—"
      );

      setText(
        "energyVariance",
        "—"
      );

      showNoData(
        "energyWrap",
        "Annual_KPI worksheet missing"
      );

      return;
    }

    const annual =
      state.data.annual;

    const budgeted =
      annual.map(
        record =>
          record.budgetedEnergy
      );

    const measured =
      annual.map(
        record =>
          record.measuredEnergy
      );

    const totalBudgeted =
      budgeted.reduce(
        (sum, value) =>
          sum +
          (
            value === null
              ? 0
              : value
          ),
        0
      );

    const totalMeasured =
      measured.reduce(
        (sum, value) =>
          sum +
          (
            value === null
              ? 0
              : value
          ),
        0
      );

    /*
      Variance:
      Measured - Budgeted
    */

    const variance =
      totalMeasured -
      totalBudgeted;

    setText(
      "totalBudget",
      `${formatNumber(
        totalBudgeted
      )} MWh`
    );

    setText(
      "totalMeasured",
      `${formatNumber(
        totalMeasured
      )} MWh`
    );

    setText(
      "energyVariance",
      `${formatNumber(
        variance
      )} MWh`
    );

    const allValues = [
      ...budgeted,
      ...measured
    ].filter(
      value =>
        value !== null
    );

    if (!allValues.length) {
      showNoData(
        "energyWrap"
      );

      return;
    }

    const maximum =
      Math.max(
        ...allValues
      );

    createBarChart({
      key:
        "energy",

      wrapperId:
        "energyWrap",

      labels:
        annual.map(
          record =>
            record.shortMonth
        ),

      values:
        budgeted,

      datasetLabel:
        "Budgeted Energy",

      secondValues:
        measured,

      secondDatasetLabel:
        "Measured Energy",

      indexAxis:
        "x",

      yMin: 0,

      yMax:
        Math.ceil(
          maximum * 1.15
        ),

      yTitle:
        "Energy (MWh)",

      pixelsPerPoint: 90
    });
  }

  /* ==========================================================
     MASTER RENDER
     ========================================================== */

  function renderAll() {
    try {
      renderDashboard();
    } catch (error) {
      console.error(
        "Dashboard render error:",
        error
      );
    }

    try {
      renderPAAnalysis();
    } catch (error) {
      console.error(
        "PA render error:",
        error
      );
    }

    try {
      renderPerformance();
    } catch (error) {
      console.error(
        "Performance render error:",
        error
      );
    }

    try {
      renderCurtailment();
    } catch (error) {
      console.error(
        "Curtailment render error:",
        error
      );
    }

    try {
      renderEnergy();
    } catch (error) {
      console.error(
        "Energy render error:",
        error
      );
    }
  }

  /* ==========================================================
     DOM / FORMATTING HELPERS
     ========================================================== */

  function setText(
    id,
    value
  ) {
    const element =
      document.getElementById(
        id
      );

    if (element) {
      element.textContent =
        value;
    }
  }

  function formatNumber(
    value
  ) {
    return Number(
      value
    ).toLocaleString(
      "en-IN",
      {
        maximumFractionDigits: 2
      }
    );
  }

  function escapeHtml(
    value
  ) {
    return String(value)
      .replaceAll(
        "&",
        "&amp;"
      )
      .replaceAll(
        "<",
        "&lt;"
      )
      .replaceAll(
        ">",
        "&gt;"
      )
      .replaceAll(
        '"',
        "&quot;"
      )
      .replaceAll(
        "'",
        "&#039;"
      );
  }

  /* ==========================================================
     OPTIONAL DEBUG ACCESS
     ========================================================== */

  window.SolarDGR = {
    getData() {
      return {
        daily:
          state.data.daily,

        paEvents:
          state.data.paEvents,

        paBreakdown:
          state.data.paBreakdown,

        paLoss:
          state.data.paLoss,

        annual:
          state.data.annual,

        curtailmentIntervals:
          state.data
            .curtailmentIntervals,

        curtailmentDaily:
          state.data
            .curtailmentDaily
      };
    }
  };

})();
