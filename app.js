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
     H = From Time
     I = To Time
     J = Duration
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

  document.addEventListener(
    "DOMContentLoaded",
    init
  );


  function init() {
    bindUpload();

    bindNavigation();

    renderSheetBadges();
  }


  function bindUpload() {

    const uploadButton =
      document.getElementById(
        "uploadBtn"
      );

    const fileInput =
      document.getElementById(
        "excelFile"
      );

    if (
      !uploadButton ||
      !fileInput
    ) {
      return;
    }

    uploadButton.addEventListener(
      "click",
      () => {
        fileInput.click();
      }
    );

    fileInput.addEventListener(
      "change",
      async event => {

        const file =
          event.target.files?.[0];

        if (!file) {
          return;
        }

        setText(
          "fileName",
          file.name
        );

        setText(
          "sidebarWorkbook",
          file.name
        );

        try {

          await loadWorkbook(file);

        } catch (error) {

          console.error(
            "DGR workbook load error:",
            error
          );

          alert(
            "The workbook could not be read. Please upload a valid Excel workbook."
          );
        }
      }
    );
  }


  function bindNavigation() {

    document
      .querySelectorAll(".nav button")
      .forEach(button => {

        button.addEventListener(
          "click",
          () => {

            const page =
              button.dataset.page;

            if (!page) {
              return;
            }

            document
              .querySelectorAll(
                ".nav button"
              )
              .forEach(item => {
                item.classList.remove(
                  "active"
                );
              });

            document
              .querySelectorAll(
                ".page"
              )
              .forEach(item => {
                item.classList.remove(
                  "active"
                );
              });

            button.classList.add(
              "active"
            );

            const section =
              document.getElementById(
                `page-${page}`
              );

            if (section) {
              section.classList.add(
                "active"
              );
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

            const [
              title,
              subtitle
            ] =
              metadata[page] || [
                "Solar DGR Analytics",
                ""
              ];

            setText(
              "pageTitle",
              title
            );

            setText(
              "pageSubtitle",
              subtitle
            );
          }
        );
      });
  }


  /* ==========================================================
     WORKBOOK LOADING
     ========================================================== */

  async function loadWorkbook(file) {

    if (
      typeof XLSX ===
      "undefined"
    ) {
      throw new Error(
        "SheetJS/XLSX is not available."
      );
    }

    const buffer =
      await file.arrayBuffer();

    const workbook =
      XLSX.read(
        buffer,
        {
          type: "array",
          cellDates: true,
          cellFormula: true,
          cellNF: true,
          cellText: true
        }
      );

    state.workbook =
      workbook;

    state.sheets =
      resolveSheets(workbook);

    renderSheetBadges();

    extractAllData();

    renderAll();
  }


  function normalizeSheetName(
    name
  ) {

    return String(
      name || ""
    )
      .trim()
      .toLowerCase()
      .replace(
        /[\s_-]+/g,
        ""
      );
  }


  function resolveSheets(
    workbook
  ) {

    const sheets = {};

    for (
      const logicalName of
      REQUIRED_SHEETS
    ) {

      const wanted =
        normalizeSheetName(
          logicalName
        );

      const actualName =
        workbook.SheetNames.find(
          name =>
            normalizeSheetName(
              name
            ) === wanted
        );

      sheets[logicalName] =
        actualName
          ? workbook.Sheets[
              actualName
            ]
          : null;
    }

    return sheets;
  }


  function renderSheetBadges() {

    const container =
      document.getElementById(
        "sheetStatus"
      );

    if (!container) {
      return;
    }

    container.innerHTML = "";

    REQUIRED_SHEETS.forEach(
      logicalName => {

        const exists =
          !!state.sheets[
            logicalName
          ];

        const badge =
          document.createElement(
            "div"
          );

        badge.className =
          `sheet-badge ${
            exists
              ? "ok"
              : "missing"
          }`;

        badge.innerHTML =
          `<span class="dot"></span>` +
          escapeHtml(
            logicalName
          ) +
          (
            exists
              ? ""
              : " · missing"
          );

        container.appendChild(
          badge
        );
      }
    );
  }


  /* ==========================================================
     EXCEL CELL ACCESS
     ========================================================== */

  function getCell(
    sheet,
    column,
    row
  ) {

    if (!sheet) {
      return null;
    }

    const address =
      `${column}${row}`;

    return (
      sheet[address] ||
      null
    );
  }


  function getDisplayedValue(
    sheet,
    column,
    row
  ) {

    const cell =
      getCell(
        sheet,
        column,
        row
      );

    if (!cell) {
      return null;
    }

    if (
      cell.w !== undefined &&
      cell.w !== null &&
      String(
        cell.w
      ).trim() !== ""
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


  function getRawCell(
    sheet,
    column,
    row
  ) {

    return getCell(
      sheet,
      column,
      row
    );
  }


  function getSheetRowCount(
    sheet
  ) {

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

    return (
      range.e.r + 1
    );
  }


  /* ==========================================================
     VALIDATION + NUMBER HELPERS
     ========================================================== */

  function isInvalid(
    value
  ) {

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


  function toNumber(
    value
  ) {

    if (
      isInvalid(value)
    ) {
      return null;
    }

    if (
      typeof value ===
      "number"
    ) {

      return Number.isFinite(
        value
      )
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

    return Number.isFinite(
      number
    )
      ? number
      : null;
  }


  function toPercentage(
    value
  ) {

    const number =
      toNumber(value);

    if (
      number === null
    ) {
      return null;
    }

    return Math.abs(
      number
    ) <= 1.5
      ? number * 100
      : number;
  }


  /* ==========================================================
     DATE HELPERS
     ========================================================== */

  function excelSerialToDate(
    serial
  ) {

    if (
      !Number.isFinite(
        serial
      )
    ) {
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


  function parseDate(
    value
  ) {

    if (
      isInvalid(value)
    ) {
      return null;
    }

    if (
      value instanceof Date
    ) {

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

    if (
      typeof value ===
      "number"
    ) {
      return excelSerialToDate(
        value
      );
    }

    const text =
      String(value).trim();

    if (
      /^\d{1,2}:\d{2}(:\d{2})?$/.test(
        text
      )
    ) {
      return null;
    }

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
      date.getFullYear() ===
        year &&
      date.getMonth() ===
        month - 1 &&
      date.getDate() ===
        day
    );
  }


  function dateKey(
    date
  ) {

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


  function addDays(
    date,
    days
  ) {

    const result =
      new Date(date);

    result.setDate(
      result.getDate() +
      days
    );

    return result;
  }


  function formatDate(
    date
  ) {

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

  function parseTimeMinutes(
    value
  ) {

    if (
      isInvalid(value)
    ) {
      return null;
    }

    if (
      value instanceof Date
    ) {

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
      typeof value ===
      "number" &&
      Number.isFinite(value)
    ) {

      /*
        Excel time fraction.
        0.5 = 12:00
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

    if (!match) {
      return null;
    }

    const hours =
      Number(match[1]);

    const minutes =
      Number(match[2]);

    const seconds =
      Number(
        match[3] || 0
      );

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


  function minutesToTime(
    minutes
  ) {

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
      state.sheets[
        "Daily_KPI"
      ];

    if (!sheet) {
      return [];
    }

    const lastRow =
      getSheetRowCount(sheet);

    const records = [];

    let previousDate =
      null;

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

      if (
        dateCell &&
        dateCell.f &&
        previousDate &&
        /B\d+\s*\+\s*1/i.test(
          String(
            dateCell.f
          )
        )
      ) {

        const parsed =
          date;

        if (
          !parsed ||
          dateKey(parsed) ===
            dateKey(
              previousDate
            )
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

      const validMetricCount =
        [
          pa,
          pr,
          systemLoss,
          operatingHours
        ].filter(
          value =>
            value !== null
        ).length;

      if (
        validMetricCount === 0
      ) {
        continue;
      }

      records.push({
        date,
        key:
          dateKey(date),
        pa,
        pr,
        systemLoss,
        operatingHours,
        validMetricCount,
        sourceRow: row
      });
    }

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

      if (
        end < start
      ) {
        end += 1440;
      }

      events.push({

        issue:
          String(issue).trim(),

        start,

        end,

        duration:
          end - start,

        sourceRow:
          row
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
      state.sheets[
        "Annual_KPI"
      ];

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
            SHORT_MONTHS[
              index
            ],

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

          sourceRow:
            row
        };
      }
    );
  }


  /* ==========================================================
     CURTAILMENT RECORDS

     EXACT MAPPING:

     C = DATE
     H = FROM TIME
     I = TO TIME
     J = DURATION
     R = LOSS OF GENERATION MWh
     ========================================================== */

  function readCurtailment() {

    const sheet =
      state.sheets[
        "Curtailment records"
      ];

    if (!sheet) {

      console.error(
        "Curtailment records worksheet not found."
      );

      return {
        daily: [],
        intervals: []
      };
    }

    const lastRow =
      getSheetRowCount(sheet);

    const dailyMap =
      new Map();

    const intervals = [];


    /*
      Start at row 2 because row 1
      contains the worksheet headers.
    */

    for (
      let row = 2;
      row <= lastRow;
      row++
    ) {

      /* ------------------------------------------------------
         C = DATE
         ------------------------------------------------------ */

      const rawDate =
        getDisplayedValue(
          sheet,
          "C",
          row
        );

      const date =
        parseDate(
          rawDate
        );

      if (!date) {
        continue;
      }

      const key =
        dateKey(date);


      /* ------------------------------------------------------
         H = FROM TIME
         ------------------------------------------------------ */

      const rawStart =
        getDisplayedValue(
          sheet,
          "H",
          row
        );

      const start =
        parseTimeMinutes(
          rawStart
        );


      /* ------------------------------------------------------
         I = TO TIME
         ------------------------------------------------------ */

      const rawEnd =
        getDisplayedValue(
          sheet,
          "I",
          row
        );

      const originalEnd =
        parseTimeMinutes(
          rawEnd
        );


      /* ------------------------------------------------------
         J = DURATION
         ------------------------------------------------------ */

      const rawDuration =
        getDisplayedValue(
          sheet,
          "J",
          row
        );

      const duration =
        toNumber(
          rawDuration
        );


      /* ------------------------------------------------------
         R = LOSS OF GENERATION MWh
         ------------------------------------------------------ */

      const rawLoss =
        getDisplayedValue(
          sheet,
          "R",
          row
        );

      const loss =
        toNumber(
          rawLoss
        );


      /* ------------------------------------------------------
         DAILY CURTAILMENT LOSS

         Every R value belonging to
         the same C date is summed.

         R remains MWh.

         It is NOT converted to %.
         ------------------------------------------------------ */

      if (
        loss !== null
      ) {

        if (
          !dailyMap.has(key)
        ) {

          dailyMap.set(
            key,
            {
              date,
              loss: 0,
              intervalCount: 0
            }
          );
        }

        const daily =
          dailyMap.get(key);

        daily.loss +=
          loss;

        daily.intervalCount +=
          1;
      }


      /* ------------------------------------------------------
         CURTAILMENT DURATION TABLE

         Every Excel record remains
         an individual record.

         H = From
         I = To
         J = Duration

         No merging.
         ------------------------------------------------------ */

      if (
        start !== null &&
        originalEnd !== null
      ) {

        let end =
          originalEnd;

        /*
          Overnight interval:
          23:30 → 01:00

          Internally:
          23:30 → 25:00
        */

        if (
          end < start
        ) {
          end += 1440;
        }


        /*
          Display duration from J.

          J may be:

          75
          = 75 minutes

          or

          0.0520833
          = Excel time fraction
          = 75 minutes
        */

        let durationMinutes =
          duration;

        if (
          durationMinutes !==
            null &&
          durationMinutes >= 0 &&
          durationMinutes < 1
        ) {

          durationMinutes *=
            1440;
        }


        /*
          If J is blank, only then
          calculate duration from H-I.
        */

        if (
          durationMinutes ===
          null
        ) {

          durationMinutes =
            end - start;
        }


        let durationText =
          "—";

        if (
          Number.isFinite(
            durationMinutes
          )
        ) {

          const rounded =
            Math.round(
              durationMinutes
            );

          const hours =
            Math.floor(
              rounded / 60
            );

          const minutes =
            rounded % 60;

          if (
            hours > 0
          ) {

            durationText =
              `${hours}h ${minutes}m`;

          } else {

            durationText =
              `${minutes} min`;
          }
        }


        intervals.push({

          date,

          key,

          start,

          end,

          duration:
            durationMinutes,

          durationText,

          loss,

          sourceRow:
            row
        });
      }
    }


    /* --------------------------------------------------------
       DAILY LOSS SORT
       -------------------------------------------------------- */

    const daily =
      Array.from(
        dailyMap.values()
      ).sort(
        (a, b) =>
          a.date - b.date
      );


    /* --------------------------------------------------------
       INTERVAL SORT
       -------------------------------------------------------- */

    intervals.sort(
      (a, b) => {

        const dateDifference =
          a.date - b.date;

        if (
          dateDifference !== 0
        ) {
          return dateDifference;
        }

        return (
          a.start -
          b.start
        );
      }
    );


    console.log(
      "CURTAILMENT DAILY LOSS:",
      daily
    );

    console.log(
      "CURTAILMENT INTERVALS:",
      intervals
    );


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

  function destroyChart(
    key
  ) {

    if (
      state.charts[key]
    ) {

      state.charts[
        key
      ].destroy();

      delete state.charts[
        key
      ];
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

    if (!wrapper) {
      return;
    }

    destroyChartByCanvas(
      wrapper
    );

    wrapper.innerHTML =
      "";

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

    if (!canvas) {
      return;
    }

    Object.keys(
      state.charts
    ).forEach(
      key => {

        const chart =
          state.charts[key];

        if (
          chart &&
          chart.canvas ===
            canvas
        ) {

          chart.destroy();

          delete state.charts[
            key
          ];
        }
      }
    );
  }


  function chartPixelWidth(
    count,
    pixelsPerPoint = 76
  ) {

    return Math.max(
      760,
      Math.max(
        1,
        count
      ) *
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

    wrapper.innerHTML =
      "";

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

  function createLineChart(
    config
  ) {

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

    if (
      !values.length
    ) {

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

    if (!canvas) {
      return;
    }

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

                data:
                  values,

                borderWidth: 2,

                pointRadius:
                  2.5,

                pointHoverRadius:
                  4,

                tension:
                  0.22,

                fill: false
              }]
            },

            options: {

              ...commonChartOptions(),

              scales: {

                x: {

                  ticks: {

                    autoSkip:
                      false,

                    maxRotation:
                      45,

                    minRotation:
                      45,

                    font: {
                      size: 9
                    },

                    callback(
                      value,
                      index
                    ) {

                      return (
                        index % 2 ===
                        0
                      )
                        ? this.getLabelForValue(
                            value
                          )
                        : "";
                    }
                  },

                  grid: {
                    display:
                      false
                  }
                },

                y: {

                  min:
                    yMin,

                  max:
                    yMax,

                  ticks: {

                    stepSize:
                      yStep,

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
                      true,

                    text:
                      yTitle,

                    font: {

                      size: 10,

                      weight:
                        "700"
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

  function createBarChart(
    config
  ) {

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

    if (!canvas) {
      return;
    }

    const datasets = [{

      label:
        datasetLabel,

      data:
        values,

      borderWidth: 1,

      borderRadius: 3
    }];


    if (
      secondValues
    ) {

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

                  min:
                    xMin,

                  max:
                    xMax,

                  ticks: {

                    stepSize:
                      xStep,

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

                    text:
                      xTitle,

                    font: {

                      size: 10,

                      weight:
                        "700"
                    }
                  }
                },

                y: {

                  min:
                    yMin,

                  max:
                    yMax,

                  ticks: {

                    stepSize:
                      yStep,

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

                    text:
                      yTitle,

                    font: {

                      size: 10,

                      weight:
                        "700"
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
      !state.sheets[
        "Daily_KPI"
      ] ||
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
        state.sheets[
          "Daily_KPI"
        ]
          ? "No data found"
          : "Daily_KPI worksheet missing"
      );

      showNoData(
        "dashLossChartWrap",
        state.sheets[
          "Daily_KPI"
        ]
          ? "No data found"
          : "Daily_KPI worksheet missing"
      );

      setText(
        "dashRange",
        "No Daily_KPI data"
      );

      return;
    }

    const latest =
      data[
        data.length - 1
      ];

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
      latest.operatingHours ===
        null
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

      key:
        "dashPR",

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

      percentage:
        true
    });


    const losses =
      data.filter(
        record =>
          record.systemLoss !==
          null
      );

    createLineChart({

      key:
        "dashLoss",

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

      /*
        Exact requested range:
        0%, 0.5%, 1%, 1.5%,
        2%, 2.5%, 3%
      */

      yMin: 0,

      yMax: 3,

      yStep: 0.5,

      yTitle:
        "System Loss (%)",

      percentage:
        true
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
      !state.sheets[
        "Daily_KPI"
      ]
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

      yMin:
        80,

      yMax:
        100,

      yStep:
        5,

      yTitle:
        "Plant Availability (%)",

      percentage:
        true
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

    if (
      !events.length
    ) {

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

    if (!wrapper) {
      return;
    }

    const width =
      Math.max(
        1150,
        issues.length * 125
      );

    wrapper.style.width =
      `${width}px`;

    wrapper.style.height =
      "390px";

    wrapper.innerHTML =
      "";

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


    const datasets =
      issues.map(
        issue => ({

          label:
            issue,

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

                  y:
                    issue,

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

          parsing:
            false,

          borderWidth:
            8,

          borderSkipped:
            false,

          pointRadius:
            0
        })
      );


    try {

      state.charts.paGantt =
        new Chart(
          canvas,
          {

            type:
              "bar",

            data: {

              labels:
                issues,

              datasets
            },

            options: {

              ...commonChartOptions(),

              indexAxis:
                "y",

              scales: {

                x: {

                  type:
                    "linear",

                  min:
                    0,

                  max:
                    1440,

                  ticks: {

                    stepSize:
                      30,

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

                    display:
                      true,

                    text:
                      "Time of day",

                    font: {

                      size: 10,

                      weight:
                        "700"
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
                  display:
                    false
                },

                tooltip: {

                  callbacks: {

                    title(items) {

                      return (
                        items[0]
                          ?.raw
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

    if (
      !data.length
    ) {

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

      xMin:
        0,

      xMax:
        13,

      xStep:
        1,

      xTitle:
        "Breakdown Time (minutes)",

      pixelsPerPoint:
        20
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

    if (
      !data.length
    ) {

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
      Math.max(
        ...values
      );

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

      yMin:
        0,

      yMax:
        Math.max(
          1,
          Math.ceil(
            maximum * 1.15
          )
        ),

      yTitle:
        "System Loss (MWh)",

      pixelsPerPoint:
        60
    });
  }


  /* ==========================================================
     PERFORMANCE
     ========================================================== */

  function renderPerformance() {

    const daily =
      state.data.daily;

    if (
      !state.sheets[
        "Daily_KPI"
      ]
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

        yMin:
          0,

        yMax:
          100,

        yStep:
          20,

        yTitle:
          "PR (%)",

        percentage:
          true
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

        yMin:
          0,

        yMax:
          Math.max(
            24,
            Math.ceil(
              maxHours
            )
          ),

        yStep:
          2,

        yTitle:
          "Operating Hours"
      });


      const losses =
        daily.filter(
          record =>
            record.systemLoss !==
            null
        );

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

        /*
          Exact requested range:
          0 to 3%
          0.5% ticks
        */

        yMin:
          0,

        yMax:
          3,

        yStep:
          0.5,

        yTitle:
          "System Loss (%)",

        percentage:
          true
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
      !state.sheets[
        "Annual_KPI"
      ]
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

      xMin:
        0,

      xMax:
        100,

      xStep:
        20,

      xTitle:
        "Performance Ratio (%)",

      percentage:
        true,

      pixelsPerPoint:
        20
    });
  }


  /* ==========================================================
     CURTAILMENT
     ========================================================== */

    /* ==========================================================
     CURTAILMENT
     ========================================================== */

  function renderCurtailment() {

    renderCurtailmentKPI();

    renderCurtailmentTable();

    renderCurtailmentTrend();

    renderCurtailmentGantt();
  }


  /* ==========================================================
     CURTAILMENT KPI
     ========================================================== */

  function renderCurtailmentKPI() {

    const data =
      state.data.curtailmentDaily;

    if (
      !state.sheets[
        "Curtailment records"
      ]
    ) {
      return;
    }

    if (
      !data.length
    ) {
      setText(
        "curtailmentTotal",
        "—"
      );

      setText(
        "curtailmentIntervals",
        "—"
      );

      return;
    }

    const totalLoss =
      data.reduce(
        (sum, record) =>
          sum +
          (
            Number.isFinite(
              record.loss
            )
              ? record.loss
              : 0
          ),
        0
      );

    const totalIntervals =
      data.reduce(
        (sum, record) =>
          sum +
          (
            Number.isFinite(
              record.intervalCount
            )
              ? record.intervalCount
              : 0
          ),
        0
      );

    setText(
      "curtailmentTotal",
      `${totalLoss.toFixed(2)} MWh`
    );

    setText(
      "curtailmentIntervals",
      String(totalIntervals)
    );
  }


  /* ==========================================================
     CURTAILMENT TABLE
     ========================================================== */

  function renderCurtailmentTable() {

    const data =
      state.data.curtailmentIntervals;

    const container =
      document.getElementById(
        "curtailmentTableWrap"
      );

    if (!container) {
      return;
    }

    if (
      !state.sheets[
        "Curtailment records"
      ]
    ) {

      container.innerHTML =
        `<div class="no-data">
          Curtailment records worksheet missing
        </div>`;

      return;
    }

    if (
      !data.length
    ) {

      container.innerHTML =
        `<div class="no-data">
          No curtailment data found
        </div>`;

      return;
    }

    let html = `
      <div class="table-scroll">
        <table class="data-table">

          <thead>
            <tr>
              <th>Date</th>
              <th>From</th>
              <th>To</th>
              <th>Duration</th>
              <th>Loss of Generation (MWh)</th>
            </tr>
          </thead>

          <tbody>
    `;

    data.forEach(
      record => {

        html += `
          <tr>

            <td>
              ${escapeHtml(
                formatDate(
                  record.date
                )
              )}
            </td>

            <td>
              ${escapeHtml(
                minutesToTime(
                  record.start
                )
              )}
            </td>

            <td>
              ${escapeHtml(
                minutesToTime(
                  record.end
                )
              )}
            </td>

            <td>
              ${escapeHtml(
                record.durationText ||
                "—"
              )}
            </td>

            <td>
              ${
                record.loss === null
                  ? "—"
                  : `${record.loss.toFixed(2)}`
              }
            </td>

          </tr>
        `;
      }
    );

    html += `
          </tbody>
        </table>
      </div>
    `;

    container.innerHTML =
      html;
  }


  /* ==========================================================
     CURTAILMENT DAILY LOSS TREND
     ========================================================== */

  function renderCurtailmentTrend() {

    const data =
      state.data.curtailmentDaily;

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

    if (
      !data.length
    ) {

      showNoData(
        "curtailmentTrendWrap"
      );

      return;
    }

    createLineChart({

      key:
        "curtailmentTrend",

      wrapperId:
        "curtailmentTrendWrap",

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
            record.loss
        ),

      datasetLabel:
        "Loss of Generation (MWh)",

      yMin:
        0,

      yMax:
        Math.max(
          1,
          Math.ceil(
            Math.max(
              ...data.map(
                record =>
                  record.loss || 0
              )
            ) * 1.15
          )
        ),

      yTitle:
        "Loss of Generation (MWh)",

      pixelsPerPoint:
        76
    });
  }


  /* ==========================================================
     CURTAILMENT GANTT
     
     Uses:
       C = Date
       H = From
       I = To

     Displays daytime curtailment:
       06:00 → 18:00

     Each Excel record remains an individual interval.
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

    const source =
      state.data.curtailmentIntervals;

    if (
      !source.length
    ) {

      showNoData(
        "curtailmentGanttWrap"
      );

      return;
    }


    /*
      Daytime window:
      06:00 = 360 minutes
      18:00 = 1080 minutes
    */

    const DAY_START =
      360;

    const DAY_END =
      1080;


    /*
      Keep only intervals that
      overlap the daytime window.
    */

    const intervals =
      source
        .map(
          record => {

            const clippedStart =
              Math.max(
                record.start,
                DAY_START
              );

            const clippedEnd =
              Math.min(
                record.end,
                DAY_END
              );

            if (
              clippedEnd <=
              clippedStart
            ) {
              return null;
            }

            return {

              ...record,

              chartStart:
                clippedStart,

              chartEnd:
                clippedEnd
            };
          }
        )
        .filter(
          record =>
            record !== null
        );


    if (
      !intervals.length
    ) {

      showNoData(
        "curtailmentGanttWrap",
        "No daytime curtailment intervals found"
      );

      return;
    }


    /*
      Use complete date keys so
      different years do not collide.
    */

    const uniqueDates = [
      ...new Set(
        intervals.map(
          record =>
            record.key
        )
      )
    ];


    const dateLabels =
      uniqueDates.map(
        key => {

          const record =
            intervals.find(
              item =>
                item.key === key
            );

          return formatDate(
            record.date
          );
        }
      );


    const wrapper =
      document.getElementById(
        "curtailmentGanttWrap"
      );

    if (!wrapper) {
      return;
    }


    const width =
      Math.max(
        1150,
        uniqueDates.length *
          85
      );

    const height =
      Math.max(
        330,
        Math.min(
          700,
          uniqueDates.length *
            32 +
            100
        )
      );

    wrapper.style.width =
      `${width}px`;

    wrapper.style.height =
      `${height}px`;

    wrapper.innerHTML =
      "";


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


    const datasets =
      intervals.map(
        (record, index) => ({

          label:
            `Curtailment ${index + 1}`,

          data: [{

            x: [
              record.chartStart,
              record.chartEnd
            ],

            y:
              record.key,

            start:
              record.start,

            end:
              record.end,

            chartStart:
              record.chartStart,

            chartEnd:
              record.chartEnd,

            duration:
              record.duration,

            durationText:
              record.durationText,

            loss:
              record.loss,

            sourceRow:
              record.sourceRow
          }],

          parsing:
            false,

          borderWidth:
            10,

          borderSkipped:
            false,

          pointRadius:
            0
        })
      );


    try {

      state.charts[
        "curtailmentGantt"
      ] =
        new Chart(
          canvas,
          {

            type:
              "bar",

            data: {

              labels:
                uniqueDates,

              datasets
            },

            options: {

              ...commonChartOptions(),

              indexAxis:
                "y",

              scales: {

                x: {

                  type:
                    "linear",

                  min:
                    DAY_START,

                  max:
                    DAY_END,

                  ticks: {

                    stepSize:
                      30,

                    font: {
                      size: 9
                    },

                    callback(
                      value
                    ) {

                      return minutesToTime(
                        value
                      );
                    }
                  },

                  title: {

                    display:
                      true,

                    text:
                      "Time of day",

                    font: {

                      size: 10,

                      weight:
                        "700"
                    }
                  }
                },

                y: {

                  type:
                    "category",

                  labels:
                    uniqueDates,

                  ticks: {

                    font: {
                      size: 9
                    },

                    callback(
                      value
                    ) {

                      return dateLabels[
                        value
                      ] || value;
                    }
                  }
                }
              },

              plugins: {

                legend: {
                  display:
                    false
                },

                tooltip: {

                  callbacks: {

                    title() {
                      return "Curtailment Interval";
                    },

                    label(context) {

                      const raw =
                        context.raw;

                      const lines = [

                        `Date: ${formatDate(
                          intervals[
                            context.datasetIndex
                          ].date
                        )}`,

                        `From: ${minutesToTime(
                          raw.start
                        )}`,

                        `To: ${minutesToTime(
                          raw.end
                        )}`,

                        `Duration: ${
                          raw.durationText ||
                          "—"
                        }`

                      ];

                      if (
                        raw.loss !==
                        null &&
                        raw.loss !==
                        undefined
                      ) {

                        lines.push(
                          `Loss: ${Number(
                            raw.loss
                          ).toFixed(
                            2
                          )} MWh`
                        );
                      }

                      lines.push(
                        `Excel Row: ${raw.sourceRow}`
                      );

                      return lines;
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
