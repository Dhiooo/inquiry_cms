(() => {
  /* =========================================================
     ELEMENT HELPERS & STATE
  ========================================================= */

  const $ = id => document.getElementById(id);

  const state = {
    file: null,
    rows: [],
    checked: [],
    issues: [],
    filter: "all",
    page: 1,
    stats: {
      total: 0,
      ready: 0,
      incomplete: 0,
      fatal: 0,
      add: 0
    }
  };

  const PAGE_SIZE = 5;

  const validSources = [
    "Instagram",
    "Facebook",
    "WhatsApp (WA)",
    "Walk in",
    "Web",
    "GMaps",
    "Referral",
    "Status Ads",
    "Others Admin",
    "Unknown"
  ];

  const sourceAliases = {
    refferal: "Referral",
    referral: "Referral",
    wa: "WhatsApp (WA)",
    whatsapp: "WhatsApp (WA)",
    "whatsapp (wa)": "WhatsApp (WA)",
    "others admin": "Others Admin",
    gmaps: "GMaps",
    "walk in": "Walk in",
    "status ads": "Status Ads"
  };

  const uploadZone = $("uploadZone");
  const spreadsheetInput = $("spreadsheetInput");
  const validateButton = $("validateBtn");

  /* =========================================================
     GENERAL HELPERS
  ========================================================= */

  function normalizeHeader(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function cellText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).trim();
  }

  function hasValue(value) {
    if (value === null || value === undefined || value === false) {
      return false;
    }

    if (typeof value === "string") {
      return value.trim() !== "";
    }

    return true;
  }

  function formatFileSize(bytes) {
    if (bytes < 1024) {
      return `${bytes} B`;
    }

    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(1)} KB`;
    }

    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function normalizeSource(value) {
    const originalValue = cellText(value);

    if (!originalValue) {
      return "";
    }

    const lowerValue = originalValue.toLowerCase();

    if (sourceAliases[lowerValue]) {
      return sourceAliases[lowerValue];
    }

    const validSource = validSources.find(
      source => source.toLowerCase() === lowerValue
    );

    return validSource || originalValue;
  }

  function normalizeDate(value) {
    if (!value) {
      return "";
    }

    // Jika SheetJS menghasilkan Date object
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const year = value.getFullYear();
      const month = String(value.getMonth() + 1).padStart(2, "0");
      const day = String(value.getDate()).padStart(2, "0");

      return `${year}-${month}-${day}`;
    }

    // Jika tanggal disimpan Excel sebagai serial number
    if (typeof value === "number" && window.XLSX) {
      const parsedDate = XLSX.SSF.parse_date_code(value);

      if (parsedDate) {
        const year = parsedDate.y;
        const month = String(parsedDate.m).padStart(2, "0");
        const day = String(parsedDate.d).padStart(2, "0");

        return `${year}-${month}-${day}`;
      }
    }

    // Jika tanggal berupa string
    const textValue = cellText(value);

    if (!textValue) {
      return "";
    }

    const parsedDate = new Date(textValue);

    if (Number.isNaN(parsedDate.getTime())) {
      return "";
    }

    const year = parsedDate.getFullYear();
    const month = String(parsedDate.getMonth() + 1).padStart(2, "0");
    const day = String(parsedDate.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  function findColumn(headers, possibleNames) {
    const normalizedHeaders = headers.map(normalizeHeader);

    for (const possibleName of possibleNames) {
      const index = normalizedHeaders.indexOf(
        normalizeHeader(possibleName)
      );

      if (index >= 0) {
        return index;
      }
    }

    return -1;
  }

  /* =========================================================
     UPLOAD ERROR
  ========================================================= */

  function showUploadError(message = "") {
    uploadZone.classList.toggle("has-error", Boolean(message));
    $("uploadError").textContent = message;
  }

  /* =========================================================
     RESET VALIDATION RESULT
  ========================================================= */

  function clearValidationResult() {
    state.rows = [];
    state.checked = [];
    state.issues = [];

    state.filter = "all";
    state.page = 1;

    state.stats = {
      total: 0,
      ready: 0,
      incomplete: 0,
      fatal: 0,
      add: 0
    };

    $("vresult").style.display = "none";

    $("addBtn").disabled = true;
    $("addBtn").textContent = "Add Inquiry";

    if (state.file) {
      $("fsummary").innerHTML =
        'Klik <b>Validate Inquiries</b> untuk memeriksa file.';
    } else {
      $("fsummary").innerHTML =
        "Pilih spreadsheet terlebih dahulu.";
    }
  }

  /* =========================================================
     ACCEPT SELECTED FILE
  ========================================================= */

  function acceptFile(file) {
    const extension = (file.name.split(".").pop() || "").toLowerCase();

    const supportedExtensions = ["xlsx", "xls", "csv"];

    if (!supportedExtensions.includes(extension)) {
      showUploadError(
        "Format file tidak didukung. Gunakan .xlsx, .xls, atau .csv."
      );

      return;
    }

    state.file = file;

    clearValidationResult();
    showUploadError("");

    uploadZone.classList.add("has-file");

    $("uploadTitle").textContent = file.name;

    $("uploadHelp").textContent =
      `${formatFileSize(file.size)} · siap divalidasi`;

    $("chooseFileBtn").textContent = "Ganti File";

    validateButton.disabled = false;
  }

  /* =========================================================
     READ SPREADSHEET
  ========================================================= */

  async function readSpreadsheet() {
    if (!window.XLSX) {
      throw new Error(
        "Library pembaca spreadsheet gagal dimuat. Periksa koneksi internet lalu refresh."
      );
    }

    if (!state.file) {
      throw new Error("Pilih spreadsheet terlebih dahulu.");
    }

    const fileBuffer = await state.file.arrayBuffer();

    const workbook = XLSX.read(fileBuffer, {
      type: "array",
      cellDates: true
    });

    if (!workbook.SheetNames.length) {
      throw new Error("Spreadsheet tidak memiliki sheet.");
    }

    // Membaca sheet pertama
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    const spreadsheetRows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: true,
      defval: null,
      // Pertahankan baris kosong agar nomor ROW sama dengan baris asli Excel.
      blankrows: true
    });

    if (!spreadsheetRows.length) {
      throw new Error("Spreadsheet kosong.");
    }

    /* ---------------------------------------------------------
       CARI HEADER

       Sistem mencari baris yang memiliki:
       - Phone Number
       - Phone
       - No HP
       - Nomor Telepon
    --------------------------------------------------------- */

    const headerIndex = spreadsheetRows.findIndex(row =>
      row.some(cell => {
        const header = normalizeHeader(cell);

        return [
          "phone number",
          "phone",
          "no hp",
          "nomor telepon"
        ].includes(header);
      })
    );

    if (headerIndex < 0) {
      throw new Error(
        'Header "Phone Number" tidak ditemukan.'
      );
    }

    const headers = spreadsheetRows[headerIndex];

    /* ---------------------------------------------------------
       CARI POSISI MASING-MASING KOLOM
    --------------------------------------------------------- */

    const columns = {
      date: findColumn(headers, [
        "Date",
        "Inquiry Date"
      ]),

      phone: findColumn(headers, [
        "Phone Number",
        "Phone",
        "No HP",
        "Nomor Telepon"
      ]),

      name: findColumn(headers, [
        "Name (Optional)",
        "Name",
        "Student Name"
      ]),

      source: findColumn(headers, [
        "Source Ads",
        "Source"
      ]),

      chat: findColumn(headers, [
        "Chat",
        "Status"
      ]),

      note: findColumn(headers, [
        "Notes & FUP",
        "Notes",
        "Contact Note"
      ])
    };

    if (columns.phone < 0) {
      throw new Error(
        'Kolom "Phone Number" tidak ditemukan.'
      );
    }

    /* ---------------------------------------------------------
       MEMBACA SETIAP BARIS

       Hanya kolom utama inquiry yang diperiksa:
       - Date
       - Phone Number
       - Name
       - Source Ads
       - Chat
       - Notes & FUP

       Kolom tambahan seperti:
       - Done Fup
       - Tgl Fup
       - Hasil fup
       - merged cell
       - formatting

       TIDAK membuat baris dianggap sebagai inquiry.
    --------------------------------------------------------- */

    const inquiryRows = [];

    for (
      let rowIndex = headerIndex + 1;
      rowIndex < spreadsheetRows.length;
      rowIndex++
    ) {
      const row = spreadsheetRows[rowIndex];

      /*
       * Hanya lima kolom inti yang menentukan apakah sebuah baris
       * benar-benar merupakan inquiry. Notes dan kolom FUP tidak
       * boleh membuat baris kosong ikut terhitung.
       */
      const rawDate =
        columns.date >= 0 ? row[columns.date] : null;
      const phone =
        columns.phone >= 0 ? row[columns.phone] : null;
      const rawName =
        columns.name >= 0 ? row[columns.name] : null;
      const rawSource =
        columns.source >= 0 ? row[columns.source] : null;
      const rawChat =
        columns.chat >= 0 ? row[columns.chat] : null;

      const coreValues = [
        rawDate,
        phone,
        rawName,
        rawSource,
        rawChat
      ];

      // Abaikan baris yang hanya berisi formatting, Notes/FUP,
      // checkbox, formula kosong, atau data di kolom tambahan.
      if (!coreValues.some(hasValue)) {
        continue;
      }

      /*
       * Abaikan judul section seperti "INQUIRY AUG". Judul section
       * biasanya hanya mengisi kolom Date dengan teks yang bukan tanggal.
       */
      const onlyDateIsFilled =
        hasValue(rawDate) &&
        !hasValue(phone) &&
        !hasValue(rawName) &&
        !hasValue(rawSource) &&
        !hasValue(rawChat);

      if (onlyDateIsFilled && !normalizeDate(rawDate)) {
        continue;
      }

      inquiryRows.push({
        // Nomor baris asli di spreadsheet
        r: rowIndex + 1,

        date: normalizeDate(
          columns.date >= 0
            ? row[columns.date]
            : null
        ),

        phone: phone,

        name: cellText(
          columns.name >= 0
            ? row[columns.name]
            : null
        ),

        source: normalizeSource(
          columns.source >= 0
            ? row[columns.source]
            : null
        ),

        chat: cellText(
          columns.chat >= 0
            ? row[columns.chat]
            : null
        ).toUpperCase(),

        note: cellText(
          columns.note >= 0
            ? row[columns.note]
            : null
        )
      });
    }

    if (!inquiryRows.length) {
      throw new Error(
        "Tidak ada data inquiry yang dapat dibaca."
      );
    }

    return inquiryRows;
  }

  /* =========================================================
     VALIDATE INQUIRY ROWS
  ========================================================= */

  function validateRows(rows) {
    const existingPhoneNumbers = new Map();
    const spreadsheetPhoneNumbers = new Map();

    /*
     * Masukkan nomor telepon yang sudah ada
     * pada dashboard ke daftar duplicate check.
     */
    dashRows.forEach((row, index) => {
      const phoneKey = keyFromCodePhone(
        row.code,
        row.phone
      );

      if (phoneKey) {
        existingPhoneNumbers.set(
          phoneKey,
          row.student ||
            row.parent ||
            `data CMS ${index + 1}`
        );
      }
    });

    state.checked = rows.map(rawRow => {
      const normalizedPhone = normalizePhone(
        rawRow.phone
      );

      const issues = [];
      const categories = [];

      let severity = "ready";

      function addIssue(category, message, type) {
        categories.push(category);
        issues.push([message, type]);

        if (type === "fatal") {
          severity = "fatal";
        } else if (severity === "ready") {
          severity = "warning";
        }
      }

      /* -------------------------------------------------------
         PHONE VALIDATION
      ------------------------------------------------------- */

      if (!normalizedPhone.ok) {
        addIssue(
          "phone",
          "Phone Number - belum diisi",
          "fatal"
        );
      } else {
        const phoneKey = normalizedPhone.digits;

        // Duplicate di dalam spreadsheet
        if (spreadsheetPhoneNumbers.has(phoneKey)) {
          const duplicateRow =
            spreadsheetPhoneNumbers.get(phoneKey);

          addIssue(
            "phone",
            `Phone Number - duplikat dengan baris ${duplicateRow}`,
            "fatal"
          );
        }

        // Duplicate dengan data dashboard
        else if (existingPhoneNumbers.has(phoneKey)) {
          const existingName =
            existingPhoneNumbers.get(phoneKey);

          addIssue(
            "phone",
            `Phone Number - sudah ada (${existingName})`,
            "fatal"
          );
        }

        // Nomor belum terdaftar
        else {
          spreadsheetPhoneNumbers.set(
            phoneKey,
            rawRow.r
          );
        }
      }

      /* -------------------------------------------------------
         DATE VALIDATION
      ------------------------------------------------------- */

      if (!rawRow.date) {
        addIssue(
          "date",
          "Inquiry Date - belum diisi / tidak terbaca",
          "warning"
        );
      }

      /* -------------------------------------------------------
         SOURCE VALIDATION
      ------------------------------------------------------- */

      if (!rawRow.source) {
        addIssue(
          "source",
          "Source Ads - belum diisi",
          "warning"
        );
      } else if (!validSources.includes(rawRow.source)) {
        addIssue(
          "source",
          `Source Ads - “${rawRow.source}” tidak tersedia`,
          "warning"
        );
      }

      /* -------------------------------------------------------
         CHAT STATUS VALIDATION
      ------------------------------------------------------- */

      if (!rawRow.chat) {
        addIssue(
          "chat",
          "Chat - belum diisi (default A)",
          "warning"
        );
      } else if (!STATUS[rawRow.chat]) {
        addIssue(
          "chat",
          `Chat - “${rawRow.chat}” tidak ditemukan (default A)`,
          "warning"
        );
      }

      return {
        raw: rawRow,
        phone: normalizedPhone,
        severity,
        categories: [...new Set(categories)],
        issues
      };
    });

    /* ---------------------------------------------------------
       DATA YANG DITAMPILKAN DI VALIDATION TABLE

       Ready tidak ditampilkan.
       Hanya incomplete dan fatal yang ditampilkan.
    --------------------------------------------------------- */

    state.issues = state.checked
      .filter(row => row.severity !== "ready")
      .map(row => ({
        row: row.raw.r,

        name: row.raw.name || "Tanpa nama",

        meta: row.phone.ok
          ? `${row.phone.code || ""} ${row.phone.pretty}`.trim()
          : "Nomor kosong",

        severity: row.severity,

        categories: row.categories,

        issues: row.issues
      }));

    /* ---------------------------------------------------------
       HITUNG JUMLAH
    --------------------------------------------------------- */

    state.stats.total = state.checked.length;

    state.stats.ready = state.checked.filter(
      row => row.severity === "ready"
    ).length;

    state.stats.incomplete = state.checked.filter(
      row => row.severity === "warning"
    ).length;

    state.stats.fatal = state.checked.filter(
      row => row.severity === "fatal"
    ).length;

    state.stats.add =
      state.stats.ready +
      state.stats.incomplete;
  }

  /* =========================================================
     UPDATE VALIDATION SUMMARY
  ========================================================= */

  function updateValidationSummary() {
    const summaryNumbers = document.querySelectorAll(
      ".ss-summary .ss-snum"
    );

    if (summaryNumbers.length >= 3) {
      summaryNumbers[0].textContent =
        state.stats.ready;

      summaryNumbers[1].textContent =
        state.stats.incomplete;

      summaryNumbers[2].textContent =
        state.stats.fatal;
    }

    const categoryCounts = {
      all: state.issues.length,
      phone: 0,
      date: 0,
      source: 0,
      chat: 0
    };

    state.issues.forEach(row => {
      row.categories.forEach(category => {
        if (
          Object.prototype.hasOwnProperty.call(
            categoryCounts,
            category
          )
        ) {
          categoryCounts[category]++;
        }
      });
    });

    document
      .querySelectorAll("#vresult .ss-ifilter")
      .forEach(button => {
        const countElement =
          button.querySelector(".ss-count");

        if (countElement) {
          countElement.textContent =
            categoryCounts[button.dataset.filter] || 0;
        }
      });

    const readyHidden =
      document.querySelector(".ss-readyhidden");

    if (readyHidden) {
      readyHidden.textContent =
        `${state.stats.ready} ready rows hidden`;
    }

    $("fsummary").innerHTML = `
      <b>${state.stats.add}</b> akan ditambahkan
      · ${state.stats.ready} ready
      + ${state.stats.incomplete} incomplete
      · <b style="color:var(--red)">
          ${state.stats.fatal}
        </b> fatal dilewati
      · ${state.stats.total} total terbaca
    `;

    $("addBtn").disabled =
      state.stats.add === 0;

    $("addBtn").textContent =
      `Add Inquiry (${state.stats.add})`;

    $("cfConfirm").textContent =
      `Yes, Add ${state.stats.add} Inquiries`;

          /* =====================================================
       UPDATE CONFIRMATION MODAL SECARA DINAMIS
    ===================================================== */

    const confirmationWarningTitle =
      document.querySelector(
        ".cf-issue.warning .cf-ititle"
      );

    const confirmationFatalTitle =
      document.querySelector(
        ".cf-issue.fatal .cf-ititle"
      );

    const confirmationSummaryTitle =
      document.querySelector(
        ".cf-summary .cf-stitle"
      );

    const confirmationSummarySubtitle =
      document.querySelector(
        ".cf-summary .cf-ssub"
      );

    const confirmationNote =
      document.querySelector(".cf-note");

    const confirmationSubtitle =
      document.querySelector(".cf-subtitle");

    const warningCard =
      document.querySelector(".cf-issue.warning");

    const fatalCard =
      document.querySelector(".cf-issue.fatal");

    /* Jumlah incomplete */

    if (confirmationWarningTitle) {
      confirmationWarningTitle.textContent =
        `${state.stats.incomplete} inquiry belum lengkap`;
    }

    /* Jumlah fatal */

    if (confirmationFatalTitle) {
      confirmationFatalTitle.textContent =
        `${state.stats.fatal} inquiry memiliki masalah fatal`;
    }

    /* Jumlah yang benar-benar akan ditambahkan */

    if (confirmationSummaryTitle) {
      confirmationSummaryTitle.textContent =
        `${state.stats.add} inquiries will be added`;
    }

    /* Ready + incomplete */

    if (confirmationSummarySubtitle) {
      confirmationSummarySubtitle.textContent =
        `${state.stats.ready} siap di-import + ` +
        `${state.stats.incomplete} belum lengkap`;
    }

    /* Catatan konfirmasi */

    if (confirmationNote) {
      confirmationNote.innerHTML = `
        <strong>Lanjutkan proses?</strong>
        Dengan melanjutkan,
        ${state.stats.incomplete} inquiry yang belum lengkap
        akan tetap masuk dan
        ${state.stats.fatal} inquiry fatal
        tidak akan ikut di-import.
      `;
    }

    /* Subtitle modal */

    if (confirmationSubtitle) {
      if (
        state.stats.incomplete === 0 &&
        state.stats.fatal === 0
      ) {
        confirmationSubtitle.textContent =
          "Semua inquiry siap ditambahkan.";
      } else {
        confirmationSubtitle.textContent =
          "Masih ada data yang belum lengkap atau memiliki " +
          "masalah fatal. Periksa ringkasan berikut sebelum melanjutkan.";
      }
    }

    /*
     * Sembunyikan kartu warning jika tidak ada incomplete.
     * Sembunyikan kartu fatal jika tidak ada fatal.
     */

    if (warningCard) {
      warningCard.style.display =
        state.stats.incomplete > 0
          ? ""
          : "none";
    }

    if (fatalCard) {
      fatalCard.style.display =
        state.stats.fatal > 0
          ? ""
          : "none";
    }
  }

  /* =========================================================
     FILTER VALIDATION ISSUES
  ========================================================= */

  function getFilteredIssues() {
    let rows = [...state.issues];

    // Filter kategori
    if (state.filter !== "all") {
      rows = rows.filter(row =>
        row.categories.includes(state.filter)
      );
    }

    // Filter severity
    const severity = $("ssSeverity").value;

    if (severity !== "all") {
      rows = rows.filter(
        row => row.severity === severity
      );
    }

    // Search
    const searchQuery = $("ssSearch")
      .value
      .trim()
      .toLowerCase();

    if (searchQuery) {
      rows = rows.filter(row => {
        const searchableText = [
          row.row,
          row.name,
          row.meta,
          ...row.issues.flat()
        ]
          .join(" ")
          .toLowerCase();

        return searchableText.includes(searchQuery);
      });
    }

    // Sorting
    const sort = $("ssSort").value;

    rows.sort((a, b) => {
      if (sort === "row-asc") {
        return a.row - b.row;
      }

      if (sort === "row-desc") {
        return b.row - a.row;
      }

      if (sort === "warning-first") {
        if (a.severity !== b.severity) {
          return a.severity === "warning"
            ? -1
            : 1;
        }

        return a.row - b.row;
      }

      // Default: fatal first
      if (a.severity !== b.severity) {
        return a.severity === "fatal"
          ? -1
          : 1;
      }

      return a.row - b.row;
    });

    return rows;
  }

  /* =========================================================
     RENDER VALIDATION TABLE
  ========================================================= */

  function renderValidationTable() {
    const rows = getFilteredIssues();

    const totalPages = Math.max(
      1,
      Math.ceil(rows.length / PAGE_SIZE)
    );

    state.page = Math.min(
      state.page,
      totalPages
    );

    state.page = Math.max(
      state.page,
      1
    );

    const startIndex =
      (state.page - 1) * PAGE_SIZE;

    const visibleRows = rows.slice(
      startIndex,
      startIndex + PAGE_SIZE
    );

    if (!visibleRows.length) {
      $("ssRows").innerHTML = `
        <div class="ss-empty">
          Tidak ada inquiry yang cocok.
        </div>
      `;
    } else {
      $("ssRows").innerHTML = visibleRows
        .map(row => {
          const issueChips = row.issues
            .map(([message, type]) => {
              const className =
                type === "fatal"
                  ? "cf"
                  : "cw";

              return `
                <span class="ss-chip ${className}">
                  ${esc(message)}
                </span>
              `;
            })
            .join("");

          const badgeClass =
            row.severity === "fatal"
              ? "bf"
              : "bw";

          const badgeLabel =
            row.severity === "fatal"
              ? "FATAL"
              : "WARNING";

          const resultLabel =
            row.severity === "fatal"
              ? "Terblokir"
              : "Tetap di-import";

          return `
            <div class="ss-datarow ${
              row.severity === "fatal"
                ? "fatalrow"
                : ""
            }">
              <div class="ss-cell">
                <div class="ss-rownum">
                  ${row.row}
                </div>
              </div>

              <div class="ss-cell">
                <div class="ss-person">
                  ${esc(row.name)}
                </div>

                <div class="ss-meta">
                  ${esc(row.meta)}
                </div>
              </div>

              <div class="ss-cell">
                <div class="ss-chips">
                  ${issueChips}
                </div>
              </div>

              <div class="ss-cell">
                <span class="ss-badge ${badgeClass}">
                  ${row.issues.length}
                  ${badgeLabel}
                </span>

                <span class="ss-result">
                  ${resultLabel}
                </span>
              </div>
            </div>
          `;
        })
        .join("");
    }

    const shownStart = rows.length
      ? startIndex + 1
      : 0;

    const shownEnd = Math.min(
      startIndex + PAGE_SIZE,
      rows.length
    );

    $("ssPagerText").textContent =
      `Showing ${shownStart}–${shownEnd} ` +
      `of ${rows.length} inquiries ` +
      `· Rows per page: ${PAGE_SIZE}`;

    let pageButtons = `
      <button
        type="button"
        class="ss-page"
        data-page="previous"
        ${state.page === 1 ? "disabled" : ""}
      >
        ‹
      </button>
    `;

    for (
      let pageNumber = 1;
      pageNumber <= totalPages;
      pageNumber++
    ) {
      pageButtons += `
        <button
          type="button"
          class="ss-page ${
            state.page === pageNumber
              ? "active"
              : ""
          }"
          data-page="${pageNumber}"
        >
          ${pageNumber}
        </button>
      `;
    }

    pageButtons += `
      <button
        type="button"
        class="ss-page"
        data-page="next"
        ${
          state.page === totalPages
            ? "disabled"
            : ""
        }
      >
        ›
      </button>
    `;

    $("ssPages").innerHTML = pageButtons;

    document
      .querySelectorAll("#ssPages [data-page]")
      .forEach(button => {
        button.onclick = () => {
          const selectedPage =
            button.dataset.page;

          if (selectedPage === "previous") {
            state.page = Math.max(
              1,
              state.page - 1
            );
          } else if (selectedPage === "next") {
            state.page = Math.min(
              totalPages,
              state.page + 1
            );
          } else {
            state.page = Number(selectedPage);
          }

          renderValidationTable();
        };
      });
  }

  /* =========================================================
     CHOOSE FILE EVENTS
  ========================================================= */

  $("chooseFileBtn").onclick = event => {
    event.stopPropagation();
    spreadsheetInput.click();
  };

  uploadZone.onclick = event => {
    const targetId = event.target.id;

    if (
      targetId !== "validateBtn" &&
      targetId !== "chooseFileBtn"
    ) {
      spreadsheetInput.click();
    }
  };

  uploadZone.onkeydown = event => {
    if (
      event.key === "Enter" ||
      event.key === " "
    ) {
      event.preventDefault();
      spreadsheetInput.click();
    }
  };

  spreadsheetInput.onchange = () => {
    const selectedFile =
      spreadsheetInput.files[0];

    if (selectedFile) {
      acceptFile(selectedFile);
    }
  };

  /* =========================================================
     DRAG AND DROP EVENTS
  ========================================================= */

  ["dragenter", "dragover"].forEach(
    eventName => {
      uploadZone.addEventListener(
        eventName,
        event => {
          event.preventDefault();
          uploadZone.classList.add("drag");
        }
      );
    }
  );

  ["dragleave", "drop"].forEach(
    eventName => {
      uploadZone.addEventListener(
        eventName,
        event => {
          event.preventDefault();
          uploadZone.classList.remove("drag");
        }
      );
    }
  );

  uploadZone.addEventListener(
    "drop",
    event => {
      const droppedFile =
        event.dataTransfer.files[0];

      if (droppedFile) {
        acceptFile(droppedFile);
      }
    }
  );

  /* =========================================================
     VALIDATE BUTTON
  ========================================================= */

  validateButton.onclick = async event => {
    event.stopPropagation();

    validateButton.disabled = true;
    validateButton.textContent = "Reading...";

    showUploadError("");

    try {
      const rows = await readSpreadsheet();

      state.rows = rows;

      validateRows(rows);

      state.filter = "all";
      state.page = 1;

      $("ssSearch").value = "";
      $("ssSeverity").value = "all";
      $("ssSort").value = "fatal-first";

      document
        .querySelectorAll("#vresult .ss-ifilter")
        .forEach(button => {
          button.classList.toggle(
            "active",
            button.dataset.filter === "all"
          );
        });

      $("vresult").style.display = "block";

      updateValidationSummary();
      renderValidationTable();

      try {
        $("vresult").scrollIntoView({
          behavior: "smooth",
          block: "nearest"
        });
      } catch (error) {
        // Browser lama mungkin tidak mendukung scroll behavior.
      }
    } catch (error) {
      showUploadError(
        error.message ||
          "Spreadsheet gagal dibaca."
      );
    } finally {
      validateButton.disabled = false;
      validateButton.textContent =
        "Validate Inquiries";
    }
  };

  /* =========================================================
     ISSUE FILTER EVENTS
  ========================================================= */

  document
    .querySelectorAll("#vresult .ss-ifilter")
    .forEach(button => {
      button.onclick = () => {
        document
          .querySelectorAll("#vresult .ss-ifilter")
          .forEach(otherButton => {
            otherButton.classList.remove("active");
          });

        button.classList.add("active");

        state.filter =
          button.dataset.filter;

        state.page = 1;

        renderValidationTable();
      };
    });

  $("ssSearch").oninput = () => {
    state.page = 1;
    renderValidationTable();
  };

  $("ssSeverity").onchange = () => {
    state.page = 1;
    renderValidationTable();
  };

  $("ssSort").onchange = () => {
    state.page = 1;
    renderValidationTable();
  };

  /* =========================================================
     OPEN ADD CONFIRMATION
  ========================================================= */

  $("addBtn").onclick = () => {
    if (state.stats.add === 0) {
      return;
    }

    $("overlay4").classList.add("show");
  };

  /* =========================================================
     CONFIRM ADD INQUIRY
  ========================================================= */

  $("cfConfirm").onclick = () => {
    const acceptedRows = state.checked.filter(
      row => row.severity !== "fatal"
    );

    // Hilangkan highlight data sebelumnya
    dashRows.forEach(row => {
      row._new = false;
    });

    const newDashboardRows = acceptedRows.map(
      row => {
        const raw = row.raw;
        const phone = row.phone;

        return {
          branch: "HQ Training",

          student: raw.name || "",

          parent: "",

          source: (
            raw.source || ""
          ).toUpperCase(),

          country: phone.country || "",

          code: phone.code || "",

          phone: phone.pretty || "",

          date: raw.date
            ? fmtDate(raw.date)
            : "—",

          // Gunakan A jika status kosong/tidak valid
          chat: STATUS[raw.chat]
            ? raw.chat
            : "A",

          note: raw.note || "",

          _new: true,

          _incomplete:
            row.severity === "warning"
        };
      }
    );

    dashRows = [
      ...newDashboardRows,
      ...dashRows
    ];

    renderDash();

    $("overlay4").classList.remove("show");

    closeModal();

    toast(
      `${newDashboardRows.length} inquiry ditambahkan` +
      ` · ${state.stats.fatal} fatal dilewati.`
    );
  };

  /* =========================================================
     INITIAL STATE
  ========================================================= */

  clearValidationResult();
})();