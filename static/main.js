/* ============================================================
   Profiling System — main.js
   ============================================================ */

(function () {
  "use strict";

  // ── Elements ──────────────────────────────────────────────
  const selectJenis    = document.getElementById("jenis");
  const step2          = document.getElementById("step-2");
  const step3          = document.getElementById("step-3");
  const step4          = document.getElementById("step-4");
  const uploadSalesWrap= document.getElementById("upload-sales-wrap");
  const dataLabelMain  = document.getElementById("data-label-main");

  const fileDataInput  = document.getElementById("file_data");
  const fileDataChosen = document.getElementById("data-chosen");
  const fileDataName   = document.getElementById("data-chosen-name");
  const sheetDataWrap  = document.getElementById("sheet-data-wrap");
  const sheetDataSelect= document.getElementById("sheet_data");

  const fileSalesInput = document.getElementById("file_sales");
  const fileSalesChosen= document.getElementById("sales-chosen");
  const fileSalesName  = document.getElementById("sales-chosen-name");
  const sheetSalesWrap = document.getElementById("sheet-sales-wrap");
  const sheetSalesSelect= document.getElementById("sheet_sales");

  const formProses     = document.getElementById("form-proses");
  const btnProses      = document.getElementById("btn-proses");
  const loading        = document.getElementById("loading");

  const formSection    = document.getElementById("form-section");
  const resultSection  = document.getElementById("result-section");
  const errorSection   = document.getElementById("error-section");

  const resultFilename = document.getElementById("result-filename");
  const resultStats    = document.getElementById("result-stats");
  const previewThead   = document.getElementById("preview-thead");
  const previewTbody   = document.getElementById("preview-tbody");
  const btnDownload    = document.getElementById("btn-download");
  const btnUlang       = document.getElementById("btn-ulang");
  const btnUlangErr    = document.getElementById("btn-ulang-err");
  const errorMsg       = document.getElementById("error-msg");

  // ── Warna per kolom (index 1-based) ─────────────────────
  const WARNA_INBOUND = {
    hdr: {
      1:"#C6EFCE", 2:"#DAFDE0", 3:"#D9D9D9", 4:"#BDD7EE",
      5:"#FCE4D6", 6:"#FCE4D6", 7:"#E2EFDA", 8:"#E2EFDA",
      9:"#E8D5F5", 10:"#E8D5F5", 11:"#FFF2CC", 12:"#DDEBF7", 13:"#DDEBF7"
    },
    data: {
      1:"#EBF5EB", 2:"#F0FAF0", 3:"#F2F2F2", 4:"#DEEAF1",
      5:"#FEF3EE", 6:"#FEF3EE", 7:"#F2F8EE", 8:"#F2F8EE",
      9:"#F5EEFA", 10:"#F5EEFA", 11:"#FFFAE6", 12:"#EEF5FB", 13:"#EEF5FB"
    }
  };

  const WARNA_OUTBOUND = {
    hdr: {
      1:"#C6EFCE", 2:"#DAFDE0", 3:"#D9D9D9", 4:"#BDD7EE",
      5:"#FCE4D6", 6:"#FCE4D6", 7:"#E2EFDA", 8:"#E2EFDA",
      9:"#E8D5F5", 10:"#E8D5F5", 11:"#FFF2CC",
      12:"#DDEBF7", 13:"#DDEBF7", 14:"#FFD7D7"
    },
    data: {
      1:"#EBF5EB", 2:"#F0FAF0", 3:"#F2F2F2", 4:"#DEEAF1",
      5:"#FEF3EE", 6:"#FEF3EE", 7:"#F2F8EE", 8:"#F2F8EE",
      9:"#F5EEFA", 10:"#F5EEFA", 11:"#FFFAE6",
      12:"#EEF5FB", 13:"#EEF5FB", 14:"#FFEAEA"
    }
  };

  let currentFile = null;   // { b64, filename }

  // ── Step 1: Pilih Jenis ───────────────────────────────────
  selectJenis.addEventListener("change", function () {
    const val = this.value;
    step2.classList.remove("hidden");
    step3.classList.remove("hidden");
    step4.classList.remove("hidden");

    // Reset upload
    resetUpload();

    if (val === "outbound") {
      uploadSalesWrap.classList.remove("hidden");
      dataLabelMain.textContent = "Klik untuk upload Data Kotor Outbound";
      fileSalesInput.required = true;
    } else {
      uploadSalesWrap.classList.add("hidden");
      fileSalesInput.required = false;
      if (val === "homi") {
        dataLabelMain.textContent = "Klik untuk upload Data Kotor HOMi";
      } else {
        dataLabelMain.textContent = "Klik untuk upload Data Kotor Website";
      }
    }
  });

  function resetUpload() {
    fileDataInput.value = "";
    fileSalesInput.value = "";
    fileDataChosen.classList.add("hidden");
    fileSalesChosen.classList.add("hidden");
    sheetDataWrap.classList.add("hidden");
    sheetSalesWrap.classList.add("hidden");
    sheetDataSelect.innerHTML = "";
    sheetSalesSelect.innerHTML = "";
  }

  // ── Inspeksi Sheet Excel ─────────────────────────────────
  async function inspectSheets(file, selectElem, wrapElem, isSales) {
    if (!file) {
      wrapElem.classList.add("hidden");
      selectElem.innerHTML = "";
      return;
    }
    const fd = new FormData();
    fd.append("file", file);
    try {
      const res = await fetch("/api/inspect-sheets", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok || !data.success || !data.sheets) {
        wrapElem.classList.add("hidden");
        return;
      }
      const sheets = data.sheets;
      if (sheets.length <= 1) {
        // Hanya 1 sheet, otomatis terpilih tanpa perlu dropdown
        wrapElem.classList.add("hidden");
        selectElem.innerHTML = sheets.length ? `<option value="${sheets[0]}" selected>${sheets[0]}</option>` : "";
        return;
      }

      // Lebih dari 1 sheet -> tampilkan dropdown
      selectElem.innerHTML = "";
      const bulan = (document.getElementById("bulan").value || "").toLowerCase();

      if (isSales) {
        const optAll = document.createElement("option");
        optAll.value = "__ALL__";
        optAll.textContent = `📊 Gabungkan Semua Sheet (${sheets.length} Sheet)`;
        selectElem.appendChild(optAll);
      }

      let matchedIndex = -1;
      sheets.forEach((s, idx) => {
        const opt = document.createElement("option");
        opt.value = s;
        opt.textContent = isSales ? `Hanya Sheet: ${s}` : `Sheet: ${s}`;
        const sLower = s.toLowerCase();
        if (bulan && (sLower.includes(bulan.slice(0, 3)) || bulan.includes(sLower))) {
          matchedIndex = isSales ? idx + 1 : idx; // offset 1 if __ALL__ is first
        }
        selectElem.appendChild(opt);
      });

      // Otomatis pilih sheet yang cocok dengan bulan jika ada
      if (matchedIndex >= 0) {
        selectElem.selectedIndex = matchedIndex;
      }

      wrapElem.classList.remove("hidden");
    } catch (e) {
      console.warn("Gagal membaca sheets:", e);
      wrapElem.classList.add("hidden");
    }
  }

  // ── File chosen indicator & sheet inspection ─────────────
  fileDataInput.addEventListener("change", function () {
    if (this.files.length) {
      fileDataName.textContent = this.files[0].name;
      fileDataChosen.classList.remove("hidden");
      inspectSheets(this.files[0], sheetDataSelect, sheetDataWrap, false);
    } else {
      sheetDataWrap.classList.add("hidden");
    }
  });

  fileSalesInput.addEventListener("change", function () {
    if (this.files.length) {
      fileSalesName.textContent = this.files[0].name;
      fileSalesChosen.classList.remove("hidden");
      inspectSheets(this.files[0], sheetSalesSelect, sheetSalesWrap, true);
    } else {
      sheetSalesWrap.classList.add("hidden");
    }
  });

  // Drag & drop visual feedback
  [document.getElementById("upload-data-area"),
   document.getElementById("upload-sales-area")].forEach(area => {
    if (!area) return;
    area.addEventListener("dragover", e => {
      e.preventDefault();
      area.classList.add("drag-over");
    });
    area.addEventListener("dragleave", () => area.classList.remove("drag-over"));
    area.addEventListener("drop", () => area.classList.remove("drag-over"));
  });

  // ── Form Submit ──────────────────────────────────────────
  formProses.addEventListener("submit", async function (e) {
    e.preventDefault();

    const jenis = selectJenis.value;
    const bulan = document.getElementById("bulan").value;
    const tahun = document.getElementById("tahun").value;

    // Validasi manual
    if (!jenis) { alert("Pilih jenis laporan terlebih dahulu."); return; }
    if (!bulan) { alert("Pilih bulan terlebih dahulu."); return; }
    if (!tahun) { alert("Masukkan tahun terlebih dahulu."); return; }
    if (!fileDataInput.files.length) { alert("Upload file data kotor terlebih dahulu."); return; }
    if (jenis === "outbound" && !fileSalesInput.files.length) {
      alert("Upload file data sales terlebih dahulu."); return;
    }

    // Tampilkan loading
    loading.classList.remove("hidden");
    btnProses.disabled = true;
    hideResults();

    try {
      const fd = new FormData(formProses);
      const res = await fetch("/api/proses", { method: "POST", body: fd });
      const data = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(data.error || "Terjadi kesalahan tidak diketahui.");
      }

      currentFile = { b64: data.file_b64, filename: data.filename };
      renderResult(data);

    } catch (err) {
      showError(err.message);
    } finally {
      loading.classList.add("hidden");
      btnProses.disabled = false;
    }
  });

  // ── Render Result ────────────────────────────────────────
  function renderResult(data) {
    const jenis = selectJenis.value;
    const warna = (jenis === "outbound") ? WARNA_OUTBOUND : WARNA_INBOUND;

    // Nama file
    resultFilename.textContent = "📄 " + data.filename;

    // Stats
    const jumlahDealer = data.rows.length;
    const totalLeads   = data.rows.reduce((s, r) => s + (parseInt(r[3]) || 0), 0);
    resultStats.innerHTML = `
      <div class="stat-pill">Dealer Aktif: <strong>${jumlahDealer}</strong></div>
      <div class="stat-pill">Total Leads: <strong>${totalLeads.toLocaleString("id-ID")}</strong></div>
      <div class="stat-pill">Periode: <strong>${data.rows[0]?.[0] || "-"} ${data.rows[0]?.[1] || ""}</strong></div>
    `;

    // Header tabel
    const tr = document.createElement("tr");
    data.kolom.forEach((label, i) => {
      const th = document.createElement("th");
      th.textContent = label;
      const bg = warna.hdr[i + 1] || "#e2e8f0";
      th.style.backgroundColor = bg;
      th.style.color = "#1a1d27";
      tr.appendChild(th);
    });
    previewThead.innerHTML = "";
    previewThead.appendChild(tr);

    // Body tabel
    previewTbody.innerHTML = "";
    data.rows.forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((val, i) => {
        const td = document.createElement("td");
        td.textContent = (val === null || val === undefined) ? "-" : val;
        const bg = warna.data[i + 1] || "#f8fafc";
        td.style.backgroundColor = bg;
        td.style.color = "#1a1d27";
        tr.appendChild(td);
      });
      previewTbody.appendChild(tr);
    });

    // Tampilkan result
    formSection.classList.add("hidden");
    resultSection.classList.remove("hidden");
    errorSection.classList.add("hidden");

    // Scroll ke hasil
    resultSection.scrollIntoView({ behavior: "smooth" });
  }

  // ── Download ────────────────────────────────────────────
  btnDownload.addEventListener("click", function () {
    if (!currentFile) return;
    const bin = atob(currentFile.b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = currentFile.filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });

  // ── Ulang ───────────────────────────────────────────────
  function ulang() {
    currentFile = null;
    formProses.reset();
    resetUpload();
    step2.classList.add("hidden");
    step3.classList.add("hidden");
    step4.classList.add("hidden");
    uploadSalesWrap.classList.add("hidden");

    formSection.classList.remove("hidden");
    resultSection.classList.add("hidden");
    errorSection.classList.add("hidden");
    formSection.scrollIntoView({ behavior: "smooth" });
  }

  btnUlang.addEventListener("click", ulang);
  btnUlangErr.addEventListener("click", ulang);

  // ── Error ───────────────────────────────────────────────
  function showError(msg) {
    errorMsg.textContent = msg;
    errorSection.classList.remove("hidden");
    errorSection.scrollIntoView({ behavior: "smooth" });
  }

  function hideResults() {
    resultSection.classList.add("hidden");
    errorSection.classList.add("hidden");
  }

})();
