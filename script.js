document.getElementById("inputExcel").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  document.querySelectorAll(".status-info").forEach((el) => (el.innerText = "Sedang memproses data..."));

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const normalisasiTeks = (teks) => teks.replace(/[^a-zA-Z0-9]/g, "").toUpperCase();

      // Timestamp diproses
      const sekarang = new Date();
      const timestamp = sekarang.toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
      document.getElementById("timestamp-info") && (document.getElementById("timestamp-info").innerText = "Diproses pada: " + timestamp);

      // Cari sheet berdasarkan keyword (tidak peduli tanggal)
      function cariSheet(keyword) {
        return workbook.SheetNames.find((n) => n.trim().toUpperCase().includes(keyword.toUpperCase()));
      }

      function prosesSheetData(keywordSheet, kolomDibutuhkan, fungsiFilter, elementIdData, elementIdStatus, fungsiMap) {
        const statusElement = document.getElementById(elementIdStatus);
        const container = document.getElementById(elementIdData);

        const actualSheetName = cariSheet(keywordSheet);

        if (!actualSheetName) {
          statusElement.innerText = `Kesalahan: Sheet "${keywordSheet}" tidak ditemukan!`;
          container.innerHTML = "";
          return;
        }

        const sheet = workbook.Sheets[actualSheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: "" });

        if (jsonData.length === 0) {
          statusElement.innerText = `Sheet ${actualSheetName} kosong.`;
          return;
        }

        const dataTerfilter = jsonData.filter(fungsiFilter).map((row) => {
          let rowBaru = {};
          kolomDibutuhkan.forEach((kolomTarget) => {
            const kolomTargetNormal = normalisasiTeks(kolomTarget);
            const keyAsli = Object.keys(row).find((k) => normalisasiTeks(k) === kolomTargetNormal);

            let nilaiAsli = keyAsli && row[keyAsli] !== undefined ? String(row[keyAsli]).trim() : "";

            if (fungsiMap) {
              nilaiAsli = fungsiMap(kolomTarget, nilaiAsli, row);
            }

            rowBaru[kolomTarget] = nilaiAsli === "" ? "-" : nilaiAsli;
          });
          return rowBaru;
        });

        if (dataTerfilter.length === 0) {
          container.innerHTML = `<p style='color:red;'>Filter gagal menemukan data yang cocok.</p>`;
          statusElement.innerText = `Gagal memfilter sheet ${actualSheetName}.`;
        } else {
          tampilkanTabel(dataTerfilter, elementIdData, actualSheetName);
          statusElement.innerText = `Berhasil memproses ${dataTerfilter.length} data. | ${timestamp}`;
        }
      }

      // 1. Eksekusi Data BILLPER
      prosesSheetData(
        "BILLPER APRIL",
        ["CCA", "SND", "PAID", "SND_GROUP", "NCLI", "NAMA_NCLI", "BILL_AMOUNT", "NOMER TLP", "STO_DESC", "PRODUK", "BUNDLING", "USAGE_DESC", "NAMA"],
        (row) => {
          const keyPaid = Object.keys(row).find((k) => normalisasiTeks(k) === "PAID");
          const valPaid = keyPaid && row[keyPaid] !== undefined ? String(row[keyPaid]).trim() : "";
          const cekPaid = valPaid === "" || valPaid === "-" || valPaid.toUpperCase() === "#N/A";

          const keySto = Object.keys(row).find((k) => normalisasiTeks(k) === "STODESC");
          const valSto = keySto && row[keySto] !== undefined ? String(row[keySto]).toUpperCase().trim() : "";
          const cekSto = valSto.startsWith("GIN") || valSto.startsWith("UBU") || valSto.startsWith("TPS");

          return cekPaid && cekSto;
        },
        "data-billper",
        "status-billper",
        function (kolomTarget, nilai) {
          if (kolomTarget === "PAID") return "0";
          return nilai;
        },
      );

      // 2. Eksekusi Data PRANPC
      prosesSheetData(
        "PRANPC APRIL",
        ["SND", "PAID MARET", "PAID APRIL", "NOMOR TELP", "DATEL", "NAMA PELANGGAN", "USAGE_DESC", "UMUR CUSTOMER", "HASIL CARING"],
        (row) => {
          const keyDatel = Object.keys(row).find((k) => normalisasiTeks(k).includes("DATEL"));
          const valDatel = keyDatel && row[keyDatel] ? String(row[keyDatel]).toUpperCase() : "";
          const cekDatel = valDatel.includes("91804") || valDatel.includes("91084") || valDatel.includes("GIANYAR");

          const keyPaidFeb = Object.keys(row).find((k) => normalisasiTeks(k).includes("PAIDFEBRUARI"));
          const valPaidFeb = keyPaidFeb && row[keyPaidFeb] !== undefined ? String(row[keyPaidFeb]).trim() : "";
          const cekFeb = valPaidFeb === "0" || valPaidFeb === "" || valPaidFeb === "-";

          const keyPaidMar = Object.keys(row).find((k) => normalisasiTeks(k).includes("PAIDMARET"));
          const valPaidMar = keyPaidMar && row[keyPaidMar] !== undefined ? String(row[keyPaidMar]).trim() : "";
          const cekMar = valPaidMar === "0" || valPaidMar === "" || valPaidMar === "-" || valPaidMar.includes("1900") || valPaidMar.includes("1899");

          return cekDatel && cekFeb && cekMar;
        },
        "data-pranpc",
        "status-pranpc",
        function (kolomTarget, nilai) {
          if (kolomTarget === "PAID MARET") return "00.01.1900";
          if (kolomTarget === "PAID APRIL") return "0";
          return nilai;
        },
      );

      // 3. Eksekusi Data C3MR
      prosesSheetData(
        "C3MR APRIL",
        ["SND", "SND_GROUP", "PAID", "NCLI", "DATEL", "NAMA PELANGGAN", "USAGE_DESC", "BILL_AMOUNT"],
        (row) => {
          const keyDatel = Object.keys(row).find((k) => normalisasiTeks(k).includes("DATEL"));
          const valDatel = keyDatel && row[keyDatel] ? String(row[keyDatel]).toUpperCase() : "";
          const cekDatel = valDatel.includes("91804") || valDatel.includes("91084") || valDatel.includes("GIANYAR");

          const keyPaid = Object.keys(row).find((k) => normalisasiTeks(k) === normalisasiTeks("PAID"));
          const valPaid = keyPaid && row[keyPaid] !== undefined ? String(row[keyPaid]).trim().toUpperCase() : "";
          const cekPaid = valPaid === "0" || valPaid === "" || valPaid === "-";

          return cekDatel && cekPaid;
        },
        "data-c3mr",
        "status-c3mr",
        function (kolomTarget, nilai) {
          if (kolomTarget === "PAID") {
            const teksNilai = String(nilai).trim();
            if (teksNilai === "0" || teksNilai === "" || teksNilai === "-") return "0";
          }
          return nilai;
        },
      );
    } catch (error) {
      console.error(error);
      document.querySelectorAll(".status-info").forEach((el) => (el.innerText = "Terjadi kesalahan sistem saat membaca file."));
    }
  };
  reader.readAsArrayBuffer(file);
});

function tampilkanTabel(data, containerId, namaSheet) {
  const container = document.getElementById(containerId);

  let html = "<table><thead><tr>";
  Object.keys(data[0]).forEach((header) => (html += `<th>${header}</th>`));
  html += "</tr></thead><tbody>";

  data.forEach((row) => {
    html += "<tr>";
    Object.values(row).forEach((isi) => (html += `<td>${isi}</td>`));
    html += "</tr>";
  });

  html += "</tbody></table>";
  container.innerHTML = html;

  if (!document.getElementById("action-buttons")) {
    const btnWrap = document.createElement("div");
    btnWrap.id = "action-buttons";
    btnWrap.style = "margin: 16px 0; display: flex; gap: 10px;";
    btnWrap.innerHTML = `
      <button id="btn-download-semua" onclick="downloadSemua()" style="padding:8px 20px;cursor:pointer;background:#1a7f4b;color:white;border:none;border-radius:4px;font-size:14px;font-weight:bold;">
        ⬇ Download Semua
      </button>
      <button id="btn-kirim-telegram" onclick="kirimKeTelegram()" style="padding:8px 20px;cursor:pointer;background:#0088cc;color:white;border:none;border-radius:4px;font-size:14px;font-weight:bold;">
        ✈️ Kirim ke Telegram
      </button>
    `;
    document.getElementById("data-billper").before(btnWrap);
  }
}

function downloadSemua() {
  const wb = XLSX.utils.book_new();
  const sheets = [
    { id: "data-billper", nama: "BILLPER" },
    { id: "data-pranpc", nama: "PRANPC" },
    { id: "data-c3mr", nama: "C3MR" },
  ];

  let adaData = false;
  sheets.forEach(({ id, nama }) => {
    const table = document.querySelector(`#${id} table`);
    if (table) {
      const ws = XLSX.utils.table_to_sheet(table);
      XLSX.utils.book_append_sheet(wb, ws, nama);
      adaData = true;
    }
  });

  if (!adaData) {
    alert("Belum ada data yang bisa didownload!");
    return;
  }

  const sekarang = new Date();
  const tgl = `${sekarang.getDate().toString().padStart(2, "0")}${(sekarang.getMonth() + 1).toString().padStart(2, "0")}${sekarang.getFullYear()}`;
  XLSX.writeFile(wb, `HASIL_FILTER_GIANYAR_${tgl}.xlsx`);
}

function kirimKeTelegram() {
  const wb = XLSX.utils.book_new();
  const tableBillper = document.querySelector("#data-billper table");

  if (!tableBillper) {
    alert("Data BILLPER kosong atau belum diproses!");
    return;
  }

  const ws = XLSX.utils.table_to_sheet(tableBillper);
  XLSX.utils.book_append_sheet(wb, ws, "BILLPER");

  const btnKirim = document.getElementById("btn-kirim-telegram");
  btnKirim.innerText = "⏳ Mengirim BILLPER...";
  btnKirim.disabled = true;

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  const sekarang = new Date();
  const tgl = `${sekarang.getDate().toString().padStart(2, "0")}${(sekarang.getMonth() + 1).toString().padStart(2, "0")}${sekarang.getFullYear()}`;
  const namaFile = `HASIL_BILLPER_GIANYAR_${tgl}.xlsx`;

  const formData = new FormData();
  formData.append("file", blob, namaFile);

  fetch("http://127.0.0.1:5000/upload-excel", {
    method: "POST",
    body: formData,
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.status === "success") {
        alert("✅ Data BILLPER Berhasil dikirim ke Telegram!");
      } else {
        alert("❌ Terjadi kesalahan pada server bot.");
      }
      btnKirim.innerText = "✈️ Kirim ke Telegram";
      btnKirim.disabled = false;
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("❌ Gagal mengirim. Pastikan server Flask sedang berjalan.");
      btnKirim.innerText = "✈️ Kirim ke Telegram";
      btnKirim.disabled = false;
    });
}

function cariData() {
  const input = document.getElementById("kolomCari").value.toLowerCase();
  const rows = document.querySelectorAll(".table-container table tbody tr");
  rows.forEach((row) => {
    row.style.display = row.textContent.toLowerCase().includes(input) ? "" : "none";
  });
}
