// ==========================================
// PEMBERSIH CACHE OTOMATIS (MENCEGAH MEMORI PENUH)
// ==========================================
function bersihkanCacheLama() {
  const keysToRemove = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    // Cari semua memori telkom lama kecuali versi 14
    if (key && (key.startsWith("telkom_poi_state_") || key.startsWith("telkom_active_zones_"))) {
      if (key !== "telkom_poi_state_v14" && key !== "telkom_active_zones_v14") {
        keysToRemove.push(key);
      }
    }
  }
  // Hapus semua sampah memori lama untuk melonggarkan 5MB limit browser
  keysToRemove.forEach((k) => localStorage.removeItem(k));
}
bersihkanCacheLama(); // Jalankan saat file diload

// ==========================================
// UPLOAD 1: FILTER TABEL EXCEL
// ==========================================
document.getElementById("inputExcelFilter").addEventListener("change", function (e) {
  const file = e.target.files[0];
  const fileNameDisplay = document.getElementById("file-name-filter");

  if (!file) {
    if (fileNameDisplay) fileNameDisplay.innerText = "Belum ada file";
    return;
  }

  if (fileNameDisplay) fileNameDisplay.innerText = file.name;

  document.querySelectorAll(".status-info").forEach((el) => {
    el.innerText = "Sedang memproses tabel...";
  });

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const normalisasiTeks = (teks) => teks.replace(/[^a-zA-Z0-9]/g, "").toUpperCase();

      const sekarang = new Date();
      const timestamp = sekarang.toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
      document.getElementById("timestamp-info") && (document.getElementById("timestamp-info").innerText = "Update Terakhir: " + timestamp);

      function cariSheet(keyword) {
        return workbook.SheetNames.find((n) => n.trim().toUpperCase().includes(keyword.toUpperCase()));
      }

      function prosesSheetData(keywordSheet, kolomDibutuhkan, fungsiFilter, elementIdData, elementIdStatus, fungsiMap) {
        const statusElement = document.getElementById(elementIdStatus);
        const container = document.getElementById(elementIdData);
        const actualSheetName = cariSheet(keywordSheet);

        if (!actualSheetName) {
          statusElement.innerText = `Catatan: Sheet "${keywordSheet}" tidak ditemukan di file ini.`;
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
            if (fungsiMap) nilaiAsli = fungsiMap(kolomTarget, nilaiAsli, row);
            rowBaru[kolomTarget] = nilaiAsli === "" ? "-" : nilaiAsli;
          });
          return rowBaru;
        });

        if (dataTerfilter.length === 0) {
          container.innerHTML = `<p style='color:red;'>Filter gagal menemukan data.</p>`;
          statusElement.innerText = `Gagal memfilter sheet ${actualSheetName}.`;
        } else {
          tampilkanTabel(dataTerfilter, elementIdData, actualSheetName);
          statusElement.innerText = `Berhasil memproses ${dataTerfilter.length} data.`;
        }
      }

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
      alert("Terjadi kesalahan sistem saat membaca file Excel filter.");
    }
  };
  reader.readAsArrayBuffer(file);
});

function cariDataTabel() {
  const input = document.getElementById("kolomCari").value.toLowerCase();
  const rows = document.querySelectorAll(".table-container table tbody tr");
  rows.forEach((row) => {
    row.style.display = row.textContent.toLowerCase().includes(input) ? "" : "none";
  });
}

function tampilkanTabel(data, containerId, namaSheet) {
  const container = document.getElementById(containerId);
  let html = "<table><thead><tr>";
  Object.keys(data[0]).forEach((header) => {
    html += `<th>${header}</th>`;
  });
  html += "</tr></thead><tbody>";
  data.forEach((row) => {
    html += "<tr>";
    Object.values(row).forEach((isi) => {
      html += `<td>${isi}</td>`;
    });
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
      XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(table), nama);
      adaData = true;
    }
  });
  if (!adaData) return alert("Belum ada data yang bisa didownload!");
  XLSX.writeFile(wb, `HASIL_FILTER_GIANYAR.xlsx`);
}

function kirimKeTelegram() {
  alert("Sistem kirim Telegram membutuhkan backend aktif. Pastikan server lokal Anda menyala.");
}

// ==========================================
// UPLOAD 2: KHUSUS CSV / POI MAPS
// ==========================================
document.getElementById("inputCSVPOI").addEventListener("change", function (e) {
  const file = e.target.files[0];
  const fileNameDisplay = document.getElementById("file-name-poi");

  if (!file) {
    if (fileNameDisplay) fileNameDisplay.innerText = "Belum ada file";
    return;
  }
  if (fileNameDisplay) fileNameDisplay.innerText = file.name;

  const reader = new FileReader();
  const isCSV = file.name.toLowerCase().endsWith(".csv");

  reader.onload = function (evt) {
    try {
      let dataSheetPertama;

      if (isCSV) {
        const workbook = XLSX.read(evt.target.result, { type: "string" });
        const sheetPertama = workbook.SheetNames[0];
        dataSheetPertama = XLSX.utils.sheet_to_json(workbook.Sheets[sheetPertama], { defval: "" });
      } else {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetPertama = workbook.SheetNames[0];
        dataSheetPertama = XLSX.utils.sheet_to_json(workbook.Sheets[sheetPertama], { defval: "" });
      }

      let isValid = false;
      if (dataSheetPertama && dataSheetPertama.length > 0) {
        const firstRow = dataSheetPertama[0];
        const keys = Object.keys(firstRow).map((k) => k.toLowerCase().trim());
        if (keys.includes("latitude") && keys.includes("longitude")) {
          isValid = true;
        }
      }

      if (isValid) {
        integrasikanCSVkeZona(dataSheetPertama);
      } else {
        alert("File gagal dimuat ke Peta. Pastikan file memiliki kolom yang bernama persis 'Latitude' dan 'Longitude'.");
      }
    } catch (err) {
      console.error(err);
      alert("Gagal memproses file. Detail error: " + err.message);
    }
  };

  if (isCSV) {
    reader.readAsText(file, "UTF-8");
  } else {
    reader.readAsArrayBuffer(file);
  }
});

// ==========================================
// KODE PETA POINT OF INTEREST & ZONASI WARNA
// ==========================================
let map;
let zoneCircles = {};
let poiMarkers = L.layerGroup();

// Cache versi 14 (Aman dari Quota Exceeded karena versi lama sudah dihancurkan)
let poiState = JSON.parse(localStorage.getItem("telkom_poi_state_v14")) || {};
const RAD_KECIL = 400;

const baseZones = [
  { id: "g1", name: "Gianyar Kota", lat: -8.5414, lng: 115.3288, radius: RAD_KECIL, kecamatan: "Gianyar" },
  { id: "g2", name: "Lebih", lat: -8.5665, lng: 115.337, radius: RAD_KECIL, kecamatan: "Gianyar" },
  { id: "g3", name: "Tulikup", lat: -8.5484, lng: 115.3521, radius: RAD_KECIL, kecamatan: "Gianyar" },
  { id: "g4", name: "Sidan", lat: -8.5298, lng: 115.3486, radius: RAD_KECIL, kecamatan: "Gianyar" },
  { id: "g5", name: "Beng", lat: -8.5342, lng: 115.3281, radius: RAD_KECIL, kecamatan: "Gianyar" },
  { id: "u1", name: "Ubud Kota", lat: -8.5069, lng: 115.2625, radius: RAD_KECIL, kecamatan: "Ubud" },
  { id: "u2", name: "Mas", lat: -8.5323, lng: 115.2706, radius: RAD_KECIL, kecamatan: "Ubud" },
  { id: "u3", name: "Singakerta", lat: -8.5375, lng: 115.2497, radius: RAD_KECIL, kecamatan: "Ubud" },
  { id: "u4", name: "Peliatan", lat: -8.5204, lng: 115.2711, radius: RAD_KECIL, kecamatan: "Ubud" },
  { id: "u5", name: "Sayan", lat: -8.5042, lng: 115.2452, radius: RAD_KECIL, kecamatan: "Ubud" },
  { id: "u6", name: "Kedewatan", lat: -8.4831, lng: 115.2486, radius: RAD_KECIL, kecamatan: "Ubud" },
  { id: "t1", name: "Tampaksiring", lat: -8.4357, lng: 115.3117, radius: RAD_KECIL, kecamatan: "Tampaksiring" },
  { id: "t2", name: "Pejeng", lat: -8.5094, lng: 115.2952, radius: RAD_KECIL, kecamatan: "Tampaksiring" },
  { id: "t3", name: "Manukaya", lat: -8.4152, lng: 115.3161, radius: RAD_KECIL, kecamatan: "Tampaksiring" },
  { id: "t4", name: "Sanding", lat: -8.4552, lng: 115.3021, radius: RAD_KECIL, kecamatan: "Tampaksiring" },
  { id: "te1", name: "Tegallalang", lat: -8.4326, lng: 115.2787, radius: RAD_KECIL, kecamatan: "Tegallalang" },
  { id: "te2", name: "Kedisan", lat: -8.4111, lng: 115.2864, radius: RAD_KECIL, kecamatan: "Tegallalang" },
  { id: "te3", name: "Taro", lat: -8.3681, lng: 115.2801, radius: RAD_KECIL, kecamatan: "Tegallalang" },
  { id: "te4", name: "Sebatu", lat: -8.3963, lng: 115.2941, radius: RAD_KECIL, kecamatan: "Tegallalang" },
  { id: "b1", name: "Blahbatuh", lat: -8.5661, lng: 115.3005, radius: RAD_KECIL, kecamatan: "Blahbatuh" },
  { id: "b2", name: "Keramas", lat: -8.5912, lng: 115.3175, radius: RAD_KECIL, kecamatan: "Blahbatuh" },
  { id: "b3", name: "Bedulu", lat: -8.5218, lng: 115.3001, radius: RAD_KECIL, kecamatan: "Blahbatuh" },
  { id: "b4", name: "Belega", lat: -8.5583, lng: 115.3101, radius: RAD_KECIL, kecamatan: "Blahbatuh" },
  { id: "b5", name: "Saba", lat: -8.5833, lng: 115.3012, radius: RAD_KECIL, kecamatan: "Blahbatuh" },
  { id: "p1", name: "Payangan", lat: -8.3615, lng: 115.2514, radius: RAD_KECIL, kecamatan: "Payangan" },
  { id: "p2", name: "Kerta", lat: -8.3241, lng: 115.2621, radius: RAD_KECIL, kecamatan: "Payangan" },
  { id: "p3", name: "Buahan", lat: -8.3391, lng: 115.2451, radius: RAD_KECIL, kecamatan: "Payangan" },
  { id: "p4", name: "Melinggih", lat: -8.3912, lng: 115.2531, radius: RAD_KECIL, kecamatan: "Payangan" },
];

let activeZones = JSON.parse(localStorage.getItem("telkom_active_zones_v14")) || JSON.parse(JSON.stringify(baseZones));
let activeZoneId = null;

function getWarnaKecamatan(kecamatan) {
  switch (kecamatan) {
    case "Gianyar":
      return "#3498db";
    case "Ubud":
      return "#9b59b6";
    case "Tampaksiring":
      return "#e67e22";
    case "Tegallalang":
      return "#e74c3c";
    case "Blahbatuh":
      return "#f1c40f";
    case "Payangan":
      return "#1abc9c";
    default:
      return "#34495e";
  }
}

document.addEventListener("DOMContentLoaded", function () {
  initMapLeaflet();
  hitungUlangStatistikGlobal();
});

function initMapLeaflet() {
  map = L.map("map").setView([-8.45, 115.3], 11);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "Map data © OpenStreetMap contributors",
  }).addTo(map);

  poiMarkers.addTo(map);
  renderSeluruhPeta();
}

function renderSeluruhPeta() {
  const kecTerpilih = document.getElementById("filterKecamatan").value;
  const searchInput = document.getElementById("searchGlobalPOI").value.toLowerCase();
  poiMarkers.clearLayers();

  activeZones.forEach((zone) => {
    let circle = zoneCircles[zone.id];

    if (!circle) {
      circle = L.circle([zone.lat, zone.lng], {
        fillOpacity: 0.2,
        radius: zone.radius,
        interactive: true,
        weight: 2,
      });
      circle.bindTooltip(`<b>${zone.name}</b><br>Kec. ${zone.kecamatan}`, { className: "custom-tooltip" });

      circle.on("click", function () {
        activeZoneId = zone.id;
        document.getElementById("searchGlobalPOI").value = "";
        const panel = document.getElementById("poi-panel");
        panel.innerHTML = `<div style="text-align: center; margin-top: 40px;">
                             <h3 style="color: #2980b9;">⏳ Sedang Memproses...</h3>
                             <p style="color: #7f8c8d;">Memuat data prospek di area <b>${zone.name}</b></p>
                           </div>`;
        setTimeout(() => muatTargetZonaOverpass(zone), 100);
      });
      zoneCircles[zone.id] = circle;
    }

    let isMatchKec = kecTerpilih === "ALL" || zone.kecamatan === kecTerpilih;
    let isMatchSearch = false;
    let matchingPlacesIds = [];

    if (searchInput === "") {
      isMatchSearch = true;
      if (poiState[zone.id] && poiState[zone.id].places) {
        matchingPlacesIds = Object.keys(poiState[zone.id].places);
      }
    } else {
      if (zone.name.toLowerCase().includes(searchInput) || zone.kecamatan.toLowerCase().includes(searchInput)) {
        isMatchSearch = true;
        if (poiState[zone.id] && poiState[zone.id].places) {
          matchingPlacesIds = Object.keys(poiState[zone.id].places);
        }
      }

      if (poiState[zone.id] && poiState[zone.id].places) {
        for (const id in poiState[zone.id].places) {
          const p = poiState[zone.id].places[id];
          if (p.name.toLowerCase().includes(searchInput) || p.vicinity.toLowerCase().includes(searchInput)) {
            isMatchSearch = true;
            if (!matchingPlacesIds.includes(id)) matchingPlacesIds.push(id);
          }
        }
      }
    }

    if (isMatchKec && isMatchSearch) {
      if (!map.hasLayer(circle)) map.addLayer(circle);

      const data = poiState[zone.id];
      const baseColor = getWarnaKecamatan(zone.kecamatan);
      let circleColor = baseColor;

      if (data && data.total > 0 && data.visitedCount === data.total) {
        circleColor = "#27ae60";
      }

      circle.setStyle({ fillColor: circleColor, color: circleColor });

      if (data && data.places) {
        matchingPlacesIds.forEach((id) => {
          const p = data.places[id];
          const warnaPin = p.visited ? "#27ae60" : baseColor;

          const gmapsLink = `https://www.google.com/maps/search/?api=1&query=${p.lat},${p.lng}`;

          const marker = L.circleMarker([p.lat, p.lng], {
            radius: 7,
            fillColor: warnaPin,
            color: "#ffffff",
            weight: 1.5,
            fillOpacity: 0.95,
          }).bindPopup(`
            <b>${p.name}</b><br>
            ${p.vicinity}<br>
            <a href="${gmapsLink}" target="_blank" style="color:#3498db; text-decoration:none; font-weight:bold; display:block; margin:6px 0;">
               🗺️ Buka di Google Maps
            </a>
            Status: <b>${p.visited ? "Sudah Dikunjungi" : "Belum"}</b>
          `);

          poiMarkers.addLayer(marker);
        });
      }
    } else {
      if (map.hasLayer(circle)) map.removeLayer(circle);
    }
  });
}

function filterZonaKecamatan() {
  renderSeluruhPeta();
  const kecTerpilih = document.getElementById("filterKecamatan").value;
  if (activeZoneId && kecTerpilih !== "ALL") {
    const activeZone = activeZones.find((z) => z.id === activeZoneId);
    if (activeZone && activeZone.kecamatan !== kecTerpilih) {
      document.getElementById("poi-panel").innerHTML = "<p style='text-align: center; color: #7f8c8d; margin-top: 50px;'>Klik salah satu zona di peta untuk memuat target spesifik.</p>";
      activeZoneId = null;
    }
  }
}

function hitungJarakMeter(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const p1 = (lat1 * Math.PI) / 180;
  const p2 = (lat2 * Math.PI) / 180;
  const dp = ((lat2 - lat1) * Math.PI) / 180;
  const dl = ((lon2 - lon1) * Math.PI) / 180;
  const a = Math.sin(dp / 2) * Math.sin(dp / 2) + Math.cos(p1) * Math.cos(p2) * Math.sin(dl / 2) * Math.sin(dl / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function hitungUlangStatistikGlobal() {
  let countMaps = 0;
  let countCSV = 0;

  for (const zoneId in poiState) {
    const dataZone = poiState[zoneId];
    if (!dataZone.places) continue;

    for (const placeId in dataZone.places) {
      const place = dataZone.places[placeId];
      if (place.isCsv) countCSV++;
      else countMaps++;
    }
  }

  document.getElementById("stat-maps").innerText = countMaps;
  document.getElementById("stat-csv").innerText = countCSV;
  document.getElementById("stat-total").innerText = countMaps + countCSV;
}

// ========================================================
// SISTEM PEMETAAN PINTAR CSV (MEMAKAI TRY-CATCH KHUSUS MEMORI)
// ========================================================
function integrasikanCSVkeZona(dataExcel) {
  let jumlahTitikTerintegrasi = 0;
  let jumlahZonaBaru = 0;

  dataExcel.forEach((row, index) => {
    let latRaw, lngRaw, namaRaw, infoRaw, segmentRaw;

    Object.keys(row).forEach((k) => {
      const keyLow = k.toLowerCase().trim();
      if (keyLow === "latitude") latRaw = row[k];
      else if (keyLow === "longitude") lngRaw = row[k];
      else if (keyLow === "name" || keyLow === "nama") namaRaw = row[k];
      else if (keyLow === "adress" || keyLow === "alamat") infoRaw = row[k];
      else if (keyLow === "segment" || keyLow === "sub segment") segmentRaw = row[k];
    });

    const lat = parseFloat(latRaw);
    const lng = parseFloat(lngRaw);
    const nama = namaRaw || "Target CSV";
    const info = infoRaw || "";
    const segment = segmentRaw || "Data CSV";

    if (!isNaN(lat) && !isNaN(lng)) {
      let assignedZone = null;
      let closestDist = Infinity;

      activeZones.forEach((zone) => {
        const jarak = hitungJarakMeter(lat, lng, zone.lat, zone.lng);
        if (jarak <= zone.radius && jarak < closestDist) {
          closestDist = jarak;
          assignedZone = zone;
        }
      });

      if (!assignedZone) {
        let nearestKec = "Lainnya";
        let minDist = Infinity;

        baseZones.forEach((bz) => {
          let d = hitungJarakMeter(lat, lng, bz.lat, bz.lng);
          if (d < minDist) {
            minDist = d;
            nearestKec = bz.kecamatan;
          }
        });

        const newZoneId = "dyn_" + Date.now() + "_" + index;
        assignedZone = {
          id: newZoneId,
          name: "Area Ekspansi " + nearestKec,
          lat: lat,
          lng: lng,
          radius: RAD_KECIL,
          kecamatan: nearestKec,
        };

        activeZones.push(assignedZone);
        jumlahZonaBaru++;
      }

      if (!poiState[assignedZone.id]) {
        poiState[assignedZone.id] = { overpassLoaded: false, places: {}, total: 0, visitedCount: 0 };
      }

      const idUnik = "csv_" + index + "_" + assignedZone.id;
      if (!poiState[assignedZone.id].places[idUnik]) {
        poiState[assignedZone.id].places[idUnik] = {
          name: nama,
          lat: lat,
          lng: lng,
          vicinity: info + " (" + segment + ")",
          visited: false,
          isCsv: true,
        };
        poiState[assignedZone.id].total++;
        jumlahTitikTerintegrasi++;
      }
    }
  });

  simpanStatus();

  try {
    localStorage.setItem("telkom_active_zones_v14", JSON.stringify(activeZones));
  } catch (err) {
    if (err.name === "QuotaExceededError") {
      console.warn("Storage Penuh. Tidak dapat menyimpan zona baru secara permanen.");
    }
  }

  renderSeluruhPeta();
  hitungUlangStatistikGlobal();

  alert(`Sukses Memuat Data!\n\n${jumlahTitikTerintegrasi} data berhasil dipetakan.\nSistem otomatis menciptakan ${jumlahZonaBaru} Area Lingkaran baru.`);
}

// Fungsi Simpan yang sudah dilengkapi pelindung agar tidak menampilkan popup error menakutkan
function simpanStatus() {
  try {
    localStorage.setItem("telkom_poi_state_v14", JSON.stringify(poiState));
  } catch (err) {
    if (err.name === "QuotaExceededError") {
      alert(
        "Peringatan: Memori browser Anda (LocalStorage) sudah mencapai batas maksimal (biasanya 5MB). Pemetaan saat ini berhasil ditampilkan di layar, namun progres centang mungkin tidak tersimpan setelah browser ditutup. Untuk mengatasi ini secara permanen, bersihkan History/Cache browser Anda.",
      );
    } else {
      console.error(err);
    }
  }
}

function muatTargetZonaOverpass(zone) {
  const panel = document.getElementById("poi-panel");

  if (!poiState[zone.id]) {
    poiState[zone.id] = { overpassLoaded: false, places: {}, total: 0, visitedCount: 0 };
  }

  if (poiState[zone.id].overpassLoaded) {
    renderSeluruhPeta();
    renderDaftarTarget(zone);
    return;
  }

  const query = `[out:json][timeout:25];
  (
    node["amenity"~"cafe|school|clinic|hospital|restaurant|bank"](around:${zone.radius},${zone.lat},${zone.lng});
    way["amenity"~"cafe|school|clinic|hospital|restaurant|bank"](around:${zone.radius},${zone.lat},${zone.lng});
    node["tourism"~"hotel|resort|villa|hostel|guest_house|attraction"](around:${zone.radius},${zone.lat},${zone.lng});
    way["tourism"~"hotel|resort|villa|hostel|guest_house|attraction"](around:${zone.radius},${zone.lat},${zone.lng});
    node["shop"~"supermarket|convenience"](around:${zone.radius},${zone.lat},${zone.lng});
    way["shop"~"supermarket|convenience"](around:${zone.radius},${zone.lat},${zone.lng});
  );
  out center;`;

  fetch(`https://overpass-api.de/api/interpreter?data=${encodeURIComponent(query)}`)
    .then((response) => response.json())
    .then((data) => {
      const elements = data.elements;
      if (elements && elements.length > 0) {
        let validElements = elements.filter((el) => {
          const lat = el.lat || (el.center && el.center.lat);
          const lon = el.lon || (el.center && el.center.lon);
          return el.tags && el.tags.name && lat && lon;
        });

        validElements.forEach(function (el) {
          const lat = el.lat || (el.center && el.center.lat);
          const lon = el.lon || (el.center && el.center.lon);
          const tipe = el.tags.amenity || el.tags.tourism || el.tags.shop || "Target Maps";

          const idUnik = "maps_" + el.id;
          if (!poiState[zone.id].places[idUnik]) {
            poiState[zone.id].places[idUnik] = {
              name: el.tags.name,
              lat: lat,
              lng: lon,
              vicinity: "Kategori: " + tipe,
              visited: false,
              isCsv: false,
            };
            poiState[zone.id].total++;
          }
        });
      }

      poiState[zone.id].overpassLoaded = true;
      simpanStatus();
      renderSeluruhPeta();
      renderDaftarTarget(zone);
      hitungUlangStatistikGlobal();
    })
    .catch((error) => {
      console.error(error);
      panel.innerHTML = `
        <div style="text-align: center; margin-top: 40px;">
          <h3 style="color: #e74c3c;">❌ Koneksi Maps Gagal</h3>
          <button onclick="muatTargetZonaOverpass(activeZones.find(function(z) { return z.id === '${zone.id}'; }))" style="padding:8px 16px; background:#3498db; color:white; border:none; border-radius:4px; cursor:pointer;">🔄 Coba Lagi</button>
        </div>
      `;
    });
}

// ========================================================
// FUNGSI COPY TO CLIPBOARD REPORT
// ========================================================
function salinTeksKeClipboard(teks, btnId, teksAwal) {
  const btn = document.getElementById(btnId);
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(teks).then(() => efekTombolSukses(btn, teksAwal));
  } else {
    const textArea = document.createElement("textarea");
    textArea.value = teks;
    textArea.style.position = "fixed";
    textArea.style.left = "-999999px";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
      document.execCommand("copy");
      efekTombolSukses(btn, teksAwal);
    } catch (err) {
      alert("Browser menolak fitur salin otomatis. Silakan blok dan salin manual.");
    }
    textArea.remove();
  }
}

function efekTombolSukses(btn, teksAwal) {
  if (!btn) return;
  btn.innerText = "✅ Tersalin!";
  btn.style.background = "#27ae60";
  setTimeout(() => {
    btn.innerText = teksAwal;
    btn.style.background = "#8e44ad";
  }, 2000);
}

function salinDaftarGlobal(input) {
  let textToCopy = `📋 Laporan Target POI - Pencarian: "${input}"\n`;
  textToCopy += `=========================================\n\n`;

  let count = 1;
  for (const zoneId in poiState) {
    const dataZone = poiState[zoneId];
    if (!dataZone.places) continue;

    for (const placeId in dataZone.places) {
      const place = dataZone.places[placeId];
      if (place.name.toLowerCase().includes(input) || place.vicinity.toLowerCase().includes(input)) {
        const sumber = place.isCsv ? "CSV" : "MAPS";
        const status = place.visited ? "✅ Sudah Dicentang" : "❌ Belum Kunjungan";
        const link = `https://www.google.com/maps/search/?api=1&query=${place.lat},${place.lng}`;
        const zObj = activeZones.find((z) => z.id === zoneId);
        const namaZ = zObj ? zObj.name : zoneId;

        textToCopy += `${count}. ${place.name} [${sumber}]\n`;
        textToCopy += `   Area: ${namaZ}\n`;
        textToCopy += `   Detail: ${place.vicinity}\n`;
        textToCopy += `   Maps: ${link}\n`;
        textToCopy += `   Status: ${status}\n\n`;
        count++;
      }
    }
  }
  salinTeksKeClipboard(textToCopy, "btn-copy-global", "📋 Salin Hasil Pencarian");
}

function salinDaftarArea(zoneId) {
  const data = poiState[zoneId];
  const searchInput = document.getElementById("kolomCariPOI") ? document.getElementById("kolomCariPOI").value.toLowerCase() : "";
  const namaZona = activeZones.find((z) => z.id === zoneId).name;

  let textToCopy = `📋 Laporan Target POI - Area: ${namaZona}\n`;
  textToCopy += `=========================================\n\n`;

  let count = 1;
  for (const placeId in data.places) {
    const place = data.places[placeId];
    if (searchInput && !place.name.toLowerCase().includes(searchInput) && !place.vicinity.toLowerCase().includes(searchInput)) {
      continue;
    }

    const sumber = place.isCsv ? "CSV" : "MAPS";
    const status = place.visited ? "✅ Sudah Dicentang" : "❌ Belum Kunjungan";
    const link = `https://www.google.com/maps/search/?api=1&query=${place.lat},${place.lng}`;

    textToCopy += `${count}. ${place.name} [${sumber}]\n`;
    textToCopy += `   Detail: ${place.vicinity}\n`;
    textToCopy += `   Maps: ${link}\n`;
    textToCopy += `   Status: ${status}\n\n`;
    count++;
  }
  salinTeksKeClipboard(textToCopy, "btn-copy-area", "📋 Salin Daftar Area");
}

function cariGlobalPOI() {
  const input = document.getElementById("searchGlobalPOI").value.toLowerCase();
  const panel = document.getElementById("poi-panel");

  if (input.trim() === "") {
    renderSeluruhPeta();
    if (activeZoneId) {
      const zone = activeZones.find((z) => z.id === activeZoneId);
      if (zone) renderDaftarTarget(zone);
    } else {
      panel.innerHTML = "<p style='text-align: center; color: #7f8c8d; margin-top: 50px;'>Klik salah satu zona di peta untuk memuat target spesifik.</p>";
    }
    return;
  }

  let html = `<h3>🔍 Hasil Pencarian: "${input}"</h3>`;
  html += `<button onclick="salinDaftarGlobal('${input}')" id="btn-copy-global" style="width:100%; margin-bottom:15px; padding:10px 15px; background:#8e44ad; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold; transition:all 0.3s;">📋 Salin Hasil Pencarian</button>`;
  html += `<div id="poi-list-container">`;
  let jumlahDitemukan = 0;

  for (const zoneId in poiState) {
    const dataZone = poiState[zoneId];
    if (!dataZone.places) continue;

    for (const placeId in dataZone.places) {
      const place = dataZone.places[placeId];
      if (place.name.toLowerCase().includes(input) || place.vicinity.toLowerCase().includes(input)) {
        jumlahDitemukan++;
        const isChecked = place.visited ? "checked" : "";
        const badge = place.isCsv ? `<span class="badge-csv">CSV</span>` : `<span class="badge-maps">MAPS</span>`;
        const zObj = activeZones.find((z) => z.id === zoneId);
        const namaZ = zObj ? zObj.name : zoneId;

        const gmapsLink = `https://www.google.com/maps/search/?api=1&query=${place.lat},${place.lng}`;

        html += `
          <div class="target-item poi-item">
            <input type="checkbox" id="glob_${placeId}" onchange="ubahStatusKunjungan('${zoneId}', '${placeId}', this.checked)" ${isChecked}>
            <div class="target-info">
              <h4>${place.name} ${badge} <span style="font-weight:normal; font-size:11px; color:#7f8c8d;">(Area: ${namaZ})</span></h4>
              <p style="font-size:11px; margin: 3px 0;">
                <a href="${gmapsLink}" target="_blank" style="color:#2980b9; font-weight:bold; text-decoration:none;">
                  📍 ${place.lat.toFixed(5)}, ${place.lng.toFixed(5)} (Buka Maps)
                </a>
              </p>
              <p>${place.vicinity}</p>
            </div>
          </div>
        `;
      }
    }
  }

  if (jumlahDitemukan === 0) {
    html += `<p style='text-align:center; color:#7f8c8d; margin-top:20px;'>Tidak ada target ditemukan.</p>`;
  }

  html += `</div>`;
  panel.innerHTML = html;
  renderSeluruhPeta();
}

function cariDataPOI() {
  const input = document.getElementById("kolomCariPOI").value.toLowerCase();
  const items = document.querySelectorAll("#poi-list-container .poi-item");
  items.forEach((item) => {
    const text = item.textContent.toLowerCase();
    item.style.display = text.includes(input) ? "" : "none";
  });
}

function renderDaftarTarget(zone) {
  const panel = document.getElementById("poi-panel");
  const data = poiState[zone.id];

  if (data.total === 0) {
    panel.innerHTML = "<p style='text-align:center; color:#7f8c8d; margin-top:30px;'>Tidak ada target yang ditemukan di area " + zone.name + ".</p>";
    return;
  }

  let html = `<h3>Area: ${zone.name}</h3>`;
  html += `<p style='font-size:13px; color:gray;'>Total Target: ${data.total} | Dikunjungi: <span id='visit-count-${zone.id}'>${data.visitedCount}</span></p>`;

  html += `<div style="display:flex; gap:10px; margin-bottom:15px;">
             <input type="text" id="kolomCariPOI" onkeyup="cariDataPOI()" placeholder="🔍 Cari di area ini..." style="flex:1; padding:10px; border-radius:4px; border:1px solid #3498db; box-sizing:border-box; outline:none;">
             <button onclick="salinDaftarArea('${zone.id}')" id="btn-copy-area" style="padding:10px 15px; background:#8e44ad; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold; white-space:nowrap; transition:all 0.3s;">📋 Salin Daftar</button>
           </div>`;

  html += `<div id="poi-list-container">`;

  for (const placeId in data.places) {
    const place = data.places[placeId];
    const isChecked = place.visited ? "checked" : "";
    const badge = place.isCsv ? `<span class="badge-csv">CSV</span>` : `<span class="badge-maps">MAPS</span>`;

    const gmapsLink = `https://www.google.com/maps/search/?api=1&query=${place.lat},${place.lng}`;

    html += `
      <div class="target-item poi-item">
        <input type="checkbox" id="${placeId}" onchange="ubahStatusKunjungan('${zone.id}', '${placeId}', this.checked)" ${isChecked}>
        <div class="target-info">
          <h4>${place.name} ${badge}</h4>
          <p style="font-size:11px; margin: 3px 0;">
            <a href="${gmapsLink}" target="_blank" style="color:#2980b9; font-weight:bold; text-decoration:none;">
              📍 ${place.lat.toFixed(5)}, ${place.lng.toFixed(5)} (Buka Maps)
            </a>
          </p>
          <p>${place.vicinity}</p>
        </div>
      </div>
    `;
  }
  html += `</div>`;
  panel.innerHTML = html;
}

function ubahStatusKunjungan(zoneId, placeId, isChecked) {
  const data = poiState[zoneId];
  if (data && data.places[placeId]) {
    const p = data.places[placeId];
    const statusLama = p.visited;
    p.visited = isChecked;

    if (isChecked && !statusLama) data.visitedCount++;
    else if (!isChecked && statusLama) data.visitedCount--;

    const visitCounter = document.getElementById("visit-count-" + zoneId);
    if (visitCounter) visitCounter.innerText = data.visitedCount;

    const globCheck = document.getElementById("glob_" + placeId);
    if (globCheck) globCheck.checked = isChecked;

    const localCheck = document.getElementById(placeId);
    if (localCheck) localCheck.checked = isChecked;

    simpanStatus();
    renderSeluruhPeta();
  }
}
