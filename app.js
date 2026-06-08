// ================================================================
// GOOGLE APPS SCRIPT — Penerima Data Kuis
// Cara pakai:
// 1. Buka script.google.com → New Project
// 2. Paste seluruh kode ini
// 3. Ganti SPREADSHEET_ID di bawah dengan ID Google Sheets kamu
// 4. Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Salin URL deployment → paste ke APPS_SCRIPT_URL di kuis.js
// ================================================================

const SPREADSHEET_ID = "GANTI_DENGAN_ID_SPREADSHEET_KAMU";

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── Sheet 1: Data Mentah Siswa ──
    let sheetData = ss.getSheetByName("Data Kuis");
    if (!sheetData) {
      sheetData = ss.insertSheet("Data Kuis");
      sheetData.appendRow([
        "No", "Timestamp", "Nama", "Kelas", "Sekolah",
        "Skor (x/10)", "Nilai (0–100)",
        "S1","S2","S3","S4","S5","S6","S7","S8","S9","S10",
        "Durasi (detik)"
      ]);
      // Format header
      const header = sheetData.getRange(1, 1, 1, 19);
      header.setBackground("#5C0E0E");
      header.setFontColor("#FFFFFF");
      header.setFontWeight("bold");
      sheetData.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents);
    const nomor = sheetData.getLastRow(); // otomatis nomor urut

    sheetData.appendRow([
      nomor,
      new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" }),
      data.nama,
      data.kelas,
      data.sekolah || "-",
      data.skor + "/10",
      data.nilai,
      ...data.jawaban,   // 10 kolom: "✓ Benar" atau "✗ Salah"
      data.durasi || "-"
    ]);

    // ── Sheet 2: Leaderboard ──
    let sheetBoard = ss.getSheetByName("Leaderboard");
    if (!sheetBoard) {
      sheetBoard = ss.insertSheet("Leaderboard");
      sheetBoard.appendRow(["Rank","Nama","Kelas","Sekolah","Nilai","Waktu"]);
      const hBoard = sheetBoard.getRange(1, 1, 1, 6);
      hBoard.setBackground("#5C0E0E");
      hBoard.setFontColor("#FFFFFF");
      hBoard.setFontWeight("bold");
      sheetBoard.setFrozenRows(1);
    }

    sheetBoard.appendRow([
      "",  // rank dikosongkan, bisa pakai formula
      data.nama,
      data.kelas,
      data.sekolah || "-",
      data.nilai,
      new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" })
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", pesan: "Data berhasil disimpan." }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", pesan: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET: untuk leaderboard (dipanggil dari web) ──
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Leaderboard");
    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    const data = rows
      .map(r => ({
        nama: r[1],
        kelas: r[2],
        sekolah: r[3],
        nilai: r[4],
        waktu: r[5]
      }))
      .filter(r => r.nama)
      .sort((a, b) => b.nilai - a.nilai)
      .slice(0, 20); // top 20

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", data: data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", data: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}