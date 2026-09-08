/**
 * GROUND CHECK PENDATAAN DESA CANTIK - BACKEND (Apps Script)
 * ---------------------------------------------------------
 * Cara pakai:
 * 1. Buka Google Sheet tujuan (yang sudah dibuat di Drive).
 * 2. Extensions > Apps Script.
 * 3. Hapus isi default Code.gs, tempel SELURUH isi file ini.
 * 4. Save (Ctrl+S).
 * 5. Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Klik Deploy, salin "Web app URL" yang muncul.
 * 7. Tempel URL itu ke js/config.js (APPS_SCRIPT_URL) di project survei.
 *
 * Setiap kali kode ini diubah, harus bikin deployment baru (New deployment),
 * bukan cuma Save, supaya perubahan ikut aktif di URL yang sama
 * (atau pilih "Manage deployments" > edit > New version).
 */

// Header kolom untuk tiap sheet. Urutan ini yang dipakai saat menulis baris baru.
const SHEET_HEADERS = {
  Rumah_Tangga: [
    "timestamp", "pic_nama", "pic_telp", "pic_email",
    "nama_kepala_keluarga", "nik_kepala_keluarga", "no_kk",
    "provinsi", "kabupaten_kota", "kecamatan", "kelurahan_desa", "rw", "rt", "alamat_lengkap",
    "nama_informan", "nama_petugas",
    "q201_status_kepemilikan_bangunan", "q202_status_kepemilikan_lahan", "q203_luas_lantai_m2",
    "q204_jenis_lantai", "q205_jenis_dinding", "q206_jenis_atap", "q207_sumber_air_minum",
    "q208_fasilitas_bab", "q209_jenis_kloset", "q210_pembuangan_akhir_tinja",
    "art1_nama", "art1_deskripsi_pekerjaan", "art1_penghasilan",
    "art2_nama", "art2_deskripsi_pekerjaan", "art2_penghasilan",
    "art3_nama", "art3_deskripsi_pekerjaan", "art3_penghasilan",
    "catatan", "foto_links",
  ],
  Individu_Anggota: [
    "timestamp", "pic_nama", "pic_telp", "pic_email",
    "nama_kepala_keluarga", "nik_kepala_keluarga", "no_kk",
    "provinsi", "kabupaten_kota", "kecamatan", "kelurahan_desa", "rw", "rt", "alamat_lengkap",
    "nama_informan", "nama_petugas",
    "anggota_ke", "nama_anggota", "nik_anggota", "jenis_kelamin", "tanggal_lahir", "umur", "hubungan_kk",
    "partisipasi_sekolah", "pendidikan_tertinggi", "kelas_tertinggi", "ijazah_sttb",
    "bekerja", "lapangan_usaha", "deskripsi_pekerjaan", "status_pekerjaan_utama", "pendapatan_sebulan",
    "catatan", "foto_links",
  ],
};

const FOTO_FOLDER_NAME = "Foto_GroundCheck_DesaCantik";

function doGet(e) {
  return ContentService.createTextOutput(
    "Ground Check Desa Cantik - Apps Script aktif.",
  ).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheetName = body.sheet;
    const data = body.data || {};
    const photos = body.photos || [];

    const headers = SHEET_HEADERS[sheetName];
    if (!headers) {
      return jsonResponse({ success: false, error: "Nama sheet tidak dikenali: " + sheetName });
    }

    const sheet = getOrCreateSheet(sheetName, headers);

    // Simpan foto ke Drive (kalau ada), lalu gabungkan link-nya
    let fotoLinks = "";
    if (photos.length > 0) {
      fotoLinks = savePhotosToDrive(photos, sheetName, data);
    }

    // Susun baris sesuai urutan header
    const rowValues = headers.map((h) => {
      if (h === "timestamp") return new Date();
      if (h === "foto_links") return fotoLinks;
      return data[h] !== undefined && data[h] !== null ? data[h] : "";
    });

    sheet.appendRow(rowValues);

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#6C5CE7")
      .setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);

    // Hapus sheet default kosong "Sheet1" kalau masih ada dan bukan sheet ini
    const defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet && ss.getSheets().length > 1) {
      const isEmpty = defaultSheet.getLastRow() === 0;
      if (isEmpty) ss.deleteSheet(defaultSheet);
    }
  }

  return sheet;
}

function savePhotosToDrive(photos, sheetName, data) {
  const root = DriveApp.getRootFolder();
  let folder;
  const folders = root.getFoldersByName(FOTO_FOLDER_NAME);
  folder = folders.hasNext() ? folders.next() : root.createFolder(FOTO_FOLDER_NAME);

  // Subfolder per submission biar rapi: nama KK + waktu
  const label = (data.nama_kepala_keluarga || "tanpa_nama") + "_" + new Date().getTime();
  const subFolder = folder.createFolder(label);

  const links = [];
  photos.forEach((p) => {
    try {
      const base64 = p.base64.split(",").pop(); // buang prefix "data:image/...;base64,"
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        p.mimeType || "image/jpeg",
        p.filename || "foto.jpg",
      );
      const file = subFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      links.push(file.getUrl());
    } catch (err) {
      links.push("GAGAL_UPLOAD:" + (p.filename || "unknown"));
    }
  });

  return links.join(", ");
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
