// =====================================================================
// SurveyStore - penyimpanan sementara data survei antar-halaman
// Karena situs ini multi-page (bukan SPA), data tiap step disimpan ke
// localStorage supaya tidak hilang saat pindah halaman, lalu dikumpulkan
// jadi satu payload di halaman terakhir (Catatan) untuk dikirim ke Sheets.
// =====================================================================
const SurveyStore = {
  KEY: "gcdc_survey_data",

  getAll() {
    try {
      return JSON.parse(localStorage.getItem(this.KEY)) || {};
    } catch (e) {
      return {};
    }
  },

  // Gabungkan (merge) field-field baru ke data yang sudah ada
  save(partial) {
    const current = this.getAll();
    const merged = Object.assign({}, current, partial);
    localStorage.setItem(this.KEY, JSON.stringify(merged));
    return merged;
  },

  clear() {
    localStorage.removeItem(this.KEY);
  },

  // Ambil data PIC yang disimpan di index2.html (key lama: data_pic)
  getPic() {
    try {
      return JSON.parse(localStorage.getItem("data_pic")) || {};
    } catch (e) {
      return {};
    }
  },

  // Konversi File -> base64 data URL, dipakai untuk kirim foto ke Apps Script
  fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  },

  // Kompres & resize foto sebelum disimpan ke localStorage / dikirim ke server.
  // Foto asli dari kamera HP bisa 3-8MB; localStorage browser biasanya cuma
  // muat sekitar 5-10MB total, jadi foto WAJIB dikecilkan dulu.
  // Hasilnya berupa data URL JPEG (base64), max lebar 1280px, quality 0.7.
  compressImage(file, maxWidth = 1280, quality = 0.7) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const reader = new FileReader();

      reader.onload = (e) => {
        img.onload = () => {
          const scale = Math.min(1, maxWidth / img.width);
          const canvas = document.createElement("canvas");
          canvas.width = img.width * scale;
          canvas.height = img.height * scale;
          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          resolve(canvas.toDataURL("image/jpeg", quality));
        };
        img.onerror = reject;
        img.src = e.target.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  },

  // Simpan array foto (sudah dikompres, format {filename, mimeType, base64})
  // ke localStorage terpisah dari data teks, supaya gampang dibersihkan.
  savePhotos(photos) {
    localStorage.setItem("gcdc_survey_photos", JSON.stringify(photos));
  },

  getPhotos() {
    try {
      return JSON.parse(localStorage.getItem("gcdc_survey_photos")) || [];
    } catch (e) {
      return [];
    }
  },

  clearPhotos() {
    localStorage.removeItem("gcdc_survey_photos");
  },

  // Kirim data ke Apps Script Web App.
  // sheetName: "Rumah_Tangga" atau "Individu_Anggota"
  // dataObject: object dengan key = nama kolom di sheet
  // photos: array {filename, mimeType, base64} hasil compressImage() (opsional)
  async submit(sheetName, dataObject, photos) {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        sheet: sheetName,
        data: dataObject,
        photos: photos || [],
      }),
    });

    return res.json();
  },
};
