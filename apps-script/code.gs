function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Aplikasi Pengeluaran Kapal');
}

/* ================= LOGIN ================= */

function login(username, password) {
  const sh = SpreadsheetApp.getActive().getSheetByName('MASTER_USER');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { success: true, role: data[i][2], user: username };
    }
  }
  return { success: false };
}

/* ================= MASTER ================= */

function getMaster() {
  const ss = SpreadsheetApp.getActive();
  return {
    kapal: ss.getSheetByName('MASTER_KAPAL').getDataRange().getValues(),
    pos: ss.getSheetByName('MASTER_POS').getDataRange().getValues()
  };
}

/* ================= PENGELUARAN ================= */

function simpanPengeluaran(d) {
  SpreadsheetApp.getActive()
    .getSheetByName('PENGELUARAN_HARIAN')
    .appendRow([
      d.tanggal, d.kapal, d.kategori, d.pos,
      d.nominal, d.keterangan, d.user, new Date()
    ]);
}

function getPengeluaran() {
  return SpreadsheetApp.getActive()
    .getSheetByName('PENGELUARAN_HARIAN')
    .getDataRange().getValues();
}

function updatePengeluaran(row, data) {
  SpreadsheetApp.getActive()
    .getSheetByName('PENGELUARAN_HARIAN')
    .getRange(row, 1, 1, data.length)
    .setValues([data]);
}

function deletePengeluaran(row) {
  SpreadsheetApp.getActive()
    .getSheetByName('PENGELUARAN_HARIAN')
    .deleteRow(row);
}

/* ================= GUEST ================= */

function getProfil() {
  const ss = SpreadsheetApp.getActive();
  return {
    kapal: ss.getSheetByName('PROFIL_KAPAL').getDataRange().getValues(),
    pengurus: ss.getSheetByName('PENGURUS_KAPAL').getDataRange().getValues()
  };
}
