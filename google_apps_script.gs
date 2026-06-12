/**
 * FAMILY FINANCE SUPER APP - BACKEND REST API
 * Google Apps Script Controller for Google Sheets
 * Deploy as a Web App (Access: Anyone, even anonymous)
 */

function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "get_dashboard") {
    return handleGetDashboard();
  }
  if (action === "cronAutoDebit") {
    return handleCronAutoDebit();
  }
  
  return createJSONOutput({ status: "error", message: "Action doGet not found" });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJSONOutput({ status: "error", message: "Invalid JSON format" });
  }
  
  var action = data.action;
  var result;
  
  if (action === "login") {
    result = handleLogin(data.email, data.password);
  } else if (action === "add_transaksi") {
    result = handleAddTransaksi(data);
  } else if (action === "save_pocket") {
    result = handleSavePocket(data);
  } else if (action === "delete_pocket") {
    result = handleDeletePocket(data.id);
  } else if (action === "save_template") {
    result = handleSaveTemplate(data);
  } else if (action === "delete_template") {
    result = handleDeleteTemplate(data.id);
  } else if (action === "save_debt") {
    result = handleSaveDebt(data);
  } else if (action === "pay_debt") {
    result = handlePayDebt(data);
  } else if (action === "add_egg_sale") {
    result = handleAddEggSale(data);
  } else {
    result = { status: "error", message: "Action doPost '" + action + "' not recognized" };
  }
  
  return createJSONOutput(result);
}

function createJSONOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 1. AUTHENTICATION SERVICE
// ==========================================
function handleLogin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("User_Access");
  if (!sheet) return { status: "error", message: "Tab User_Access tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colUser = headers.indexOf("Username");
  var colPass = headers.indexOf("Password");
  var colNama = headers.indexOf("Nama");
  var colRole = headers.indexOf("Role");
  var colXp = headers.indexOf("Skor_Xp");
  var colLevel = headers.indexOf("Level_Saat_Ini");
  var colTitle = headers.indexOf("Title_Grade");
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colUser]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      if (String(data[i][colPass]) === String(password)) {
        return {
          status: "success",
          username: data[i][colUser],
          nama: data[i][colNama],
          role: data[i][colRole],
          xp: Number(data[i][colXp]),
          level: Number(data[i][colLevel]),
          title: data[i][colTitle]
        };
      } else {
        return { status: "error", message: "Password salah!" };
      }
    }
  }
  return { status: "error", message: "Username tidak terdaftar!" };
}

// ==========================================
// 2. DASHBOARD SYNC SERVICE
// ==========================================
function handleGetDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // A. Fetch Pockets
  var sheetPos = ss.getSheetByName("Pos_Keuangan");
  var posList = [];
  if (sheetPos) {
    var rawPos = sheetPos.getDataRange().getValues();
    var hPos = rawPos[0];
    for (var i = 1; i < rawPos.length; i++) {
      if (!rawPos[i][0]) continue;
      posList.push({
        ID_Pos: rawPos[i][hPos.indexOf("ID_Pos")],
        Nama_Pos: rawPos[i][hPos.indexOf("Nama_Pos")],
        Kategori: rawPos[i][hPos.indexOf("Kategori")],
        Pemilik: rawPos[i][hPos.indexOf("Pemilik")],
        Saldo_Saat_Ini: Number(rawPos[i][hPos.indexOf("Saldo_Saat_Ini")]),
        Kategori_Kesehatan: rawPos[i][hPos.indexOf("Kategori_Kesehatan")]
      });
    }
  }
  
  // B. Fetch Transactions and Build virtual ledger
  var sheetTx = ss.getSheetByName("Transaksi");
  var ledgerList = [];
  var rawTxs = [];
  if (sheetTx) {
    var rawTx = sheetTx.getDataRange().getValues();
    var hTx = rawTx[0];
    for (var i = 1; i < rawTx.length; i++) {
      if (!rawTx[i][0]) continue;
      var txId = rawTx[i][hTx.indexOf("ID_Transaksi")];
      var tgl = Utilities.formatDate(new Date(rawTx[i][hTx.indexOf("Tanggal")]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var jenis = rawTx[i][hTx.indexOf("Jenis")];
      var comp = rawTx[i][hTx.indexOf("Komponen_Gaji")];
      var sumber = rawTx[i][hTx.indexOf("Sumber_Pos")];
      var tujuan = rawTx[i][hTx.indexOf("Tujuan_Pos")];
      var jumlah = Number(rawTx[i][hTx.indexOf("Jumlah")]);
      var ket = rawTx[i][hTx.indexOf("Keterangan")];
      var tag = rawTx[i][hTx.indexOf("Tag_Kategori")];
      var bukti = rawTx[i][hTx.indexOf("Bukti_Transfer")];
      var oleh = rawTx[i][hTx.indexOf("Dibuat_Oleh")];
      
      rawTxs.push({
        ID_Transaksi: txId, Tanggal: tgl, Jenis: jenis, Komponen_Gaji: comp,
        Sumber_Pos: sumber, Tujuan_Pos: tujuan, Jumlah: jumlah, Keterangan: ket,
        Tag_Kategori: tag, Bukti_Transfer: bukti, Dibuat_Oleh: oleh
      });

      // Generate ledger records matching the visual tab's ledger format
      if (jenis === "Pemasukan") {
        ledgerList.push({
          ID_Pos: tujuan,
          Tipe: "Masuk",
          Vol: jumlah,
          Ket: (comp ? "Gaji: " : "") + ket,
          Time: tgl
        });
      } else if (jenis === "Pengeluaran") {
        ledgerList.push({
          ID_Pos: sumber,
          Tipe: "Keluar",
          Vol: jumlah,
          Ket: ket,
          Time: tgl
        });
      } else if (jenis === "Transfer") {
        ledgerList.push({
          ID_Pos: sumber,
          Tipe: "Keluar",
          Vol: jumlah,
          Ket: "Pindah ke " + getPosName(posList, tujuan) + ": " + ket,
          Time: tgl
        });
        ledgerList.push({
          ID_Pos: tujuan,
          Tipe: "Masuk",
          Vol: jumlah,
          Ket: "Terima dari " + getPosName(posList, sumber) + ": " + ket,
          Time: tgl
        });
      } else if (jenis === "Pinjam") {
        ledgerList.push({
          ID_Pos: tujuan,
          Tipe: "Masuk",
          Vol: jumlah,
          Ket: "Pinjam dari " + sumber + ": " + ket,
          Time: tgl
        });
      } else if (jenis === "BayarHutang") {
        ledgerList.push({
          ID_Pos: tujuan, // Tabung pembayar
          Tipe: "Keluar",
          Vol: jumlah,
          Ket: "Bayar cicilan ke " + sumber + ": " + ket,
          Time: tgl
        });
      }
    }
  }
  
  // Sort ledger by newest first
  ledgerList.reverse();
  
  // C. Fetch Templates
  var sheetTpl = ss.getSheetByName("Template_Pengeluaran");
  var tplList = [];
  if (sheetTpl) {
    var rawTpl = sheetTpl.getDataRange().getValues();
    var hTpl = rawTpl[0];
    for (var i = 1; i < rawTpl.length; i++) {
      if (!rawTpl[i][0]) continue;
      tplList.push({
        ID_Template: rawTpl[i][hTpl.indexOf("ID_Template")],
        Kategori_Template: rawTpl[i][hTpl.indexOf("Kategori_Template")],
        Nama_Template: rawTpl[i][hTpl.indexOf("Nama_Template")],
        Nominal_Standar: Number(rawTpl[i][hTpl.indexOf("Nominal_Standar")]),
        Pos_Default: rawTpl[i][hTpl.indexOf("Pos_Default")],
        Tanggal_Auto_Input: rawTpl[i][hTpl.indexOf("Tanggal_Auto_Input")] ? Number(rawTpl[i][hTpl.indexOf("Tanggal_Auto_Input")]) : null,
        Status_Auto: rawTpl[i][hTpl.indexOf("Status_Auto")]
      });
    }
  }

  // D. Fetch Settings
  var sheetSettings = ss.getSheetByName("Settings_App");
  var settingsObj = {};
  if (sheetSettings) {
    var rawSettings = sheetSettings.getDataRange().getValues();
    for (var i = 1; i < rawSettings.length; i++) {
      if (rawSettings[i][0]) {
        settingsObj[rawSettings[i][0]] = rawSettings[i][1];
      }
    }
  }

  // E. Fetch Debts
  var sheetDebts = ss.getSheetByName("Utang_Piutang");
  var debtsList = [];
  if (sheetDebts) {
    var rawDebts = sheetDebts.getDataRange().getValues();
    var hDebts = rawDebts[0];
    for (var i = 1; i < rawDebts.length; i++) {
      if (!rawDebts[i][0]) continue;
      debtsList.push({
        ID_Utang: rawDebts[i][hDebts.indexOf("ID_Utang")],
        Pemberi_Pinjaman: rawDebts[i][hDebts.indexOf("Pemberi_Pinjaman")],
        Penerima_Pinjaman: rawDebts[i][hDebts.indexOf("Penerima_Pinjaman")],
        Total_Pinjaman: Number(rawDebts[i][hDebts.indexOf("Total_Pinjaman")]),
        Tenor_Cicilan: Number(rawDebts[i][hDebts.indexOf("Tenor_Cicilan")]),
        Sisa_Hutang: Number(rawDebts[i][hDebts.indexOf("Sisa_Hutang")]),
        Status: rawDebts[i][hDebts.indexOf("Status")]
      });
    }
  }

  // F. Fetch Installments Log
  var sheetLog = ss.getSheetByName("Log_Cicilan");
  var logList = [];
  if (sheetLog) {
    var rawLog = sheetLog.getDataRange().getValues();
    var hLog = rawLog[0];
    for (var i = 1; i < rawLog.length; i++) {
      if (!rawLog[i][0]) continue;
      logList.push({
        ID_Cicilan: rawLog[i][hLog.indexOf("ID_Cicilan")],
        ID_Utang: rawLog[i][hLog.indexOf("ID_Utang")],
        Cicilan_Ke: Number(rawLog[i][hLog.indexOf("Cicilan_Ke")]),
        Tanggal_Bayar: Utilities.formatDate(new Date(rawLog[i][hLog.indexOf("Tanggal_Bayar")]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        Jumlah_Bayar: Number(rawLog[i][hLog.indexOf("Jumlah_Bayar")]),
        Sisa_Utang_Baru: Number(rawLog[i][hLog.indexOf("Sisa_Utang_Baru")]),
        Bukti_SS_Transfer: rawLog[i][hLog.indexOf("Bukti_SS_Transfer")],
        Dibuat_Oleh: rawLog[i][hLog.indexOf("Dibuat_Oleh")]
      });
    }
  }

  // G. Fetch Egg Sales
  var sheetEgg = ss.getSheetByName("Penjualan_Telur");
  var eggList = [];
  if (sheetEgg) {
    var rawEgg = sheetEgg.getDataRange().getValues();
    var hEgg = rawEgg[0];
    for (var i = 1; i < rawEgg.length; i++) {
      if (!rawEgg[i][0]) continue;
      eggList.push({
        ID_Jual: rawEgg[i][hEgg.indexOf("ID_Jual")],
        Tanggal: Utilities.formatDate(new Date(rawEgg[i][hEgg.indexOf("Tanggal")]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        Nama_Pembeli: rawEgg[i][hEgg.indexOf("Nama_Pembeli")],
        Jumlah_Telur: Number(rawEgg[i][hEgg.indexOf("Jumlah_Telur")]),
        Total_Harga: Number(rawEgg[i][hEgg.indexOf("Total_Harga")]),
        Status_Bayar: rawEgg[i][hEgg.indexOf("Status_Bayar")],
        Pos_Masuk: rawEgg[i][hEgg.indexOf("Pos_Masuk")]
      });
    }
  }

  return {
    status: "success",
    pos: posList,
    ledger: ledgerList,
    raw_txs: rawTxs,
    templates: tplList,
    settings: settingsObj,
    debts: debtsList,
    installments: logList,
    egg_sales: eggList
  };
}

function getPosName(posList, id) {
  for (var i = 0; i < posList.length; i++) {
    if (posList[i].ID_Pos === id) return posList[i].Nama_Pos;
  }
  return id;
}

// ==========================================
// 3. TRANSACTION ENGINE
// ==========================================
function handleAddTransaksi(tx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTx = ss.getSheetByName("Transaksi");
  if (!sheetTx) return { status: "error", message: "Tab Transaksi tidak ditemukan" };
  
  var id = tx.id || "TX" + new Date().getTime() + Math.floor(Math.random() * 1000);
  var tgl = tx.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var jenis = tx.jenis; // Pengeluaran, Pemasukan, Transfer, Pinjam, BayarHutang
  var comp = tx.komponen_gaji || ""; // JSON string
  var sumber = tx.sumber_pos || "";
  var tujuan = tx.tujuan_pos || "";
  var jumlah = Number(tx.jumlah);
  var ket = tx.keterangan || "";
  var tag = tx.tag_kategori || "Umum";
  var bukti = tx.bukti_transfer || ""; // Base64
  var oleh = tx.dibuat_oleh || "System";
  
  // Write transaction row
  sheetTx.appendRow([id, tgl, jenis, comp, sumber, tujuan, jumlah, ket, tag, bukti, oleh]);
  
  // Update Balances dynamically based on transaction types
  var sheetPos = ss.getSheetByName("Pos_Keuangan");
  if (sheetPos) {
    var dataPos = sheetPos.getDataRange().getValues();
    var hPos = dataPos[0];
    var colId = hPos.indexOf("ID_Pos");
    var colBal = hPos.indexOf("Saldo_Saat_Ini");
    
    for (var i = 1; i < dataPos.length; i++) {
      // Deduct from Source
      if (sumber && dataPos[i][colId] === sumber) {
        var oldBal = Number(dataPos[i][colBal]);
        sheetPos.getRange(i + 1, colBal + 1).setValue(oldBal - jumlah);
      }
      // Add to Destination
      if (tujuan && dataPos[i][colId] === tujuan) {
        var oldBal = Number(dataPos[i][colBal]);
        sheetPos.getRange(i + 1, colBal + 1).setValue(oldBal + jumlah);
      }
    }
  }

  // Award +10 XP to user
  var xpAwarded = awardXp(oleh, 10);
  
  return { 
    status: "success", 
    message: "Transaksi berhasil dicatat!", 
    id: id,
    xp_gained: 10,
    current_xp: xpAwarded.xp,
    current_level: xpAwarded.level,
    level_up: xpAwarded.level_up
  };
}

function awardXp(username, amount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("User_Access");
  if (!sheet) return { xp: 0, level: 1, level_up: false };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colUser = headers.indexOf("Username");
  var colXp = headers.indexOf("Skor_Xp");
  var colLevel = headers.indexOf("Level_Saat_Ini");
  var colTitle = headers.indexOf("Title_Grade");
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colUser]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      var xp = Number(data[i][colXp]) + amount;
      var level = Number(data[i][colLevel]);
      var levelUp = false;
      
      // Level up checks (every 100 XP triggers a new level)
      if (xp >= 100) {
        level += Math.floor(xp / 100);
        xp = xp % 100;
        levelUp = true;
      }
      
      // Update grade title dynamically
      var title = "Financial Newbie";
      if (level >= 10) title = "Wealth Master";
      else if (level >= 5) title = "Budget General";
      else if (level >= 2) title = "Savings Knight";
      
      sheet.getRange(i + 1, colXp + 1).setValue(xp);
      sheet.getRange(i + 1, colLevel + 1).setValue(level);
      sheet.getRange(i + 1, colTitle + 1).setValue(title);
      
      return { xp: xp, level: level, level_up: levelUp };
    }
  }
  return { xp: 0, level: 1, level_up: false };
}

// ==========================================
// 4. POCKETS MANAGEMENT (CRUD)
// ==========================================
function handleSavePocket(pos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pos_Keuangan");
  if (!sheet) return { status: "error", message: "Tab Pos_Keuangan tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colId = headers.indexOf("ID_Pos");
  var colNama = headers.indexOf("Nama_Pos");
  var colKat = headers.indexOf("Kategori");
  var colPemilik = headers.indexOf("Pemilik");
  var colBal = headers.indexOf("Saldo_Saat_Ini");
  var colHealth = headers.indexOf("Kategori_Kesehatan");
  
  var id = pos.id;
  var name = pos.nama;
  var kat = pos.kategori; // Utama, Tabungan, Bisnis
  var pemilik = pos.pemilik || "Bersama";
  var health = pos.kategori_kesehatan || "Umum";
  
  // If editing an existing pocket
  if (id) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][colId] === id) {
        sheet.getRange(i + 1, colNama + 1).setValue(name);
        sheet.getRange(i + 1, colKat + 1).setValue(kat);
        sheet.getRange(i + 1, colPemilik + 1).setValue(pemilik);
        sheet.getRange(i + 1, colHealth + 1).setValue(health);
        return { status: "success", message: "Kantong berhasil diperbarui!" };
      }
    }
  }
  
  // If adding new pocket
  var newId = "P" + String(data.length).padStart(3, '0') + Math.floor(Math.random() * 10);
  sheet.appendRow([newId, name, kat, pemilik, 0, health]);
  return { status: "success", message: "Kantong baru berhasil dibuat!", id: newId };
}

function handleDeletePocket(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pos_Keuangan");
  if (!sheet) return { status: "error", message: "Tab Pos_Keuangan tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colId = headers.indexOf("ID_Pos");
  var colBal = headers.indexOf("Saldo_Saat_Ini");
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][colId] === id) {
      var balance = Number(data[i][colBal]);
      
      // Re-route balance to Main Pocket (P001) if deleting a sub-pocket
      if (id !== "P001" && balance > 0) {
        for (var j = 1; j < data.length; j++) {
          if (data[j][colId] === "P001") {
            var mainBal = Number(data[j][colBal]);
            sheet.getRange(j + 1, colBal + 1).setValue(mainBal + balance);
            
            // Record transfer transaction to show trace
            var sheetTx = ss.getSheetByName("Transaksi");
            if (sheetTx) {
              sheetTx.appendRow([
                "TX_DEL_" + new Date().getTime(),
                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
                "Transfer",
                "",
                id,
                "P001",
                balance,
                "Auto cleanup balance from deleted pocket: " + data[i][headers.indexOf("Nama_Pos")],
                "System",
                "",
                "System"
              ]);
            }
            break;
          }
        }
      }
      
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Kantong berhasil dihapus!" };
    }
  }
  return { status: "error", message: "Kantong tidak ditemukan" };
}

// ==========================================
// 5. TEMPLATES MANAGEMENT (CRUD)
// ==========================================
function handleSaveTemplate(tpl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Template_Pengeluaran");
  if (!sheet) return { status: "error", message: "Tab Template_Pengeluaran tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colId = headers.indexOf("ID_Template");
  var colKat = headers.indexOf("Kategori_Template");
  var colNama = headers.indexOf("Nama_Template");
  var colNom = headers.indexOf("Nominal_Standar");
  var colPos = headers.indexOf("Pos_Default");
  var colTgl = headers.indexOf("Tanggal_Auto_Input");
  var colStatus = headers.indexOf("Status_Auto");
  
  var id = tpl.id;
  var kat = tpl.kategori_template; // Pengeluaran, Pemasukan
  var nama = tpl.nama_template;
  var nominal = Number(tpl.nominal_standar);
  var posDef = tpl.pos_default;
  var autoDay = tpl.tanggal_auto_input ? Number(tpl.tanggal_auto_input) : "";
  var statusAuto = tpl.status_auto || "Manual"; // Manual, Auto
  
  if (id) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][colId] === id) {
        sheet.getRange(i + 1, colKat + 1).setValue(kat);
        sheet.getRange(i + 1, colNama + 1).setValue(nama);
        sheet.getRange(i + 1, colNom + 1).setValue(nominal);
        sheet.getRange(i + 1, colPos + 1).setValue(posDef);
        sheet.getRange(i + 1, colTgl + 1).setValue(autoDay);
        sheet.getRange(i + 1, colStatus + 1).setValue(statusAuto);
        return { status: "success", message: "Template berhasil diperbarui!" };
      }
    }
  }
  
  var newId = "T" + String(data.length).padStart(3, '0') + Math.floor(Math.random() * 10);
  sheet.appendRow([newId, kat, nama, nominal, posDef, autoDay, statusAuto]);
  return { status: "success", message: "Template berhasil dibuat!", id: newId };
}

function handleDeleteTemplate(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Template_Pengeluaran");
  if (!sheet) return { status: "error", message: "Tab Template_Pengeluaran tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var colId = data[0].indexOf("ID_Template");
  for (var i = 1; i < data.length; i++) {
    if (data[i][colId] === id) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Template berhasil dihapus!" };
    }
  }
  return { status: "error", message: "Template tidak ditemukan" };
}

// ==========================================
// 6. UTANG PIUTANG SERVICES
// ==========================================
function handleSaveDebt(debt) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Utang_Piutang");
  if (!sheet) return { status: "error", message: "Tab Utang_Piutang tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colId = headers.indexOf("ID_Utang");
  
  var newId = "U" + String(data.length).padStart(3, '0');
  var total = Number(debt.total_pinjaman);
  
  // ID_Utang, Pemberi_Pinjaman, Penerima_Pinjaman, Total_Pinjaman, Tenor_Cicilan, Sisa_Hutang, Status
  sheet.appendRow([
    newId,
    debt.pemberi_pinjaman,
    debt.penerima_pinjaman,
    total,
    Number(debt.tenor_cicilan),
    total, // Sisa hutang awal = total pinjaman
    "Belum Lunas"
  ]);
  
  // Record dynamic incoming transaction if it affects a pocket (represented as Tujuan_Pos)
  if (debt.pos_masuk) {
    var txObj = {
      jenis: "Pinjam",
      sumber_pos: debt.pemberi_pinjaman,
      tujuan_pos: debt.pos_masuk,
      jumlah: total,
      keterangan: "Pinjaman masuk dari " + debt.pemberi_pinjaman,
      tag_kategori: "Hutang",
      dibuat_oleh: debt.dibuat_oleh
    };
    handleAddTransaksi(txObj);
  }
  
  return { status: "success", message: "Catatan utang berhasil didaftarkan!", id: newId };
}

function handlePayDebt(payment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDebts = ss.getSheetByName("Utang_Piutang");
  var sheetLog = ss.getSheetByName("Log_Cicilan");
  if (!sheetDebts || !sheetLog) return { status: "error", message: "Tab Utang/Cicilan tidak ditemukan" };
  
  var debtId = payment.id_utang;
  var bayar = Number(payment.jumlah_bayar);
  var bukti = payment.bukti_ss_transfer || "";
  var oleh = payment.dibuat_oleh || "System";
  
  // 1. Fetch debt details
  var dataDebts = sheetDebts.getDataRange().getValues();
  var hDebts = dataDebts[0];
  var colId = hDebts.indexOf("ID_Utang");
  var colSisa = hDebts.indexOf("Sisa_Hutang");
  var colStatus = hDebts.indexOf("Status");
  var colPemberi = hDebts.indexOf("Pemberi_Pinjaman");
  
  var sisaLama = 0;
  var pemberi = "";
  var rowIdx = -1;
  
  for (var i = 1; i < dataDebts.length; i++) {
    if (dataDebts[i][colId] === debtId) {
      sisaLama = Number(dataDebts[i][colSisa]);
      pemberi = dataDebts[i][colPemberi];
      rowIdx = i + 1;
      break;
    }
  }
  
  if (rowIdx === -1) return { status: "error", message: "ID Utang tidak ditemukan" };
  if (sisaLama <= 0) return { status: "error", message: "Utang sudah lunas!" };
  
  // 2. Compute remaining
  var sisaBaru = Math.max(0, sisaLama - bayar);
  var statusBaru = sisaBaru === 0 ? "Lunas" : "Belum Lunas";
  
  // Update Utang_Piutang sheet
  sheetDebts.getRange(rowIdx, colSisa + 1).setValue(sisaBaru);
  sheetDebts.getRange(rowIdx, colStatus + 1).setValue(statusBaru);
  
  // 3. Append to Log_Cicilan
  var dataLog = sheetLog.getDataRange().getValues();
  var newCicilanId = "C" + String(dataLog.length).padStart(3, '0');
  
  // Count how many payments exist for this debt
  var cicilanKe = 1;
  for (var i = 1; i < dataLog.length; i++) {
    if (dataLog[i][1] === debtId) {
      cicilanKe++;
    }
  }
  
  // ID_Cicilan, ID_Utang, Cicilan_Ke, Tanggal_Bayar, Jumlah_Bayar, Sisa_Utang_Baru, Bukti_SS_Transfer, Dibuat_Oleh
  sheetLog.appendRow([
    newCicilanId,
    debtId,
    cicilanKe,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    bayar,
    sisaBaru,
    bukti,
    oleh
  ]);
  
  // 4. Record dynamic transaction if pocket is used to pay (Deduct from payment.pos_sumber)
  if (payment.pos_sumber) {
    var txObj = {
      jenis: "BayarHutang",
      sumber_pos: pemberi, // Label pemberi pinjaman
      tujuan_pos: payment.pos_sumber, // Pos pembayar (e.g. P001) yang saldonya dikurangi
      jumlah: bayar,
      keterangan: "Cicilan ke-" + cicilanKe + " untuk utang " + debtId,
      tag_kategori: "Cicilan",
      dibuat_oleh: oleh
    };
    handleAddTransaksi(txObj);
  }
  
  return { status: "success", message: "Pembayaran cicilan berhasil dicatat!", sisa_hutang: sisaBaru, status: statusBaru };
}

// ==========================================
// 7. EGG SALES (BUSINESS LEDGER)
// ==========================================
function handleAddEggSale(sale) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Penjualan_Telur");
  if (!sheet) return { status: "error", message: "Tab Penjualan_Telur tidak ditemukan" };
  
  var data = sheet.getDataRange().getValues();
  var newId = "TLR" + String(data.length).padStart(3, '0') + Math.floor(Math.random() * 10);
  var tgl = sale.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var qty = Number(sale.jumlah_telur);
  var total = Number(sale.total_harga);
  
  // ID_Jual, Tanggal, Nama_Pembeli, Jumlah_Telur, Total_Harga, Status_Bayar, Pos_Masuk
  sheet.appendRow([
    newId,
    tgl,
    sale.nama_pembeli,
    qty,
    total,
    sale.status_bayar || "Lunas",
    sale.pos_masuk || "P001"
  ]);
  
  // Record dynamic incoming transaction to the pocket
  if (sale.pos_masuk && sale.status_bayar === "Lunas") {
    var txObj = {
      jenis: "Pemasukan",
      sumber_pos: "EKSTERNAL",
      tujuan_pos: sale.pos_masuk,
      jumlah: total,
      keterangan: "Jual " + qty + " telur ke " + sale.nama_pembeli,
      tag_kategori: "Bisnis",
      dibuat_oleh: sale.dibuat_oleh || "System"
    };
    handleAddTransaksi(txObj);
  }
  
  return { status: "success", message: "Penjualan telur berhasil dicatat!", id: newId };
}

// ==========================================
// 8. AUTO-DEBIT SCHEDULER (CRON CRON)
// ==========================================
function handleCronAutoDebit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTpl = ss.getSheetByName("Template_Pengeluaran");
  if (!sheetTpl) return { status: "error", message: "Tab Template_Pengeluaran tidak ditemukan" };
  
  var today = new Date().getDate(); // Current day of the month (1-31)
  var rawTpls = sheetTpl.getDataRange().getValues();
  var hTpl = rawTpls[0];
  
  var colId = hTpl.indexOf("ID_Template");
  var colKat = hTpl.indexOf("Kategori_Template");
  var colNama = hTpl.indexOf("Nama_Template");
  var colNom = hTpl.indexOf("Nominal_Standar");
  var colPos = hTpl.indexOf("Pos_Default");
  var colTgl = hTpl.indexOf("Tanggal_Auto_Input");
  var colStatus = hTpl.indexOf("Status_Auto");
  
  var processed = 0;
  
  for (var i = 1; i < rawTpls.length; i++) {
    if (rawTpls[i][colStatus] === "Auto" && Number(rawTpls[i][colTgl]) === today) {
      var kat = rawTpls[i][colKat];
      var name = rawTpls[i][colNama];
      var amount = Number(rawTpls[i][colNom]);
      var pos = rawTpls[i][colPos];
      
      // Auto-execute transaction
      var txObj = {
        jenis: kat,
        sumber_pos: kat === "Pengeluaran" ? pos : "",
        tujuan_pos: kat === "Pemasukan" ? pos : "",
        jumlah: amount,
        keterangan: "Auto Debit: " + name,
        tag_kategori: "Auto-Debit",
        dibuat_oleh: "Auto-Cron"
      };
      
      handleAddTransaksi(txObj);
      processed++;
    }
  }
  
  return { status: "success", message: "Cron auto-debit completed", processed_templates: processed, date: today };
}
