// --- UI & Menu ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Kirim Email Otomatis')
    .addItem('Buka Pengaturan', 'showSettingsDialog')
    .addToUi();
}

function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(1200) // Ubah lebar agar editor HTML terlihat lebih baik
    .setHeight(900); // Ubah tinggi
  SpreadsheetApp.getUi().showModalDialog(html, 'Pengaturan Kirim Email Otomatis');
}

// --- Data Spreadsheet ---
function getSheetHeaders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers;
  } catch (e) {
    logError(`Gagal mendapatkan header kolom: ${e}`);
    return [];
  }
}


// --- Template Email ---
function saveTemplate(templateName, templateBody) {
  try {
    const templateMap = loadTemplates() || {};
    templateMap[templateName] = templateBody;
    PropertiesService.getUserProperties().setProperty("emailTemplates", JSON.stringify(templateMap));
    return { status: 'success', message: 'Template berhasil disimpan.' };
  } catch (e) {
    logError(`Gagal menyimpan template: ${e}`);
    return { status: 'error', message: `Gagal menyimpan template: ${e}` };
  }
}


function loadTemplates() {
  try {
    const templateJson = PropertiesService.getUserProperties().getProperty("emailTemplates");
    if (templateJson) {
      return JSON.parse(templateJson);
    } else {
      return {};
    }
  } catch (e) {
    logError(`Gagal memuat template: ${e}`);
    return {};
  }
}


function deleteTemplate(templateName) {
  try {
    const templateMap = loadTemplates() || {};
    delete templateMap[templateName];
    PropertiesService.getUserProperties().setProperty("emailTemplates", JSON.stringify(templateMap));
    return { status: 'success', message: 'Template berhasil dihapus.' };
  } catch (e) {
    logError(`Gagal menghapus template: ${e}`);
    return { status: 'error', message: `Gagal menghapus template: ${e}` };
  }

}



// ---  Pengaturan Output File PDF ---
function selectFolder() {
  const folder = SpreadsheetApp.getUi().selectFolder('Pilih Folder Tujuan');
  return folder ? folder.getId() : null;
}

function getFolderName(folderId) {
  if (!folderId) return null;
  try {
    const folder = DriveApp.getFolderById(folderId);
    return folder.getName()
  }
  catch (e) {
    logError(`Gagal mendapatkan nama folder: ${e}`);
    return null;
  }
}

// --- Pengiriman Email ---
function sendEmails(config) {
  let logData = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNumber = i + 2;
    let status = 'Berhasil';
    let message = '';
    let emailDetails = {
      to: '',
      subject: '',
      body: '',
      cc: '',
      bcc: '',
    };

    try {
      // ---  Kondisi Pengiriman ---
      if (config.conditionalEnabled) {
        const conditionColumnIndex = headers.indexOf(config.conditionColumn);
        if (conditionColumnIndex === -1) {
          status = 'Gagal';
          message = 'Kolom kondisi tidak ditemukan';
          logData.push({
            time: new Date(),
            row: rowNumber,
            ...emailDetails,
            status: status,
            message: message
          });
          continue;
        }

        const rowValue = row[conditionColumnIndex];
        const conditionValue = config.conditionValue;
        const operator = config.conditionOperator;

        let conditionMet = false;
        switch (operator) {
          case 'sama dengan': conditionMet = rowValue === conditionValue; break;
          case 'tidak sama dengan': conditionMet = rowValue !== conditionValue; break;
          case 'berisi': conditionMet = String(rowValue).includes(conditionValue); break;
          case 'tidak berisi': conditionMet = !String(rowValue).includes(conditionValue); break;
          case 'lebih besar dari': conditionMet = rowValue > conditionValue; break;
          case 'lebih kecil dari': conditionMet = rowValue < conditionValue; break;
          default: conditionMet = false;
        }


        if (!conditionMet) {
          status = 'Diabaikan (Kondisi Tidak Terpenuhi)';
          logData.push({
            time: new Date(),
            row: rowNumber,
            ...emailDetails,
            status: status,
            message: message
          });
          continue;
        }

      }

      // --- Ambil data email ---
      const toColumnIndex = headers.indexOf(config.toColumn);
      const ccColumnIndex = config.ccColumn === "Tidak Ada" ? -1 : headers.indexOf(config.ccColumn);
      const bccColumnIndex = config.bccColumn === "Tidak Ada" ? -1 : headers.indexOf(config.bccColumn);
      const subjectColumnIndex = config.subjectColumn === "Tidak Ada" ? -1 : headers.indexOf(config.subjectColumn);

      if (toColumnIndex === -1) {
        status = 'Gagal';
        message = 'Kolom penerima email tidak ditemukan.';
        logData.push({
          time: new Date(),
          row: rowNumber,
          ...emailDetails,
          status: status,
          message: message
        });
        continue;
      }


      const toEmail = String(row[toColumnIndex]).trim();
      const ccEmails = ccColumnIndex !== -1 ? String(row[ccColumnIndex]).trim() : config.defaultCC;
      const bccEmails = bccColumnIndex !== -1 ? String(row[bccColumnIndex]).trim() : config.defaultBCC;
      const subject = subjectColumnIndex !== -1 ? String(row[subjectColumnIndex]).trim() : config.defaultSubject;

      // --- Output File PDF ---
      let pdfBlob = null;
      if (config.pdfEnabled) {
        try {
          const filename = getPdfFilename(config.pdfFilename, row, headers);
          pdfBlob = createPdf(config.emailBody, filename);
        } catch (pdfError) {
          status = "Gagal";
          message = `Gagal membuat file PDF: ${pdfError}`
        }
      }

      emailDetails = {
        to: toEmail,
        subject: subject,
        body: config.emailBody,
        cc: ccEmails,
        bcc: bccEmails
      };


      // --- Kirim Email ---
      MailApp.sendEmail({
        to: toEmail,
        subject: subject,
        htmlBody: config.emailBody,
        cc: ccEmails,
        bcc: bccEmails,
        attachments: pdfBlob ? [pdfBlob] : [],
        replyTo: config.replyTo,
        from: config.from
      });


      if (status === 'Berhasil') {
        message = "Email berhasil dikirim."
      }

    } catch (e) {
      status = 'Gagal';
      message = `Gagal mengirim email: ${e}`
      logError(`Gagal mengirim email baris ke-${rowNumber}: ${e}`);
    } finally {
      logData.push({
        time: new Date(),
        row: rowNumber,
        ...emailDetails,
        status: status,
        message: message
      });

    }
  }

  logDataToSheet(logData);
  return { status: 'success', message: 'Email telah berhasil dikirimkan. Periksa log untuk detailnya' };
}


//--- Get PDF Filename ---
function getPdfFilename(filenameTemplate, row, headers) {
  if (!filenameTemplate) return "output.pdf";
  let filename = filenameTemplate;
  for (let i = 0; i < headers.length; i++) {
    const mergeTag = `<<${headers[i]}>>`;
    filename = filename.replace(mergeTag, row[i]);
  }

  return filename + ".pdf";
}

// --- Create PDF---
function createPdf(content, filename) {
  try {
    const htmlOutput = HtmlService.createHtmlOutput(content);
    const pdfBlob = htmlOutput.getAs(MimeType.PDF);
    pdfBlob.setName(filename);
    return pdfBlob;
  } catch (e) {
    logError(`Gagal membuat PDF: ${e}`);
    throw new Error(`Gagal membuat PDF: ${e}`)
  }
}


// --- Log ke Sheet ---
function logDataToSheet(logData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Log Email');
    if (!logSheet) {
      logSheet = ss.insertSheet('Log Email');
      logSheet.appendRow(['Waktu', 'Baris', 'Email Tujuan', 'Subjek', 'Body', 'CC', 'BCC', 'Status', 'Pesan']);
    }

    logData.forEach(log => {
      logSheet.appendRow([
        log.time,
        log.row,
        log.to,
        log.subject,
        log.body,
        log.cc,
        log.bcc,
        log.status,
        log.message
      ]);
    });
  } catch (e) {
    logError(`Gagal menulis log: ${e}`);
  }
}

// --- Penanganan Error ---
function logError(error) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName('Log Error');
    if (!errorSheet) {
      errorSheet = ss.insertSheet('Log Error');
      errorSheet.appendRow(['Waktu', 'Pesan Error']);
    }
    errorSheet.appendRow([new Date(), error]);
  } catch (e) {
    console.error(`Gagal menulis log error: ${e}`, error);
  }
}


// --- Fungsi  Penjadwalan ---
function createTimeDrivenTrigger(triggerTime, config) {
  try {
    const date = new Date(triggerTime);
    const now = new Date();
    if (date <= now) {
      return { status: 'error', message: 'Waktu penjadwalan harus di masa depan.' };
    }

    ScriptApp.newTrigger('scheduledSendEmails')
      .timeBased()
      .at(date)
      .create();

    PropertiesService.getUserProperties().setProperty('emailConfig', JSON.stringify(config));

    return { status: 'success', message: `Pengiriman email dijadwalkan pada ${date}.` };
  } catch (e) {
    logError(`Gagal membuat penjadwalan: ${e}`);
    return { status: 'error', message: `Gagal membuat penjadwalan: ${e}` };
  }
}

function scheduledSendEmails() {
  const configJson = PropertiesService.getUserProperties().getProperty('emailConfig');
  if (configJson) {
    const config = JSON.parse(configJson);
    sendEmails(config);
  } else {
    logError(`Tidak ada konfigurasi yang tersimpan untuk penjadwalan`);
  }

  // Delete trigger setelah dijalankan
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'scheduledSendEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'scheduledSendEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function testEmail(config) {
  try {

    const result = MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "Test Email Pengaturan",
      htmlBody: config.emailBody,
      cc: config.defaultCC,
      bcc: config.defaultBCC,
      replyTo: config.replyTo,
      from: config.from
    });

    return { status: 'success', message: 'Test email berhasil dikirimkan.' };
  } catch (e) {
    logError(`Gagal mengirimkan test email: ${e}`);
    return { status: 'error', message: `Gagal mengirimkan test email: ${e}` };
  }
}

//--- Save Config ---
function saveConfig(config) {
  try {
    PropertiesService.getUserProperties().setProperty('emailConfig', JSON.stringify(config));
    return { status: 'success', message: 'Pengaturan berhasil disimpan.' };
  } catch (e) {
    logError(`Gagal menyimpan pengaturan: ${e}`);
    return { status: 'error', message: `Gagal menyimpan pengaturan: ${e}` };
  }
}


function loadConfig() {
  try {
    const configJson = PropertiesService.getUserProperties().getProperty('emailConfig');
    if (configJson) {
      return JSON.parse(configJson);
    } else {
      return null;
    }
  } catch (e) {
    logError(`Gagal memuat konfigurasi: ${e}`);
    return null;
  }
}
