# 🚀 Email Marketing Blast dengan Google Apps Script  

Ekstensi ini memungkinkan kamu mengirim **email marketing secara massal langsung dari Google Spreadsheet** menggunakan Google Apps Script.  
Dilengkapi dengan berbagai fitur untuk mempermudah pengelolaan email campaign.  

## ✨ Fitur  
✅ **Mengirim email ke banyak penerima secara otomatis**  
✅ **Dukungan format HTML untuk email profesional**  
✅ **Jadwal pengiriman email otomatis**  
✅ **Deteksi email berhasil terkirim atau gagal**  
✅ **Tanpa perlu menggunakan layanan pihak ketiga**  

---

## 📥 Instalasi  

1. **Buka Google Spreadsheet**  
2. **Buka menu** `Extensions` → `Apps Script`  
3. **Hapus isi kode default** di editor Apps Script  
4. **Salin & Tempel kode di bawah ini** ke editor Apps Script  
5. **Simpan dan jalankan script**  

---

## 📜 Kode 1: `Code.gs` (Google Apps Script)  

```javascript
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
```
## Cara Menggunakan:
Buka Apps Script
Tambahkan kode di atas ke dalam editor Apps Script
Jalankan fungsi doGet() untuk memulai
Kirim email ke penerima yang ada di Spreadsheet

## 📜 Kode 2: `index.html` (Google Apps Script)  

``` HTML
<!DOCTYPE html>
<html>

<head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600&display=swap" rel="stylesheet">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600&display=swap');

        :root {
            --primary-color: #000000;
            --secondary-color: #333333;
            --background-light: #F4F4F4;
            --text-color: #000000;
            --border-color: #000000;
            --soft-shadow: 4px 4px 0 rgba(0, 0, 0, 0.1);
            --spacing-sm: 10px;
            --spacing-md: 15px;
            --spacing-lg: 20px;
            --border-radius: 8px;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Poppins', sans-serif;
            background-color: var(--background-light);
            color: var(--text-color);
            line-height: 1.6;
            padding: var(--spacing-lg);
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: var(--spacing-lg);
            border: 3px solid var(--border-color);
            border-radius: var(--border-radius);
            box-shadow: var(--soft-shadow);
        }

        /* Typography */
        h1,
        h2,
        h3,
        h4,
        h5,
        h6 {
            color: var(--primary-color);
            margin-bottom: var(--spacing-md);
        }

        /* Form Styling */
        .form-section {
            background-color: white;
            border: 3px solid var(--border-color);
            border-radius: var(--border-radius);
            padding: var(--spacing-md);
            margin-bottom: var(--spacing-lg);
            box-shadow: var(--soft-shadow);
        }

        label {
            display: block;
            margin-bottom: var(--spacing-sm);
            font-weight: 600;
            color: var(--text-color);
        }

        input[type="text"],
        input[type="email"],
        input[type="password"],
        input[type="datetime-local"],
        select,
        textarea {
            width: 100%;
            padding: var(--spacing-sm);
            margin-bottom: var(--spacing-md);
            border: 3px solid var(--border-color);
            border-radius: var(--border-radius);
            background-color: white;
            font-family: 'Poppins', sans-serif;
            color: var(--text-color);
            transition: all 0.3s ease;
            box-shadow: var(--soft-shadow);
        }

        input[type="text"]:focus,
        input[type="email"]:focus,
        input[type="password"]:focus,
        input[type="datetime-local"]:focus,
        select:focus,
        textarea:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 6px 6px 0 rgba(0, 0, 0, 0.2);
        }

        textarea {
            min-height: 150px;
            resize: vertical;
        }

        /* Button Styling */
        button,
        .btn {
            font-family: 'Poppins', sans-serif;
            background-color: var(--primary-color);
            color: white;
            padding: var(--spacing-sm) var(--spacing-md);
            border: 3px solid var(--border-color);
            border-radius: var(--border-radius);
            cursor: pointer;
            transition: all 0.3s ease;
            font-weight: 600;
            box-shadow: var(--soft-shadow);
            display: inline-block;
            text-align: center;
            text-decoration: none;
        }

        button:hover,
        .btn:hover {
            transform: translate(2px, 2px);
            box-shadow: 2px 2px 0 rgba(0, 0, 0, 0.1);
        }

        button.secondary,
        .btn.secondary {
            background-color: #666666;
        }

        button.danger,
        .btn.danger {
            background-color: #CC0000;
        }

        /* Flex Layout */
        .flex {
            display: flex;
            flex-wrap: wrap;
            gap: var(--spacing-md);
        }

        .flex>* {
            flex: 1 1 250px;
        }

        /* Modal Styling */
        .modal {
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.5);
            display: none;
            justify-content: center;
            align-items: center;
        }

        .modal-content {
            background-color: white;
            padding: var(--spacing-lg);
            border-radius: var(--border-radius);
            border: 3px solid var(--border-color);
            width: 90%;
            max-width: 600px;
            box-shadow: var(--soft-shadow);
        }

        /* List Styling */
        .template-list {
            list-style: none;
            padding: 0;
            max-height: 250px;
            overflow-y: auto;
            border: 3px solid var(--border-color);
            border-radius: var(--border-radius);
        }

        .template-list li {
            padding: var(--spacing-sm);
            border-bottom: 3px solid var(--border-color);
            cursor: pointer;
            transition: background-color 0.3s ease;
            color: var(--text-color);
        }

        .template-list li:hover {
            background-color: #f0f0f0;
        }

        /* Status Messages */
        .status-message {
            padding: var(--spacing-md);
            border-radius: var(--border-radius);
            margin-top: var(--spacing-md);
            font-weight: 600;
            border: 3px solid var(--border-color);
            box-shadow: var(--soft-shadow);
        }

        .status-message.success {
            background-color: #E6F3E6;
            color: #006400;
        }

        .status-message.error {
            background-color: #F3E6E6;
            color: #8B0000;
        }

        /* Loading Overlay */
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(255, 255, 255, 0.8);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 1500;
        }

        .loading-spinner {
            border: 8px solid #f3f3f3;
            border-top: 8px solid var(--primary-color);
            border-radius: 50%;
            width: 60px;
            height: 60px;
            animation: spin 1.5s linear infinite;
        }

        @keyframes spin {
            0% {
                transform: rotate(0deg);
            }

            100% {
                transform: rotate(360deg);
            }
        }

        /* Email Editor Specific Styles */
        .email-editor-container {
            display: flex;
            flex-direction: column;
            gap: var(--spacing-sm);
        }

        .email-editor-toolbar {
            display: flex;
            gap: var(--spacing-sm);
            margin-bottom: var(--spacing-sm);
        }

        .email-editor-toolbar select,
        .email-editor-toolbar button {
            padding: var(--spacing-sm);
            border-radius: var(--border-radius);
            border: 3px solid var(--border-color);
            background-color: white;
            color: var(--text-color);
            cursor: pointer;
            box-shadow: var(--soft-shadow);
        }

        #emailEditor,
        #emailPreview {
            border: 3px solid var(--border-color);
            border-radius: var(--border-radius);
            min-height: 300px;
            background-color: white;
            color: var(--text-color);
            padding: var(--spacing-md);
            box-shadow: var(--soft-shadow);
        }

        #emailPreview {
            overflow-y: auto;
        }

        /* Ensure all text in preview is black */
        #emailPreview *,
        #emailPreview p,
        #emailPreview span,
        #emailPreview div,
        #emailPreview h1,
        #emailPreview h2,
        #emailPreview h3,
        #emailPreview h4,
        #emailPreview h5,
        #emailPreview h6,
        #emailPreview a,
        #emailPreview strong,
        #emailPreview em,
        #emailPreview b,
        #emailPreview i {
            color: var(--text-color) !important;
        }

        /* Responsive Adjustments */
        @media screen and (max-width: 768px) {
            body {
                padding: var(--spacing-md);
            }

            .container {
                padding: var(--spacing-md);
            }

            .flex {
                flex-direction: column;
            }

            .flex>* {
                flex-basis: 100%;
            }

            .modal-content {
                width: 95%;
                padding: var(--spacing-md);
            }

            button,
            .btn {
                width: 100%;
                margin-bottom: var(--spacing-sm);
            }

            input[type="text"],
            input[type="email"],
            input[type="password"],
            input[type="datetime-local"],
            select,
            textarea {
                width: 100%;
            }

            .email-editor-toolbar {
                flex-direction: column;
            }

            .email-editor-toolbar select,
            .email-editor-toolbar button {
                width: 100%;
                margin-bottom: var(--spacing-sm);
            }
        }

        /* Scrollbar Styling */
        #emailPreview::-webkit-scrollbar {
            width: 10px;
        }

        #emailPreview::-webkit-scrollbar-track {
            background: #f1f1f1;
        }

        #emailPreview::-webkit-scrollbar-thumb {
            background: #888;
            border-radius: 5px;
        }

        #emailPreview::-webkit-scrollbar-thumb:hover {
            background: #555;
        }
    </style>
</head>

<body>
    <div class="container">
        <h2>Pengaturan Email Otomatis</h2>

        <div class="form-section">
            <label for="toColumn">Kolom Penerima (To):</label>
            <select id="toColumn"></select>

            <label for="ccColumn">Kolom CC:</label>
            <select id="ccColumn"></select>

            <label for="bccColumn">Kolom BCC:</label>
            <select id="bccColumn"></select>

            <label for="subjectColumn">Kolom Subjek:</label>
            <select id="subjectColumn"></select>
        </div>
        <div class="form-section">
            <label>Mode Input Email:</label>
            <select id="emailInputMode">
                <option value="manual">Input Manual</option>
                <option value="template">Gunakan Template</option>
            </select>

            <div id="manualInputSection" style="display:block;">
                <div class="email-editor-container">
                    <div class="email-editor-toolbar">
                        <select id="fontSelect">
                            <option value="Arial, sans-serif">Arial</option>
                            <option value="Verdana, sans-serif">Verdana</option>
                            <option value="Tahoma, sans-serif">Tahoma</option>
                            <option value="Helvetica, sans-serif">Helvetica</option>
                            <option value="'Times New Roman', serif">Times New Roman</option>
                            <option value="'Courier New', monospace">Courier New</option>
                            <option value="Georgia, serif">Georgia</option>
                            <option value="Garamond, serif">Garamond</option>
                            <option value="'Trebuchet MS', sans-serif">Trebuchet MS</option>
                            <option value="'Segoe UI', sans-serif">Segoe UI</option>
                            <option value="'Roboto', sans-serif">Roboto</option>
                            <option value="'Open Sans', sans-serif">Open Sans</option>
                            <option value="'Lato', sans-serif">Lato</option>
                            <option value="Poppins, sans-serif">Poppins</option>
                        </select>
                        <button type="button" id="toggleEditorModeBtn">Tampilkan HTML</button>

                    </div>
                    <iframe id="emailEditor" style="display:block;"></iframe>

                    <div id="emailPreview" style="display:none"></div>

                </div>

            </div>
            <div id="templateSection" style="display:none;">
                <label>Pilih Template:</label>
                <input type="text" id="templateSearchInput" placeholder="Cari Template">
                <ul id="templateList" class="template-list"></ul>
                <div class="template-action">
                    <button id="openTemplateModal" type="button">Kelola Template</button>
                </div>
            </div>
        </div>
        <div class="form-section">
            <label>Pengaturan Default</label>

            <div class="flex">
                <div> <label for="defaultCC">CC Default:</label>
                    <input type="text" id="defaultCC" placeholder="Email CC Default (pisahkan koma)" />
                </div>
                <div> <label for="defaultBCC">BCC Default:</label>
                    <input type="text" id="defaultBCC" placeholder="Email BCC Default (pisahkan koma)" />
                </div>
                <div> <label for="defaultSubject">Subjek Default:</label>
                    <input type="text" id="defaultSubject" placeholder="Subjek Email Default" />
                </div>
                <div> <label for="fromEmail">Email Pengirim (From):</label>
                    <input type="text" id="fromEmail" placeholder="Email Pengirim (From)" />
                </div>
                <div> <label for="replyToEmail">Email Balasan (Reply-To):</label>
                    <input type="text" id="replyToEmail" placeholder="Email Balasan (Reply-To)" />
                </div>
            </div>
        </div>

        <div class="form-section">
            <label>Pengaturan Output PDF:</label>
            <label><input type="checkbox" id="pdfEnabled" /> Aktifkan Output PDF</label>

            <div id="pdfOptions" style="display:none">
                <label for="pdfFolder">Folder Penyimpanan:</label>
                <div style="display:flex; gap:5px; align-items:center;"> <input type="text" id="pdfFolder" readonly/>
                    <button type="button" id="selectFolderBtn" class="secondary">Pilih Folder</button>
                </div>
                <label for="pdfFilename">Nama File PDF (Merge Tag <<nama_kolom>>):</label>
                <input type="text" id="pdfFilename" placeholder="Nama File PDF (Optional)">

            </div>
        </div>
        <div class="form-section">
            <label>Pengaturan Kondisi Pengiriman:</label>
            <label><input type="checkbox" id="conditionalEnabled" /> Aktifkan Kondisi Pengiriman</label>
            <div id="conditionOptions" style="display:none">
                <label for="conditionColumn">Kolom Kondisi:</label>
                <select id="conditionColumn"></select>

                <label for="conditionOperator">Operator Kondisi:</label>
                <select id="conditionOperator">
                    <option value="sama dengan">Sama Dengan</option>
                    <option value="tidak sama dengan">Tidak Sama Dengan</option>
                    <option value="berisi">Berisi</option>
                    <option value="tidak berisi">Tidak Berisi</option>
                    <option value="lebih besar dari">Lebih Besar Dari</option>
                    <option value="lebih kecil dari">Lebih Kecil Dari</option>
                </select>
                <label for="conditionValue">Nilai Kondisi:</label>
                <input type="text" id="conditionValue" />
            </div>
        </div>
        <div class="form-section">
            <div class="btn-group">
                <button type="button" id="saveSettingsBtn">Simpan Pengaturan</button>
                <button type="button" id="sendTestEmailBtn" class="secondary">Test Email</button>
                <button type="button" id="sendEmailsBtn">Kirim Email</button>
            </div>
            <label>Jadwalkan Email</label>
            <div style="display:flex; gap:5px; align-items:center;">
                <input type="datetime-local" id="scheduleTime">
                <button type="button" id="scheduleEmailsBtn" class="secondary">Jadwalkan</button>
                <button type="button" id="deleteSchedule" class="danger">Hapus Jadwal</button>
            </div>
        </div>
        <div class="status-message" id="statusMessage"></div>
    </div>

    <!-- Modal untuk Template -->
    <div id="templateModal" class="modal">
        <div class="modal-content">
            <span class="close-button" id="closeTemplateModal">×</span>
            <h2>Kelola Template Email</h2>
            <label for="templateName">Nama Template:</label>
            <input type="text" id="templateName" />
            <label for="templateBody">Isi Template:</label>
            <textarea id="templateBody"></textarea>

            <div class="btn-group">
                <button type="button" id="saveTemplateBtn">Simpan Template</button>
            </div>
            <hr>
            <input type="text" id="templateSearch" placeholder="Cari Template">
            <ul id="allTemplateList" class="template-list"></ul>
        </div>
    </div>

    <div class="loading-overlay" id="loadingOverlay">
        <div class="loading-spinner"></div>
    </div>

    <script>
         const toColumnSelect = document.getElementById('toColumn');
        const ccColumnSelect = document.getElementById('ccColumn');
        const bccColumnSelect = document.getElementById('bccColumn');
        const subjectColumnSelect = document.getElementById('subjectColumn');
        const emailInputModeSelect = document.getElementById('emailInputMode');
        const manualInputSection = document.getElementById('manualInputSection');
        const templateSection = document.getElementById('templateSection');
        const emailEditor = document.getElementById('emailEditor');
        const emailPreview = document.getElementById('emailPreview');
        const templateList = document.getElementById('templateList');
        const openTemplateModalBtn = document.getElementById('openTemplateModal');
        const pdfEnabledCheckbox = document.getElementById('pdfEnabled');
        const pdfOptionsDiv = document.getElementById('pdfOptions');
        const pdfFolderInput = document.getElementById('pdfFolder');
        const selectFolderBtn = document.getElementById('selectFolderBtn');
        const pdfFilenameInput = document.getElementById('pdfFilename');
        const sendEmailsBtn = document.getElementById('sendEmailsBtn');
        const scheduleEmailsBtn = document.getElementById('scheduleEmailsBtn');
        const scheduleTimeInput = document.getElementById('scheduleTime');
        const deleteScheduleBtn = document.getElementById('deleteSchedule');
        const conditionalEnabledCheckbox = document.getElementById('conditionalEnabled');
        const conditionOptionsDiv = document.getElementById('conditionOptions');
        const conditionColumnSelect = document.getElementById('conditionColumn');
        const conditionOperatorSelect = document.getElementById('conditionOperator');
        const conditionValueInput = document.getElementById('conditionValue');
        const statusMessage = document.getElementById('statusMessage');
        const templateModal = document.getElementById('templateModal');
        const closeTemplateModalBtn = document.getElementById('closeTemplateModal');
        const templateNameInput = document.getElementById('templateName');
        const templateBodyTextarea = document.getElementById('templateBody');
        const saveTemplateBtn = document.getElementById('saveTemplateBtn');
        const allTemplateList = document.getElementById('allTemplateList');
        const templateSearchInput = document.getElementById('templateSearch');
        const templateSearch = document.getElementById('templateSearchInput');
        const defaultCC = document.getElementById('defaultCC');
        const defaultBCC = document.getElementById('defaultBCC');
        const defaultSubject = document.getElementById('defaultSubject');
        const fromEmail = document.getElementById('fromEmail');
        const replyToEmail = document.getElementById('replyToEmail');
        const saveSettingsBtn = document.getElementById('saveSettingsBtn');
        const sendTestEmailBtn = document.getElementById('sendTestEmailBtn');
        const loadingOverlay = document.getElementById('loadingOverlay');
        const fontSelect = document.getElementById('fontSelect');
        const toggleEditorModeBtn = document.getElementById('toggleEditorModeBtn');

        let allHeaders = [];
        let currentTemplate = null;
        let isHtmlMode = false;
        let editorDocument;

        function showLoading() {
            loadingOverlay.style.display = 'flex';
        }

        function hideLoading() {
            loadingOverlay.style.display = 'none';
        }

        function showStatusMessage(message, type) {
            statusMessage.textContent = message;
            statusMessage.classList.remove('success', 'error', 'info');
            statusMessage.classList.add(type);
            statusMessage.style.display = 'block';
        }


        function clearStatusMessage() {
            statusMessage.style.display = 'none';
            statusMessage.textContent = '';
        }


        function clearTemplateFields() {
            templateNameInput.value = '';
            templateBodyTextarea.value = '';
        }

        // Function to set the font of the email editor
        function setEditorFont(fontFamily) {
            if (editorDocument) {
                editorDocument.body.style.fontFamily = fontFamily;
            }
        }

        // Set default font
        function setDefaultFont(defaultFont) {
            if (fontSelect.value === "Poppins, sans-serif") {
                setEditorFont(defaultFont);
            } else {
                setEditorFont(fontSelect.value)
            }
        }


        //Initialize the iframe editor
        function initEditor() {
            editorDocument = emailEditor.contentDocument || emailEditor.contentWindow.document;
            editorDocument.designMode = "on";
            setDefaultFont("'Poppins', sans-serif") // Set Poppins as default font
            editorDocument.body.addEventListener('paste', handlePaste);
        }

        // Populate Select Options
        function populateColumns(selector, includeNone = false) {
            selector.innerHTML = '';
            if (includeNone) {
                const noneOption = document.createElement('option');
                noneOption.value = "Tidak Ada";
                noneOption.text = "Tidak Ada";
                selector.appendChild(noneOption);
            }
            allHeaders.forEach(header => {
                const option = document.createElement('option');
                option.value = header;
                option.text = header;
                selector.appendChild(option);
            });
        }

        //Load Config
        function loadConfig() {
            showLoading();
            google.script.run
                .withSuccessHandler(config => {
                    hideLoading();
                    if (config) {
                        toColumnSelect.value = config.toColumn;
                        ccColumnSelect.value = config.ccColumn || 'Tidak Ada';
                        bccColumnSelect.value = config.bccColumn || 'Tidak Ada';
                        subjectColumnSelect.value = config.subjectColumn || 'Tidak Ada';
                        emailInputModeSelect.value = config.emailInputMode;
                        if (config.emailInputMode === 'manual') {
                            manualInputSection.style.display = 'block';
                            templateSection.style.display = 'none';
                        } else if (config.emailInputMode === 'template') {
                            manualInputSection.style.display = 'none';
                            templateSection.style.display = 'block';
                        }
                        if (config.emailBody) {
                            if (isHtmlMode) {
                                emailEditor.style.display = 'none';
                                emailPreview.style.display = 'block';
                                emailPreview.innerHTML = config.emailBody;
                                // Inject CSS from config email body
                                injectCssToPreview(config.emailBody);
                            } else {
                                emailEditor.style.display = 'block';
                                emailPreview.style.display = 'none';
                                if (editorDocument) {
                                    editorDocument.body.innerHTML = config.emailBody;
                                }
                            }
                        }
                        pdfEnabledCheckbox.checked = config.pdfEnabled;
                        if (config.pdfEnabled) {
                            pdfOptionsDiv.style.display = 'block';
                        } else {
                            pdfOptionsDiv.style.display = 'none';
                        }
                        pdfFolderInput.value = config.pdfFolder || '';
                        pdfFilenameInput.value = config.pdfFilename || '';
                        conditionalEnabledCheckbox.checked = config.conditionalEnabled;
                        if (config.conditionalEnabled) {
                            conditionOptionsDiv.style.display = 'block';
                        } else {
                            conditionOptionsDiv.style.display = 'none';
                        }
                        conditionColumnSelect.value = config.conditionColumn || '';
                        conditionOperatorSelect.value = config.conditionOperator || 'sama dengan';
                        conditionValueInput.value = config.conditionValue || '';
                        defaultCC.value = config.defaultCC || '';
                        defaultBCC.value = config.defaultBCC || '';
                        defaultSubject.value = config.defaultSubject || '';
                        fromEmail.value = config.from || '';
                        replyToEmail.value = config.replyTo || '';
                        if (config.fontFamily) {
                            fontSelect.value = config.fontFamily
                            setEditorFont(config.fontFamily)
                        }

                        loadTemplates();
                    } else {
                        loadHeaders();
                    }
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage('Gagal memuat pengaturan: ' + error, 'error');
                })
                .loadConfig();
        }

        // Load headers
        function loadHeaders() {
            showLoading();
            google.script.run
                .withSuccessHandler(headers => {
                    hideLoading();
                    allHeaders = headers;
                    populateColumns(toColumnSelect);
                    populateColumns(ccColumnSelect, true);
                    populateColumns(bccColumnSelect, true);
                    populateColumns(subjectColumnSelect, true);
                    populateColumns(conditionColumnSelect);
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage('Gagal memuat header kolom: ' + error, 'error');
                })
                .getSheetHeaders();
        }



        // Load Templates to List
        function loadTemplates() {
            google.script.run
                .withSuccessHandler(templates => {
                    if (templates) {
                        const filteredTemplates = filterTemplates(templates);
                        renderTemplateList(filteredTemplates);
                    } else {
                        renderTemplateList({});
                    }
                })
                .withFailureHandler(error => {
                    showStatusMessage("Gagal memuat template: " + error, 'error');
                    renderTemplateList({});
                })
                .loadTemplates();
        }

        // Template filter
        function filterTemplates(templates) {
            const filter = templateSearch.value.toLowerCase();
            const filteredTemplates = {};

            for (const key in templates) {
                if (key.toLowerCase().includes(filter)) {
                    filteredTemplates[key] = templates[key];
                }
            }
            return filteredTemplates;
        }


        function renderTemplateList(templates) {
            templateList.innerHTML = '';
            for (const key in templates) {
                const listItem = document.createElement('li');
                listItem.textContent = key;
                listItem.addEventListener('click', () => {
                    if (editorDocument) {
                        editorDocument.body.innerHTML = templates[key];
                    }
                    currentTemplate = key;
                });
                templateList.appendChild(listItem);
            }

            if (Object.keys(templates).length === 0) {
                const listItem = document.createElement('li');
                listItem.textContent = "Tidak ada template";
                templateList.appendChild(listItem);
            }
        }

        function renderAllTemplateList(templates) {
            allTemplateList.innerHTML = '';
            for (const key in templates) {
                const listItem = document.createElement('li');
                listItem.textContent = key;
                listItem.addEventListener('click', () => {
                    templateBodyTextarea.value = templates[key];
                    templateNameInput.value = key;
                });
                const actionDiv = document.createElement('div');
                actionDiv.classList.add('template-action');
                const deleteButton = document.createElement('button');
                deleteButton.textContent = 'Hapus';
                deleteButton.classList.add('danger');
                deleteButton.addEventListener('click', (event) => {
                    event.stopPropagation();
                    deleteTemplate(key);
                });
                actionDiv.appendChild(deleteButton);
                listItem.appendChild(actionDiv);
                allTemplateList.appendChild(listItem);
            }

            if (Object.keys(templates).length === 0) {
                const listItem = document.createElement('li');
                listItem.textContent = "Tidak ada template";
                allTemplateList.appendChild(listItem);
            }
        }


        //  Save Template
        function saveTemplate() {
            const templateName = templateNameInput.value;
            const templateBody = editorDocument.body.innerHTML;

            if (!templateName) {
                showStatusMessage("Nama template harus diisi", 'error');
                return;
            }
            if (!templateBody) {
                showStatusMessage("Isi template harus diisi", 'error');
                return;
            }

            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'success') {
                        showStatusMessage(response.message, 'success');
                        clearTemplateFields();
                        loadTemplates();
                        loadAllTemplates();
                    } else {
                        showStatusMessage(response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    showStatusMessage("Gagal menyimpan template: " + error, 'error');
                })
                .saveTemplate(templateName, templateBody);
        }


        //  Delete Template
        function deleteTemplate(templateName) {
            showLoading();
            google.script.run
                .withSuccessHandler(response => {
                    hideLoading();
                    if (response.status === 'success') {
                        showStatusMessage(response.message, 'success');
                        loadTemplates();
                        loadAllTemplates();
                    } else {
                        showStatusMessage(response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage("Gagal menghapus template: " + error, 'error');
                })
                .deleteTemplate(templateName);
        }

        //load All Templates
        function loadAllTemplates() {
            google.script.run
                .withSuccessHandler(templates => {
                    if (templates) {
                        renderAllTemplateList(templates);
                    } else {
                        renderAllTemplateList({});
                    }
                })
                .withFailureHandler(error => {
                    showStatusMessage("Gagal memuat template: " + error, 'error');
                    renderAllTemplateList({});
                })
                .loadTemplates();
        }


        // Select folder
        function selectFolder() {
            google.script.run
                .withSuccessHandler(folderId => {
                    if (folderId) {
                        pdfFolderInput.value = folderId;
                        google.script.run.withSuccessHandler(folderName => {
                            if (folderName) {
                                pdfFolderInput.value = folderName
                            }
                        }).getFolderName(folderId)
                    } else {
                        pdfFolderInput.value = '';
                    }
                })
                .withFailureHandler(error => {
                    showStatusMessage("Gagal memilih folder: " + error, 'error');
                })
                .selectFolder();
        }

        // Save settings to storage
        function saveConfig() {
            const config = {
                toColumn: toColumnSelect.value,
                ccColumn: ccColumnSelect.value,
                bccColumn: bccColumnSelect.value,
                subjectColumn: subjectColumnSelect.value,
                emailInputMode: emailInputModeSelect.value,
                emailBody: isHtmlMode ? emailPreview.innerHTML : editorDocument.body.innerHTML,
                pdfEnabled: pdfEnabledCheckbox.checked,
                pdfFolder: pdfFolderInput.value,
                pdfFilename: pdfFilenameInput.value,
                conditionalEnabled: conditionalEnabledCheckbox.checked,
                conditionColumn: conditionColumnSelect.value,
                conditionOperator: conditionOperatorSelect.value,
                conditionValue: conditionValueInput.value,
                defaultCC: defaultCC.value,
                defaultBCC: defaultBCC.value,
                defaultSubject: defaultSubject.value,
                from: fromEmail.value,
                replyTo: replyToEmail.value,
                fontFamily: fontSelect.value
            };

            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'success') {
                        showStatusMessage(response.message, 'success');
                    } else {
                        showStatusMessage(response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    showStatusMessage('Gagal menyimpan pengaturan: ' + error, 'error');
                })
                .saveConfig(config);
        }

        // Handle Paste in Iframe
        function handlePaste(e) {
            e.preventDefault();
            let text = e.clipboardData.getData('text/html');
            if (!text) {
                text = e.clipboardData.getData('text/plain');
            }

            if (text) {
                if (isHtmlMode) {
                    // Handle HTML paste for Preview
                    emailPreview.innerHTML = text;
                    injectCssToPreview(text);
                } else {
                    // Handle HTML paste for Editor
                    editorDocument.execCommand('insertHTML', false, text);
                }
            }
        }

        function injectCssToPreview(html) {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = html;
            let styles = '';
            let promises = [];


            // Extract style tags
            const styleTags = tempDiv.querySelectorAll('style');
            styleTags.forEach(tag => {
                styles += tag.textContent;
                tag.remove()
            });


            // Extract linked CSS files
            const linkTags = tempDiv.querySelectorAll('link[rel="stylesheet"]');
            linkTags.forEach(linkTag => {
                promises.push(
                 fetch(linkTag.href)
                        .then(response => {
                             if (!response.ok) {
                                throw new Error(`Failed to fetch CSS from ${linkTag.href}: ${response.status} ${response.statusText}`);
                             }
                             return response.text();
                          })
                        .then(css => styles += css)
                       .catch(error => {
                          console.error('Failed to fetch CSS:', error);
                           return ''; // Resolve with an empty string for the promise so Promise.all doesn't reject
                     })
                 );
                 linkTag.remove()
             });


            // Wait for all CSS to be fetched, then inject it
            Promise.all(promises)
                    .then(() => {
                        const previewDoc = emailPreview.contentDocument || emailPreview.contentWindow.document;
                            if (previewDoc) {
                                let styleElement = previewDoc.getElementById('injected-styles');
                                    if (!styleElement) {
                                        styleElement = previewDoc.createElement('style');
                                        styleElement.id = 'injected-styles';
                                        previewDoc.head.appendChild(styleElement);
                                    }
                                styleElement.textContent = styles;

                                 // Append HTML content without style and link tags
                                 previewDoc.body.innerHTML = tempDiv.innerHTML
                            }

                    })
                    .catch(error => console.error("Error injecting CSS", error));
        }

        // Function untuk mengirim Email
        function sendEmails(scheduled = false) {
            clearStatusMessage();
            const config = {
                toColumn: toColumnSelect.value,
                ccColumn: ccColumnSelect.value,
                bccColumn: bccColumnSelect.value,
                subjectColumn: subjectColumnSelect.value,
                emailInputMode: emailInputModeSelect.value,
                emailBody: isHtmlMode ? emailPreview.innerHTML : editorDocument.body.innerHTML,
                pdfEnabled: pdfEnabledCheckbox.checked,
                pdfFolder: pdfFolderInput.value,
                pdfFilename: pdfFilenameInput.value,
                conditionalEnabled: conditionalEnabledCheckbox.checked,
                conditionColumn: conditionColumnSelect.value,
                conditionOperator: conditionOperatorSelect.value,
                conditionValue: conditionValueInput.value,
                defaultCC: defaultCC.value,
                defaultBCC: defaultBCC.value,
                defaultSubject: defaultSubject.value,
                from: fromEmail.value,
                replyTo: replyToEmail.value
            };

            if (!config.toColumn) {
                showStatusMessage('Penerima (To) wajib dipilih', 'error');
                return;
            }


            if (config.emailInputMode === 'template' && !editorDocument.body.innerHTML) {
                showStatusMessage('Isi template wajib diisi.', 'error')
                return;
            }


            if (config.emailInputMode === 'manual' && !config.emailBody) {
                showStatusMessage('Isi email wajib diisi.', 'error');
                return;
            }


            showLoading();

            google.script.run
                .withSuccessHandler(response => {
                    hideLoading();
                    if (response.status === 'success') {
                        showStatusMessage(response.message, 'success');
                    } else {
                        showStatusMessage(response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage("Gagal mengirim email: " + error, 'error');
                })
                .sendEmails(config);
        }



        // Function untuk mengirim email test
        function testEmail() {
            clearStatusMessage();
            const config = {
                toColumn: toColumnSelect.value,
                ccColumn: ccColumnSelect.value,
                bccColumn: bccColumnSelect.value,
                subjectColumn: subjectColumnSelect.value,
                emailInputMode: emailInputModeSelect.value,
                emailBody: isHtmlMode ? emailPreview.innerHTML : editorDocument.body.innerHTML,
                defaultCC: defaultCC.value,
                defaultBCC: defaultBCC.value,
                defaultSubject: defaultSubject.value,
                from: fromEmail.value,
                replyTo: replyToEmail.value
            };

            if (config.emailInputMode === 'template' && !editorDocument.body.innerHTML) {
                showStatusMessage('Isi template wajib diisi.', 'error')
                return;
            }

            if (config.emailInputMode === 'manual' && !config.emailBody) {
                showStatusMessage('Isi email wajib diisi.', 'error');
                return;
            }
            showLoading();
            google.script.run
                .withSuccessHandler(response => {
                    hideLoading();
                    if (response.status === 'success') {
                        showStatusMessage(response.message, 'success');
                    } else {
                        showStatusMessage(response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage("Gagal mengirim test email: " + error, 'error');
                })
                .testEmail(config);
        }


        // Penjadwalan email
        function scheduleEmails() {
            const scheduleTime = scheduleTimeInput.value;
            if (!scheduleTime) {
                showStatusMessage('Waktu penjadwalan wajib diisi', 'error');
                return;
            }

            const config = {
                toColumn: toColumnSelect.value,
                ccColumn: ccColumnSelect.value,
                bccColumn: bccColumnSelect.value,
                subjectColumn: subjectColumnSelect.value,
                emailInputMode: emailInputModeSelect.value,
                emailBody: isHtmlMode ? emailPreview.innerHTML : editorDocument.body.innerHTML,
                pdfEnabled: pdfEnabledCheckbox.checked,
                pdfFolder: pdfFolderInput.value,
                pdfFilename: pdfFilenameInput.value,
                conditionalEnabled: conditionalEnabledCheckbox.checked,
                conditionColumn: conditionColumnSelect.value,
                conditionOperator: conditionOperatorSelect.value,
                conditionValue: conditionValueInput.value,
                defaultCC: defaultCC.value,
                defaultBCC: defaultBCC.value,
                defaultSubject: defaultSubject.value,
                from: fromEmail.value,
                replyTo: replyToEmail.value
            };

            if (!config.toColumn) {
                showStatusMessage('Penerima (To) wajib dipilih', 'error');
                return;
            }


            if (config.emailInputMode === 'template' && !editorDocument.body.innerHTML) {
                showStatusMessage('Isi template wajib diisi.', 'error')
                return;
            }


            if (config.emailInputMode === 'manual' && !config.emailBody) {
                showStatusMessage('Isi email wajib diisi.', 'error');
                return;
            }

            showLoading();
            google.script.run
                .withSuccessHandler(response => {
                    hideLoading();
                    if (response.status === 'success') {
                        showStatusMessage(response.message, 'success');
                    } else {
                        showStatusMessage(response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage("Gagal menjadwalkan email: " + error, 'error');
                })
                .createTimeDrivenTrigger(scheduleTime, config);
        }

        // delete triggers
        function deleteTriggers() {
            showLoading();
            google.script.run
                .withSuccessHandler(() => {
                    hideLoading();
                    showStatusMessage("Jadwal email telah dibatalkan.", 'info');
                })
                .withFailureHandler(error => {
                    hideLoading();
                    showStatusMessage("Gagal membatalkan jadwal email: " + error, 'error');
                })
                .deleteTriggers();
        }


        // --- Event Listener ---
        emailInputModeSelect.addEventListener('change', () => {
            if (emailInputModeSelect.value === 'manual') {
                manualInputSection.style.display = 'block';
                templateSection.style.display = 'none';
                if (!isHtmlMode && editorDocument) {
                    editorDocument.body.innerHTML = ''
                    setDefaultFont("'Poppins', sans-serif")
                }
            } else if (emailInputModeSelect.value === 'template') {
                manualInputSection.style.display = 'none';
                templateSection.style.display = 'block';
            }
        });

        pdfEnabledCheckbox.addEventListener('change', () => {
            pdfOptionsDiv.style.display = pdfEnabledCheckbox.checked ? 'block' : 'none';
        });


        conditionalEnabledCheckbox.addEventListener('change', () => {
            conditionOptionsDiv.style.display = conditionalEnabledCheckbox.checked ? 'block' : 'none';
        });


        openTemplateModalBtn.addEventListener('click', () => {
            templateModal.style.display = "flex";
            loadAllTemplates();
        });


        closeTemplateModalBtn.addEventListener('click', () => {
            templateModal.style.display = "none";
            clearTemplateFields();
        });


        selectFolderBtn.addEventListener('click', selectFolder);
        sendEmailsBtn.addEventListener('click', sendEmails);
        scheduleEmailsBtn.addEventListener('click', scheduleEmails);
        deleteScheduleBtn.addEventListener('click', deleteTriggers);
        saveSettingsBtn.addEventListener('click', saveConfig);
        sendTestEmailBtn.addEventListener('click', testEmail);
        saveTemplateBtn.addEventListener('click', saveTemplate);
        templateSearch.addEventListener('input', () => {
            loadAllTemplates();
        });
        templateSearchInput.addEventListener('input', () => {
            loadTemplates();
        });

        fontSelect.addEventListener('change', () => {
            setEditorFont(fontSelect.value);
        });

        toggleEditorModeBtn.addEventListener('click', () => {
            isHtmlMode = !isHtmlMode;

            if (isHtmlMode) {
                emailEditor.style.display = 'none';
                emailPreview.style.display = 'block';

                if (editorDocument) {
                emailPreview.innerHTML = editorDocument.body.innerHTML;
                    // inject style
                    injectCssToPreview(editorDocument.body.innerHTML)
                }
                toggleEditorModeBtn.textContent = 'Tampilkan Editor';
            } else {
                emailEditor.style.display = 'block';
                emailPreview.style.display = 'none';
                toggleEditorModeBtn.textContent = 'Tampilkan HTML';
            }
        });

        // Load configuration on start
        loadConfig();
        initEditor();
    </script>
</body>

</html>
```
## Cara Menggunakan:
1. Buka Google Apps Script.
2. Tambahkan file `index.html` melalui menu **File** → **New** → **HTML**.
3. Salin kode HTML di atas ke dalam file `index.html`.
4. Jalankan fungsi `doGet()` untuk menampilkan form.

## Pengaturan Tambahan:
- Gunakan format HTML untuk email profesional.
- Tambahkan fitur kustomisasi subjek atau isi email.
- Gunakan pemrosesan batch jika jumlah email besar, untuk menghindari limit dari Google.

📌 Selamat mencoba dan semoga sukses dengan email marketing kamu!

