function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Siva Physiocare')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. PROCESS REGISTRATION
function processForm(formObject) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('patients');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { 
    if (data[i][0] === formObject.pid) throw new Error("Patient ID already exists!");
  }
  sheet.appendRow([formObject.pid, formObject.name, formObject.age, formObject.gender, formObject.phone, formObject.address, new Date()]);
  return "Success";
}

// 2. VERIFY PATIENT
function checkPatientExists(pid) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('patients');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(pid).trim()) return { exists: true, name: data[i][1] };
  }
  return { exists: false };
}

// 3. PROCESS VISIT
function processVisitForm(formObject) {
  const visitSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('visits');
  visitSheet.appendRow([formObject.visit_id, formObject.pid, formObject.visit_date, formObject.diagnosis, formObject.reason, formObject.follow_up_date,formObject.category, new Date()]);
  return "Visit Logged";
}

// 4. FETCH VISITS FOR PAYMENT
function getPatientAndVisits(pid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let pData = ss.getSheetByName('patients').getDataRange().getValues();
  let patientName = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][0]).trim() === String(pid).trim()) { patientName = pData[i][1]; break; }
  }
  if (!patientName) return { exists: false };
  
  let patientVisits = [];
  const vSheet = ss.getSheetByName('visits');
  if (vSheet) {
    const vData = vSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim() === String(pid).trim()) {
        patientVisits.push({ visit_id: vData[i][0], date: vData[i][2], diagnosis: vData[i][3] });
      }
    }
  }
  return { exists: true, name: patientName, visits: patientVisits };
}

// 5. PROCESS PAYMENT (UPDATED WITH DATES & NOTES)
function processPaymentForm(formObject) {
  const paySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('payments');
  paySheet.appendRow([
    formObject.payment_id, formObject.pay_visit_id, formObject.pid, formObject.payment_date,
    formObject.period_from, formObject.period_to, // NEW DATES
    formObject.total_amount, formObject.paid_amount, formObject.balance, 
    formObject.notes, // NEW NOTES
    new Date()
  ]);
  return "Payment Saved Successfully";
}

// 6. DASHBOARD HISTORY
function getPatientComprehensiveHistory(pid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pData = ss.getSheetByName('patients').getDataRange().getDisplayValues();
  let patientInfo = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][0]).trim() === String(pid).trim()) {
      patientInfo = { pid: pData[i][0], name: pData[i][1], age: pData[i][2], gender: pData[i][3], phone: pData[i][4], address: pData[i][5], registered: pData[i][6] }; break;
    }
  }
  if (!patientInfo) return { exists: false };
  
  let visits = [];
  const vSheet = ss.getSheetByName('visits');
  if (vSheet) {
    const vData = vSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim() === String(pid).trim()) {
        visits.push({ visit_id: vData[i][0], date: vData[i][2], diagnosis: vData[i][3], reason: vData[i][4], follow_up: vData[i][5] });
      }
    }
  }
  
  let payments = [], totalBilled = 0, totalPaid = 0;
  const paySheet = ss.getSheetByName('payments');
  if (paySheet) {
    const payData = paySheet.getDataRange().getDisplayValues();
    for (let i = 1; i < payData.length; i++) {
      if (String(payData[i][2]).trim() === String(pid).trim()) {
        const tAmt = parseFloat(payData[i][6].replace(/,/g, '')) || 0; // Index 6 is Total
        const pAmt = parseFloat(payData[i][7].replace(/,/g, '')) || 0; // Index 7 is Paid
        totalBilled += tAmt; totalPaid += pAmt;
        payments.push({ payment_id: payData[i][0], visit_id: payData[i][1], date: payData[i][3], total: tAmt, paid: pAmt, balance: parseFloat(payData[i][8].replace(/,/g, '')) || 0 });
      }
    }
  }
  return { exists: true, patient: patientInfo, visits: visits, payments: payments, financialSummary: { totalBilled: totalBilled, totalPaid: totalPaid, outstandingBalance: totalBilled - totalPaid } };
}


// ==========================================
// 7. PROCESS MEDICAL REPORT (UPDATED WITH REFERRED BY)
// ==========================================
function processReportForm(formObject, fileDataArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName('reports');
  
  let folderName = "Patient_Medical_Files";
  let folders = DriveApp.getFoldersByName(folderName);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  
  let newFileUrls = [];
  if (fileDataArray && fileDataArray.length > 0) {
    fileDataArray.forEach(f => {
      let blob = Utilities.newBlob(Utilities.base64Decode(f.base64), f.mimeType, f.filename);
      let file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      newFileUrls.push(file.getName() + "::" + file.getUrl()); 
    });
  }

  const data = reportSheet.getDataRange().getValues();
  let existingRow = -1;
  let existingFiles = "";
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(formObject.pid).trim()) {
      existingRow = i + 1; 
      existingFiles = data[i][28] || ""; // MOVED TO INDEX 28 DUE TO NEW COLUMN
      break;
    }
  }

  let allFiles = existingFiles;
  if (newFileUrls.length > 0) {
    if (existingFiles && existingFiles.trim() !== "") {
      allFiles = existingFiles + "||" + newFileUrls.join("||");
    } else {
      allFiles = newFileUrls.join("||");
    }
  }

  const isChecked = (val) => val === 'on' ? 'Yes' : 'No';
  const rowData = [
    formObject.pid, isChecked(formObject.bp), isChecked(formObject.dm), isChecked(formObject.epilepsy),
    isChecked(formObject.metal_implant), isChecked(formObject.medications), isChecked(formObject.heart_diseases),
    isChecked(formObject.pregnant), isChecked(formObject.pacemaker), isChecked(formObject.surgeries),
    isChecked(formObject.thyroid), formObject.notes, isChecked(formObject.injury), isChecked(formObject.trauma),
    isChecked(formObject.fracture), isChecked(formObject.cumulative_trauma), isChecked(formObject.others),
    formObject.chief_complaints, formObject.past_history, formObject.biological, formObject.psychological,
    formObject.sociological, formObject.objective_examination, formObject.investigation, formObject.diagnosis,
    formObject.treatment, formObject.home_advices, 
    formObject.referred_by, // <-- NEW REFERRED BY DATA ADDED HERE
    allFiles, new Date() 
  ];

  if (existingRow > -1) {
    reportSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    return "Report Updated Successfully";
  } else {
    reportSheet.appendRow(rowData);
    return "Report Created Successfully";
  }
}

// ==========================================
// 8. FETCH MEDICAL REPORT (UPDATED TO FETCH UPDATES LOG)
// ==========================================
function getMedicalReport(pid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('patients');
  let patientInfo = null;
  
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][0]).trim() === String(pid).trim()) {
        patientInfo = { name: pData[i][1], age: pData[i][2], gender: pData[i][3] }; break;
      }
    }
  }
  if (!patientInfo) return { exists: false, error: "Patient ID not found." };
  
  const rSheet = ss.getSheetByName('reports');
  const rData = rSheet.getDataRange().getDisplayValues();
  let report = null;
  
  for (let i = 1; i < rData.length; i++) {
    if (String(rData[i][0]).trim() === String(pid).trim()) {
      let fileString = rData[i][28];
      
      let filesArray = [];
      if (fileString && typeof fileString === 'string') {
        filesArray = fileString.split("||")
          .filter(f => f.trim() !== "")
          .map(f => { let parts = f.split("::"); return { name: parts[0], url: parts[1] }; })
          .filter(f => f.url && f.url !== "undefined");
      }
      
      report = {
        bp: rData[i][1], dm: rData[i][2], epilepsy: rData[i][3], metal_implant: rData[i][4], medications: rData[i][5],
        heart_diseases: rData[i][6], pregnant: rData[i][7], pacemaker: rData[i][8], surgeries: rData[i][9], thyroid: rData[i][10],
        notes: rData[i][11], injury: rData[i][12], trauma: rData[i][13], fracture: rData[i][14], cumulative_trauma: rData[i][15], others: rData[i][16],
        chief_complaints: rData[i][17], past_history: rData[i][18], biological: rData[i][19], psychological: rData[i][20],
        sociological: rData[i][21], objective_examination: rData[i][22], investigation: rData[i][23], diagnosis: rData[i][24],
        treatment: rData[i][25], home_advices: rData[i][26], referred_by: rData[i][27], files: filesArray, last_updated: rData[i][29]
      };
      break;
    }
  }

  // FETCH PATIENT UPDATES LOG
  const uSheet = ss.getSheetByName('updates');
  let updatesArray = [];
  if (uSheet) {
    const uData = uSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < uData.length; i++) {
      if (String(uData[i][0]).trim() === String(pid).trim()) {
        updatesArray.push({ date: uData[i][1], notes: uData[i][2] });
      }
    }
  }

  return { exists: true, patient: patientInfo, report: report, updates: updatesArray };
}


// ==========================================
// 9. PROCESS PATIENT UPDATES
// ==========================================
function processUpdateForm(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('updates');
  if (!sheet) sheet = ss.insertSheet('updates'); // Auto-create sheet if missing
  
  sheet.appendRow([
    formObject.pid,
    formObject.update_date,
    formObject.update_notes,
    new Date() // Timestamp
  ]);
  return "Update Saved";
}

// ==========================================
// 10. PROCESS PRODUCTS
// ==========================================
function processProductForm(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('products');
  if (!sheet) sheet = ss.insertSheet('products'); // Auto-create sheet if missing
  
  sheet.appendRow([
    formObject.pid,
    formObject.product_date,
    formObject.product_name,
    formObject.quantity,
    formObject.price,
    new Date() // Timestamp
  ]);
  return "Product Saved";
}
