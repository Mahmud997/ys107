function doGet(e) {
  const action = e.parameter.action;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === 'login') {
    const login = e.parameter.login;
    const password = e.parameter.password;
    const usersSheet = sheet.getSheetByName('users');
    const usersData = usersSheet.getDataRange().getValues();
    
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] === login && usersData[i][1] === password) {
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          role: usersData[i][2],
          name: usersData[i][3]
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getStudents') {
    const studentsSheet = sheet.getSheetByName('students');
    const data = studentsSheet.getDataRange().getValues();
    const students = data.slice(1).map(row => ({ id: row[0], name: row[1], class: row[2] }));
    return ContentService.createTextOutput(JSON.stringify(students))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getRecords') {
    const recordsSheet = sheet.getSheetByName('records');
    const data = recordsSheet.getDataRange().getValues();
    const records = data.slice(1).map(row => ({ 
      id: row[0], studentId: row[1], date: row[2], 
      time: row[3], status: row[4], note: row[5], teacher: row[6] 
    }));
    return ContentService.createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'addRecord') {
    const recordsSheet = sheet.getSheetByName('records');
    const row = [
      Date.now().toString(36) + Math.random().toString(36).slice(2,6),
      e.parameter.studentId,
      e.parameter.date,
      e.parameter.time,
      e.parameter.status,
      e.parameter.note || '',
      e.parameter.teacher || 'Дежурный'
    ];
    recordsSheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'addStudent') {
    const studentsSheet = sheet.getSheetByName('students');
    const row = [
      Date.now().toString(36) + Math.random().toString(36).slice(2,6),
      e.parameter.name,
      e.parameter.class
    ];
    studentsSheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'removeStudent') {
    const studentsSheet = sheet.getSheetByName('students');
    const data = studentsSheet.getDataRange().getValues();
    let rowToDelete = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === e.parameter.id) {
        rowToDelete = i + 1;
        break;
      }
    }
    if (rowToDelete > 0) {
      studentsSheet.deleteRow(rowToDelete);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'removeStudentRecords') {
    const recordsSheet = sheet.getSheetByName('records');
    const data = recordsSheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === e.parameter.studentId) {
        rowsToDelete.push(i + 1);
      }
    }
    rowsToDelete.sort((a, b) => b - a);
    for (const row of rowsToDelete) {
      recordsSheet.deleteRow(row);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Неизвестное действие' }))
    .setMimeType(ContentService.MimeType.JSON);
}
