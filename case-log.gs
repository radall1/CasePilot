// Define named constants for template IDs and destination folder ID
const TEMPLATE_STUDENT_0 = '1YZeAV9U9QGehjz89AYtw2I_jhsq_6xgBO32ASH-k2do';      // URL for Case Report
const TEMPLATE_STUDENT_1 = '1XVDjJXlKSvu2UbuWFNk2__M0LVT4a3yi2wrlmq44nzA';      // URL for Notification Memo
const TEMPLATE_STUDENT_2 = '1WaSLVn8m_7U38928jXgB9Cy6zx396c1rO5xpHEXg0B4';      // URL for Materials Request
const TEMPLATE_STUDENT_3 = '1XesdryqUjzPkJcAJ6GE_HQ623pMj9osldfR2baBjHlE';      // URL for Implicating Student Testimony
const TEMPLATE_STUDENT_4 = '1KjmsIR34dDJcfmD80uRPb7Iw7tT5oxmFNLyXt3kz6C0';      // URL for Interview Invitation
const TEMPLATE_STUDENT_5 = '1ARYEm2TY9eeelMyZri8i_qClA6TLnDkos8TlFJg8rDg';      // URL for Case Debriefing
const TEMPLATE_STUDENT_6 = '13js-uRqq2KlQh2FI454s7BM40ZsAxp6NGy1DX5GPhn4';      // URL for Professor Notification

const TEMPLATE_FACULTY_0 = '1bpQKf_ADN7WRNMv4i_bWyHGk_jELEsArRHIfpH_8OjU';      // URL for Case Report
const TEMPLATE_FACULTY_1 = '1RO27FGHuEPdoNfsqP2OlyWKv1bURo5QdLxENjsCTTUk';      // URL for Notification Memo
const TEMPLATE_FACULTY_2 = '1jSN5S0Vr33IKGLgbWf7DS7BvnPGfgwtO2Me2g3OPqGs';      // URL for Materials Request
const TEMPLATE_FACULTY_4 = '1q2wlJYivleIZ8uzmh3KtZ60jhCErwvxzPFdx10JRLSw';      // URL for Interview Invitation
const TEMPLATE_FACULTY_5 = '1h4xI36stoa-xtS2VeGhH4hPN73l5L7K1k_jv-CGNlcA';      // URL for Case Debriefing
const TEMPLATE_FACULTY_6 = '1Bewj8O5o3ynASZ46df8QATNADBp0Zle5SkYeeO7VWJ0';      // URL for Professor Notification

const DESTINATION_FOLDER = '1bn8rbz56HrvLcX6jiH5yP_AhqpFQyGIe';

// Define placeholders for student and faculty cases
const PLACEHOLDERS_STUDENT = {
  'case number': 0,
  'ProfFN': 2,
  'ProfLN': 3,
  'dept': 25,
  'class': 26,
  'section': 27,
  'courseterm': 28,
  'examdaydate': 29,
  'examnumber': 30,
  'examtopic': 31,
  'examformat': 32,
  'studentname': 33,
  'relationship': 34,
  'resources': 35
};

const PLACEHOLDERS_FACULTY = {
  'case number': 0,
  'ProfFN': 2,
  'ProfLN': 3,
  'affiliation': 4,
  'anon': 6,
  'dept': 7,
  'class': 8,
  'section': 9,
  'courseterm': 10,
  'examdaydate': 11,
  'examnumber': 12,
  'examtopic': 13,
  'examformat': 14,
  'studentname': 15,
  'examgrade': 16,
  'discussionoutcome': 17,
  'implicationdes': 18,
  'desiredres': 19,
  'relationship': 20,
  'resources': 21,
  'additional': 22,
  'colleague': 23,
  'followup': 24
};

// Triggered when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Case Management');
   
  // Get data from the 'Form Center' sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Center');
  const rows = sheet.getDataRange().getValues();

  // Add menu items based on case availability
  if (!rows[1]) {
    menu.addItem('No Cases Identified', 'empty');
  } else {
    menu.addItem(`Create Case Material for Case ${rows[1][0]}`, 'caseReport');
  }
  
  menu.addToUi();
}

// Placeholder for empty function
function empty() {}

// Replaces placeholders in the document body with actual values
function replacePlaceholders(body, placeholders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Center');
  const rows = sheet.getDataRange().getValues();
  
  for (const placeholder in placeholders) {
    body.replaceText(`{{${placeholder}}}`, rows[1][placeholders[placeholder]]);
  }
}

// Generates a document from a template and replaces placeholders
function generateDocumentFromTemplate(templateId, destinationFolder, fileName, placeholders) {
  try {
    const googleDocTemplate = DriveApp.getFileById(templateId);

    const copy = googleDocTemplate.makeCopy(fileName, destinationFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    replacePlaceholders(body, placeholders);

    doc.saveAndClose();
  } catch (error) {
    Logger.log(`Error generating document: ${error}`);
  }
}

// Generates case materials for a student case
function caseStudent(rows) {

  const caseFolder = DriveApp.getFolderById(DESTINATION_FOLDER).createFolder(`${rows[1][0]}`);

  generateDocumentFromTemplate(TEMPLATE_STUDENT_0, caseFolder, `${rows[1][0]} - Case Report`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_1, caseFolder, `${rows[1][0]} - Notification Memo`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_2, caseFolder, `${rows[1][0]} - Materials Request`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_3, caseFolder, `${rows[1][0]} - Implicating Student Testimony`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_4, caseFolder, `${rows[1][0]} - Interview Invitation`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_5, caseFolder, `${rows[1][0]} - Case Debriefing`, PLACEHOLDERS_STUDENT);
  generateDocumentFromTemplate(TEMPLATE_STUDENT_6, caseFolder, `${rows[1][0]} - Professor Notification`, PLACEHOLDERS_STUDENT);
}

// Generates case materials for a faculty case
function caseFaculty(rows) {

  const caseFolder = DriveApp.getFolderById(DESTINATION_FOLDER).createFolder(`${rows[1][0]}`);

  generateDocumentFromTemplate(TEMPLATE_FACULTY_0, caseFolder, `${rows[1][0]} - Case Report`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_1, caseFolder, `${rows[1][0]} - Notification Memo`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_2, caseFolder, `${rows[1][0]} - Materials Request`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_4, caseFolder, `${rows[1][0]} - Interview Invitation`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_5, caseFolder, `${rows[1][0]} - Case Debriefing`, PLACEHOLDERS_FACULTY);
  generateDocumentFromTemplate(TEMPLATE_FACULTY_6, caseFolder, `${rows[1][0]} - Professor Notification`, PLACEHOLDERS_FACULTY);
}

// Entry point for generating case materials
function caseReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Center');
  const rows = sheet.getDataRange().getValues();

  // Determine case type and generate materials accordingly
  if (rows[1][5] === 'Student-Implicated: Faculty/staff member reporting a student-implicated Honor System violation.') {
    caseStudent(rows);
  } else if (rows[1][5] === 'Faculty-Implicated: Faculty/staff member reporting a faculty-implicated Honor System violation.') {
    caseFaculty(rows);
  }
}
