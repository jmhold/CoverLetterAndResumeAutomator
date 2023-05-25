//This lets you to have different templates for different "pratices" or types of roles you want to apply to
//
const Practices = Object.freeze({
  Executive:"exec",
  Management:"mgt",
  Researcher:"rsch",
  Product:"prd"
})

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Resume');
  menu.addItem('Generate Docs', 'generateDocs');
  menu.addToUi();
}

function generateDocs() {
  const rootFolder = DriveApp.getFolderById('ROOT_FOLDER_ID'); //Root Folder
  // const rootFolder = DriveApp.getFolderById('TEST_FOLDER_ID'); //Test Folder
  const today = new Date().toLocaleDateString('en-US', {year: 'numeric', month: 'long', day: 'numeric'});

  const entries = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Entries'); //Sheet you are using to add job application info

  const rows = entries
    .getDataRange()
    .getValues();


  rows.forEach(function(row, index){
    const docData = {
      jobLink: String(row[0]).trim(),
      orgName: String(row[1]).trim(),
      hiringManager: String(row[2]).trim(),
      role: String(row[3]).trim(),
      practice: String(row[4]).trim(),
      keyPhrases: String(row[5]).trim(),
      coverLetterCell: row[6],
      resumeCell: row[7],
      today: today,
    }

    if (index === 0) return; // This returns true if it's the first row (where the headings are)
    if (docData.coverLetterCell && docData.resumeCell) return; // This returns true if there is something in the "Cover Letter Link" and "Resume Link" cells
        
    const destinationFolder = detectFolder(rootFolder, docData.orgName) ? DriveApp.getFolderById(detectFolder(rootFolder, docData.orgName)) : rootFolder.createFolder(docData.orgName);
    
    if (!docData.coverLetterCell) {
      entries.getRange(index + 1, 7).setRichTextValue(generateCoverLetter(docData, destinationFolder));
    }
    
    if (!docData.resumeCell) {
      entries.getRange(index + 1, 8).setRichTextValue(generateResume(docData, destinationFolder));
    }
    
    if(String(docData.jobLink).startsWith('http')) entries.getRange(index + 1, 1).setValue(formatHyperlink(docData.jobLink));

  })  

}

function generateCoverLetter(docData, destinationFolder) {
  var coverLeterTemplateId;

  switch(docData.practice) {
    case Practices.Executive:
      coverLeterTemplateId = 'EXEC_COVER_LETTER_TEMPLATE_ID'
      break;
    case Practices.Management:
      coverLeterTemplateId = 'MGT_COVER_LETTER_TEMPLATE_ID'
      break;
    case Practices.Researcher:
      coverLeterTemplateId = 'RSCH_COVER_LETTER_TEMPLATE_ID'
      break;
    case Practices.Product:
      coverLeterTemplateId = 'PRD_COVER_LETTER_TEMPLATE_ID' 
      break;
    default:
      coverLeterTemplateId = 'DEFAULT_COVER_LETTER_TEMPLATE_ID'
      break;
  }

  const coverLetterTemplate = DriveApp.getFileById(coverLeterTemplateId); //Cover Letter Template File

  docData.hiringManager = docData.hiringManager ? docData.hiringManager : 'Hiring Manager';

  const formattedOrgName = formatString(docData.orgName);
  const formattedRole = formatString(docData.role);

  const copy = coverLetterTemplate.makeCopy(`Jarrod_Holder_${formattedOrgName}_${formattedRole}_Cover_Letter`, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText('{{hiring_manager}}', docData.hiringManager);
  body.replaceText('{{org_name}}', docData.orgName);
  body.replaceText('{{todays_date}}', docData.today);

  doc.saveAndClose();

  const pdfBlob = doc.getAs('application/pdf');
  const pdfDoc = destinationFolder.createFile(pdfBlob);

  return SpreadsheetApp.newRichTextValue()
    .setText("doc | pdf")
    .setLinkUrl(0,3, doc.getUrl())
    .setLinkUrl(6,9,pdfDoc.getUrl())
    .build();
}

function generateResume(docData, destinationFolder) {
  var resumeTemplateId; 

  switch(docData.practice) {
    case Practices.Executive:
      resumeTemplateId = 'EXEC_RESUME_TEMPLATE_ID'
      break;
    case Practices.Management:
      resumeTemplateId = 'MGT_RESUME_TEMPLATE_ID'
      break;
    case Practices.Researcher:
      resumeTemplateId = 'RSCH_RESUME_TEMPLATE_ID'
      break;
    case Practices.Product:
      resumeTemplateId = 'PRD_RESUME_TEMPLATE_ID'
      break;
    default:
      resumeTemplateId = 'DEFAULT_RESUME_TEMPLATE_ID'
      break;
  }

  const resumeTemplate = DriveApp.getFileById(resumeTemplateId)
  const seperator = " ▪ "
  
  const formattedOrgName = formatString(docData.orgName);
  const formattedRole = formatString(docData.role);
  
  const mgtPhrases = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Mgt_Phrases');

  const rows = mgtPhrases
    .getDataRange()
    .getValues();
  
  const keyPhrasesArray = new String(docData.keyPhrases).split("//");
  
  var specializedPhrasesArray = [];
  var wwtJobDescPhrasesArray = [];
  rows.forEach(function(row, index){
    if (index === 0) return; // This returns true if it's the first row (where the headings are)
    row[0] ? specializedPhrasesArray.push(row[0]) : null;
    row[2] ? wwtJobDescPhrasesArray.push(row[2]) : null;
  })

  const specializedPhrases = specializedPhrasesArray.join(seperator);
  const jobDescKeyPhrases = keyPhrasesArray.join(seperator);
  const wwtJobDescPhrases = wwtJobDescPhrasesArray.join(`\n`);

  const copy = resumeTemplate.makeCopy(`Jarrod_Holder_${formattedOrgName}_${formattedRole}_Resume`, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText('{{specialized_skill_set}}', specializedPhrases);
  body.replaceText('{{job_desc_key_phrases}}', jobDescKeyPhrases);
  body.replaceText('{{wwt_job_desc}}', wwtJobDescPhrases);

  doc.saveAndClose();

  const pdfBlob = doc.getAs('application/pdf');
  const pdfDoc = destinationFolder.createFile(pdfBlob);

  return SpreadsheetApp.newRichTextValue()
    .setText("doc | pdf")
    .setLinkUrl(0,3, doc.getUrl())
    .setLinkUrl(6,9,pdfDoc.getUrl())
    .build();
}

function detectFolder(rootFolder, folderName) {
  const folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    var folder = folders.next();
    return folder.getId();
  }
  return false;
}

function formatHyperlink(link) {
  return `=HYPERLINK("${link}", "Link")`;
}

function formatString(str) {
  return new String(str).replaceAll(" ", "_");
}
