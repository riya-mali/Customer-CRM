function doGet() {
  // Yeh line aapke naye dark UI (index1) ko screen par load karegi
  return HtmlService.createTemplateFromFile('index1').evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle("Nexus Support CRM v2");
}

const CENTRAL_REGISTRY_ID = "1oUReZFrtTWHN9I5e-53fM20MYjheIAWFxSkjrXRLd78";

// 1. DATA READ KARNE KA FUNCTION
function downloadMatrixStream() {
  var activeWorkbook = SpreadsheetApp.openById(CENTRAL_REGISTRY_ID);
  var mainRecordsTab = activeWorkbook.getSheets()[0]; 
  var rawMatrixData = mainRecordsTab.getDataRange().getValues();
  var sanitizedResponseArray = [];
  
  for (var rowIdx = 1; rowIdx < rawMatrixData.length; rowIdx++) {
    if(rawMatrixData[rowIdx][0]) { 
      sanitizedResponseArray.push({
        ticket_id: rawMatrixData[rowIdx][0],
        customer_name: rawMatrixData[rowIdx][1],
        customer_email: rawMatrixData[rowIdx][2],
        subject: rawMatrixData[rowIdx][3],
        description: rawMatrixData[rowIdx][4],
        status: rawMatrixData[rowIdx][5],
        created_at: rawMatrixData[rowIdx][6],
        priority: rawMatrixData[rowIdx][8] || "Standard"
      });
    }
  }
  return JSON.stringify(sanitizedResponseArray);
}

// 2. DATA WRITE KARNE KA FUNCTION (Ab yahan Tamanna ya jo likhogi wahi save hoga)
function injectNewServiceTicket(payloadData) {
  var activeWorkbook = SpreadsheetApp.openById(CENTRAL_REGISTRY_ID);
  var mainRecordsTab = activeWorkbook.getSheets()[0];
  var generatedId = "TKT-" + Math.floor(100000 + Math.random() * 899999);
  var systemTimestamp = new Date().toISOString();
  
  mainRecordsTab.appendRow([
    generatedId,
    payloadData.customer_name,   // Column B: Name
    payloadData.customer_email,  // Column C: Email
    payloadData.subject,         // Column D: Subject
    payloadData.description,     // Column E: Description
    "Open",                      // Column F: Status
    systemTimestamp,             // Column G: Timestamp
    systemTimestamp,             // Column H: Updated At
    payloadData.priority || "Standard" // Column I: Priority
  ]);
  return JSON.stringify({ status: "SUCCESS_INJECT", ticket_id: generatedId });
}

// 3. TICKET DETAILS READ KARNE KA FUNCTION
function extractSingleIdentityBlock(targetTicketId) {
  var activeWorkbook = SpreadsheetApp.openById(CENTRAL_REGISTRY_ID);
  var mainRecordsTab = activeWorkbook.getSheets()[0];
  var timelineNotesTab = activeWorkbook.getSheets()[1]; 
  
  var primaryDataset = mainRecordsTab.getDataRange().getValues();
  var matchingResultNode = null;
  
  for (var i = 1; i < primaryDataset.length; i++) {
    if (primaryDataset[i][0] == targetTicketId) {
      matchingResultNode = {
        ticket_id: primaryDataset[i][0],
        customer_name: primaryDataset[i][1],
        customer_email: primaryDataset[i][2],
        subject: primaryDataset[i][3],
        description: primaryDataset[i][4],
        status: primaryDataset[i][5],
        created_at: primaryDataset[i][6],
        priority: primaryDataset[i][8] || "Standard",
        notes: []
      };
      break;
    }
  }
  
  if (matchingResultNode && timelineNotesTab) {
    try {
      var historicNotesData = timelineNotesTab.getDataRange().getValues();
      for (var j = 1; j < historicNotesData.length; j++) {
        if (historicNotesData[j][0] == targetTicketId) {
          matchingResultNode.notes.push({
            note_text: historicNotesData[j][1],
            created_at: historicNotesData[j][2]
          });
        }
      }
    } catch(err) {}
  }
  return JSON.stringify(matchingResultNode);
}

// 4. STATUS CHANGE KARNE KA FUNCTION
function modifyServiceTicketState(stateMutationRequest) {
  var activeWorkbook = SpreadsheetApp.openById(CENTRAL_REGISTRY_ID);
  var mainRecordsTab = activeWorkbook.getSheets()[0];
  var timelineNotesTab = activeWorkbook.getSheets()[1];
  var currentActionTime = new Date().toISOString();
  
  var rowRecords = mainRecordsTab.getDataRange().getValues();
  for (var r = 1; r < rowRecords.length; r++) {
    if (rowRecords[r][0] == stateMutationRequest.ticket_id) {
      mainRecordsTab.getRange(r + 1, 6).setValue(stateMutationRequest.status); 
      mainRecordsTab.getRange(r + 1, 8).setValue(currentActionTime); 
      break;
    }
  }
  
  if (stateMutationRequest.notes && stateMutationRequest.notes.trim() !== "" && timelineNotesTab) {
    timelineNotesTab.appendRow([
      stateMutationRequest.ticket_id, 
      stateMutationRequest.notes, 
      currentActionTime
    ]);
  }
  return JSON.stringify({ status: "SUCCESS_MUTATION" });
}