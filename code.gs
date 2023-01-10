function autoFillPSTemplateGoogleDoc(e) {

    // declare variables from Google Sheet
    let timeStamp = e.values[0];
    let investorName = e.values[1];
    let emailID = e.values[2]
    
    // convert values from column 5 of Google Sheet to string
    const goals = e.values[4].toString();

    // convert values from column 27 of Google Sheet to string
    const returnExpectationString = e.values[26].toString();
    let returnExpectation = parseFloat(returnExpectationString);

    // convert values from column 17 of Google Sheet to string
    let actualBalanceString = e.values[16].toString();
    let actualBalance = parseFloat(actualBalanceString);

    let evaluationFrequency = e.values[23].toString();


    // declare goal variables
    let goal1 = ""
    let goal2 = ""
    let goal3 = ""
    
    //create an array and parse values from CSV format, store them in an array
    goalsArr = goals.split(',')
    if (goalsArr.length >= 1)
        goal1 = goalsArr[0]
        if (goalsArr.length >= 2)
            goal2 = goalsArr[1]
            if (goalsArr.length >= 3)
                goal3 = goalsArr[2]
                
                //grab the template file ID to modify
                const file = DriveApp.getFileById(templateID);
    
    //grab the Google Drive folder ID to place the modied file into
    var folder = DriveApp.getFolderById(folderID)
    
    //create a copy of the template file to modify, save using the naming conventions below
    var copy = file.makeCopy(investorName + ' Investment Policy', folder);
    
    //modify the Google Drive file
    var doc = DocumentApp.openById(copy.getId());
    
    var body = doc.getBody();
    
    body.replaceText('%InvestorName%', investorName);
    body.replaceText('%Date%', timeStamp);
    
    body.replaceText('%Goal1%', goal1.trim())
    body.replaceText('%Goal2%', goal2.trim())
    body.replaceText('%Goal3%', goal3.trim())

    // Calling the handleRiskTolerance function/method
    handleRiskTolerance(body, returnExpectation);

    // Calling the assessFinancialKnowledge function/method
    assessFinancialKnowledge(actualBalance)
    // Callling the evaluationFrequency function/method
    getEvaluationFrequency(body, evaluationFrequency)

    doc.saveAndClose();
    
    //find the file that was just modified, convert to PDF, attach to e-mail, send e-mail
    var attach = DriveApp.getFileById(copy.getId());
    var pdfattach = attach.getAs(MimeType.PDF);
    MailApp.sendEmail(emailID, subject, emailBody, { attachments: [pdfattach] });
}

// This function determines the risk tolerance the user is willing to take based on their return expectations
function handleRiskTolerance(body, returnExpectation) {
  if (returnExpectation <= 8) {
    body.replaceText('%risk tolerance%', 'low');
  } else if (returnExpectation > 8 && returnExpectation <= 12) {
    body.replaceText('%risk tolerance%', 'medium');
  } else if (returnExpectation > 12) {
    body.replaceText('%risk tolerance%', 'high');
  } else {
    body.replaceText('%risk tolerance%', 'low (default)');
  }
}

//The function compares the expected balance to the actual balance and returns a string indicating the user's financial knowledge level.
function assessFinancialKnowledge(actualBalance) {
  let expectedBalance = 110.41;
  if (expectedBalance === actualBalance) {
    return 'expert';
  } else if (Math.abs(expectedBalance - actualBalance) < 0.1 * expectedBalance) {
    return 'advanced';
  } else if (Math.abs(expectedBalance - actualBalance) < 0.25 * expectedBalance) {
    return 'intermediate';
  } else {
    return 'novice';
  }
}

function getEvaluationFrequency(body, evaluationFrequency) {
  switch (evaluationFrequency) {
    case 'Annually':
      body.replaceText('%evaluationFrequency%', 'Annually');
    case 'Quarterly':
      body.replaceText('%evaluationFrequency%', 'Quarterly');
    case 'Semi-annually':
      body.replaceText('%evaluationFrequency%', 'Semi-annually');
    case 'Monthly':
      body.replaceText('%evaluationFrequency%', 'Monthly');
    default:
      body.replaceText('%evaluationFrequency%', 'Annually (default)');
  }
}
