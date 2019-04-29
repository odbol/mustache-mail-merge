var loadingImg = 'https://lh6.googleusercontent.com/-S87nMBe6KWE/TuB9dR48F0I/AAAAAAAAByQ/0Z96LirzDqg/s27/load.gif';
var ss = SpreadsheetApp.getActiveSpreadsheet();

function onInstall() {
  onOpen();
}

function onOpen() {
  ss.addMenu("Mail Merge", [{
    name: "Standard Merge",
    functionName: "openDialogStandardMerge"
  }, {
    name: "Import contacts",
    functionName: "selectGroup"
  }]);
  /*,{
    name: "Scheduled mail merge",
    functionName: "startingPageforScheduledMerge"
  }*/
}



function openDialogStandardMerge() {
  var html = HtmlService
      .createTemplateFromFile('SelectCampaign');
  SpreadsheetApp.getUi()
      .showModalDialog(html.evaluate(), 'Mail Merge');
}

function gmailGetAliases() {
  var userEmail = Session.getEffectiveUser().getEmail();
  var chosenFrom = [userEmail];
  var aliases = GmailApp.getAliases();
  if(aliases != null && aliases.length > 0){
    for(i in aliases) chosenFrom.push(aliases[i]);
  }

  return aliases;
}



function gmailGetDrafts() {
  var templates = GmailApp.search("in:drafts");
  return _.map(templates, function (template) {
    return {
      subject: template.getFirstMessageSubject(),
      id: template.getId()
    };
  })
  .filter(function(t) { return t.subject; });
}

function gmailGetFromName() {
  return UserProperties.getProperty('mustacheChosenName') || Session.getEffectiveUser().getEmail();
}

function gmailGetGlobalCC() {
  return UserProperties.getProperty('mustacheGlobalCC');
}

function processForm(formObject) {
  var selectedTemplate = GmailApp.getThreadById(formObject.chosenTemplate).getMessages()[0];
  var user = Session.getEffectiveUser().getEmail();
  var name = formObject.chosenName;
  var from = formObject.chosenFrom;
  var cc = formObject.ccAddr;
  
  // save choices for multiple runnings
  UserProperties.setProperty('mustacheChosenName', name);
  UserProperties.setProperty('mustacheGlobalCC', cc);
  
  merge('gmail', selectedTemplate, name, from, cc);

  return true;
}



function merge(kind, selectedTemplate, name, from, cc) {
  var dataSheet = ss.getActiveSheet();
  var headers = createHeaderIfNotFound_('Merge status');
  var dataRange = dataSheet.getDataRange();
  //////////////////////////////////////////////////////////////////////////////
  // Get inline images and make sure they stay as inline images
  //////////////////////////////////////////////////////////////////////////////
  var emailTemplate = selectedTemplate.getBody();
  var rawContent = selectedTemplate.getRawContent();
  var attachments = selectedTemplate.getAttachments();
  cc = cc || selectedTemplate.getCc();
  var bcc = selectedTemplate.getBcc();

  //Logger.log("id: " + selectedTemplate.getId());

  //Logger.log(emailTemplate);
  //Logger.log("rawcontent: ");
  //Logger.log(rawContent);

  //Logger.log("attachments: ");
  //Logger.log(attachments);


  var regMessageId = new RegExp(selectedTemplate.getId(), "g");
  //if (emailTemplate.match(regMessageId) != null) {
    var inlineImages = {};
    var imgVars = emailTemplate.match(/<img[^>]+>/g);
    var imgToReplace = [];
    if(imgVars != null){
      for (var i = 0; i < imgVars.length; i++) {
        //Logger.log("imgVars: " + imgVars[i]);

        //if (imgVars[i].search(regMessageId) != -1) {
          var id = imgVars[i].match(/cid:([^&"]+)[&"]/);
          //Logger.log("imgVars id: " + id);
          if (id != null) {
            id = id[1];
            var temp = rawContent.split(id)[1];
            temp = temp.substr(temp.lastIndexOf('Content-Type'));
            var imgTitle = imgVars[i].match(/alt="([^"]+)"/);
            if (imgTitle) {
              imgTitle = imgTitle[1];
            }
            var contentType = temp.match(/Content-Type: ([^;]+);/);
            contentType = (contentType != null) ? contentType[1] : "image/jpeg";
            var b64c1 = rawContent.lastIndexOf(id) + id.length + 3; // first character in image base64
            var b64cn = rawContent.substr(b64c1).indexOf("--") - 3; // last character in image base64
            var imgb64 = rawContent.substring(b64c1, b64c1 + b64cn + 1); // is this fragile or safe enough?
            var imgblob = Utilities.newBlob(Utilities.base64Decode(imgb64), contentType, id); // decode and blob
            imgToReplace.push([imgTitle, imgVars[i], id, imgblob]);

            //Logger.log("imgToReplace: " + imgTitle + " : " + imgVars[i] + " : " + id);
          }
        //}
      }
    }
    for (var i = 0; i < imgToReplace.length; i++) {
      inlineImages[imgToReplace[i][2]] = imgToReplace[i][3];
      var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + imgToReplace[i][2] + "\"");
      emailTemplate = emailTemplate.replace(imgToReplace[i][1], newImg);
    }
  //}
  //////////////////////////////////////////////////////////////////////////////
  var mergeData = {
    template: emailTemplate,
    subject: selectedTemplate.getSubject(),
    plainText : selectedTemplate.getPlainBody(),
    attachments: attachments,
    name: name,
    from: from,
    cc: cc,
    bcc: bcc,
    inlineImages: inlineImages
  }
  
  var objects = getRowsData(dataSheet, dataRange);
  for (var i = 0; i < objects.length; ++i) {
    var rowData = objects[i];
    if (rowData.mergeStatus != "Done" && rowData.mergeStatus != "0") {
      try {
        processRow(rowData, kind, mergeData);
        dataSheet.getRange(i + 2, headers.indexOf('Merge status') + 1).setValue("Done").clearFormat().setComment(new Date());
      }
      catch (e) {
        dataSheet.getRange(i + 2, headers.indexOf('Merge status') + 1).setValue("Error").setBackground('red').setComment(e.message);
      }
    }
  }
}

function processRow(rowData, kind, mergeData) {
  if (!rowData.emailAddress) 
    throw { message: "Missing email address" };
  
  
  var emailText = fillInTemplateFromObject(mergeData.template, rowData);
  var emailSubject = fillInTemplateFromObject(mergeData.subject, rowData);
  var plainTextBody = fillInTemplateFromObject(mergeData.plainText, rowData);
  mergeData['htmlBody'] = emailText;
  if(rowData.cc != undefined) mergeData.cc = rowData.cc;
  if(rowData.bcc != undefined) mergeData.bcc = rowData.bcc;
  GmailApp.sendEmail(rowData.emailAddress, emailSubject, plainTextBody, mergeData);
}

// holds the compiled mustache template for speed.
var mustacheTemplate = null;

// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance <<Column name>>
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker <<Column name>>
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  ////Logger.log('got template: ' + template);
  
  var  normalizedData = {},
       result;
  
  _.each(data, function(value, key, list) {
    normalizedData[normalizeHeader(key)] = value ? value.toString().trim() : '';
  });
  result = Mustache.render(template, normalizedData);
  
  ////Logger.log('rendered: ' + result);
  ////Logger.log('rendered with vars: ' + _.keys(normalizedData).join(',') + _.values(normalizedData).join(','));  
  return result;
}