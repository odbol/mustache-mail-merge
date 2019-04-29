var loadingImg = 'https://lh6.googleusercontent.com/-S87nMBe6KWE/TuB9dR48F0I/AAAAAAAAByQ/0Z96LirzDqg/s27/load.gif';


function gmailGetGroups() {
    return _.map(ContactsApp.getContactGroups(), function (g) {
        return {
            name: g.getName()
        }
    });
}

function processImportForm(formObject) {
    return importGroup(formObject.groups);
}



function selectGroup() {
    var html = HtmlService
      .createTemplateFromFile('ImportContacts');
  SpreadsheetApp.getUi()
      .showModalDialog(html.evaluate(), 'Import Contacts');
}

function importGroup(groups) {
    var headers = createHeaderIfNotFound_('Full Name');
    headers = createHeaderIfNotFound_('First Name');
    headers = createHeaderIfNotFound_('Last Name');
    headers = createHeaderIfNotFound_('Email Address');
    headers = createHeaderIfNotFound_('Company');
    var sheet = ss.getActiveSheet();
    var group = ContactsApp.getContactGroup(groups);
    var contacts = ContactsApp.getContactsByGroup(group);
    var row = sheet.getLastRow() + 1;
    for (i in contacts) {
        sheet.getRange(row, headers.indexOf('Full Name') + 1).setValue(contacts[i].getFullName());
        sheet.getRange(row, headers.indexOf('First Name') + 1).setValue(contacts[i].getGivenName());
        sheet.getRange(row, headers.indexOf('Last Name') + 1).setValue(contacts[i].getFamilyName());
        if (contacts[i].getEmails()[0] != undefined) sheet.getRange(row, headers.indexOf('Email Address') + 1).setValue(contacts[i].getEmails()[0].getAddress());
        if (contacts[i].getCompanies()[0] != undefined) sheet.getRange(row, headers.indexOf('Company') + 1).setValue(contacts[i].getCompanies()[0].getCompanyName());
        // Add custom fields  
        var customFields = contacts[i].getCustomFields();
        for (j in customFields) {
            var label = customFields[j].getLabel();
            if (headers.indexOf(label) == -1) headers = createHeaderIfNotFound_(label);
            sheet.getRange(row, headers.indexOf(label) + 1).setValue(customFields[j].getValue());
        }
        row++;
    }

    return contacts.length;
}

function createHeaderIfNotFound_(value) {
    var sheet = ss.getActiveSheet();
    var lastColumn = sheet.getLastColumn();
    if (lastColumn == 0) {
        sheet.getRange(1, lastColumn + 1).setValue(value);
        return lastColumn;
    } else {
        var headers = sheet.getRange(1, 1, 1, lastColumn).getValues();
        if (headers[0].indexOf(value) == -1) {
            sheet.getRange(1, lastColumn + 1).setValue(value);
            headers[0].push(value);
        }
    }
    return headers[0];
}
