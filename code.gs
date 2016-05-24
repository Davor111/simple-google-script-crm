/* 
Simple CRM based on Google App Script by David Mitterer

HOW TO: 

Provide sheet IDs for Client and User sheet. Just create two new Google Sheets and copy the ID from the URL.
For example: 
https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_ID

Give it access to your Drive Account when asked. 
IMPORTANT: leave the first row BLANK or provide row names like "id, name, address, etc" for your data. 

This feautres address validation via the built in Google Maps Geocoder API. 
Address validation works only for German addresses as this was built for a German company. 

Click on 'Add Client' to see it.  

*/

var sheetUserID = "INSERT_YOUR_ID_FOR_YOUR_USER_HERE"; // 
var sheetClient = "INSERT_YOUR_ID_FOR_YOUR_CLIENTS_HERE"; //


function saveUser(user, password) {
    //you can use only one sheet. Just change the array at the end instead of the ID 
    var sheet = SpreadsheetApp.openById(sheetUserID).getSheets()[0];

    var lastRow = sheet.getLastRow();

    var lastCell = sheet.getRange(lastRow, 4);
    var userID = lastCell.getValue() + 1;
    sheet.appendRow([Date(), user, password, userID])

    return userID

};


function checkPassword(user, password) {
    var sheet = SpreadsheetApp.openById(sheetUserID).getSheets()[0];
    var lastRow = sheet.getLastRow();
    var result = sheet.getRange(2, 2, lastRow, 3).getValues();
    Logger.log(result);

    var i = 0;
    for (i; i < lastRow; i++) {
        if (user == result[i][0]) {

            if (password == result[i][1]) {
                return result[i]
            }

        }
    }
    return false
}


// gets all clients from list 
function getClientList() {
    var sheet = SpreadsheetApp.openById(sheetClient).getSheets()[0];
    var lastRow = sheet.getLastRow()
    var result = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // -1 bc for some reason empty row appears #rewrite

    return makeNiceFormat(result)
}


// various search methods
function searchStudent(search, index, callback) {
    var sheet = SpreadsheetApp.openById(sheetClient).getSheets()[0];
    var lastRow = sheet.getLastRow();
    var result = sheet.getRange(2, 1, lastRow, 11).getValues();
    var i = 0;


    // index indicates what is searched for in sheet (0 = id, 1 = name, 2 = address, ..". If undefined then name is searched for
    if (!index) { index = 1 };


    switch (index) {

        case "0": // faster searching by stopping after first hit (id is unique)
            for (i; i < lastRow; i++) {
                if (search == result[i][index]) {
                    return result[i];
                };
            };

        case "idNiceFormat": // faster searching by stopping after first hit, but with nice Object returned
            var searchResult = []
            for (i; i < lastRow; i++) {
                if (search == result[i][0]) {
                    searchResult.push(result[i]);
                    return makeNiceFormat(searchResult);
                };

            };

        case "row": // special because if only row index is needed
            for (i; i < lastRow; i++) {
                if (search == result[i][0]) {
                    callback(i);
                    break;
                }

            }


        default: // index = the row in the sheet
            var searchResult = [];
            for (i; i < lastRow; i++) {
                var sear = new RegExp(search, "i", "m");
                if (sear.test(result[i][index])) {
                    searchResult.push(result[i])
                }
            }

            return makeNiceFormat(searchResult);

    };

};


function makeNiceFormat(result) {
    var clientData = [];
    var i = 0;
    for (i; i < result.length; i++) {
        clientData.push({
            id: result[i][0],
            name: result[i][1],
            address: result[i][2],
            plz: result[i][3],
            city: result[i][4],
            country: result[i][5],
            comment: result[i][6],
            consultant: result[i][7],
            consultantID: result[i][8],
            telephone: result[i][9],
            email: result[i][10]
        }
        );
    }
    return clientData

};



// gets Google Maps address for verifying data
function getGeoData(searchAdd) {

    var response = Maps.newGeocoder()
        // sets region to Germany 
        .setRegion('de')
        .setBounds(51.02503, 5.2959, 51.22537, 16.10119) // change them as you want
        .geocode(searchAdd);

    return response

};


//adds new Client

function addNewClient(client) {

    var sheet = SpreadsheetApp.openById(sheetClient).getSheets()[0];
    var lastRow = sheet.getLastRow();

    var result = sheet.getRange(2, 1, lastRow, 1).getValues();

    //fix for number sortin and gets new ID for student
    function sortNumber(a, b) {
        return a - b;
    }
    result = result.sort(sortNumber);

    var lastId = result[result.length - 1];
    lastId = parseInt(lastId) + 1;

    sheet.appendRow([lastId, client.name, client.address, client.plz, client.city, client.country, client.comment, client.consultant, client.consultantID, client.telephone, client.email])

    return "Client successfully added"
};


// deletes Client
function deleteClient(id) {
    var sheet = SpreadsheetApp.openById(sheetClient).getSheets()[0];

    searchStudent(id, "row", function (x) {
        sheet.deleteRow(x + 2)
    })

};


//edit Client

function editClient(client) {

    var clientRange = [client.id, client.name, client.address, client.plz, client.city, client.country, client.comment, client.consultant, client.consultantID, client.telephone, client.email]
    var sheet = SpreadsheetApp.openById(sheetClient).getSheets()[0];

    searchStudent(client.id, "row", function (row) {
        var range = sheet.getRange(row + 2, 1, 1, 11);
        range.setValues([clientRange]);

    })
    return "Client successfully edited"
};


// Initiales first page
function doGet() {

    return HtmlService
        .createTemplateFromFile('index')

        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("CRM by dm").addMetaTag('viewport', 'width=device-width, initial-scale=1');
};

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
};

// not needed, but for future use
function loadFile(file) {
    return HtmlService.createHtmlOutputFromFile(file).getContent()
};