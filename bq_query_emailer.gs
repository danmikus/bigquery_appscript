//Runs a query and emails the results to recipients
//IMPORTANT: make sure to enable the bq library in Apps Script and turn the API on in the GCP project before starting.
//For reference I used 'bq' for the library name instead of the default 'bigquery'

function calculateFileName_() {
  //Calculates the files name to be used for file generation
  var fileDate = new Date()
  return Utilities.formatDate(fileDate, "EST", "YYYYMMdd") + "_report.csv";
}

function runQuery(projectId, query) {
  //Runs the query and returns the query results
  
  var queryRequest = bq.newQueryRequest();
  queryRequest.query = query
  queryRequest.useLegacySql = false
  
  var queryResults = bq.Jobs.query(queryRequest,projectId);
  
  var resultCount = queryResults.totalRows;
  var resultSchema = queryResults.schema.fields;
  var resultValues = new Array(resultCount);
  var fieldNames = new Array();
  var tableRows = queryResults.rows;
 
  for (var i = 0; i < resultSchema.length; i++) {
    fieldNames[i] = resultSchema[i].name
  };
  
  for (var i = 0; i < tableRows.length; i++) {
    var cols = tableRows[i].f;
    resultValues[i] = new Array(cols.length);
    // For each column, add values to the result array
    for (var j = 0; j < cols.length; j++) {
      resultValues[i][j] = cols[j].v;
    }
  }
  
  resultValues.splice(0, 0, fieldNames)
  
  return resultValues
}

function convertRangeToCsv(csvFileName,data) {
  try {

    var csvData = undefined;

    // Loop through the data in the range and build a string with the CSV data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvData = csv;
    }
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
  
  var file = Utilities.newBlob(csvData, "text/csv", csvFileName)
  return file
}

function sendEmail(recipient, subject, message, attachment){
  
  MailApp.sendEmail(recipient, subject, message, {
     name: 'Automatated Reporting Emailer',
     attachments: [attachment]
 })
}

function main() {
  
  var projectId = XXX //insert your bigquery project id here
  
  //enter your query below (example template below)
  var query = "SELECT \
    XXX \
  FROM \
    `XXX \
  JOIN \
    XXX \
  ON \
    XXX \
  WHERE \
    XXX \
    AND \
    XXX \
  GROUP BY \
    XXX \
  ORDER BY \
    XXX"
    ;
  
  var recipient = "XXX, XXX, XXX"; //enter the email recipients here
  var subject = "XXX - " + calculateFileName_(); //enter the email subject here
  var message = "Hello everyone,\n\nPlease see the attachment for this week's file. Thanks!\n\nRegards,\nDan\n\n"; //enter the email subjec here
  
  var queryResults = runQuery(projectId, query)
  var csvFileName = calculateFileName_()
  var createdFile = convertRangeToCsv(csvFileName, queryResults)
  
  sendEmail(recipient, subject, message, createdFile)

}
