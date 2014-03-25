/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Create Letters",
    functionName : "createLetters"
  }];
  sheet.addMenu("Script Center Menu", entries);
};

function getTodayInfo(){
  var now = new Date();
  var days = [
    'Domingo',
    'Lunes',
    'Martes',
    'Miercoles',
    'Jueves',
    'Viernes',
    'Sabado',
  ];
  
  var months = [
    'Enero',
    'Febrero',
    'Marzo',
    'Abril',
    'Mayo',
    'Junio',
    'Julio',
    'Agosto',
    'Septiembre',
    'Octubre',
    'Noviembre',
    'Diciembre',
  ];
  var dates = {};
  
  dates['day_month'] = Math.ceil(now.getDate());
  dates['day'] = days[ now.getDay()];
  dates['month'] = months[now.getMonth()];
  dates['year'] = now.getFullYear();
  
  return dates;
}

/*
*Creates different functions that makes the following actions.
*
* 1 -Selects a range of lines.
* 2- Create documents based on the basic info of the range.
*/


/************************************ Employees Related Functions *******************************************/


function selectEmployees(){
 /*
 * The purpose of this function is to select all the employees of the company and return an array
 * with all the basic information they have, that way info can be managed easily with "Atributes"
 * of the users.
 *
 */
 
 //Open Original document with the employees list 
 var employees = SpreadsheetApp.openById("0Ag-OGPzUmTv5dFBzcUgzb040aF9hSXhzTzN4VkRMU0E").getRangeByName('employees').getValues();
 
 //Create array for the array of employees.
  var employeesArray = [];
  
  
 //Construct an array with all the atributes of the original list selected, this is implemented through aloop that returns
 //an array to the executioner
  
  for (i = 1; i < employees.length; i++){
       
       //Create new array to insert the info of the employee i;
       var employeeArray = [];
       employeeArray['code'] = employees[i][0];
       employeeArray['sysName'] = employees[i][1];
       employeeArray['email'] = employees[i][2];
       employeeArray['fullName'] = employees[i][3];
       employeeArray['position'] = employees[i][4];
       employeeArray['id'] = employees[i][5];
       employeeArray['beginningDate'] = employees[i][6];
       employeeArray['insuranceAFP'] = employees[i][7];
       employeeArray['sex'] = employees[i][8];
       
       //append the array with the employee i info to the whole company array;
       employeesArray.push( employeeArray);
                                      
       }
  
       return employeesArray;
  
}

function getInfo(data, key){
  /*
   * The purpose of this function is to get the array of information belonging to an employee
   * based on the key and values.  
   */
    
    //Create the array of all the employees
    var employeesArray = selectEmployees();
    var answerArray = [];  
    
    for (i = 0; i < employeesArray.length; i++){
        
        //if the data I am looking for does match I select the first array I find.
        if (employeesArray[i][k] == data){
         
            answerArray.push(employeesArray[i][key]);
      
         }
  
      }
      
  if (answerArray != [] ){
      return answerArray;
  }else{
      return "";  
    
  }
 }

/************************************************ Human Resources Letters Functions ******************************/


  function createLetters(){
    
    var employees   = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('employees').getValues();
    var today = getTodayInfo(); 
    
        for (i =35; i< 40/*employees.length**/; i++){
              
              //Creates a new document, can be verified for the document to see if the document exist or not.
               
              var doc = DocumentApp.create(employees[i][3] + '- Bank Letter');
              
              // Append a Few  space paragraphs
              var s1 = doc.appendParagraph("");
              var s2 = doc.appendParagraph("");
              var s3 = doc.appendParagraph("");
              var s4 = doc.appendParagraph("");
          
              // Append Salutation
              var location = doc.appendParagraph("Santo Domingo, Republica Dominicana");
              
              //This can be to be based on the actual date. 
              var date = doc.appendParagraph(today.day + ', ' + today.day_month +' de ' + today.month+' del ' + today.year);
          
              //more space
              var s5 = doc.appendParagraph("");
              var salutation = doc.appendParagraph('A quien corresponda, Banco Popular Dominicano,')
              
              //More Space for the content
              var s6 = doc.appendParagraph("");
              var employeeName = employees[i][3];   
              var employeeId = employees[i][5];
              
              var content = 'Por esta vía le solicitamos la apertura de cuenta  de nómina a ';
              if(employees[i][8] == 'F'){
                content += 'la empleada ';
              }else{
                Logger.log(employees[i][8]);
                content += 'el empleado '
              }

                  content += employeeName + ' cédula ' + employeeId;
                  content += ' bajo la empresa Stance Data S.R.L. de RNC número 130877025 en la cual labora.';
                  
              var letterContent = doc.appendParagraph(content);
              
              var s7 = doc.appendParagraph(""); 
              var thanks = doc.appendParagraph('Gracias,');
              
              var managerName = doc.appendParagraph('Leonardo Jimenez');
              var managerTitle = doc.appendParagraph('Gerente General');
               
              }
    
  }
