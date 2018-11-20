
// FOR MONTHLY CONSOLIDATION
function setLastRowMonthly() {

 }




function mailLeaveDetails(){
  var attendanceSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  
  var controllerSheet = SpreadsheetApp.openById("1O9LE0bNoQBnWW33rCzZG6l_EcReY_SQMvZTdRuNQPUg");
  var controlSheet = controllerSheet.getSheetByName("controller");
  var emailIds = controlSheet.getRange('B3').getValue();
  var isSendMail = controlSheet.getRange('B4').getValue();
  var subjects = controlSheet.getRange('B5').getValue();
  var sendersEmail = controlSheet.getRange('B89').getValue();
  var rowCount = attendanceSheet.getLastRow();
  var startRow = attendanceSheet.getRange("H1").getValue();
  rowCount = rowCount -2;
  Logger.log(" total row : "+rowCount);
  Logger.log(" start row : "+startRow);
  var todaysDate = new Date();
  var htmlContent = '<p> <b> <center> <H2> K.S.C.C.F Ltd </H2> </center> </b> </p> <p><center> <b> KOZHIKODE REGION  </b></center> </p> <p><center> <b> EMPLOYEES LEAVE ON '+todaysDate+' </b> </center> </p>'
  var htmlTable = '<table border="2"cellspacing="0" cellpadding="4" style="width:100%">';
  htmlTable = htmlTable + '<tr><td><b>SL NO</b></td><td><b> TIME </b></td><td><b> BRANCH NAME </b></td><td><b> EMPLOYEE NAME </b></td><td><b> ATTENDANCE </b></td></tr>';
  var slNo = 1;
  for (var i = rowCount; i > startRow; i--) {
     var row = "";
     row = '<tr><td>'+slNo+'</td><td>'+attendanceSheet.getRange('A'+i).getValue()+'</td><td>'+attendanceSheet.getRange('B'+i).getValue()+'</td><td>'+attendanceSheet.getRange('C'+i).getValue()+'</td><td>'+attendanceSheet.getRange('D'+i).getValue()+'</td></tr>';
     htmlTable = htmlTable + row;
     slNo = slNo + 1;
     //Logger.log(attendanceSheet.getRange('B'+rowCount).getValue());
     //Logger.log(row);
   }
   htmlTable = htmlTable + '</table>';
      var footer = '<br><br><p align="right"><b> <a href="https://goo.gl/5aTn41"> IT SECTION </a> </b></p>';
      var message = htmlContent + htmlTable + footer;
      var subject =  subjects;
      var optAdvancedArgs = {name: sendersEmail , bcc : "nijesh.zgc@gmail.com", htmlBody: message};
   if(isSendMail =='YES') {
     MailApp.sendEmail(emailIds, subject , "Email Body" , optAdvancedArgs);
     }
   Logger.log(htmlTable);
   Logger.log("mailLeaveDetails script run succesfull");
}

// HELPS FOR BACKUP
function setMonthlyLastRow() {
  var attendanceSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  var attnSheet = attendanceSheet.getSheetByName("ATTENDANCE")
  var rowCounts = attnSheet.getLastRow();
  attnSheet.getRange("F9").setValue(rowCounts-1);
  
    
  var dataSheet = attendanceSheet.getSheetByName("DATA");
  var rowCountDataSheet = dataSheet.getLastRow();
  var rowCount = rowCounts-2;
  var monthList = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  var todayDate = new Date();
  var thisYear = todayDate.getYear();
  var month = todayDate.getMonth();
  var day = todayDate.getDate();
  
  var row  = rowCountDataSheet+1;
  var actualMonth = month +1;
  dataSheet.getRange("A"+row).setValue(thisYear);
  dataSheet.getRange("B"+row).setValue(monthList[month-1]);
  dataSheet.getRange("C"+row).setValue(thisYear+"-"+month);
  dataSheet.getRange("D"+row).setValue(rowCount);
  
  Logger.log("The last row : "+rowCount);
  Logger.log("setLastRow script run succesfull");
  
}
 
 //for sinija  at regional office - currently not in use
function mailMonthlyLeaveDetails(){
  var attendanceSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  var rowCount = attendanceSheet.getLastRow();
  var startRow = attendanceSheet.getRange("F9").getValue();
  Logger.log(" total row : "+rowCount);
  Logger.log(" start row : "+startRow);
  var todaysDate = new Date();
  var htmlContent = '<p> <b> <center> <H2> K.S.C.C.F Ltd </H2> </center> </b> </p> <p><center> <b> KOZHIKODE REGION  </b></center> </p> <p><center> <b> EMPLOYEES LEAVE ON '+todaysDate+' </b> </center> </p>'
  var htmlTable = '<table border="2"cellspacing="0" cellpadding="4" style="width:100%">';
  htmlTable = htmlTable + '<tr><td><b>SL NO</b></td><td><b> BRANCH NAME </b></td><td><b> EMPLOYEE NAME </b></td><td><b> TYPE OF LEAVE </b></td></tr>';
  var slNo = 1;
  for (var i = rowCount; i > startRow; i--) {
     var row = "";
     row = '<tr><td>'+slNo+'</td><td>'+attendanceSheet.getRange('B'+i).getValue()+'</td><td>'+attendanceSheet.getRange('C'+i).getValue()+'</td><td>'+attendanceSheet.getRange('D'+i).getValue()+'</td></tr>';
     htmlTable = htmlTable + row;
     slNo = slNo + 1;
     //Logger.log(attendanceSheet.getRange('B'+rowCount).getValue());
     //Logger.log(row);
   }
   htmlTable = htmlTable + '</table>';
   var footer = '<br><br><p align="right"><b> ADMINISTRATION SECTION  </b></p>';
   var message = htmlContent + htmlTable + footer;
   var subject =  ' EMPLOYEES LEAVE IN THIS MONTH ';
  var optAdvancedArgs = {name: " CFED KZD REPORTS ", bcc : "nijesh.zgc@gmail.com", htmlBody: message};
   if(rowCount > startRow) {
     MailApp.sendEmail("bithesh@gmail.com", subject , "Email Body" , optAdvancedArgs);
     }
   Logger.log(htmlTable);
   Logger.log("mailLeaveDetails script run succesfull");
}


function changeSheetDaily(){
  var alphaList = ["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
  var attendanceSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  var sheetAttn = attendanceSheet.getSheetByName("ATTENDANCE");
  var startRow = 2;
  var rowCount = sheetAttn.getLastRow();
  var rowLimit = rowCount -  3;
  Logger.log(rowLimit);
  attendanceSheet.getSheetByName("ATTENDANCE").hideRows(startRow,rowLimit);
  var lastColumn = sheetAttn.getLastColumn();
  Logger.log(" L : "+lastColumn)
  var columnToHide = sheetAttn.getRange("G:"+alphaList[lastColumn]);
  Logger.log(alphaList[lastColumn]);
  sheetAttn.hideColumn(columnToHide);
  sheetAttn.unhideColumn(sheetAttn.getRange("A:F"));
  sheetAttn.unhideRow(sheetAttn.getRange("1:2"));
  sheetAttn.unhideRow(sheetAttn.getRange((rowCount-1)+":"+rowCount));
  var todayDate = new Date();
  var thisYear = todayDate.getYear();
  var monthList = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  var weekList = ["SUN","MON","TUE","WED","THU","FRI","SAT"];
  var weekDay = todayDate.getDay();
  var actualMonth = todayDate.getMonth();
  var day = todayDate.getDate();
  attendanceSheet.rename("ATTENDANCE_ON_"+thisYear+"_"+monthList[actualMonth]+"_"+day);
  sheetAttn.getRange("B1").setValue("Date : "+thisYear+" "+monthList[actualMonth]+" "+day+" "+weekList[weekDay]);
  sheetAttn.getRange("F3").setValue(rowCount-2);
  Logger.log("clear sheet Daily script run succesfull");
  var consolidationSheet = attendanceSheet.getSheetByName("CONSOLIDATION");
  consolidationSheet.hideSheet();
}



function sendEmailwithPdf() {
      var sheetToPdf = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
      var controllerSheet = SpreadsheetApp.openById("1O9LE0bNoQBnWW33rCzZG6l_EcReY_SQMvZTdRuNQPUg");
      var controlSheet = controllerSheet.getSheetByName("controller");
      var consolidationSheet = sheetToPdf.getSheetByName("CONSOLIDATION");
      var emailIds = controlSheet.getRange('B1').getValue();
      var isSendMail = controlSheet.getRange('B2').getValue();
      var subject = controlSheet.getRange('B5').getValue();
      var senderEmail = controlSheet.getRange('B89').getValue();
      consolidationSheet.showSheet();
      var attendanceSheet = sheetToPdf.getSheetByName("ATTENDANCE");
      var startRow = attendanceSheet.getRange("F3").getValue();
      var attndCount = startRow + 1;
      attendanceSheet.getRange("F4").setValue('=COUNTIF(D'+attndCount+':D,"PRESENT")');
      attendanceSheet.getRange("F5").setValue('=COUNTIF(D'+attndCount+':D,"FULL DAY(L)")'); //ON DUTY
      attendanceSheet.getRange("F6").setValue('=SUM(COUNTIF(D'+attndCount+':D,"HALF DAY - MORNING(L)"),COUNTIF(D5353:D,"HALF DAY - AFTER NOON(L)"))');
      attendanceSheet.getRange("F7").setValue('=COUNTIF(D'+attndCount+':D,"ON DUTY")');
      attendanceSheet.getRange("F8").setValue(new Date());
      var totPresent = 0;
      var totAbsent = 0;
      var totHalfDays = 0;
      var totOnDuties = 0;
      totPresent = attendanceSheet.getRange("F4").getValue();
      totAbsent = attendanceSheet.getRange("F5").getValue();
      totHalfDays = attendanceSheet.getRange("F6").getValue();
      totOnDuties = attendanceSheet.getRange("F7").getValue();
      var totMarked = totPresent + totAbsent + totHalfDays + totOnDuties;
      
      var todayDate = new Date();
     var htmlBody = '<br/><p><br/> <font size="4" face="verdana">'+totPresent+'</font> presents <font size="4" face="verdana">'+totAbsent+'</font> absents  <font size="4" face="verdana">'+totHalfDays+'</font> half days <font size="4" face="verdana">'+totOnDuties+'</font> on duties in our region from <font size="5">'+totMarked+'</font> total attendance marked <br/><p> See the changes in the Google Document till '+todayDate+' <br/><br/> Open the current version of the Google Document : <a href="https://goo.gl/Gm8HuR"> Click here</a> <br/><br/> Powered by Google Sheets <br/> --- <br/> Want to stop receiving this email? <a href="https://goo.gl/I7GjOL"> Click here </a> </p> ';
     //var htmlBody = '<p> See the changes in the Google Document till '+todayDate+' <br/><br/> Open the current version of the Google Document : <a href="https://goo.gl/Gm8HuR"> Click here</a> <br/><br/> Powered by Google Sheets <br/> --- <br/> Want to stop receiving this email? <a href="https://goo.gl/I7GjOL"> Click here </a> </p> ';
     //var htmlBody = '';
     var subject =  subject;
      var optAdvancedArgs = {name:senderEmail ,bcc :"nijesh.zgc@gmail.com", htmlBody: htmlBody, attachments: sheetToPdf.getAs(MimeType.PDF) };
      var rowCount = sheetToPdf.getLastRow();
      
      if(isSendMail =='YES') {
        MailApp.sendEmail(emailIds, subject , "Email Body" , optAdvancedArgs);
      }
      //consolidationSheet.hideSheet();
}


function backupMonthly(){
  var months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  var todayDate = new Date();
  var year = todayDate.getYear();
  var sheetToPdf = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  var controllerSheet = SpreadsheetApp.openById("1O9LE0bNoQBnWW33rCzZG6l_EcReY_SQMvZTdRuNQPUg");
  var controlSheet = controllerSheet.getSheetByName("controller");
  var attnSheet = controllerSheet.getSheetByName("ATTENDANCE");
  var startRow = attnSheet.getRange("F9").getValue();
  var rowCount = attnSheet.getLastRow();
  var diff = rowCount - startRow;
  Logger.log(diff);
  sheetToPdf.getSheetByName("ATTENDANCE").showRows(startRow , diff-4);
  var actualMonth = todayDate.getMonth();
  var previousMonth = actualMonth - 1;
  attnSheet.getRange("B1").setValue("Month : "+year+" "+months[previousMonth]);
    if(actualMonth==0){
      previousMonth = 11;
    }
     var destFolder = DriveApp.getFolderById("0B0mtudSQYtQpeDBzNTF3czljLU0");
     DriveApp.getFileById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo").makeCopy("BK_RO_ATTENDANCE_"+year+"_"+months[previousMonth], destFolder);
 // sheetToPdf.getRange("I1").setValue(rowCount-2);
     var senderEmail = controlSheet.getRange('B89').getValue();
     var subject = controlSheet.getRange('B5').getValue();
     var htmlBody = '<br/> <p> Backup created on '+todayDate+'<br/><br/> Backup location : <a href="https://goo.gl/m4JAP8"> Location </a> <br/><br/> File Name  : <a href="https://goo.gl/m4JAP8"> BK_RO_ATTENDANCE_'+year+"_"+months[previousMonth]+' </a><br/> To modify details <a href="https://goo.gl/I7GjOL"> Click here </a> <br/><br/> Open the current version of the Google Document : <a href="https://goo.gl/Gm8HuR"> Spreadsheet </a> <br/><br/> Powered by Google Sheets <br/> --- <br/> Want to stop receiving this email? <a href="https://goo.gl/I7GjOL"> Click here </a> </p> ';
     var subject =  subject;
     var optAdvancedArgs = {name: senderEmail ,bcc :"nijesh.zgc@gmail.com", htmlBody: htmlBody, attachments: sheetToPdf.getAs(MimeType.PDF) };
     MailApp.sendEmail(senderEmail, subject , "Email Body" , optAdvancedArgs);
      
  }


 function test(){
 
 var attendanceSpreadSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
 var attendanceSheet = attendanceSpreadSheet.getSheetByName("ATTENDANCE");
  attendanceSheet.getRange("L2").setValue('=COUNTIF(D5354:D,"PRESENT")');
  attendanceSheet.getRange("M2").setValue('=COUNTIF(D5354:D,"FULL DAY(L)")'); //ON DUTY
  attendanceSheet.getRange("N2").setValue('=SUM(COUNTIF(D5353:D,"HALF DAY - MORNING(L)"),COUNTIF(D5353:D,"HALF DAY - AFTER NOON(L)"))');
  attendanceSheet.getRange("O2").setValue('=COUNTIF(D5354:D,"ON DUTY")');
  attendanceSheet.getRange("P2").setValue(new Date());
 //Var a = ArrayFormula(CountIF(A:A="PRESENT"));
 //Logger.log(a);
 
     //var i = 0;
    // var a = "hello";
     //var b = "hell";
     //var c = "a";
    // var d = "aat";
     //var n = a.localeCompare(b);
     // Logger.log(c.length);
    // b="";
    // i = b.length;
   //  Logger.log(i);
  //var attendanceSpreadSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  
  //var attendanceSheet = attendanceSpreadSheet.getSheetByName("ATTENDANCE");
  //var consolidationSheet = attendanceSpreadSheet.getSheetByName("CONSOLIDATION");
  //consolidationSheet.showSheet();
 
 } 

//created on 2/5/2017 (time 11:32 pm)  RUNNING TIME DAILY NOON
function dailyConsolidation(){
  var attendanceSpreadSheet = SpreadsheetApp.openById("11GyynG3eJcscLpMpsjvA4yyPKCMe6bBKR3ctwmvRqEo");
  
  var attendanceSheet = attendanceSpreadSheet.getSheetByName("ATTENDANCE");
  var consolidationSheet = attendanceSpreadSheet.getSheetByName("CONSOLIDATION");
  
  var startRowAttendance = consolidationSheet.getRange("H1").getValue();
  var lastRowAttendance = attendanceSheet.getLastRow();
  var lastRowOfAttendanceArray = lastRowAttendance - startRowAttendance - 1;
  
  
  
 
  
  var branchListAS = attendanceSheet.getRange(startRowAttendance,2,lastRowOfAttendanceArray,1).getValues();
  var empListAS = attendanceSheet.getRange(startRowAttendance,3,lastRowOfAttendanceArray,1).getValues();
  var attndListAS = attendanceSheet.getRange(startRowAttendance,4,lastRowOfAttendanceArray,1).getValues();
  
  var currentAttndRowPstn = startRowAttendance;
  
  for(var rowCount = 0; rowCount < branchListAS.length; rowCount++){
  
          var branchListCS = consolidationSheet.getRange(1,1,consolidationSheet.getLastRow(),1).getValues();
          var empListCS = consolidationSheet.getRange(1,2,consolidationSheet.getLastRow(),1).getValues();
          
          var lastRowCS = consolidationSheet.getLastRow();
          var lastConsolidationRow = lastRowCS + 1;
          
          var branchNameAS = branchListAS[rowCount].toString();
          var employeeNameAS = empListAS[rowCount].toString();
          var attendanceTypeAS = attndListAS[rowCount].toString();  // present,onduty etc
          var attendanceType = attendanceTypeAS;
          
          currentAttndRowPstn = currentAttndRowPstn + 1;
          consolidationSheet.getRange("H1").setValue(currentAttndRowPstn);
          
          var isEntryMatched = 0;
          var row = 0;
          var input = 1;
          
               for(var rowCS = 0; rowCS < branchListCS.length; rowCS++){
          
                
                      var branchNameCS = branchListCS[rowCS].toString();
                      var employeeNameCS = empListCS[rowCS].toString();
                
                      var as = branchNameAS.concat(employeeNameAS);
                      var cs = branchNameCS.concat(employeeNameCS);
                
                      var value = as.localeCompare(cs);
                      
                        if(value == 0){
                          row = rowCS + 1;
                          isEntryMatched = 1;
                          attendanceType = "OTHERS";
                          break;
                         } // if value
                      
                }// for rowCS
                
              if(isEntryMatched > 0){
                  //var choose = attendanceType;
                  switch(attendanceTypeAS){
                         case "PRESENT":
                               value = consolidationSheet.getRange("C"+row).getValue();
                               consolidationSheet.getRange("C"+row).setValue(value+1);
                               consolidationSheet.getRange("I"+row).setValue(new Date());
                               
                              // Logger.log('----'+value);
                               break;
                        case "FULL DAY(L)":
                               value = consolidationSheet.getRange("D"+row).getValue();
                               consolidationSheet.getRange("D"+row).setValue(value+1);
                               consolidationSheet.getRange("I"+row).setValue(new Date());
                               
                               //Logger.log('----'+value);
                               break;
                        case "HALF DAY - MORNING(L)":
                               value = consolidationSheet.getRange("E"+row).getValue();
                               consolidationSheet.getRange("E"+row).setValue(value+1);
                               consolidationSheet.getRange("I"+row).setValue(new Date());
                               
                               //Logger.log('----'+value);
                                break;
                         case "HALF DAY - AFTER NOON(L)":
                               value = consolidationSheet.getRange("F"+row).getValue();
                               consolidationSheet.getRange("F"+row).setValue(value+1);
                               consolidationSheet.getRange("I"+row).setValue(new Date());
                               
                               //Logger.log('----'+value);
                                break;
                        case "ON DUTY":
                               value = consolidationSheet.getRange("G"+row).getValue();
                               consolidationSheet.getRange("G"+row).setValue(value+1);
                               consolidationSheet.getRange("I"+row).setValue(new Date());
                               
                               //Logger.log('----'+value);
                              break;
                        default:      
                             }
                        
                        //isInserted = 0;
                        
                        attendanceType = "OTHERS";
                        input = 0;
                        branchNameAS = "";
                        employeeNameAS = "";
              }else{
              
               var others = "OTHERS";
               var n = attendanceType.localeCompare(others);
               Logger.log("s :"+n+attendanceType);
               
                if(n!=0) {
                        
                        //var choose = attendanceType;
                        consolidationSheet.getRange("A"+lastConsolidationRow).setValue(branchNameAS);
                        consolidationSheet.getRange("B"+lastConsolidationRow).setValue(employeeNameAS);
                        
                        //Logger.log(" Before switch : "+choose);
                        
                        switch(attendanceType){
                         
                       
                        case "FULL DAY(L)":
                               consolidationSheet.getRange("D"+lastConsolidationRow).setValue(input);
                               consolidationSheet.getRange("C"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("E"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("F"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("G"+lastConsolidationRow).setValue(0);
                        break;
                        case "HALF DAY - MORNING(L)":
                               consolidationSheet.getRange("E"+lastConsolidationRow).setValue(input);
                               consolidationSheet.getRange("C"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("D"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("F"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("G"+lastConsolidationRow).setValue(0);
                        break;
                        case "HALF DAY - AFTER NOON(L)":
                               consolidationSheet.getRange("F"+lastConsolidationRow).setValue(input);
                               consolidationSheet.getRange("C"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("E"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("D"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("G"+lastConsolidationRow).setValue(0);
                        break;
                        case "ON DUTY":
                               consolidationSheet.getRange("G"+lastConsolidationRow).setValue(input);
                               consolidationSheet.getRange("C"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("E"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("F"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("D"+lastConsolidationRow).setValue(0);
                        break;
                        case "PRESENT":
                               consolidationSheet.getRange("C"+lastConsolidationRow).setValue(input);
                               consolidationSheet.getRange("D"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("E"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("F"+lastConsolidationRow).setValue(0);
                               consolidationSheet.getRange("G"+lastConsolidationRow).setValue(0);
                        break;
                        default :
                               Logger.log("test : "+attendanceType);
                               var values = [[ "0", "0", "0" , "0", "0"]];
                               consolidationSheet.getRange("C"+lastConsolidationRow+":G"+lastConsolidationRow).setValue(values);
                              
                        } //end switch
                        
                       }// if attendanceType >0 n!=0
                       
              if(n==0){
                  consolidationSheet.hideRows(lastConsolidationRow);
              } 
              
              }
              
                
  
  } //for rowCount

} //end consolidationNew


