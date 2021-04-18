/* โค้ด.gs 
เครดิตและอ่านรายละเอียด : https://github.com/jamiewilson/form-to-google-sheets
เครดิตต้นฉบับ original from: http://mashe.hawksey.info/2014/07/google-sheets-as-a-database-insert-with-apps-script-using-postget-methods-with-ajax-example/
อัพเดทโค้ด 18 เมษายน 2564 เพิ่มระบบสร้างไฟล์ PDF ใบสมัคร , ส่ง อีเมล , แจ้งเตือนทางไลน์กลุ่ม
*/

var sheetName = 'Sheet1'
var scriptProp = PropertiesService.getScriptProperties()

function intialSetup () {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  scriptProp.setProperty('key', activeSpreadsheet.getId())
}

function doPost (e) {
  var lock = LockService.getScriptLock()
  lock.tryLock(10000)

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
    var sheet = doc.getSheetByName(sheetName)

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    var nextRow = sheet.getLastRow() + 1

    var newRow = headers.map(function(header) {
      return header === 'timestamp' ? new Date() : e.parameter[header]
    })

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])

/* ----------------------------------------------------------------------------------------------------------------------------------------*/
/* สร้าง pdf เครดิต ครูสมพงษ์ โพคาศรี email: Spkorat0125@gmail.com Tel : 0956659190 Line : guytrue fb: https://www.facebook.com/spkorat0125 */

//============สร้าง pdf==========================
    var strYear = parseInt(Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy")) + 543;
    var strMonth = Utilities.formatDate(new Date(), "Asia/Bangkok", "M");
    var strDay = Utilities.formatDate(new Date(), "Asia/Bangkok", "d");
    var strhour=Utilities.formatDate(new Date(), "Asia/Bangkok", "HH");
    var strMinute=Utilities.formatDate(new Date(), "Asia/Bangkok", "mm");
    
    var strMonthCut = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    
    var strMonthThai = strMonthCut[strMonth];  
    //var DatetimeLine=strDay+' '+strMonthThai+' '+strYear+ ' เวลา '+strhour+':'+strMinute+' น.';
    var DatetimeFile=strDay+' '+strMonthThai+' '+strYear+ ' เวลา '+strhour+'.'+strMinute;
    
    
    var SlideFile = "ID สไลด์ไฟล์แม่แบบ"; // ID สไลด์ไฟล์แม่แบบ
    const tempFolder = DriveApp.getFolderById("ID โฟลเดอร์ temp"); // ID โฟลเดอร์ temp
    const pdfFolder = DriveApp.getFolderById("ID โฟลเดอร์ PDF"); // ID โฟลเดอร์ PDF
            
            
//==================ส่วนสำหรับสร้างสำเนาไฟล์ต้นฉบับ=======================
  //var Slide_TempFile_Copy = DriveApp.getFileById(SlideFile).makeCopy(tempFolder);
    var Slide_TempFile_Copy = DriveApp.getFileById(SlideFile);              
    var Slide_File_CopyStud = Slide_TempFile_Copy.makeCopy('สมัครเรียน ม.1 '+newRow[3]+newRow[4]+" "+newRow[5]+" "+DatetimeFile,tempFolder); 
    var SlideCopyId = Slide_File_CopyStud.getId();
    var SlideNewCopy = SlidesApp.openById(SlideCopyId);
    var slides = SlideNewCopy.getSlides();
    var TemplateSlide = slides[0]; 
    var shapes = TemplateSlide.getShapes();
           
//=========================ส่วนของการผนวกข้อมูลกับเอกสาร========================================   
    //var Image_URL1 = 'https://doc.google.com/uc?export=view&id='+ ID_image;
    //var Image_URL2 = 'https://doc.google.com/uc?export=view&id='+ ID_sign;    
    //TemplateSlide.insertImage(Image_URL1, 196, 13, 30, 40).bringToFront().getBorder().setWeight(1); // Left  , top ,width , height + Border
    //TemplateSlide.insertImage(Image_URL2, 164, 182, 37, 26).bringToFront(); // Left  , top ,width , height
    //var strMonthFull = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];                
    //var strBHDYear = parseInt(Utilities.formatDate(newRow[6], "Asia/Bangkok", "yyyy")) + 543;              
    //var strBHDMonth = Utilities.formatDate(newRow[6], "Asia/Bangkok", "M");
    //var strBHDDay = Utilities.formatDate(newRow[6], "Asia/Bangkok", "d");
    //var MonthId = strMonthFull[strBHDMonth];    
 
    shapes.forEach(function (shape) {
    shape.getText().replaceAllText('{service}',newRow[1]);
    shape.getText().replaceAllText('{reg_type}',newRow[2]);
    shape.getText().replaceAllText('{prefix}',newRow[3]);
    shape.getText().replaceAllText('{name}',newRow[4]);
    shape.getText().replaceAllText('{lastname}',newRow[5]);
    shape.getText().replaceAllText('{birthday}',newRow[6]);
    shape.getText().replaceAllText('{idcard}',newRow[7]);
    shape.getText().replaceAllText('{race}',newRow[8]);
    shape.getText().replaceAllText('{nationality}',newRow[9]);
    shape.getText().replaceAllText('{religion}',newRow[10]);
    shape.getText().replaceAllText('{house_no}',newRow[11]);
    shape.getText().replaceAllText('{village_no}',newRow[12]);
    shape.getText().replaceAllText('{village}',newRow[13]);
    shape.getText().replaceAllText('{road}',newRow[14]);
    shape.getText().replaceAllText('{alley}',newRow[15]);
    shape.getText().replaceAllText('{district}',newRow[16]);
    shape.getText().replaceAllText('{amphoe}',newRow[17]);
    shape.getText().replaceAllText('{province}',newRow[18]);
    shape.getText().replaceAllText('{zipcode}',newRow[19]);
    shape.getText().replaceAllText('{student_phone}',newRow[20]);
    shape.getText().replaceAllText('{school}',newRow[21]);
    shape.getText().replaceAllText('{district1}',newRow[22]);
    shape.getText().replaceAllText('{amphoe1}',newRow[23]);
    shape.getText().replaceAllText('{province1}',newRow[24]);
    shape.getText().replaceAllText('{gpa}',newRow[25]);
    shape.getText().replaceAllText('{school_type}',newRow[26]);
    shape.getText().replaceAllText('{disability}',newRow[27]);
    shape.getText().replaceAllText('{father}',newRow[28]);
    shape.getText().replaceAllText('{father_occupation}',newRow[29]);
    shape.getText().replaceAllText('{father_phone}',newRow[30]);
    shape.getText().replaceAllText('{mother}',newRow[31]);
    shape.getText().replaceAllText('{mother_occupation}',newRow[32]);
    shape.getText().replaceAllText('{mother_phone}',newRow[33]);
    shape.getText().replaceAllText('{parent}',newRow[34]);
    shape.getText().replaceAllText('{parent_occupation}',newRow[35]);
    shape.getText().replaceAllText('{parent_phone}',newRow[36]);
    shape.getText().replaceAllText('{relationship}',newRow[37]);
    });
    
    var text_data = '📣 นักเรียนสมัครเรียนออนไลน์ ระดับชั้น ม.1\n';
    text_data += 'วันที่ '+DatetimeFile+" น."+'\nชื่อ-นามสกุล : '+newRow[3]+newRow[4]+" "+newRow[5];
    sendLineNotify(text_data);
    
    var pdfName ="สมัครเรียน ม.1 " + newRow[3]+newRow[4]+" "+newRow[5]+" "+DatetimeFile
    
    SlideNewCopy.saveAndClose();
    
    // ======================สร้างไฟล์ pdf========================
    
    //var newPDFFile = DriveApp.createFile(Slide_File_CopyStud.getAs("application/pdf")); //ไฟล์ที่ผสานข้อมูลแล้ว
    //const pdfContentBlob = Slide_File_CopyStud.getAs(MimeType.PDF);
    const pdfContentBlob = Slide_File_CopyStud.getAs(MimeType.PDF); 
    var newPDFFile=pdfFolder.createFile(pdfContentBlob).setName(pdfName+".pdf");
    //tempFolder.removeFile(Slide_TempFile_Copy);
    
    //======================ส่วนการส่งอีเมล์=========================
    var email = "xxx@gmail.com"; //ส่งเมลไปที่เจ้าหน้าที่
    MailApp.sendEmail(email, "สมัครเรียนออนไลน์", "จาก โรงเรียนวัดไร่ขิงวิทยา ท่านได้ทำการลงทะเบียนเรียนด้วยระบบออนไลน์ กรุณาตรวจสอบข้อมูล", {attachments: [newPDFFile],});
    
    //=====================ลบไฟล์สำเนาออก=========================
    // Slide_TempFile_Copy.setTrashed(true); // ไฟล์ google slide สำเนาต้นฉบับ หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก
    // newPDFFile.setTrashed(true); // ไฟล์ PDF หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก
    // Slide_File_CopyStud.setTrashed(true); // ไฟล์ google slide สำเนาต้นฉบับที่ถูกแทนที่ด้วยข้อความใหม่ หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก
/* ----------------------------------------------------------------------------------------------------------------------------------------*/
 
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
      .setMimeType(ContentService.MimeType.JSON)
  }

  catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
      .setMimeType(ContentService.MimeType.JSON)
  }

  finally {
    lock.releaseLock()
  }
}

//==============ส่วนฟังก์ชั่นแจ้งเตือนไลน์====================
function sendLineNotify(message) {

    var token = ["กรอก Token ID"]; //ใส่ access token
    var options = {
        "method": "post",
        "payload": "message=" + message,
        "headers": {
            "Authorization": "Bearer " + token
        }
    };

    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}