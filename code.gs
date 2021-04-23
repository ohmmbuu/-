/* โค้ด.gs 
ระบบรับสมัครนักเรียน พัฒนาโดย นายจิรศักดิ์ จิรสาโรช E-mail: niddeaw.n@gmail.com Tel : 0806393969
เครดิตและอ่านรายละเอียด : https://github.com/jamiewilson/form-to-google-sheets
เครดิตต้นฉบับ original from: http://mashe.hawksey.info/2014/07/google-sheets-as-a-database-insert-with-apps-script-using-postget-methods-with-ajax-example/

อัพเดทโค้ด 23 เมษายน 2564 เพิ่มระบบสร้างไฟล์ PDF ใบสมัคร , ส่ง อีเมล , แจ้งเตือนทางไลน์กลุ่ม , อัพโหลดรูปภาพ เครดิต ครูเก๋ 

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
  	const folderId = "ID_โฟลเดอร์รูปภาพ";  // ID โฟลเดอร์รูปภาพ

	const blob = Utilities.newBlob(JSON.parse(e.postData.contents), e.parameter.mimeType, e.parameter.filename);
	const file = DriveApp.getFolderById(folderId).createFile(blob);
	const responseObj = {filename: file.getName(), fileId: file.getId(), fileUrl: file.getUrl()};

	var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1')
	var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2')

    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
    var sheet = doc.getSheetByName('Sheet1') // ระบุชื่อชีต

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    var nextRow = sheet.getLastRow() + 1

    var newRow = headers.map(function(header) {
      return header === 'timestamp' ? new Date() : e.parameter[header]
    })

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])

	sheet.getRange(ss.getLastRow(),41).setValue(responseObj['fileId']) // เพิ่ม URL ของไฟล์ภาพที่ตำแหน่งแถวล่าสุด, คอลัมภ์ที่ **แก้ไข
	var getim_id = sheet.getRange(ss.getLastRow(),41).getDisplayValue() // แสดงค่า ID ภาพ คอลัมภ์ที่ **แก้ไข
	var IMAGE_URL_1 = 'https://doc.google.com/uc?export=view&id='+ getim_id;

	sheet.getRange(ss.getLastRow(),42).setValue(IMAGE_URL_1) // คอลัมภ์ที่ **แก้ไข
	
/* -------------------------------------------------------------------------------------------------------------------------------*/
/* ระบบสร้างไฟล์ PDF ใบสมัคร , ส่ง อีเมล , แจ้งเตือนทางไลน์กลุ่ม
/* เครดิต ครูสมพงษ์ โพคาศรี E-mail: Spkorat0125@gmail.com Tel : 0956659190 
/* Line : guytrue fb: https://www.facebook.com/spkorat0125 */

// สร้าง pdf กำหนดไฟล์แม่แบบและโฟลเดอร์ที่ใช้งาน --------------------------------------------------------------------------------
    var SlideFile = "ID_สไลด์ไฟล์แม่แบบ"; // ID สไลด์ไฟล์แม่แบบ
    const tempFolder = DriveApp.getFolderById("ID_โฟลเดอร์_temp"); // ID โฟลเดอร์ temp
    const pdfFolder = DriveApp.getFolderById("ID_โฟลเดอร์_PDF"); // ID โฟลเดอร์ PDF
            
// ส่วนสำหรับสร้างสำเนาไฟล์ต้นฉบับ ---------------------------------------------------------------------------------------------
    var strYear = parseInt(Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy")) + 543;
    var strMonth = Utilities.formatDate(new Date(), "Asia/Bangkok", "M");
    var strDay = Utilities.formatDate(new Date(), "Asia/Bangkok", "d");
    var strhour=Utilities.formatDate(new Date(), "Asia/Bangkok", "HH");
    var strMinute=Utilities.formatDate(new Date(), "Asia/Bangkok", "mm");
    var strMonthCut = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    var strMonthThai = strMonthCut[strMonth];  
    var DatetimeFile=strDay+' '+strMonthThai+' '+strYear+ ' เวลา '+strhour+'.'+strMinute;

    var Slide_TempFile_Copy = DriveApp.getFileById(SlideFile);              
    var Slide_File_CopyStud = Slide_TempFile_Copy.makeCopy('ม.1 '+newRow[3]+newRow[4]+" "+newRow[5]+" "+DatetimeFile,tempFolder); 
    var SlideCopyId = Slide_File_CopyStud.getId();
    var SlideNewCopy = SlidesApp.openById(SlideCopyId);
    var slides = SlideNewCopy.getSlides();
    var TemplateSlide = slides[0]; 
    var shapes = TemplateSlide.getShapes();
	
	TemplateSlide.insertImage(IMAGE_URL_1,195,10,50,40).getBorder().setWeight(1) // แทรกรูปภาพ กำหนดตำแหน่งและขนาดภาพ insertImage(imageUrl, left, top, width, height)
	
// ส่วนของการผนวกข้อมูลกับเอกสาร (แทนที่ข้อความด้วยข้อมูล) ------------------------------------------------------------------   
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
    shape.getText().replaceAllText('{zipcode1}',newRow[25]);
    shape.getText().replaceAllText('{gpa}',newRow[26]);
    shape.getText().replaceAllText('{school_type}',newRow[27]);
    shape.getText().replaceAllText('{disability}',newRow[28]);
    shape.getText().replaceAllText('{father}',newRow[29]);
    shape.getText().replaceAllText('{father_occupation}',newRow[30]);
    shape.getText().replaceAllText('{father_phone}',newRow[31]);
    shape.getText().replaceAllText('{mother}',newRow[32]);
    shape.getText().replaceAllText('{mother_occupation}',newRow[33]);
    shape.getText().replaceAllText('{mother_phone}',newRow[34]);
    shape.getText().replaceAllText('{parent}',newRow[35]);
    shape.getText().replaceAllText('{parent_occupation}',newRow[36]);
    shape.getText().replaceAllText('{parent_phone}',newRow[37]);
    shape.getText().replaceAllText('{relationship}',newRow[38]);
});

    var pdfName ="ม.1 " + newRow[3]+newRow[4]+" "+newRow[5]+" "+DatetimeFile
    SlideNewCopy.saveAndClose();
    
// สร้างไฟล์ pdf ---------------------------------------------------------------------------------------------------------------
    const pdfContentBlob = Slide_File_CopyStud.getAs(MimeType.PDF); 
    var newPDFFile=pdfFolder.createFile(pdfContentBlob).setName(pdfName+".pdf"); 
    //tempFolder.removeFile(Slide_File_CopyStud); // ลบไฟล์สำเนาสไลด์ หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก
    
// ส่วนการส่งอีเมล์ -------------------------------------------------------------------------------------------------------------
    var email = "xxx@gmail.com"; //ส่งเมลไปที่เจ้าหน้าที่
    MailApp.sendEmail(email, "สมัครเรียนออนไลน์", "จาก โรงเรียนวัดไร่ขิงวิทยา ท่านได้ทำการลงทะเบียนเรียนด้วยระบบออนไลน์ กรุณาตรวจสอบข้อมูล", {attachments: [newPDFFile],});
    
// ลบไฟล์สำเนาออก -----------------------------------------------------------------------------------------------------------
    // Slide_TempFile_Copy.setTrashed(true); // ไฟล์ google slide สำเนาต้นฉบับ หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก
    // newPDFFile.setTrashed(true); // ไฟล์ PDF หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก
    // Slide_File_CopyStud.setTrashed(true); // ไฟล์ google slide สำเนาต้นฉบับที่ถูกแทนที่ด้วยข้อความใหม่ หากต้องการลบไฟล์ให้ลบเครื่องหมาย // ด้านหน้าออก

// กำหนดตัวแปรให้กับข้อความที่จะส่งไลน์แจ้งเตือน -------------------------------------------------------------------------------
	var xd = newPDFFile.getUrl()
	var nationid = newRow[7]
	var pnone = newRow[36]
	var re_xx = nationid.slice(8, 13);
	var re_phone = pnone.slice(5,10)
	var id_doc = "ps-"+re_xx+"-"+re_phone
	addlink(xd,id_doc)
	var sht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2")
	var shot_url = sht.getRange("b1").getValue()
	var text_data = '📣 สมัครเรียนระดับชั้น ม.1\n';
    text_data += 'วันที่ '+DatetimeFile+" น."+'\nชื่อ-นามสกุล : '+newRow[3]+newRow[4]+" "+newRow[5];
    sendLineNotify(text_data);
/* -----------------------------------------------------------------------------------------------------------------------------*/
 
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

// ส่วนฟังก์ชั่นแจ้งเตือนไลน์ -------------------------------------------------------------------------------------------------------
function sendLineNotify(message) {

    var token = ["xxx"]; // ใส่ access token Line
    var options = {
        "method": "post",
        "payload": "message=" + message,
        "headers": {
            "Authorization": "Bearer " + token
        }
    };

    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

//============================ ส่วนเพิ่มเติม การเพิมลิงค์และเพิ่มเมนู By gukkghu ========================================
var Route ={};
    Route.path = function(route,callback){
    Route[route] = callback;
    
    }

function  addlink(xd,id_doc){
	 var ws = SpreadsheetApp.getActiveSpreadsheet()
	 var sheet1 = ws.getSheetByName("Sheet1")
	 var sheet2 = ws.getSheetByName("Sheet2")
	 var lr = sheet1.getLastRow()
		Logger.log(lr)
	sheet1.getRange(lr,40).setValue(xd) // ID ลิงค์ภาพ

	sheet2.getRange("a1").setValue(xd) // ลิงค์ PDF
}

function getSheetData()  { 
	var ss= SpreadsheetApp.getActiveSpreadsheet();
	var dataSheet = ss.getSheetByName('Sheet1'); 
	var dataRange = dataSheet.getDataRange();
	var dataValues = dataRange.getDisplayValues();  
 
   return dataValues;

}
function getSheetDatas2()  { 
	var ss= SpreadsheetApp.getActiveSpreadsheet();
	var dataSheet = ss.getSheetByName('Sheet2'); 
	var dataRange = dataSheet.getDataRange();
	var dataValues1 = dataRange.getDisplayValues(); 
 }
 
function doGet(e) {
  Route.path("index1",loadForm);
  // Route.path("index",loadForm2);

   if(Route[e.parameters.v]) {
   return Route[e.parameters.v]();
   }else {
   return render("menu");
   }
  }

//ส่วนย่อยเรียกการทำงานหน้า page
function loadForm(){
	return render("index1");
}

function getUrl(){
	var url =ScriptApp.getService().getUrl();
	return url;
	Logger.log(url)
}

function include(filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function render(file, argsObject){
var tmp = HtmlService.createTemplateFromFile(file);
	if(argsObject) {
var keys = Object.keys (argsObject);
	keys.forEach(function(key){
	tmp[key] = argsObject[key]; 
   }); 
 } 
return tmp.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode. ALLOWALL);
}