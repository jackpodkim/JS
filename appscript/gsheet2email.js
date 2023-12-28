
function getConfirmMessage(
    timestamp, admin, requester, dept, position, work_start_date, leave_type, leave_start,
    leave_end, leave_message ){
var htmlOutput = HtmlService.createHtmlOutputFromFile('Confirm'); // Message is the name of the HTML file

var message = htmlOutput.getContent()
message = message.replace("%timestamp", timestamp);
message = message.replace("%admin", admin);
message = message.replace("%requester", requester);
message = message.replace("%requester", requester);
message = message.replace("%dept", dept);
message = message.replace("%position", position);
message = message.replace("%work_start_date", work_start_date);
message = message.replace("%leave_type", leave_type);
message = message.replace("%leave_start", leave_start);
message = message.replace("%leave_end", leave_end);
message = message.replace("%leave_message", leave_message);

return message;
}

function getData(){
ss = SpreadsheetApp.openById('SOME ID'); 
sheet = ss.getSheets()[0];
var data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
return data;  
}


// standard time needs to be parsed
function parse_time(time){
time.setHours( time.getHours() + 14 )
var r = time.toLocaleDateString();
var res = r.split('/')[2] + '/' + r.split('/')[0] + '/' + r.split('/')[1];
return res
}


function addToCalendar(data_dict){
// calendar setup	
var cals = CalendarApp.getAllOwnedCalendars()
for(i=0;i<cals.length;i++){
  if(cals[i].getName()=='calendar name here'){
    var gCal = cals[i]
    break
  }
}
// console.log(gCal.getName());
// console.log(data_dict);


// calendar input
var title = data_dict['name'] + ' [' + data_dict['leave type'] +', ' + data_dict['leave days'] + 'days]'
var startDate = new Date(parse_time(data_dict['leave start date']))
var endDate = new Date(parse_time(data_dict['leave end date']))
endDate = new Date(endDate.setDate(endDate.getDate()+1)) // needs +1
var calDescription = data_dict['leave details']
console.log(title, startDate, endDate, calDescription);


// 1 day vs many days use different param
if(data_dict['leave days'] > 1){
    gCal.createAllDayEvent(
      title, 
      startDate, 
      endDate, 
      {description: calDescription}
    ) // multiple days
  } 
else{
    gCal.createAllDayEvent(
      title, 
      startDate, 
      {description: calDescription}
    ) // single day
  }
}

function addAllRecordsToCalendar(){
// calendar setup	
var cals = CalendarApp.getAllOwnedCalendars()
for(i=0;i<cals.length;i++){
  if(cals[i].getName()=='some title here'){
    var gCal = cals[i]
    break
  }
}
console.log(gCal.getName());

var data = getData();
var keys = data[0]; //header

// console.log(data.length)

// delete existing records
// loop to add all
var date = new Date();
var firstDay = new Date(2020, 0, 1); // 2020 jan 1
var lastDay = new Date(date.getFullYear(), date.getMonth()+10, 0); //end of current month

caldata = gCal.getEvents(firstDay, lastDay)
console.log('calendar records to delete: ', caldata.length)

// delete previous calendar entries 
if(caldata){
  for(i=0;i<caldata.length;i++){
    caldata[i].deleteEvent()
    }
}

// add to calendar loop
console.log('calendar records to add: ', data.length)
for(i=1;i<data.length;i++){ // data record index starts from 1 (0 = headers) 
  // retrieve data
  var vals = data[i]; 
  var data_dict = {};
  for (var k in keys) {
    data_dict[keys[k]] = vals[k]
  }
  // console.log(data_dict)

  // calendar input
  var title = data_dict['name'] + ' [' + data_dict['leave type'] +', ' + data_dict['leave days'] + 'days]'
  var startDate = new Date(parse_time(data_dict['leave start date']))
  var endDate = new Date(parse_time(data_dict['leave end date']))
  endDate = new Date(endDate.setDate(endDate.getDate()+1)) // needs +1
  var calDescription = data_dict['leave details']
  // console.log(title, startDate, endDate, calDescription);

  // 1 day vs many days use different param
  if(data_dict['leave days'] > 1){
      gCal.createAllDayEvent(
        title, 
        startDate, 
        endDate, 
        {description: calDescription}
      ) // multiple days
    } 
  else{
      gCal.createAllDayEvent(
        title, 
        startDate, 
        {description: calDescription}
      ) // single day
    }
}
console.log('added total of ' + i + ' records to the calendar')
}

function main(){
var data = getData();
var keys = data[0]; //header
var vals = data[data.length - 1]; //last row

var data_dict = {};

for (var k in keys) {
  data_dict[keys[k]] = vals[k]
}
// console.log(data[data.length - 1]);
// console.log(data_dict);
// Logger.log(data_dict);

console.log(keys);
var timestamp = parse_time(data_dict[keys[0]]); // timestamp 
var admin = data_dict[keys[3]]; //HR header
var requester = data_dict[keys[1]]; //employee name header
var recipientEmail = data_dict[keys[2]] //employee email header
var dept = data_dict[keys[9]]; // employee dept
var position = data_dict[keys[10]]; // employee title
var work_start_date = parse_time(data_dict[keys[11]]); // employee initial work start day
var leave_type = data_dict[keys[4]]; // vacation type
var leave_start = parse_time(data_dict[keys[5]]); // vacation start date 
var leave_end = parse_time(data_dict[keys[6]]); // vacation end date 
var leave_message = data_dict[keys[8]]; // leave detail

var message = getConfirmMessage(
    timestamp, admin, requester, dept, position, work_start_date, leave_type, leave_start,
    leave_end, leave_message );

// send confirmation email
if (recipientEmail.length > 0) {
  var subject = 'this is the leave request confirmation email. 휴가신청 접수 확인 메일입니다.';
  var message = getConfirmMessage(
    timestamp, admin, requester, dept, position, work_start_date, leave_type, leave_start,
    leave_end, leave_message );
  
  MailApp.sendEmail(recipientEmail, subject, message, {htmlBody : message, noReply: true});
  
  //당담자 보내기
  // default 3명
  var ceo_email = '??@???.com';
  var cfo_email = '??@???.com';
  var hr_email = '??@???.com';
  var daigeun_email = '??@???.com';
  var kyungrok_email = '??@???.com';
  var junghan_email = '??@???.com';

  
  if (admin == 'except this guy') {
    var admin_email = '??@???.com';
    subject = requester + ' 님의 휴가 신청 내용 입니다.';
    MailApp.sendEmail(ceo_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(cfo_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(hr_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(admin_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(daigeun_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(kyungrok_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(junghan_email, subject, message, {htmlBody : message, noReply: true});
  } else {
    // 당담자 상관없이 전체 전송 ( 빼고.. )
    subject = requester + ' 님의 휴가 신청 내용 입니다.';
    MailApp.sendEmail(ceo_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(cfo_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(hr_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(daigeun_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(kyungrok_email, subject, message, {htmlBody : message, noReply: true});
    MailApp.sendEmail(junghan_email, subject, message, {htmlBody : message, noReply: true});
  }
}

// add to calendar (2023/12/21)
addToCalendar(data_dict);
} // main wrapper

function test(){
var data = getData();
var keys = data[0]; //header
var vals = data[data.length - 2]; //last row

var data_dict = {};

for (var k in keys) {
  data_dict[keys[k]] = vals[k]
}
console.log(data_dict)
addToCalendar(data_dict)
}

function testCal(){
// calendar setup	
var cals = CalendarApp.getAllOwnedCalendars()
for(i=0;i<cals.length;i++){
  if(cals[i].getName()=='연차 캘린더'){
    var gCal = cals[i]
    break
  }
}
console.log(gCal.getName());

var data = getData();
var keys = data[0]; //header

for (i=800; i<data.length; i++){
  var vals = data[i]; //last row
  var data_dict = {};
  for (var k in keys) {
    data_dict[keys[k]] = vals[k]
  };
  console.log(i);
  console.log(data_dict);
  // addToCalendar(data_dict) 
}
}
