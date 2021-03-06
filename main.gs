function deleteEvent(event) {
  if (typeof event != 'undefined') {
    Logger.log("Deleting event %s", event.getTitle())
    event.deleteEvent()
  }
}

function deleteEvents(eventCal,startTime,endTime,title){

        //  var oldEvents = eventCal.getEvents(startTime, endTime, {search: title});  
          var oldEvents = eventCal.getEvents(startTime, endTime);
          Logger.log("oldEvents %s", oldEvents);
          for (var j = 0; j < oldEvents.length; j++){  
             Logger.log("oldEvents[j] %s", oldEvents[j]);
             deleteEvent(oldEvents[j]);
          }
          

}

function min(a, b){
  if (a<b && a != ''){
    return a
  } else {
    return b
  }

}

function max(a, b){
  if (a>b && a != ''){
    return a
  } else {
    return b
  }

}

function addHours(time, amount_hours){

    var endTime = new Date(time);
    //  var addTime = new Date(3600000);  
   var endTime_1h = new Date(endTime.getTime()+amount_hours*3600000);
   Logger.log("New Time is %s", endTime_1h)
   return endTime_1h
}



function main(){
// инициализация
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet0 = spreadsheet.getSheets()[0];
var sheet1 = spreadsheet.getSheets()[1];
//var calendarId = sheet1.getRange('K2').getValue();
var AntonGoryniaCalendarId = '@group.calendar.google.com';
var PMcalendarId = '@group.calendar.google.com';
var BulatCalendarId ='@group.calendar.google.com';
var AntonKCalendarId ='@group.calendar.google.com';
var MarinaPCalendarId ='@group.calendar.google.com';
  
var AntonGEventCal = CalendarApp.getCalendarById(AntonGoryniaCalendarId);
var PMeventCal = CalendarApp.getCalendarById(PMcalendarId);
var BulatEventCal = CalendarApp.getCalendarById(BulatCalendarId);
var AntonKEventCal = CalendarApp.getCalendarById(AntonKCalendarId);
var MarinaPEventCal = CalendarApp.getCalendarById(MarinaPCalendarId);
  
//для удаления каленадря и да 119 = 2019 а 11 - это декабрь
var fromDate = new Date(2020,0,1,0,0,0);
var toDate = new Date(2020,11,31,0,0,0);
 Logger.log("fromeDate %s", fromDate);   
 Logger.log("toDate %s", toDate); 
  
var timeZone = "GMT+03:00"
var lastrow = sheet0.getLastRow();
var startTime = 0;
var addTime = 0;  
  
var data = sheet0.getDataRange().getValues() ;  
  
var date_start_bt = [];  
var date_end_bt = []; 
var date_start_meeting = []; 
var date_start_meetingS = [];
var date_end_meeting = [];
var date_end_meetingS = [];
var meeting_days = [];
var total_days = [];
var q = [];
var year = [];
var city = [];
var area = [];
var country = [];
var type = [];
var topic = [];
var speaker = [];
var speaker_type = [];
var status = [];
 // процес инициализации и сброса календаря 
  
deleteEvents(AntonGEventCal,fromDate,toDate,title);
deleteEvents(BulatEventCal,fromDate,toDate,title);
deleteEvents(AntonKEventCal,fromDate,toDate,title); 
deleteEvents(MarinaPEventCal,fromDate,toDate,title);
deleteEvents(PMeventCal,fromDate,toDate,title);
  // i = 99 - что бы начать с 100 строки где начинается 2020 
  
  for (var i = 99; i < data.length; i++) {        
   //        date_start_bt[i] = data[i][0];  
   //        date_end_bt[i] = data[i][1]; 
   //        date_start_meeting[i] = data[i][2]; 
           //date_start_meetingS[i] = Utilities.formatDate(new Date( date_start_meeting), "GMT+3", "dd-MM-yyyy");
   //        date_end_meeting[i] = data[i][3];
          // date_end_meetingS[i] = Utilities.formatDate(new Date( date_end_meeting), "GMT+3", "dd-MM-yyyy");
           meeting_days[i] = data[i][4];
           total_days[i] = data[i][5];
           q[i] = data[i][6];
           year[i] = data[i][7];
           city = data[i][8];
           area[i] = data[i][9];
           country[i] = data[i][10];
           type = data[i][11];
           topic[i] = data[i][12];
           speaker[i] = data[i][13];
           speaker_type[i] = data[i][14];
           status[i] = data[i][15];  
// удалить дишние переменные       
           date_start_bt = data[i][0];  
           date_end_bt = data[i][1]; 
           date_start_meeting = data[i][2]; 
           date_start_meetingS = Utilities.formatDate(new Date( date_start_meeting), "GMT+3", "dd-MM-yyyy");
           date_end_meeting = data[i][3];
           date_end_meetingS = Utilities.formatDate(new Date( date_end_meeting), "GMT+3", "dd-MM-yyyy");

  
          var title = speaker[i] +' '+ topic[i];
          Logger.log('title  : '+title); 
    
 

          var startTime = min(date_start_bt, date_start_meeting );
          var endTime = max(date_end_bt,date_end_meeting );
  

    
          if (date_start_meeting == date_end_meeting && date_start_meeting == ''  ){
              var startTime = new Date(3600000);
              var endTime = new Date(2*3600000);
          }
    
          //Logger.log('startTime  : '+startTime); 
    
          // добавляем 1 час к дате.
        //   var endTime = new Date(endTime);
         //  var addTime = new Date(3600000);  
         //  var endTime_1h = new Date(endTime.getTime()+2*3600000);
           var endTime_1h = addHours(endTime, 1);
  
           
    
           //добавляем события в календарь.  
          // addEventToCal(eventCal,title, speaker, startTime, endTime, city,date_start_meetingS , date_end_meetingS, total_days);
           

           if (speaker[i] =='Антон Горыня') {
             var event = AntonGEventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.RED);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           } 
           if (speaker[i] =='Марина Привалова') {
             var event = MarinaPEventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.PALE_RED);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Антон Кузин') {
             var event = AntonKEventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.BLUE);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Булат Хафизов') {
             var event = BulatEventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.GREEN);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Максим Савельев') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.CYAN);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Константин Белянин') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.GRAY);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Tatiana Tan') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.MAUVE);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Julia Yu') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.ORANGE);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Иван Григорьев') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.PALE_BLUE);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Ray Zhang') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.YELLOW);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Борис Дятлов') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.PALE_GREEN);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }    
           if (speaker[i] =='Дмитрий Джан') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.PALE_GREEN);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }
           if (speaker[i] =='Max Zhang ') {             
             var event = PMeventCal.createEvent(title,   new Date(startTime),    new Date(endTime_1h), {location: city, description: 'Тип мероприятия: '+ type +'\nДаты мероприятия: ' + date_start_meetingS + '--' + date_end_meetingS} ).setColor(CalendarApp.EventColor.PALE_GREEN);
             if (total_days[i] == '1'){
                 event.setAllDayDate(endTime);               
               }
           }

  
    
    
  } 

}
