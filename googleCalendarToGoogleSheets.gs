// 1st task => Open trigger
// DO NOT REMOVE THIRD COMMENT => Needed to get user data for different user
//@NotOnlyCurrentDoc 
var ui = SpreadsheetApp.getUi();
// ATTENTION:
// To Enable IonOpen function to run without crush, create open trigger on app script menu, and add IonOpen to it
function IonOpen() {
       ui.createMenu('Мероприятия')
      // .addItem('Показать', 'showAlert')   
      .addItem('Показать сайдбар', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('showSidebar')
      .setTitle('Календарь пользователя');
      ui.showSidebar(html);
}

// Getting user's approval
// function showAlert() {
//   var result = ui.alert(
//     'Для запуска скрипта необходимо подтвердить получение данных вашего календаря.',
//     'Продолжить?',
//     ui.ButtonSet.YES_NO);
  
//   // Process the user's response.
//   if (result == ui.Button.YES) {
//     ui.alert('Спасибо. Теперь нажмите кнопку меню, затем боковую панель.')
//   } else {
//     ui.alert('Отменено');
//   }
// }

// 4th task => Fetch user's all calendar
// Header
var sheet = SpreadsheetApp.getActiveSheet();
var header = [["Весь календарь пользователя", "Сегодня", "Завтра", "Через 2 дня", "Через 3 дней",	"Через 4 дня"]]
var headerRange = sheet.getRange(1,1,1,6);
headerRange.setValues(header);
// Day setter
var today = new Date();
var fiveDaysForward = new Date(today.getTime() + (5 * 24 * 60 * 60 * 1000));
var allCalendarOfUser = CalendarApp.getAllCalendars();

// First Calendar 
function fetchFirstCalendar() {
  var firstCalendarEvents = allCalendarOfUser[0].getEvents(today, fiveDaysForward);
  var firstCalendarName = [[allCalendarOfUser[0].getName()]]
  var firstCalendarNameRange = sheet.getRange(2,1,1,1)
  firstCalendarNameRange.setValues(firstCalendarName)
  //  Looping through events array
  for (var i=0;i<5;i++) {
    var row=i+2;
    var firstCalendar = [[firstCalendarEvents[i].getTitle()]];
    var firstCalendarRange=sheet.getRange(2,row,1,1);
    firstCalendarRange.setValues(firstCalendar);
 }
}

// Second Calendar
function fetchSecondCalendar() {
  var secondCalendarEvents = allCalendarOfUser[1].getEvents(today, fiveDaysForward);
  var secondCalendarName = [[allCalendarOfUser[1].getName()]]
  var secondCalendarNameRange = sheet.getRange(3,1,1,1)
  secondCalendarNameRange.setValues(secondCalendarName)
  //  Looping through events array
  for (var i=0;i<5;i++) {
    var row=i+2;
    var secondCalendar = [[secondCalendarEvents[i].getTitle()]];
    var secondCalendarRange=sheet.getRange(3,row,1,1);
    secondCalendarRange.setValues(secondCalendar);
  }
}

// Third Calendar
function fetchThirdCalendar() {
  var thirdCalendarEvents = allCalendarOfUser[2].getEvents(today, fiveDaysForward);
  var thirdCalendarName = [[allCalendarOfUser[2].getName()]]
  var thirdCalendarNameRange = sheet.getRange(4,1,1,1)
  thirdCalendarNameRange.setValues(thirdCalendarName)
  //  Looping through events array
  for (var i=0;i<5;i++) {
    var row=i+2;
    var thirdCalendar = [[thirdCalendarEvents[i].getTitle()]];
    var thirdCalendarRange=sheet.getRange(4,row,1,1);
    thirdCalendarRange.setValues(thirdCalendar);
  }
}
