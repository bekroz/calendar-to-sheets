// 1st task => Open trigger
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Мероприятия')
      .addItem('Показать сайдбар', 'showSidebar')
      .addToUi();
}

// 3rd task => Sidebar
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('showSidebar')
      .setTitle('Сайдбар')
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

// 4th task => Fetch user's all calendar
// Header
var sheet = SpreadsheetApp.getActiveSheet();
var header = [[ "Весь календарь пользователя", "Сегодня", "Завтра", "Через 2 дня", "Через 3 дней",	"Через 4 дня"]]
var headerRange = sheet.getRange(1,1,1,6);
headerRange.setValues(header);
// Day setter
var today = new Date();
var fiveDaysForward = new Date(today.getTime() + (5 * 24 * 60 * 60 * 1000));
var allCalendarOfUser = CalendarApp.getAllCalendars();

// First => Personal Calendar 
function fetchPersonalCalendar() {
  var personalCalendarID = Session.getActiveUser().getEmail();
  var personalEvents = CalendarApp.getCalendarById(personalCalendarID).getEvents(today, fiveDaysForward);
  var personalCalendar = [[ "Личный календарь", personalEvents[0].getTitle(), personalEvents[1].getTitle(), personalEvents[2].getTitle(), personalEvents[3].getTitle(), personalEvents[4].getTitle()]]
  var firstCalendarRange = sheet.getRange(2,1,1,6);
firstCalendarRange.setValues(personalCalendar);
}

// Second Calendar
function fetchFamilyCalendar() {
  // TODO: Create function to Get FamilyCalendarID
  var secondCalendarID = allCalendarOfUser[1];
  var secondCalendarEvents = CalendarApp.getCalendarById(secondCalendarID).getEvents(today, fiveDaysForward);
  var personalCalendar = [["Семейный календарь", secondCalendarEvents[0].getTitle(), secondCalendarEvents[1].getTitle(), secondCalendarEvents[2].getTitle(), secondCalendarEvents[3].getTitle(), secondCalendarEvents[4].getTitle()]]
  var secondCalendarRange = sheet.getRange(3,1,1,6);
secondCalendarRange.setValues(secondCalendar);
}

// Third Calendar
function fetchWorkCalendar() {
  // TODO: Create function to Get WorkCalendarID
  
  var thirdCalendarID = Session.allCalendarOfUser[2];
  var thirdCalendarEvents = CalendarApp.getCalendarById(thirdCalendarID).thirdCalendarEvents(today, fiveDaysForward);
  var thirdCalendar = [["Рабочий календарь", workEvents[0].getTitle(), thirdCalendarEvents[1].getTitle(), thirdCalendarEvents[2].getTitle(), thirdCalendarEvents[3].getTitle(), thirdCalendarEvents[4].getTitle()]]
  var thirdCalendarRange = sheet.getRange(4,1,1,6);
thirdCalendarRange.setValues(thirdCalendar);
}
