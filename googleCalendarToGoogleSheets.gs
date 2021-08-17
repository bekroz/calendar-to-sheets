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

// First => Personal Calendar 
var personalCalendarID = "bekrozikoff@gmail.com";
var personalEvents = CalendarApp.getCalendarById(personalCalendarID).getEvents(today, fiveDaysForward);
Logger.log('Number of events: ' + personalEvents.length);
var personalCalendar = [[ "Личный календарь", personalEvents[0].getTitle(), personalEvents[1].getTitle(), personalEvents[2].getTitle(), personalEvents[3].getTitle(), personalEvents[4].getTitle()]]
var firstCalendarRange = sheet.getRange(2,1,1,6);
firstCalendarRange.setValues(personalCalendar);

// Second => Family Calendar
var familyCalendarID = "bekrozikoff@gmail.com";
var familyEvents = CalendarApp.getCalendarById(familyCalendarID).getEvents(today, fiveDaysForward);
Logger.log('Number of events: ' + familyEvents.length);
var personalCalendar = [["Семейный календарь", familyEvents[0].getTitle(), familyEvents[1].getTitle(), familyEvents[2].getTitle(), familyEvents[3].getTitle(), familyEvents[4].getTitle()]]
var firstCalendarRange = sheet.getRange(3,1,1,6);
firstCalendarRange.setValues(personalCalendar);
// Second => Work Calendar

// Third => Work Calendar
var familyCalendarID = "bekrozikoff@gmail.com";
var workEvents = CalendarApp.getCalendarById(familyCalendarID).getEvents(today, fiveDaysForward);
Logger.log('Number of events: ' + familyEvents.length);
var workCalendar = [["Рабочий календарь", workEvents[0].getTitle(), workEvents[1].getTitle(), workEvents[2].getTitle(), workEvents[3].getTitle(), workEvents[4].getTitle()]]
var firstCalendarRange = sheet.getRange(4,1,1,6);
firstCalendarRange.setValues(workCalendar);