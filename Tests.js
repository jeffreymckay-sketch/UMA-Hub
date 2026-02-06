function FORCE_CALENDAR_AUTH() {
  console.log("Attempting to connect to Calendar API...");
  // This line forces Google to check if you have allowed Calendar access
  const cals = Calendar.CalendarList.list();
  console.log("Success! Found " + cals.items.length + " calendars.");
}