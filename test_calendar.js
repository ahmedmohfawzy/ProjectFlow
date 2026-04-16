const fs = require('fs');
const js = fs.readFileSync('js/calendar.js', 'utf8');
eval(js);
try {
    let start = new Date('2024-01-01');
    let end = new Date('2024-01-10');
    console.log("getWorkingDays:", WorkCalendar.getWorkingDays(start, end));
    console.log("addWorkingDays:", WorkCalendar.addWorkingDays(start, 5));
} catch(e) { console.error("ERR:", e); }
