function generateRotaForNextFourWeeks() {
    const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Preference form response sheet
    const rotaSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Active sheet for rota 
    rotaSheet.clear(); // Clear the current sheet 

    // Set the start date to the upcoming Monday
    const startDate = getNextMonday(new Date());

    // Set the end date to 4 weeks (20 working days) after the start date 
    const endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + (4 * 7) - 1); // 4 weeks inclusive of the start day 

    // Fetch preference form responses 
    const formData = formSheet.getDataRange().getValues();

    // Fetch bank holidays from helper function 'fetchBankHolidays'
    const holidays = fetchBankHolidays();

    // Set first cell of header and first cell of second row 
    const headerRow = ['Employee Name']; // Labeled 'Employee Name'
    const dateRow = ['']; // Set as blank to keep alignment 
    const holidayColumns = []; // Array to store which columns are bank holidays

    // Track how many office days each employee has per week
    let officeDaysPerWeek = {};
    const employees = extractUniqueEmployees(formData);

    // Initialize each employee with 0 office days per week
    employees.forEach(employee => {
        officeDaysPerWeek[employee] = { week1: 0, week2: 0, week3: 0, week4: 0 };
    });

    // Generate headers for each workday over 4 weeks eg. Monday, Tuesday...
    let currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        if (currentDate.getDay() >= 1 && currentDate.getDay() <= 5) { // Check if current day is a weekday 
            headerRow.push(getDayName(currentDate)); // Add short day name to sheet
            dateRow.push(formatDateShort(currentDate)); // Add formatted date to sheet 

            // Check if it's a bank holiday
            if (isBankHoliday(currentDate, holidays)) {
                holidayColumns.push(headerRow.length); // Track the column if it's a holiday
            }
        }
        currentDate.setDate(currentDate.getDate() + 1); // Move to next day
    }

    // Set headers in the sheet 
    rotaSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]); // Set day names 
    rotaSheet.getRange(2, 1, 1, dateRow.length).setValues([dateRow]); // Set corresponding date 

    // Populate the shifts for each employee
    for (let e = 0; e < employees.length; e++) {
        const employeeRow = [employees[e]];  // Start the row with the employee's name
        
        // Reset currentDate to the start date of the 4-week period
        currentDate = new Date(startDate);
        let columnIndex = 1; // Column index starts after "Employee Name"
        
        while (currentDate <= endDate) {
            if (currentDate.getDay() >= 1 && currentDate.getDay() <= 5) { // Only Monday to Friday
                const weekNumber = getWeekNumber(currentDate); // Determine the week number

                if (isBankHoliday(currentDate, holidays)) {
                    employeeRow.push('B/H'); // Mark as "B/H" for bank holiday
                } else {
                    // Determine if the employee needs more office days
                    const needsOffice = officeDaysPerWeek[employees[e]][`week${weekNumber}`] < 2;
                    const shift = assignShiftForDay(currentDate, employees[e], needsOffice, weekNumber, officeDaysPerWeek, formData);
                    employeeRow.push(shift);
                }
                columnIndex++;
            }
            currentDate.setDate(currentDate.getDate() + 1); // Move to the next day
        }
        
        // Add the employee row to the sheet
        rotaSheet.getRange(3 + e, 1, 1, employeeRow.length).setValues([employeeRow]);
    }

    // Mark entire columns as "B/H" for holidays
    holidayColumns.forEach(function(colIndex) {
        rotaSheet.getRange(3, colIndex, employees.length, 1).setValue('B/H');
    });

    SpreadsheetApp.getUi().alert('Rota for the next 4 weeks generated successfully, with bank holidays marked!');
}

// Helper function to assign shifts for a day (home or office) and ensure at least one manager is in the office
function assignShiftForDay(currentDate, employee, needsOffice, weekNumber, officeDaysPerWeek, formData) {
    const dayOfWeek = currentDate.getDay(); // 1 = Monday, 5 = Friday
    const shifts = (dayOfWeek === 5) ? ['Early', 'Mid'] : ['Early', 'Mid', 'Late']; // Friday has only Early and Mid shifts

    let shiftType = shifts[Math.floor(Math.random() * shifts.length)]; // Randomly assign a shift type

    const isManager = isEmployeeManager(employee, formData); // Check if the employee is a manager

    // Prioritize assigning a manager to the office if no manager has been assigned yet
    if (isManager && !isManagerAssignedToOffice(currentDate, employees, formData)) {
        officeDaysPerWeek[employee][`week${weekNumber}`]++; // Increment office days for the manager
        return `${shiftType} (Office)`; // Assign manager to the office
    }

    // Assign office or home shift based on availability
    if (needsOffice) {
        officeDaysPerWeek[employee][`week${weekNumber}`]++; // Increment office days for the employee
        return `${shiftType} (Office)`; // Assign office shift
    } else {
        return `${shiftType} (Home)`; // Assign home shift
    }
}

// Helper function to check if a manager has been assigned to the office for the day
function isManagerAssignedToOffice(currentDate, employees, formData) {
    for (let i = 0; i < employees.length; i++) {
        const isManager = isEmployeeManager(employees[i], formData);
        if (isManager) {
            return true; // A manager has already been assigned to the office
        }
    }
    return false; // No manager assigned to the office yet
}

// Helper function to check if an employee is a manager based on the form data
function isEmployeeManager(employee, formData) {
    for (let i = 1; i < formData.length; i++) {
        if (formData[i][1] === employee) {
            return formData[i][6] === 'true'; // Column 6 identifies if the employee is a manager
        }
    }
    return false;
}

// Helper function to fetch bank holidays
function fetchBankHolidays() {
    const url = "https://www.gov.uk/bank-holidays.json"; // URL is defined within the function
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    // Fetching for "England and Wales"
    return data["england-and-wales"].events.map(function(event) {
      return new Date(event.date); // Parse the date
    });
}

// Helper function to calculate the next Monday from the current date
function getNextMonday(date) {
    const dayOfWeek = date.getDay(); // Sunday = 0, Monday = 1, etc.
    const daysUntilNextMonday = (dayOfWeek === 0) ? 1 : 8 - dayOfWeek; // How many days until next Monday
    date.setDate(date.getDate() + daysUntilNextMonday);
    return date;
}

// Helper function to get the name of the day (e.g., "MON")
function getDayName(date) {
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return days[date.getDay()].substring(0, 3).toUpperCase(); // Returns 'MON', 'TUES', etc.
}

// Helper function to format a date as "DD-MMM" (e.g., "30-Dec")
function formatDateShort(date) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return date.getDate() + '-' + months[date.getMonth()];
}

// Helper function to extract unique employees from form data
function extractUniqueEmployees(formData) {
    const employees = [];
    for (let i = 1; i < formData.length; i++) {
      const employee = formData[i][1];  // Employee Name is in column 2
      if (employees.indexOf(employee) === -1) {
        employees.push(employee);  // Only add if not already in the list
      }
    }
    return employees;
}

// Helper function to check if a given date is a bank holiday
function isBankHoliday(date, holidays) {
    for (let i = 0; i < holidays.length; i++) {
      if (holidays[i].getTime() === date.getTime()) {
        return true; // It's a bank holiday
      }
    }
    return false;
}

// Helper function to determine the week number (1 to 4)
function getWeekNumber(currentDate) {
    const firstMonday = getNextMonday(new Date(currentDate.getFullYear(), currentDate.getMonth(), 1));
    const daysSinceFirstMonday = Math.floor((currentDate - firstMonday) / (1000 * 60 * 60 * 24));
    return Math.ceil((daysSinceFirstMonday + 1) / 7); // Get the week number (1-4)
}
