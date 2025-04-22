var SLACK_BOT_TOKEN = "xoxb-XXXXXXXXXX"; // Replace with your bot token

// Log Values in speadsheet
// var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
// var lastRow = sheet.getLastRow();
// var logCell = sheet.getRange(lastRow + 1, 1);
// logCell.setValue(JSON.stringify(e.postData));

// Reurn message for get requests
function doGet(e) {
    return ContentService.createTextOutput("Slack Leave Management Bot is Active");
}

// Handle users post request for checking and adding leaves
function doPost(e) {
    var params = e.parameter;
    var commandText = params.text.split(" ");
    var userId = params.user_id;
    var userName = params.user_name;

    // add logs in speadsheet
    // var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // var lastRow = sheet.getLastRow();
    // var logCell = sheet.getRange(lastRow + 1, 1);
    // logCell.setValue(JSON.stringify(e.postData));

    doTaskAsync(userId, userName, commandText);
    return ContentService.createTextOutput("Please wait, I am processing your request."); // Immediate 200 OK response
}

// Handle Leave management i.e. check and add leaves asynchronously
function doTaskAsync(userId, userName, commandText) {
    // test data start
    // userId = 1;
    // userName = "employee.name";
    // commandText = ["Check"];
    // test data end

    Utilities.sleep(500); // Optional delay for better processing
    if (ensureCurrentYearSheet() === 1 && ensureCurrentUser(userName) === 1) {
        if (commandText[0].toLowerCase() === "check") {
            checkLeaveBalance(userId, userName);
            // sendSlackMessage(userId, checkLeaveBalance(userName));
        } else if (commandText[0].toLowerCase() === "add") {
            if (commandText.length < 2) {
                // sendSlackMessage(userId, "Please provide a date. Usage: `/leave add YYYY-MM-DD`.");
                return;
            }
            sendSlackMessage(userId, addLeave(userId, userName, commandText));
        } else {
            sendSlackMessage(userId, "Invalid command! Use `/leave check` or `/leave add [s | c] YYYY-MM-DD`.");
        }
    } else {
        sendSlackMessage(userId, "An Error has occurred, please try again later or contact your manager.");
    }
}

// Check if the leave management sheet for current year exists, if not then create a new sheet
function ensureCurrentYearSheet() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const currentYear = new Date().getFullYear().toString();
        const sheetExists = ss.getSheets().some((sheet) => sheet.getName() === currentYear);

        if (sheetExists) {
            return 1;
        }

        const newSheet = ss.insertSheet(currentYear);

        const months = [
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ];

        // Total columns: 1 (Employee Name) + (12 * 2) for Sick & Casual + 1 Annual + 3 for Leaves
        const totalNeededCols = 1 + months.length * 2 + 1 + 3; // = 29
        if (newSheet.getMaxColumns() < totalNeededCols) {
            newSheet.insertColumnsAfter(newSheet.getMaxColumns(), totalNeededCols - newSheet.getMaxColumns());
        }

        const mainHeaders = ["Employee Name"];
        const subHeaders = [""];

        let currentCol = 2;

        for (let i = 0; i < months.length; i++) {
            mainHeaders.push(months[i], "");
            subHeaders.push("Sick", "Casual");

            // Merge 2 columns for each month
            newSheet.getRange(1, currentCol, 1, 2).merge().setHorizontalAlignment("center");
            currentCol += 2;
        }

        // Add the "Annual" leave column as a standalone header
        mainHeaders.push("Annual");
        subHeaders.push("");

        // Add Leaves Consumed and Total Leaves
        mainHeaders.push("Leaves Consumed", "Total Leaves");
        subHeaders.push("", "");

        // Set header values
        newSheet.getRange(1, 1, 1, mainHeaders.length).setValues([mainHeaders]);
        newSheet.getRange(2, 1, 1, subHeaders.length).setValues([subHeaders]);

        // Formatting
        newSheet.getRange("1:2").setFontWeight("bold");
        newSheet.setFrozenRows(2);
        // === Background colors ===
        const headerRange = newSheet.getRange(1, 1, 1, mainHeaders.length);
        headerRange.setBackground("#d9ead3"); // light green
        for (let i = 0; i < months.length; i++) {
            const colStart = 2 + i * 2; // Starting col for each month's sub-columns
            const subHeaderRange = newSheet.getRange(2, colStart, 1, 2);
            subHeaderRange.setHorizontalAlignment("center");
            subHeaderRange.setBackground("#fff2cc"); // light yellow
        }

        return 1;
    } catch (e) {
        return 0;
    }
}

// Check if row for current user exists, if not then create a new row
function ensureCurrentUser(userName) {
    const currentYear = new Date().getFullYear().toString();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentYear);
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
        // Start from row 3 (index 2)
        if (data[i][0] === userName) {
            return 1;
        }
    }

    // Build new row data for User
    const months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    const totalCols = 1 + months.length * 2 + 3; // 1 name + 24 monthly + 3 end cols
    const newRow = new Array(totalCols).fill("");

    newRow[0] = userName;

    // Add Annual, Consumed, and Total Leaves
    const annualCol = 1 + months.length * 2; // after month columns
    newRow[annualCol] = 0;
    newRow[annualCol + 1] = 0;
    newRow[annualCol + 2] = 20;

    sheet.appendRow(newRow);
    return 1;
}

// Get the user's leave records and return a string
function checkLeaveBalance(userId, userName) {
    try {
        const currentYear = new Date().getFullYear().toString();
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentYear);
        const data = sheet.getDataRange().getValues();

        const months = [
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ];

        const currentMonth = new Date().toLocaleString("default", { month: "long" });
        const monthIndex = months.indexOf(currentMonth);

        if (monthIndex === -1) {
            return `<@${userId}>, Unable to find current month column.`;
        }

        // Calculate column indexes
        const startCol = 1 + monthIndex * 2; // Sick
        const casualCol = startCol + 1;
        const annualCol = 1 + months.length * 2; // Annual
        const leavesConsumedCol = annualCol + 1;
        const leavesTotalCol = annualCol + 2;

        for (let i = 2; i < data.length; i++) {
            // Start from row 3 (index 2)
            if (data[i][0] === userName) {
                const sickCount = parseFloat(data[i][startCol]) || 0;
                const casualCount = parseFloat(data[i][casualCol]) || 0;
                const annualCount = parseFloat(data[i][annualCol]) || 0;
                const consumedCount = parseFloat(data[i][leavesConsumedCol]) || 0;
                const totalCount = parseFloat(data[i][leavesTotalCol]) || 0;
                let remainingCount = totalCount - consumedCount;
                return `<@${userId}>, You have taken ${sickCount} sick leaves and ${casualCount} casual leaves this month. You have consumed ${annualCount} annual leaves. Your remaining leaves are ${remainingCount} leaves.`;
            }
        }
        return `<@${userId}>, No record found for username "${userName}".`;
    } catch (e) {
        return `<@${userId}>, Error occurred while checking your leave balance: ${e.message}`;
    }
}

// Add new leaves for users
function addLeave(userID, userName, leaveDate) {
    const currentYear = new Date().getFullYear().toString();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentYear);
    sheet.appendRow([userName, leaveDate]); // Simply adds a new row with the user and date
    return `Leave added for ${userName} on ${leaveDate}.`;
}

function sendSlackMessage(userId, message) {
    var url = "https://slack.com/api/chat.postMessage";
    var options = {
        method: "post",
        headers: {
            Authorization: "Bearer " + SLACK_BOT_TOKEN,
            "Content-Type": "application/json",
        },
        payload: JSON.stringify({
            channel: userId, // Send message directly to the user
            text: message,
        }),
    };

    UrlFetchApp.fetch(url, options);
}
