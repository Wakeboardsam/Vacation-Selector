// Final complete Code.gs file - Adds turn data to public view
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Vacation Week Selection System');
}
function getParticipantNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Turn Management');
  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return names.flat();
}
function verifyUser(name, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Turn Management');
  const data = sheet.getDataRange().getValues();
  const userRow = data.find(row => row[0] === name);
  if (userRow) {
    if (String(userRow[1]) === String(pin)) { return "Success"; }
  }
  return "Invalid PIN";
}
function getDashboardData(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekSheet = ss.getSheetByName('Week Availability');
  const configSheet = ss.getSheetByName('Config');
  const turnData = turnSheet.getDataRange().getValues();
  const weekData = weekSheet.getDataRange().getValues();
  const currentRound = configSheet.getRange("B2").getValue();
  turnData.shift();
  weekData.shift();
  const userRow = turnData.find(row => row[0] === name);
  const currentUser = { name: userRow[0], queuePosition: userRow[2], status: userRow[3] };
  const turnQueue = turnData.map(row => ({ name: row[0], queuePosition: row[2], status: row[3] }));
  const availableWeeks = weekData.filter(row => row[6] > 0).map(row => ({
    displayDate: row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }),
    valueDate: row[0].getTime(),
    classification: row[1],
    spotsRemaining: row[6]
  }));
  return {
    currentUser: currentUser,
    turnQueue: turnQueue,
    availableWeeks: availableWeeks,
    currentRound: currentRound
  };
}
function getPublicCalendarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekSheet = ss.getSheetByName('Week Availability');
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekData = weekSheet.getDataRange().getValues();
  const turnData = turnSheet.getDataRange().getValues();
  weekData.shift();
  turnData.shift();

  const calendarData = weekData.map(row => ({
    startDate: row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }),
    classification: row[1],
    person1: row[2], person2: row[3], person3: row[4], person4: row[5],
    spotsRemaining: row[6]
  }));

  const turnQueue = turnData.map(row => ({
    name: row[0],
    queuePosition: row[2],
    status: row[3]
  }));
  
  return {
      calendarData: calendarData,
      turnQueue: turnQueue
  };
}
function processSelection(selectionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekSheet = ss.getSheetByName('Week Availability');
  const configSheet = ss.getSheetByName('Config');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const turnData = turnSheet.getDataRange().getValues();
    const weekData = weekSheet.getDataRange().getValues();
    const currentRound = configSheet.getRange("B2").getValue();
    const turnHeaders = turnData.shift();
    weekData.shift();
    const nameIndex = turnHeaders.indexOf('Name');
    const statusIndex = turnHeaders.indexOf('Status');
    const weeksSelectedIndex = turnHeaders.indexOf('WeeksSelected');
    const queuePosIndex = turnHeaders.indexOf('QueuePosition');
    const userRowIndex = turnData.findIndex(row => row[nameIndex] === selectionData.name);
    if (userRowIndex === -1) { return { success: false, message: "User not found." }; }
    const userStatus = turnData[userRowIndex][statusIndex];
    if (!['Active', 'Standby', 'Backup'].includes(userStatus)) { return { success: false, message: "It is not your turn to make a selection." }; }
    if (selectionData.week1 && selectionData.week2) {
      const week1Info = weekData.find(row => row[0].getTime() == selectionData.week1);
      const week2Info = weekData.find(row => row[0].getTime() == selectionData.week2);
      if (!week1Info || !week2Info) { return { success: false, message: "Error: One of the selected weeks could not be found." }; }
      if (week1Info[1] === 'Prime' || week2Info[1] === 'Prime') { return { success: false, message: "Invalid selection. You cannot include a Prime week when selecting two weeks." }; }
    }
    const weeksPickedCount = selectionData.week2 ? 2 : 1;
    const currentWeeksSelected = turnData[userRowIndex][weeksSelectedIndex];
    turnSheet.getRange(userRowIndex + 2, weeksSelectedIndex + 1).setValue(currentWeeksSelected + weeksPickedCount);
    turnSheet.getRange(userRowIndex + 2, statusIndex + 1).setValue('Completed');
    const weeksToUpdate = [selectionData.week1, selectionData.week2].filter(w => w);
    weeksToUpdate.forEach(weekValue => {
      const weekIndex = weekData.findIndex(row => row[0].getTime() == weekValue);
      if (weekIndex !== -1) {
        const weekRowValues = weekSheet.getRange(weekIndex + 2, 3, 1, 5).getValues()[0];
        for (let i = 0; i < 4; i++) {
          if (!weekRowValues[i]) {
            weekSheet.getRange(weekIndex + 2, i + 3).setValue(selectionData.name);
            break;
          }
        }
        const spots = weekRowValues[4];
        if (spots > 0) { weekSheet.getRange(weekIndex + 2, 7).setValue(spots - 1); }
      }
    });
    let updatedTurnData = turnSheet.getDataRange().getValues();
    updatedTurnData.shift();
    const isRoundOver = updatedTurnData.every(row => row[weeksSelectedIndex] >= currentRound);
    let nextRound = currentRound;
    if (isRoundOver) {
      nextRound++;
      configSheet.getRange("B2").setValue(nextRound);
    }
    let eligibleUsers = [];
    updatedTurnData.forEach((row, index) => {
      if (row[weeksSelectedIndex] < nextRound) {
        eligibleUsers.push({ data: row, originalIndex: index + 2 });
      }
    });
    const isEvenRound = nextRound % 2 === 0;
    if (isEvenRound) {
      eligibleUsers.sort((a, b) => b.data[queuePosIndex] - a.data[queuePosIndex]);
    } else {
      eligibleUsers.sort((a, b) => a.data[queuePosIndex] - b.data[queuePosIndex]);
    }
    if (isRoundOver) {
      turnSheet.getRange(2, statusIndex + 1, turnSheet.getLastRow() - 1, 1).setValue('Waiting');
    } else {
      const turnSheetData = turnSheet.getDataRange().getValues();
      turnSheetData.shift();
      turnSheetData.forEach((row, index) => {
        if(row[statusIndex] !== 'Completed') {
          turnSheet.getRange(index + 2, statusIndex + 1).setValue('Waiting');
        }
      });
    }
    if (eligibleUsers.length > 0) turnSheet.getRange(eligibleUsers[0].originalIndex, statusIndex + 1).setValue('Active');
    if (eligibleUsers.length > 1) turnSheet.getRange(eligibleUsers[1].originalIndex, statusIndex + 1).setValue('Standby');
    if (eligibleUsers.length > 2) turnSheet.getRange(eligibleUsers[2].originalIndex, statusIndex + 1).setValue('Backup');
    return { success: true, message: "Selection recorded." };
  } catch (e) {
    return { success: false, message: "An error occurred: " + e.message };
  } finally {
    lock.releaseLock();
  }
}
