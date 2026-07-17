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
  const turnHeaders = turnData.shift();
  weekData.shift();

  const nameIdx = turnHeaders.indexOf('Name');
  const senIdx = turnHeaders.indexOf('SeniorityPosition');
  const lotIdx = turnHeaders.indexOf('LotteryPosition');
  const statusIdx = turnHeaders.indexOf('Status');

  const userRow = turnData.find(row => row[nameIdx] === name);
  const queuePos = currentRound === 1 ? userRow[senIdx] : userRow[lotIdx];
  const currentUser = { name: userRow[nameIdx], queuePosition: queuePos, status: userRow[statusIdx] };

  const turnQueue = turnData.map(row => ({
      name: row[nameIdx],
      queuePosition: currentRound === 1 ? row[senIdx] : row[lotIdx],
      status: row[statusIdx]
  }));

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
  const configSheet = ss.getSheetByName('Config');
  const currentRound = configSheet.getRange("B2").getValue();
  const weekData = weekSheet.getDataRange().getValues();
  const turnData = turnSheet.getDataRange().getValues();
  const turnHeaders = turnData.shift();
  weekData.shift();

  const calendarData = weekData.map(row => ({
    startDate: row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }),
    classification: row[1],
    person1: row[2], person2: row[3], person3: row[4], person4: row[5],
    spotsRemaining: row[6]
  }));

  const nameIdx = turnHeaders.indexOf('Name');
  const senIdx = turnHeaders.indexOf('SeniorityPosition');
  const lotIdx = turnHeaders.indexOf('LotteryPosition');
  const statusIdx = turnHeaders.indexOf('Status');

  const turnQueue = turnData.map(row => ({
    name: row[nameIdx],
    queuePosition: currentRound === 1 ? row[senIdx] : row[lotIdx],
    status: row[statusIdx]
  }));
  
  return {
      calendarData: calendarData,
      turnQueue: turnQueue,
      currentRound: currentRound
  };
}
function setupSpreadsheetSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  if (!turnSheet) return "Turn Management sheet not found.";

  const headersRange = turnSheet.getRange(1, 1, 1, turnSheet.getLastColumn());
  let headers = headersRange.getValues()[0];

  let modifications = false;

  // 1. Rename QueuePosition to SeniorityPosition
  let qIndex = headers.indexOf('QueuePosition');
  if (qIndex !== -1) {
    turnSheet.getRange(1, qIndex + 1).setValue('SeniorityPosition');
    headers[qIndex] = 'SeniorityPosition';
    modifications = true;
  }

  // 2. Add LotteryPosition if missing
  if (headers.indexOf('LotteryPosition') === -1) {
    let newCol = headers.length + 1;
    turnSheet.getRange(1, newCol).setValue('LotteryPosition');
    headers.push('LotteryPosition');
    modifications = true;
  }

  // 3. Add SkipNextTurn if missing
  if (headers.indexOf('SkipNextTurn') === -1) {
    let newCol = headers.length + 1;
    turnSheet.getRange(1, newCol).setValue('SkipNextTurn');

    // Initialize to false
    if (turnSheet.getLastRow() > 1) {
      turnSheet.getRange(2, newCol, turnSheet.getLastRow() - 1, 1).setValue(false);
    }
    headers.push('SkipNextTurn');
    modifications = true;
  }

  return modifications ? "Schema updated successfully." : "Schema already up to date.";
}

function validateSchema(turnData, currentRound) {
    const headers = turnData[0];
    const required = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
    for (let req of required) {
        if (headers.indexOf(req) === -1) {
            return { valid: false, message: "Missing required column: " + req + ". Please run setupSpreadsheetSchema()." };
        }
    }

    // Check LotteryPosition setup if moving past round 1
    if (currentRound > 1) {
        const lotIdx = headers.indexOf('LotteryPosition');
        let usedPositions = new Set();
        for (let i = 1; i < turnData.length; i++) {
            let val = turnData[i][lotIdx];
            if (!val || val === "") return { valid: false, message: "LotteryPosition is missing values for some participants. Cannot proceed past Round 1." };
            if (usedPositions.has(val)) return { valid: false, message: "LotteryPosition contains duplicate values. Must be a unique order." };
            usedPositions.add(val);
        }
    }
    return { valid: true };
}

function processSelection(selectionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekSheet = ss.getSheetByName('Week Availability');
  const configSheet = ss.getSheetByName('Config');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const turnDataRaw = turnSheet.getDataRange().getValues();
    const currentRound = configSheet.getRange("B2").getValue();

    // Schema validation
    const schemaCheck = validateSchema(turnDataRaw, currentRound);
    if (!schemaCheck.valid) { return { success: false, message: "System setup error: " + schemaCheck.message }; }

    // Check if system is completely full
    const weekData = weekSheet.getDataRange().getValues();
    weekData.shift(); // remove header
    let totalSpotsRemaining = 0;
    weekData.forEach(row => {
        let empty = 0;
        for(let i=2; i<=5; i++) { if (!row[i]) empty++; }
        totalSpotsRemaining += empty;
    });
    if (totalSpotsRemaining === 0) {
        return { success: false, message: "Selection Complete: All available vacation slots are filled." };
    }

    const turnHeaders = turnDataRaw.shift();
    const turnData = turnDataRaw;

    const nameIdx = turnHeaders.indexOf('Name');
    const statusIdx = turnHeaders.indexOf('Status');
    const weeksSelectedIdx = turnHeaders.indexOf('WeeksSelected');
    const senPosIdx = turnHeaders.indexOf('SeniorityPosition');
    const lotPosIdx = turnHeaders.indexOf('LotteryPosition');
    const skipIdx = turnHeaders.indexOf('SkipNextTurn');

    const userRowIndex = turnData.findIndex(row => row[nameIdx] === selectionData.name);
    if (userRowIndex === -1) { return { success: false, message: "User not found." }; }

    const userStatus = turnData[userRowIndex][statusIdx];
    if (!['Active', 'Standby', 'Backup'].includes(userStatus)) {
        return { success: false, message: "It is not your turn to make a selection." };
    }

    if (!selectionData.week1) {
        return { success: false, message: "Missing primary week selection." };
    }

    if (currentRound === 1 && selectionData.week2) {
        return { success: false, message: "Invalid selection. You can only select exactly ONE week during Round 1." };
    }

    if (selectionData.week1 === selectionData.week2) {
        return { success: false, message: "You cannot select the same week twice in one submission." };
    }

    let w1Index = weekData.findIndex(row => row[0].getTime() == selectionData.week1);
    let w2Index = selectionData.week2 ? weekData.findIndex(row => row[0].getTime() == selectionData.week2) : -1;

    if (w1Index === -1 || (selectionData.week2 && w2Index === -1)) {
        return { success: false, message: "One of the selected weeks does not exist." };
    }

    let w1Data = weekData[w1Index];
    let w2Data = selectionData.week2 ? weekData[w2Index] : null;

    // Check max capacity and existing spots
    let w1EmptySlots = 0;
    for (let i=2; i<=5; i++) { if (!w1Data[i]) w1EmptySlots++; }
    if (w1EmptySlots === 0) return { success: false, message: "Primary week is full." };

    if (w2Data) {
        let w2EmptySlots = 0;
        for (let i=2; i<=5; i++) { if (!w2Data[i]) w2EmptySlots++; }
        if (w2EmptySlots === 0) return { success: false, message: "Secondary week is full." };
    }

    // Check double booking in the same week
    for(let i=2; i<=5; i++){
        if (w1Data[i] === selectionData.name) return { success: false, message: "You are already booked for the primary week." };
        if (w2Data && w2Data[i] === selectionData.name) return { success: false, message: "You are already booked for the secondary week." };
    }

    if (selectionData.week2) {
        if (w1Data[1] === 'Prime' || w2Data[1] === 'Prime') {
            return { success: false, message: "Invalid selection. You cannot include a Prime week when selecting two weeks." };
        }
    }

    // If validation passes, apply changes atomically.
    let w1TargetCol = -1;
    for(let i=2; i<=5; i++){
        if(!weekSheet.getRange(w1Index + 2, i + 1).getValue()) {
            w1TargetCol = i + 1;
            break;
        }
    }
    if (w1TargetCol === -1) throw new Error("Concurrency error: primary week filled up.");

    let w2TargetCol = -1;
    if (selectionData.week2) {
        for(let i=2; i<=5; i++){
            if(!weekSheet.getRange(w2Index + 2, i + 1).getValue()) {
                w2TargetCol = i + 1;
                break;
            }
        }
        if (w2TargetCol === -1) throw new Error("Concurrency error: secondary week filled up.");
    }

    // Perform writes
    weekSheet.getRange(w1Index + 2, w1TargetCol).setValue(selectionData.name);
    weekSheet.getRange(w1Index + 2, 7).setValue(w1EmptySlots - 1);
    totalSpotsRemaining -= 1;

    if (selectionData.week2) {
        weekSheet.getRange(w2Index + 2, w2TargetCol).setValue(selectionData.name);
        let currentSpots = (w2Data ? (4 - (w2Data.filter((_, idx) => idx >= 2 && idx <= 5 && w2Data[idx]).length)) : 0);
        let newW2Spots = Math.max(0, currentSpots - 1);
        weekSheet.getRange(w2Index + 2, 7).setValue(newW2Spots);
        totalSpotsRemaining -= 1;
    }

    // Update Turn Sheet for current user
    const weeksPickedCount = selectionData.week2 ? 2 : 1;
    const currentWeeksSelected = turnData[userRowIndex][weeksSelectedIdx];
    turnSheet.getRange(userRowIndex + 2, weeksSelectedIdx + 1).setValue(currentWeeksSelected + weeksPickedCount);
    turnSheet.getRange(userRowIndex + 2, statusIdx + 1).setValue('Completed');

    if (selectionData.week2) {
        turnSheet.getRange(userRowIndex + 2, skipIdx + 1).setValue(true);
    }

    // If we just filled the last spot in the whole sheet, clear queue and exit early
    if (totalSpotsRemaining === 0) {
        let finalResetData = turnSheet.getDataRange().getValues();
        finalResetData.shift();
        finalResetData.forEach((row, index) => {
            turnSheet.getRange(index + 2, statusIdx + 1).setValue('Waiting');
        });
        return { success: true, message: "Selection recorded. Selection process is now complete." };
    }

    // RE-EVALUATE QUEUE STATE (Loop to clear skips, advancing round if necessary)
    let nextRound = currentRound;
    let filledSlots = 0;
    const windowSlots = ['Active', 'Standby', 'Backup'];
    let activatedUserIndices = [];

    // Reset all non-Completed to Waiting
    let resetTurnData = turnSheet.getDataRange().getValues();
    resetTurnData.shift();
    resetTurnData.forEach((row, index) => {
        if (row[statusIdx] !== 'Completed') {
            turnSheet.getRange(index + 2, statusIdx + 1).setValue('Waiting');
        }
    });

    while (filledSlots < 3) {
        let currentTurnData = turnSheet.getDataRange().getValues();
        currentTurnData.shift();

        let isRoundOver = currentTurnData.every(row => row[statusIdx] === 'Completed');

        if (isRoundOver) {
            nextRound++;
            configSheet.getRange("B2").setValue(nextRound);
            currentTurnData.forEach((row, index) => {
                turnSheet.getRange(index + 2, statusIdx + 1).setValue('Waiting');
                currentTurnData[index][statusIdx] = 'Waiting';
            });
            // Also need to validate the newly started round for schema
            let validation = validateSchema([turnHeaders].concat(currentTurnData), nextRound);
            if(!validation.valid) throw new Error("Blocked advancing round: " + validation.message);
        }

        let eligibleUsers = [];
        currentTurnData.forEach((row, index) => {
            if (row[statusIdx] !== 'Completed' && !activatedUserIndices.includes(index + 2)) {
                eligibleUsers.push({ data: row, originalIndex: index + 2 });
            }
        });

        if (eligibleUsers.length === 0) break; // End of everyone if no one left

        if (nextRound === 1) {
            eligibleUsers.sort((a, b) => a.data[senPosIdx] - b.data[senPosIdx]);
        } else {
            const isEvenRound = nextRound % 2 === 0;
            if (isEvenRound) { // Round 2, 4 -> Top-to-bottom
                eligibleUsers.sort((a, b) => a.data[lotPosIdx] - b.data[lotPosIdx]);
            } else { // Round 3, 5 -> Bottom-to-top
                eligibleUsers.sort((a, b) => b.data[lotPosIdx] - a.data[lotPosIdx]);
            }
        }

        let candidate = eligibleUsers[0]; // Take top candidate

        if (candidate.data[skipIdx] === true) {
            // Consume skip and mark completed for this round
            turnSheet.getRange(candidate.originalIndex, skipIdx + 1).setValue(false);
            turnSheet.getRange(candidate.originalIndex, statusIdx + 1).setValue('Completed');
        } else {
            // Add to window
            turnSheet.getRange(candidate.originalIndex, statusIdx + 1).setValue(windowSlots[filledSlots]);
            activatedUserIndices.push(candidate.originalIndex);
            filledSlots++;
        }
    }

    return { success: true, message: "Selection recorded." };
  } catch (e) {
    return { success: false, message: "An error occurred: " + e.message };
  } finally {
    lock.releaseLock();
  }
}
