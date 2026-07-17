// Final complete Code.gs file - Adds turn data to public view
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Vacation Week Selection System');
}
function normalizeClassification(value) {
  const normalized = String(value || '').trim().toLowerCase();

  if (normalized === 'prime') return 'Prime';
  if (
    normalized === 'non-prime' ||
    normalized === 'non prime' ||
    normalized === 'nonprime'
  ) {
    return 'Non-Prime';
  }

  return null;
}

function buildAvailableWeekData(rows) {
  return rows
    .filter(row => Number(row[6]) > 0)
    .map(row => ({
      displayDate: row[0] instanceof Date ? row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }) : String(row[0]),
      valueDate: row[0] instanceof Date ? row[0].getTime() : null,
      classification: normalizeClassification(row[1]),
      spotsRemaining: Number(row[6]),
      originalClassification: row[1]
    }));
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
function calculateQueueWindow(turnDataRaw, currentRound) {
    const turnData = [...turnDataRaw];
    const turnHeaders = turnData.shift();

    const nameIdx = turnHeaders.indexOf('Name');
    const statusIdx = turnHeaders.indexOf('Status');
    const senPosIdx = turnHeaders.indexOf('SeniorityPosition');
    const lotPosIdx = turnHeaders.indexOf('LotteryPosition');
    const skipIdx = turnHeaders.indexOf('SkipNextTurn');

    // Create an array of objects to sort easily
    let queue = turnData.map((row, index) => ({
        originalIndex: index + 1, // +1 because we shifted headers, and another +1 for 1-based indexing in sheets if needed, but we'll use original row data
        name: row[nameIdx],
        status: row[statusIdx],
        skipNextTurn: row[skipIdx],
        seniorityPosition: row[senPosIdx],
        lotteryPosition: row[lotPosIdx],
        queuePosition: currentRound === 1 ? row[senPosIdx] : row[lotPosIdx],
        computedStatus: 'Waiting'
    }));

    // Sort queue based on current round rules
    if (currentRound === 1) {
        queue.sort((a, b) => a.seniorityPosition - b.seniorityPosition);
    } else {
        const isEvenRound = currentRound % 2 === 0;
        if (isEvenRound) {
            queue.sort((a, b) => a.lotteryPosition - b.lotteryPosition);
        } else {
            queue.sort((a, b) => b.lotteryPosition - a.lotteryPosition);
        }
    }

    // Find the anchor: the first person who is not 'Completed' and not 'skipped' (if we were consuming skips)
    // Wait, the requirement says: "Do not consume SkipNextTurn in read-only getter functions. Consume skips only under the script lock during state transitions."
    // And "A completed or skipped interior position leaves a hole; it must not pull a fourth person into the window while the earliest unfinished person remains."

    let anchorIndex = queue.findIndex(person => person.status !== 'Completed');

    if (anchorIndex !== -1) {
        const windowSlots = ['Active', 'Standby', 'Backup'];
        // The permitted window consists of that queue position and the next two queue positions only.
        for (let offset = 0; offset < 3; offset++) {
            const personIndex = anchorIndex + offset;
            if (personIndex < queue.length) {
                let person = queue[personIndex];
                if (person.status !== 'Completed' && person.skipNextTurn !== true) {
                    person.computedStatus = windowSlots[offset];
                } else if (person.skipNextTurn === true && person.status !== 'Completed') {
                    // If they are skipping, they leave a hole.
                    person.computedStatus = 'Waiting';
                } else if (person.status === 'Completed') {
                     // If they are completed but inside the window (can happen if someone completes out of order), they leave a hole
                    person.computedStatus = 'Completed';
                }
            }
        }
    } else {
        // Everyone is completed
        queue.forEach(p => p.computedStatus = 'Completed');
    }

    // Sort back to original order or keep sorted order for return. We'll return sorted queue.
    return queue;
}

function getDashboardData(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekSheet = ss.getSheetByName('Week Availability');
  const configSheet = ss.getSheetByName('Config');
  const turnDataRaw = turnSheet.getDataRange().getValues();
  const weekData = weekSheet.getDataRange().getValues();
  const currentRound = configSheet.getRange("B2").getValue();

  weekData.shift();

  const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);

  const userObj = queueWindow.find(p => p.name === name);
  const currentUser = { name: userObj.name, queuePosition: userObj.queuePosition, status: userObj.computedStatus };

  const turnQueue = queueWindow.map(p => ({
      name: p.name,
      queuePosition: p.queuePosition,
      status: p.computedStatus
  }));

  const availableWeeks = buildAvailableWeekData(weekData);

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
  const turnDataRaw = turnSheet.getDataRange().getValues();
  weekData.shift();

  const calendarData = weekData.map(row => ({
    startDate: row[0] instanceof Date ? row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }) : String(row[0]),
    classification: normalizeClassification(row[1]) || row[1], // fallback for calendar display
    person1: row[2], person2: row[3], person3: row[4], person4: row[5],
    spotsRemaining: row[6]
  }));

  const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);

  const turnQueue = queueWindow.map(p => ({
    name: p.name,
    queuePosition: p.queuePosition,
    status: p.computedStatus
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
    return { valid: true };
}

function checkLotteryReady(turnData) {
    const headers = turnData[0];
    const lotIdx = headers.indexOf('LotteryPosition');
    if (lotIdx === -1) return false;
    let usedPositions = new Set();
    for (let i = 1; i < turnData.length; i++) {
        let val = turnData[i][lotIdx];
        if (val === "" || val === null || val === undefined) return false;
        if (usedPositions.has(val)) return false;
        usedPositions.add(val);
    }
    return true;
}

function initializeLotteryRound() {
    setupSpreadsheetSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const turnSheet = ss.getSheetByName('Turn Management');
    const configSheet = ss.getSheetByName('Config');

    const currentRound = configSheet.getRange("B2").getValue();
    if (currentRound !== 1) {
        return "Not in Round 1. No action taken.";
    }

    let turnData = turnSheet.getDataRange().getValues();
    const headers = turnData[0];
    const statusIdx = headers.indexOf('Status');
    const lotPosIdx = headers.indexOf('LotteryPosition');

    let isRound1Over = true;
    for (let i=1; i<turnData.length; i++) {
        if (turnData[i][statusIdx] !== 'Completed') {
            isRound1Over = false;
            break;
        }
    }

    if (!isRound1Over) {
        return "Round 1 is not fully complete. No action taken.";
    }

    if (!checkLotteryReady(turnData)) {
        return "LotteryPosition is not correctly populated. Must have exactly one unique value per participant.";
    }

    // Advance to Round 2
    configSheet.getRange("B2").setValue(2);

    // Reset all to Waiting and setup Active window
    let eligibleUsers = [];
    for (let i=1; i<turnData.length; i++) {
        turnSheet.getRange(i+1, statusIdx+1).setValue('Waiting');
        eligibleUsers.push({ data: turnData[i], originalIndex: i+1 });
    }

    // Status is calculated dynamically now, we only need to reset Completed to Waiting

    return "Lottery Round 2 Initialized Successfully.";
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

    const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);
    const userObj = queueWindow.find(p => p.name === selectionData.name);

    if (!userObj) { return { success: false, message: "User not found." }; }

    if (!['Active', 'Standby', 'Backup'].includes(userObj.computedStatus)) {
        return { success: false, message: "It is not your turn to make a selection." };
    }

    // We also need the userRowIndex in the original turnData array
    const userRowIndex = turnData.findIndex(row => row[nameIdx] === selectionData.name);

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
        if (normalizeClassification(w1Data[1]) === 'Prime' || normalizeClassification(w2Data[1]) === 'Prime') {
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
        return { success: true, message: "Selection recorded. Selection process is now complete." };
    }

    // Since authorization relies on computed queue window and 'Completed' status,
    // we don't strictly need to write 'Waiting', 'Active', etc. to the sheet anymore
    // except for skips. Let's consume skips that fall within the *new* window.
    // Wait, the rule says: "Consume skips only under the script lock during state transitions."
    // If a person inside the new 3-person window has SkipNextTurn = true, we consume it and mark them Completed.

    let loopGuard = 0;
    let nextRound = currentRound;

    while (loopGuard < 100) {
        loopGuard++;
        let currentTurnDataRaw = turnSheet.getDataRange().getValues();
        let queueWindowData = calculateQueueWindow(currentTurnDataRaw, nextRound);

        // Find if anyone in the new window (offset 0, 1, 2 from anchor) has skipNextTurn = true
        let anchorIndex = queueWindowData.findIndex(person => person.status !== 'Completed');

        if (anchorIndex === -1) {
             // Round is over
             if (nextRound === 1) {
                 return { success: true, message: "Selection recorded. Round 1 complete — awaiting lottery setup." };
             } else {
                 nextRound++;
                 configSheet.getRange("B2").setValue(nextRound);
                 let rows = turnSheet.getDataRange().getValues();
                 rows.shift();
                 rows.forEach((row, index) => {
                     turnSheet.getRange(index + 2, statusIdx + 1).setValue('Waiting');
                 });
                 continue; // re-evaluate for the new round
             }
        }

        let skippedSomeone = false;
        for (let offset = 0; offset < 3; offset++) {
            const personIndex = anchorIndex + offset;
            if (personIndex < queueWindowData.length) {
                let person = queueWindowData[personIndex];
                if (person.status !== 'Completed' && person.skipNextTurn === true) {
                    // Consume skip
                    let originalDataRow = turnData.findIndex(row => row[nameIdx] === person.name);
                    turnSheet.getRange(originalDataRow + 2, skipIdx + 1).setValue(false);
                    turnSheet.getRange(originalDataRow + 2, statusIdx + 1).setValue('Completed');
                    skippedSomeone = true;
                }
            }
        }

        if (!skippedSomeone) {
            break; // stable state
        }
    }

    return { success: true, message: "Selection recorded." };
  } catch (e) {
    return { success: false, message: "An error occurred: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function testQueueWindowBehavior() {
    // 1, 2, 3 Active
    const headers = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
    const turnDataRaw = [
        headers,
        ['Person1', '1234', 1, 'Waiting', 0, 1, false],
        ['Person2', '1234', 2, 'Waiting', 0, 2, false],
        ['Person3', '1234', 3, 'Waiting', 0, 3, false],
        ['Person4', '1234', 4, 'Waiting', 0, 4, false],
        ['Person5', '1234', 5, 'Waiting', 0, 5, false],
    ];

    let computed = calculateQueueWindow(turnDataRaw, 1);
    if (computed[0].computedStatus !== 'Active' || computed[1].computedStatus !== 'Standby' || computed[2].computedStatus !== 'Backup') {
        throw new Error("Initial window incorrect");
    }
    if (computed[3].computedStatus !== 'Waiting') throw new Error("Person 4 should be waiting");

    // Person 2 completes
    turnDataRaw[2][3] = 'Completed';
    computed = calculateQueueWindow(turnDataRaw, 1);

    if (computed[0].computedStatus !== 'Active') throw new Error("Person 1 should be Active");
    if (computed[1].computedStatus !== 'Completed') throw new Error("Person 2 should be Completed");
    if (computed[2].computedStatus !== 'Backup') throw new Error("Person 3 should be Backup");
    if (computed[3].computedStatus !== 'Waiting') throw new Error("Person 4 should be Waiting");

    // Person 1 completes
    turnDataRaw[1][3] = 'Completed';
    computed = calculateQueueWindow(turnDataRaw, 1);

    if (computed[2].computedStatus !== 'Active') throw new Error("Person 3 should be Active");
    if (computed[3].computedStatus !== 'Standby') throw new Error("Person 4 should be Standby");
    if (computed[4].computedStatus !== 'Backup') throw new Error("Person 5 should be Backup");

    console.log('PASS: testQueueWindowBehavior');
}

function testQueueWindowSkipNextTurn() {
    const headers = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
    const turnDataRaw = [
        headers,
        ['Person1', '1234', 1, 'Waiting', 0, 1, false],
        ['Person2', '1234', 2, 'Waiting', 0, 2, true], // Skipping
        ['Person3', '1234', 3, 'Waiting', 0, 3, false],
        ['Person4', '1234', 4, 'Waiting', 0, 4, false],
    ];

    let computed = calculateQueueWindow(turnDataRaw, 1);
    if (computed[0].computedStatus !== 'Active') throw new Error("Person 1 should be Active");
    if (computed[1].computedStatus !== 'Waiting') throw new Error("Person 2 should be Waiting (skipped)");
    if (computed[2].computedStatus !== 'Backup') throw new Error("Person 3 should be Backup");
    if (computed[3].computedStatus !== 'Waiting') throw new Error("Person 4 should be Waiting (hole is left)");

    console.log('PASS: testQueueWindowSkipNextTurn');
}
function runTests() {
  testQueueWindowBehavior();
  testQueueWindowSkipNextTurn();
}

function testRound1SelectableWeeks() {
  const rows = [
    [new Date('2026-06-01'), ' Prime ', '', '', '', '', 4],
    [new Date('2026-10-05'), 'non-prime', '', '', '', '', 3]
  ];

  const weeks = buildAvailableWeekData(rows);

  const primeCount = weeks.filter(
    week => week.classification === 'Prime'
  ).length;

  const nonPrimeCount = weeks.filter(
    week => week.classification === 'Non-Prime'
  ).length;

  if (primeCount !== 1) {
    throw new Error('Expected one selectable Prime week.');
  }

  if (nonPrimeCount !== 1) {
    throw new Error('Expected one selectable Non-Prime week.');
  }

  console.log('PASS: Round 1 includes Prime and Non-Prime options.');
}

function testInvalidClassification() {
  const rows = [
    [new Date('2026-06-01'), ' Prime ', '', '', '', '', 4],
    [new Date('2026-10-05'), 'UnknownType', '', '', '', '', 3]
  ];

  const weeks = buildAvailableWeekData(rows);
  const invalid = weeks.filter(w => w.classification === null);

  if (invalid.length !== 1) {
    throw new Error('Expected one invalid classification.');
  }

  console.log('PASS: Invalid classification properly detected.');
}

function runMoreTests() {
  testRound1SelectableWeeks();
  testInvalidClassification();
}
