651:function processSelection(selectionData) {
652-  const ss = SpreadsheetApp.getActiveSpreadsheet();
653-  const turnSheet = ss.getSheetByName('Turn Management');
654-  const weekSheet = ss.getSheetByName('Week Availability');
655-  const configSheet = ss.getSheetByName('Config');
656-  const lock = LockService.getScriptLock();
657-  lock.waitLock(30000);
658-  try {
659-    const turnDataRaw = turnSheet.getDataRange().getValues();
660-    const currentRound = configSheet.getRange("B2").getValue();
661-
662-    // Schema validation
663-    const schemaCheck = validateSchema(turnDataRaw, currentRound);
664-    if (!schemaCheck.valid) { return { success: false, message: "System setup error: " + schemaCheck.message }; }
665-
666-    // Check if system is completely full
667-    const weekDataRaw = weekSheet.getDataRange().getValues();
668-    const weekData = weekDataRaw.slice(1);
669-    let totalSpotsRemaining = 0;
670-    weekData.forEach(row => {
671-        let empty = 0;
672-        for(let i=2; i<=5; i++) { if (!row[i]) empty++; }
673-        totalSpotsRemaining += empty;
674-    });
675-    if (totalSpotsRemaining === 0) {
676-        return { success: false, message: "Selection Complete: All available vacation slots are filled." };
677-    }
678-
679-    const turnHeaders = turnDataRaw[0];
680-    const turnData = turnDataRaw.slice(1);
681-
682-    const nameIdx = turnHeaders.indexOf('Name');
683-    const statusIdx = turnHeaders.indexOf('Status');
684-    const weeksSelectedIdx = turnHeaders.indexOf('WeeksSelected');
685-    const senPosIdx = turnHeaders.indexOf('SeniorityPosition');
686-    const lotPosIdx = turnHeaders.indexOf('LotteryPosition');
687-    const skipIdx = turnHeaders.indexOf('SkipNextTurn');
688-
689-    const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);
690-    const userObj = queueWindow.find(p => p.name === selectionData.name);
691-
692-    if (!userObj) { return { success: false, message: "User not found." }; }
693-
694-    if (!['Active', 'Standby', 'Backup'].includes(userObj.computedStatus)) {
695-        return { success: false, message: "It is not your turn to make a selection." };
696-    }
697-
698-    // We also need the userRowIndex in the original turnData array
699-    const userRowIndex = turnData.findIndex(row => row[nameIdx] === selectionData.name);
700-
701-    if (!selectionData.week1) {
702-        return { success: false, message: "Missing primary week selection." };
703-    }
704-
705-    if (currentRound === 1 && selectionData.week2) {
706-        return { success: false, message: "Invalid selection. You can only select exactly ONE week during Round 1." };
707-    }
708-
709-    if (selectionData.week1 === selectionData.week2) {
710-        return { success: false, message: "You cannot select the same week twice in one submission." };
711-    }
712-
713-    let w1Index = weekData.findIndex(row => row[0].getTime() == selectionData.week1);
714-    let w2Index = selectionData.week2 ? weekData.findIndex(row => row[0].getTime() == selectionData.week2) : -1;
715-
716-    if (w1Index === -1 || (selectionData.week2 && w2Index === -1)) {
717-        return { success: false, message: "One of the selected weeks does not exist." };
718-    }
719-
720-    let w1Data = weekData[w1Index];
721-    let w2Data = selectionData.week2 ? weekData[w2Index] : null;
722-
723-    // Validate classifications tightly on the backend before making any writes
724-    const class1 = normalizeClassification(w1Data[1]);
725-    const class2 = w2Data ? normalizeClassification(w2Data[1]) : null;
726-
727-    if (!class1 || (w2Data && !class2)) {
728-        return { success: false, message: "Selected week has an invalid classification." };
729-    }
730-
731-    if (w2Data && (class1 !== "Non-Prime" || class2 !== "Non-Prime")) {
732-        return { success: false, message: "Two-week selections must both be Non-Prime." };
733-    }
734-
735-    // Check max capacity and existing spots
736-    let w1EmptySlots = 0;
737-    for (let i=2; i<=5; i++) { if (!w1Data[i]) w1EmptySlots++; }
738-    if (w1EmptySlots === 0) return { success: false, message: "Primary week is full." };
739-
740-    if (w2Data) {
741-        let w2EmptySlots = 0;
742-        for (let i=2; i<=5; i++) { if (!w2Data[i]) w2EmptySlots++; }
743-        if (w2EmptySlots === 0) return { success: false, message: "Secondary week is full." };
744-    }
745-
746-    // Check double booking in the same week
747-    for(let i=2; i<=5; i++){
748-        if (w1Data[i] === selectionData.name) return { success: false, message: "You are already booked for the primary week." };
749-        if (w2Data && w2Data[i] === selectionData.name) return { success: false, message: "You are already booked for the secondary week." };
750-    }
751-
752-    // If validation passes, apply changes atomically.
753-    let w1TargetCol = -1;
754-    for(let i=2; i<=5; i++){
755-        if(!weekSheet.getRange(w1Index + 2, i + 1).getValue()) {
756-            w1TargetCol = i + 1;
757-            break;
758-        }
759-    }
760-    if (w1TargetCol === -1) throw new Error("Concurrency error: primary week filled up.");
761-
762-    let w2TargetCol = -1;
763-    if (selectionData.week2) {
764-        for(let i=2; i<=5; i++){
765-            if(!weekSheet.getRange(w2Index + 2, i + 1).getValue()) {
766-                w2TargetCol = i + 1;
767-                break;
768-            }
769-        }
770-        if (w2TargetCol === -1) throw new Error("Concurrency error: secondary week filled up.");
771-    }
772-
773-    // Perform writes
774-    weekSheet.getRange(w1Index + 2, w1TargetCol).setValue(selectionData.name);
775-    weekSheet.getRange(w1Index + 2, 7).setValue(w1EmptySlots - 1);
776-    totalSpotsRemaining -= 1;
777-
778-    if (selectionData.week2) {
779-        weekSheet.getRange(w2Index + 2, w2TargetCol).setValue(selectionData.name);
780-        let currentSpots = (w2Data ? (4 - (w2Data.filter((_, idx) => idx >= 2 && idx <= 5 && w2Data[idx]).length)) : 0);
781-        let newW2Spots = Math.max(0, currentSpots - 1);
782-        weekSheet.getRange(w2Index + 2, 7).setValue(newW2Spots);
783-        totalSpotsRemaining -= 1;
784-    }
785-
786-    // Update Turn Sheet for current user
787-    const weeksPickedCount = selectionData.week2 ? 2 : 1;
788-    const currentWeeksSelected = turnData[userRowIndex][weeksSelectedIdx];
789-    turnSheet.getRange(userRowIndex + 2, weeksSelectedIdx + 1).setValue(currentWeeksSelected + weeksPickedCount);
790-    turnSheet.getRange(userRowIndex + 2, statusIdx + 1).setValue('Completed');
791-
792-    if (selectionData.week2) {
793-        turnSheet.getRange(userRowIndex + 2, skipIdx + 1).setValue(true);
794-    }
795-
796-    // If we just filled the last spot in the whole sheet, clear queue and exit early
797-    if (totalSpotsRemaining === 0) {
798-        return { success: true, message: "Selection recorded. Selection process is now complete." };
799-    }
800-
801-    // Since authorization relies on computed queue window and 'Completed' status,
802-    // we don't strictly need to write 'Waiting', 'Active', etc. to the sheet anymore
803-    // except for skips. Let's consume skips that fall within the *new* window.
804-    // Wait, the rule says: "Consume skips only under the script lock during state transitions."
805-    // If a person inside the new 3-person window has SkipNextTurn = true, we consume it and mark them Completed.
806-
807-    let loopGuard = 0;
808-    let nextRound = currentRound;
809-
810-    while (loopGuard < 100) {
811-        loopGuard++;
812-        let currentTurnDataRaw = turnSheet.getDataRange().getValues();
813-        let queueWindowData = calculateQueueWindow(currentTurnDataRaw, nextRound);
814-
815-        // Find if anyone in the new window (offset 0, 1, 2 from anchor) has skipNextTurn = true
816-        let anchorIndex = queueWindowData.findIndex(person => person.status !== 'Completed');
817-
818-        if (anchorIndex === -1) {
819-             // Round is over
820-             if (nextRound === 1) {
821-                 return { success: true, message: "Selection recorded. Round 1 complete — awaiting lottery setup." };
822-             } else {
823-                 nextRound++;
824-                 configSheet.getRange("B2").setValue(nextRound);
825-                 let rows = turnSheet.getDataRange().getValues();
826-                 rows.shift();
827-                 rows.forEach((row, index) => {
828-                     turnSheet.getRange(index + 2, statusIdx + 1).setValue('Waiting');
829-                 });
830-                 continue; // re-evaluate for the new round
831-             }
832-        }
833-
834-        let skippedSomeone = false;
835-        for (let offset = 0; offset < 3; offset++) {
836-            const personIndex = anchorIndex + offset;
837-            if (personIndex < queueWindowData.length) {
838-                let person = queueWindowData[personIndex];
839-                if (person.status !== 'Completed' && person.skipNextTurn === true) {
840-                    // Consume skip
841-                    let originalDataRow = turnData.findIndex(row => row[nameIdx] === person.name);
842-                    turnSheet.getRange(originalDataRow + 2, skipIdx + 1).setValue(false);
843-                    turnSheet.getRange(originalDataRow + 2, statusIdx + 1).setValue('Completed');
844-                    skippedSomeone = true;
845-                }
846-            }
847-        }
848-
849-        if (!skippedSomeone) {
850-            break; // stable state
851-        }
