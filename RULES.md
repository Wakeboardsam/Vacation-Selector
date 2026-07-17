# Vacation Selection Rules

## Setup & Administrator Steps
1. **Merge/Deploy:** Merge the PR and deploy the Google Apps Script code.
2. **Schema Migration:** Run the `setupSpreadsheetSchema()` function from the Google Apps Script editor. This safely migrates the old `QueuePosition` to `SeniorityPosition` and adds `LotteryPosition` and `SkipNextTurn` columns.
3. **Verify Seniority:** Confirm the `SeniorityPosition` column is correct for Round 1 (1 to 21).
4. **Lottery Setup:** Enter one unique `LotteryPosition` for every participant in the spreadsheet.
5. **Initialize Lottery Round:** After Round 1 (the seniority round) is fully complete for all participants, an administrator MUST run `initializeLotteryRound()` from the Apps Script editor. This will advance the round to 2 and setup the new Active window.
6. **Confirm Window:** Confirm that Round 2 has `Active`, `Standby`, and `Backup` statuses assigned.

## Selection Order
* **Round 1 (Seniority Round):** Participants select exactly one week in seniority order.
* **Remaining Rounds (Lottery Rounds):** Participants select in a separately maintained lottery order, snaking down and up the list on alternating passes. Round 2 goes Top-to-Bottom. Round 3 goes Bottom-to-Top, etc.

## Week Hierarchy (Outcomes)
* **PRIME_ONE:** A participant selects exactly one prime week. They advance normally.
* **NONPRIME_ONE:** A participant selects exactly one non-prime week. They advance normally.
* **NONPRIME_TWO:** A participant selects two non-prime weeks. The participant advances normally for the current round, but their `SkipNextTurn` flag is set to true. The next time they enter the active window, the flag is cleared and their turn is skipped for that round.

## Capacity & Validation
* **Capacity Limit:** Each vacation week has a maximum capacity of four participants.
* **Double Booking:** A participant cannot be recorded twice in the same week.
* **Simultaneous Picks:** A participant cannot pick two identical weeks in the same turn submission.
* **Prime Rule:** If a participant selects two weeks in one turn, both must be non-prime. Prime weeks can only be selected alone.
* **Selection Complete:** When all slots in all weeks are completely full, the queue is cleared, and users are notified that the selection process has ended. No further turns can be taken.
