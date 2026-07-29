# Vacation Selection System Expansion

This is the expanded Vacation, Weekend, Holiday, and Transfer Selector system for Google Apps Script.

## High-Level Architecture
- **Apps Script Backend (`Code.gs`):** Serves the HTML, parses sheet data, implements strict queue calculation with serpentine boundary logic, and safely proxies Twilio SMS notifications outside the lock.
- **Frontend (`Index.html`, `Components.html`, `JavaScript.html`):** Renders the dashboard and picker screens without frameworks. Validates against phase-specific rules. Contains an annual rules-acknowledgment gateway.
- **Google Sheets Database:** The single source of truth. Config controls the current phase state. Participant Config is the canonical roster. Turn Management tracks live queue status.

## Migration and Setup
1. **Schema Migration:** Deploy the new Apps Script code and run `setupSpreadsheetSchema()`. It safely migrates `Name`, `PIN`, `PhoneNumber`, `SeniorityPosition`, and `LotteryPosition` into a new `Participant Config` sheet. It preserves all existing Turn Management data so active selections are uninterrupted. It also generates new sheets for `Admin Options`, `Transfer Offers`, `Rules & Tips`, and Holiday coverages.
2. **Fresh Year Empty Setup:** Run `checkNewYearSetupReadiness()` to ensure all past operational data (weeks, weekends, turn statuses, and transfer histories) are cleared by the administrator.
3. **Auto-Fill / Randomize:** Once empty, run `autoFillRandomize(year, year)` to generate M-F vacation dates, Sat-Sun weekend coverage, 6 observed official holidays, and soft holidays. It will randomize Lottery positions while keeping Seniority intact.
4. **Setup Review:** The system enters `SETUP_REVIEW`. The administrator MUST explicitly mark which auto-generated week is **Spring Break** and confirm **Christmas** via the "Special Week" column on the `Week Availability` sheet.
5. **Confirm Setup:** After review, run `confirmSetup()`. The system is now ready.

## Administrator Workflow
Phase transitions are strictly manual after Setup (excluding the auto-transition from Round 1 to Round 2). An admin will run the respective start functions (e.g. `beginSeniorityRound()`, `beginWeekendPhase()`, `beginHolidayVolunteerPhase()`, `beginTransferRound()`).
When direct manual edits are made to assignments on the sheet, the admin must run `refreshReconcileFromSheet()` to synchronize capacities, targets, and queue readiness without ever destroying the manual edits.
**Twilio Warning:** Credentials are intentionally stored in the `Admin Options` tab. Any Google Sheet Editor can view or copy them. They are masked in UI and NEVER returned to the frontend JSON API.

## Participant Workflow
1. **Gateway:** Upon login, if the rules are unacknowledged for the active year, the user must view `Rules & Tips` and select their Holiday Volunteer / Transfer preferences.
2. **Vacation Phase:** R1 is Seniority. R2+ is Lottery serpentine. User picks 1 Prime alone, or 1-2 Non-Primes. Two Non-Primes forfeit the next turn. Strict boundaries apply: At an endpoint, the entire descending window must finish before ascending begins.
3. **Weekend Phase:** User picks 1 First Call position. Consecutive, Adjacency, and Holiday warnings are displayed. If a nearby Holiday is available, they can reserve it simultaneously in one atomic transaction.
4. **Holiday Phase (Volunteer/Mandatory):** "Pass" is allowed only in Volunteer. Mandatory forces selection via a strict 3-tier algorithm based on prior and current holiday counts.
5. **Transfer Phase:** Givers offer multiple assignments as single items into a locked pool. Receivers then claim them via serpentine queue order.

## Testing & Automation
Tests verify strict multi-person directional bounds, concurrency lock wins, API credential masking, auto-fill blocking, and stable deduplicated SMS notification timers (360-min reminder, 720-min admin alert).

### Passed Tests
- `testStrictSerpentineBoundary`: PASS (verified boundary pausing).
- `testTransferOfferLocking`: PASS (verified pool lock checks).
- `testTwilioTokenNonExposure`: PASS (verified API data masking).
- `testQueueWindowBehavior`: PASS (verified Active assignment queue).
