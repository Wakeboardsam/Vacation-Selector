# Vacation Selection System

A complete integrated system for managing Vacations, Weekend Coverage, Holidays, and Transfers via Google Apps Script and Google Sheets.

## Architecture
The system operates securely out of a single monolithic lock service dispatcher `processSelection` bridging Google Sheets rows dynamically based on calculated `Participant Config` eligibility states.

### Core Flows
1. **Annual Setup and Confirmation**: The `autoFillRandomize` calculates exact bounds (accounting for years stretching into New Year margins naturally) and exact Easter/holiday coordinates dynamically.
2. **Vacation**: Tiered limit-checks based on Config states (Round 1 = `VACATION_SENIORITY`, Round 2+ = `VACATION_RANDOM`).
3. **Weekend Phase / Holiday Phase**: Rotates independently over serpentine order.
4. **Transfer Phase**: Two stage Offer/Receiver lock model.

### Security
Credentials (e.g., Twilio) are safely maintained out of Apps Script properties explicitly into an `Admin Options` sheet which never broadcasts to frontend endpoints. A `Session Token` mechanism is employed masking exact ID rows dynamically.

## Migration
The setup process uses an idempotent `setupSpreadsheetSchema` capable of extracting legacy `Turn Management` states gracefully into `Participant Config` structures without mutating raw Active assignment blocks.

## Admin Options & Warnings
Credentials, particularly Twilio SIDs and Tokens, are now centrally managed under the `Admin Options` sheet.
**WARNING**: These spreadsheet cells are visible to anyone with Editor permissions to the workbook. Tokens are completely abstracted from the frontend REST payloads and cannot be accessed via Developer Tools by Standard users.

### Scheduled Notifications
- Immediate alerts fire on dynamic queue transition outside the `LockService`.
- 360-minute reminders and 720-minute Administrator alerts map to a dedicated `_processScheduledTimers` apps script trigger using the unique Active timestamp boundary for perfect deduplication.
- Ensure a 10-15 minute time-driven trigger is configured inside Google Apps Script mapped to `_processScheduledTimers`.

## Playwright Tests
To verify UI compliance offline:
```bash
npm install
npx playwright test
```
