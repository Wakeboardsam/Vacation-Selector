# Vacation Selection Rules

## Setup & Schema Requirements
The `Turn Management` sheet MUST contain the following exact columns:
`Name | PIN | SeniorityPosition | Status | WeeksSelected | LotteryPosition | SkipNextTurn`
*   `SeniorityPosition`: Used in Round 1 strictly from top to bottom.
*   `LotteryPosition`: Used in Round 2 and beyond, maintaining a strict alternating snake order. Cannot be empty or duplicate.
*   The script includes a `setupSpreadsheetSchema()` function to safely initialize these columns from the legacy setup without overwriting data.

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
