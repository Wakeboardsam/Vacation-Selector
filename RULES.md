# Vacation Selection Rules

## Selection Order
* **Round 1 (Seniority Round):** Participants select exactly one week in seniority order.
* **Remaining Rounds (Lottery Rounds):** Participants select in a separately maintained lottery order, snaking down and up the list on alternating passes.

## Week Hierarchy (Outcomes)
* **PRIME_ONE:** A participant selects exactly one prime week. They advance normally.
* **NONPRIME_ONE:** A participant selects exactly one non-prime week. They advance normally.
* **NONPRIME_TWO:** A participant selects two non-prime weeks. The participant advances normally for the current round, but their `SkipNextTurn` flag is set to true. The next time they enter the active window, the flag is cleared and their turn is skipped for that round.

## Capacity & Validation
* **Capacity Limit:** Each vacation week has a maximum capacity of four participants.
* **Double Booking:** A participant cannot be recorded twice in the same week.
* **Simultaneous Picks:** A participant cannot pick two identical weeks in the same turn submission.
* **Prime Rule:** If a participant selects two weeks in one turn, both must be non-prime. Prime weeks can only be selected alone.
