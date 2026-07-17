# Step 1 — Define helper functions

- Get-UrgencyLevel — takes $DaysRemaining and returns EXPIRED (< 0 days), CRITICAL (≤ 30), WARNING (≤ 60), NOTICE (≤ 90), or OK.
- Get-UrgencyColor — maps each urgency level to a console color (Red for EXPIRED/CRITICAL, Yellow for WARNING, Cyan for NOTICE, Green for default/OK).
- Write-ConsoleSummary — filters results down to anything not OK; if nothing's flagged, prints a clean "all good" message and returns. Otherwise -prints a report header (scan date + flagged count), then loops through EXPIRED → CRITICAL → WARNING → NOTICE, printing each group's items (sorted by days remaining) with app name, credential type, days label, expiry date, and name.

<img width="870" height="850" alt="Screenshot 2026-07-16 at 22 02 31" src="https://github.com/user-attachments/assets/19f0937f-92f6-45b6-8b6b-2d5525e5ac77" />


