\*\*Project:\*\* TrackerUI (Vanilla HTML/CSS/JS Progressive Web App)  
\*\*Current Task:\*\* Implement Batch Status Updates and Auto-Fill Graceful Degradation for the testing phase.

\#\#\# 1\. Batch Status Update Architecture  
\*   \*\*UI Integration:\*\* Add a permanent checkbox column to the left edge of the main data table dashboard.  
\*   \*\*Action Dock:\*\* Implement a floating action bar at the bottom of the screen that only renders when one or more row checkboxes are ticked. It must contain a status dropdown and an "Apply" button.  
\*   \*\*Data Processing & State Sync:\*\*  
    1\. Iterate over the selected rows and map them using their immutable UUIDv4 keys (located in hidden Column K).  
    2\. \*\*CRITICAL:\*\* Strictly preserve the existing UUIDv4s. Do not overwrite or regenerate a key during a status change, or it will break the historical rollover linking.  
    3\. Update the Request Status (Column J) in the active state array first.  
    4\. Re-render the DOM to reflect the new statuses instantly.  
    5\. Trigger a single \`showDirectoryPicker\` write-back to update the local Excel workbook in one batch operation.

\#\#\# 2\. Auto-Fill Defensive Programming & Mock Data  
\*   \*\*Graceful Degradation:\*\* The \`Master Accounts\` reference sheet has been intentionally removed from the new workbook bootstrap function. Implement an \`if (workbook.Sheets\['Master Accounts'\])\` safety check. If the sheet is missing, do not throw a fatal error. Gracefully degrade the UI by disabling the auto-fill event listeners or greying out the fields so manual typing can continue.  
\*   \*\*Mock Testing Environment:\*\* Create a hardcoded \`mockMasterAccounts.js\` module containing dummy data structured exactly like the real sheet (e.g., generic names like "Auto Import Tbilisi LLC", fake Master/Bidder numbers, and generic yard numbers like 110).   
\*   \*\*Dev Toggle:\*\* Implement initialization logic to check if the app is running on \`localhost\` (or a designated dev state). If true, bypass the Excel sheet lookup and populate the auto-fill dropdowns using the mock array. This allows rigorous testing of the keystroke events and UUID mappings without requiring real company data.

\#\#\# 3\. Future Scope (Parked \- Do not implement yet)  
\*   \*\*Mass Report Filter:\*\* A client-side filter for 380,000-row workbooks utilizing O(1) \`Set\` logic and \`{header: 1}\` lightweight SheetJS parsing to avoid V8 memory crashes.  
\*   \*\*Smart Scratchpad:\*\* A split-pane raw text input that uses real-time Regex to automatically extract 8-digit lot numbers and alphanumeric title tracking numbers into a structured table, with eventual UUID-linking to specific tickets.


Quick additional fixes:
Need to add rollover from previous sheet.

Completed status still rolls over.  

Dark mode still has the text showing white in the dropdown for the workbook selection. Highlights show black text correctly. The background being white causes the text be invisible until highlighted.
