# Setup Instrustions

1. Create a folder structure in GoogleDrive that works best for you. The script uses folder and document IDs, so the actual folder structure isn't important. Mine looks something like this:
- Job Apps (root folder)
    - templates
        - resume template files
        - cover letter template files
    - _Job_App_Entries_ (main spreadsheet)

2. Open _Job_App_Entries_ and delete any example text
3. Above the menu bar, select Extensions -> Apps Scripts (this should open up a new window)
4. Name the project, copy/paste provided code from index.js into code editor and replace example code like file and folder IDs
5. Save and Run (You may need to sign into your Google account and "Review Permissions")
6. Go back to the window with \_\Job_App_Entries\_\ open and refresh the page
7. There should be a new menu item titled "AutoFill Resume" with a "Generate Docs" item inside it.

That's it! Now just add entries. The script automatically skips the rows whose "Cover Letter Link" and "Resume Link" cells aren't empty. If you ever need to regenerate an entry, just delete values in these cells and and run "Generate Docs." 