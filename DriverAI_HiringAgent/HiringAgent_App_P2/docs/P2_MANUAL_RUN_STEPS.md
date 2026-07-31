# P2 Client Run Guide

P2 scores only new or genuinely updated candidates waiting in the SharePoint
queue. A normal live run does not rescore candidates already marked `Scored` or
`Rejected`.

The old Windows task `HiringAgent P2 Daily 9AM` is disabled.

## Option A: Run from the Desktop App

This is the recommended method for clients.

### Open the app

1. Go to the Windows desktop.
2. Double-click **DriverAI Hiring Agent P2**.
3. Wait for the **DriverAI Hiring Agent** window to appear.

If the desktop shortcut is unavailable, open File Explorer and double-click:

```text
C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2\Start.bat
```

### Run candidates now

1. On the **Home** screen, find **Batch**.
2. Select or type any positive whole number:
   - `1` processes up to one waiting candidate.
   - `25` processes up to 25 waiting candidates.
   - `50` processes up to 50 waiting candidates.
3. Click **Run Live** once.
4. Leave the app and internet connection open while P2 runs.
5. Wait until the activity log reports `Finished`.
6. Check that the final summary shows `Errors: 0`.

While P2 is running, the Home progress panel shows:

- The current candidate number, such as `Candidate 3 of 25`.
- The candidate Application ID, name, and row number at the start of the run.
- The current activity, such as downloading, extracting, matching, or saving.
- How many candidates remain in the current batch.
- The estimated queue remaining during the run and the refreshed actual count
  when the batch finishes.

The same progress details appear in the activity log and in CMD.

If fewer candidates are waiting than the selected batch size, P2 processes only
the candidates available. Clicking **Run Live** again is safe. When no work is
waiting, the log reports:

```text
No unscored candidates - queue is empty.
```

### What the live run does

- Processes new candidates.
- Processes candidates with a genuinely updated resume or clarification reply.
- Leaves previously scored or rejected candidates untouched.
- Keeps unclear-location candidates on the main sheet for human review.
- Moves only confirmed non-US candidates to the Rejected sheet.
- Sends eligible clarification or decline emails.
- Refreshes SharePoint `Candidate_List_Results.xlsx` after the batch finishes.

Only candidates with final status `Scored` appear in
`Candidate_List_Results.xlsx`.

### Schedule an automatic daily run

1. On **Home**, click **Schedule**.
2. Set **Candidates per run**.
3. Enter a time such as `9am`, `2:30pm`, or `17:45`.
4. Click **Add**.
5. Turn on **Turn on daily schedule**.
6. Close the Schedule window but keep the P2 app open.

The in-app schedule works only while the P2 app is open and the computer is
awake. Closing the app stops scheduled runs. The old Windows 9 AM task remains
disabled.

To stop using the schedule:

1. Click **Schedule**.
2. Turn off **Turn on daily schedule**.
3. Close the Schedule window.

## Option B: Run from Command Prompt

Use these steps when the desktop app is unavailable or when technical staff
need a direct command.

### Open the correct folder

1. Press the Windows key.
2. Type `cmd`.
3. Open **Command Prompt**.
4. Enter this command exactly:

```bat
cd /d C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2
```

The command line should now begin with:

```text
C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2>
```

Do not run the P2 command from `C:\Users\tracy>` because Windows will not find
the P2 virtual environment or `bot.py`.

### Test the connection

This test does not update candidates or send email:

```bat
.\.venv\Scripts\python.exe bot.py --test-sharepoint run
```

Continue when the final line says:

```text
PASS - SharePoint files, workbook, table, and P2 mail permission are ready.
```

### Optional read-only preview

This checks the live queue without writing to SharePoint or sending email:

```bat
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 25 --dry-run
```

Change `25` to any positive whole number.

### Run live

Use the number of candidates wanted for this run:

```bat
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 25
```

Examples:

```bat
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 1
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 10
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 50
```

For a large backlog:

```bat
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 500
```

Leave Command Prompt open until the summary and `Finished` lines appear. Run
the command again if more candidates remain.

During candidate processing, CMD displays lines similar to:

```text
P2 Progress | state=candidate | current=3 | batch=25 | batch_left=22 |
queue_left=39 | row=128 | ref=APP-123 | name=Jane Smith |
activity=Downloading resume
```

After the batch, `state=finish` reports the refreshed number of candidates
still waiting.

### Understand the summary

- `Processed`: candidates completed in this run.
- `Geo review`: unclear locations kept on the main sheet for review.
- `Rejected`: confirmed non-US candidates moved to Rejected.
- `Declines`: decline email results.
- `Info reqs`: clarification email results.
- `Errors`: should be zero.

Press `Ctrl+C` only when a running command must be stopped. Do not use `--watch`
for normal client operation.

## Check the Client Results

After the live run finishes, open:

```text
https://led1234567.sharepoint.com/sites/CandidateList_HiringAgent/Shared Documents/Candidate_List_Results.xlsx
```

The workbook is rebuilt from all current `Scored` rows. Previous scored rows
remain in the result workbook but are not rescored. New, review, rejected, and
processing-error rows are not included.

## Quick Troubleshooting

### Windows says "The system cannot find the path specified"

Run the folder command first:

```bat
cd /d C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2
```

Then run the P2 command again.

### The app does not appear

1. Wait 10 seconds.
2. Check the Windows taskbar for **DriverAI Hiring Agent**.
3. Try the desktop shortcut only once more.
4. If it still does not appear, double-click:

```text
C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2\Start.bat
```

### The summary shows an error

Do not repeatedly click **Run Live**. Keep the log visible and send the final
error lines to technical support.
