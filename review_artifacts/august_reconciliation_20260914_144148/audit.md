# August 2026 Live Reconciliation

Generated: 2026-09-14T14:42:08

## P1 and P2 masters

- P1 CandidateList has 29 candidates total: 27 August and 2 September.
- P1 Rejected has zero candidate records. Its one data-body row contains only the calculated Resume Link formula/spacer.
- P2 Master has 27 August candidates: 16 on CandidateList (15 Scored, 1 Needs Review - Location Confirmation) and 11 on Rejected (all Rejected - Non-USA Location).

## Resume folders

- `/Candidate_Resumes/2026/August` contains 29 non-empty resume files.
- All 16 exact filenames referenced by P2 Master's August CandidateList exist and download successfully.
- All 11 exact filenames referenced by P2 Master's August Rejected sheet exist and download successfully, but they are still in the August folder.
- `/Candidate_Resumes/2026/Rejected` contains 100 historical files and zero of the exact filenames for the current 11 P2 Master rejected rows.
- The August folder therefore is not an exact accepted-only folder: it contains 16 current main files, 11 current rejected files, and 2 files with no active master row.
- Extra files: `Jeffrey_Kennedy_MC1A.pdf` and `Pragya_Mittal_MDBA.pdf`. Both are non-empty and downloadable.

## Missing active candidate rows

- `APP-20260813-1733-MC1A` ? Jeffrey John Kennedy ? Scored ? `jeffreysajjlins@gmail.com`.
- `APP-20260805-0805-MDBA` ? Pragya Mittal ? Scored ? `pragyamittal926@gmail.com`.
- Both records still exist in `candidates.db` with `active=0`, and both resumes remain in the August folder.

## August Archive

- 38 messages are in Archive for the UTC calendar month used by `trigger_reset.py --month 2026-08`.
- All 38 are read; zero are unread, so there is no August P1 failure message left waiting.
- 30 messages contain a PDF/DOCX resume, representing 29 unique senders. Vibha sent two resume messages and correctly maps to one candidate.
- 27 unique resume senders have a current P1/P2 row. The two without rows are Pragya Mittal and Jeffrey John Kennedy.
- The other 8 messages contain no resume. Five are text-only replies whose original resume emails remain safely unread in earlier Archive months: Mukesh Sharma, Muhammad Usama Bilal, Aakash Gangji, Rajat Gade, and Nikhil Teja Rudraram. They require the older-month replay before appearing in the new masters.

## Client result sheet

- The live `Candidate_List_Results.xlsx` has 15 candidates.
- P2 Master currently has 16 active Scored candidates. Renu Jaitly (`APP-20260821-0812-MCQA`) is missing from the result sheet because the latest publication was blocked by the workbook lock.
- Vibha Swaminathan's result row (`APP-20260805-0931-PMCA`) incorrectly contains Shalin Edward's name, profile, roles, and resume link. P2 Master itself has Vibha and Shalin correctly separated.
- Jeffrey and Pragya are also absent because their database records are inactive and they no longer appear in P2 Master.
- After restoring Jeffrey and Pragya, the complete scored population is 18, subject to the normal result-export eligibility rules.
