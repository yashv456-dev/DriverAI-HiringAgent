# Resume availability guards — 2026-09-06

P2 now excludes unchanged completed duplicate retries before applying the batch
limit. It also defers a missing/inaccessible resume without incrementing Retry
Count or rejecting the applicant for a processing error.

The trigger was two stale Main rows for Mona Rawat and Yash Shah. Both also have
completed non-USA rejection rows for the same application and version. Their Main
links return 404. The actual readable PDFs are under different names in Rejected;
the historical row/file mismatch is documented in
[the read-only diagnostic](../stage2a/resume_download_diagnostic.json).

## Code changes

- `_scoring_queue_rows` skips only an unreadable-resume retry with a completed
  Scored/non-USA-rejected copy matching Application ID, email, received date,
  last-updated date, original filename, and application-update count. Missing
  evidence, new submissions, and changed versions remain eligible. No rows are
  deleted by this check.
- `_download_resume_text` raises `ResumeUnavailableError` when an attachment
  cannot be downloaded. The scoring loop defers it and logs the actual error.
  Successfully downloaded files that produce no readable text still use the
  existing parsing/OCR review path. A missing second attachment also defers the
  candidate rather than scoring incomplete input.
- `SharePointClient.download_resume` falls back from the dated folder to the
  legacy root only on 404. Authentication, authorization, throttling, and server
  failures retain their original error instead of being masked by a root lookup.

This change does not repair historical Resume URLs or select older files by
name/email alone. The unavailable originals still require a verified link/file
repair before those applications can be scored. Applicant identity alone does
not establish which resume version belongs to a newer submission.

## Verification

| Suite | Result |
|---|---:|
| `test_p2.py` | 1393 passed / 0 failed |
| `test_p1.py` | 706 passed / 0 failed |
| `test_store.py` | 59 passed / 0 failed (unchanged existing count) |
| `test_resume_availability.py` | 10 tests passed |

New tests exercise completed copies on both sheets, preservation of new and
changed applications, failure to verify completion, correct legacy fallback,
unmasked access/server failures, missing files, unreadable downloaded bytes,
partial attachment sets, and pipeline deferral in both dry-run and mocked live
mode with Retry Count already at 2. The latter asserts no model call, review,
rejection, row write, or retry-count change.

Live verification command:

```powershell
.venv/Scripts/python.exe -u bot.py --score-sharepoint --batch-size 2 --dry-run --ai
```

The process additionally set `HIRING_SUPPRESS_EMAILS=true` and
`HIRING_ERROR_EMAIL=false`; configuration files were not changed.

The [dry run](dry_run_2.log) completed in **41.882s**:

- Six unchanged completed retries were skipped, including Mona and Yash.
  Eligible queue size changed from 19 to 13; no tenant rows were removed.
- `APP-20260804-2208-WQMA` and `APP-20260804-1923-WPKA` were selected next.
  Both also returned 404 for all expected resume names.
- Both were reported as **Deferred - resume unavailable**. Neither reached
  extraction or scoring. Their status and Retry Count were left unchanged.
- Summary: processed 0, rejected 0, errors 0, deferred 2; zero emails sent.
- No tenant mutations were performed. No historical link/file repair was applied.

Production diffs: [sharepoint_scoring.py.diff](sharepoint_scoring.py.diff) and
[sharepoint_client.py.diff](sharepoint_client.py.diff). Scoring prompts, extraction
settings, geo rules, the column contract, and the row-level give-up limit remain
unchanged by this task. No List migration or grounded-extraction rework was started.
