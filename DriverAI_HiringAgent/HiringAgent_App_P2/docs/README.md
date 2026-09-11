# P2 documentation

Reviewed September 8, 2026. Start with the [current P1/P2 architecture](../../CURRENT_ARCHITECTURE.md), which describes the implemented flow, recent changes, and incomplete SQLite integration.

| Document | Purpose |
|---|---|
| [September 8 reliability review](ARCHITECTURE_REVIEW_20260908.md) | Current local configuration, reproduced integration failures, and completion priorities |
| [Client handoff](../CLIENT_COMPUTER_START_HERE.md) | Environment/model setup, tenant targets, retained historical defects |
| [Manual run steps](P2_MANUAL_RUN_STEPS.md) | Desktop and CLI operation |
| [P1 runbook](../../HiringAgent_P1/docs/P1_RUNBOOK.md) | Power Automate installation and operations |
| [SQLite architecture](P2_SQLITE_ARCHITECTURE_PLAN.md) | Latest target design; partial implementation is not a completed cutover |
| [Implementation checklist](P1_P2_IMPLEMENTATION_STEPS.md) | Acceptance requirements and September 6 containment record |
| [Extraction benchmark](benchmarks/stage2a/REPORT.md) | September 6 measurements, raw data, and source snapshots |
| [Resume availability](benchmarks/resume_availability/REPORT.md) | Missing-file deferral and completed-retry filtering evidence |
| [Frozen row audit](../P2_Logs/audits/master_row_audit_20260906/ROW_BY_ROW_REPORT.md) | Historical source/row reconciliation evidence |
| [Model definition](client_handoff/qwen3-1.7b-p2.Modelfile) | Custom model reconstruction input |

### Desktop GUI (6 tabs)

| Tab | Description |
|---|---|
| **Home** | Run live pipeline, view system health and Ollama status |
| **Job Descriptions** | View, edit, and refresh job descriptions |
| **SharePoint** | Check SharePoint connection and test credentials |
| **Resumes** | Manage downloaded candidate resumes |
| **Candidates** | Search, filter, and review candidate applications |
| **Settings** | Configure system preferences and API keys |

P1 still writes SharePoint Excel; the regular P2 queue still reads it. The factory fallback is Excel, but the checked local `.env` selects SQLite. That selection changes store access without completing the independent intake, recovery, and master-publication design. See the architecture review for the resulting consistency gaps.

Applicant mail is suppressed in the checked local configuration; admin alerts are enabled. A normal live run can change rows, move/rename resumes, and upload results. An empty queue can still trigger maintenance. The scoring `--dry-run` preview is distinct from `--export-results`, which uploads a workbook.

The old List-migration roadmaps, P1 TDD, P2 detailed summary, operator prompt, and older cloud/deployment guides were retired on September 7. Recent evidence and current setup/operations guides remain; use their linked source files for implementation details. Historical benchmark/test counts describe their recorded runs, not a new validation of today's code.
