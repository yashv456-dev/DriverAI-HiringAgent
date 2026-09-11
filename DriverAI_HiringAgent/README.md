# HiringAgent P1 and P2

Reviewed against the local source and configuration on September 7, 2026. Cloud deployment state was not checked.

- [Current architecture and implementation gaps](CURRENT_ARCHITECTURE.md)
- [Client setup and handoff](HiringAgent_P2/CLIENT_COMPUTER_START_HERE.md)
- [P1 installation and operations](HiringAgent_P1/docs/P1_RUNBOOK.md)
- [P2 documentation index](HiringAgent_P2/docs/README.md)

P1 captures applications in Power Automate and saves resumes and candidate rows to SharePoint. P2 reads that Excel queue, extracts and scores resumes using Ollama, and publishes candidate results. SQLite support exists locally, but the independent database-backed intake design is incomplete.

The September 7 documentation cleanup removed the superseded SharePoint List roadmaps, P1 TDD, old P2 summary, operator prompt, cloud deployment guide, and duplicate P2 deployment runbook. Recent design/implementation notes, benchmark evidence, audit reports, model handoff files, and operational guides remain. Application code, tests, configuration, flow packages, databases, resumes, and results were preserved.
