# DriverAI Hiring Agent (P1 & P2)

Dual-phase AI hiring pipeline and evaluation engine for DriverAI.

- [Current Architecture](CURRENT_ARCHITECTURE.md)
- [Client Computer Setup & Handoff](HiringAgent_P2/CLIENT_COMPUTER_START_HERE.md)
- [P2 Manual Operations Guide](HiringAgent_P2/docs/P2_MANUAL_RUN_STEPS.md)
- [P1 Installation & Runbook](HiringAgent_P1/docs/P1_RUNBOOK.md)
- [P2 Documentation Index](HiringAgent_P2/docs/README.md)

### Overview
- **P1 (Intake)**: Captures applications via Power Automate, filters spam, extracts metadata, and deposits resumes to SharePoint / local staging.
- **P2 (Intelligence & Web App)**: Extracts text via multi-tier PDF/DOCX/OCR ladders, runs candidate scoring via Gemini 3.1 Flash-Lite & Ollama, maps against a 105-role matching matrix, and powers the **GitHub Primer Dark Web Dashboard (`http://localhost:8000`)** with Command Palette, live stepper, and 1-click candidate outreach.

