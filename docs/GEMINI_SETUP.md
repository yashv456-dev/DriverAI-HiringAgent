# Google Gemini Cloud AI Setup & Verification Guide

This guide walks you through configuring Google Gemini API as the primary AI brain for the DriverAI Hiring Agent.

---

## ⚡ Why Gemini for DriverAI?

The DriverAI evaluation engine uses Gemini for:
1. **4-Step Real-Time Resume Extraction**:
   - Parses unstructured candidate resumes (PDF, DOCX, scanned images) into validated entities: Full Name, Email, Phone, US Geo validation, Core Skills, and Degrees.
2. **105-Role Semantic Matrix Matching**:
   - Scores candidates against all 105 DriverAI technical roles on a strict 0–100 match scale with contextual reasoning.
3. **6-Axis Skill Radar Scoring**:
   - Generates normalized radar scores across: Core Engineering, Distributed Systems, ML/AI, Production Reliability, Leadership, and Communication.
4. **Interactive Recruiter Copilot**:
   - Powers real-time chat on the recruiter dashboard, answering complex queries like *"Find all US candidates with Raft consensus experience"*.

---

## 🔑 Step 1: Obtain a Gemini API Key

1. Navigate to **[Google AI Studio](https://aistudio.google.com/)**.
2. Sign in with your Google account.
3. In the left navigation, click **Get API key**.
4. Click **Create API key** (you can create it in a new project or select an existing Google Cloud project).
5. Copy the generated API key (format: `AIzaSy...`).

---

## ⚙️ Step 2: Configure Environment Variables

Create or edit your `.env` file (either at repository root or in `DriverAI_HiringAgent/HiringAgent_P2/.env`):

```ini
# Google Gemini API Key
GEMINI_API_KEY=AIzaSyYourGeminiApiKeyHere

# Model selection (default: gemini-2.5-flash)
HIRING_GEMINI_MODEL=gemini-2.5-flash

# Master toggle (true by default if key is present)
HIRING_GEMINI_ENABLED=true
```

### Supported Models
| Model Name | Recommendation | Latency / Accuracy |
|---|---|---|
| **`gemini-2.5-flash`** | **Recommended (Default)** | Sub-second extraction with top reasoning accuracy |
| **`gemini-2.0-flash`** | High-performance alternative | Extremely fast processing for high-volume batches |
| **`gemini-1.5-flash`** | Long-context alternative | Ideal for deeply nested, 10+ page executive CVs |

---

## 🧪 Step 3: Test & Verify Connectivity

### Method 1: CLI Diagnostic Tool
Run the built-in doctor diagnostic:
```bash
cd DriverAI_HiringAgent/HiringAgent_P2
python bot.py --doctor
```
Or with Docker:
```bash
docker compose run --rm cli --doctor
```

**Expected Output**:
```text
✓ AI Brain: Gemini Cloud AI (gemini-2.5-flash) (API connected)
✓ Active roles loaded: 105 roles from cache
✓ Storage backend: SQLite (candidates.db)
```

### Method 2: Web Server Health Check
While the web server is running (`http://localhost:8000`), run:
```bash
curl http://localhost:8000/health
```

**Response**:
```json
{
  "status": "healthy",
  "gemini": {
    "configured": true,
    "model": "gemini-2.5-flash"
  },
  "roles_loaded": 105
}
```

---

## 🛡️ Resilient Multi-Tier Fallback Hierarchy

The system never halts if cloud connectivity is interrupted or quotas are exceeded:

```text
┌────────────────────────────────────────────────────────┐
│  Tier 1: Google Gemini API (High-speed Cloud AI)       │
│                ▼ (if offline or no key)                │
│  Tier 2: Local Ollama (Private local LLM weights)      │
│                ▼ (if Ollama offline)                   │
│  Tier 3: Deterministic Scorer (Regex & Keyword Matrix) │
└────────────────────────────────────────────────────────┘
```

1. **Tier 1 (Gemini Cloud AI)**: Default when `GEMINI_API_KEY` is provided.
2. **Tier 2 (Local Ollama)**: Auto-detected at `http://localhost:11434` (e.g. `qwen3:1.7b` or `qwen3-1.7b-p2`).
3. **Tier 3 (Deterministic Scorer)**: Hardened, zero-dependency keyword matcher based on `jd_roles_cache.json`.
