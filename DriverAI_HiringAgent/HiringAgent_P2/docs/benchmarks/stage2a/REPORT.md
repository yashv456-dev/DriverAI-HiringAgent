# Stage 2A extraction measurements — 2026-09-06

The fixed 3,628-character CV's median smart extraction fell from **5.621s to
3.023s (46.2%)**. All three final field dictionaries equal their corresponding
baseline dictionaries. No measured call timed out or reached the output cap.

For the current implementation, read [CURRENT_ARCHITECTURE.md](../../../../CURRENT_ARCHITECTURE.md). This retained September 6 measurement covers extraction latency only; the older migration roadmap has been retired.

The starting workspace already had `timeout: 180` and `require_ai: true`.
The budget was retained; only its explanation was updated, after steps 2–4.
Consequently the budget step measures an unchanged setting, not a 60-to-180
performance improvement. The earlier 62.1s client measurement still motivates
180s as approximately three times that duration. Scoring remains at 150s.

## Method and reproducibility

- Live local endpoint: `http://localhost:11434`, Ollama **0.33.3**,
  `qwen3-1.7b-p2:latest`.
- Loaded model digest:
  `338d99fd07ec88391dae70f0b1227f80bcecc2f574d6ba1e387e0d889d8176d0`.
- Input: [resume.txt](resume.txt), a synthetic, ordinary two-page software CV.
  It contains an explicit email address and “6 years of experience.” No applicant
  records were fetched. This is not the earlier Doreen CV or a hardware-matched
  reproduction of the previous 62.1s result.
- Three consecutive `extract_candidate_details_smart` runs at each stage.
  The median includes deterministic parsing and result merging. No warm-up call
  was discarded; the first baseline run includes greater loading/prompt overhead.
- Then one call each to `ai_recheck_fields`, `infer_looking_for_role`, and
  `infer_missing_portfolios`, using the same resume. Auxiliary inputs are fixed
  from the baseline, so their request content is comparable across stages.
- Role inference and portfolio inference were explicitly exercised even though
  the normal pipeline could skip them. This is a per-call benchmark, not a
  pipeline throughput measurement. The 106-role scoring call was not exercised.
- Measurements were sequential against Ollama. There were no tenant calls,
  tenant mutations, or messages. No end-to-end tenant run was needed.

Run from P2 using a **new** stage directory name:

```powershell
.venv/Scripts/python.exe docs/benchmarks/stage2a/benchmark.py another_run
```

The harness retains request payloads, raw responses, token counts, stop reasons,
model/template metadata, per-run fields, and source snapshots under each stage.
`auxiliary_fields.json` fixes the auxiliary-call inputs for repeat runs.

## Timings, measured before and after each step

All values are seconds. Extraction is the median of three full smart-extractor
calls. Other columns time one HTTP call each, including receipt of the full body.

| Cumulative stage | Extraction median | Recheck | Role note | Portfolios |
|---|---:|---:|---:|---:|
| 1. Baseline | 5.621 | 5.366 | 4.014 | 3.630 |
| 2. `think: false` | 2.844 | 2.924 | 2.487 | 2.493 |
| 3. `num_predict: 512` | 2.886 | 2.895 | 2.450 | 2.473 |
| 4. JSON schemas | 2.987 | 2.953 | 2.352 | 2.459 |
| 5. Retain 180s budget | 3.020 | 2.967 | 2.382 | 2.466 |
| 6. One timeout retry | 3.023 | 2.971 | 2.386 | 2.484 |

| Stage | Extraction run 1 | Run 2 | Run 3 |
|---|---:|---:|---:|
| Baseline | 7.676 | 5.563 | 5.621 |
| Thinking off | 4.374 | 2.841 | 2.844 |
| Bounded | 2.916 | 2.826 | 2.886 |
| Schema | 3.089 | 2.890 | 2.987 |
| Budget | 3.120 | 3.013 | 3.020 |
| Retry | 3.109 | 2.998 | 3.023 |

HTTP-only extraction medians were **5.563s before / 2.965s after**; the small
difference from the headline medians is local parsing and merging. Small changes
between the later stages do not establish a performance effect. The clear win
is disabling thinking: **−2.777s / −49.4%** at that step. Budget and retry do not
accelerate successful requests.

## Thinking, output length, and schemas

The exported `docs/client_handoff/qwen3-1.7b-p2.Modelfile` appends `/no_think`
and prefills an empty `<think>` block. However, `/api/show` revealed that the live
model uses a different template. The raw baseline response contains a separate
`message.thinking` value, reaching **2,434 characters** in the extraction runs.
See [the raw baseline responses](01_baseline/raw.json) and
[the live template snapshot](01_baseline/metadata.json).

Setting `think: false` explicitly on all four extraction requests eliminated
that field's content in every subsequent measured response. The installed model
and exported Modelfile were not changed. The API switch follows
[Ollama's thinking documentation](https://docs.ollama.com/capabilities/thinking).

The largest baseline generation was **666 tokens including thinking**; its
answer was 421 characters. After disabling thinking, the largest observed
generation across all four call types was **121 tokens / 421 answer characters**.
That measurement selected **512 tokens**, over four times the observed maximum.
With schemas, the maximum rose to **124 tokens**, still with 421 answer characters
and more than four times headroom. This is an observed bound from this fixture,
not a claim that every future CV fits within 124 tokens.

Every request now supplies an object schema with required string fields and
`additionalProperties: false`: seven existing candidate fields for extraction
and recheck, `summary` for role inference, and `linkedin/github/other` for
portfolios. This uses
[Ollama's structured-output API](https://docs.ollama.com/capabilities/structured-outputs).
All measured schema responses had exactly the required keys, string values,
`done_reason: stop`, and token counts below 512. Existing `json.loads` exception
handling remains, including for truncated or otherwise malformed responses.

## Field comparison

All three extraction outputs in **every stage** equal their corresponding
baseline output, including fields beyond those requested for this comparison.

| Requested field | Baseline | Final |
|---|---|---|
| `full_name` | Alex Morgan | Same |
| `phone` | +1 (512) 555-0142 | Same |
| `email` | Absent from smart result | Same; not fixed |
| `location` | Austin, Texas | Same |
| `country` | United States | Same |
| `years_exp` | No such result key | Same; actual `experience` is `6 years` |
| `education` | Bachelor of Science in Computer Science, University of Texas at Austin | Same |
| `skills` | Python, FastAPI, Django, React, TypeScript, JavaScript, SQL, PostgreSQL, Docker, AWS, Git, GitHub Actions, Linux, pytest, REST API, HTML, CSS, REST | Same |

The email limitation is pre-existing: the deterministic parser finds the email,
but the smart result does not carry it forward. Neither the model prompt nor
`_ai_fields` requests/returns email. Experience is already extracted
deterministically into `experience`; looking up `years_exp` would show a blank.
No email propagation or result-key changes were made in this latency task.

The raw unconstrained non-thinking first pass omitted `phone`; deterministic
merging preserved the correct phone. The schema restored `phone` to the raw
model response too. The independent recheck and its prompt remain intact.

## Retry and validation

All four extraction calls share a transport helper with at most two attempts.
Only `requests.Timeout` triggers the second attempt. Its `ReadTimeout` and
`ConnectTimeout` subclasses are included, matching existing timeout provenance;
a plain `ConnectionError` is not retried. HTTP status checks and JSON parsing
remain outside the loop, so those failures never trigger a retry. Exhaustion
preserves the existing fallback and `EXTRACTION_SOURCE` classification and logs
both attempts. The budget is **180s per attempt**, so two fully exhausted
attempts can take about 360s. The row-level retry counter is untouched.

| Suite | Untouched baseline | Final |
|---|---:|---:|
| `test_p2.py` | 1393/0 | 1393/0 |
| `test_p1.py` | 706/0 | 706/0 |
| `test_store.py` | **59/0** | **59/0** |
| New `test_extraction_transport.py` | — | 9 tests, all pass |

**Acceptance-count discrepancy:** the supplied workspace already contains six
more store checks than the requested 53. Section J (`require_ai`) accounts for
those six. The suite was not edited or reduced to manufacture the older count.

The new transport tests cover all four calls: successful one-attempt requests,
all three timeout classes followed by recovery, two-timeout exhaustion, connection
failure, HTTP failure, malformed JSON, timeout followed by connection/JSON failure,
and recovered model provenance. Recorded benchmark outputs were also checked for
field equality, schema shape, absence of thinking, and normal completion.

`score_retry_max: 3`, `row_delay_seconds: 2`, `scoring_timeout: 150`, the scoring
prompt, geo functions, and the column contract were preserved. Parsed config
values equal the initial config exactly. P2 `suppress_emails` remains true and
P1 `send_applicant_emails` remains false. SharePoint migration and grounded
extraction were not started.

Application changes are reviewable in [extraction.py.diff](extraction.py.diff)
and [config.yaml.diff](config.yaml.diff). Baseline and final suite logs and all
six measurement stages are retained alongside this report.
