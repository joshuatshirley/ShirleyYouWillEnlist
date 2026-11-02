# local-bank — practice questions

## What's here

| Path | What |
|---|---|
| `raw/` | Questions  (Diagnostic, Practice Exams One/Two, AFQT Focus Exam) |
| `images/` | The equation images referenced by `raw/` |
| `verified/` | Checked versions: images transcribed, every answer re-solved, explanations corrected, `status` + `check_note` per item |
| `check/` | Second, blind re-solve of every AR/MK question plus explanation audit |
| `fixes.json` | Final decisions: items excluded as ambiguous, and patches for book errors |
| `build_practice.py` | Validates everything and writes `practice-data.js` |
| `practice-data.js` | What the kiosk loads (`window.__EST_PRACTICE__`) |
| `tips-data.js` | Extra Test-Taking Tips(`window.__EST_TIPS__`). Hand-edited, no build step |

## Counts (2026-11-01)

AR 120 · MK 100 · WK 136 · PC 59 = 415. Five items excluded as ambiguous
(see `fixes.json`). The book's in-print tests are the whole source here;
its three online-only tests aren't in this copy.
