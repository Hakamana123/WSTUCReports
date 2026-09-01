# Re-registration Principles & Templates — reverse-engineered from Grant's AB4 file

Source: `2026 AB4 for 2026 SPR - Reregistration List v1.1` (Grant's own worked
file, 2026-07-07). ~3,600 de-duplicated students across 9 sheets, each carrying
a `Rereg Principles` label, a `Template` label, and the resulting
`Prep / B1–B4 Name` advice.

The system has three layers:

```
student's situation  ──►  Rereg Principle  ──►  Template  ──►  what to register
     (data)               (a label)            (a rule)        (Prep + B1–4)
```

Findings are split by how solid they are:

- **A. Mechanical & verified** — reproduced against the file, match rates given.
- **B. A human judgement call** — Grant's, not a formula. Don't harden it.
- **C. Structure known, details not fully pinned down** — needs the Spring
  offering list (still being scraped) and, for Template 1c, at least one more
  factor that looks like Grant's judgement.

---

## Vocabulary

`All Subjects` bitmask (Grant's notation), `1` = still to pass, `0` = passed:

| Program type | Shape | Groups |
|---|---|---|
| Diploma (7188–7198) | `[p1+p2] + [b1+b2+b3+b4] + [b5+b6] + [e]` | 2 prep · 4 sem-1 blocks · 2 sem-2 blocks · elective count |
| Nursing (9031) | `[b1+b2+b3+b4] + [b5+b6+b7+b8]` | 8 blocks, no prep, no electives |

- **Prep 1** = GEDU0016, **Prep 2** = GEDU0017.
- **Slot _n_** = the subject that runs at position _n_ of the program's pattern
  (slot 1–4 = semester 1, slot 5–6 = semester 2 for diplomas).
- **Cohort age** (from `Start Semester`):
  - **current** = started `26 AUT`, `25 SPR`, or `25 SUM`
  - **old** = started `25 AUT` or `24 …`

---

## A. Mechanical & verified

### A1. `Sem 1 Subject Count`  (100% match)

```
Sem 1 Subject Count = (Prep 1 still to pass ? 1 : 0) + (# of slots 1–4 still to pass)
```

i.e. how many of the first-semester subjects (GEDU0016 + the four block
subjects) the student still owes. It is **not** affected by prep 2, slots 5–6,
or electives.

### A2. Principle → Template  (1-to-1)

| Rereg Principle | Template |
|---|---|
| On Pattern | Template 1a |
| Mostly Progressing (diploma) / Stay with Cohort (nursing) | Template 1c |
| Unsatisfactory progress in S1 (diploma) / Start Again (nursing) | Template 1b |
| 3+ Sessions | Template 2 |
| Overall lack of success | Template 3 |
| Complete | Transition |

### A3. Choosing the Principle

```
if nothing outstanding:                              → Complete            (Transition)

elif cohort age == current:
    if Sem 1 Subject Count == 0:                      → On Pattern          (1a)
    elif Sem 1 Subject Count is 1–2:                  → Mostly Progressing / Stay with Cohort  (1c)
    else  (Sem 1 Subject Count ≥ 3):                  → Unsatisfactory progress in S1          (1b)
        └─ nursing + no Spring subjects available:    → Start Again          (1b)

elif cohort age == old:
    if total still-to-pass ≤ 6:                       → 3+ Sessions          (2)
    else (total still-to-pass ≥ 7):                   → Overall lack of success (3)
```

Verified boundaries:
- **On Pattern**: every On-Pattern student has `Sem 1 Subject Count == 0`
  (1,132/1,132 current-cohort diplomas).
- **Old-cohort split at total-outstanding 6 / 7**: 3+ Sessions runs 1–6
  (max 6), Overall lack of success runs 7–10 (min 7). No overlap.

### A4. Template 3 sub-rule — Conditional Enrolment cap  (clean on standing)

Within Template 3, the shape depends on progression standing:

| Standing | What Template 3 registers |
|---|---|
| **Conditional Enrolment** | **minimal**: Prep 1 + slot 1 only (≈25–30cp — the CE cap) |
| At Risk / blank / Good Standing | **full semester-1 restart** (same as 1b) |

82/82 of the "minimal" Template-3 rows are Conditional Enrolment; the 179
"restart" rows are not. This confirms the 30cp CE cap already in the tool —
it's the same mechanism, showing up here as a modifier on the template rather
than a template of its own.

### A5. Prep rule  (consistent across every template)

```
Prep 1 (GEDU0016) still to pass   → advise GEDU0016
else Prep 2 (GEDU0017) still to pass → advise GEDU0017
else                              → no prep
```

(A rare Template-2 variant advises "GEDU0016 and GEDU0017" together — 20 of 84
both-preps-outstanding rows. Treat as an exception, not the rule.)

### A6. Transition  (100% match)

`Complete` / Template Transition → register nothing in every slot. The student
has finished the diploma and moves to their bachelor degree.

### A7. Template 1a  (95% match — "On Pattern")

Student is on cohort pace (`Sem 1 Subject Count == 0`), 1–5 subjects ahead.

```
Prep    = prep rule (usually GEDU0017)
B1, B2  = slots 5, 6  (whichever are still to pass)
B3, B4  = "+1 elective" each, down to the elective count
```

Nursing 1a: B1–B4 = slots 5, 6, 7, 8.

---

## B. A human judgement call — NOT a formula

### The Mostly Progressing ↔ Unsatisfactory boundary at `Sem 1 Subject Count == 2`

- `Sem 1 Subject Count` ≤ 1  → Mostly Progressing, always.
- `Sem 1 Subject Count` ≥ 3  → Unsatisfactory, always.
- **`Sem 1 Subject Count` == 2** → ~82% Mostly Progressing, ~18% Unsatisfactory,
  and nothing in this file separates the two cleanly (not total outstanding,
  not which blocks failed, not standing).

This is Grant weighing each borderline student by hand. The rule
"`== 2` → Mostly Progressing" is a reasonable default but it is a **default,
not the principle**. This is exactly why the template matters more than the
principle: the template, once the principle is set, is mechanical; the
principle has one soft edge that needs Grant.

---

## C. Structure known — details not fully pinned down

These templates place subjects into blocks by slot position, but a naïve
reproduction only lands ~50–70%. Two things are missing:

1. **The Spring offering list.** Some subjects genuinely never run in Spring and
   are always skipped — e.g. `ENGR1053` (7193 slot 4) appears as advised **0**
   times in the whole file. This list is still being scraped from the handbook.

2. **At least one more factor in Template 1c's B3/B4 (unexplained).** Whether a
   failed semester-1 subject gets caught up into B3/B4 or deferred (block → an
   elective) is **not** cleanly predictable. Example, program 7189, slot 4
   (`HLTH1021`) still to pass:
   - `[1+1] + [0+0+0+1] + [1+1] + [2]` → B3/B4 = elective, elective  (deferred)
   - `[0+1] + [0+0+0+1] + [1+1] + [2]` → B3/B4 = elective, HLTH1021   (caught up)

   The only visible difference is prep 1 status, which shouldn't matter to a
   block slot — so this is most likely another of Grant's per-student calls.
   `HLTH1021` itself is advised 108 times elsewhere, so "not offered" is **not**
   the explanation here.

### Template 1b — "Unsatisfactory progress in S1" / "Start Again"

Repeat semester 1.

```
Prep    = prep rule (usually GEDU0016)
B1–B4   = slots 1, 2, 3, 4
          - slot still to pass  → that subject
          - slot already passed → "+1 elective" if electives still needed, else blank
```

Nursing "Start Again" (no Spring subjects available): register nothing now,
restart next Autumn.

### Template 1c — "Mostly Progressing" / "Stay with Cohort"

Stay with the cohort; *sometimes* catch up a miss around it.

```
Prep    = prep rule
B1, B2  = slots 5, 6  (cohort's current subjects)
B3, B4  = SOMETIMES a failed slot-1–4 subject, SOMETIMES an elective
          (see C2 — the choice isn't cleanly predictable from this file)
```

Nursing 1c: B1–B4 = slots 5–8; the 1–2 early misses wait for next Autumn.

### Template 2 — "3+ Sessions"

Old cohort, little left, lots of runway — deliberately light.

```
Prep    = prep rule
Blocks  = the outstanding subjects that are offered in Spring, in slot order;
          a lone outstanding elective goes to B3
          - many students get ONLY the prep, or only one elective
```

### Template 3 (non-CE) — "Overall lack of success"

Old cohort, failed almost everything → full semester-1 restart, same shape as 1b.
(CE students under Template 3 → the minimal cap, see A4.)

---

## How this maps onto the current tool

| Current tool | Grant's system |
|---|---|
| "on-track" path | Template 1a ✓ (matches 95%) |
| "cohort-first, backlog displaces electives" | Template 1c — right idea, but Grant defers the catch-up more often than the tool does (C2) |
| CE 30cp cap → 3 subjects, prep to Summer | Template 3 / CE modifier ✓ (A4) |
| "nothing outstanding" | Transition ✓ |
| earliest-outstanding-first for old cohorts | splits into Template 2 vs 3 by total load ✗ (not yet) |
| — | Template 1b (restart semester 1) ✗ (not yet) |
| — | the Principle label itself ✗ (not yet surfaced) |

**Next step when the Spring offering list lands:** rebuild the advice around
"pick the Principle (A3) → apply the Template (A7 / C)", and emit the Principle
label as a column so a coach sees *why* each plan looks the way it does.
