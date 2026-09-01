# Re-registration advice — clean rebuild (v2)

## Goal
One input file in, same file back out, with five new columns —
`Prep / Block 1-4 Registration Advice` — holding the recommendation, plus an
`Advice Reason` column. The file's own `... Registration` columns (current
enrolment) are left untouched so a coach can compare advice vs actual.

No pattern_table.json, no live-roster merge, no history. Everything needed is in
the one workbook, except the Spring subject-offering list (see Inputs).

---

## Inputs

### 1. The progression file (uploaded)
Sheet `Query1`, header row 1. One row per student. Columns used:

| Column | Meaning |
|---|---|
| `COMMENCEMENT_PERIOD` | when the student started — drives "where is their cohort now" |
| `PROGRAM_CD` | 1 of 13 programs |
| `Prep 1 Status` … `Subject 8 Status` | per slot: `"CODE Completed"` = passed, `1` = **still to pass**, blank = slot not used by this program |
| `Electives Needed` | integer 0/1/2, or `Not Applicable` / `No Elective Required` |
| `Prep/Block 1-4 Registration` | current enrolment — **ignored on read**, overwritten on write |

**Slot → subject code map** is derived from the file itself: for each program,
whatever code sits in "Subject N Status = CODE Completed" for classmates is the
canonical slot-N subject. (Cross-checked: this reproduces the old
pattern_table.json "26 AUT" block exactly.)

### 2. Spring offering list (you are scraping this from the handbook)
Needed shape — per program, which slots run in Spring and in which block:

```
PROGRAM_CD, slot_number, spring_block   # e.g. 7197, 5, 1   /   7197, 6, 2
```
Plus the prep subjects' Spring availability.
Until supplied, the tool falls back to "slots run in pattern order, earliest
unpassed first" and flags every placement as unverified.

---

## Program shapes

| Programs | Prep? | Modular subjects | Session load |
|---|---|---|---|
| 7188–7198 | GEDU0016 (prep 1) + GEDU0017 (prep 2) | 6, in slots 1–6 | prep + 4 modular = 55cp |
| 9031 (Nursing) | none | 8, in slots 1–8 | 4 modular (1 per block); Autumn = 1–4, Spring = 5–8 |
| 9034 (Policing) | none | 2, in slots 1–2 | 1 per block, only 2 subjects total |

1 subject per block, 4 blocks per session. Prep subject runs alongside the whole
session (not in a block).

---

## The cohort clock

"Where is the student's cohort this Spring" = `start_slot`, the pattern position
the cohort reaches this session. Only the mappings checked against real data are
encoded (`COHORT_START_SLOT` in the code):

| COMMENCEMENT_PERIOD | start_slot | note |
|---|---|---|
| 26 - Autumn Block 1 | 5 | did slots 1–4 in Autumn |
| 26 - Autumn Block 3 | 3 | started mid-Autumn, did slots 1–2 |
| 26 - Spring Block 1 / 26 - Spring | 1 | starting now |
| anything starting 2025 or earlier | 1 | continuing student — clears outstanding earliest-first, no caveat |
| any other 2026 string (SC1/WSTC, Summer…) | 1 | flagged "not recognised, sanity-check" |

---

## Building one student's plan

1. **Outstanding list** = slots where status == `1`, mapped to codes, in slot order.
2. **Order** = cohort slots (`slot >= start_slot`) first, then backlog
   (`slot < start_slot`), each group earliest-first.
3. **Fill 4 blocks** walking that order:
   - not offered in Spring → set aside as *carry*
   - blocks already full → set aside as *following session*
   - otherwise → place; count it if it was a backlog slot
4. **Electives**: one `+1 elective` per leftover block, down to
   `Electives Needed` minus the backlog subjects placed. Any electives that
   don't fit → "pushed back" (if backlog displaced them) or "roll to the
   following session" (just a >1-session load).
5. **Prep**: one prep subject per session — the earliest outstanding one that's
   offered; any others → "prep for a later session".
6. **Reason**: one plain-English line assembled from the above.

Never exceed 4 block subjects + 1 prep.

### Progression standing (`Progression Outcome` column)

| Standing | Rule |
|---|---|
| Good Standing / At Risk / blank (new starters) | full load — prep + 4 subjects. At Risk gets a "monitor" note, no restriction. |
| **Conditional Enrolment** | 30cp cap → **3 modular subjects max** (Block 4 left blank), and the prep subject moves to **Summer** (prep runs in every Summer block), so the 30cp is all modular progress. Electives still fill leftover of the 3 slots; a 2-core student gets 1 elective + 1 deferred. |
| Exclusion | no advice — blank columns, reason "not eligible to re-register". |
| Deferred / Leave of Absence | plan still computed, reason prefixed "confirm the student is returning". |

### Worked example — behind student (Josiah's scenario)
7197, 26-Autumn intake, failed slot 3 (EDUC1013), passed 4/5/6, owes prep 2.

- Prep: GEDU0017
- Cohort Spring slots = 5, 6 → EDUC1012 (Blk1), TEAC2067 (Blk2)
- Backlog: EDUC1013 retake → Blk3 (displaces elective 1)
- Blk4: +1 elective
- Reason: *"Retaking EDUC1013 in Block 3 (slot-7 position); 1 elective deferred to a later session."*

---

## Output
- Same workbook, `Query1` sheet.
- Five columns added: `Prep Registration Advice`, `Block 1-4 Registration
  Advice` — subject codes / `"+1 elective"` text / blank.
- Original `... Registration` columns untouched (current enrolment, for compare).
- New column `Advice Reason` — one plain-English sentence.
- Rows with no live cohort and a large backlog still get a best-effort plan +
  a reason flagging them for coach review.

---

## Validation (done, 2026-09-01)
Against the 1,930 `26 - Autumn Block 1` rows that already have a Spring
enrolment in the file:

- Block 1 subject matches existing enrolment: **95%**
- Block 2 subject matches: **95%**
- Prep matches: **76%** — the gap is almost entirely students who owe *both*
  prep subjects: the tool advises GEDU0016 (prep 1) first, coaches often
  enrolled GEDU0017. **Open question for Josiah.**

Blocks 3–4 diverge by design (elective choice is the coach's).

## Decisions (Josiah, 2026-09-01)
- **Prep order when both outstanding:** GEDU0016 (prep 1) first, GEDU0017 later.
  (Tool already does this; the 76% is messy real enrolment differing, not a bug.)
- **9034 electives:** ignore `Electives Needed` for 9034 — it's a 2-subject
  program with no elective structure. Handled via `NO_ELECTIVE_PROGRAMS`.
- **Heavily-behind students:** keep cohort subjects first, then re-takes — even
  when that puts a later subject in an earlier block. No special case.

## Still open
- The handbook Spring-offering scrape — only 9031 is seeded so far.
