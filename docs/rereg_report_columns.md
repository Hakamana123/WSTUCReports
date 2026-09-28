# Re-registration report — what each column means

One reference so coaches, team leads and the data team read the report the
same way. Column letters are for the **Coach View** tab (also the layout of
every per-coach file in the zip download). The full sheet (`Query1`) has the
raw extract first and the same advice columns after it (`AK`–`AU`).

Why each student gets the template they do is in
[rereg_principles_and_templates.md](rereg_principles_and_templates.md).

---

## How to read the advice cells (S–W)

| You see | It means |
|---|---|
| **Black text** | Register this subject in this block. This is the advice. |
| *Grey italic* (download) / `(brackets)` (on-screen preview) | For reference only. **Don't register it.** See the three cases below. |
| `+1 elective` | Register any elective in this block. |
| Blank | Nothing to register in this block. |
| `You do not need to register in a subject` | Nothing to register this session. |

**Grey shows up in three situations:**

1. **Mid-semester report** (target ends in `Block 3`). Blocks 1–2 (T, U) have
   already started. They show what the student **was enrolled in**, with its
   result:
   - `GEDU1001 ✓`: passed
   - `GEDU1001 ✗`: failed (F, FNS, E or W)
   - `GEDU1001` with no mark: enrolled, no grade recorded yet
   - `0`: **not enrolled** in that block at all

   Only Blocks 3–4 (V, W) are advice. Anything still owed from Blocks 1–2 is in
   the Advice Reason ("still owes X, Y — take next Autumn").
2. **Paused students** (Deferred / Leave of Absence, column I). **Every** advice
   cell is grey, because the plan is provisional until they confirm they're
   returning. Grey here does **not** mean a past registration.
3. **Summer report.** A grey subject is one the student still needs that
   doesn't run in Summer at their campus (or the Summer block is full). They
   take it in a later session. The Advice Reason says which.

> The ✓ / ✗ / 0 marks exist only on the mid-semester report, and only for this
> session's Blocks 1–2, because that's the only place the file carries a result
> for each block. Elsewhere a failed subject and a never-attempted one look the
> same in the data (both just "still to pass").

---

## Column by column (Coach View)

| Col | Column | What it tells you |
|---|---|---|
| A | STUDENT_ID | Student number. |
| B–D | FIRST_NAME, LAST_NAME, PREFERRED_NAME | Name. Use the preferred name in comms. |
| E | INSTITUTION_EMAIL_ADDRESS | Student email. |
| F | Coach | Assigned success coach. Also the file each student lands in for the per-coach split. |
| G | PROGRAM_CD | Diploma program code (e.g. 7196; 9031 = Nursing). |
| H | COMMENCEMENT_PERIOD | When they started. Decides current vs old cohort, and "commencing". |
| I | STUDY_PATH_STATUS | `Active Study Path`, or **Deferred / Leave of Absence** (= paused). |
| J | Progression Outcome | Raw standing from the extract: Good Standing, At Risk, Conditional Enrolment, Exclusion. **Blank for some students.** Column K explains the blanks. |
| K | Study Status | Column J with no silent blanks: the outcome if there is one, else `Not yet assessed (commencing)`, `Not assessed (paused)` or **`No outcome recorded - review`** (a student who should have an outcome and doesn't). |
| L | Other Enrolment | `Also enrolled: program X (coach)` when the same student has a second program row. Check which one is real. |
| M | Student Status | Count of what's left: "Outstanding: 1 prep, 2 core, 2 electives", or "All passed". |
| N | Progress Bar | The same as a bar and percentage. |
| O | Progress | Position by position across the whole course: `✓` passed, `✗` **still to pass**. **Not the same as the Block 1–2 marks.** Here `✗` covers failed *and* not yet attempted. |
| P | Withdrawal Flag | `ADVISE WITHDRAWAL` (red) = a commencing student who passed neither Block 1 nor Block 2 this session. **Never set for a paused student.** They get a note in the Advice Reason instead. |
| Q | Messaging Template | Which email/tab: Commencing, Template 1A / 1B / 1C / 2 / 3, Transition. Paused students are also collected on the *Paused – check enrolment* tab of the per-coach files. |
| R | Rereg Principle | The situation behind the template (On Pattern, Mostly Progressing, 3+ Sessions, …). |
| S | Prep Registration Advice | Prep subject to register (GEDU0016 / GEDU0017), if any. |
| T | Block 1 Registration Advice | Block 1 advice. **On a mid-semester report: what they took in Block 1 + result (grey).** |
| U | Block 2 Registration Advice | Block 2 advice. **On a mid-semester report: what they took in Block 2 + result (grey).** |
| V | Block 3 Registration Advice | Block 3 subject to register. |
| W | Block 4 Registration Advice | Block 4 subject to register. |
| X | Earliest Completion | Soonest they could finish at full load. `(est.)` = a projection; blank when it can't be estimated. |
| Y | Advice Reason | The plain-English why: what to register, what's owed later, the Conditional Enrolment cap, paused notes, and assumptions to sanity-check. **Read this before contacting the student.** |

In the per-coach files, S–W also carry the subject name after each code
(`GEDU1001 — Name`).

---

## Rules worth knowing

- **Conditional Enrolment = 30 credit points per semester, max.** The cap
  includes anything already taken in Blocks 1–2 this session. Order inside the
  cap: failed / owed subjects first, then electives, then prep. What doesn't fit
  is listed under "Still to pass (later session)".
- **Exclusion**: no advice. Refer to the coach.
- **Paused**: advice is provisional (grey), no withdrawal advice, no comms until
  they confirm they're returning.
- **Full sheet only** (`Query1`): `Advice Source` says which engine produced the
  row (Grant calculator, position-based, v2 rule-tree, withdrawal, Summer). It's
  for checking, not for coaches.
