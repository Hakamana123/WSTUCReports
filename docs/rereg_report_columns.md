# Re-registration report — what each column means

One reference so coaches, team leads and the data team read the report the
same way. Column letters are for the **Coach View** tab (also the layout of
every per-coach file in the zip download). The full sheet (`Query1`) has the
raw extract first and the same advice columns after it (`AK`–`AU`).

Coaches read this as the *Report Column Guide* page in the app
([rereg_column_guide.html](rereg_column_guide.html)). Change both together.

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
| O | Progress | Position by position across the whole course: `✓` passed, `◐` registered now (not passed yet), `○` still to pass. `○` covers failed *and* not yet attempted, because the file can't tell them apart. That's why it isn't `✗`, which means "failed" in T/U. |
| P | Withdrawal Flag | Mid-semester only, commencing students with **no pass** in Blocks 1–2 this session. See [the flags](#withdrawal-flag-column-p) below. Blank for everyone else, and **always blank for a paused student** (they get a note in the Advice Reason instead). |
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

## Withdrawal Flag (column P)

Only on a mid-semester report, and only for a **commencing, not paused**
student who has **no pass** in Blocks 1–2 this session:

| Flag | When | What to do |
|---|---|---|
| **ADVISE WITHDRAWAL** (red) | At least one ✗ in Blocks 1–2 (the other block can be ✗ or `0`), **and** registered in Block 3 or 4 | Withdrawal conversation: drop Blocks 3–4, restart next semester. No subject advice. The Advice Reason lists the failed subjects. |
| **CHECK ENROLMENT** (amber) | Not enrolled in Block 1 or 2 (`0`/`0`), **or** no Block 3/4 registration, so there's nothing to withdraw from | Not a withdrawal. Confirm the enrolment is right, or whether they're still studying. Normal advice is shown. |
| **AWAITING GRADE** | Enrolled in Block 1 or 2 with no grade recorded yet | Nothing yet. Re-run once grades are in. Normal advice is shown. |

The Advice Reason starts with a line explaining which case applies.

---

## Rules worth knowing

- **Conditional Enrolment = 30 credit points per semester, max.** On a
  mid-semester report the cap applies to the rest of this semester and starts
  from what the student is **actually enrolled in** this session: 10cp for each
  Block 1–2 subject (passed or failed) and 15cp for a prep they're doing now
  ("Prep in progress"). The Advice Reason shows the total, e.g. "Already
  enrolled this session: 35cp of 30". Blocks 3–4 only get what still fits, in
  the order failed / owed subjects, then electives, then prep. If nothing fits
  it says "30cp cap reached". What doesn't fit is listed under "Still to pass
  (later session)".
- **Exclusion**: no advice. Refer to the coach.
- **Paused**: advice is provisional (grey), no withdrawal advice, no comms until
  they confirm they're returning.
- **Full sheet only** (`Query1`): `Advice Source` says which engine produced the
  row (Grant calculator, position-based, v2 rule-tree, withdrawal, Summer). It's
  for checking, not for coaches.
