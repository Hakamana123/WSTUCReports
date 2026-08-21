"""
GLOSSARY — plain-language definitions of every term that shows up in the
advisory report (columns, confidence tiers, on-pattern statuses, bucket
names). Exists because the report's own audience (including the person
who built it) found "grounded/candidate" and bucket names like
off_pattern_partial opaque without an explanation next to them.

Kept as one place so the dashboard and the exported Excel workbook's
Glossary sheet can't drift apart - both read from the same dicts here
rather than each writing their own copy of the wording.
"""

from __future__ import annotations

COLUMN_GLOSSARY: list[tuple[str, str]] = [
    ("Student ID", "The student's ID number, from the uploaded roster."),
    ("Student Name", "The student's name, from the uploaded roster. Blank if the "
                      "source file didn't include it."),
    ("Program", "The WSU program code the student is enrolled in. 7188-7198 are "
                "diplomas; 9031/9034 are the University Preparation Program (UPP)."),
    ("On pattern", "Whether the subjects the student registered for match what "
                    "they're expected to take next, based on their program and "
                    "start date. Y = matches. Partial = some but not all of the "
                    "checkable subjects match. N = doesn't match. Unknown = can't "
                    "be checked yet for this student - see Reason."),
    ("Reason (if Unknown)", "Why 'On pattern' can't be determined for this "
                             "student - e.g. their program/start-date combination "
                             "isn't supported by the pattern data yet."),
    ("Subjects advised", "What the student's pattern says they should be "
                          "registered for, for the next two blocks (Spring Block "
                          "1 and 2 only - the pattern data doesn't cover further "
                          "blocks yet, see the technical notes in "
                          "report_builder.py if useful)."),
    ("Subjects registered", "What the student is actually registered for right "
                             "now, across all 4 blocks plus the prep subject."),
    ("Bucket", "Which advisory category the student has been sorted into - see "
               "the Buckets section of this glossary for what each one means."),
    ("Bucket confidence", "How solid that categorisation is - see the Confidence "
                           "tiers section of this glossary."),
    ("Advice", "The recommended next step, where one has actually been "
               "confirmed. Most buckets don't have confirmed advice yet - see "
               "Confidence tiers."),
]

CONFIDENCE_GLOSSARY: list[tuple[str, str]] = [
    ("High", "Both what's going on with this student's registration, AND what "
             "to do about it, have been confirmed. Safe to act on."),
    ("Medium", "We've confirmed what's actually going on with this student "
               "(e.g. they really are over a credit cap, or really are "
               "registered for the wrong subjects) - but we haven't yet "
               "confirmed what to do about it. Needs a decision before advising "
               "the student."),
    ("Low", "Even the classification itself is our best guess from the data, "
            "not yet confirmed at all. Treat as a starting point for "
            "discussion, not a finished rule."),
]

ON_PATTERN_GLOSSARY: list[tuple[str, str]] = [
    ("Y", "Registered subjects match what the pattern expects."),
    ("Partial", "Some, but not all, of the checkable subjects match."),
    ("N", "Registered subjects don't match what the pattern expects."),
    ("Unknown", "Can't be checked yet for this student - see the Reason column "
                "for why."),
]

# Kept in sync with SHORT_LABELS in pages/4_Reregistration_Advisory.py -
# every bucket name bucketing.py can produce should have an entry here.
BUCKET_GLOSSARY: list[tuple[str, str]] = [
    ("on_pattern_continuing",
     "Fully registered and on track. No action needed - this is the only "
     "bucket where the recommended action is fully confirmed."),
    ("on_pattern_partial_registration",
     "On track for what they've registered so far - just hasn't finished "
     "registering for all 4 blocks yet. Not the same as being off pattern."),
    ("on_pattern_at_risk_monitoring",
     "Fully registered and on track subject-wise, but their standing is below "
     "Good Standing - unclear whether they still need coach support despite "
     "correct registration."),
    ("off_pattern_partial",
     "Partially registered, and what they have registered doesn't match what's "
     "expected - a genuinely off-pattern case, not just an incomplete one."),
    ("exception_registered_wrong_subjects",
     "Fully registered, but for subjects that don't match what's expected. "
     "What to do about it (swap subjects? something else?) isn't confirmed."),
    ("exception_full_registration_unverified",
     "Fully registered, but we can't check whether the subjects are correct - "
     "their program/start-date combination isn't supported by the pattern data "
     "yet."),
    ("exception_partial_registration_unverified",
     "Partially registered, same 'can't check' issue as "
     "exception_full_registration_unverified."),
    ("zero_registration_unclear",
     "Good Standing, but hasn't registered for anything next semester, and no "
     "term start date was picked - so it's not clear if that's a real problem "
     "or just early/normal timing. Pick a term start date to resolve this into "
     "one of the two buckets below."),
    ("zero_registration_too_early",
     "Good Standing, hasn't registered yet - but the term hasn't started, so "
     "this is fully expected. No action needed."),
    ("zero_registration_overdue",
     "Good Standing, hasn't registered, and at least one block they should "
     "already be registered for has started - confirmed overdue, though what "
     "to actually do about it isn't confirmed yet."),
    ("success_coach_outreach",
     "A returning (continuing) student whose standing has dropped below Good "
     "Standing. Likely candidate for coach outreach, pending confirmation of "
     "the exact rule."),
    ("reapply_next_semester",
     "A newly-commencing student with a below-Good-Standing result and no "
     "registration at all - may need to restart next semester rather than "
     "continue."),
    ("exception_exclusion_still_registered",
     "Excluded from their program, but still shows an active registration - a "
     "known data anomaly that needs checking (expected lag, or should the "
     "registration have been pulled?)."),
    ("exception_no_commencement_period",
     "We don't know when this student started studying, so none of the "
     "pattern checks can run for them at all."),
    ("exception_ce_over_credit_cap",
     "On Conditional Enrolment and registered for more credit points than the "
     "confirmed 30cp cap allows."),
    ("exception_manual_review",
     "Doesn't fit any of the rules above - needs a human to look at it "
     "directly. Should rarely if ever appear given the current data."),
]


def glossary_dataframe():
    """Assemble the full glossary as a flat table, suitable for writing to a
    spreadsheet sheet or rendering in the dashboard. Local pandas import so
    modules that don't need it (e.g. a future non-Streamlit consumer) aren't
    forced to depend on it.
    """
    import pandas as pd

    rows = []
    for term, definition in COLUMN_GLOSSARY:
        rows.append({"Section": "Report columns", "Term": term, "Meaning": definition})
    for term, definition in CONFIDENCE_GLOSSARY:
        rows.append({"Section": "Bucket confidence tiers", "Term": term, "Meaning": definition})
    for term, definition in ON_PATTERN_GLOSSARY:
        rows.append({"Section": "On pattern values", "Term": term, "Meaning": definition})
    for term, definition in BUCKET_GLOSSARY:
        # Displayed report sheets show bucket names with underscores turned to
        # spaces (see pages/4_Reregistration_Advisory.py) - match that exact
        # format here so a reader can look a term up directly, rather than
        # having to mentally translate underscores back in.
        rows.append({"Section": "Buckets", "Term": term.replace("_", " "), "Meaning": definition})
    return pd.DataFrame(rows)
