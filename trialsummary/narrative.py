"""Plain-language highlights and action items, built from computed numbers only (no free text)."""
from __future__ import annotations

from datetime import timedelta


def _p(x):
    return "–" if x is None else f"{x:.0%}"


def _d(x):
    return "–" if x is None else f"{x:.0f}"


def _prev_month_label(period) -> str:
    return (period.start - timedelta(days=1)).strftime("%B %Y")


def sponsor_highlights(h: dict, period) -> list[str]:
    """High-level progress for the sponsor: enrollment, study completion, data progress. No query detail."""
    m = period.label
    if h["enrolled_prev_month"] is None:
        trend = ""
    elif h["enrolled_month"] > h["enrolled_prev_month"]:
        trend = f", up from {h['enrolled_prev_month']} in {_prev_month_label(period)}"
    elif h["enrolled_month"] < h["enrolled_prev_month"]:
        trend = f", down from {h['enrolled_prev_month']} in {_prev_month_label(period)}"
    else:
        trend = f", the same as in {_prev_month_label(period)}"
    completed_pct = h["completed_to_date"] / h["enrolled_to_date"] if h["enrolled_to_date"] else None
    return [
        f"{h['enrolled_month']} participants enrolled in {m}{trend}. Total enrollment is now {h['enrolled_to_date']:,}.",
        f"{h['active_sites']} sites are active and enrolled an average of {h['per_site_month']:.1f} participants each in {m}.",
        f"{h['completed_to_date']:,} participants have completed the study, {_p(completed_pct)} of those enrolled.",
        f"Pathology reports are {_p(h['reports_pct'])} complete: {h['reports_to_date']:,} of {h['surgery_to_date']:,} participants "
        f"with surgery have a report uploaded.",
        f"{_p(h['window_pct_month'])} of blood samples registered in {m} reached the central lab within the required window.",
    ]


def sponsor_actions(h: dict, period) -> list[str]:
    """Items that need site attention, kept high level. Internal follow-up such as open queries is left out."""
    m = period.label
    out = []
    if h["late_month"] or h["same_day_month"]:
        parts = []
        if h["late_month"]:
            parts.append(f"{h['late_month']} of {h['samples_month']} samples registered in {m} arrived outside the required window")
        if h["same_day_month"]:
            parts.append(f"{h['same_day_month']} sample pairs were shipped together instead of separately")
        out.append("Sample shipping: " + ", and ".join(parts) + ". Sites should ship each sample promptly after collection.")
    missing = h["missing_overdue"] + h["no_sample_overdue"]
    if missing:
        out.append(f"Missing samples: {missing} participants are missing one or both blood samples past the registration deadline.")
    if h["path_over_60"]:
        out.append(f"Pathology reports: {h['path_over_60']} participants are more than 60 days past surgery without a report uploaded.")
    return out


def highlights(h: dict, period, completeness: bool = True, audience: str = "internal") -> list[str]:
    if audience == "sponsor":
        return sponsor_highlights(h, period)
    m = period.label
    out = [
        f"{h['enrolled_month']} participants enrolled in {m}, and {h['consented_month']} consented. "
        f"{h['enrolled_to_date']} participants are enrolled to date across {h['active_sites']} active sites.",
        f"The {m} enrollment rate was {h['per_site_month']:.2f} per active site per month, "
        f"or {h['per_site_week']:.2f} per active site per week.",
        f"CRF completeness is {_p(h['crf_pct'])}: {h['crf_done']:,} of {h['crf_expected']:,} expected forms are complete. "
        f"ePRO completion is {_p(h['epro_pct'])}: {h['epro_done']:,} of {h['epro_assigned']:,} assigned forms, with {h['epro_open']:,} still open.",
        f"Pathology reports are uploaded for {h['reports_to_date']} of {h['surgery_to_date']} participants with surgery complete "
        f"({_p(h['reports_pct'])}). The median time from enrollment to upload is {_d(h['turn_median'])} days.",
        f"{_p(h['window_pct_month'])} of samples registered in {m} were inside their registration window "
        f"({h['window_n_month']} of {h['window_d_month']}).",
        f"{h['q_open'] + h['q_reviewing']} queries were unresolved at month end: {h['q_open']} open and {h['q_reviewing']} in review. "
        f"The median time open is {_d(h['q_median_days'])} days.",
    ]
    if not completeness:
        out = [x for x in out if not x.startswith("CRF completeness")]
    return out


def actions(h: dict, sections: dict, period, audience: str = "internal") -> list[str]:
    if audience == "sponsor":
        return sponsor_actions(h, period)
    T = {t.key: t for s in sections.values() for t in s.tables}
    out = []
    if h["missing_overdue"]:
        out.append(f"{h['missing_overdue']} participants have a sample with no pair registered after the window closed.")
    if h["no_sample_overdue"]:
        out.append(f"{h['no_sample_overdue']} enrolled participants have no sample registered at the central lab after 14 days.")
    late = len(T["blood_out_of_window"].rows)
    if late:
        out.append(f"{late} samples registered in {period.label} were outside their registration window.")
    if h["same_day_month"]:
        out.append(f"{h['same_day_month']} DX1 and DX2 pairs were registered on the same day in {period.label}.")
    long_wait = sum(1 for r in T["psf_awaiting"].rows if r[-1] > 60)
    if long_wait:
        out.append(f"{long_wait} participants are still waiting for a pathology report more than 60 days after surgery.")
    if h["q_over_90"]:
        out.append(f"{h['q_over_90']} queries have been open or in review for more than 90 days.")
    return out


def definitions(period, meta, completeness: bool = True) -> list[tuple[str, str]]:
    out = [
        ("Reporting month", f"{period.label}. Month counts use events dated 1 to {period.last_day.day} {period.start:%B %Y} (UTC)."),
        ("To date", f"All events up to the end of {period.last_day.day} {period.start:%B %Y}."),
        ("Sources", f"Citrus exports from {meta['citrus_export']:%d %b %Y %H:%M} UTC. "
                    f"LabConnect inventory report dated {meta['inventory_date']}. Site lookup table for site type and activation date."),
        ("Weeks", "Monday to Sunday. Part weeks at the start or end of the month count only days inside the month."),
        ("Enrollment rate", f"Per site per week = enrollments in the month ÷ {period.weeks_in_month:.2f} weeks. "
                            "Per site per month = enrollments in the month. The 'per active site' rates divide by the number of "
                            "sites activated by month end. Per month since activation = enrolled to date ÷ months since activation."),
        ("CRF completeness", "Complete or Locked forms ÷ forms expected. Forms expected excludes forms marked 'Unavailable', "
                             "which are not yet open to the participant. Snapshot at the Citrus export time."),
        ("ePRO completeness", "Each assigned ePRO form is Completed, Open, Expired or Skipped. Percentages use forms assigned. "
                              "Overdue forms are open forms past their due date."),
        ("Surgery complete", "A surgery date is entered on the Post-Surgical Findings form."),
        ("Pathology report uploaded", "The Post-Surgical Findings form is complete with a report file. "
                                      "The upload date is the form's 'Completed At' time."),
        ("Sample registration window", "DX1 within 14 days of collection. DX2 within 28 days of collection. "
                                       "DX1 and DX2 must be registered on different days because they ship separately."),
        ("Missing pair", "One sample of the pair is registered and the other is not. Overdue when the missing sample's window has passed."),
        ("Blood sample month view", "Counts samples registered at the central lab in the month."),
        ("Query status at month end", "Closed if closed by month end. Reviewing if the last update before month end moved it to review. "
                                      "Open otherwise."),
        ("Days open", "Days from query creation to the end of the month, for open and reviewing queries only."),
    ]
    if not completeness:
        out = [d for d in out if d[0] not in ("CRF completeness", "ePRO completeness")]
    return out


SECTION_ORDER = ["enrollment", "completeness", "post_surgical", "blood", "queries"]


def _d0(x):
    return "\u2013" if x is None else f"{x:.0f}"


def digest(sections: dict, head: dict, period, included: list[str]) -> dict:
    """Page-1 'Inside this report' rows and the key-figure strip that opens each section.
    Every value comes from numbers already in the report. Mirrored by digest() in app/src/engine.js."""
    from .charts import chart_specs, pct_text as p  # local import: charts does not depend on narrative

    specs, h, m = chart_specs(sections, period, head), head, period.short
    completed_pct = h["completed_to_date"] / h["enrolled_to_date"] if h["enrolled_to_date"] else None
    findings = {
        "enrollment": specs["enrollment_trend"]["title"],
        "completeness": f"{p(h['crf_pct'])} of expected CRF forms and {p(h['epro_pct'])} of assigned ePRO forms are complete",
        "post_surgical": (f"Pathology reports are {p(h['reports_pct'])} complete; {h['path_over_60']} of {h['awaiting']} "
                          f"still awaited are more than 60 days past surgery"),
        "blood": specs["sample_timeliness"]["title"],
        "queries": specs["queries_by_site"]["title"],
    }
    figures = {
        "enrollment": [
            (f"{h['enrolled_month']:,}", f"Enrolled in {m}", f"{h['consented_month']:,} consented in {m}"),
            (f"{h['enrolled_to_date']:,}", "Enrolled to date", f"{h['consented_to_date']:,} consented to date"),
            (f"{h['completed_to_date']:,}", "Completed the study", f"{p(completed_pct)} of enrolled"),
            (f"{h['active_sites']}", "Active sites", f"{h['per_site_month']:.2f} enrolled per site in {m}"),
        ],
        "completeness": [
            (p(h["crf_pct"]), "CRF forms complete", f"{h['crf_done']:,} of {h['crf_expected']:,} expected"),
            (p(h["epro_pct"]), "ePRO forms completed", f"{h['epro_done']:,} of {h['epro_assigned']:,} assigned"),
            (f"{h['crf_outstanding']:,}", "CRF forms outstanding", "open, not yet complete"),
            (f"{h['crf_open_query']:,}", "CRF forms with an open query", "snapshot at export"),
        ],
        "post_surgical": [
            (f"{h['surgery_month']:,}", f"Surgeries in {m}", f"{h['surgery_to_date']:,} to date"),
            (f"{h['reports_month']:,}", f"Reports uploaded in {m}", f"{h['reports_to_date']:,} to date"),
            (p(h["reports_pct"]), "Reports uploaded", f"of {h['surgery_to_date']:,} surgeries complete"),
            (f"{_d0(h['turn_median'])} days", "Median enrollment to upload", f"range {_d0(h['turn_min'])}\u2013{_d0(h['turn_max'])} days"),
        ],
        "blood": [
            (f"{h['samples_month']:,}", f"Samples registered in {m}", f"{h['samples_to_date']:,} to date"),
            (p(h["window_pct_month"]), "Inside registration window", f"{h['window_n_month']:,} of {h['window_d_month']:,} in {m}"),
            (f"{h['missing_overdue'] + h['no_sample_overdue']:,}", "Participants missing samples", "past the registration deadline"),
            (f"{h['same_day_month']:,}", "Pairs shipped together", f"registered on the same day in {m}"),
        ],
        "queries": [
            (f"{h['q_open'] + h['q_reviewing']:,}", "Unresolved queries", f"{h['q_open']:,} open, {h['q_reviewing']:,} reviewing"),
            (f"{h['q_raised_month']:,}", f"Raised in {m}", f"{h['q_closed_month']:,} closed in {m}"),
            (f"{_d0(h['q_median_days'])} days", "Median days open", "open and reviewing queries"),
            (f"{h['q_over_90']:,}", "Open more than 90 days", "open or in review"),
        ],
    }
    keys = [k for k in SECTION_ORDER if k in included]
    return {"inside": [[sections[k].title, findings[k]] for k in keys],
            "figures": {k: [list(f) for f in figures[k]] for k in keys}}
