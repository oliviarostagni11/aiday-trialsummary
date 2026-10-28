"""Compute every number in the monthly summary.

Each section returns Table objects. A Table holds rows of plain values plus a column spec,
so the Excel, Word and totals-CSV writers all render the same numbers.
Every percentage has its denominator in a visible column or in the column header.
"""
from __future__ import annotations

import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone

from .load import lab_date, ts

UTC = timezone.utc
SAMPLE_WINDOW_DAYS = {"CLEODX1": 14, "CLEODX2": 28}
NO_SAMPLE_REASON = "Unable to Obtain Study Blood Sample"
CRF_FORM_ORDER = ["Informed Consent", "Inclusion Criteria", "Exclusion Criteria", "Demographics", "Medical History",
                  "Health Information Collection", "CA 125 Test, Pre-Surgery", "Enrollment",
                  "Study Blood Sample Collection & Processing", "Post-Surgical Findings", "Study Completion"]
COMPLETE_STATES = ("Complete", "Locked")


# ---------------------------------------------------------------- table model
@dataclass
class Col:
    header: str
    kind: str = "int"  # text | int | pct | dec | rate | date | range
    width: float | None = None


@dataclass
class Table:
    key: str
    title: str
    cols: list[Col]
    rows: list[list] = field(default_factory=list)
    note: str = ""
    total_rows: set[int] = field(default_factory=set)  # row indexes styled as totals
    detail: bool = False  # participant- or sample-level list, kept out of the totals CSV

    def add(self, row: list, total: bool = False):
        if total:
            self.total_rows.add(len(self.rows))
        self.rows.append(row)


@dataclass
class Section:
    key: str
    title: str
    intro: str
    tables: list[Table]


# ---------------------------------------------------------------- helpers
def pct(n, d):
    return None if not d else n / d


def stats(values):
    """mean, median, min, max of a list of numbers, or Nones when empty."""
    if not values:
        return None, None, None, None
    return statistics.mean(values), statistics.median(values), min(values), max(values)


def fmt_day(d: date) -> str:
    return f"{d.day} {d.strftime('%b')}"


class Period:
    """The report month. Start is inclusive, end is exclusive (00:00 UTC on the 1st of next month)."""

    def __init__(self, month: str):
        y, m = (int(x) for x in month.split("-"))
        self.start = datetime(y, m, 1, tzinfo=UTC)
        self.end = datetime(y + (m == 12), m % 12 + 1, 1, tzinfo=UTC)
        self.first_day = self.start.date()
        self.last_day = (self.end - timedelta(days=1)).date()
        self.label = self.start.strftime("%B %Y")
        self.short = self.start.strftime("%b %Y")
        self.days = (self.end - self.start).days
        self.weeks_in_month = self.days / 7

    def in_month(self, dt) -> bool:
        if dt is None:
            return False
        if isinstance(dt, datetime):
            return self.start <= dt < self.end
        return self.first_day <= dt <= self.last_day

    def by_end(self, dt) -> bool:
        if dt is None:
            return False
        if isinstance(dt, datetime):
            return dt < self.end
        return dt <= self.last_day

    def weeks(self):
        """Monday-to-Sunday weeks that overlap the month. Part weeks count only days inside the month."""
        monday = self.first_day - timedelta(days=self.first_day.weekday())
        out = []
        while monday <= self.last_day:
            sunday = monday + timedelta(days=6)
            lo, hi = max(monday, self.first_day), min(sunday, self.last_day)
            partial = (lo, hi) != (monday, sunday)
            label = f"{fmt_day(monday)}–{fmt_day(sunday)}" + ("*" if partial else "")
            out.append((lo, hi, label))
            monday += timedelta(days=7)
        return out


# ---------------------------------------------------------------- shared participant view
def prepare(data: dict, period: Period) -> dict:
    sites = {s["site_id"]: s for s in data["sites"]}
    people = []
    for r in data["participants"]:
        p = {
            "pid": r["Participant ID"], "site_id": r["Site ID"], "site": r["Site"],
            "consent": ts(r["Consent at"]), "enrolled": ts(r["Enrolled at"]), "completed": ts(r["Completed at"]),
            "sf": ts(r["Screen failure at"]), "wd": ts(r["Withdrawal at"]), "wd_reason": r["Withdrawal reason"],
        }
        e = period.by_end
        p["status_eom"] = ("Completed" if e(p["completed"]) else "Withdrawn" if e(p["wd"]) else
                           "Screen failed" if e(p["sf"]) else "Enrolled, active" if e(p["enrolled"]) else
                           "Consented, not yet enrolled" if e(p["consent"]) else None)
        people.append(p)
    active_sites = [s for s in data["sites"] if site_active(s, period)]
    return {"sites": sites, "active_sites": active_sites, "people": people,
            "by_pid": {p["pid"]: p for p in people}}


def site_active(s, period) -> bool:
    """Activated by month end and not terminated before the month started.
    A terminated site with no termination date counts as inactive."""
    if s.get("activation_date") and date.fromisoformat(s["activation_date"]) > period.last_day:
        return False
    if s.get("terminated") == "Yes":
        return bool(s.get("termination_date")) and date.fromisoformat(s["termination_date"]) >= period.first_day
    return True


def site_name(ctx, site_id):
    s = ctx["sites"].get(site_id)
    return s["site"] if s else (site_id or "(no site)")


# ---------------------------------------------------------------- 1. enrollment
def enrollment(data, ctx, period: Period) -> Section:
    people = ctx["people"]
    consented = [p for p in people if period.by_end(p["consent"])]
    enrolled = [p for p in people if period.by_end(p["enrolled"])]
    m = period.in_month

    funnel = Table("funnel", "Recruitment summary",
                   [Col("Measure", "text", 40), Col(period.label), Col(f"To date (to {fmt_day(period.last_day)} {period.last_day.year})")],
                   note="Counts use the date of each event. Citrus holds no pre-screening date, so pre-screening is not shown.")
    for label, key in [("Consented", "consent"), ("Enrolled", "enrolled"), ("Screen failures", "sf"),
                       ("Withdrawals", "wd"), ("Completed study", "completed")]:
        funnel.add([label, sum(m(p[key]) for p in people), sum(period.by_end(p[key]) for p in people)])

    conv = Table("conversion", "Consent to enrollment",
                 [Col("Measure", "text", 40), Col("n"), Col("Denominator", "text", 28), Col("%", "pct")])
    c_month = [p for p in people if m(p["consent"])]
    n = sum(1 for p in c_month if p["enrolled"] and period.by_end(p["enrolled"]))
    conv.add([f"Consented in {period.short} and enrolled by month end", n, f"{len(c_month)} consented in {period.short}", pct(n, len(c_month))])
    conv.add(["Enrolled to date", len(enrolled), f"{len(consented)} consented to date", pct(len(enrolled), len(consented))])
    sf = sum(period.by_end(p["sf"]) for p in people)
    conv.add(["Screen failures to date", sf, f"{len(consented)} consented to date", pct(sf, len(consented))])

    status = Table("status_eom", f"Participant status at month end ({fmt_day(period.last_day)} {period.last_day.year})",
                   [Col("Status", "text", 40), Col("Participants"), Col("% of consented", "pct")],
                   note=f"Denominator: {len(consented)} participants consented to date.")
    order = ["Consented, not yet enrolled", "Enrolled, active", "Completed", "Withdrawn", "Screen failed"]
    counts = Counter(p["status_eom"] for p in consented)
    for s in order:
        status.add([s, counts.get(s, 0), pct(counts.get(s, 0), len(consented))])
    status.add(["Total consented", len(consented), pct(len(consented), len(consented))], total=True)

    # by site and week
    weeks = period.weeks()
    wk_cols = [Col(lbl) for _, _, lbl in weeks]
    by_site = Table("enrol_by_site_week", f"Enrollments by site and week, {period.label}",
                    [Col("Site", "text", 38), Col("Type", "text", 10)] + wk_cols +
                    [Col(f"{period.short} total"), Col("Per week", "rate"), Col("Enrolled to date"),
                     Col("Months active", "dec"), Col("Per month since activation", "rate")],
                    note=("Weeks run Monday to Sunday. * Part week: only days inside the month count. "
                          f"Per week = month total ÷ {period.weeks_in_month:.2f} weeks. "
                          "Per month since activation = enrolled to date ÷ months since the site activation date. "
                          "Activation dates and site types come from the site lookup table."))
    site_enrol = defaultdict(list)
    for p in enrolled:
        site_enrol[p["site_id"]].append(p["enrolled"].date())
    tot_week = [0] * len(weeks)
    active_ids = {s["site_id"] for s in ctx["active_sites"]}
    shown = [s for s in data["sites"] if s["site_id"] in active_ids or site_enrol.get(s["site_id"])]
    for s in sorted(shown, key=lambda s: s["site"]):
        ds = site_enrol.get(s["site_id"], [])
        wk = [sum(lo <= d <= hi for d in ds) for lo, hi, _ in weeks]
        tot_week = [a + b for a, b in zip(tot_week, wk)]
        month_n = sum(period.in_month(d) for d in ds)
        months_active = ((period.last_day - date.fromisoformat(s["activation_date"])).days / 30.4375
                         if s.get("activation_date") else None)
        label = s["site"] + (" (terminated)" if s["site_id"] not in active_ids else "")
        by_site.add([label, s["site_type"]] + wk +
                     [month_n, month_n / period.weeks_in_month, len(ds), months_active,
                      len(ds) / months_active if months_active else None])
    n_active = len(ctx["active_sites"])
    month_total = sum(tot_week)
    by_site.add(["All sites", ""] + tot_week + [month_total, month_total / period.weeks_in_month, len(enrolled), None, None], total=True)
    by_site.add([f"Average per active site ({n_active} sites)", ""] + [w / n_active for w in tot_week] +
                [month_total / n_active, month_total / period.weeks_in_month / n_active, len(enrolled) / n_active, None, None],
                total=True)
    for i in range(2, 2 + len(weeks)):
        by_site.cols[i].kind = "int"

    by_type = Table("enrol_by_type", "Enrollment by site type",
                    [Col("Group", "text", 22), Col("Active sites"), Col(f"Enrolled {period.short}"),
                     Col("Per site per month", "rate"), Col("Per site per week", "rate"),
                     Col("Enrolled to date"), Col("% of all enrolled to date", "pct")],
                    note=(f"Denominator for %: {len(enrolled)} participants enrolled to date. Month rates use active sites only. "
                          "Enrolled to date includes terminated sites. Site types come from the CLEO-001 site list."))
    types = list(dict.fromkeys(s["site_type"] for s in data["sites"]))
    groups = [(t, "site_type", t) for t in types] + [("Academic", "academic_status", "Academic"),
                                                    ("Non-academic", "academic_status", "Non-academic")]
    for label, field_, value in groups:
        ids = {s["site_id"] for s in ctx["active_sites"] if s[field_] == value}
        all_ids = {s["site_id"] for s in data["sites"] if s[field_] == value}
        mo = sum(1 for p in enrolled if p["site_id"] in ids and m(p["enrolled"]))
        td = sum(1 for p in enrolled if p["site_id"] in all_ids)
        by_type.add([label, len(ids), mo, pct(mo, len(ids)), pct(mo, len(ids)) and mo / len(ids) / period.weeks_in_month,
                     td, pct(td, len(enrolled))])

    trend = Table("enrol_trend", "Enrollments by month",
                  [Col("Month", "text", 12), Col("Consented"), Col("Enrolled"), Col("Cumulative enrolled")])
    first = min((p["enrolled"] for p in enrolled), default=period.start)
    y, mo_ = first.year, first.month
    cum = 0
    while (y, mo_) <= (period.start.year, period.start.month):
        key = (y, mo_)
        e_n = sum(1 for p in enrolled if (p["enrolled"].year, p["enrolled"].month) == key)
        c_n = sum(1 for p in consented if (p["consent"].year, p["consent"].month) == key)
        cum += e_n
        trend.add([date(y, mo_, 1).strftime("%b %Y"), c_n, e_n, cum])
        y, mo_ = (y + 1, 1) if mo_ == 12 else (y, mo_ + 1)

    return Section("enrollment", "Enrollment",
                   f"Enrollment activity in {period.label} and cumulative to month end.",
                   [funnel, conv, status, by_site, by_type, trend])


# ---------------------------------------------------------------- 2. completeness
def completeness(data, ctx, period: Period) -> Section:
    crf = data["crf"]

    def crf_row(label, rows):
        expected = [r for r in rows if r["status"] != "Unavailable"]
        c = Counter(r["status"] for r in expected)
        done = sum(c[s] for s in COMPLETE_STATES)
        return [label, len(expected), done, pct(done, len(expected)), c["Available"], c["Open Query"], c["Skipped"]]

    cols = lambda first: [Col(first, "text", 38), Col("Forms expected"), Col("Complete"), Col("% complete", "pct"),
                          Col("Outstanding"), Col("Open query"), Col("Skipped")]
    note = ("Forms expected = all forms open to the participant (status not 'Unavailable'). "
            "Complete = 'Complete' or 'Locked'. Outstanding = 'Available', not yet completed.")
    by_site = Table("crf_by_site", "CRF completeness by site", cols("Site"), note=note)
    groups = defaultdict(list)
    for r in crf:
        groups[r["site_id"]].append(r)
    for sid in sorted(groups, key=lambda s: site_name(ctx, s)):
        by_site.add(crf_row(site_name(ctx, sid), groups[sid]))
    by_site.add(crf_row("All sites", crf), total=True)

    by_form = Table("crf_by_form", "CRF completeness by form", cols("Form"), note=note)
    forms = defaultdict(list)
    for r in crf:
        forms[r["form_title"]].append(r)
    for f in sorted(forms, key=lambda f: CRF_FORM_ORDER.index(f) if f in CRF_FORM_ORDER else 99):
        by_form.add(crf_row(f, forms[f]))
    by_form.add(crf_row("All forms", crf), total=True)

    def epro_row(label, rows):
        t = Counter()
        for r in rows:
            for k in ("num_available", "num_expired", "num_completed", "num_skipped", "supp_num_overdue"):
                t[k] += int(r[k] or 0)
        assigned = t["num_available"] + t["num_expired"] + t["num_completed"] + t["num_skipped"]
        return [label, assigned, t["num_completed"], pct(t["num_completed"], assigned), t["num_available"],
                pct(t["num_available"], assigned), t["num_expired"], pct(t["num_expired"], assigned),
                t["num_skipped"], pct(t["num_skipped"], assigned), t["supp_num_overdue"]]

    ecols = lambda first: [Col(first, "text", 38), Col("Assigned"), Col("Completed"), Col("% completed", "pct"),
                           Col("Open"), Col("% open", "pct"), Col("Expired"), Col("% expired", "pct"),
                           Col("Skipped"), Col("% skipped", "pct"), Col("Overdue (open)")]
    enote = "Denominator: ePRO forms assigned. Open = still available to the participant. Overdue is a subset of open."
    e_form = Table("epro_by_form", "ePRO completeness by form", ecols("Form"), note=enote)
    ef = defaultdict(list)
    for r in data["epro"]:
        ef[r["form_title"]].append(r)
    for f in sorted(ef):
        e_form.add(epro_row(f, ef[f]))
    e_form.add(epro_row("All ePRO forms", data["epro"]), total=True)

    e_site = Table("epro_by_site", "ePRO completeness by site", ecols("Site"), note=enote)
    es = defaultdict(list)
    for r in data["epro"]:
        es[r["site_id"]].append(r)
    for sid in sorted(es, key=lambda s: site_name(ctx, s)):
        e_site.add(epro_row(site_name(ctx, sid), es[sid]))
    e_site.add(epro_row("All sites", data["epro"]), total=True)

    snap = data["meta"]["citrus_export"]
    return Section("completeness", "Data completeness",
                   (f"Completeness is a snapshot from the Citrus export on {snap:%d %b %Y}. "
                    "Citrus does not date form status changes, so month-end values cannot be rebuilt."),
                   [by_site, by_form, e_form, e_site])


# ---------------------------------------------------------------- 3. post-surgical findings
def post_surgical(data, ctx, period: Period) -> Section:
    sd_col = next(k for k in data["psf"][0] if "(Surgery Date)" in k) if data["psf"] else None
    rpt_col = next(k for k in data["psf"][0] if "(Pathology Report)" in k) if data["psf"] else None
    psf = {}
    for r in data["psf"]:
        surgery = date.fromisoformat(r[sd_col]) if r[sd_col] else None
        has_report = r[rpt_col].startswith(("permanent/", "temporary/"))
        psf[r["Participant ID"]] = {"surgery": surgery,
                                    "upload": ts(r["Completed At"]) if (r["Form Status"] == "COMPLETE" and has_report) else None}
    enrolled = [p for p in ctx["people"] if period.by_end(p["enrolled"])]
    rows = []
    for p in enrolled:
        f = psf.get(p["pid"], {})
        surgery = f.get("surgery") if period.by_end(f.get("surgery")) else None
        upload = f.get("upload") if period.by_end(f.get("upload")) else None
        rows.append({**p, "surgery": surgery, "upload": upload,
                     "days": (upload - p["enrolled"]).total_seconds() / 86400 if upload else None})
    surg = [r for r in rows if r["surgery"]]
    upl = [r for r in rows if r["upload"]]
    waiting = [r for r in surg if not r["upload"]]

    summ = Table("psf_summary", "Surgery and pathology reports",
                 [Col("Measure", "text", 40), Col(period.label), Col("To date"), Col("Denominator (to date)", "text", 30),
                  Col("% (to date)", "pct")],
                 note="Surgery complete = a surgery date is entered. Report uploaded = Post-Surgical Findings form complete with a report file.")
    summ.add(["Participants enrolled", sum(period.in_month(r["enrolled"]) for r in rows), len(rows), "", None])
    summ.add(["Surgery complete", sum(period.in_month(r["surgery"]) for r in surg), len(surg),
              f"{len(rows)} enrolled", pct(len(surg), len(rows))])
    summ.add(["Pathology report uploaded", sum(period.in_month(r["upload"]) for r in upl), len(upl),
              f"{len(surg)} with surgery complete", pct(len(upl), len(surg))])
    summ.add(["Awaiting pathology report", None, len(waiting), f"{len(surg)} with surgery complete", pct(len(waiting), len(surg))])

    turn = Table("psf_turnaround", "Days from enrollment to pathology report upload",
                 [Col("Reports uploaded", "text", 40), Col("Reports"), Col("Mean days", "dec"), Col("Median days", "dec"),
                  Col("Shortest", "dec"), Col("Longest", "dec")],
                 note="Upload date = 'Completed At' on the Post-Surgical Findings form.")
    for label, subset in [(f"In {period.label}", [r for r in upl if period.in_month(r["upload"])]), ("To date", upl)]:
        turn.add([label, len(subset), *stats([r["days"] for r in subset])])

    by_site = Table("psf_by_site", "Surgery and pathology reports by site",
                    [Col("Site", "text", 38), Col("Enrolled"), Col("Surgery complete"), Col("% of enrolled", "pct"),
                     Col("Report uploaded"), Col("% of surgeries", "pct"), Col("Awaiting report"),
                     Col("Median days to upload", "dec"), Col("Mean days to upload", "dec")])
    g = defaultdict(list)
    for r in rows:
        g[r["site_id"]].append(r)
    for sid in sorted(g, key=lambda s: site_name(ctx, s)):
        rs = g[sid]
        s_ = [r for r in rs if r["surgery"]]
        u_ = [r for r in rs if r["upload"]]
        mean, med, _, _ = stats([r["days"] for r in u_])
        by_site.add([site_name(ctx, sid), len(rs), len(s_), pct(len(s_), len(rs)), len(u_), pct(len(u_), len(s_)),
                     len(s_) - len(u_), med, mean])
    mean, med, _, _ = stats([r["days"] for r in upl])
    by_site.add(["All sites", len(rows), len(surg), pct(len(surg), len(rows)), len(upl), pct(len(upl), len(surg)),
                 len(waiting), med, mean], total=True)

    wl = Table("psf_awaiting", "Participants awaiting a pathology report",
               [Col("Participant", "text", 12), Col("Site", "text", 38), Col("Enrolled", "date"), Col("Surgery date", "date"),
                Col("Days since surgery")], detail=True,
               note=f"Days since surgery counted to {fmt_day(period.last_day)} {period.last_day.year}.")
    for r in sorted(waiting, key=lambda r: r["surgery"]):
        wl.add([r["pid"], site_name(ctx, r["site_id"]), r["enrolled"].date(), r["surgery"], (period.last_day - r["surgery"]).days])

    return Section("post_surgical", "Post-surgical findings",
                   "Surgeries, pathology report uploads and turnaround for enrolled participants.",
                   [summ, turn, by_site, wl])


# ---------------------------------------------------------------- 4. blood samples
def blood_samples(data, ctx, period: Period) -> Section:
    samples = []
    for r in data["inventory"]:
        coll, reg = lab_date(r["Date of Collection"]), lab_date(r["Date Registered"])
        if not period.by_end(reg):
            continue  # registered after month end
        code = r["Test Code"]
        days = (reg - coll).days
        samples.append({"req": r["Requisition Number"], "barcode": r["LabMatrix Barcode"], "pid": r["Subject ID"],
                        "site": r["Site"], "code": code, "coll": coll, "reg": reg, "days": days,
                        "limit": SAMPLE_WINDOW_DAYS[code], "in_window": days <= SAMPLE_WINDOW_DAYS[code],
                        "status": r["Current Status"]})
    code_to_site = {s["lab_site_code"]: s["site"] for s in data["sites"]}
    reqs = defaultdict(dict)
    for s in samples:
        reqs[s["req"]][s["code"]] = s
    same_day = [v for v in reqs.values() if len(v) == 2 and v["CLEODX1"]["reg"] == v["CLEODX2"]["reg"]]

    def window_rows(subset, same_day_pairs):
        d1 = [s for s in subset if s["code"] == "CLEODX1"]
        d2 = [s for s in subset if s["code"] == "CLEODX2"]
        ok = lambda xs: sum(s["in_window"] for s in xs)
        return [
            ("Samples registered (DX1 + DX2)", len(subset), None, None),
            ("DX1 registered within 14 days of collection", ok(d1), len(d1), pct(ok(d1), len(d1))),
            ("DX2 registered within 28 days of collection", ok(d2), len(d2), pct(ok(d2), len(d2))),
            ("All samples within their window", ok(subset), len(subset), pct(ok(subset), len(subset))),
            ("Samples outside their window", len(subset) - ok(subset), len(subset), pct(len(subset) - ok(subset), len(subset))),
            ("DX1 and DX2 registered on the same day (pairs)", len(same_day_pairs), None, None),
        ]

    month = [s for s in samples if period.in_month(s["reg"])]
    month_same = [v for v in same_day if period.in_month(v["CLEODX1"]["reg"])]
    summ = Table("blood_summary", "Sample registration at the central lab",
                 [Col("Measure", "text", 46), Col(f"{period.short} n"), Col(f"{period.short} denominator"), Col(f"{period.short} %", "pct"),
                  Col("To date n"), Col("To date denominator"), Col("To date %", "pct")],
                 note=(f"{period.label} counts samples registered in the month. "
                       "Each DX1 must be registered within 14 days of collection and each DX2 within 28 days. "
                       "DX1 and DX2 must ship separately, so a same-day registration is a deviation."))
    for a, b in zip(window_rows(month, month_same), window_rows(samples, same_day)):
        summ.add([a[0], a[1], a[2], a[3], b[1], b[2], b[3]])

    # pair status at month end
    enrolled = [p for p in ctx["people"] if period.by_end(p["enrolled"])]
    pid_reqs = defaultdict(set)
    for s in samples:
        pid_reqs[s["pid"]].add(s["req"])
    pairs = Table("blood_pairs", "Sample pair status at month end",
                  [Col("Status", "text", 46), Col("Participants"), Col("% of enrolled", "pct")],
                  note=(f"Denominator: {len(enrolled)} participants enrolled to date. "
                        "Overdue = the registration window for the missing sample has passed. Pending = still inside the window."))
    missing_list = Table("blood_missing_pair", "Samples missing their pair",
                         [Col("Requisition", "text", 12), Col("Participant", "text", 12), Col("Site", "text", 34),
                          Col("Collected", "date"), Col("Registered", "text", 10), Col("Missing", "text", 10),
                          Col("Days since collection"), Col("Status", "text", 10)], detail=True)
    no_sample = Table("blood_no_sample", "Enrolled participants with no sample registered",
                      [Col("Participant", "text", 12), Col("Site", "text", 34), Col("Enrolled", "date"),
                       Col("Days since enrollment"), Col("Status", "text", 34)], detail=True)
    c = Counter()
    for p in enrolled:
        if not pid_reqs.get(p["pid"]):
            age = (period.last_day - p["enrolled"].date()).days
            if p["wd_reason"] == NO_SAMPLE_REASON:
                c["No sample: withdrawn, sample not obtained"] += 1
                st = "Withdrawn: sample not obtained"
            elif age > SAMPLE_WINDOW_DAYS["CLEODX1"]:
                c["No sample registered: overdue"] += 1
                st = "Overdue"
            else:
                c["No sample registered: pending"] += 1
                st = "Pending"
            no_sample.add([p["pid"], site_name(ctx, p["site_id"]), p["enrolled"].date(), age, st])
            continue
        for req in pid_reqs[p["pid"]]:
            got = reqs[req]
            if len(got) == 2:
                c["Both samples registered"] += 1
                continue
            have = next(iter(got.values()))
            missing = "CLEODX2" if have["code"] == "CLEODX1" else "CLEODX1"
            age = (period.last_day - have["coll"]).days
            overdue = age > SAMPLE_WINDOW_DAYS[missing]
            c[f"{missing[-3:].replace('DX', 'DX')} missing: {'overdue' if overdue else 'pending'}"] += 1
            missing_list.add([req, p["pid"], code_to_site.get(have["site"], have["site"]), have["coll"], have["code"][-3:],
                              missing[-3:], age, "Overdue" if overdue else "Pending"])
    for label in ["Both samples registered", "DX2 missing: pending", "DX2 missing: overdue", "DX1 missing: pending",
                  "DX1 missing: overdue", "No sample registered: pending", "No sample registered: overdue",
                  "No sample: withdrawn, sample not obtained"]:
        pairs.add([label, c.get(label, 0), pct(c.get(label, 0), len(enrolled))])
    pairs.add(["Total enrolled", len(enrolled), pct(len(enrolled), len(enrolled))], total=True)
    missing_list.rows.sort(key=lambda r: (r[7] != "Overdue", -r[6]))

    by_site = Table("blood_by_site", "Sample registration by site, to date",
                    [Col("Site", "text", 34), Col("Samples registered"), Col("DX1 registered"), Col("DX1 within 14 days"),
                     Col("% DX1 in window", "pct"), Col("DX2 registered"), Col("DX2 within 28 days"), Col("% DX2 in window", "pct"),
                     Col("Same-day pairs"), Col("Missing pair: overdue")])
    overdue_by_site = Counter(r[2] for r in missing_list.rows if r[7] == "Overdue")
    same_by_site = Counter(code_to_site.get(v["CLEODX1"]["site"]) for v in same_day)
    g = defaultdict(list)
    for s in samples:
        g[code_to_site.get(s["site"], s["site"])].append(s)

    def site_row(name, xs, same, od):
        d1 = [s for s in xs if s["code"] == "CLEODX1"]
        d2 = [s for s in xs if s["code"] == "CLEODX2"]
        o1, o2 = sum(s["in_window"] for s in d1), sum(s["in_window"] for s in d2)
        return [name, len(xs), len(d1), o1, pct(o1, len(d1)), len(d2), o2, pct(o2, len(d2)), same, od]
    for name in sorted(g):
        by_site.add(site_row(name, g[name], same_by_site.get(name, 0), overdue_by_site.get(name, 0)))
    by_site.add(site_row("All sites", samples, len(same_day), sum(overdue_by_site.values())), total=True)

    trend = Table("blood_trend", "Sample registration by month",
                  [Col("Month", "text", 12), Col("Registered"), Col("Within window"), Col("Outside window"), Col("% within window", "pct")],
                  note="Samples registered at the central lab in each month, last 12 months. Denominator for %: samples registered that month.")
    y, mo_ = period.start.year, period.start.month
    months = []
    for _ in range(12):
        months.append((y, mo_))
        y, mo_ = (y - 1, 12) if mo_ == 1 else (y, mo_ - 1)
    for (y, mo_) in reversed(months):
        xs = [s for s in samples if (s["reg"].year, s["reg"].month) == (y, mo_)]
        ok = sum(s["in_window"] for s in xs)
        trend.add([date(y, mo_, 1).strftime("%b %Y"), len(xs), ok, len(xs) - ok, pct(ok, len(xs))])

    late = Table("blood_out_of_window", f"Samples registered outside their window, {period.label}",
                 [Col("Barcode", "text", 18), Col("Participant", "text", 12), Col("Site", "text", 34), Col("Collected", "date"),
                  Col("Registered", "date"), Col("Days to register"), Col("Limit (days)")], detail=True)
    for s in sorted((s for s in month if not s["in_window"]), key=lambda s: -s["days"]):
        late.add([s["barcode"], s["pid"], code_to_site.get(s["site"], s["site"]), s["coll"], s["reg"], s["days"], s["limit"]])

    same_list = Table("blood_same_day", "Pairs registered on the same day, to date",
                      [Col("Requisition", "text", 12), Col("Participant", "text", 12), Col("Site", "text", 34),
                       Col("Collected", "date"), Col("Registered", "date")], detail=True)
    for v in sorted(same_day, key=lambda v: v["CLEODX1"]["reg"], reverse=True):
        a = v["CLEODX1"]
        same_list.add([a["req"], a["pid"], code_to_site.get(a["site"], a["site"]), a["coll"], a["reg"]])

    loc = Table("blood_location", "Where registered samples are now",
                [Col("Current status", "text", 30), Col("Samples"), Col("% of registered", "pct")],
                note=f"Status from the inventory report dated {data['meta']['inventory_date']}. Denominator: {len(samples)} samples registered to month end.")
    lc = Counter(s["status"] for s in samples)
    for k in ["In Inventory", "Transferred", "Destroyed"]:
        loc.add([k, lc.get(k, 0), pct(lc.get(k, 0), len(samples))])

    return Section("blood", "Blood samples",
                   f"Central lab registration from the LabConnect inventory report dated {data['meta']['inventory_date']}.",
                   [summ, trend, pairs, by_site, loc, missing_list, no_sample, late, same_list])


# ---------------------------------------------------------------- 5. data queries
def queries(data, ctx, period: Period) -> Section:
    qs = []
    for r in data["queries"]:
        created, updated = ts(r["Created At"]), ts(r["Updated at"])
        if not period.by_end(created):
            continue
        closed_by_end = r["Status"] == "Closed" and period.by_end(updated)
        if closed_by_end:
            state = "Closed"
        elif r["Status"] == "Reviewing" and period.by_end(updated):
            state = "Reviewing"
        else:
            state = "Open"
        qs.append({"site_id": r["Site ID"], "type": r["Type"], "form": r["Form Title"], "state": state,
                   "created": created, "closed_in_month": closed_by_end and period.in_month(updated),
                   "days_open": None if state == "Closed" else (period.end - created).total_seconds() / 86400})
    total = len(qs)
    st = Counter(q["state"] for q in qs)
    summ = Table("query_summary", "Query status at month end",
                 [Col("Measure", "text", 40), Col("Queries"), Col("% of all queries raised", "pct")],
                 note=f"Denominator: {total} queries raised to month end. Reviewing = the site responded and a reviewer is checking it.")
    for k in ["Open", "Reviewing", "Closed"]:
        summ.add([k, st.get(k, 0), pct(st.get(k, 0), total)])
    summ.add(["Total raised to date", total, pct(total, total)], total=True)
    summ.add([f"Raised in {period.label}", sum(period.in_month(q["created"]) for q in qs), None])
    summ.add([f"Closed in {period.label}", sum(q["closed_in_month"] for q in qs), None])
    summ.add(["Open or reviewing for more than 90 days", sum(1 for q in qs if (q["days_open"] or 0) > 90), None])

    days = Table("query_days_open", "Days open for unresolved queries",
                 [Col("Status at month end", "text", 40), Col("Queries"), Col("Mean days open", "dec"),
                  Col("Median days open", "dec"), Col("Shortest", "dec"), Col("Longest", "dec")],
                 note=f"Days open = days from creation to the end of {fmt_day(period.last_day)} {period.last_day.year}.")
    for label, states in [("Open", ("Open",)), ("Reviewing", ("Reviewing",)), ("Open and reviewing", ("Open", "Reviewing"))]:
        xs = [q["days_open"] for q in qs if q["state"] in states]
        days.add([label, len(xs), *stats(xs)], total=label == "Open and reviewing")

    def grouped(key, title, label, namer=lambda k: k, order=None):
        t = Table(f"query_by_{key.replace('_id', '')}", title,
                  [Col(label, "text", 38), Col("Total"), Col("Open"), Col("% open", "pct"), Col("Reviewing"),
                   Col("% reviewing", "pct"), Col("Closed"), Col("% closed", "pct"), Col("Mean days open", "dec"),
                   Col("Median days open", "dec"), Col("Range of days open", "range", 14)],
                  note="Percentages use the group's total queries. Days open covers open and reviewing queries.")
        g = defaultdict(list)
        for q in qs:
            g[q[key]].append(q)

        def row(name, xs):
            c = Counter(q["state"] for q in xs)
            mean, med, lo, hi = stats([q["days_open"] for q in xs if q["days_open"] is not None])
            return [name, len(xs), c["Open"], pct(c["Open"], len(xs)), c["Reviewing"], pct(c["Reviewing"], len(xs)),
                    c["Closed"], pct(c["Closed"], len(xs)), mean, med, (lo, hi) if lo is not None else None]
        for k in (order or sorted(g, key=lambda k: namer(k))):
            if k in g:
                t.add(row(namer(k), g[k]))
        t.add(row("All", qs), total=True)
        return t

    by_site = grouped("site_id", "Queries by site", "Site", namer=lambda s: site_name(ctx, s))
    by_type = grouped("type", "Queries by type", "Query type")
    by_form = grouped("form", "Queries by form", "Form")
    return Section("queries", "Data queries",
                   "Query status at month end, built from each query's creation and last-update dates.",
                   [summ, days, by_site, by_type, by_form])


# ---------------------------------------------------------------- headline figures
def headline(sections: dict, period: Period, n_active: int) -> dict:
    """Pull the key figures used on the summary page. Each value names its own denominator."""
    T = {t.key: t for s in sections.values() for t in s.tables}
    funnel = {r[0]: r for r in T["funnel"].rows}
    crf_all = T["crf_by_site"].rows[-1]
    epro_all = T["epro_by_form"].rows[-1]
    psf = {r[0]: r for r in T["psf_summary"].rows}
    turn = T["psf_turnaround"].rows[-1]
    blood = {r[0]: r for r in T["blood_summary"].rows}
    pairs = {r[0]: r for r in T["blood_pairs"].rows}
    qsum = {r[0]: r for r in T["query_summary"].rows}
    qdays = T["query_days_open"].rows[-1]
    week_rows = T["enrol_by_site_week"].rows
    avg_row = week_rows[-1]
    in_win = blood["All samples within their window"]
    return {
        "enrolled_month": funnel["Enrolled"][1], "enrolled_to_date": funnel["Enrolled"][2],
        "consented_month": funnel["Consented"][1], "consented_to_date": funnel["Consented"][2],
        "completed_month": funnel["Completed study"][1],
        "per_site_month": avg_row[-5], "per_site_week": avg_row[-4],
        "active_sites": n_active,
        "crf_pct": crf_all[3], "crf_done": crf_all[2], "crf_expected": crf_all[1],
        "epro_pct": epro_all[3], "epro_done": epro_all[2], "epro_assigned": epro_all[1], "epro_open": epro_all[4],
        "surgery_to_date": psf["Surgery complete"][2], "surgery_month": psf["Surgery complete"][1],
        "reports_to_date": psf["Pathology report uploaded"][2], "reports_month": psf["Pathology report uploaded"][1],
        "reports_pct": psf["Pathology report uploaded"][4], "awaiting": psf["Awaiting pathology report"][2],
        "turn_median": turn[3], "turn_mean": turn[2], "turn_min": turn[4], "turn_max": turn[5],
        "samples_month": blood["Samples registered (DX1 + DX2)"][1],
        "window_pct_month": in_win[3], "window_n_month": in_win[1], "window_d_month": in_win[2],
        "window_pct_to_date": in_win[6],
        "same_day_month": blood["DX1 and DX2 registered on the same day (pairs)"][1],
        "missing_overdue": pairs["DX2 missing: overdue"][1] + pairs["DX1 missing: overdue"][1],
        "no_sample_overdue": pairs["No sample registered: overdue"][1],
        "q_open": qsum["Open"][1], "q_reviewing": qsum["Reviewing"][1], "q_total": qsum["Total raised to date"][1],
        "q_raised_month": qsum[f"Raised in {period.label}"][1], "q_closed_month": qsum[f"Closed in {period.label}"][1],
        "q_median_days": qdays[3], "q_mean_days": qdays[2],
        "q_over_90": qsum["Open or reviewing for more than 90 days"][1],
        "completed_to_date": funnel["Completed study"][2],
        "enrolled_prev_month": T["enrol_trend"].rows[-2][2] if len(T["enrol_trend"].rows) >= 2 else None,
        "late_month": len(T["blood_out_of_window"].rows),
        "path_over_60": sum(1 for r in T["psf_awaiting"].rows if r[-1] > 60),
        "samples_to_date": blood["Samples registered (DX1 + DX2)"][4],
        "crf_outstanding": crf_all[4], "crf_open_query": crf_all[5],
    }


def build(data: dict, month: str) -> tuple[Period, dict, dict]:
    period = Period(month)
    ctx = prepare(data, period)
    sections = {s.key: s for s in (enrollment(data, ctx, period), completeness(data, ctx, period),
                                    post_surgical(data, ctx, period), blood_samples(data, ctx, period),
                                    queries(data, ctx, period))}
    return period, sections, headline(sections, period, len(ctx["active_sites"]))
