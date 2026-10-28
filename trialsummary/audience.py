"""Who a summary is for, and what that audience may see."""

AUDIENCES = {
    # Sponsors never see participant-level data and do not get data completeness.
    "sponsor": {"label": "Sponsor", "exclude_sections": {"completeness"}, "detail_tables": False},
    # Internal readers get every section and the participant-level action lists.
    "internal": {"label": "Internal", "exclude_sections": set(), "detail_tables": True},
}


def sections_for(sections: dict, audience: str) -> dict:
    a = AUDIENCES[audience]
    return {k: s for k, s in sections.items() if k not in a["exclude_sections"]}


def tables_for(section, audience: str) -> list:
    return [t for t in section.tables if AUDIENCES[audience]["detail_tables"] or not t.detail]
