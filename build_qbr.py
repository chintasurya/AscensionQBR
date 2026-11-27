"""Ascension QBR, FY27 PI-1 (July to mid-October 2026), built on the IG template with IGPPT's ig_kit.

    python build_qbr.py out.pptx [--photo portrait.jpg] [--igppt <path to IGPPT/igppt>]

Executive style: no icons, numbers highlighted. Gradient template backgrounds on the title and agenda
slides; plain black (Dark Title Only) on every other slide, including the reflection. Validates with IGPPT's validate.py.

Numbers come from the Jira ILM export (71 items after removing ILM-1148, a PI-2 story) plus the 10 Sprint 5
items in the sprint status report: 81 items, 2 cancelled (ILM-1064, ILM-1101), 79 in scope, all complete.
"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
IGPPT = Path(sys.argv[sys.argv.index("--igppt") + 1]) if "--igppt" in sys.argv else HERE.parent / "IGPPT" / "igppt"
sys.path.insert(0, str(IGPPT / "scripts"))

from ig_kit import (BOTTOM, CYAN, GREEN, LEFT, MAGENTA, NAVY, ORANGE, RIGHT, WIDTH, YELLOW,  # noqa: E402
                    Deck, columns, fit_size, text_height)

LOGO = HERE / "assets" / "ascension-logo-reversed.png"
DIM_LINE = "2A2E35"


# ------------------------------------------------------------------ building blocks
def big_number(c, x, y, w, value, label, color, size=40, label_size=12, label_h=0.5):
    c.text(x, y, w, size / 72 * 1.3, value, size=size, color=color, light=True)
    c.text(x, y + size / 72 * 1.4, w, label_h, label, size=label_size, color=c.t.muted)


def stage_tracker(c, y, stages):
    """stages = [(name, state, note)], state: done, current, next, paused."""
    n = len(stages)
    w, xs = columns(n, gap=0)
    d = 0.3
    for i, (x, (name, state, note)) in enumerate(zip(xs, stages)):
        cx = x + w / 2
        fill = {"done": GREEN, "current": YELLOW}.get(state)
        line = {"next": "7A828D", "paused": ORANGE}.get(state)
        c.oval(cx - d / 2, y, d, fill=fill, line=line, line_width=1.5)
        if i < n - 1:
            c.line(cx + d / 2 + 0.08, y + d / 2, xs[i + 1] + w / 2 - d / 2 - 0.08, y + d / 2,
                   color=GREEN if state == "done" else "4A515C", width=2)
        c.text(x, y + 0.42, w, 0.3, name, size=13, bold=True, align="center")
        col = {"done": GREEN, "current": YELLOW, "paused": ORANGE}.get(state, c.t.dim)
        c.text(x, y + 0.74, w, 0.26, note.upper(), size=8.5, bold=True, spacing=1.5, color=col, align="center")


def feature_slide(deck, title, stages, highlights, rows, notes=""):
    """Stage tracker; highlighted outcomes on the left (value line, explanation line, same style for every item);
    problem / solution / impact summaries on the right."""
    c = deck.slide(title, eyebrow="Key Features Built", notes=notes)
    stage_tracker(c, 1.95, stages)
    y0, h = 3.25, BOTTOM - 3.25
    pw = 3.75
    c.card(LEFT, y0, pw, h)
    step = (h - 0.3) / len(highlights)
    for k, (value, label, color) in enumerate(highlights):
        y = y0 + 0.18 + k * step
        c.text(LEFT + 0.3, y, pw - 0.6, 0.4, value, size=20, bold=True, color=color)
        c.text(LEFT + 0.3, y + 0.4, pw - 0.6, step - 0.45, label, size=11.5, color=c.t.muted)
    x, lw = LEFT + pw + 0.3, 1.9
    tx, tw = x + lw, RIGHT - x - lw
    rh = h / len(rows)
    for k, (label, color, text) in enumerate(rows):
        yy = y0 + k * rh
        if k:
            c.line(x, yy, RIGHT, yy, color=DIM_LINE, width=1)
        c.label(x, yy + 0.2, lw - 0.15, label, color=color)
        c.text(tx, yy + 0.14, tw, rh - 0.26, text, fit=(14, 10.5))
    return c


# ------------------------------------------------------------------ deck
def build(out, photo=None):
    deck = Deck(theme="dark", background="black")

    # 1 Title (gradient)
    s = deck.title("Quarterly Business Review", "Ascension Provider Services | FY27 PI-1", "July to mid-October 2026",
                   notes="FY27 PI-1: sprints 1 to 6 plus the IP sprint, July to mid-October 2026.")
    from ig_kit import Canvas, DARK
    Canvas(s, DARK).image(LOGO, (13.333 - 4.2) / 2, 4.65, w=4.2)

    # 2 Agenda (gradient)
    deck.agenda(["Reflection", "PI-1 at a Glance", "Objectives and Business Impact", "Key Features Built",
                 "Operational Efficiency", "Delivery Metrics", "Looking Ahead to PI-2"])


    # 3 Reflection (plain black IG layout, quote in Trenda italic, portrait on the right)
    refl = deck._new("Dark Title Only", notes="Why this reflection: PI-1 was a team result. Business users and the "
                     "IG team built, tested and validated Texas CIN and DPA together, and both go-lives in PI-2 "
                     "depend on that partnership continuing.")
    deck._fill(refl, 0, ""); deck._fill(refl, 21, "Reflection")
    rc = Canvas(refl, DARK)
    photo = photo or HERE / "assets" / "jordan-reflection.png"
    rc.image(photo, 8.55, 0.55, h=6.35)
    rc.text(LEFT, 2.05, 7.6, 2.9, "“Talent wins games, but teamwork and intelligence win championships.”",
            size=40, bold=True, italic=True, valign="bottom")
    rc.text(LEFT, 5.2, 7.6, 0.45, "Michael Jordan", size=18, italic=True, color=CYAN)

    # 4 PI-1 at a glance
    c = deck.slide("PI-1 at a Glance", eyebrow="PI-1 at a Glance",
                   subtitle="Seven sprints, July to mid-October 2026: every item in scope completed")
    top, bot = c.top, BOTTOM - 0.1
    c.card(LEFT, top, 5.5, bot - top, border=CYAN)
    c.text(LEFT + 0.4, top + 0.3, 4.8, 1.65, "79/79", size=96, color=CYAN, light=True)
    c.text(LEFT + 0.4, top + 2.1, 4.8, 0.4, "Work items completed", size=20, bold=True)
    c.text(LEFT + 0.4, top + 2.65, 4.7, 1.0, "100% of PI-1 scope: 38 stories, 32 tests, 5 designs and 4 spikes.",
           size=14, color=c.t.muted)
    th = (bot - top - 0.4) / 3
    for k, (v, head, desc, col) in enumerate([
            ("111", "Story points delivered", "15.9 points per sprint on average", YELLOW),
            ("15", "Weekly regression cycles", "Every week from 6 July to 12 October", MAGENTA),
            ("2", "Programs ready for go-live", "Texas CIN and Digital Provider Application", GREEN)]):
        c.stat_tile(6.4, top + k * (th + 0.2), RIGHT - 6.4, th, v, head, desc, color=col, number_w=1.6, number_size=44)

    # 5 Objectives and business impact
    c = deck.slide("Objectives and Business Impact", eyebrow="Objectives")
    cols = [("Objective", 2.3), ("Key focus", 2.3), ("Status", 2.05), ("Key results", 2.3), ("Business impact", 3.18)]
    xs, x = [], LEFT
    for _, w in cols:
        xs.append(x); x += w
    for (name, w), x in zip(cols, xs):
        c.label(x + 0.18, 1.92, w - 0.3, name, color=CYAN)
    rows = [
        ("Expand the network to Texas (Texas CIN)", "Onboard Texas CIN as a new market on Network Services Salesforce",
         "Built; business UAT started; go-live in PI-2", "Texas record type, sharing, routing, stages",
         "100%", "Texas records kept separate", CYAN),
        ("Digital Provider Application (DPA)", "One digital intake for provider onboarding",
         "Built; UAT in progress; go-live in PI-2", "Auto invite, lead conversion, data mapping",
         "1", "digital form for onboarding", YELLOW),
        ("Preserve legacy org history", "Move historical records and files to Enterprise",
         "Built; UAT and production pending", "Historical data and file migration built",
         "1", "org for current and old records", MAGENTA),
        ("Automate provider dues", "Discovery and solution design",
         "Design complete; build to be planned", "Architecture, design document, data model",
         "7", "object dues model ready", GREEN)]
    ry, rh, gap = 2.27, 0.8, 0.18
    for k, (obj, focus, status, result, num, num_label, col) in enumerate(rows):
        y = ry + k * (rh + gap)
        c.card(LEFT, y, WIDTH, rh)
        c.text(xs[0] + 0.18, y + 0.1, cols[0][1] - 0.3, rh - 0.2, obj, size=12.5, bold=True, valign="middle")
        c.text(xs[1] + 0.18, y + 0.1, cols[1][1] - 0.3, rh - 0.2, focus, size=11.5, color=c.t.muted, valign="middle")
        c.text(xs[2] + 0.18, y + 0.1, cols[2][1] - 0.3, rh - 0.2, status, size=11.5, bold=True, color=col, valign="middle")
        c.text(xs[3] + 0.18, y + 0.1, cols[3][1] - 0.3, rh - 0.2, result, size=11.5, color=c.t.muted, valign="middle")
        c.text(xs[4] + 0.18, y + 0.08, 1.15, rh - 0.16, num, size=28, color=col, light=True, valign="middle")
        c.text(xs[4] + 1.4, y + 0.1, cols[4][1] - 1.55, rh - 0.2, num_label, size=11.5, valign="middle")
    oy = ry + 4 * (rh + gap) + 0.05
    c.card(LEFT, oy, WIDTH, BOTTOM - oy, fill=NAVY, border=CYAN)
    c.text(LEFT + 0.3, oy, WIDTH - 0.6, BOTTOM - oy, [
        {"text": "PI-1 OUTCOME: ", "size": 11, "bold": True, "color": CYAN, "spacing": 2},
        {"text": "79/79", "size": 20, "bold": True, "color": YELLOW}, {"text": " items completed | ", "size": 13},
        {"text": "111", "size": 20, "bold": True, "color": YELLOW}, {"text": " story points | ", "size": 13},
        {"text": "2", "size": 20, "bold": True, "color": YELLOW}, {"text": " programs ready for go-live in PI-2", "size": 13}],
        valign="middle")

    # 6 to 8 Key features built
    feature_slide(deck, "Texas CIN: Built and in Business UAT",
                  [("Discovery", "done", "Complete"), ("Design", "done", "Complete"), ("Build", "done", "Complete"),
                   ("Business UAT", "current", "Started"), ("Go-live", "next", "PI-2")],
                  [("100%", "Record separation between Texas CIN and National.", CYAN),
                   ("1 platform", "One Salesforce framework for multiple markets.", YELLOW),
                   ("Easy to onboard", "New markets such as Texas Seton join the same model.", MAGENTA),
                   ("Clear ownership", "Each team owns and sees only its own records and work.", GREEN)],
                  [("Problem", CYAN, "Texas CIN runs its own intake, stages, credentialing, contracting and participation, with "
                    "Texas-only ownership and provider changes that must flow across records. By hand: duplicates, "
                    "broken hierarchies and missed handoffs."),
                   ("Solution", YELLOW, "A Texas CIN record type from Lead to Contract, kept apart from National by market-based "
                    "sharing on one automation framework. Leads convert to the right provider structure, TIN "
                    "integrity is enforced and work routes to the Texas team."),
                   ("Impact", MAGENTA, "One Salesforce view of the Texas provider lifecycle, with "
                    "restricted record access, consistent data, less duplicate entry, clear status and pipeline "
                    "reporting, and lower risk when provider details change.")],
                  notes="Sharing: National and Texas CIN public groups, ownership-based sharing rules and the NS "
                        "Market field. Cascades originate from Healthcare Provider so they never cross markets. "
                        "Texas stages include Network Interest Form, M&C Committee Approval and CPA Created - TX.")
    feature_slide(deck, "Digital Provider Application: In UAT",
                  [("Discovery", "done", "Complete"), ("Design", "done", "Complete"), ("Build", "done", "Complete"),
                   ("Business UAT", "current", "In progress"), ("Go-live", "next", "PI-2")],
                  [("1 digital form", "One entry point for provider onboarding.", CYAN),
                   ("9 steps automated", "Intake automated end to end, invite to documents.", YELLOW),
                   ("Less manual work", "No re-keying from emails, documents and spreadsheets.", MAGENTA),
                   ("Better provider data", "Complete and consistent from the first submission.", GREEN)],
                  [("Problem", CYAN, "Onboarding needs tax and NPI, licenses, insurance, credentialing details and documents. "
                    "Staff re-keyed data, records drifted apart, attachments stayed on the lead and contacts were "
                    "created by hand."),
                   ("Solution", YELLOW, "One digital application in Salesforce. It invites eligible leads, prevents duplicate "
                    "invitations, converts the lead to an account and opportunity, maps data to provider records, "
                    "adds office-staff contacts and relinks documents."),
                   ("Impact", MAGENTA, "A single, standard entry point for onboarding, documents that follow the provider, a "
                    "faster path to contracting and credentialing, and staff focused on providers, not data entry.")])
    feature_slide(deck, "Legacy History: Built, UAT Pending",
                  [("Discovery", "done", "Complete"), ("Design", "done", "Complete"), ("Build", "done", "Complete"),
                   ("UAT", "paused", "Pending"), ("Production", "paused", "Pending")],
                  [("1 org", "Current and historical records in one place.", CYAN),
                   ("100%", "Of history kept apart from live data.", YELLOW),
                   ("Easy access", "Old records found in Enterprise Salesforce, no legacy org login.", MAGENTA)],
                  [("Problem", CYAN, "Network Services history sat in the legacy org, needed for audit, "
                    "compliance and reporting without cluttering the Health Cloud model."),
                   ("Solution", YELLOW, "Historical objects built in the Enterprise org, records migrated with "
                    "Data Loader, and files migrated with their record links kept."),
                   ("Impact", MAGENTA, "Easier access to old records for audit, compliance and reporting. UAT "
                    "and production are pending while the team focuses on the Texas CIN go-live.")])

    # 9 Discovery complete
    c = deck.slide("Discovery Complete for Two Future Builds", eyebrow="Key Features Built")
    w, (x1, x2) = columns(2, gap=0.3)
    for x, head, status, nums, pts, col in [
            (x1, "Provider Dues", "Discovery and design complete", [("7", "object data model configured"), ("6", "design and discovery items")],
             ["High-level architecture and solution design document", "Payment page branding started with UX"], YELLOW),
            (x2, "Texas ACO Intake", "Discovery complete", [("3", "manual tools to replace"), ("4", "processes in scope")],
             ["Intake, contracting, credentialing and rosters", "Salesforce future-state approach drafted"], MAGENTA)]:
        y, h = 1.95, BOTTOM - 1.95
        c.card(x, y, w, h)
        c.text(x + 0.35, y + 0.3, w - 0.7, 0.45, head, size=22, bold=True)
        c.text(x + 0.35, y + 0.82, w - 0.7, 0.3, status.upper(), size=10, bold=True, spacing=2, color=col)
        for k, (v, lab) in enumerate(nums):
            big_number(c, x + 0.35 + k * (w - 0.7) / 2, y + 1.35, (w - 0.9) / 2, v, lab, col, size=44, label_size=12)
        c.bullets(x + 0.35, y + 2.85, w - 0.7, h - 3.05, pts, fit=(15, 11), gap=8)

    # 10 Operational efficiency
    c = deck.slide("Operational Efficiency", eyebrow="Operational Efficiency",
                   subtitle="13 enhancements and support items completed alongside the major programs")
    cards = [("2", "Net Par letters automated", "Sent on Opportunity stage change. Live in production since 8 September.", GREEN),
             ("4", "HFN trigger scenarios", "Re-parenting, fee schedule, account type and facility type changes.", CYAN),
             ("3", "APDH data updates", "Provider fields added to the payload; real-time practitioner feed in DEV.", YELLOW),
             ("3", "DDG data support items", "Hierarchy model testing and bulk load testing.", MAGENTA),
             ("2", "Platform readiness items", "DocuSign security upgrade and a production support handbook.", ORANGE),
             ("15", "Weekly regression cycles", "Every release tested against the full regression suite.", CYAN)]
    w, xs = columns(3)
    h = (BOTTOM - c.top - 0.2) / 2
    for k, (v, head, desc, col) in enumerate(cards):
        x, y = xs[k % 3], c.top + (k // 3) * (h + 0.2)
        c.card(x, y, w, h)
        c.text(x + 0.3, y + 0.2, w - 0.6, 0.85, v, size=44, color=col, light=True)
        c.text(x + 0.3, y + 1.05, w - 0.6, 0.35, head, size=15, bold=True)
        c.text(x + 0.3, y + 1.45, w - 0.6, h - 1.6, desc, fit=(13, 10), color=c.t.muted)

    # 11 Delivery metrics (velocity, work mix and KPIs on one slide)
    c = deck.slide("Delivery Metrics", eyebrow="Delivery Metrics")
    w, xs = columns(4)
    for x, (v, lab, col) in zip(xs, [("79/79", "Work items completed", CYAN), ("111", "Story points delivered", YELLOW),
                                      ("15.9", "Average points per sprint", MAGENTA), ("15", "Regression cycles", GREEN)]):
        c.card(x, 1.9, w, 1.25)
        c.text(x + 0.3, 1.98, w - 0.6, 0.72, v, size=38, color=col, light=True)
        c.text(x + 0.3, 2.72, w - 0.6, 0.3, lab, size=12, bold=True)
    c.label(LEFT, 3.45, 6, "Story points per sprint", color=CYAN)
    c.chart(LEFT, 3.75, 7.4, BOTTOM - 3.75, "column", ["S1", "S2", "S3", "S4", "S5", "S6", "IP"],
            [{"name": "Story points", "values": [21, 20, 16, 20, 10, 16, 8]}])
    c.label(8.4, 3.45, 4, "Completed items by type", color=CYAN)
    c.chart(8.4, 3.75, RIGHT - 8.4, BOTTOM - 3.75, "bar", ["Stories", "Tests", "Designs", "Spikes"],
            [{"name": "Items", "values": [38, 32, 5, 4]}], colors=[YELLOW])

    # 12 Looking ahead to PI-2
    c = deck.slide("Looking Ahead to PI-2", eyebrow="Looking Ahead")
    y, h = 1.95, BOTTOM - 1.95
    lw = 7.7
    c.card(LEFT, y, lw, h, border=GREEN)
    c.label(LEFT + 0.35, y + 0.3, lw - 0.7, "In PI-2", color=GREEN)
    items = [("Texas CIN go-live", "Production release for Texas CIN users."),
             ("Texas hypercare", "Close support for Texas users after go-live."),
             ("DPA go-live", "UAT sign-off, then release of the Digital Provider Application."),
             ("Minor enhancements", "Operational efficiency requests from Network Services teams.")]
    step = (h - 0.85) / len(items)
    for k, (t, d) in enumerate(items):
        yy = y + 0.75 + k * step
        c.text(LEFT + 0.35, yy, 0.9, 0.7, f"0{k + 1}", size=32, color=GREEN, light=True)
        c.text(LEFT + 1.35, yy + 0.02, lw - 1.7, 0.35, t, size=17, bold=True)
        c.text(LEFT + 1.35, yy + 0.4, lw - 1.7, 0.35, d, size=12.5, color=c.t.muted)
    rx = LEFT + lw + 0.3
    rw = RIGHT - rx
    c.card(rx, y, rw, h)
    c.label(rx + 0.35, y + 0.3, rw - 0.7, "Not in PI-2", color=ORANGE)
    later = [("Legacy history migration", "Built; UAT and production paused for Texas CIN."),
             ("Provider Dues build", "Discovery and design complete."),
             ("Texas ACO intake", "Discovery complete.")]
    step = (h - 0.85) / len(later)
    for k, (t, d) in enumerate(later):
        yy = y + 0.75 + k * step
        c.text(rx + 0.35, yy, rw - 0.7, 0.35, t, size=15, bold=True)
        c.text(rx + 0.35, yy + 0.4, rw - 0.7, 0.6, d, size=12.5, color=c.t.muted)

    # 13 Decisions and support needed
    c = deck.slide("Decisions and Support Needed", eyebrow="Looking Ahead")
    c.bullets(LEFT, c.top + 0.1, 7.9, BOTTOM - c.top - 0.1, [
        "Confirm the Texas CIN go-live date and the hypercare window",
        "DPA UAT sign-off and go-live date",
        "Agree when Provider Dues and Texas ACO intake move to build",
        "Schedule legacy history UAT and production once Texas is live"], fit=(20, 12), gap=14)
    x, cw = 8.85, RIGHT - 8.85
    c.card(x, c.top, cw, 3.4, border=MAGENTA)
    c.text(x + 0.3, c.top + 0.25, cw - 0.6, 1.1, "2", size=60, color=MAGENTA, light=True)
    c.text(x + 0.3, c.top + 1.4, cw - 0.6, 0.5, "Go-lives planned in PI-2", size=15, bold=True)
    c.text(x + 0.3, c.top + 1.95, cw - 0.6, 1.2, "Texas CIN and the Digital Provider Application.", size=13, color=c.t.muted)

    # 14 Thank you
    deck.closing("thank-you")
    deck.save(out)


def _note(c, text):
    y = BOTTOM - 0.62
    c.card(LEFT, y, WIDTH, 0.62, fill=NAVY, border=NAVY)
    c.text(LEFT + 0.25, y, WIDTH - 0.5, 0.62, text, size=fit_size(text, WIDTH - 0.5, 0.5, 13, 10), valign="middle")


if __name__ == "__main__":
    out = Path(sys.argv[1])
    photo = Path(sys.argv[sys.argv.index("--photo") + 1]) if "--photo" in sys.argv else None
    build(out, photo)
    cmd = [sys.executable, str(IGPPT / "scripts" / "validate.py"), str(out), "--words", str(HERE / "igppt-words.txt")]
    sys.exit(subprocess.call(cmd))
