#!/usr/bin/env python3
"""Build the AdvisAR 4-slide hackathon deck.

Every figure on these slides was read out of the hack-team-12 repository. Three
portfolios appear in the deck and are never blended into one number:

  A. Northstar, 600 applications, CORRECTED workbook -- the headline run.
     Source: data/northstar/northstar-600-corrected-summary.md, reproduced by
     re-running engine/score_northstar_600_corrected.py:
         gross annual avoidable      $25,816,000
         one-time transition cost     $8,407,000  (32.6% of gross)
         net first year              $17,409,000
         portfolio run cost         $354,330,000; the CIO's 15% = $53,149,500
         retain 301 / invest 174 / retire 81 / consolidate 44 / replace 0
         301 of 600 all-pass, so 299 carry an action
         35 overlap groups, yet only 44 rows are consolidated
         c_consumption_price_variance unscored on 600 of 600 rows
         regression: the 20 applications shared with the 20-app run, 0 moved
     The superseded first cut of this workbook ($27,552,000 gross /
     $18,416,000 net / replace 7) is NOT quoted anywhere in the deck.
     No confidence-split dollar figure is quoted for this portfolio: the
     source carries no metered or consumption cost column, so the split
     cannot be computed for it.

  B. Northstar, 20 applications -- the detailed walkthrough.
     Source: data/northstar/Northstar-Disposition-Analysis-v3.xlsx, "Savings"
     sheet (PORTFOLIO row) and northstar-dispositions-v3.csv:
         run-rate $43,250,000 | gross $9,150,000 | transition $3,850,000
         net first year $5,300,000 | 15% target $6,487,500, not met
         retain 11 / consolidate 7 / invest 1 / retire 1
         confidence: 17 high, 2 medium, 1 needs validation
         6 overlap clusters

  C. The team's own 20-application synthetic portfolio -- the parity fixture.
     Source: data/synthetic-portfolio/applications-v2.csv, and the measured
     browser-vs-engine comparison in docs/wireframe-README.md:
         329 fields compared, 329 matched, 0 mismatches
         $22,057,000 run-rate | $5,818,716 net
         600-row scale test: 189 ms to parse, 133 ms to score and render

Other figures:
  - 18 inputs / 16-row table / 3.0 gate -> engine/generate_dataset.py, engine/README.md
  - 125 / 57 / 68 / 21 columns    -> data/client-intake/client-intake-requirements.md
  - 65 requirements, seven steps  -> docs/requirements/README.md
  - upload wiring, refusal message, three table filters -> read out of
    src/apprat-ai-wireframe-v2.html, not out of its README

Nothing in this deck describes what the deployed page renders: no session that
builds this deck can load it. Behaviour is described from the committed source.

Run:  python3 docs/build_deck.py docs/AdvisAR-Hackathon-Deck.pptx
"""

import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

# ---------------------------------------------------------------- palette
NAVY   = RGBColor(0x0F, 0x24, 0x38)
SLATE  = RGBColor(0x44, 0x57, 0x6D)
MUTED  = RGBColor(0x7C, 0x8A, 0x9B)
ACCENT = RGBColor(0x1F, 0x7A, 0x8C)
PANEL  = RGBColor(0xF1, 0xF4, 0xF7)
RULE   = RGBColor(0xD8, 0xDF, 0xE6)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
ONNAVY = RGBColor(0xB8, 0xC6, 0xD2)
DIVIDE = RGBColor(0x2C, 0x44, 0x5C)

FONT = "Calibri"
SW, SH = 13.333, 7.5
ML = 0.72
CW = SW - 2 * ML          # 11.893
FLOOR = 6.72              # nothing body-level may cross this

prs = Presentation()
prs.slide_width  = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]
NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# ---------------------------------------------------------------- helpers
def textbox(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, *, size, color, bold=False, italic=False, first=False,
         space_before=0, space_after=0, line=1.18, align=PP_ALIGN.LEFT):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line
    r = p.add_run()
    r.text = text
    f = r.font
    f.name, f.size, f.bold, f.italic = FONT, Pt(size), bold, italic
    f.color.rgb = color
    return p


def rich(tf, chunks, *, size, color, first=False, space_after=0, line=1.16):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_before = Pt(0)
    p.space_after = Pt(space_after)
    p.line_spacing = line
    for text, bold, col in chunks:
        r = p.add_run()
        r.text = text
        f = r.font
        f.name, f.size, f.bold = FONT, Pt(size), bold
        f.color.rgb = col if col is not None else color
    return p


def bulletize(p, char="▪", indent=0.20, color=ACCENT):
    """Real bullet glyph + hanging indent, in schema-legal child order."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(indent))))
    pPr.set("indent", str(-Emu(Inches(indent))))
    clr = etree.SubElement(pPr, f"{{{NS}}}buClr")
    etree.SubElement(clr, f"{{{NS}}}srgbClr").set("val", str(color))
    etree.SubElement(pPr, f"{{{NS}}}buFont").set("typeface", "Arial")
    etree.SubElement(pPr, f"{{{NS}}}buChar").set("char", char)
    return p


def bullet(tf, chunks, *, size=13.5, color=SLATE, space_after=7, first=False):
    p = rich(tf, chunks, size=size, color=color, first=first,
             space_after=space_after)
    return bulletize(p)


def rect(slide, x, y, w, h, fill, *, shape=MSO_SHAPE.RECTANGLE, adj=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if adj is not None:
        s.adjustments[0] = adj
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return s


def header(slide, eyebrow, headline, sub=None):
    """Accent rule, eyebrow, one-line headline, optional sub. Returns next y."""
    rect(slide, ML, 0.46, 0.62, 0.055, ACCENT)
    tf = textbox(slide, ML, 0.62, CW, 0.28)
    para(tf, eyebrow.upper(), size=10.5, color=MUTED, bold=True, first=True)
    tf = textbox(slide, ML, 0.94, CW, 0.46)
    para(tf, headline, size=27, color=NAVY, bold=True, first=True, line=1.06)
    if sub:
        tf = textbox(slide, ML, 1.56, CW * 0.90, 0.30)
        para(tf, sub, size=13.5, color=SLATE, first=True, line=1.2)
        return 2.08
    return 1.62


def column_head(slide, x, y, w, text):
    tf = textbox(slide, x, y, w, 0.26)
    para(tf, text.upper(), size=10.5, color=ACCENT, bold=True, first=True)
    rect(slide, x, y + 0.29, w, 0.011, RULE)
    return y + 0.46


def stat_panel(slide, x, y, w, h, number, label, *, dark=False, num_size=40):
    rect(slide, x, y, w, h, NAVY if dark else PANEL,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.055)
    rect(slide, x, y + 0.16, 0.055, h - 0.32, ACCENT)
    tf = textbox(slide, x + 0.34, y + 0.18, w - 0.62, h - 0.36,
                 anchor=MSO_ANCHOR.MIDDLE)
    para(tf, number, size=num_size, color=WHITE if dark else NAVY, bold=True,
         first=True, line=1.0)
    para(tf, label, size=11, color=ONNAVY if dark else SLATE,
         space_before=5, line=1.18)


def footnote(slide, text):
    rect(slide, ML, SH - 0.78, CW, 0.011, RULE)
    tf = textbox(slide, ML, SH - 0.62, CW - 2.75, 0.34)
    para(tf, text, size=9.5, color=MUTED, italic=True, first=True, line=1.18)


def page_tag(slide, n, label):
    tf = textbox(slide, SW - ML - 2.60, SH - 0.62, 2.60, 0.22)
    para(tf, f"AdvisAR  ·  {label}  ·  {n} / 4", size=9.5, color=MUTED,
         first=True, align=PP_ALIGN.RIGHT)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text.strip()


# ================================================================ SLIDE 1
s1 = prs.slides.add_slide(BLANK)

rect(s1, ML, 0.46, 0.62, 0.055, ACCENT)
tf = textbox(s1, ML, 0.62, CW, 0.28)
para(tf, "ABERDEEN ADVISORS  ·  TEAM 12, HALLUCINATORS  ·  PROMPT #5",
     size=10.5, color=MUTED, bold=True, first=True)

tf = textbox(s1, ML, 0.90, CW, 0.74)
para(tf, "AdvisAR", size=44, color=NAVY, bold=True, first=True, line=1.0)

tf = textbox(s1, ML, 1.72, CW * 0.80, 0.50)
para(tf, "Application rationalization that returns a defensible disposition "
         "for every application in a portfolio — with the evidence attached.",
     size=14, color=SLATE, first=True, line=1.2)

rect(s1, ML, 2.34, CW, 0.011, RULE)

LW = 7.35
RX = 8.55
RW = ML + CW - RX          # 4.063

yl = column_head(s1, ML, 2.54, LW,
                 "The problem — Northstar Global Health, a health-system CIO")
tf = textbox(s1, ML, yl, LW, 3.4)
bullet(tf, [("600 applications, SaaS and AI tools. ", True, NAVY),
            ("Nobody can defend the portfolio line by line.", False, None)],
       first=True)
bullet(tf, [("Rationalization today is a spreadsheet exercise: ", True, NAVY),
            ("months of interviews, a facilitated workshop per application, "
             "inconsistent judgement between them.", False, None)])
bullet(tf, [("Decisions arrive without a path back to the evidence, ", True, NAVY),
            ("so each one gets relitigated in the next meeting.", False, None)])
bullet(tf, [("Healthcare removes the easy answer. ", True, NAVY),
            ("Patient-care criticality and compliance outrank cost — the "
             "expensive application is often the one that must not move.",
             False, None)])
bullet(tf, [("The time goes early. ", True, NAVY),
            ("Across the seven-step process, it is inventory, analysis and "
             "the decision itself that consume the engagement. Execution is "
             "not the bottleneck.", False, None)])

stat_panel(s1, RX, 2.54, RW, 1.58, "600",
           "applications, scored end to end in one run — not a sample of the "
           "portfolio. By hand it is a workshop each, which is why this never "
           "finishes.", dark=True, num_size=42)

rect(s1, RX, 4.32, RW, 1.22, PANEL, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
tf = textbox(s1, RX + 0.30, 4.54, RW - 0.60, 0.85)
para(tf, "THE SEVEN-STEP PROCESS WE BUILT TO", size=9.5, color=ACCENT,
     bold=True, first=True)
para(tf, "Discover → Inventory → Analyse → Decide → Roadmap → Execute "
         "→ Monitor", size=12, color=NAVY, bold=True, space_before=6, line=1.25)

footnote(s1, "Northstar Global Health is a fictional organisation. Every "
             "figure in this deck is synthetic and illustrative — product and "
             "vendor names are real; nothing here describes an actual client.")
page_tag(s1, 1, "Problem")

notes(s1, """
Northstar Global Health is our health system: 600 applications, SaaS and AI tools, and a CIO under
pressure to take cost out of technology without breaking a clinical workflow that somebody depends
on. And 600 is not a number we are waving at - we scored all six hundred of them end to end, in one
run, and you will see what came out on the impact slide.

The way this gets done today is a spreadsheet and a calendar. You interview owners, you run a
facilitated workshop per application, and because those workshops happen months apart with
different people in the room, the judgement is not consistent between them. That is why
rationalization rarely finishes at portfolio scale - it is not that nobody wants to do it, it is
that the unit cost per application is a meeting.

And the decisions that come out have no traceable path back to evidence. So when a business owner
objects, there is nothing to point at, and the decision gets relitigated. That is the real
failure mode: not a wrong answer, an unarguable one.

Healthcare makes it harder in a specific way. You cannot rank by cost and cut from the top. In
Northstar's portfolio the largest single line item is the core EHR at 8.75 million a year, and it
is exactly the thing you must not touch. Patient-care criticality and compliance outrank cost, so
the expensive application is very often the one that has to stay.

We built against a seven-step rationalization process - discover, inventory, analyse, decide,
roadmap, execute, monitor - and wrote 65 requirements mapped across all seven. Worth saying where
the time actually goes: steps two through four. Gathering the data, mapping the overlap, and
reaching a decision. Execution is not the bottleneck. Getting to a decision you can defend is the
bottleneck, and that is the part we automated.

One disclaimer that covers the whole deck: Northstar is fictional and every number you will see
is synthetic. The product and vendor names are real; the portfolio is not.
""")


# ================================================================ SLIDE 2
s2 = prs.slides.add_slide(BLANK)
y = header(s2, "Solution · AdvisAR",
           "Four gates, a 16-row lookup, and an honest intake list",
           "Every application comes out with one of five dispositions, a "
           "priority, a written rationale and a confidence level.")

LW2 = 6.10
RX2 = ML + LW2 + 0.66      # 7.48
RW2 = ML + CW - RX2        # 5.133

yl = column_head(s2, ML, y, LW2, "How a disposition is reached")
tf = textbox(s2, ML, yl, LW2, 3.4)
bullet(tf, [("18 scored inputs, ", True, NAVY),
            ("1–5 in half steps — 5 always the favourable end, so a high risk "
             "score means low risk.", False, None)], first=True)
bullet(tf, [("Four dimensions: ", True, NAVY),
            ("business value, technical health, cost efficiency, risk.",
             False, None)])
bullet(tf, [("Each gated at 3.0 of 5. ", True, NAVY),
            ("At or above passes, below fails.", False, None)])
bullet(tf, [("The pass/fail key is looked up in a 16-row decision table ",
             True, NAVY),
            ("returning one of five dispositions — invest, retain, "
             "consolidate, replace, retire — plus a priority. The key doubles "
             "as the rationale.", False, None)])
bullet(tf, [("Two guardrails then run: ", True, NAVY),
            ("retire and replace are barred on an application still ramping; a "
             "duplicate inside an overlap cluster is folded into the named "
             "survivor instead.", False, None)])
bullet(tf, [("Thresholds, weights and all 16 rows are configuration, not code.",
             True, NAVY)])

yr = column_head(s2, RX2, y, RW2, "The intake story — the differentiator")
tf = textbox(s2, RX2, yr, RW2, 2.8)
bullet(tf, [("We tell the client exactly what to send. ", True, NAVY),
            ("The model is 125 columns; the client supplies 57 and AdvisAR "
             "derives the other 68.", False, None)], first=True)
bullet(tf, [("Most of it is files, not a form ", True, NAVY),
            ("— five system extracts plus eight questions per application, "
             "about a ten-minute conversation.", False, None)])
bullet(tf, [("A 21-item minimum viable intake ", True, NAVY),
            ("still returns a disposition and a priority for every "
             "application, at capped confidence by design.", False, None)])
bullet(tf, [("65 requirements ", True, NAVY),
            ("mapped across the seven process steps.", False, None)])

stat_panel(s2, RX2, yr + 2.62, RW2, 1.22, "57 of 125",
           "columns is all we ask a client for. The tool derives the other 68 "
           "— and says so, column by column.", dark=True, num_size=32)

footnote(s2, "Scoring model, thresholds and the 16-row table: engine/  ·  "
             "column tiers and the minimum intake: data/client-intake/  ·  "
             "65 requirements: docs/requirements/")
page_tag(s2, 2, "Solution")

notes(s2, """
This is AdvisAR, and the mechanics are deliberately boring, because boring is what survives a
steering committee.

Eighteen inputs per application, each scored one to five in half steps. Five is always the
favourable end - that matters for risk and cost, where five means controlled and cheap, so you
never have to remember which way a number points. Sixteen of the eighteen carry weight; two,
absolute cost band and end-user perceived quality, are collected and stored at weight zero,
because we decided cost should move priority without being able to condemn something on its own.

Those roll into four dimensions - business value, technical health, cost efficiency, risk - and
each dimension is gated at 3.0. At or above, it passes. Below, it fails. So every application
produces a four-character pass/fail key, and that key is looked up in a sixteen-row table which
returns one of five dispositions plus a priority. Every possible combination has a row, so there
is no undefined case, and the key itself is the rationale: "business value passes, technical
health passes, cost efficiency fails, risk passes" is a sentence a CIO can argue with.

Two guardrails run after the lookup, and they matter. First, retire and replace are barred
outright for an application that has not finished ramping up - if the gates say retire on a
young pilot, that becomes a funded invest instead, because you paid for it and have not learned
anything from it yet. Second, if an application is the duplicate inside an overlap cluster and
there is a named survivor holding the same capability, it is folded into that survivor rather
than switched off outright - the capability persists, so somebody's workflow does not simply
disappear. On Northstar those two guardrails moved five rows. Both are written onto the row, so
you can always see what the gates said and what overrode it.

Now the part I would actually sell, which is intake. The hardest question in a rationalization
engagement is "what do you need from us", and the usual answer is a two-hundred-question survey
that comes back half empty. Our answer is precise: the model has 125 columns, we ask you for 57,
and we derive 68. Most of it is files rather than typing - five system extracts, plus eight
questions per application aimed at the named owner, which is about ten minutes each.

And there is a floor. A 21-item minimum intake still returns a disposition and a priority for
every application. It does not pretend to be as good: confidence is capped at medium, because
that intake populates only about half the fields the model consumes, and high confidence is
deliberately unreachable there. That is the honest version of "start now with what you have".
""")


# ================================================================ SLIDE 3
s3 = prs.slides.add_slide(BLANK)
y = header(s3, "Demo · Northstar's 20 applications, and the same model at 600",
           "Intake workbook in, defensible shortlist out",
           "All of it runs today: the 20-application walkthrough, the same "
           "model at 600, and the same scoring in the browser.")

steps = [
    ("1", "Intake",       "Client-supplied columns,\nopened read-only"),
    ("2", "Score & gate", "18 inputs → 4 dimensions,\neach gated at 3.0"),
    ("3", "Look up",      "16-row table, then\nthe two guardrails"),
    ("4", "Explain",      "Disposition, priority,\nrationale, confidence"),
    ("5", "Shortlist",    "11 no-action published,\n9 rows to argue about"),
]
px, pw, gap = ML, 2.18, 0.245
for i, (n, title, body) in enumerate(steps):
    last = i == len(steps) - 1
    rect(s3, px, y, pw, 1.10, NAVY if last else PANEL,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.085)
    tb = textbox(s3, px + 0.20, y + 0.17, pw - 0.40, 0.80)
    para(tb, f"{n}   {title}", size=12, color=WHITE if last else NAVY,
         bold=True, first=True)
    para(tb, body, size=9.5, color=ONNAVY if last else SLATE,
         space_before=4, line=1.15)
    if not last:
        ta = textbox(s3, px + pw + 0.02, y + 0.40, gap - 0.04, 0.30,
                     anchor=MSO_ANCHOR.MIDDLE)
        para(ta, "›", size=16, color=ACCENT, bold=True, first=True,
             align=PP_ALIGN.CENTER)
    px += pw + gap

y2 = y + 1.44
yl = column_head(s3, ML, y2, LW2, "What runs today")
tf = textbox(s3, ML, yl, LW2, 2.8)
bullet(tf, [("The command-line engine is the reference implementation. ",
             True, NAVY),
            ("One dependency — openpyxl — no build step. It writes a 9-sheet "
             "workbook: dispositions, clusters, savings, agreement with the "
             "client's own labels, assumptions and sanity checks.",
             False, None)],
       first=True)
bullet(tf, [("Every row explains itself — ", True, NAVY),
            ("which dimension failed, at what score, the successor it folds "
             "into, and the evidence gap behind its confidence. On these 20: "
             "17 high, 2 medium, 1 needs validation.", False, None)])
bullet(tf, [("The same engine, unchanged, scored 600 applications ", True, NAVY),
            ("in one run — seconds, not weeks — and all 20 above kept their "
             "dispositions: 0 moved.", False, None)])

yr = column_head(s3, RX2, y2, RW2, "The web app — upload and score")
tf = textbox(s3, RX2, yr, RW2, 2.0)
bullet(tf, [("Upload a workbook or CSV and score it live. ", True, ACCENT),
            ("A real file dialogue, .xlsx or .csv in our intake schema. The "
             "page rewrites itself from the file: figures, decision mix, "
             "guardrails, clusters, and a per-row table with rationale and "
             "dimension scores. No server, no internet.", False, None)],
       first=True)
bullet(tf, [("It agrees with the engine, measured: ", True, NAVY),
            ("329 of 329 values match, 0 mismatches.", False, None)])

stat_panel(s3, RX2, yr + 1.78, RW2, 0.94, "329 / 329",
           "browser-versus-engine values agreed, 0 mismatches.",
           num_size=26)

footnote(s3, "Engine: engine/score_northstar_v3.py and "
             "score_northstar_600_corrected.py  ·  browser scoring and the "
             "measured parity: src/apprat-ai-wireframe-v2.html, "
             "docs/wireframe-README.md  ·  deployed at hack-team-12.vercel.app")
page_tag(s3, 3, "Demo")

notes(s3, """
Let me walk Northstar's flow, and I will be precise about what is built and what is not.

Step one, intake. A workbook of the columns we asked the client for. The engine opens it
read-only and never writes back to it.

Step two, scoring. Eighteen inputs derived and scored, rolled into the four dimensions, each
gated at 3.0. Where an input has no source it is skipped and the dimension renormalises over what
is actually populated - it does not silently score a blank as a zero.

Step three, the lookup and the two guardrails I just described.

Step four is the one that makes this usable. Every row comes out with its rationale written in
prose - not just "consolidate", but which dimension failed, at what score, which input drove it,
and which application it folds into. Plus a confidence level and the specific evidence gap behind
that level. On Northstar's twenty applications: seventeen rows at high confidence, two at medium,
and one - the video-conferencing tool - held as needs validation rather than given a
recommendation at all, because its telehealth and conference-room dependencies were not
evidenced. That sixth state, needs validation, is deliberately not one of the five dispositions.

Step five is the point. Eleven applications come back retain: no action, no spend. Nine rows are
contested - seven consolidate, one invest, one retire. Those nine are the shortlist a CIO
actually argues about, and they fall out of six overlap clusters the engine found: the two EHRs,
collaboration, ambient clinical documentation, ITSM, analytics, and HCM. In each cluster the
engine names the survivor and says why. The core EHR holds all three contested capabilities as
primary, with 29,400 active users, and scores retain independently - so the second EHR folds into
it rather than the other way round.

On how it runs: Python 3 with exactly one third-party dependency, openpyxl. No build step, no
package manager, nothing to compile. That command-line engine stays the reference implementation -
if the browser and the engine ever disagree, the engine is right by definition. Output is a
nine-sheet workbook, including a sanity-checks sheet and a sheet comparing our dispositions
against the client's own labels, which is how we found the places we disagree with her, and why.

Then we ran the same engine, unmodified, over a six-hundred-application portfolio. Seconds, not
weeks: the scoring step is a shade over two seconds, and about seven end to end including writing
every output workbook. Treat that as an order of magnitude rather than a benchmark - it is a
laptop-class measurement and the committed summary from a slower machine records nine and a
quarter seconds end to end. The important half of that run is the regression: the twenty
applications shared with the walkthrough came out of the six hundred with exactly the same
dispositions, none moved.

Now the web app, and here is what is real. Upload inventory opens a genuine file dialogue that
accepts .xlsx and .csv. Analyze portfolio scores that file in the browser and rewrites the page
from it - the headline figures, the savings bars, the decision mix, the guardrail counts, the
redundancy clusters, and the workbench table, where clicking a row opens the evidence rail rebuilt
from the computed result: four dimension scores with pass or fail, the pattern key, the priority,
the confidence and a written rationale. Three filters - capability, recommendation, critical
operation - narrow that table, and the headline figures deliberately do not move when you filter,
so a narrowed table can never be read as a changed total. There is no server, no build step and no
network call in the whole path: the file never leaves the machine, which also means this works on
a conference-centre wifi that does not.

We checked the browser against the engine rather than assuming it: on our own twenty-application
synthetic portfolio, 329 compared values matched, zero mismatches, at 22,057,000 of spend and
5,818,716 net. That is per application the four dimension scores, the four gate verdicts, the
pattern key, the disposition, the priority, the confidence and the net saving. Six hundred rows in
the browser parse in 189 milliseconds and score and render in 133 - fast enough that there is no
progress bar, because nothing waits.

One boundary to state plainly, and I would rather say it than have it found: it scores a portfolio
in our intake schema. It is not a universal reader. Hand it an inventory in a different column
vocabulary - as both raw Northstar workbooks are - and it refuses: it names the file, the row and
column count, and lists the headings it does look for, app_id, app_name, primary_capability,
ov_patient_care_criticality, th_supportability, c_cost_per_active_user_vs_peers, r_technical_risk.
Case, spaces, hyphens and underscores do not matter. Mapping another vocabulary onto those names
is a real piece of work and we have not automated it.

IF AN UPLOAD IS REFUSED ON STAGE, say this: "That is the tool refusing a file it cannot score
rather than inventing an answer for it - it wants our intake column names, and it just told you
which ones. Here is the same portfolio in that schema." Then upload
data/northstar/northstar-600-corrected-tool-vocabulary.csv, which is the file to demo from.

And to be exact about what I am claiming: the page is deployed at hack-team-12.vercel.app, and
everything I have just described about its behaviour is read out of the committed source in
src/apprat-ai-wireframe-v2.html and the measured comparison in docs/wireframe-README.md. Export
decisions and Generate executive readout are still deliberately inert - we left them visibly
unwired rather than faking them.
""")


# ================================================================ SLIDE 4
s4 = prs.slides.add_slide(BLANK)
y = header(s4, "Impact · Business value and path to market",
           "$17.41M net in year one at 600 applications")

BH = 1.28
rect(s4, ML, y, CW, BH, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.05)
rect(s4, ML, y + 0.18, 0.055, BH - 0.36, ACCENT)

tfm = textbox(s4, ML + 0.34, y + 0.20, 3.00, BH - 0.40, anchor=MSO_ANCHOR.MIDDLE)
para(tfm, "$17.41M", size=38, color=WHITE, bold=True, first=True, line=1.0)
para(tfm, "net first-year benefit, Northstar's 600", size=11, color=ONNAVY,
     space_before=4)

facts = [
    ("$25.82M",   "gross annual avoidable\n— recurs from year two"),
    ("$8.41M",    "one-time transition cost\nnetted out (32.6% of gross)"),
    ("301 of 600","need no action at all —\nretain, no spend"),
    ("$354.33M",  "portfolio run-rate the\nsavings come off"),
]
fx = ML + 3.50
fw = (CW - 3.50 - 0.16) / 4
for i, (num, lbl) in enumerate(facts):
    if i:
        rect(s4, fx - 0.05, y + 0.30, 0.008, BH - 0.60, DIVIDE)
    tff = textbox(s4, fx + 0.13, y + 0.24, fw - 0.24, BH - 0.48,
                  anchor=MSO_ANCHOR.MIDDLE)
    para(tff, num, size=18, color=WHITE, bold=True, first=True, line=1.0)
    para(tff, lbl, size=9.5, color=RGBColor(0xA8, 0xB8, 0xC6), space_before=4,
         line=1.16)
    fx += fw

y2 = y + BH + 0.30
CW3 = (CW - 0.90) / 3
cols = [ML, ML + CW3 + 0.45, ML + 2 * (CW3 + 0.45)]

yc = column_head(s4, cols[0], y2, CW3, "What changes operationally")
tf = textbox(s4, cols[0], yc, CW3, 2.6)
bullet(tf, [("The spread across 600: ", True, NAVY),
            ("retain 301, invest 174, retire 81, consolidate 44, replace 0.",
             False, None)],
       first=True, size=12.5)
bullet(tf, [("No-action outcomes get published, ", True, NAVY),
            ("with a route to object — not a meeting. 301 of the 600 need no "
             "further conversation; nobody defends all 301, anybody can "
             "challenge one.", False, None)], size=12.5)
bullet(tf, [("Only the contested minority ", True, NAVY),
            ("gets the full seven-step treatment, so scarce facilitation goes "
             "where it can change an answer — which is what makes 600 "
             "tractable, not just 20.", False, None)], size=12.5)

yc = column_head(s4, cols[1], y2, CW3, "Why it is repeatable")
tf = textbox(s4, cols[1], yc, CW3, 2.6)
bullet(tf, [("The model is portfolio-agnostic. ", True, NAVY),
            ("The engine was built against a different, independently "
             "generated 20-application portfolio; Northstar was built "
             "separately and scored through the same code with no changes.",
             False, None)], first=True, size=12.5)
bullet(tf, [("Nothing sector-specific is hard-coded. ", True, NAVY),
            ("Thresholds, weights and all 16 table rows are configuration, so "
             "a new sector is a re-tune, not a rebuild.", False, None)],
       size=12.5)
bullet(tf, [("So it drops into any Aberdeen engagement, ", True, NAVY),
            ("healthcare or not, as a repeatable offering rather than a "
             "one-off analysis.", False, None)], size=12.5)

yc = column_head(s4, cols[2], y2, CW3, "What the data limits · what's next")
tf = textbox(s4, cols[2], yc, CW3, 2.6)
bullet(tf, [("One cost input — consumption price variance — is absent from "
             "this client's source data, ", True, NAVY),
            ("so it is left unscored rather than assumed.", False, None)],
       first=True, size=12.5)
bullet(tf, [("Consolidation is gated by migration evidence, not redundancy "
             "alone: ", True, NAVY),
            ("overlap is wide — 35 groups — but only 44 applications can be "
             "shown to land somewhere. A limit of the data supplied, not of "
             "the method.", False, None)], size=12.5)
bullet(tf, [("Next: a real peer cost benchmark, ", True, NAVY),
            ("and closing the derived-column pipeline — 31 of 68 columns are "
             "still authored by hand.", False, None)], size=12.5)

footnote(s4, "Figures above are the 600-application Northstar run: "
             "northstar-600-corrected-summary.md. The 20-application "
             "walkthrough ($5.30M net) and our own 20-application dataset "
             "($5.82M net) are separate portfolios; nothing is blended. The "
             "CIO's 15% target, $53.15M, is not met on first-year net.")
page_tag(s4, 4, "Impact")

notes(s4, """
The money first, and this is the six-hundred-application portfolio, not a sample of it. Northstar's
600 applications carry a 354.33 million dollar annual run cost. We identify 25.816 million of gross
annual avoidable cost, net out 8.407 million of one-time transition cost - migration, interface
cutover, contract exit, which is 32.6 percent of the gross - and land at 17.409 million net in the
first year. From year two the gross recurs, because the transition cost is one time only. All of
those figures come from the engine run recorded in northstar-600-corrected-summary.md, and I
re-ran the scoring script to reproduce them rather than reading them off a slide.

The spread is retain 301, invest 174, retire 81, consolidate 44, replace 0. Nothing lands in
replace, because every application that might have been replaced clears its gates.

The twenty-application walkthrough you saw on the previous slide is the detailed version of the
same method - 43.25 million of run-rate, 5.30 million net, eleven retains and nine contested rows -
and it holds inside the six hundred: all twenty came out with the same dispositions they had, none
moved. Please keep the three portfolios apart if you ask me about them. The corrected 600-app
Northstar run is the headline. The 20-app Northstar run is the walkthrough. The team's own
20-application synthetic portfolio is the fixture we test the browser against. Three different
things; I am not adding them up.

Now the honesty points, because they are the reason to believe the rest. First, the target. The
CIO's stated savings target is fifteen percent of run cost, which is 53.15 million dollars. We do
not meet it on first-year net - not at six hundred applications and not at twenty. I would rather
say that plainly than reclassify a transition cost to make a number.

Second, one cost input is missing from her source data: consumption price variance. There is no
metered or consumption cost line anywhere in her workbook, so that single input is left unscored
across the portfolio rather than assumed. It carries the lowest weight in the model, and the other
inputs renormalise around it, so it does not condemn a row - it is one missing portfolio-wide
input, not a penalty applied per application. It is also why you will see a zero where the tool
would otherwise show a high-confidence split for this portfolio: that zero is absent data, not a
result, and I am not quoting a safe-savings figure for the six hundred at all. The twenty-
application portfolio does carry that input, which is why the workbook can split its number.

Third, consolidation. The functional overlap in this portfolio is wide - thirty-five overlap groups,
some with twenty members - but only forty-four applications are consolidated. That is deliberate:
consolidation is gated on migration evidence and a named survivor, not on redundancy alone. If we
cannot see that a capability actually lands somewhere, we do not switch anything off. So the
absorbable count is limited by the evidence in the data we were given, not by the method - richer
migration and dependency evidence would raise it, and that is an intake conversation, not a code
change.

What actually changes operationally matters more to me than the number. 301 of the 600 come back
retain - no action, no spend. Today each of those still costs you a meeting to conclude nothing.
Under AdvisAR they are published as decisions with the evidence attached and a route to object.
Nobody has to defend all 301; anybody can challenge one. And the contested minority gets the full
seven-step treatment - the workshops, the dependency checks, the sequencing. Scarce facilitation
goes only where it can change an answer, and that is exactly what makes six hundred tractable
rather than twenty.

Now path to market, which is the reusability line. The scoring engine was built against a
completely different twenty-application portfolio - a separately generated dataset with a
different shape. Northstar was constructed independently, by hand, as a test of the model, and it
ran through the same code with no changes at all. That is the claim: nothing about healthcare is
hard-coded. The thresholds, the weights and all sixteen rows of the disposition table are
configuration rather than branching logic, so taking this into manufacturing or financial
services is a re-tune of a config table, not a rebuild. It drops into an Aberdeen engagement as a
repeatable offering with a known intake list and a known deliverable, instead of a bespoke
analysis we rebuild every time.

One practical note for the demo. The file we upload on stage is
data/northstar/northstar-600-corrected-tool-vocabulary.csv - the same 600 applications translated
into the tool's own column names. I checked that it totals to the same money as the engine run:
25,816,000 gross and 17,409,000 net, to the dollar, with the same 301 / 174 / 81 / 44 / 0 spread.
So the figures on this slide and the figures on the screen are the same figures, and if they ever
diverge the engine run in the summary file is the one to trust.

Two things we owe you before this is client-ready. One: the peer cost band. Cost per active user
against peers is the single most influential input in the cost dimension, and right now it is
modelled from the portfolio itself rather than measured against an external benchmark. We label
it as modelled in every output, but a live engagement needs a real data source. Two: the pipeline
in front of the engine. Of the 68 columns the tool derives, 31 are currently authored by hand
rather than computed - AI classification, vendor end-of-support lookups, the clustering step, and
five band mappings. When those are supplied, all 37 implemented derivations reproduce exactly, on
every row, with zero mismatches. So the engine is right; the intake pipeline is the unfinished
part, and the five band mappings are the cheapest place to start.
""")


out = sys.argv[1] if len(sys.argv) > 1 else "AdvisAR-Hackathon-Deck.pptx"
prs.save(out)
print("saved", out)
