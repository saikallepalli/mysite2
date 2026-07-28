const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, Header, Footer, PageNumber,
  LevelFormat, convertInchesToTwip
} = require("docx");

const NAVY = "1F3864";
const GOLD = "B8860B";
const GREY = "F2F2F2";
const DARKGREY = "595959";

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { color: NAVY, space: 4, style: BorderStyle.SINGLE, size: 6 } },
    children: [new TextRun({ text, bold: true, color: NAVY, size: 30 })],
  });
}
function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, color: GOLD, size: 24 })],
  });
}
function P(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 276 },
    children: [new TextRun({ text, size: 21, ...opts })],
  });
}
function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 90, line: 270 },
    children: [new TextRun({ text, size: 21, ...opts })],
  });
}
function quote(text) {
  return new Paragraph({
    indent: { left: 360 },
    border: { left: { color: GOLD, space: 8, style: BorderStyle.SINGLE, size: 12 } },
    spacing: { after: 160 },
    children: [new TextRun({ text, italics: true, size: 20, color: DARKGREY })],
  });
}

function cell(text, { bold = false, shade = null, color = "000000", width = 1000, align = AlignmentType.LEFT } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, size: 19, color })],
    })],
  });
}

function makeTable(headers, rows, widths) {
  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: headers.map((h, i) => cell(h, { bold: true, shade: NAVY, color: "FFFFFF", width: widths[i], align: AlignmentType.CENTER })),
  });
  const dataRows = rows.map((r, idx) => new TableRow({
    cantSplit: true,
    children: r.map((v, i) => cell(v, { width: widths[i], shade: idx % 2 === 1 ? GREY : null, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
  }));
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...dataRows],
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.2) } } } }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { color: "BFBFBF", space: 4, style: BorderStyle.SINGLE, size: 4 } },
          children: [new TextRun({ text: "Dr. Reddy's Foundation — Unannounced Centre Visit Analysis", size: 16, color: DARKGREY })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Page ", size: 16, color: DARKGREY }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, color: DARKGREY })],
        })],
      }),
    },
    children: [
      new Paragraph({ spacing: { before: 1200 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Unannounced Monthly Centre Visit", bold: true, size: 48, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "Second Level Manager Field Audit — Analysis & Insights", bold: true, size: 32, color: GOLD })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 40 },
        children: [new TextRun({ text: "Data source: Google Form responses, 1–21 July 2026", size: 22, color: DARKGREY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [new TextRun({ text: "41 unannounced visits · 41 centres · 9 Second Level Managers", size: 22, color: DARKGREY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 800 },
        children: [new TextRun({ text: "Prepared for: MIS, Dr. Reddy's Foundation", size: 22, color: DARKGREY })] }),

      H1("1. Executive Summary"),
      P("Between 1 and 21 July 2026, nine Second Level Managers (SLMs) completed 41 unannounced visits across 41 skilling centres, covering roughly 50 checklist points per visit plus written action points and three reflective questions. This is a genuinely useful dataset — completion is high and several SLMs are writing rich, evidence-based notes — but it also surfaces real risk that needs leadership attention, most notably around placement-data integrity."),
      P("Four headline findings:"),
      bullet("Alumni Hand-holding is the single weakest process in the entire checklist — 26 of 39 applicable visits (66.7%) flagged “Need Improvement,” meaning most centres are not systematically following up with placed and unplaced alumni as designed."),
      bullet("Employer Connect is the second-biggest gap — AHs are not regularly meeting employers or maintaining the Top 20 Employer format in 60.5% of applicable visits, directly threatening placement pipelines."),
      bullet("Data veracity red flags were found at multiple centres: alumni verification calls uncovered aspirants shown as “placed” who had in fact not joined or already left their jobs, and several centres had DOB/age document mismatches (including one under-age enrolment caught and removed). These are disciplinary-grade findings, not routine process gaps.", { bold: true }),
      bullet("Manager-level reporting quality varies sharply. Some SLMs (notably Amit Chakraborty and Sanjay Kumar Tarai) write long, specific, verification-heavy notes that actively catch fraud and documentation errors; others log the same volume of visits with far shorter notes and higher “Not Applicable” usage, which is worth a calibration conversation rather than an assumption that their centres are simply cleaner."),
      P("Aspirant-facing basics are in good shape: only 8.4% of Aspirant Interaction checks needed improvement, and zero visits reported aspirants being asked for money, which is an important integrity signal in itself. Infrastructure and data-check items sit in the middle (20–24% Need Improvement), driven mostly by ageing furniture/electronics and incomplete registration documentation."),

      H1("2. Coverage Snapshot"),
      makeTable(
        ["Metric", "Value"],
        [
          ["Total visit responses", "41"],
          ["Unique centres visited", "41 (no repeat visits in the period)"],
          ["Second Level Managers reporting", "9"],
          ["Date range covered", "1 Jul 2026 – 21 Jul 2026"],
          ["Checklist items per visit", "~50 items across 4 sections + 4 free-text action logs + 3 reflective questions"],
          ["Programs represented", "GROW Youth, HQHCS, CFRSI, Grow-Green-Solar, Accenture/CITI/JPM-partnered Youth batches, and others"],
        ],
        [5400, 4100]
      ),
      P("The four checklist sections and their approximate time budgets, as designed in the form, are: 1) Data Check (~30 min), 2) Infra Check (~15 min), 3) Process Check (~30 min), 4) Aspirant Interaction (~45 min)."),

      H1("3. Section-Wise Compliance Scorecard"),
      P("For every checklist item, SLMs marked “All is Well,” “Need Improvement,” or “Not Applicable.” The table below aggregates all items within each section (Not Applicable rows excluded from the % Need Improvement base)."),
      makeTable(
        ["Section", "All is Well", "Need Improvement", "Not Applicable", "% Need Improvement"],
        [
          ["Data Check (veracity & records)", "219", "68", "41", "23.7%"],
          ["Infra Check", "225", "58", "45", "20.5%"],
          ["Process Check", "475", "198", "229", "29.4%"],
          ["Aspirant Interaction", "347", "32", "72", "8.4%"],
        ],
        [3600, 1550, 1900, 1550, 1900]
      ),
      P("Process Check — the section covering employer connect, batch planning, weekly meetings, CA diaries and GROW Manual completion — is proportionally the weakest, with nearly 3 in 10 applicable checks flagged. Aspirant Interaction is comfortably the strongest, suggesting frontline delivery and aspirant welfare are being handled better than back-office process discipline."),

      H2("3.1 The 10 Weakest Checklist Items"),
      P("Ranked by % Need Improvement among visits where the item was applicable (n = applicable visits)."),
      makeTable(
        ["Checklist Item", "Section", "% NI", "n"],
        [
          ["Verify retention — Alumni Hand-holding Format updated as designed", "Data Check", "66.7%", "39"],
          ["AH holds a quality weekly team meeting with documented action points", "Process Check", "61.5%", "39"],
          ["AH regularly meets employers & updates Top 20 Employer format", "Process Check", "60.5%", "38"],
          ["CA’s diary — guided daily by AH during planning days", "Process Check", "55.3%", "38"],
          ["Aspirant Mock Interview Score Tracker updated in L&D Google Sheet", "Process Check", "55.0%", "40"],
          ["Any pending infra repairs", "Infra Check", "43.9%", "41"],
          ["AH announces next batch-start date on notice boards by the 3rd", "Process Check", "43.9%", "41"],
          ["Planning for next batch-start announcement", "Process Check", "39.0%", "41"],
          ["All centre teams completed GROW Manual Course on Skillfy", "Process Check", "39.0%", "41"],
          ["AH plans & shares monthly enrolment/placement numbers by the 3rd", "Process Check", "37.5%", "40"],
        ],
        [5300, 1600, 1500, 1100]
      ),
      P("Every one of the ten weakest items is a management-discipline / cadence issue, not a training-delivery or aspirant-welfare issue — which is reassuring in one sense (aspirants are being looked after) but flags that AHs are not being held to the reporting rhythm the process expects (weekly meetings, employer visits, batch-start notices, portal updates by fixed dates)."),

      H1("4. Deep Dive by Focus Area"),

      H2("4.1 Data Veracity"),
      P("This is the highest-stakes section in the checklist because it is the one place designed to catch fabricated or inflated numbers. Overall, Data Check items were flagged in 23.7% of applicable checks, but the pattern within the section matters more than the average:"),
      bullet("Alumni retention/placement follow-up (66.7% Need Improvement) is by far the weakest data-veracity control — without it, the Foundation has no independent check on whether reported placements are holding."),
      bullet("Aspirant Registration Forms had issues in 31.7% of visits (13 of 41) — commonly missing signatures, missing Aadhaar back-page copies, or acceptance of computer-generated mark sheets in place of original certificates (flagged explicitly at 4 centres)."),
      bullet("Staff attendance registers / geo-tagged photos were flagged in 30.8% of visits — several notes mention specific missing dates of geo-tagged attendance photos, which weakens the audit trail for aspirant presence."),
      bullet("Placement-verification calls (random alumni calling) were the least-flagged item on paper (7.9%), but the free-text notes tell a more serious story — see the red-flag cases below, which a simple Yes/No tick does not fully capture."),
      P("Red-flag cases surfaced through SLM verification calls (direct quotes from the form):", { bold: true }),
      quote("Danapur — “11 Aspirant verification initiated... 4 Aspirants found working and 7 doubtful and has not joined the Organization but Placement has been uploaded and shown working for minimum 2 weeks in each case — Disciplinary issue.”"),
      quote("Ghaziabad — “Called 5 Aspirants – 2 Candidates confirmed they had not joined any Organization but shown as Placed is Cogito (346798 & 348902), 1 candidate is SP but shown as placed in Cogito (353697)... Cogito Placement needs Physical visit by SM in next visit.”"),
      quote("Bareilly — “Contacted 4 Aspirants – 1 Parent reverted and shared the candidate left job in 7 days (Anjalee)...”"),
      quote("Kidwainagar — “1 Below age candidate (Candidate ID-366421) has been uploaded in Portal with wrong DOB in Batch-87 — have removed the candidate from Batch and discussed with Team on disciplinary issue.”"),
      P("All four of these findings came from centres visited by the same manager, Amit Chakraborty, whose notes consistently include named candidate IDs, call outcomes, and explicit disciplinary framing. Six visits in total (5 under Amit Chakraborty, 1 under Gagandeep Singh Sachar) flagged DOB/age-document mismatches. Taken together, this is the clearest actionable signal in the dataset: some centres are carrying placement records that do not survive a phone call, and the rate at which this is being caught tracks the rigor of the manager’s verification, not necessarily the underlying rate of the problem. The Foundation should treat this as a sampling issue — if only one SLM out of nine is consistently doing physical/telephonic placement verification at this level of depth, the true prevalence across all 41 centres is very likely higher than the 3–4 cases visible here."),

      H2("4.2 Process Gaps"),
      P("Process Check is the weakest section overall (29.4% Need Improvement) and the gaps cluster around cadence and documentation rather than substance:"),
      bullet("Weekly team meetings are happening in most centres but are frequently undocumented — no Minutes of Meeting, no photo shared, or MoM kept only in hard copy without action points/ownership/timelines (61.5% of visits)."),
      bullet("Employer connect is the most consequential process gap: AHs are not meeting employers at the expected cadence and the Top 20 Employer format is stale or incomplete in 60.5% of visits — this is the leading indicator that will show up as a placement shortfall next quarter if unaddressed."),
      bullet("CA diaries are not being reviewed daily by AHs in 55.3% of visits — several notes mention CAs sharing plans only informally over WhatsApp rather than maintaining the diary as designed."),
      bullet("GROW Manual Course completion on Skillfy is pending for at least one team member in 39% of visits, most often newly joined trainers/CAs or staff whose courses are “frozen” on the LMS and need L&D intervention."),
      bullet("Batch-start dates are missing from notice boards in 43.9% of visits, and next-batch planning itself was flagged in 39% — both feed directly into the “batch on time” goal called out in the reflective questions."),
      bullet("AH/CA hiring delays and resignations were explicitly called out at 3 centres (Berhampur, Kidwainagar, Vadodara Citi), each triggering cascading gaps in almost every other process area for that visit — vacant AH/CA roles are a strong leading indicator to watch centre-by-centre."),

      H2("4.3 Batch Screening (where applicable)"),
      P("Screening-related items were only applicable to a subset of centres (mostly technical/GROW-Tech programs), which is expected given most visited batches are GROW Youth programs where full technical screening does not apply:"),
      bullet("Counselling and Screening Forms of recently started batches were in order in 38 of 41 visits (92.7%); the 3 exceptions involved missing AH/aspirant signatures on hard-copy forms."),
      bullet("“100% of enrolled students undergone screening (technical courses)” was applicable at only 14 centres, and all 14 were compliant — a clean result, but the small base (14 of 41) means screening discipline for technical programs has not really been stress-tested this cycle."),
      bullet("Target-group enrolment (right age/education profile) was flagged in 19.5% of visits, including B.Tech candidates found enrolled in a program meant for a different target group, and a small number of near-boundary age cases needing special-consideration approval."),

      H2("4.4 Aspirant Interaction"),
      P("This section is the strongest in the audit (8.4% Need Improvement across 11 items) and the qualitative signal is reassuring:"),
      bullet("Zero visits reported any aspirant being asked for money for training or placement — across all 41 unannounced checks, this is a meaningful clean result on the single most serious integrity question asked directly to aspirants."),
      bullet("All aspirants confirmed they were told the kind of job to expect during counselling, and none reported disinterest in jobs without being asked why they enrolled — counselling quality at the point of entry looks sound."),
      bullet("The two softer gaps are LMS/Skillfy app enrolment (6 visits: 15%) — several aspirants were practising only on desktops rather than their own phones — and referral-scheme awareness (7 visits: 17.5%), both easily closed with a one-time orientation session, which several SLMs already did on the spot during the visit."),

      H2("4.5 Hostel Visit"),
      P("Only a minority of centres in this cycle have an attached hostel, so this focus area has a much smaller sample: 6–9 applicable visits out of 41 depending on the specific item."),
      bullet("Where hostels exist, food quality and cleanliness were the main issues raised (flagged directly in text at 3 centres), with SLMs escalating to hostel vendors on the spot in at least one case."),
      bullet("Aspirant Hostel Forms were properly filled in 6 of 7 applicable visits; aspirants themselves raised hostel-quality concerns in 2 of 9 applicable interaction checks."),
      bullet("Given the small base, hostel findings should be read centre-by-centre rather than as a system-wide trend — but every hostel-linked centre this cycle had at least one open item, so a dedicated hostel-only audit for that subset of centres would be worthwhile."),

      H2("4.6 Employer Connect"),
      P("Flagged as the second-worst item in the entire checklist (60.5% Need Improvement, 23 of 38 applicable visits) — this deserves to be called out as its own priority rather than folded into general process gaps, given its direct line to placement outcomes."),
      bullet("The most common issue is the Top 20 Employer format not being kept current, followed by AHs not meeting the expected number of employers per week (one SLM guidance line asked an AH to meet “5 Employers a week” going forward)."),
      bullet("New-employer portal registration is, by contrast, in good shape (94.7% compliant) — the gap is in relationship maintenance and cadence, not in the one-time registration step."),
      bullet("In the reflective “placement %” question, the single most common commitment SLMs made themselves was to personally share employer leads from their own network — a sign that SLMs are actively compensating for AH-side employer-connect gaps rather than only flagging them."),

      H2("4.7 Centre Hand-holding"),
      P("Read together, Alumni Hand-holding compliance, weekly-meeting documentation, CA-diary mentoring and GROW Manual completion form the “hand-holding backbone” of centre management — and it is the area needing the most support:"),
      bullet("Alumni Hand-holding Format updates are the weakest single item system-wide (66.7% Need Improvement) — typically batches updated only through Contact 1 or 2 of a multi-contact tracking format, or not started at all for recently completed batches."),
      bullet("Weekly team meetings are happening in most centres but rarely produce a proper, ownership-and-timeline MoM — several SLMs had to introduce their own Google Sheet MoM template on the spot because none existed."),
      bullet("GROW Manual Course completion gaps concentrate in newly joined trainers/CAs and staff facing LMS technical issues (“frozen” courses) — this is an L&D/IT coordination fix as much as an AH accountability issue."),
      bullet("Where an AH had recently resigned or was newly appointed (3 centres), hand-holding gaps compounded across almost every checklist section — confirming that AH continuity is the single biggest predictor of overall centre health in this dataset."),

      H1("5. Insights on Second Level Managers"),
      P("The following table summarises each SLM’s visit volume, the share of checklist items they marked “Need Improvement” vs. “Not Applicable,” and the average length of their written action-point notes (across the 4 free-text fields per visit) — a rough proxy for how much verification and detail went into each visit."),
      makeTable(
        ["SLM", "Visits", "% Need Improvement", "% Not Applicable", "Avg. words in notes"],
        [
          ["Amit Chakraborty", "9", "24.1%", "19.0%", "307"],
          ["Sanjay Kumar Tarai", "1", "34.1%", "16.3%", "468"],
          ["Bhargav Challa", "3", "33.0%", "23.8%", "186"],
          ["Barna", "9", "24.6%", "22.4%", "52"],
          ["Gagandeep Singh Sachar", "10", "21.3%", "20.4%", "131"],
          ["S Jakir", "2", "12.9%", "13.3%", "154"],
          ["Phani", "2", "14.5%", "22.4%", "48"],
          ["Mahaboob Basha", "3", "15.6%", "12.9%", "39"],
          ["Raghavender Rao", "2", "11.0%", "7.1%", "19"],
        ],
        [3100, 1300, 1900, 1600, 2550]
      ),
      P("Three patterns stand out:", { bold: true }),
      bullet("Depth of verification varies more than centre quality likely does. Amit Chakraborty and Sanjay Kumar Tarai write by far the longest, most specific notes (300–470 words on average, with named Batch/Candidate IDs and call outcomes) and, not coincidentally, are the only two SLMs whose reports surfaced hard disciplinary findings (ghost placements, under-age enrolment). This strongly suggests the true issue rate at other SLMs’ centres is under-reported rather than genuinely lower — their % Need Improvement is similar or even lower, but their notes rarely go deep enough to catch a fabricated placement."),
      bullet("Barna and Gagandeep Singh Sachar cover the most ground (9 and 10 visits respectively) but write comparatively short notes (52 and 131 words) relative to Amit’s 307 — worth a coaching conversation on note quality, since these two SLMs alone account for nearly half of all visits this cycle and set the tone for what “thorough” looks like across the program."),
      bullet("Raghavender Rao’s two visits have the lowest Need Improvement rate (11.0%) but also by far the shortest notes (19 words average) and lowest NA usage — combined with nearly empty action-point fields, this looks more like light-touch reporting than genuinely clean centres, and is the one profile in this table that most needs a quality check on the visit itself, not just the write-up."),
      P("None of this is a judgement on centre performance — it is a reporting-quality signal. The Foundation should treat SLM note depth as its own KPI: a visit that ticks “All is Well” 50 times with no elaboration is far less useful than one that ticks “All is Well” 40 times and explains the 10 exceptions with names, dates and next steps."),

      H1("6. Recommendations"),
      bullet("Make random alumni/placement verification calls (not just a tick-box) mandatory practice for every SLM, modelled on Amit Chakraborty’s method — named candidate IDs, call outcome, and explicit escalation when a candidate denies working. Consider a quarterly cross-check where one SLM independently re-verifies a sample of another SLM’s “All is Well” placements.", { bold: true }),
      bullet("Fix Alumni Hand-holding first — it is the single worst-performing control (66.7% Need Improvement) and it is also the Foundation’s main defence against inflated placement numbers. A simple weekly reminder/dashboard per AH, tied to the Contact 1/2/3 cadence, would likely close most of this gap."),
      bullet("Put Employer Connect on a written cadence (e.g., a minimum number of employer meetings per week, tracked in the Top 20 Employer sheet and reviewed by SLMs at each visit) — it is the process gap most directly linked to future placement %."),
      bullet("Standardise the weekly-meeting MoM format across all centres so SLMs are not building their own Google Sheets on the spot, and make photo/WhatsApp submission of the MoM a hard requirement, not a request."),
      bullet("Escalate GROW Manual Course completion gaps to L&D/IT as a system issue — multiple notes mention courses “frozen” on Skillfy for specific staff, which is outside the AH’s control to fix alone."),
      bullet("Treat AH/CA vacancy as a red-flag trigger — every centre with a recent resignation or vacancy showed compounding gaps across sections; these centres should get a follow-up visit inside 2–3 weeks rather than waiting for the next monthly cycle."),
      bullet("Coach note-taking quality alongside visit compliance. Share Amit Chakraborty’s and Sanjay Kumar Tarai’s action-point write-ups as the internal benchmark for what a useful unannounced-visit report looks like, and consider a minimum word-count or structured template (issue / evidence / owner / date) for the free-text fields."),
      bullet("Run a dedicated documentation clean-up drive on registration forms — computer-generated mark sheets and DOB mismatches between Aadhaar and school certificates were flagged at multiple centres and are an easy compliance/audit risk to close proactively."),

      new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "— End of report —", italics: true, color: DARKGREY, size: 20 })] }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/sessions/affectionate-busy-goodall/mnt/outputs/DRF_Centre_Visit_Analysis_Report.docx", buf);
  console.log("done");
});
