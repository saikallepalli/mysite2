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
        children: [new TextRun({ text: "Data source: Google Form responses, 1–23 July 2026 (updated dataset, 74 visits)", size: 22, color: DARKGREY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [new TextRun({ text: "74 unannounced visits · 74 centres · 9 Second Level Managers", size: 22, color: DARKGREY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 800 },
        children: [new TextRun({ text: "Prepared for: MIS, Dr. Reddy's Foundation", size: 22, color: DARKGREY })] }),

      H1("1. Executive Summary"),
      P("This update adds just one visit since the last refresh — 74 unannounced visits now on record between 1 and 23 July 2026 (up from 73, and 41 at first analysis) — the same nine Second Level Managers (SLMs), now covering 74 distinct centres. With only one new data point, this cycle is best read as confirmation that the picture from the last refresh is stable, not a new trend: every figure below moved by well under a percentage point."),
      P("Five headline findings this cycle:"),
      bullet("Alumni Hand-holding is still the single weakest control — 42 of 68 applicable visits (61.8%) flagged “Need Improvement,” essentially flat versus 62.7% at the last refresh and still well below the 66.7% seen two refreshes ago."),
      bullet("Weekly team meeting documentation (54.2%) and Employer Connect (43.7%) remain the next-largest process gaps — both essentially unchanged from 54.9% and 44.3% last refresh, holding onto the earlier 16-point improvement in Employer Connect rather than slipping back."),
      bullet("No new data-veracity red flags in the additional visit — the ghost-placement and DOB/age-mismatch findings remain confined to the same six centres identified across the last two refreshes (five under Amit Chakraborty, one under Gagandeep Singh Sachar), now unchanged for three consecutive updates.", { bold: true }),
      bullet("The activity-planner delay item flagged last cycle holds steady at 40.4% (19 of 47 applicable visits, unchanged) — still worth a dedicated look, but not worsening."),
      bullet("Phani's reporting-depth gap is unchanged: now at 14 visits (up 1) with notes still averaging just 16 words — the lightest documentation of any high-volume SLM, and still the clearest coaching priority flagged last cycle."),
      P("Aspirant Interaction remains the strongest section (6.9% Need Improvement), and zero visits across all 74 have reported an aspirant being asked for money for training or placement — the integrity signal from earlier cycles continues to hold at scale."),

      H1("2. Coverage Snapshot"),
      makeTable(
        ["Metric", "Value"],
        [
          ["Total visit responses", "74 (up from 73 at last refresh, 63 before that, 41 at first analysis)"],
          ["Unique centres visited", "74 (no repeat visits in the period)"],
          ["Second Level Managers reporting", "9"],
          ["Date range covered", "1 Jul 2026 – 23 Jul 2026"],
          ["Checklist items per visit", "~50 items across 4 sections + 4 free-text action logs + 3 reflective questions"],
          ["Programs represented", "GROW Youth, HQHCS, CFRSI, Grow-Green-Solar, Accenture/CITI/JPM-partnered Youth batches, and others"],
        ],
        [5400, 4100]
      ),

      H1("3. Section-Wise Compliance Scorecard"),
      makeTable(
        ["Section", "All is Well", "Need Improvement", "Not Applicable", "% Need Improvement"],
        [
          ["Data Check (veracity & records)", "414", "100", "78", "19.5%"],
          ["Infra Check", "411", "109", "72", "21.0%"],
          ["Process Check", "956", "283", "389", "22.8%"],
          ["Aspirant Interaction", "648", "48", "118", "6.9%"],
        ],
        [3600, 1550, 1900, 1550, 1900]
      ),
      P("All four sections are essentially flat versus the last refresh (Process Check 22.8% vs. 23.2%; Data Check 19.5% vs. 19.7%; Infra Check 21.0% vs. 20.7%; Aspirant Interaction 6.9% vs. 7.0%) — a single additional visit, as expected, doesn't move the picture. Process Check remains the weakest section by a clear margin."),

      H2("3.1 The 12 Weakest Checklist Items"),
      P("Ranked by count of “Need Improvement” flags among applicable visits (n = applicable visits)."),
      makeTable(
        ["Checklist Item", "Section", "NI / n", "% NI"],
        [
          ["Verify retention — Alumni Hand-holding Format updated as designed", "Data Check", "42 / 68", "61.8%"],
          ["AH holds a quality weekly team meeting with documented action points", "Process Check", "39 / 72", "54.2%"],
          ["AH regularly meets employers & updates Top 20 Employer format", "Process Check", "31 / 71", "43.7%"],
          ["Any pending infra repairs", "Infra Check", "29 / 74", "39.2%"],
          ["CA’s diary — guided daily by AH during planning days", "Process Check", "28 / 69", "40.6%"],
          ["Any scrap material cleared with support of Admin", "Infra Check", "25 / 72", "34.7%"],
          ["Status of digital infra — PCs, smartboards, tablets, laptops", "Infra Check", "25 / 74", "33.8%"],
          ["Aspirant Mock Interview Score Tracker updated in L&D Google Sheet", "Process Check", "24 / 70", "34.3%"],
          ["AH plans & shares monthly enrolment/placement numbers by the 3rd", "Process Check", "21 / 73", "28.8%"],
          ["AH announces next batch-start date on notice boards by the 3rd", "Process Check", "21 / 74", "28.4%"],
          ["Planning for next batch-start announcement", "Process Check", "20 / 73", "27.4%"],
          ["Refer to the activity planner — check if any activities are delayed", "Process Check", "19 / 47", "40.4%"],
        ],
        [5200, 1600, 1400, 1300]
      ),
      P("The list is unchanged in composition from the last refresh, with one minor reshuffle: scrap-material clearance and digital-infra status each picked up one more flag and now edge out the Mock Interview tracker for 6th/7th place. Otherwise the ranking and rates are stable."),

      H1("4. Deep Dive by Focus Area"),

      H2("4.1 Data Veracity"),
      P("No new red-flag cases were found among the 10 visits added since the last refresh. The six previously identified cases stand unchanged:"),
      quote("Danapur — “11 Aspirant verification initiated... 4 Aspirants found working and 7 doubtful and has not joined the Organization but Placement has been uploaded and shown working for minimum 2 weeks in each case — Disciplinary issue.”"),
      quote("Ghaziabad — “Called 5 Aspirants – 2 Candidates confirmed they had not joined any Organization but shown as Placed is Cogito (346798 & 348902)... Cogito Placement needs Physical visit by SM in next visit.”"),
      quote("Bareilly — “Contacted 4 Aspirants – 1 Parent reverted and shared the candidate left job in 7 days (Anjalee)...”"),
      P("Plus DOB/age-document mismatches at Dumdum, Kidwainagar, Alambagh, Lucknow-Indiranagar (all Amit Chakraborty) and Vadodara Citi (Gagandeep Singh Sachar). The concentration under a single SLM's verification method is unchanged from the last cycle — this remains the strongest evidence that detection depth, not underlying prevalence, explains why these cases cluster where they do."),
      bullet("Alumni retention/placement follow-up is still the weakest data-veracity control at 61.8% Need Improvement — essentially unchanged since the last refresh, holding the ~4-point improvement gained two refreshes ago."),
      bullet("Aspirant Registration Forms and staff-attendance/geo-tag issues continue to appear at a similar rate to previous cycles — no material change."),

      H2("4.2 Process Gaps"),
      P("Process Check remains the weakest section (22.8%, essentially flat versus 23.2% last refresh). Nothing material shifted this cycle:"),
      bullet("Weekly team meeting documentation is flat (54.2% vs. 54.9%) — still the second-largest gap system-wide."),
      bullet("Employer Connect is holding its prior gain (43.7% vs. 44.3% last refresh, both well below the 60.5% seen two refreshes ago) — the push on employer-meeting cadence looks to be sticking rather than a one-off blip."),
      bullet("CA diary review by AHs (40.6%) and Mock Interview tracker maintenance (34.3%) remain persistent, moderate-severity gaps, both essentially flat this cycle."),
      bullet("The activity-planner delay check is the one area that got materially worse in relative ranking, newly entering the top 12 — worth a dedicated look next cycle to see if this is a real process slippage or an artifact of more SLMs starting to check this item consistently."),

      H2("4.3 Batch Screening (where applicable)"),
      P("Screening remains applicable mainly to technical/GROW-Tech programs, consistent with prior cycles."),
      bullet("Counselling and Screening Forms continue to show a high compliance rate, consistent with the ~93% seen previously."),
      bullet("Target-group enrolment issues (wrong age/education profile) persist at a similar low-to-moderate rate as before, still worth routine monitoring rather than urgent escalation."),

      H2("4.4 Aspirant Interaction"),
      P("Still the strongest section (6.9% Need Improvement, flat versus 7.0%). Zero visits across all 74 have reported an aspirant being asked for money for training or placement — this clean result has now held across four consecutive data refreshes (41, 63, 73, and 74 visits), which is a meaningful integrity signal in its own right."),
      bullet("LMS/Skillfy app enrolment and referral-scheme awareness remain the two softer, easily-closed gaps in this section, at rates broadly consistent with prior cycles."),

      H2("4.5 Hostel Visit"),
      P("Hostel-linked centres remain a minority of the total, so this focus area continues to have a small sample. Findings are consistent with prior cycles — food quality and cleanliness are the recurring issues where hostels exist — and should still be read centre-by-centre rather than as a system trend."),

      H2("4.6 Employer Connect"),
      P("This remains the most-improved focus area of the last two refreshes: 43.7% Need Improvement, essentially flat versus 44.3% last cycle and still well down from 60.5% two refreshes ago. It remains a top-3 gap system-wide, so the improvement should be reinforced rather than treated as solved, but it is holding rather than eroding."),

      H2("4.7 Centre Hand-holding"),
      P("The hand-holding backbone (Alumni Hand-holding, weekly-meeting documentation, CA-diary mentoring, GROW Manual completion) shows a similar pattern to before, with modest improvement across the board:"),
      bullet("Alumni Hand-holding remains the single weakest item system-wide (61.8%, essentially flat versus 62.7% last refresh)."),
      bullet("GROW Manual Course completion gaps still concentrate in newly joined trainers/CAs and LMS technical issues — an L&D/IT coordination item as before."),
      bullet("AH/CA hiring delays or resignations were explicitly flagged at 3 centres this cycle (a similar count to before) — continue to treat these as a trigger for a follow-up visit inside 2–3 weeks."),

      H1("5. Insights on Second Level Managers"),
      P("Updated visit volumes, Need Improvement / Not Applicable rates, and average free-text note length per SLM (only Phani's row changed this cycle — +1 visit):"),
      makeTable(
        ["SLM", "Visits", "% Need Improvement", "% Not Applicable", "Avg. words in notes"],
        [
          ["S Jakir", "15", "13.8%", "12.5%", "102"],
          ["Phani", "14", "14.0%", "20.7%", "16"],
          ["Barna", "11", "21.8%", "21.7%", "47"],
          ["Amit Chakraborty", "10", "22.2%", "19.2%", "287"],
          ["Gagandeep Singh Sachar", "10", "21.3%", "20.4%", "131"],
          ["Bhargav Challa", "6", "23.2%", "20.7%", "124"],
          ["Mahaboob Basha", "3", "15.6%", "12.9%", "39"],
          ["Raghavender Rao", "3", "9.1%", "10.2%", "18"],
          ["Sanjay Kumar Tarai", "2", "32.1%", "17.3%", "388"],
        ],
        [3100, 1300, 1900, 1600, 2550]
      ),
      P("What's changed since the last refresh: essentially nothing except Phani's row (+1 visit, figures barely moved) — the other eight SLMs are unchanged from last cycle:", { bold: true }),
      bullet("S Jakir has become the highest-volume SLM (15 visits, up from 2 originally) while maintaining a reasonably thorough note style (102 words average) and one of the lowest Need Improvement rates (13.8%) — currently the strongest combination of volume and depth in the group."),
      bullet("Phani is now at 14 visits (up from 2 originally, 13 last refresh) with note depth still not keeping pace (16 words average, among the thinnest in the group). The volume-to-depth gap flagged last cycle persists unchanged — still the clearest coaching priority."),
      bullet("Amit Chakraborty and Sanjay Kumar Tarai remain the two SLMs whose detailed, evidence-heavy notes (287 and 388 words respectively) are the ones actually surfacing hard findings — this pattern is unchanged and still supports rolling their verification method out more broadly."),
      bullet("Raghavender Rao's profile is unchanged in character (lowest Need Improvement rate, shortest notes, lowest NA usage) — still the profile most in need of an independent quality check on the visit itself rather than just the write-up."),

      H1("6. Recommendations"),
      bullet("Keep pushing the alumni-verification-call standard modelled on Amit Chakraborty's method — it remains the only reliable way this dataset has caught fabricated placements, and no other SLM's notes have surfaced an equivalent finding across three refreshes.", { bold: true }),
      bullet("Keep reinforcing the Employer Connect gains from two refreshes ago — the improvement is holding (43.7%) rather than eroding, but it's still a top-3 gap; keep the weekly employer-meeting cadence and Top 20 Employer format review as a standing SLM checklist item rather than easing off."),
      bullet("Coach Phani specifically on note depth — now at 14 visits with still the widest volume-to-depth gap of any SLM; the structured template (issue / evidence / owner / date) recommended last cycle has not yet visibly landed."),
      bullet("Still worth investigating the activity-planner delay flag (40.4%, 19 of 47 visits, unchanged this cycle) to confirm whether it reflects a genuine process slippage or just more consistent checking."),
      bullet("Continue treating AH/CA vacancy as a red-flag trigger for a follow-up visit inside 2–3 weeks — this pattern held again this cycle."),
      bullet("Keep Alumni Hand-holding as the top fix-first priority — still the single worst-performing control despite the improvement, and still the Foundation's main defence against inflated placement numbers."),

      new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "— End of report —", italics: true, color: DARKGREY, size: 20 })] }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/sessions/affectionate-busy-goodall/mnt/outputs/work/DRF_Centre_Visit_Analysis_Report_v3.docx", buf);
  console.log("done");
});
