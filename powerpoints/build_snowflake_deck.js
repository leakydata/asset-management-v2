const fs = require("fs");
const pptxgen = require("pptxgenjs");

const BG = "0F1B2A";
const PANEL = "16283F";
const PANEL2 = "12314A";
const INK = "E8EEF5";
const DIM = "9FB3C8";
const YEL = "FFC72C";
const BLUE = "4DA3FF";
const TEAL = "2EC4B6";
const GREEN = "34C759";
const RED = "FF5A5F";
const ORANGE = "FF9F1C";
const GRAY = "8496AB";

const W = 13.33;
const F = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

function base(title, kicker) {
  const s = pres.addSlide();
  s.background = { color: BG };
  if (kicker) {
    s.addText(kicker.toUpperCase(), {
      x: 0.6, y: 0.32, w: 12.1, h: 0.3, margin: 0,
      fontFace: F, fontSize: 12, bold: true, color: YEL, charSpacing: 3,
    });
  }
  if (title) {
    s.addText(title, {
      x: 0.6, y: 0.58, w: 12.1, h: 0.75, margin: 0,
      fontFace: F, fontSize: 32, bold: true, color: INK,
    });
  }
  return s;
}

function panel(s, x, y, w, h, fill) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: fill || PANEL }, line: { color: "223650", width: 1 },
  });
}

function dot(s, x, y, color, size) {
  const d = size || 0.16;
  s.addShape("ellipse", { x, y, w: d, h: d, fill: { color }, line: { type: "none" } });
}

// ---------------------------------------------------------------- 1 · title
{
  const s = pres.addSlide();
  s.background = { color: BG };
  s.addText("CCAT Reconciliation in Snowflake", {
    x: 0.9, y: 2.15, w: 11.5, h: 1.0, margin: 0,
    fontFace: F, fontSize: 44, bold: true, color: INK,
  });
  s.addText("Detect, fix, and prove it: keeping our equipment records and Caterpillar's registry in sync", {
    x: 0.9, y: 3.2, w: 11.0, h: 0.6, margin: 0,
    fontFace: F, fontSize: 20, color: DIM,
  });
  const colors = [GREEN, BLUE, TEAL, ORANGE, RED, YEL];
  colors.forEach((c, i) => dot(s, 0.95 + i * 0.34, 4.05, c, 0.2));
  s.addText("Cleveland Brothers  ·  Data & Analytics", {
    x: 0.9, y: 6.7, w: 8, h: 0.4, margin: 0, fontFace: F, fontSize: 13, color: DIM,
  });
  s.addNotes(
    "SCRIPT: Good morning. This is about a problem every dealer has: our equipment records and Caterpillar's official ownership registry, CCAT, drift apart. Machines missing, machines on the wrong customer, machines another dealer still claims. What I'll show you is the system we built inside Snowflake, which we already own, that finds every one of those differences, fixes the safe ones automatically, routes the risky ones to people, and keeps permanent proof of everything. " +
    "CUE: The six dots are the six outcomes a machine can have. They'll come back on slide 5."
  );
}

// ------------------------------------------- 1b · why: Trifecta and DNA
{
  const s = base("The scoreboard this feeds: Trifecta and DNA", "why it matters");
  panel(s, 0.6, 1.7, 5.9, 4.15);
  s.addText("TRIFECTA", { x: 0.9, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 18, bold: true, color: YEL });
  s.addText("Caterpillar's data-quality score. A customer only counts as valid when ALL legs pass: CCID, company data, a valid contact, and 100% asset match on priority assets.", {
    x: 0.9, y: 2.42, w: 5.3, h: 1.15, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
  s.addText("92.8%", { x: 0.9, y: 3.7, w: 2.4, h: 0.7, margin: 0, fontFace: F, fontSize: 40, bold: true, color: ORANGE });
  s.addText("asset match conformance (2025 baseline) — the leg our data governance team is focused on, and the one still maintained by hand today", {
    x: 3.3, y: 3.72, w: 2.95, h: 1.3, margin: 0, fontFace: F, fontSize: 12.5, color: DIM, lineSpacing: 16,
  });
  s.addText("Worth 15% of the dealer excellence scorecard", { x: 0.9, y: 5.25, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 13.5, bold: true, color: INK });
  panel(s, 6.8, 1.7, 5.9, 4.15);
  s.addText("DNA — DATA NOTIFICATION ALERTS", { x: 7.1, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 18, bold: true, color: TEAL });
  s.addText("Caterpillar flags our prioritized data issues — asset alerts and customer alerts — sends us the top ones by value, and grades how well we resolve them.", {
    x: 7.1, y: 2.42, w: 5.3, h: 1.15, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
  s.addText("10%", { x: 7.1, y: 3.7, w: 2.0, h: 0.7, margin: 0, fontFace: F, fontSize: 40, bold: true, color: TEAL });
  s.addText("of the scorecard rides on DNA resolution — and many of the asset alerts trace back to the same root cause: ownership records that do not match reality", {
    x: 9.2, y: 3.72, w: 3.05, h: 1.3, margin: 0, fontFace: F, fontSize: 12.5, color: DIM, lineSpacing: 16,
  });
  panel(s, 0.6, 6.1, 12.1, 0.85, PANEL2);
  s.addText([
    { text: "Both metrics stand on the same ground:  ", options: { bold: true, color: YEL } },
    { text: "asset ownership records in CCAT. That is exactly what this project manages.", options: { color: INK } },
  ], { x: 0.9, y: 6.28, w: 11.55, h: 0.55, margin: 0, fontFace: F, fontSize: 15 });
  s.addNotes(
    "SCRIPT: Before the how, the why. Caterpillar grades us on two things that both come down to data quality. Trifecta is their combined score across customer, contact, and equipment: and it is strict: a customer only counts when every leg passes, including one hundred percent asset match on priority assets. Our asset match sits around ninety-three percent, it is the leg our governance team is actively focused on, and today it is maintained by hand. DNA is the alerting side: Caterpillar flags our data issues and grades how well we resolve them. Together that is a quarter of the dealer excellence scorecard riding on whether our ownership records match reality. " +
    "CUE: land on the banner: both metrics stand on CCAT ownership records: which is exactly what this project manages."
  );
}

// --------------------------------------- 1c · how this project moves them
{
  const s = base("This project is the asset-match engine", "why it matters");
  const rows = [
    { c: BLUE, t: "Machines missing from CCAT", body: "get added, audited — they enter Caterpillar's population on the right customer" },
    { c: ORANGE, t: "Machines on the wrong customer", body: "get reassigned — the exact defect that breaks a customer's asset match" },
    { c: RED, t: "Machines other dealers claim", body: "get reviewed transfers — closing the gaps no amount of data entry can fix" },
    { c: TEAL, t: "Manual upkeep becomes a pipeline", body: "nightly detection, audited fixes, and alerts prevented at the source instead of resolved after they fire" },
  ];
  rows.forEach((r, i) => {
    const y = 1.75 + i * 1.05;
    panel(s, 0.6, y, 12.1, 0.9);
    dot(s, 0.9, y + 0.36, r.c, 0.18);
    s.addText(r.t, { x: 1.25, y: y + 0.14, w: 4.35, h: 0.65, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: r.c, valign: "middle" });
    s.addText(r.body, { x: 5.7, y: y + 0.14, w: 6.75, h: 0.65, margin: 0, fontFace: F, fontSize: 13.5, color: INK, valign: "middle" });
  });
  panel(s, 0.6, 6.05, 12.1, 0.9, PANEL2);
  s.addText([
    { text: "One wrong machine fails its whole customer for Trifecta.  ", options: { bold: true, color: YEL } },
    { text: "Every machine this system fixes can flip an entire customer back to valid — the same work, scored twice: Trifecta up, alert queue down.", options: { color: INK } },
  ], { x: 0.9, y: 6.2, w: 11.55, h: 0.65, margin: 0, fontFace: F, fontSize: 14.5, lineSpacing: 19 });
  s.addNotes(
    "SCRIPT: So how does the asset management work move those needles? Line by line: machines Caterpillar has no record of get added, correctly, with an audit trail: they enter the population on the right customer. Machines sitting on the wrong customer get reassigned: and that is precisely the defect that breaks asset match. Machines other dealers still claim get resolved through reviewed transfers: gaps that no amount of manual data entry could fix from our side. And the biggest shift: what is manual upkeep today becomes a pipeline: detection every night, fixes that audit themselves, and alert causes removed before Caterpillar ever flags them. " +
    "CUE: the leverage point on the banner: Trifecta counts customers, not machines. One wrong machine fails the whole customer, so one fixed machine can flip an entire customer back to valid. That is why per-machine work moves a customer-level score. " +
    "TRANSITION: with the why on the table, here is the system."
  );
}

// ------------------------------------------------------- 2 · the whole idea
{
  const s = base("The whole system on one slide", "the big picture");
  s.addNotes(
    "SCRIPT: Three parts. DETECT: every night, a read-only process checks machines against Caterpillar's live system and classifies each one. It cannot write anything, so it is zero risk. ACT: when something needs fixing, it goes through purpose-built Snowflake procedures that preview every change before sending it and refuse anything unsafe on their own. PROVE: two permanent tables record everything: what we did, and what the data looked like at every point in time. " +
    "CUE: Land on the bottom banner: the same sweep that finds problems verifies our fixes. The system never grades its own homework."
  );
  const cards = [
    { t: "DETECT", c: BLUE, body: "A nightly, read-only sweep checks each machine against Caterpillar's live API and classifies it. It never writes anything to Caterpillar." },
    { t: "ACT", c: ORANGE, body: "Audited Snowflake procedures do the fixes: adds, customer reassignments, and (reviewed) transfers. Every write previews first and refuses unsafe cases on its own." },
    { t: "PROVE", c: YEL, body: "Two permanent tables: AUDIT records everything we did; HISTORY records every version of the data anyone ever produced. Nothing is ever edited or deleted." },
  ];
  cards.forEach((card, i) => {
    const x = 0.6 + i * 4.15;
    panel(s, x, 1.75, 3.85, 3.4);
    s.addText(card.t, { x: x + 0.3, y: 2.05, w: 3.2, h: 0.4, margin: 0, fontFace: F, fontSize: 21, bold: true, color: card.c });
    s.addText(card.body, { x: x + 0.3, y: 2.6, w: 3.25, h: 2.4, margin: 0, fontFace: F, fontSize: 14.5, color: INK, lineSpacing: 21 });
  });
  panel(s, 0.6, 5.5, 12.1, 1.15, PANEL2);
  s.addText([
    { text: "The loop closes itself:  ", options: { bold: true, color: YEL } },
    { text: "every action's result is verified by the same sweep that found the problem — the system never marks its own homework.", options: { color: INK } },
  ], { x: 0.9, y: 5.72, w: 11.6, h: 0.75, margin: 0, fontFace: F, fontSize: 15.5 });
}

// ------------------------------------------------------------- 3 · detection
{
  const s = base("The nightly sweep: CCAT_DETECT_DISCREPANCIES", "detect");
  s.addNotes(
    "SCRIPT: The sweep asks Caterpillar one question per machine: what do you have for this serial? Then it compares the answer to our records field by field. The clever part is what it does NOT do: it remembers every machine it has verified and skips anything that has not changed, so we are not hammering Caterpillar's API with repeat questions. A machine only gets re-checked when there is a reason. " +
    "NUMBERS: under a second per machine; five thousand a night is about seventy-five minutes. Full active fleet coverage builds over a few weeks and then maintains itself. " +
    "IF ASKED: it runs on a warehouse we already pay for; the marginal cost is effectively zero."
  );
  panel(s, 0.6, 1.7, 7.6, 4.9);
  s.addText([
    { text: "One live API search per machine", options: { bold: true, color: BLUE, breakLine: true } },
    { text: "Compares Caterpillar's answer against our golden equipment view, field by field.", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Smart about what it re-checks", options: { bold: true, color: BLUE, breakLine: true } },
    { text: "A machine is only re-checked if it was never checked, our side changed, the last check errored, or the check has gone stale. Everything else is skipped for free.", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Writes everything downstream", options: { bold: true, color: BLUE, breakLine: true } },
    { text: "Fills the work queues, updates each machine's state, and snapshots any changed CCAT data into history. Returns a summary of every run.", options: { color: INK, breakLine: true } },
  ], { x: 0.95, y: 2.0, w: 6.9, h: 4.4, margin: 0, fontFace: F, fontSize: 15, lineSpacing: 20 });
  panel(s, 8.5, 1.7, 4.2, 2.3, PANEL2);
  s.addText("~0.9s", { x: 8.8, y: 1.95, w: 3.6, h: 0.8, margin: 0, fontFace: F, fontSize: 44, bold: true, color: YEL });
  s.addText("per machine checked — one polite, paced API call", { x: 8.8, y: 2.8, w: 3.6, h: 0.9, margin: 0, fontFace: F, fontSize: 14, color: DIM });
  panel(s, 8.5, 4.3, 4.2, 2.3, PANEL2);
  s.addText("5,000", { x: 8.8, y: 4.55, w: 3.6, h: 0.8, margin: 0, fontFace: F, fontSize: 44, bold: true, color: YEL });
  s.addText("machines a night in about 75 minutes — full fleet coverage builds in weeks, then maintains itself", { x: 8.8, y: 5.4, w: 3.6, h: 1.1, margin: 0, fontFace: F, fontSize: 14, color: DIM });
}

// ------------------------------------------------------ 4 · tracking tables
{
  const s = base("What the sweep writes: queues, state, and memory", "detect");
  s.addNotes(
    "SCRIPT: Think of these as the to-do lists the sweep maintains. MISSING is the pipeline for adds and transfers: machines Caterpillar has nothing for under our dealer code. NON_MATCHING is machines we hold where a detail differs. ERRORS is deliberately separate: if a lookup fails, we record the failure, never a false conclusion. A failed check is never mistaken for a missing machine. And CHECK_STATE is the memory: one row per machine, its verdict, and when we last looked. That memory is what keeps the nightly sweep cheap and what powers the dashboard on the next slide."
  );
  const cells = [
    { t: "CCAT_MISSING", c: BLUE, body: "Machines with no record under our dealer code — the add and transfer pipeline. Keeps the other dealers' records for triage." },
    { t: "CCAT_NON_MATCHING", c: ORANGE, body: "Field-level differences on machines we hold — one row per differing field, judged against our best-matching CCAT record." },
    { t: "CCAT_ERRORS", c: RED, body: "Failed checks. A failure is never recorded as a discrepancy — it lands here with Caterpillar's tracking ID for support." },
    { t: "CCAT_CHECK_STATE", c: TEAL, body: "One row per machine ever checked: verdict, ownership type, customer, pending flags — the sweep's memory and the dashboard's engine." },
  ];
  cells.forEach((cell, i) => {
    const x = 0.6 + (i % 2) * 6.2;
    const y = 1.75 + Math.floor(i / 2) * 2.55;
    panel(s, x, y, 5.9, 2.3);
    dot(s, x + 0.3, y + 0.34, cell.c, 0.18);
    s.addText(cell.t, { x: x + 0.62, y: y + 0.22, w: 5.0, h: 0.4, margin: 0, fontFace: F, fontSize: 17, bold: true, color: cell.c });
    s.addText(cell.body, { x: x + 0.3, y: y + 0.75, w: 5.3, h: 1.45, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18 });
  });
}

// ------------------------------------------------------------- 5 · phases
{
  const s = base("Every machine gets a phase, an action, and a why", "classify");
  s.addNotes(
    "SCRIPT: The colors are ordered by blast radius: who else can see what we do. Green through orange, nobody outside the company is affected, which is why those can be automated with volume caps. Red is the only lane other dealers can see, and that is exactly the lane where a person reviews every single item. Gray is worth a word: some model names will always differ because Caterpillar's version is the official one, so we report those and deliberately do nothing. That is why the goal is the right customer on every machine, not a cosmetic one hundred percent match. " +
    "CUE: The dashboard query on the right is live in Snowflake, filterable by phase, exportable to Excel."
  );
  const rows = [
    { c: GREEN, t: "Done", body: "Record is right. Nothing to do." },
    { c: BLUE, t: "Phase 1 — Add", body: "Not in CCAT anywhere. Adding affects nobody. Batchable." },
    { c: TEAL, t: "Phase 2 — Quiet fix", body: "Other records exist but none are exclusive. Completes without notifying anyone." },
    { c: ORANGE, t: "Phase 3 — Fix ours", body: "We own the record: backfill missing values, or move it to the correct customer." },
    { c: RED, t: "Phase 4 — Review", body: "Transfers and shared ownership. A person decides, about ten per shift." },
    { c: GRAY, t: "Info only", body: "Model differs but CCAT is the authority — reported, no action queued." },
  ];
  rows.forEach((r, i) => {
    const y = 1.72 + i * 0.82;
    panel(s, 0.6, y, 8.0, 0.7);
    dot(s, 0.88, y + 0.26, r.c, 0.18);
    s.addText(r.t, { x: 1.2, y: y + 0.1, w: 2.6, h: 0.5, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: r.c, valign: "middle" });
    s.addText(r.body, { x: 3.85, y: y + 0.1, w: 4.6, h: 0.5, margin: 0, fontFace: F, fontSize: 12.5, color: INK, valign: "middle" });
  });
  panel(s, 8.9, 1.72, 3.8, 4.82, PANEL2);
  s.addText("One query answers it all", { x: 9.2, y: 2.0, w: 3.2, h: 0.6, margin: 0, fontFace: F, fontSize: 17, bold: true, color: YEL });
  s.addText("The phase dashboard shows, per machine:\n\nPHASE  ·  ACTION  ·  WHY\n\nplus both customer numbers, ownership type, and which other dealers are involved. The rollout plan, applied to live data.", {
    x: 9.2, y: 2.6, w: 3.25, h: 3.6, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 19,
  });
}

// ---------------------------------------------------- 6 · the animated key
{
  const s = base("The decision key, animated", "classify");
  s.addNotes(
    "SCRIPT: Rather than walk you through the rules, let three real machines show you. INTRO LINE: This is eighty-three seconds; the narrator does the work. " +
    "CUE: Click the video, let it play fully, do not talk over it. It covers: a new excavator landing in the safe add bucket, a dozer with an old rental record elsewhere landing in the quiet add bucket, and a machine another dealer actively owns landing in reviewed transfer. " +
    "AFTER: The serials in that video are real machines from our fleet; the excavator was actually added to CCAT through this system last week."
  );
  const poster = fs.readFileSync("video_poster.png").toString("base64");
  s.addMedia({
    type: "video",
    path: "../manim/final.mp4",
    cover: "image/png;base64," + poster,
    x: 0.6, y: 1.75, w: 8.4, h: 4.725,
  });
  panel(s, 9.3, 1.75, 3.4, 4.725, PANEL2);
  s.addText("83 seconds, narrated", { x: 9.6, y: 2.05, w: 2.85, h: 0.6, margin: 0, fontFace: F, fontSize: 16, bold: true, color: YEL });
  s.addText("Three real machines walk the decision tree:\n\na safe add,\na quiet add,\nand a reviewed transfer.\n\nClick to play.", {
    x: 9.6, y: 2.65, w: 2.85, h: 3.4, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 19,
  });
}

// ------------------------------------------------------------ 7 · wrappers
{
  const s = base("Acting: every write goes through an audited procedure", "act");
  s.addNotes(
    "SCRIPT: When we act, there is exactly one sanctioned door: purpose-built procedures, one per action type. You give them a serial number; they pull everything else themselves, search Caterpillar live, and refuse anything outside their lane: the record already exists, another dealer owns it, it is an attachment with a made-up serial, or Caterpillar already rejected that serial once. By default they are in preview mode: they show exactly what would be sent and send nothing. Executing is a deliberate second step. " +
    "CUE payload rule: this came from our own data review with Vinod and Michele: Caterpillar's model names are the ones their systems recognize, so we never overwrite them with ours. " +
    "CUE raw tools: the low-level API tools still exist for diagnostics, but nobody will ever be granted them: only the audited wrappers."
  );
  panel(s, 0.6, 1.7, 6.0, 4.9);
  s.addText("CCAT_ADD_ASSET  ·  CCAT_ADD_INVENTORY", { x: 0.9, y: 1.95, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: BLUE });
  s.addText("CCAT_REASSIGN_CUSTOMER", { x: 0.9, y: 2.35, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: ORANGE });
  s.addText([
    { text: "Look up the machine by serial number, pull everything else from our records", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Search CCAT live before acting — never write blind", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Refuse anything outside their lane: existing records, other owners, attachments, known-bad serials", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Dry-run by default — show exactly what would be sent, send nothing", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "On execute: one API call, one audit row, one history snapshot", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.9, y: 2.95, w: 5.4, h: 3.5, margin: 0, fontFace: F, fontSize: 13.5, lineSpacing: 18 });
  panel(s, 6.9, 1.7, 5.8, 2.3);
  s.addText("Payload rule", { x: 7.2, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: YEL });
  s.addText("Where CCAT already has a record, its model and year are the versions Caterpillar recognizes — our requests reuse them. Our data only goes to Caterpillar for machines it has never seen.", {
    x: 7.2, y: 2.42, w: 5.2, h: 1.5, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
  panel(s, 6.9, 4.3, 5.8, 2.3);
  s.addText("Raw API tools still exist", { x: 7.2, y: 4.55, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: RED });
  s.addText("The low-level add/expire/transfer procedures are diagnostics only. They are not audited, and when access rolls out, other users will only ever be granted the wrappers.", {
    x: 7.2, y: 5.02, w: 5.2, h: 1.4, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
}

// ---------------------------------------------- 7b · phase one flow diagram
{
  const s = base("Phase one flow: inventory adds (live)", "act · flow");
  s.addNotes(
    "SCRIPT: This is the pipeline that is running today, end to end. NAXT data lands each morning; a daily scan picks qualifying inventory machines, oldest first. Before any API call, the skip lists drop everything we already know about: prior adds, machines already in CCAT, machines another dealer blocks: those cost nothing. Each remaining machine gets one live search, and every answer is saved to history whether we act or not. The dry run previews the batch; execution adds each machine under our inventory account with a full audit row; failures are logged once and never retried blind. Then the one manual step: Lindsay's subscription. The same sweep that finds problems confirms each machine lands as DONE_INVENTORY. " +
    "CUE: week one through this pipe: 50 Monday plus 50 Wednesday, capped on purpose to protect the manual subscription step."
  );
  const node = (x, y, c, t, body) => {
    panel(s, x, y, 2.75, 1.55);
    s.addText(t, { x: x + 0.14, y: y + 0.1, w: 2.47, h: 0.5, margin: 0, fontFace: F, fontSize: 12.5, bold: true, color: c, lineSpacing: 14 });
    s.addText(body, { x: x + 0.14, y: y + 0.6, w: 2.47, h: 0.9, margin: 0, fontFace: F, fontSize: 10, color: INK, lineSpacing: 13 });
  };
  const arrowR = (x, y) => s.addShape("rightArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });
  const arrowL = (x, y) => s.addShape("leftArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });

  // row 1 (left to right)
  node(0.6,  1.75, BLUE, "1 · NAXT refresh", "Golden equipment data lands in Snowflake each morning");
  node(3.75, 1.75, BLUE, "2 · Candidate scan", "Blank customer · Cat make · active · New · no attachments — oldest first by modified date");
  node(6.9,  1.75, TEAL, "3 · Free skips", "Already added, already in CCAT, or dealer-blocked — dropped from history at zero API cost");
  node(10.05,1.75, TEAL, "4 · Live CCAT search", "One API search per machine; every answer saved to history");
  arrowR(3.42, 2.4); arrowR(6.57, 2.4); arrowR(9.72, 2.4);
  s.addShape("downArrow", { x: 11.28, y: 3.42, w: 0.28, h: 0.76, fill: { color: DIM }, line: { type: "none" } });

  // row 2 (right to left)
  node(10.05, 4.3, YEL,    "5 · Dry-run review", "The would_add list previewed — nothing sent yet");
  node(6.9,   4.3, GREEN,  "6 · Add to INT00495", "Executed with audit row + Cat tracking ID; failures logged once, auto-skipped after");
  node(3.75,  4.3, ORANGE, "7 · Subscription", "Lindsay registers Product Link in the dealer portal — the manual step that sets the pace");
  node(0.6,   4.3, GREEN,  "8 · Sweep verifies", "Machine lands DONE_INVENTORY on the dashboard — loop closed");
  arrowL(9.72, 4.95); arrowL(6.57, 4.95); arrowL(3.42, 4.95);

  panel(s, 0.6, 6.15, 12.1, 0.9, PANEL2);
  s.addText([
    { text: "Live since August 10.  ", options: { bold: true, color: GREEN } },
    { text: "100 machines through this pipe in week one (50 Monday + 50 Wednesday) — capped on purpose: step 7 is a person.", options: { color: INK } },
  ], { x: 0.9, y: 6.32, w: 11.55, h: 0.6, margin: 0, fontFace: F, fontSize: 14 });
}

// ---------------------------------------------- 7c · phase two flow diagram
{
  const s = base("Phase two flow: sold machine → its buyer (designed)", "act · flow");
  s.addNotes(
    "SCRIPT: Phase two is scoped and the detection half already runs: the moment a customer number appears on an inventory machine in NAXT, the dashboard flags it: forty-one machines are sitting on that list right now, each one currently a manual move for Lindsay. The new piece since the August 6 meeting: Caterpillar's customer master is shared directly into our Snowflake, so the CCID question that blocked automation: does this buyer exist properly on Cat's side: is now answerable BEFORE we act. Buyer resolves: automated, audited move; the old inventory record expires itself. Buyer missing: it routes to the Customer Admin Tool queue instead of failing mid-flight. " +
    "CUE: until this is built and approved, the manual path stays: PDI email, Lindsay moves it within about a day. The flow replaces exactly that."
  );
  const node = (x, y, w, h, c, t, body) => {
    panel(s, x, y, w, h);
    s.addText(t, { x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.55, margin: 0, fontFace: F, fontSize: 13, bold: true, color: c, lineSpacing: 15 });
    s.addText(body, { x: x + 0.16, y: y + 0.68, w: w - 0.32, h: h - 0.8, margin: 0, fontFace: F, fontSize: 10.5, color: INK, lineSpacing: 14 });
  };
  const arrowR = (x, y) => s.addShape("rightArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });

  node(0.6,  2.6, 2.7, 1.7, BLUE, "1 · Machine sells", "A customer number appears on an inventory machine in NAXT");
  node(3.75, 2.6, 2.7, 1.7, YEL,  "2 · Dashboard flags it", "P3_INV_TO_CUSTOMER — caught by the daily sweep. 41 machines on this list today");
  node(6.9,  2.6, 2.7, 1.7, TEAL, "3 · CCID check", "Cat's customer master, shared into Snowflake, answers: does the buyer resolve to a CCID?");
  arrowR(3.42, 3.32); arrowR(6.57, 3.32);

  s.addShape("line", { x: 9.68, y: 2.72, w: 0.5, h: 0.62, flipV: true, line: { color: GREEN, width: 2.25, endArrowType: "triangle" } });
  s.addShape("line", { x: 9.68, y: 3.55, w: 0.5, h: 1.5, line: { color: ORANGE, width: 2.25, endArrowType: "triangle" } });
  s.addText("CCID exists", { x: 9.42, y: 2.28, w: 1.3, h: 0.3, margin: 0, fontFace: F, fontSize: 10, bold: true, color: GREEN });
  s.addText("no CCID", { x: 9.5, y: 4.45, w: 1.1, h: 0.3, margin: 0, fontFace: F, fontSize: 10, bold: true, color: ORANGE });

  node(10.25, 1.8, 2.5, 2.05, GREEN, "4a · Automated move", "Reassign INT00495 → the buyer's account. Old record auto-expires. One audit row, one API call.");
  node(10.25, 4.35, 2.5, 2.05, ORANGE, "4b · Human queue", "Customer Admin Tool creates the CCID first — then the move retries. No mid-flight failures.");

  panel(s, 0.6, 6.55, 9.3, 0.75, PANEL2);
  s.addText([
    { text: "Until built & approved:  ", options: { bold: true, color: YEL } },
    { text: "PDI email → Lindsay moves the machine by hand (~24 h). This flow replaces exactly that step.", options: { color: INK } },
  ], { x: 0.85, y: 6.68, w: 8.9, h: 0.5, margin: 0, fontFace: F, fontSize: 12.5 });
}

// --------------------------------------------------------------- 8 · audit
{
  const s = base("CCAT_AUDIT: what we did, forever", "prove");
  s.addNotes(
    "SCRIPT: Every action, executed or failed, writes one audit row: what we did, to which machine, the customer before and after, the exact payload we sent, Caterpillar's full response, and their tracking ID, which means Cat support can trace any call we ever made. It also stores a snapshot of what Caterpillar had immediately before we acted, so if anything ever needs unwinding, we know exactly what to restore. " +
    "CUE right panel: append-only is not a policy we promise to follow; it is enforced by database permissions. The tooling can add rows and read rows. It physically cannot edit or delete them. " +
    "IF ASKED who and when: every row carries user, timestamp, and shift date: that is where per-shift reporting comes from."
  );
  panel(s, 0.6, 1.7, 7.7, 4.9);
  s.addText("Every executed or failed action writes one row:", { x: 0.95, y: 2.0, w: 6.9, h: 0.4, margin: 0, fontFace: F, fontSize: 15, bold: true, color: INK });
  s.addText([
    { text: "What: action type, phase, machine, model, ownership type", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Customer before  →  customer after (the reassignment story)", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "The exact payload sent, and Caterpillar's full response", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Caterpillar's tracking ID — their support can trace any call", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "A snapshot of CCAT taken just before we acted (the undo reference)", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Who, when, and which shift — feeds the per-shift reporting", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.95, y: 2.5, w: 7.0, h: 3.9, margin: 0, fontFace: F, fontSize: 14, lineSpacing: 19 });
  panel(s, 8.6, 1.7, 4.1, 4.9, PANEL2);
  s.addText("Append-only\nby permission", { x: 8.9, y: 2.1, w: 3.5, h: 1.1, margin: 0, fontFace: F, fontSize: 22, bold: true, color: YEL });
  s.addText("The tooling role can INSERT and SELECT — nothing else. Audit rows cannot be edited or deleted, even by accident.\n\nKept forever. This table is the compliance trail.", {
    x: 8.9, y: 3.35, w: 3.5, h: 3.0, margin: 0, fontFace: F, fontSize: 14, color: INK, lineSpacing: 19,
  });
}

// -------------------------------------------------------------- 9 · history
{
  const s = base("CCAT_HISTORY: what the data was, at every point", "prove");
  s.addNotes(
    "SCRIPT: History answers a different question: not what did we do, but what did the data look like, at every point, no matter who changed it. Every time we look at a machine, we fingerprint Caterpillar's full answer. If the fingerprint matches the last one, we store nothing. If it changed, we store the complete snapshot. So a machine checked fifty times with no change costs one row: history only grows when reality changes. " +
    "CUE: This costs zero extra API calls: it is built entirely from answers the sweep was already fetching and throwing away. " +
    "NICE DETAIL if time: even no records at all counts as a version, so a machine appearing in or vanishing from CCAT leaves a trail."
  );
  panel(s, 0.6, 1.7, 7.7, 4.9);
  s.addText([
    { text: "Every distinct version, any author", options: { bold: true, color: TEAL, breakLine: true } },
    { text: "Each check fingerprints Caterpillar's full answer for the machine. New fingerprint = the full snapshot is stored. Same fingerprint = nothing.", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Cheap by design", options: { bold: true, color: TEAL, breakLine: true } },
    { text: "A machine checked fifty times with no change stores one row, not fifty. Built entirely from data the sweep already fetches — zero extra API calls.", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Absence counts too", options: { bold: true, color: TEAL, breakLine: true } },
    { text: "“No records at all” is itself a version — a machine appearing in, or vanishing from, CCAT leaves a trail.", options: { color: INK, breakLine: true } },
  ], { x: 0.95, y: 2.0, w: 7.0, h: 4.4, margin: 0, fontFace: F, fontSize: 14.5, lineSpacing: 19 });
  panel(s, 8.6, 1.7, 4.1, 4.9, PANEL2);
  s.addText("1 row", { x: 8.9, y: 2.1, w: 3.5, h: 0.9, margin: 0, fontFace: F, fontSize: 48, bold: true, color: YEL });
  s.addText("per real change in Caterpillar's data, no matter how many times we look.\n\nAppend-only and kept forever, same as audit.", {
    x: 8.9, y: 3.1, w: 3.5, h: 2.6, margin: 0, fontFace: F, fontSize: 14, color: INK, lineSpacing: 19,
  });
}

// -------------------------------------------- 10 · audit vs history payoff
{
  const s = base("Two tables, one powerful difference", "prove");
  s.addNotes(
    "SCRIPT: Here is why having both matters. Audit is what we did. History is what the data was. Subtract one from the other and you get something we have never had before: external changes. If history shows a machine's record changed and there is no audit row from us around that time, then Caterpillar's systems, another dealer, or someone in a portal changed it: and now we see it instead of discovering it months later. " +
    "CUE: pause on the banner: this is the class of change that used to be completely invisible to us. " +
    "IF ASKED for an example: another dealer filing a transfer request against one of our machines shows up this way within a day."
  );
  panel(s, 0.6, 1.75, 5.9, 3.3);
  s.addText("AUDIT", { x: 0.9, y: 2.05, w: 5.2, h: 0.45, margin: 0, fontFace: F, fontSize: 20, bold: true, color: YEL });
  s.addText("What WE did.\nActions, payloads, outcomes, tracking IDs — signed, dated, attributable.", {
    x: 0.9, y: 2.6, w: 5.3, h: 2.2, margin: 0, fontFace: F, fontSize: 15, color: INK, lineSpacing: 21,
  });
  panel(s, 6.8, 1.75, 5.9, 3.3);
  s.addText("HISTORY", { x: 7.1, y: 2.05, w: 5.2, h: 0.45, margin: 0, fontFace: F, fontSize: 20, bold: true, color: TEAL });
  s.addText("What the data WAS.\nEvery observed version of every machine's CCAT picture, whoever changed it.", {
    x: 7.1, y: 2.6, w: 5.3, h: 2.2, margin: 0, fontFace: F, fontSize: 15, color: INK, lineSpacing: 21,
  });
  panel(s, 0.6, 5.35, 12.1, 1.3, PANEL2);
  s.addText([
    { text: "History minus audit = external changes.  ", options: { bold: true, color: YEL } },
    { text: "A version change with no audit row nearby means Caterpillar, another dealer, or a portal edit changed the data — the class of change that used to be invisible.", options: { color: INK } },
  ], { x: 0.9, y: 5.55, w: 11.55, h: 0.95, margin: 0, fontFace: F, fontSize: 15.5, lineSpacing: 21 });
}

// ------------------------------------------------------ 11 · rhythm+metrics
{
  const s = base("The operating rhythm", "run it");
  s.addNotes(
    "SCRIPT: Day to day it looks like this. Nightly, the sweep runs by itself: read-only, no supervision needed. Weekly, the safe piles get worked in capped batches: adds and quiet fixes, the work nobody else can see. And every shift, about ten reviewed decisions: the transfers and judgment calls: each one audited as it happens. " +
    "CUE metrics: the headline number is deliberately not the match rate: it is the customer assignment rate, the percentage of machines sitting on the correct customer in Caterpillar's system: because that is the outcome the business asked for. Everything else is supporting detail. " +
    "ASK if appropriate: the one approval this slide needs is scheduling the nightly sweep as an automated task."
  );
  const cols = [
    { t: "Nightly", c: BLUE, body: "The sweep checks its batch, refreshes verdicts, and grows coverage. Read-only, automatic." },
    { t: "Weekly", c: TEAL, body: "Batched safe adds and quiet fixes run under volume caps. A scoreboard snapshot tracks the backlog shrinking." },
    { t: "Every shift", c: RED, body: "About ten reviewed decisions: transfers, inbound claims, judgment calls — audited as they happen." },
  ];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    panel(s, x, 1.75, 3.85, 2.5);
    s.addText(c.t, { x: x + 0.3, y: 2.0, w: 3.2, h: 0.4, margin: 0, fontFace: F, fontSize: 18, bold: true, color: c.c });
    s.addText(c.body, { x: x + 0.3, y: 2.5, w: 3.25, h: 1.6, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18 });
  });
  panel(s, 0.6, 4.55, 12.1, 2.1, PANEL2);
  s.addText("What leadership sees, weekly:", { x: 0.9, y: 4.8, w: 11, h: 0.4, margin: 0, fontFace: F, fontSize: 15, bold: true, color: YEL });
  s.addText([
    { text: "Customer assignment rate — the headline: % of machines on the correct customer in CCAT", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 5 } },
    { text: "Coverage and match rate — how much of the fleet is verified, and how much agrees", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 5 } },
    { text: "Backlog by phase, resolutions per shift, and transfers in flight", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.9, y: 5.3, w: 11.4, h: 1.3, margin: 0, fontFace: F, fontSize: 13.5, lineSpacing: 18 });
}

// ----------------------------------------------------------- 12 · the rules
{
  const s = base("The rules that keep it safe", "guardrails");
  s.addNotes(
    "SCRIPT: I will not read all six: three matter most. No bulk claims, ever: a blanket push would fire formal transfer requests at real dealers across the network and can silently overwrite our own records: the system is built so that shortcut is impossible, not just discouraged. Shared customers are excluded automatically: when another dealer serves the same national account, the machine routes to a conversation, never a claim. And production awareness: Caterpillar has no test environment, so every write is treated as real, previewed first, and logged. " +
    "CUE: these guardrails travel with the tools themselves: they apply no matter who eventually gets access."
  );
  const rules = [
    { t: "Dry-run first, always", body: "Every write shows its exact payload and sends nothing until deliberately executed.", c: GREEN },
    { t: "No bulk claims, ever", body: "Blanket adds would fire transfer requests at real dealers and can silently overwrite our own records.", c: RED },
    { t: "Shared customers excluded", body: "When another dealer serves the same customer, the machine routes to coordination — never a claim.", c: TEAL },
    { t: "CCAT owns the model", body: "We never push our model names over Caterpillar's — their versions are what Cat systems recognize.", c: YEL },
    { t: "Production awareness", body: "Caterpillar has no test environment. Every write is treated as real, because it is.", c: ORANGE },
    { t: "Compliance tables are sacred", body: "Audit and history are append-only by permission and kept forever — off the reset list permanently.", c: BLUE },
  ];
  rules.forEach((r, i) => {
    const x = 0.6 + (i % 3) * 4.15;
    const y = 1.75 + Math.floor(i / 3) * 2.5;
    panel(s, x, y, 3.85, 2.25);
    dot(s, x + 0.3, y + 0.32, r.c, 0.16);
    s.addText(r.t, { x: x + 0.58, y: y + 0.2, w: 3.1, h: 0.65, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: r.c });
    s.addText(r.body, { x: x + 0.3, y: y + 0.85, w: 3.3, h: 1.3, margin: 0, fontFace: F, fontSize: 12, color: INK, lineSpacing: 16 });
  });
}

// ----------------------------------------------------- 13 · built vs next
{
  const s = base("Where it stands", "status");
  s.addNotes(
    "SCRIPT: Being straight about status. Everything on the left is not a plan: it is running. The sweep has checked thousands of machines. Real equipment has been added to Caterpillar's production system through the audited procedures, with the audit rows to show for it. The right column is scheduling and scale, not invention: the nightly task is a small script; the batch runner and the remaining wrappers reuse patterns already proven; access roles are written and waiting on rollout approval. " +
    "CUE: the honest framing is: the engineering risk is behind us: what remains is operational rollout."
  );
  panel(s, 0.6, 1.75, 5.9, 4.9);
  s.addText("Working today", { x: 0.9, y: 2.05, w: 5.2, h: 0.45, margin: 0, fontFace: F, fontSize: 19, bold: true, color: GREEN });
  s.addText([
    { text: "Detection sweep, proven at thousand-machine scale", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Phase dashboard: phase, action, and why per machine", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Audited procedures for adds, inventory adds, and customer reassignment — real adds executed in production", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Audit + history compliance layer", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Full documentation and a versioned SQL reference", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.9, y: 2.65, w: 5.3, h: 3.8, margin: 0, fontFace: F, fontSize: 14, lineSpacing: 19 });
  panel(s, 6.8, 1.75, 5.9, 4.9);
  s.addText("Next", { x: 7.1, y: 2.05, w: 5.2, h: 0.45, margin: 0, fontFace: F, fontSize: 19, bold: true, color: YEL });
  s.addText([
    { text: "Schedule the sweep as a nightly Snowflake task", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Batch runner for the safe-add pile, with volume caps", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Backfill and transfer wrappers (same audited pattern)", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Reader/writer access roles once rollout is approved", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Review workflow at ~10 resolutions per shift", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 7.1, y: 2.65, w: 5.3, h: 3.8, margin: 0, fontFace: F, fontSize: 14, lineSpacing: 19 });
}

// ------------------------------------------------------------- 14 · closing
{
  const s = pres.addSlide();
  s.background = { color: BG };
  s.addText("The system sorts every machine.", {
    x: 0.9, y: 2.5, w: 11.5, h: 0.8, margin: 0, fontFace: F, fontSize: 40, bold: true, color: INK,
  });
  s.addText("People decide only the ones that need judgment.", {
    x: 0.9, y: 3.45, w: 11.5, h: 0.7, margin: 0, fontFace: F, fontSize: 28, color: YEL,
  });
  s.addText("Code: SQL Reference folder (files 40–49)   ·   Docs: 10–14   ·   Dashboard: V_CCAT_PHASE_DASHBOARD", {
    x: 0.9, y: 6.5, w: 11.5, h: 0.4, margin: 0, fontFace: F, fontSize: 13, color: DIM,
  });
  s.addNotes(
    "SCRIPT: One sentence to remember: the system sorts every machine; people decide only the ones that need judgment. That is the whole design: automation where it is safe, human judgment where it matters, and a permanent record of both. Happy to take questions: and if you want the detail behind any slide, the rollout plan document and the SQL reference carry all of it. " +
    "CLOSE: the decisions I need from this room are the four in the rollout plan: approve the phased order, the nightly job, the pilot batches, and the review staffing."
  );
}

pres.writeFile({ fileName: "Snowflake_CCAT_Reconciliation.pptx" }).then(() => console.log("WROTE Snowflake_CCAT_Reconciliation.pptx"));
