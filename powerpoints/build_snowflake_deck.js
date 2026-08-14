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
    "SCRIPT: Good morning. This is about a problem every dealer has. Our equipment records and Cat's official ownership registry drift apart over time. Machines Cat has no record of, machines sitting on the wrong customer, machines another dealer still claims. What I want to show you is the system we built inside Snowflake, which we're already paying for, that finds all of those differences, fixes the safe ones on its own, sends the risky ones to a person, and keeps a permanent record of every bit of it. " +
    "CUE: the six dots are the six things that can happen to a machine. They come back on slide five."
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
  s.addText("asset match conformance (2025 baseline). The leg our data governance team is focused on, and the one still maintained by hand today", {
    x: 3.3, y: 3.72, w: 2.95, h: 1.3, margin: 0, fontFace: F, fontSize: 12.5, color: DIM, lineSpacing: 16,
  });
  s.addText("Worth 15% of the dealer excellence scorecard", { x: 0.9, y: 5.25, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 13.5, bold: true, color: INK });
  panel(s, 6.8, 1.7, 5.9, 4.15);
  s.addText("DNA · DATA NOTIFICATION ALERTS", { x: 7.1, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 18, bold: true, color: TEAL });
  s.addText("Caterpillar flags our prioritized data issues, asset alerts and customer alerts, sends us the top ones by value, and grades how well we resolve them.", {
    x: 7.1, y: 2.42, w: 5.3, h: 1.15, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
  s.addText("10%", { x: 7.1, y: 3.7, w: 2.0, h: 0.7, margin: 0, fontFace: F, fontSize: 40, bold: true, color: TEAL });
  s.addText("of the scorecard rides on DNA resolution, and many of the asset alerts trace back to the same root cause: ownership records that do not match reality", {
    x: 9.2, y: 3.72, w: 3.05, h: 1.3, margin: 0, fontFace: F, fontSize: 12.5, color: DIM, lineSpacing: 16,
  });
  panel(s, 0.6, 6.1, 12.1, 0.85, PANEL2);
  s.addText([
    { text: "Both metrics stand on the same ground:  ", options: { bold: true, color: YEL } },
    { text: "asset ownership records in CCAT. That is exactly what this project manages.", options: { color: INK } },
  ], { x: 0.9, y: 6.28, w: 11.55, h: 0.55, margin: 0, fontFace: F, fontSize: 15 });
  s.addNotes(
    "SCRIPT: Before the how, the why. Cat grades us on two things that both come down to data quality. Trifecta is their combined score across customer, contact, and equipment, and it's strict. A customer only counts when every leg passes, including a hundred percent asset match on their priority assets. We're around ninety-three percent on asset match. That's the leg our governance team is focused on, and today it's all done by hand. DNA is the alert side: Cat flags our data problems and grades how well we clear them. Put those together and about a quarter of the dealer excellence scorecard rides on whether our ownership records match reality. " +
    "CUE: land on the banner. Both of those scores sit on CCAT ownership records, and that's exactly what this project manages."
  );
}

// --------------------------------------- 1c · how this project moves them
{
  const s = base("This project is the asset-match engine", "why it matters");
  const rows = [
    { c: BLUE, t: "Machines missing from CCAT", body: "get added, audited, so they enter Caterpillar's population on the right customer" },
    { c: ORANGE, t: "Machines on the wrong customer", body: "get reassigned, the exact defect that breaks a customer's asset match" },
    { c: RED, t: "Machines other dealers claim", body: "get reviewed transfers, closing the gaps no amount of data entry can fix" },
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
    { text: "Every machine this system fixes can flip an entire customer back to valid. The same work, scored twice: Trifecta up, alert queue down.", options: { color: INK } },
  ], { x: 0.9, y: 6.2, w: 11.55, h: 0.65, margin: 0, fontFace: F, fontSize: 14.5, lineSpacing: 19 });
  s.addNotes(
    "SCRIPT: So how does this actually move those numbers? Line by line. Machines Cat has no record of get added, correctly, with an audit trail behind them, so now they're in the population on the right customer. Machines on the wrong customer get moved, and that's the exact defect that breaks asset match. Machines another dealer still claims go through a reviewed transfer, which is a gap we could never close from our side with data entry. And the biggest change is the last one: what's manual upkeep today turns into a pipeline. Detection every night, fixes that audit themselves, and alert causes gone before Cat ever flags them. " +
    "CUE: the leverage is on the banner. Trifecta counts customers, not machines. One bad machine fails the whole customer, so fixing one machine can flip that customer back to valid. That's why machine-level work moves a customer-level score. " +
    "TRANSITION: that's the why. Here's the system."
  );
}

// ------------------------------------------------------- 2 · the whole idea
{
  const s = base("The whole system on one slide", "the big picture");
  s.addNotes(
    "SCRIPT: Three parts. Detect. Every night a read-only process checks machines against Cat's live system and sorts each one. It can't write anything, so there's no risk in it. Act. When something needs fixing it goes through procedures we built for that one job. They show you the change before they send it, and they refuse anything outside their lane on their own. Prove. Two permanent tables that record what we did, and what the data looked like at every point. " +
    "CUE: land on the bottom banner. The same sweep that finds the problems is what confirms the fixes. It never grades its own homework."
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
    { text: "every action's result is verified by the same sweep that found the problem. The system never marks its own homework.", options: { color: INK } },
  ], { x: 0.9, y: 5.72, w: 11.6, h: 0.75, margin: 0, fontFace: F, fontSize: 15.5 });
}

// ------------------------------------------------------------- 3 · detection
{
  const s = base("The nightly sweep: CCAT_DETECT_DISCREPANCIES", "detect");
  s.addNotes(
    "SCRIPT: The sweep asks Cat one question per machine. What do you have for this serial? Then it compares the answer to our records, field by field. The useful part is what it doesn't do. It remembers every machine it's already checked and skips anything that hasn't changed, so we're not hammering Cat's API asking the same question over and over. A machine only gets looked at again when there's a reason. " +
    "NUMBERS: under a second per machine. Five thousand a night is about seventy-five minutes. Coverage of the active fleet builds over a few weeks and then keeps itself up. " +
    "IF ASKED: it runs on a warehouse we're already paying for, so the extra cost is basically nothing."
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
  s.addText("per machine checked: one polite, paced API call", { x: 8.8, y: 2.8, w: 3.6, h: 0.9, margin: 0, fontFace: F, fontSize: 14, color: DIM });
  panel(s, 8.5, 4.3, 4.2, 2.3, PANEL2);
  s.addText("5,000", { x: 8.8, y: 4.55, w: 3.6, h: 0.8, margin: 0, fontFace: F, fontSize: 44, bold: true, color: YEL });
  s.addText("machines a night in about 75 minutes. Full fleet coverage builds in weeks, then maintains itself", { x: 8.8, y: 5.4, w: 3.6, h: 1.1, margin: 0, fontFace: F, fontSize: 14, color: DIM });
}

// ------------------------------------------------------ 4 · tracking tables
{
  const s = base("What the sweep writes: queues, state, and memory", "detect");
  s.addNotes(
    "SCRIPT: Think of these as the to-do lists the sweep keeps. Missing is the queue for adds and transfers, the machines Cat has nothing for under our dealer code. Non-matching is machines we do hold, where some detail is off. Errors is separate on purpose: if a lookup fails, we write down that it failed, not a conclusion. A failed check never gets mistaken for a missing machine. And check-state is the memory. One row per machine, what we decided, and when we last looked. That memory is what keeps the nightly run cheap, and it's what drives the dashboard on the next slide."
  );
  const cells = [
    { t: "CCAT_MISSING", c: BLUE, body: "Machines with no record under our dealer code. The add and transfer pipeline. Keeps the other dealers' records for triage." },
    { t: "CCAT_NON_MATCHING", c: ORANGE, body: "Field-level differences on machines we hold. One row per differing field, judged against our best-matching CCAT record." },
    { t: "CCAT_ERRORS", c: RED, body: "Failed checks. A failure is never recorded as a discrepancy. It lands here with Caterpillar's tracking ID for support." },
    { t: "CCAT_CHECK_STATE", c: TEAL, body: "One row per machine ever checked: verdict, ownership type, customer, pending flags. The sweep's memory and the dashboard's engine." },
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
    "SCRIPT: The colors go in order of who else can see what we do. Green through orange, nobody outside this company is affected, which is why those can run automatically with caps on the volume. Red is the only lane other dealers can see, and that's exactly where a person looks at every single one. Gray deserves a word. Some model names will always disagree because Cat's version is the official one, so we report those and deliberately leave them alone. That's why the goal is the right customer on every machine, not a cosmetic hundred percent match. " +
    "CUE: the dashboard on the right is live in Snowflake. Filter it by phase, export it to Excel."
  );
  const rows = [
    { c: GREEN, t: "Done", body: "Record is right. Nothing to do." },
    { c: BLUE, t: "Phase 1 · Add", body: "Not in CCAT anywhere. Adding affects nobody. Batchable." },
    { c: TEAL, t: "Phase 2 · Quiet fix", body: "Other records exist but none are exclusive. Completes without notifying anyone." },
    { c: ORANGE, t: "Phase 3 · Fix ours", body: "We own the record: backfill missing values, or move it to the correct customer." },
    { c: RED, t: "Phase 4 · Review", body: "Transfers and shared ownership. A person decides, about ten per shift." },
    { c: GRAY, t: "Info only", body: "Model differs but CCAT is the authority. Reported, no action queued." },
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
    "SCRIPT: Instead of walking you through the rules, I'll let three real machines show you. INTRO LINE: it's eighty-three seconds, and the narration does the work. " +
    "CUE: click it, let it run all the way through, don't talk over it. It covers a new excavator landing in the safe-add bucket, a dozer with an old rental record somewhere else landing in the quiet-add bucket, and a machine another dealer actively owns going to reviewed transfer. " +
    "AFTER: those are real serials out of our fleet. The excavator actually went into CCAT through this system last week."
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
    "SCRIPT: When we actually change something there's one door, and it's these procedures, one for each kind of action. You hand it a serial number and it pulls everything else itself, searches Cat live, and refuses anything outside its lane. Record already exists, another dealer owns it, it's an attachment with a made-up serial, or Cat already rejected that serial once. They start in preview mode: they show you exactly what would go out, and send nothing. Actually running it is a separate, deliberate step. " +
    "CUE payload rule: this one came out of the data review with Vinod and Michele. Cat's model names are the ones their systems recognize, so we never write ours over the top of them. " +
    "CUE raw tools: the low-level API procedures are still there for troubleshooting, but nobody's getting access to those. People get the audited wrappers, that's it."
  );
  panel(s, 0.6, 1.7, 6.0, 4.9);
  s.addText("CCAT_ADD_ASSET  ·  CCAT_ADD_INVENTORY", { x: 0.9, y: 1.95, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: BLUE });
  s.addText("CCAT_REASSIGN_CUSTOMER", { x: 0.9, y: 2.35, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: ORANGE });
  s.addText([
    { text: "Look up the machine by serial number, pull everything else from our records", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Search CCAT live before acting; never write blind", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Refuse anything outside their lane: existing records, other owners, attachments, known-bad serials", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Dry-run by default: show exactly what would be sent, send nothing", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "On execute: one API call, one audit row, one history snapshot", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.9, y: 2.95, w: 5.4, h: 3.5, margin: 0, fontFace: F, fontSize: 13.5, lineSpacing: 18 });
  panel(s, 6.9, 1.7, 5.8, 2.3);
  s.addText("Payload rule", { x: 7.2, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 16, bold: true, color: YEL });
  s.addText("Where CCAT already has a record, its model and year are the versions Caterpillar recognizes, so our requests reuse them. Our data only goes to Caterpillar for machines it has never seen.", {
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
    "SCRIPT: Here's the pipeline that's running today, end to end. NAXT data lands every morning. A scan picks out the inventory machines that qualify, oldest first. Before we call Cat at all, we drop everything we already know about. Machines we added before, machines already in CCAT, machines another dealer has locked up. Those cost us nothing to skip. Whatever's left gets one search each, and we save every answer whether we act on it or not. The dry run lets me eyeball the batch first. Then the adds go under our inventory account, each one with a full audit row behind it. If one fails, it's logged once and we never blindly retry it. Then the one manual step, which is Lindsay doing the subscription. And the same sweep that finds the problems comes back around and confirms each machine landed. " +
    "CUE: week one was fifty Monday and fifty Wednesday. That cap is there on purpose, to protect the manual step."
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
  node(3.75, 1.75, BLUE, "2 · Candidate scan", "Blank customer · Cat make · active · New · no attachments. Oldest first by modified date");
  node(6.9,  1.75, TEAL, "3 · Free skips", "Already added, already in CCAT, or dealer-blocked. Dropped from history at zero API cost");
  node(10.05,1.75, TEAL, "4 · Live CCAT search", "One API search per machine; every answer saved to history");
  arrowR(3.42, 2.4); arrowR(6.57, 2.4); arrowR(9.72, 2.4);
  s.addShape("downArrow", { x: 11.28, y: 3.42, w: 0.28, h: 0.76, fill: { color: DIM }, line: { type: "none" } });

  // row 2 (right to left)
  node(10.05, 4.3, YEL,    "5 · Dry-run review", "The would_add list previewed. Nothing sent yet");
  node(6.9,   4.3, GREEN,  "6 · Add to INT00495", "Executed with audit row + Cat tracking ID; failures logged once, auto-skipped after");
  node(3.75,  4.3, ORANGE, "7 · Subscription", "Lindsay registers Product Link in the dealer portal. The manual step that sets the pace");
  node(0.6,   4.3, GREEN,  "8 · Sweep verifies", "Machine lands DONE_INVENTORY on the dashboard. Loop closed");
  arrowL(9.72, 4.95); arrowL(6.57, 4.95); arrowL(3.42, 4.95);

  panel(s, 0.6, 6.15, 12.1, 0.9, PANEL2);
  s.addText([
    { text: "Live in production.  ", options: { bold: true, color: GREEN } },
    { text: "100 machines through this pipe in the first week: 50 Monday, 50 Wednesday. Capped on purpose: step 7 is a person.", options: { color: INK } },
  ], { x: 0.9, y: 6.32, w: 11.55, h: 0.6, margin: 0, fontFace: F, fontSize: 14 });
}

// ---------------------------------------------- 7c · phase two flow diagram
{
  const s = base("Phase two flow: sold machine → its buyer (designed)", "act · flow");
  s.addNotes(
    "SCRIPT: Phase two is scoped, and the detection half already runs. The minute a customer number shows up on an inventory machine in NAXT, it lands on the dashboard. There are a hundred eighty sitting there right now, and every one of them is a manual move for Lindsay today. What's new since the August sixth meeting is that Cat's customer master is shared straight into our Snowflake. So the question that was blocking us was whether the buyer actually exists properly on Cat's side. We can answer that now, before we touch anything. If the buyer checks out, the move is automatic and audited, and the old inventory record expires on its own. If they don't, it goes on a list for the Customer Admin Tool instead of blowing up halfway through. " +
    "CUE: until this is built and approved nothing changes. PDI email, Lindsay moves it within about a day. This replaces that one step."
  );
  const node = (x, y, w, h, c, t, body) => {
    panel(s, x, y, w, h);
    s.addText(t, { x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.55, margin: 0, fontFace: F, fontSize: 13, bold: true, color: c, lineSpacing: 15 });
    s.addText(body, { x: x + 0.16, y: y + 0.68, w: w - 0.32, h: h - 0.8, margin: 0, fontFace: F, fontSize: 10.5, color: INK, lineSpacing: 14 });
  };
  const arrowR = (x, y) => s.addShape("rightArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });

  node(0.6,  2.6, 2.7, 1.7, BLUE, "1 · Machine sells", "A customer number appears on an inventory machine in NAXT");
  node(3.75, 2.6, 2.7, 1.7, YEL,  "2 · Dashboard flags it", "P3_INV_TO_CUSTOMER, caught by the daily sweep. 180 machines on this list today");
  node(6.9,  2.6, 2.7, 1.7, TEAL, "3 · CCID check", "Cat's customer master, shared into Snowflake, answers: does the buyer resolve to a CCID?");
  arrowR(3.42, 3.32); arrowR(6.57, 3.32);

  s.addShape("line", { x: 9.68, y: 2.72, w: 0.5, h: 0.62, flipV: true, line: { color: GREEN, width: 2.25, endArrowType: "triangle" } });
  s.addShape("line", { x: 9.68, y: 3.55, w: 0.5, h: 1.5, line: { color: ORANGE, width: 2.25, endArrowType: "triangle" } });
  s.addText("CCID exists", { x: 9.42, y: 2.28, w: 1.3, h: 0.3, margin: 0, fontFace: F, fontSize: 10, bold: true, color: GREEN });
  s.addText("no CCID", { x: 9.5, y: 4.45, w: 1.1, h: 0.3, margin: 0, fontFace: F, fontSize: 10, bold: true, color: ORANGE });

  node(10.25, 1.8, 2.5, 2.05, GREEN, "4a · Automated move", "Reassign INT00495 → the buyer's account. Old record auto-expires. One audit row, one API call.");
  node(10.25, 4.35, 2.5, 2.05, ORANGE, "4b · Human queue", "Customer Admin Tool creates the CCID first, then the move retries. No mid-flight failures.");

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
    "SCRIPT: Every action writes one row, whether it worked or failed. What we did, which machine, the customer before and after, the exact payload we sent, Cat's full response, and their tracking ID, which means Cat support can trace any call we've ever made. It also saves a snapshot of what Cat had right before we touched it, so if we ever have to unwind something, we know exactly what to put back. " +
    "CUE right panel: append-only isn't a policy we promise to follow. It's enforced by permissions. The tooling can add rows and read rows. It physically can't edit or delete them. " +
    "IF ASKED who and when: every row carries the user, the timestamp, and the shift date. That's where the per-shift reporting comes from."
  );
  panel(s, 0.6, 1.7, 7.7, 4.9);
  s.addText("Every executed or failed action writes one row:", { x: 0.95, y: 2.0, w: 6.9, h: 0.4, margin: 0, fontFace: F, fontSize: 15, bold: true, color: INK });
  s.addText([
    { text: "What: action type, phase, machine, model, ownership type", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Customer before  →  customer after (the reassignment story)", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "The exact payload sent, and Caterpillar's full response", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Caterpillar's tracking ID, so their support can trace any call", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "A snapshot of CCAT taken just before we acted (the undo reference)", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Who, when, and which shift, which feeds the per-shift reporting", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.95, y: 2.5, w: 7.0, h: 3.9, margin: 0, fontFace: F, fontSize: 14, lineSpacing: 19 });
  panel(s, 8.6, 1.7, 4.1, 4.9, PANEL2);
  s.addText("Append-only\nby permission", { x: 8.9, y: 2.1, w: 3.5, h: 1.1, margin: 0, fontFace: F, fontSize: 22, bold: true, color: YEL });
  s.addText("The tooling role can INSERT and SELECT, nothing else. Audit rows cannot be edited or deleted, even by accident.\n\nKept forever. This table is the compliance trail.", {
    x: 8.9, y: 3.35, w: 3.5, h: 3.0, margin: 0, fontFace: F, fontSize: 14, color: INK, lineSpacing: 19,
  });
}

// -------------------------------------------------------------- 9 · history
{
  const s = base("CCAT_HISTORY: what the data was, at every point", "prove");
  s.addNotes(
    "SCRIPT: History answers a different question. Not what did we do, but what did the data look like at any given point, no matter who changed it. Every time we look at a machine we fingerprint Cat's whole answer. Same fingerprint as last time, we store nothing. Different, we save the full snapshot. So a machine we check fifty times with no changes costs us one row. It only grows when something actually changed. " +
    "CUE: this costs zero extra API calls. It's built entirely out of answers the sweep was already getting and throwing away. " +
    "NICE DETAIL if time: no records at all counts as a version too, so a machine showing up in CCAT, or disappearing from it, leaves a trail."
  );
  panel(s, 0.6, 1.7, 7.7, 4.9);
  s.addText([
    { text: "Every distinct version, any author", options: { bold: true, color: TEAL, breakLine: true } },
    { text: "Each check fingerprints Caterpillar's full answer for the machine. New fingerprint = the full snapshot is stored. Same fingerprint = nothing.", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Cheap by design", options: { bold: true, color: TEAL, breakLine: true } },
    { text: "A machine checked fifty times with no change stores one row, not fifty. Built entirely from data the sweep already fetches, so zero extra API calls.", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Absence counts too", options: { bold: true, color: TEAL, breakLine: true } },
    { text: "“No records at all” is itself a version, so a machine appearing in, or vanishing from, CCAT leaves a trail.", options: { color: INK, breakLine: true } },
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
    "SCRIPT: Here's why having both matters. Audit is what we did. History is what the data was. Subtract one from the other and you get something we've never had: changes nobody here made. If history shows a machine's record changed and there's no audit row from us anywhere near that time, then Cat's systems, another dealer, or somebody in a portal did it. And now we see it, instead of finding out months later. " +
    "CUE: pause on the banner. This is the kind of change that used to be completely invisible to us. " +
    "IF ASKED for an example: another dealer filing a transfer against one of our machines shows up this way within a day."
  );
  panel(s, 0.6, 1.75, 5.9, 3.3);
  s.addText("AUDIT", { x: 0.9, y: 2.05, w: 5.2, h: 0.45, margin: 0, fontFace: F, fontSize: 20, bold: true, color: YEL });
  s.addText("What WE did.\nActions, payloads, outcomes, tracking IDs. Signed, dated, attributable.", {
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
    { text: "A version change with no audit row nearby means Caterpillar, another dealer, or a portal edit changed the data. That is the class of change that used to be invisible.", options: { color: INK } },
  ], { x: 0.9, y: 5.55, w: 11.55, h: 0.95, margin: 0, fontFace: F, fontSize: 15.5, lineSpacing: 21 });
}

// ------------------------------------------------------ 11 · rhythm+metrics
{
  const s = base("The operating rhythm", "run it");
  s.addNotes(
    "SCRIPT: Day to day it looks like this. At night the sweep runs on its own. Read-only, nobody watching it. Weekly we work the safe piles in capped batches, the adds and quiet fixes nobody outside the company can see. And every shift, about ten reviewed decisions: the transfers and the judgment calls, each one audited as it happens. " +
    "CUE metrics: the headline number is deliberately not the match rate. It's the customer assignment rate, meaning what percentage of machines are sitting on the right customer in Cat's system, because that's the outcome the business asked for. Everything else is supporting detail. " +
    "ASK if appropriate: the one approval I need off this slide is scheduling the nightly sweep as an automated task."
  );
  const cols = [
    { t: "Nightly", c: BLUE, body: "The sweep checks its batch, refreshes verdicts, and grows coverage. Read-only, automatic." },
    { t: "Weekly", c: TEAL, body: "Batched safe adds and quiet fixes run under volume caps. A scoreboard snapshot tracks the backlog shrinking." },
    { t: "Every shift", c: RED, body: "About ten reviewed decisions: transfers, inbound claims, judgment calls, all audited as they happen." },
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
    { text: "Customer assignment rate, the headline: % of machines on the correct customer in CCAT", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 5 } },
    { text: "Coverage and match rate: how much of the fleet is verified, and how much agrees", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 5 } },
    { text: "Backlog by phase, resolutions per shift, and transfers in flight", options: { bullet: true, color: INK, breakLine: true } },
  ], { x: 0.9, y: 5.3, w: 11.4, h: 1.3, margin: 0, fontFace: F, fontSize: 13.5, lineSpacing: 18 });
}

// ----------------------------------------------------------- 12 · the rules
{
  const s = base("The rules that keep it safe", "guardrails");
  s.addNotes(
    "SCRIPT: I won't read all six. Three of them matter most. No bulk claims, ever. A blanket push would fire formal transfer requests at real dealers across the network, and it can quietly overwrite our own records. The system is built so that shortcut isn't possible, not just discouraged. Shared customers get excluded automatically: when another dealer serves the same national account, that machine goes to a conversation, never a claim. And production awareness. Cat has no test environment, so every write gets treated as real, because it is. Previewed first, logged after. " +
    "CUE: these guardrails live in the tools themselves, so they hold no matter who ends up with access."
  );
  const rules = [
    { t: "Dry-run first, always", body: "Every write shows its exact payload and sends nothing until deliberately executed.", c: GREEN },
    { t: "No bulk claims, ever", body: "Blanket adds would fire transfer requests at real dealers and can silently overwrite our own records.", c: RED },
    { t: "Shared customers excluded", body: "When another dealer serves the same customer, the machine routes to coordination, never a claim.", c: TEAL },
    { t: "CCAT owns the model", body: "We never push our model names over Caterpillar's. Their versions are what Cat systems recognize.", c: YEL },
    { t: "Production awareness", body: "Caterpillar has no test environment. Every write is treated as real, because it is.", c: ORANGE },
    { t: "Compliance tables are sacred", body: "Audit and history are append-only by permission and kept forever. Off the reset list permanently.", c: BLUE },
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
    "SCRIPT: Straight talk on where this stands. Everything on the left isn't a plan, it's running. The sweep has checked thousands of machines. Real equipment has gone into Cat's production system through the audited procedures, and there are audit rows to prove it. The right column is scheduling and scale, not inventing anything new. The nightly task is a small script. The batch runner and the rest of the wrappers reuse patterns we've already proven. The access roles are written and waiting on approval to roll out. " +
    "CUE: the honest version is that the engineering risk is behind us. What's left is operational."
  );
  panel(s, 0.6, 1.75, 5.9, 4.9);
  s.addText("Working today", { x: 0.9, y: 2.05, w: 5.2, h: 0.45, margin: 0, fontFace: F, fontSize: 19, bold: true, color: GREEN });
  s.addText([
    { text: "Detection sweep, proven at thousand-machine scale", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Phase dashboard: phase, action, and why per machine", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Audited procedures for adds, inventory adds, and customer reassignment, with real adds executed in production", options: { bullet: true, color: INK, breakLine: true, paraSpaceAfter: 6 } },
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
    "SCRIPT: One sentence to take away: the system sorts every machine, and people only decide the ones that need judgment. That's the whole design. Automation where it's safe, a person where it matters, and a permanent record of both. Happy to take questions, and if you want the detail behind any slide, the rollout plan and the SQL reference have all of it. " +
    "CLOSE: the four decisions I need from this room are in the rollout plan. Approve the phased order, the nightly job, the pilot batches, and the review staffing."
  );
}

pres.writeFile({ fileName: "Snowflake_CCAT_Reconciliation.pptx" }).then(() => console.log("WROTE Snowflake_CCAT_Reconciliation.pptx"));
