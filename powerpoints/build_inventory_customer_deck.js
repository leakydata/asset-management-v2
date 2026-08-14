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
  s.addText("Inventory First. Then the Right Customer.", {
    x: 0.9, y: 2.15, w: 11.5, h: 1.0, margin: 0,
    fontFace: F, fontSize: 42, bold: true, color: INK,
  });
  s.addText("Phase one of CCAT automation is live — and the customer-identity work that makes phase two possible", {
    x: 0.9, y: 3.2, w: 11.0, h: 0.6, margin: 0,
    fontFace: F, fontSize: 20, color: DIM,
  });
  [GREEN, BLUE, TEAL, ORANGE].forEach((c, i) => dot(s, 0.95 + i * 0.34, 4.05, c, 0.2));
  s.addText("Cleveland Brothers  ·  Data & Analytics", {
    x: 0.9, y: 6.7, w: 8, h: 0.4, margin: 0, fontFace: F, fontSize: 13, color: DIM,
  });
  s.addNotes(
    "SCRIPT: Two things to cover today and they're connected. First, the inventory automation we agreed on back on the sixth is running. We ran a hundred machines through it this week, fifty Monday and fifty Wednesday, and I'll walk you through how it works. Second is something we hit while working the mismatch list: we figured out how to tell whether a machine is actually on the wrong customer, using Cat's own data. That turned out to matter more than I expected, and it's what unblocks phase two. " +
    "CUE: both of them land on the same scorecard, so that's where I'll start."
  );
}

// ------------------------------------------------- 2 · trifecta / DNA
{
  const s = base("The scoreboard this feeds: Trifecta and DNA", "why it matters");
  panel(s, 0.6, 1.7, 5.9, 4.15);
  s.addText("TRIFECTA", { x: 0.9, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 18, bold: true, color: YEL });
  s.addText("Caterpillar's data-quality score. A customer only counts when ALL legs pass: CCID, company data, a valid contact, and 100% asset match on priority assets.", {
    x: 0.9, y: 2.42, w: 5.3, h: 1.15, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
  s.addText("92.8%", { x: 0.9, y: 3.7, w: 2.4, h: 0.7, margin: 0, fontFace: F, fontSize: 40, bold: true, color: ORANGE });
  s.addText("asset match (2025 baseline) — maintained by hand today. The CCID leg sits at 92.7% — the same customer identity this project now works with directly.", {
    x: 3.3, y: 3.72, w: 2.95, h: 1.4, margin: 0, fontFace: F, fontSize: 12, color: DIM, lineSpacing: 15 });
  s.addText("Worth 15% of the dealer excellence scorecard", { x: 0.9, y: 5.25, w: 5.3, h: 0.4, margin: 0, fontFace: F, fontSize: 13.5, bold: true, color: INK });
  panel(s, 6.8, 1.7, 5.9, 4.15);
  s.addText("DNA — DATA NOTIFICATION ALERTS", { x: 7.1, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 18, bold: true, color: TEAL });
  s.addText("Caterpillar flags our prioritized data issues and grades how well we resolve them. Many asset alerts trace to the same root cause: ownership records that do not match reality.", {
    x: 7.1, y: 2.42, w: 5.3, h: 1.3, margin: 0, fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 18,
  });
  s.addText("10%", { x: 7.1, y: 3.85, w: 2.0, h: 0.7, margin: 0, fontFace: F, fontSize: 40, bold: true, color: TEAL });
  s.addText("of the scorecard rides on DNA resolution — prevention at the source beats resolving alerts after they fire", {
    x: 9.2, y: 3.87, w: 3.05, h: 1.2, margin: 0, fontFace: F, fontSize: 12.5, color: DIM, lineSpacing: 16 });
  panel(s, 0.6, 6.1, 12.1, 0.85, PANEL2);
  s.addText([
    { text: "Inventory adds grow the matched population.  ", options: { bold: true, color: GREEN } },
    { text: "Customer fixes repair the asset-match leg.  ", options: { bold: true, color: ORANGE } },
    { text: "Both stand on CCAT ownership records — exactly what this work manages.", options: { color: INK } },
  ], { x: 0.9, y: 6.28, w: 11.55, h: 0.55, margin: 0, fontFace: F, fontSize: 14.5 });
  s.addNotes(
    "SCRIPT: Quick word on why this matters upstairs. Trifecta is Cat's data quality score and it's strict. A customer only counts if every leg passes, including a hundred percent asset match on their priority assets. So one bad machine fails that whole customer. We're sitting around ninety-three percent on asset match and it's all maintained by hand today. Look at the CCID leg too, ninety-two point seven. That's the same CCID I'll be talking about in the second half. DNA is the other side of it: Cat flags our data problems and grades how fast we clear them, and a lot of those asset alerts come back to ownership records being wrong. " +
    "CUE: the banner is the point. Inventory adds grow the pool that gets matched, customer fixes repair the match rate itself. Same records, two different scores."
  );
}

// ------------------------------------------------- 3 · why inventory first
{
  const s = base("Why inventory came first", "phase one");
  const rows = [
    { c: GREEN, t: "It is the cleanest case", body: "An inventory machine has a blank customer in NAXT and belongs under exactly one well-known account (INT00495). There is no customer identity to get wrong." },
    { c: BLUE, t: "It is where the pain is", body: "More than half of Lindsay's churn: PDI email arrives, key the machine in by hand, subscribe it, move it when it sells — twenty minutes and two sittings per machine, 20–70 emails a day." },
    { c: ORANGE, t: "The backlog was real", body: "~718 New machines sat in inventory with no CCAT record, some since 2025 — invisible to Caterpillar and to the scorecard the whole time." },
    { c: TEAL, t: "The goal in one line", body: "Every machine already in CCAT, subscribed and reporting, by the time anyone needs it. When it sells, the only remaining work is one customer move." },
  ];
  rows.forEach((r, i) => {
    const y = 1.75 + i * 1.18;
    panel(s, 0.6, y, 12.1, 1.02);
    dot(s, 0.9, y + 0.42, r.c, 0.18);
    s.addText(r.t, { x: 1.25, y: y + 0.14, w: 3.6, h: 0.78, margin: 0, fontFace: F, fontSize: 15, bold: true, color: r.c, valign: "middle" });
    s.addText(r.body, { x: 5.0, y: y + 0.1, w: 7.45, h: 0.86, margin: 0, fontFace: F, fontSize: 12.5, color: INK, valign: "middle", lineSpacing: 16 });
  });
  s.addText("Rules settled August 6: Cat make · active · no attachments · fleet type New only to start · Re-rent never touched · modified-date drives detection", {
    x: 0.6, y: 6.55, w: 12.1, h: 0.5, margin: 0, fontFace: F, fontSize: 12.5, color: DIM, italic: true,
  });
  s.addNotes(
    "SCRIPT: Why inventory first, out of everything we could have started with? Because it's the one case with no judgment call in it. Customer's blank in NAXT, there's exactly one account it goes to, so there's nothing to get wrong about who owns it. That turns out to be the hard part everywhere else, and you'll see that in a few slides. It's also where the manual work piles up. Lindsay said more than half her day is this same loop, over and over. And the backlog is real: around seven hundred machines Cat has never seen, some of them sitting there since last year. " +
    "CUE: read the goal line out loud. In CCAT, subscribed, and reporting before anybody needs it, and one move left to do when it sells. That last part is phase two."
  );
}

// ------------------------------------------------- 4 · phase one flow
{
  const s = base("Phase one flow: inventory adds (live)", "phase one · flow");
  s.addNotes(
    "SCRIPT: Here's the whole thing end to end. NAXT data lands every morning. A scan picks out the machines that qualify, oldest first. Before we call Cat at all, we drop everything we already know about. Machines we added before, machines already in CCAT, machines another dealer has locked up. Those cost us nothing to skip. Whatever's left gets one search each, and we save every answer whether we act on it or not. Then a dry run so I can eyeball the list before anything goes out. Then the adds themselves, each one under our inventory account with a full audit row behind it. If one fails, it's logged once and we never blindly retry it. Step seven is Lindsay doing the subscription. And then the sweep comes back around and confirms the machine landed where it should. " +
    "CUE: the fifty-at-a-time cap is there because of step seven. The computer was never the bottleneck."
  );
  const node = (x, y, c, t, body) => {
    panel(s, x, y, 2.75, 1.55);
    s.addText(t, { x: x + 0.14, y: y + 0.1, w: 2.47, h: 0.5, margin: 0, fontFace: F, fontSize: 12.5, bold: true, color: c, lineSpacing: 14 });
    s.addText(body, { x: x + 0.14, y: y + 0.6, w: 2.47, h: 0.9, margin: 0, fontFace: F, fontSize: 10, color: INK, lineSpacing: 13 });
  };
  const arrowR = (x, y) => s.addShape("rightArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });
  const arrowL = (x, y) => s.addShape("leftArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });

  node(0.6,  1.75, BLUE, "1 · NAXT refresh", "Golden equipment data lands in Snowflake each morning");
  node(3.75, 1.75, BLUE, "2 · Candidate scan", "Blank customer · Cat make · active · New · no attachments — oldest first by modified date");
  node(6.9,  1.75, TEAL, "3 · Free skips", "Already added, already in CCAT, or dealer-blocked — dropped from history at zero API cost");
  node(10.05,1.75, TEAL, "4 · Live CCAT search", "One API search per machine; every answer saved to history");
  arrowR(3.42, 2.4); arrowR(6.57, 2.4); arrowR(9.72, 2.4);
  s.addShape("downArrow", { x: 11.28, y: 3.42, w: 0.28, h: 0.76, fill: { color: DIM }, line: { type: "none" } });

  node(10.05, 4.3, YEL,    "5 · Dry-run review", "The would_add list previewed — nothing sent yet");
  node(6.9,   4.3, GREEN,  "6 · Add to INT00495", "Executed with audit row + Cat tracking ID; failures logged once, auto-skipped after");
  node(3.75,  4.3, ORANGE, "7 · Subscription", "Lindsay registers Product Link in the dealer portal — the manual step that sets the pace");
  node(0.6,   4.3, GREEN,  "8 · Sweep verifies", "Machine lands DONE_INVENTORY on the dashboard — loop closed");
  arrowL(9.72, 4.95); arrowL(6.57, 4.95); arrowL(3.42, 4.95);

  panel(s, 0.6, 6.15, 12.1, 0.9, PANEL2);
  s.addText([
    { text: "First week:  ", options: { bold: true, color: GREEN } },
    { text: "100 machines through this pipe — 50 Monday, 50 Wednesday — 39 subscriptions registered, zero surprises. Capped on purpose: step 7 is a person.", options: { color: INK } },
  ], { x: 0.9, y: 6.32, w: 11.55, h: 0.6, margin: 0, fontFace: F, fontSize: 14 });
}

// ------------------------------------------------- 5 · the mismatch problem
{
  const s = base("Then the harder question: is the customer wrong?", "customer identity");
  panel(s, 0.6, 1.7, 5.9, 4.4);
  s.addText("What the data said", { x: 0.9, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 17, bold: true, color: RED });
  s.addText("1,327", { x: 0.9, y: 2.45, w: 3.1, h: 0.9, margin: 0, fontFace: F, fontSize: 54, bold: true, color: RED });
  s.addText("machines where the customer number on our CCAT record differs from the customer number in NAXT.", {
    x: 3.85, y: 2.55, w: 2.45, h: 1.3, margin: 0, fontFace: F, fontSize: 13, color: INK, lineSpacing: 17,
  });
  s.addText("Reassigning all 1,327 would be the obvious move — and for hundreds of them, it would be wrong.", {
    x: 0.9, y: 4.1, w: 5.3, h: 1.6, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: YEL, lineSpacing: 20,
  });
  panel(s, 6.8, 1.7, 5.9, 4.4);
  s.addText("The catch: AMERIKOHL", { x: 7.1, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 17, bold: true, color: TEAL });
  s.addText([
    { text: "C00025828   →  AMERIKOHL MINING INC", options: { color: INK, breakLine: true, paraSpaceAfter: 4 } },
    { text: "C0B094893   →  AMERIKOHL MINING INC", options: { color: INK, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Two different customer numbers. One real customer.", options: { bold: true, color: YEL, breakLine: true, paraSpaceAfter: 10 } },
    { text: "A customer number (DCN) is an account — and one customer picks up several over the years: a second branch account, a legacy number, an account at another dealer entirely.", options: { color: INK, breakLine: true } },
  ], { x: 7.1, y: 2.45, w: 5.3, h: 3.4, margin: 0, fontFace: F, fontSize: 13, lineSpacing: 17 });
  panel(s, 0.6, 6.3, 12.1, 0.75, PANEL2);
  s.addText([
    { text: "The account number is not the customer.  ", options: { bold: true, color: YEL } },
    { text: "Caterpillar's CCID — the customer identity Trifecta itself grades — is the ground truth of “who.”", options: { color: INK } },
  ], { x: 0.9, y: 6.44, w: 11.55, h: 0.5, margin: 0, fontFace: F, fontSize: 14.5 });
  s.addNotes(
    "SCRIPT: With inventory running, we went after the wrong-customer pile. Right now that's about thirteen hundred machines where our customer number and Cat's don't match. Obvious move is to go fix all of them. Here's why that would've been wrong for hundreds of them. Amerikohl. Two completely different customer numbers, and Cat says they're the same company. Customer numbers are just accounts, and a customer picks up extras over the years. A second branch, an old number from a conversion, an account at another dealer. What actually identifies the customer is the CCID. And notice, that's the same CCID Trifecta grades us on. " +
    "CUE: land the banner. The customer number is not the customer. Everything after this comes out of that one sentence."
  );
}

// ------------------------------------------------- 6 · policy + sources
{
  const s = base("The rule we set, and where the answers come from", "customer identity");
  panel(s, 0.6, 1.7, 5.9, 4.6);
  s.addText("The policy (Aug 11)", { x: 0.9, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 17, bold: true, color: YEL });
  s.addText([
    { text: "Multiple account numbers per customer are allowed — ours, multiples of ours, other dealers'.", options: { color: INK, breakLine: true, paraSpaceAfter: 8 } },
    { text: "CCID equality is the ONLY customer-correctness test.", options: { bold: true, color: INK, breakLine: true, paraSpaceAfter: 8 } },
    { text: "Same CCID, different account → nothing to fix.", options: { color: GREEN, breakLine: true, paraSpaceAfter: 4 } },
    { text: "Different CCID → truly the wrong customer.", options: { color: RED, breakLine: true, paraSpaceAfter: 4 } },
    { text: "Cat's own systems disagree → a person looks, never automation.", options: { color: ORANGE, breakLine: true } },
  ], { x: 0.9, y: 2.45, w: 5.3, h: 3.7, margin: 0, fontFace: F, fontSize: 13.5, lineSpacing: 18 });
  panel(s, 6.8, 1.7, 5.9, 4.6);
  s.addText("Four sources, cheapest first", { x: 7.1, y: 1.95, w: 5.2, h: 0.4, margin: 0, fontFace: F, fontSize: 17, bold: true, color: TEAL });
  s.addText([
    { text: "1 · Our observation history — every Cat API answer we ever stored already carried the CCID. Free, retroactive.", options: { color: INK, breakLine: true, paraSpaceAfter: 7 } },
    { text: "2 · The crosswalk table — every answer any lookup ever produced, kept forever. One API call becomes permanent knowledge.", options: { color: INK, breakLine: true, paraSpaceAfter: 7 } },
    { text: "3 · Cat's customer master, shared into Snowflake — resolves customers who own zero equipment. Refreshed by Cat, not by us.", options: { color: INK, breakLine: true, paraSpaceAfter: 7 } },
    { text: "4 · The Cat API — last resort, for genuinely new customers only.", options: { color: INK, breakLine: true } },
  ], { x: 7.1, y: 2.45, w: 5.3, h: 3.7, margin: 0, fontFace: F, fontSize: 13, lineSpacing: 17 });
  panel(s, 0.6, 6.5, 12.1, 0.6, PANEL2);
  s.addText([
    { text: "The system gets cheaper the longer it runs:  ", options: { bold: true, color: YEL } },
    { text: "every sweep, lookup, and Cat refresh widens the free layers. ~100 API calls resolved the entire backlog.", options: { color: INK } },
  ], { x: 0.9, y: 6.6, w: 11.55, h: 0.42, margin: 0, fontFace: F, fontSize: 13.5 });
  s.addNotes(
    "SCRIPT: Two things made this work. First we made a call: multiple customer numbers under one customer are fine. Ours, several of ours, another dealer's. Doesn't matter. The only test that counts is whether the CCID matches. Same CCID, nothing to fix. Different CCID, it's really wrong. And every so often Cat's own systems disagree with each other, usually after they've merged customers on their end. Those we set aside for a person to look at. We don't let the automation guess. Second is where the answers come from, and this one still bugs me a little. Every response we'd ever saved from Cat already had the CCID sitting in it. We just weren't reading that field. So we got it back to day one, for free. On top of that, we keep every answer we look up, and Cat shares their customer master right into our Snowflake. Calling the API is the last thing we try now, not the first. " +
    "CUE: the whole backlog cost about a hundred calls to clear. Asking the same question again is free."
  );
}

// ------------------------------------------------- 7 · the verdict numbers
{
  const s = base("1,327 “wrong customers” — what they really are", "customer identity");
  const rows = [
    { c: RED,    n: "723", t: "Truly wrong customer", body: "CCIDs genuinely differ — a validated, CCID-proven reassignment worklist, awaiting this group's go-ahead" },
    { c: GREEN,  n: "520", t: "Same customer, different account", body: "CCID matches — the AMERIKOHL pattern. Correct by policy. No action, forever." },
    { c: ORANGE, n: "84",  t: "Genuinely unresolved", body: "Customers unknown to Cat's own customer master — the real CCID gap, now countable — plus a handful where Cat's systems disagree; quarantined for humans" },
    { c: BLUE,   n: "4.5×", t: "The pool grew with coverage", body: "294 machines when the CCID test first ran (Aug 11) → 1,327 as the sweep widened across the fleet — same test, same discipline, at scale" },
  ];
  rows.forEach((r, i) => {
    const y = 1.75 + i * 1.14;
    panel(s, 0.6, y, 12.1, 0.98);
    s.addText(r.n, { x: 0.9, y: y + 0.08, w: 1.5, h: 0.82, margin: 0, fontFace: F, fontSize: 34, bold: true, color: r.c, valign: "middle" });
    s.addText(r.t, { x: 2.5, y: y + 0.14, w: 3.9, h: 0.7, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: r.c, valign: "middle" });
    s.addText(r.body, { x: 6.5, y: y + 0.1, w: 5.95, h: 0.78, margin: 0, fontFace: F, fontSize: 12, color: INK, valign: "middle", lineSpacing: 15 });
  });
  panel(s, 0.6, 6.4, 12.1, 0.7, PANEL2);
  s.addText([
    { text: "520 false alarms cleared forever.  ", options: { bold: true, color: YEL } },
    { text: "And the 723 that are real are proven machine by machine, not suspected. Each fix can flip a whole customer back to Trifecta-valid.", options: { color: INK } },
  ], { x: 0.9, y: 6.53, w: 11.55, h: 0.48, margin: 0, fontFace: F, fontSize: 14 });
  s.addNotes(
    "SCRIPT: So here's what the CCID test did to those thirteen hundred. Seven hundred twenty-three are genuinely on the wrong customer. That list is proven now, one machine at a time, off Cat's own data, and it's ready to work whenever you tell me to go. Five hundred twenty were the Amerikohl thing, same customer with a different account number. Those are closed for good, and without the CCID check every one of them was a wrong reassignment waiting to happen. Eighty-four we still can't answer. Mostly customers Cat's own master doesn't know about, plus a few where Cat contradicts itself. That eighty-four is the CCID gap everybody's been talking about, except now it's a number instead of a worry. Last row's worth a mention too: this pool more than quadrupled in a week as the sweep covered more of the fleet, and the test just scaled with it. " +
    "CUE: the banner. Five hundred twenty we never have to touch, seven hundred twenty-three we can prove. And remember one bad machine fails its whole customer on Trifecta, so these are customer-level fixes, not machine-level."
  );
}

// ------------------------------------------------- 8 · phase two flow
{
  const s = base("Phase two flow: sold machine → its buyer (designed)", "what's next · flow");
  s.addNotes(
    "SCRIPT: This is where the CCID work pays off past just cleanup. It unblocks phase two. The detection half already runs. The minute a customer shows up on an inventory machine in NAXT, it lands on the dashboard. There are a hundred eighty sitting there right now, and every one of those is a manual move for Lindsay today. What stopped us on the sixth was John's question. Does the buyer even have a CCID? We can answer that in Snowflake now, before we touch anything, off Cat's customer master. If the buyer checks out, the move is automatic and audited, and Cat expires the old inventory record on its own. If they don't, it goes on a list for the Customer Admin Tool instead of blowing up halfway through. " +
    "CUE: until this is built and approved nothing changes. PDI email, Lindsay, about a day. This replaces that one step."
  );
  const node = (x, y, w, h, c, t, body) => {
    panel(s, x, y, w, h);
    s.addText(t, { x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.55, margin: 0, fontFace: F, fontSize: 13, bold: true, color: c, lineSpacing: 15 });
    s.addText(body, { x: x + 0.16, y: y + 0.68, w: w - 0.32, h: h - 0.8, margin: 0, fontFace: F, fontSize: 10.5, color: INK, lineSpacing: 14 });
  };
  const arrowR = (x, y) => s.addShape("rightArrow", { x, y, w: 0.32, h: 0.26, fill: { color: DIM }, line: { type: "none" } });

  node(0.6,  2.6, 2.7, 1.7, BLUE, "1 · Machine sells", "A customer number appears on an inventory machine in NAXT");
  node(3.75, 2.6, 2.7, 1.7, YEL,  "2 · Dashboard flags it", "P3_INV_TO_CUSTOMER — caught by the daily sweep. 180 machines on this list today");
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

// ------------------------------------------------- 9 · asks / closing
{
  const s = base("Where this leaves us", "decisions");
  const rows = [
    { c: GREEN, t: "Keep the inventory cadence", body: "50 Monday + 50 Wednesday continues; the ~718 backlog clears in roughly 7 weeks at this pace — or faster if subscription bandwidth allows" },
    { c: RED, t: "Approve the reassignment lane", body: "723 CCID-proven wrong-customer machines, ready to work through the audited procedure — dry-run first, reviewed pace, same discipline as inventory" },
    { c: ORANGE, t: "The CCID gap is now a list, not a mystery", body: "84 machines whose customers are unknown to Cat's own master — named and counted; Customer Admin Tool work, and the exact customers phase two would stumble on" },
    { c: BLUE, t: "Green-light phase two design", body: "Detection runs, the CCID check exists, 180 machines are waiting — the build replaces a manual step Lindsay does today" },
  ];
  rows.forEach((r, i) => {
    const y = 1.75 + i * 1.14;
    panel(s, 0.6, y, 12.1, 0.98);
    dot(s, 0.9, y + 0.4, r.c, 0.18);
    s.addText(r.t, { x: 1.25, y: y + 0.14, w: 3.9, h: 0.7, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: r.c, valign: "middle" });
    s.addText(r.body, { x: 5.3, y: y + 0.1, w: 7.15, h: 0.78, margin: 0, fontFace: F, fontSize: 12.5, color: INK, valign: "middle", lineSpacing: 16 });
  });
  s.addText("The system sorts every machine. People decide only the ones that need judgment.", {
    x: 0.6, y: 6.5, w: 12.1, h: 0.5, margin: 0, fontFace: F, fontSize: 16, bold: true, color: YEL, align: "center",
  });
  s.addNotes(
    "SCRIPT: Four things to land on. One, the inventory runs keep going, fifty and fifty. At that pace the backlog's gone in about seven weeks, quicker if Lindsay ever gets help on the subscriptions. Two, and this is the real ask: let me start working the reassignment list. Seven hundred twenty-three machines, all proven, and I'd run it exactly like inventory. Dry run first, slow pace, everything audited. Three, no decision needed here, just so you know: the CCID gap is a list of eighty-four now, not a mystery. That's Customer Admin Tool work. Four, give me the go-ahead to design phase two. Everything it needs already exists and a hundred eighty machines are waiting on it. " +
    "CUE: end on the line at the bottom. The system sorts every machine, people decide only the ones that need judgment."
  );
}

pres.writeFile({ fileName: "CCAT_Inventory_and_Customer_Identity.pptx" }).then(() => console.log("WROTE CCAT_Inventory_and_Customer_Identity.pptx"));
