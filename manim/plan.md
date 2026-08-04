# The Decision Key

## Overview
- **Topic**: How the CCAT reconciliation system decides what is safe to do with one machine
- **Hook**: "One machine. A few questions."
- **Target Audience**: Cleveland Brothers leadership; no technical prerequisites
- **Estimated Length**: ~32 seconds
- **Key Insight**: The system sorts every machine automatically; people only decide the ones that need judgment
- **Resolution**: 1080p60
- **Aspect Ratio**: 16:9

## Narrative Arc
A simplified version of the rollout plan's dichotomous key is drawn as a live decision
tree. Two real machines walk it: a new excavator lands in the safe automated ADD bucket;
a machine another dealer claims lands in the human-reviewed TRANSFER bucket. The
contrast IS the message.

---

## Scene: DecisionKey (single scene, `decision_key.py`)
1. Title card: "One machine. A few questions."
2. The tree draws itself: 3 question boxes (blue), 3 terminals (green ADD /
   teal quiet add / red TRANSFER), YES/NO edges.
3. Machine A (`WKX53704`, green chip) traverses: no record -> nobody has one ->
   ADD. Yellow stamp: "audited + verified by next sweep".
4. Machine B (`9303`, red chip) traverses: no record -> another dealer -> they
   actively own it -> TRANSFER. Yellow stamp: "evidence checked -- about 10 per shift".
5. Closer over dimmed tree: "The system sorts every machine. / People decide only
   the ones that need judgment."

## Color Palette (matches the Remotion loop + phase dashboard)
- Background: #101c2c   - Questions: #4da3ff (blue)
- ADD: #34c759 (green)  - Quiet: #2ec4b6 (teal)  - TRANSFER: #ff5a5f (red)
- Stamps/accents: #ffc72c (Cat yellow)

## Render
    uv run --with manim manim -qh --disable_caching decision_key.py DecisionKey
    # output: media/videos/decision_key/1080p60/DecisionKey.mp4 (copied to final.mp4)

## Possible future companion
A macro-view piece (nightly sweep, phase lanes, audit ledger, verify loop) could be
built in this same Manim project as a second scene. This piece is the micro view:
one machine's path through the rules.
