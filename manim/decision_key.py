"""One machine, a few questions: the CCAT decision key, animated + narrated.

Narration: Microsoft Edge TTS (en-US-EmmaNeural), pre-generated into audio/seg1..5.mp3
    uv run --with edge-tts edge-tts --voice en-US-EmmaNeural --text "..." --write-media audio/segN.mp3
    (swap the voice for en-US-AndrewNeural and regenerate to change narrator)

Render:  uv run --with manim manim -qh --disable_caching decision_key.py DecisionKey
Output side effects: media/videos/decision_key/1080p60/DecisionKey.mp4 + .srt subtitles
"""

from manim import (
    DOWN, LEFT, RIGHT, UP,
    Arrow, Create, FadeIn, FadeOut, Flash, Indicate, RoundedRectangle,
    Scene, SurroundingRectangle, Text, VGroup, Write, config,
)

config.background_color = "#101c2c"

INK = "#e8eef5"
DIM = "#9fb3c8"
BLUE = "#4da3ff"
TEAL = "#2ec4b6"
GREEN = "#34c759"
RED = "#ff5a5f"
YELLOW = "#ffc72c"

# Measured durations of the narration segments (seconds)
NARR = {
    1: ("Every machine we track leads a double life: one record in our system, "
        "and one in Caterpillar's registry. Keeping the two in sync starts with "
        "a few simple questions.", 10.32),
    2: ("First: do we already have a record in Caterpillar's system? If not, does "
        "anyone? And if another dealer does, do they actively own it? Every answer "
        "leads to a different action, with a different level of care.", 12.62),
    3: ("Take this brand new excavator. We have no record for it, and neither does "
        "anyone else. So it lands in the safest bucket: a simple add. Nobody else "
        "is affected, the action is logged, and the next nightly sweep verifies "
        "that it landed.", 15.53),
    4: ("But sometimes another dealer holds a record without owning the machine. "
        "This dozer passed through another dealer's rental fleet years ago, and "
        "that old record is still there. Records like that are not exclusive, so "
        "our add still completes quietly. No transfer request fires, and nobody "
        "gets notified. It is logged, all the same.", 19.70),
    5: ("Now a harder case. This machine is ours in the field, but in Caterpillar's "
        "registry, another dealer actively owns it. Claiming it sends that dealer "
        "a formal transfer request. So it goes to a person first, with the evidence, "
        "at about ten reviewed decisions per shift.", 16.34),
    6: ("And that is the whole idea. The system sorts every machine. People decide "
        "only the ones that need judgment.", 7.01),
}


def box(text: str, color: str, width: float = 4.4, size: int = 28) -> VGroup:
    label = Text(text, font_size=size, color=INK, line_spacing=0.8)
    rect = RoundedRectangle(
        corner_radius=0.18, width=width,
        height=max(1.0, label.height + 0.5), color=color, stroke_width=3,
    )
    label.move_to(rect.get_center())
    return VGroup(rect, label)


def chip(serial: str, color: str) -> VGroup:
    label = Text(serial, font_size=22, color="#101c2c", weight="BOLD")
    rect = RoundedRectangle(
        corner_radius=0.12, width=label.width + 0.5, height=0.5,
        color=color, fill_color=color, fill_opacity=1, stroke_width=0,
    )
    label.move_to(rect.get_center())
    return VGroup(rect, label)


def vedge(a: VGroup, b: VGroup, text: str, color: str, side=LEFT) -> VGroup:
    arrow = Arrow(
        a.get_bottom(), b.get_top(), buff=0.15, color=color,
        stroke_width=4, max_tip_length_to_length_ratio=0.10,
    )
    label = Text(text, font_size=24, color=color, weight="BOLD")
    label.next_to(arrow.get_center(), side, buff=0.22)
    return VGroup(arrow, label)


def hedge(a: VGroup, b: VGroup, text: str, color: str) -> VGroup:
    arrow = Arrow(
        a.get_right(), b.get_left(), buff=0.15, color=color,
        stroke_width=4, max_tip_length_to_length_ratio=0.18,
    )
    label = Text(text, font_size=24, color=color, weight="BOLD")
    label.next_to(arrow.get_center(), UP, buff=0.18)
    return VGroup(arrow, label)


class DecisionKey(Scene):
    def narrate(self, n: int) -> float:
        text, dur = NARR[n]
        self.add_sound(f"audio/seg{n}.mp3")
        self.add_subcaption(text, duration=dur)
        return dur

    def travel(self, mobile: VGroup, target: VGroup) -> None:
        self.play(mobile.animate.next_to(target, UP, buff=0.1), run_time=0.9)
        self.play(Indicate(target, scale_factor=1.04), run_time=0.7)

    def construct(self) -> None:
        # ---------- beat 1: title (10.32s) ----------
        d = self.narrate(1)
        title = Text("One machine. A few questions.", font_size=52, color=INK, weight="BOLD")
        sub = Text("How the system decides what is safe to do", font_size=30, color=DIM)
        sub.next_to(title, DOWN, buff=0.35)
        self.play(Write(title), FadeIn(sub, shift=UP * 0.3), run_time=1.2)
        self.wait(d - 2.4 + 0.3)
        self.play(FadeOut(title), FadeOut(sub), run_time=1.2)

        # ---------- beat 2: the tree draws itself (12.62s) ----------
        d = self.narrate(2)
        q1 = box("Do WE have a\nCCAT record for it?", BLUE).move_to(UP * 2.55)
        q2 = box("Does ANY dealer\nhold a record?", BLUE).move_to(LEFT * 2.6 + UP * 0.45)
        q3 = box("Do they actively\nOWN it?", BLUE, width=4.2).move_to(RIGHT * 3.8 + UP * 0.45)
        t_add = box("ADD to CCAT\nsafe -- affects nobody", GREEN, width=4.0, size=26).move_to(LEFT * 5.0 + DOWN * 2.45)
        t_quiet = box("Add quietly\nno dealer notified", TEAL, width=3.5, size=26).move_to(LEFT * 0.4 + DOWN * 2.45)
        t_transfer = box("TRANSFER request\na person reviews first", RED, width=4.4, size=26).move_to(RIGHT * 4.3 + DOWN * 2.45)

        e1 = vedge(q1, q2, "NO", DIM, side=LEFT)
        e2a = vedge(q2, t_add, "NO", GREEN, side=LEFT)
        e2b = hedge(q2, q3, "YES", DIM)
        e3a = vedge(q3, t_quiet, "NO", TEAL, side=LEFT)
        e3b = vedge(q3, t_transfer, "YES", RED, side=RIGHT)

        self.play(FadeIn(q1, shift=DOWN * 0.3), run_time=1.4)
        self.play(Create(e1[0]), FadeIn(e1[1]), FadeIn(q2, shift=DOWN * 0.3), run_time=1.6)
        self.play(
            Create(e2a[0]), FadeIn(e2a[1]), FadeIn(t_add, shift=DOWN * 0.2),
            Create(e2b[0]), FadeIn(e2b[1]), FadeIn(q3, shift=RIGHT * 0.2),
            run_time=1.8,
        )
        self.play(
            Create(e3a[0]), FadeIn(e3a[1]), FadeIn(t_quiet, shift=DOWN * 0.2),
            Create(e3b[0]), FadeIn(e3b[1]), FadeIn(t_transfer, shift=DOWN * 0.2),
            run_time=1.8,
        )
        self.wait(d - 6.6 + 0.3)

        # ---------- beat 3: machine A, the safe add (15.53s) ----------
        d = self.narrate(3)
        a = chip("WKX53704", GREEN).to_corner(UP + LEFT).shift(DOWN * 0.1 + RIGHT * 0.1)
        a_tag = Text("new excavator, not in CCAT", font_size=22, color=DIM)
        a_tag.next_to(a, DOWN, buff=0.12).align_to(a, LEFT)
        self.play(FadeIn(a, shift=RIGHT * 0.4), FadeIn(a_tag), run_time=0.8)
        self.wait(1.6)
        self.travel(a, q1)
        self.play(FadeOut(a_tag), run_time=0.3)
        self.wait(0.8)
        self.travel(a, q2)
        self.wait(0.8)
        self.play(a.animate.next_to(t_add, UP, buff=0.1), run_time=0.9)
        glowa = SurroundingRectangle(t_add, color=GREEN, buff=0.1, stroke_width=5)
        stamp_a = Text("audited + verified by next sweep", font_size=22, color=YELLOW)
        stamp_a.next_to(t_add, DOWN, buff=0.18).align_to(t_add, LEFT)
        self.play(Create(glowa), FadeIn(stamp_a), Flash(t_add.get_center(), color=GREEN, line_length=0.35), run_time=1.0)
        self.wait(d - 9.8 + 0.3)

        # ---------- beat 4: machine C, the quiet add (19.70s) ----------
        d = self.narrate(4)
        c = chip("6GK02175", TEAL).to_corner(UP + LEFT).shift(DOWN * 0.1 + RIGHT * 0.1)
        c_tag = Text("dozer -- old rental record elsewhere", font_size=22, color=DIM)
        c_tag.next_to(c, DOWN, buff=0.12).align_to(c, LEFT)
        self.play(FadeIn(c, shift=DOWN * 0.3), FadeIn(c_tag), run_time=0.8)
        self.wait(1.6)
        self.travel(c, q1)
        self.play(FadeOut(c_tag), run_time=0.3)
        self.wait(0.6)
        self.travel(c, q2)
        self.wait(0.6)
        self.travel(c, q3)
        self.wait(0.6)
        self.play(c.animate.next_to(t_quiet, UP, buff=0.1), run_time=0.9)
        glowc = SurroundingRectangle(t_quiet, color=TEAL, buff=0.1, stroke_width=5)
        stamp_c = Text("no transfer fires -- still audited", font_size=22, color=YELLOW)
        stamp_c.next_to(t_quiet, DOWN, buff=0.18)
        self.play(Create(glowc), FadeIn(stamp_c), Flash(t_quiet.get_center(), color=TEAL, line_length=0.35), run_time=1.0)
        self.wait(d - 11.2 + 0.3)

        # ---------- beat 5: machine B, the reviewed transfer (16.34s) ----------
        d = self.narrate(5)
        b = chip("9303", RED).to_corner(UP + RIGHT).shift(DOWN * 0.1 + LEFT * 0.1)
        b_tag = Text("ours in the field,\ntheirs in CCAT", font_size=22, color=DIM, line_spacing=0.8)
        b_tag.next_to(b, DOWN, buff=0.12).align_to(b, RIGHT)
        self.play(FadeIn(b, shift=LEFT * 0.4), FadeIn(b_tag), run_time=0.8)
        self.wait(1.6)
        self.travel(b, q1)
        self.play(FadeOut(b_tag), run_time=0.3)
        self.wait(0.6)
        self.travel(b, q2)
        self.wait(0.6)
        self.travel(b, q3)
        self.wait(0.6)
        self.play(b.animate.next_to(t_transfer, UP, buff=0.1), run_time=0.9)
        glowb = SurroundingRectangle(t_transfer, color=RED, buff=0.1, stroke_width=5)
        stamp_b = Text("evidence checked -- about 10 per shift", font_size=22, color=YELLOW)
        stamp_b.next_to(t_transfer, DOWN, buff=0.18)
        self.play(Create(glowb), FadeIn(stamp_b), Flash(t_transfer.get_center(), color=RED, line_length=0.35), run_time=1.0)
        self.wait(d - 12.2 + 0.3)

        # ---------- beat 6: closer (7.01s) ----------
        d = self.narrate(6)
        everything = VGroup(
            q1, q2, q3, t_add, t_quiet, t_transfer,
            e1, e2a, e2b, e3a, e3b, a, b, c, glowa, glowb, glowc,
            stamp_a, stamp_b, stamp_c,
        )
        self.play(everything.animate.set_opacity(0.15), run_time=0.8)
        closer1 = Text("The system sorts every machine.", font_size=46, color=INK, weight="BOLD").move_to(UP * 0.5)
        closer2 = Text("People decide only the ones that need judgment.", font_size=38, color=YELLOW)
        closer2.next_to(closer1, DOWN, buff=0.45)
        self.play(Write(closer1), run_time=1.2)
        self.play(FadeIn(closer2, shift=UP * 0.3), run_time=1.0)
        self.wait(d - 3.0 + 1.5)
