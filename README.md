# Presentation Pacer

A lightweight VBA macro for PowerPoint that shows a segmented countdown clock in the top-right corner during a slide show, so a presenter always knows how much time is left for the current part of the talk.

## What it does

During the slide show (F5), a clock appears in the top-right corner:

- **Green**: on schedule
- **Orange**: less than 30s left in the current segment's time budget
- **Red with "+"**: over budget (shows by how much)

The clock **resets automatically** whenever you move into a new slide range, and counts down the time allotted to that segment only, so the presenter immediately sees how much time remains for their part, regardless of how much time has already elapsed. For example:

- Slides 1–5 → clock starts at **5:00** and counts down
- Slides 6–10 → on entering slide 6, the clock resets and starts at **15:00**

## Files

- `PrezentacjaZegar.bas`: main VBA module (`ModZegar`) with the clock logic and segment definitions
- `ClsPresEvents.cls`: class module (`KlasaZdarzen`) that wires up slide show events

## Setup

1. Enable the Developer tab (File → Options → Customize Ribbon → check "Developer")
2. Open the VBA editor (Developer → Visual Basic, or Alt+F11)
3. Import `PrezentacjaZegar.bas` (adds module `ModZegar`)
4. Edit the `ranges` array in `GetSegmentInfo` to match your agenda:
   ```vba
   ranges = Array( _
       Array(1, 5, 5), _
       Array(6, 10, 15), _
       Array(11, 999, 30) _
   )
   ```
   Each row is `(start slide, end slide, duration of this segment in minutes)`.
5. Import `ClsPresEvents.cls` (adds class `KlasaZdarzen`)
6. Manually run macro `InitEvents` once per session (Alt+F8 → `InitEvents` → Run), before starting the slide show (`Auto_Open` doesn't run in `.pptm` files)
7. Save the file as `.pptm` (macro-enabled presentation)
8. Enable macros on first open, run `InitEvents` again, then start the slide show (F5)

## Notes

- The clock draws an invisible-to-nobody textbox named `ClockBox` only on the slide currently shown during the slide show, so it won't clutter other slides in edit view.
- If you want the clock visible only on the presenter's screen and not the audience's, VBA alone can't split that: a separate Presenter View plus a browser-based tool on your monitor would be needed instead.
- The clock only resets when entering a **different** segment than the one currently counting down. Moving between slides within the same segment doesn't reset it.
- Going back to a previous segment also resets the clock to the full duration of that segment (it doesn't remember how much time was already used).

## Platform notes

`PrezentacjaZegar.bas` is written for **macOS** PowerPoint, where `Application.OnTime` is not available. Instead, `StartClock` runs a polling loop (`Timer` + `DoEvents`) that calls `UpdateClock` roughly once per second, so the clock still ticks every second even without `Application.OnTime`. It also refreshes immediately on slide change via `App_SlideShowNextSlide`.

## License

MIT
