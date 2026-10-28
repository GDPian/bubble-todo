# Getting iPhone steps into Bubbles

A web app cannot read Apple Health, so the iPhone sends the step total itself using the Shortcuts app.
The path is: **Health -> Shortcut -> your Apps Script link -> "Steps" tab in the Hevy sheet -> Bubbles**.

## 1. Sheet side (already built)
`bubbles_api.gs` has two links (same Web app URL and token as the others):

- save: `<URL>?what=steps_set&date=2026-10-28&steps=8123&token=<token>`  (a day's total can only go up, so a late or lower reading never overwrites a higher one)
- read: `<URL>?what=steps&token=<token>` -> paste this as the Sync URL on the **10k steps** habit in Bubbles

The habit ticks itself on any day at or above `BUBBLES_STEP_GOAL` (8,000 in the script, because an iPhone-only count runs about 10% under the Health app) and the bubble shows today's running total (for example 7.4k).

## 2. The Shortcut (Shortcuts app on the iPhone)
Create a new shortcut called **Send steps** with these actions, in order:

1. **Find Health Samples**: Type = Steps, Start Date = is Today. (Leave sorting and limit alone.)
2. **Calculate Statistics**: Sum, of *Health Samples*.
3. **Round Number**: the result, to the ones place.
4. **Format Date**: Current Date, Format = Custom, `yyyy-MM-dd`.
5. **URL**: `<Web app URL>?what=steps_set&date=` *[Formatted Date]* `&steps=` *[Rounded Number]* `&token=<token>`
6. **Get Contents of URL**: that URL, method GET.

The first run asks for permission to read Health Steps and to contact script.google.com. Allow both.
Check it: run the shortcut, then open `<URL>?what=steps&token=<token>`. You should see today's steps under `latest`.

## 3. Make it run by itself
Shortcuts > Automation > + > choose a trigger, add action **Run Shortcut = Send steps**, set **Run Immediately**, turn **Notify When Run** off.

- **Time of Day** (for example 12:30, 17:30, 21:30): simplest, but iOS often keeps Health data locked while the phone is locked,
  so the shortcut can silently get nothing. If a time-of-day run does not update the sheet, this is the usual reason.
- **Open App** (for example whenever you open Safari or a messaging app): the phone is unlocked by definition, so Health is readable.
  This is the more reliable choice. Pick an app you open several times a day.
- **Charger** (when it connects to power): good for an end-of-day total.
- Add the shortcut to your Home Screen as a manual tap-to-sync fallback.

## 4. If this does not work for you
- **Health Auto Export** (paid iOS app) can send Health data to a web address on a schedule in the background. Its JSON would need a `doPost` in the script.
- **Withings**: only works if your steps are recorded by a Withings device and the Withings app permission includes activity data. The current sheet sync only asks for scale measurements.
