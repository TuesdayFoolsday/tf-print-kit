# Flow 3: adding the 5th data part (owner steps)

The nightly data no longer fits in four 18 KB files, so it now goes up as **five**:
`Dashboard Data 1.txt` … `Dashboard Data 5.txt`. Until the steps below are done, the workbook will NOT update.

1. **Data 5 file.** The nightly courier run creates `Dashboard Data 5.txt` by itself, in the same SharePoint
   folder as Data 1–4, the first time it uploads. If you want to set up Flow 3 *before* that run, create an empty
   text file with exactly that name in that folder yourself (the courier will then just overwrite it).
2. **Office Script.** In Excel, open the Office Scripts `load_dashboard_data` script and replace all of it with
   the current `dashboard/load_dashboard_data.ts` from this repo. (It is not copied into Excel automatically.)
3. **Power Automate, Flow 3.** Add one more **Get file content** step for `Dashboard Data 5.txt`. In the
   **Run script** step, fill the new `part5` box with that step's file content. Leave `part1` … `part4` as they are.
4. **Timing.** Do all of this **before 9:00pm CT on the day the pull requests are merged**. The courier runs at
   7:20pm CT and uploads 5 parts; Flow 3 loads at 9:00pm CT. If Flow 3 still passes only 4 parts, the script
   refuses with `Got 4 data parts, expected 5` and the workbook stays on old data.
5. **Check the next morning.** Open `Dashboard.xlsx`: the date in cell A2 should be that night's date (not Nov 10),
   and Bid Log / estimates should look complete (the Nov 11–12 builds had 224 bids / 42 estimates). In the Flow 3
   run history, the Run script step should show `Loaded: …` listing every table.

**Merge order:** merge the print-kit PR first, then the courier PR, then immediately do steps 2–3 (and 1 if you
create the file yourself).
