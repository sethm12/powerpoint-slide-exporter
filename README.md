# PowerPoint Slide Exporter
**PowerPoint Add-In | Automates high-resolution slide export for video production**

---

## Demo

[![Graph Exporter Demo](https://img.youtube.com/vi/phdhUYuQIuI/maxresdefault.jpg)](https://www.youtube.com/watch?v=phdhUYuQIuI)

> *Click the thumbnail to watch the full demo on YouTube.*

---

## Overview

Slide Exporter is a custom PowerPoint add-in built for a video production team that regularly exports presentation slides as images for use in Adobe Premiere Pro. It solves two real, recurring pain points: PowerPoint's inadequate default export resolution, and duplicate file name conflicts that were breaking Premiere Pro projects.

Both issues previously required either manual IT intervention on every workstation or tedious file renaming by hand. This tool eliminates both problems entirely with a guided export workflow that takes seconds.

---

## The Problems It Solves

**Problem 1 — Export resolution:** PowerPoint exports images at 72 DPI by default. Video production requires 300 DPI. The only native fix is a registry edit that had to be performed manually on every workstation — and repeated any time a new editor joined the team.

**Problem 2 — Duplicate file names:** When pulling slides from multiple previous projects into a single new video, Premiere Pro would encounter multiple files named identically (e.g. three different "Slide 1" exports from three different projects), causing project errors and requiring manual file renaming to resolve.

---

## Features

- **300 DPI export** — exports slides at proper video production resolution without any registry edits or IT involvement, on any workstation
- **Custom Ribbon integration** — surfaces as a native-feeling PowerPoint tool with dedicated Ribbon buttons, no macro menus required
- **Flexible export scope** — export all slides in the presentation, or just the currently selected slides; intelligently handles single-slide, multi-slide, and fallback-to-active-slide scenarios
- **Smart default save location** — automatically opens the folder dialog at the presentation's current save location, eliminating manual navigation every time
- **Unique project identifier prefix** — prompts for a custom ID to prepend to every exported file name (e.g. "A220_Slide1.jpg"), eliminating duplicate file name conflicts in Premiere Pro; auto-capitalizes input and strips illegal filename characters automatically
- **Overwrite with reporting** — forcefully overwrites existing exports when re-exporting, and displays a clear breakdown of how many files were new vs. overwritten so the user always knows exactly what happened
- **Cross-platform** — runs on both Windows and macOS; includes platform-specific engineering under the hood to handle differences in folder picker dialogs, file path formats, and compiler behavior between the two operating systems

---

## Impact

Built collaboratively based on direct feedback from the production team. It eliminated a recurring IT burden (the per-workstation registry edit), removed a manual file management step from every export cycle, and gave the team a reliable, repeatable workflow for a task they perform constantly throughout every project.

---

## Built With

- Microsoft PowerPoint (Office 2016+, Windows & macOS)
- VBA (Visual Basic for Applications)
- Office RibbonX Editor (custom Ribbon UI)
- AppleScript (macOS folder picker integration)

---

## Notes

This repository contains documentation and a demo only. Source code is proprietary and not included.
