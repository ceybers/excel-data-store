# History behind Excel Data Store
## 🤔 Scenario
At work, we have a lot of Excel spreadsheets. Many, many spreadsheets.

Some of the spreadsheets are linked in some way or the other. Some spreadsheets overlap and contain the same information. Some sheets pull data from other sheets, which in turn might pull _different_ day back in return.

And most importantly, the data is always changing. There are always updates: revised estimates, changes in scope, feedback on project progress, running commentary.

Its never as simple as changing one value in one cell, when that same cell needs to be updated in half a dozen other spreadsheets to make sure everything stays balanced. When a project milestone is reached, no matter where that data is entered, it might need to propagate to ten other spreadsheets that need that update.

In a perfectly organised world, everything would be sitting in a database with a strictly enforced schema. Updates would always be exact, precise, and entered slowly and carefully. There would be one and exactly one place to find a specific piece of information.

But the real world doesn't operate like that. There's never enough time for everything, and we don't always have the luxury of documenting what changed, when it changed, and it changed. The updates need to happen and they need to happen _now_. If, in a month from now, we need to go back and understand what changed, then that's a problem for then and not for now.

## 🗃️ Background
This project is an attempt to have some sort of comprehensive system that can record and track data across multiple spreadsheets. 

It needs to be generalised enough to handle as many domains of data that we can throw at it, without needing exceptions for every scenario. Some of the previous attempts worked, but only for the initial niche they were designed for, and could not scale at all. They were also rigid and couldn't handle data that wasn't in the exact correct shape.

This attempt applies the learnings of the previous systems and hopefully addresses the shortcomings.

The project takes a lot of inspiration from Git, with early proofs of concept trying to emulate the basics of how it works. As development progressed, a lot of this changed to a system more suitable to working with Excel and spreadsheets. Lots of the moving parts such as Pulls, Pushes and Commits have kept the same name, but what they actually end up doing now is different. 

---
⏏️ [Back to README](../README.md)