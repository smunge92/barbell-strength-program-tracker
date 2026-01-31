# Barbell Strength Tracker

Hey there! 👋

I'm a **vibe coder** - not a software developer by trade, just someone who wanted a better way to track my barbell training. This entire project was built using [Claude Code CLI](https://claude.ai/code) (the Max version). Yep, from the first line of code to pushing it to this GitHub repo - all done through conversation with AI. Pretty cool, right? 😊

If I can do it, you can too!

---

## What Is This?

A comprehensive Excel-based workout tracker for novice linear progression barbell training. It's based on Mark Rippetoe's program from "Starting Strength: Basic Barbell Training".

**What it does:**
- Tracks your workouts with auto-calculated target weights
- Detects stalls automatically (OK/STALL status)
- Tells you when you're ready to transition from novice to intermediate
- Logs body weight and shows your progress over time

## Features

- **9 sheets**: README, Settings, Assistance Exercises, Workout Log, Warm-Up Calculator, Body Weight Log, Progress Summary, Progress Chart, Program Phase
- **Auto-progression**: Target weights calculated from your last successful lift
- **Smart exercise introduction**: Light Squat and Chin Ups appear based on training progress
- **Mobile compatible**: Works with Excel on iOS/Android via OneDrive (no macros!)

---

## How to Generate the Tracker

Don't worry, this is easier than it looks! Just follow along step by step. 🙂

### Step 1: Install Node.js (one-time setup)

**What's Node.js?** It's a program that lets your computer run JavaScript code. You need it to generate the Excel file.

1. Go to [nodejs.org](https://nodejs.org/)
2. Click the big green button that says **"LTS"** (this is the stable version)
3. Run the downloaded installer
4. Just keep clicking "Next" and accept the defaults
5. Click "Install" and wait for it to finish

**How do I know it worked?**
- Open Command Prompt (search "cmd" in Windows Start menu)
- Type `node --version` and press Enter
- You should see a version number like `v20.11.0` - that means it worked!

### Step 2: Download this project

**Option A: Download as ZIP (easiest)**
1. On this GitHub page, click the green **"Code"** button near the top
2. Click **"Download ZIP"**
3. Find the downloaded ZIP file and extract it (right-click → "Extract All")
4. Remember where you extracted it!

**Option B: Clone with Git (if you're feeling fancy)**
```
git clone https://github.com/smunge92/barbell-strength-program-tracker.git
```

### Step 3: Open Command Prompt in the project folder

This is where people usually get stuck, so let me explain a few ways to do it:

**Easy way (Windows):**
1. Open the folder where you extracted/downloaded the project
2. Click in the address bar at the top (where it shows the folder path)
3. Type `cmd` and press Enter
4. A black Command Prompt window will open, already in the right folder!

**Alternative way:**
1. Open Command Prompt (search "cmd" in Start menu)
2. Type `cd ` (with a space after it)
3. Drag and drop the project folder into the Command Prompt window
4. Press Enter

You'll know you're in the right place if you see the folder path ending with something like `barbell-strength-tracker`.

### Step 4: Install the dependencies

**What are dependencies?** They're helper tools that the code needs to work. Think of them like ingredients for a recipe.

In your Command Prompt, type:
```
npm install
```
Press Enter and wait. You'll see a bunch of text scrolling - that's normal! It's downloading the ingredients.

**When it's done**, you'll see something like "added 85 packages" and get a fresh prompt line.

### Step 5: Generate your tracker!

This is the fun part. Type:
```
node create-tracker.js
```
Press Enter.

You should see:
```
Barbell Strength Tracker v2.5 created successfully!
File: Barbell_Strength_Tracker.xlsx
```

### Step 6: Open and enjoy! 🎉

Look in the project folder - you'll see a shiny new file called `Barbell_Strength_Tracker.xlsx`.

Double-click it to open in Excel, and you're ready to start tracking your gains!

### Step 7: Sync to OneDrive/SharePoint for mobile access 📱

Want to log your workouts from your phone at the gym? Here's how:

**Why this works great:**
- ✅ No macros - just formulas, so it works everywhere
- ✅ Real-time sync between your computer and phone
- ✅ Works offline too (syncs when you're back online)

**How to set it up:**

1. **Move the file to OneDrive:**
   - Open File Explorer
   - Find `Barbell_Strength_Tracker.xlsx` in your project folder
   - Copy it (Ctrl+C)
   - Go to your OneDrive folder (usually in the left sidebar)
   - Paste it there (Ctrl+V)

2. **Access on your phone:**
   - Download the **Microsoft Excel** app (iOS or Android)
   - Sign in with the same Microsoft account
   - Your tracker will appear in the "Recent" or "Shared" section
   - Tap to open and start logging!

**Pro tip:** Create a home screen shortcut on your phone for quick access at the gym.

---

## Customization

Edit the **Settings** sheet to customize:
- Starting weights for each lift
- Weight increments
- Stall threshold
- Chin Ups introduction week
- Light Squat percentage

## The Program

| Workout A | Workout B |
|-----------|-----------|
| Squat 3x5 | Squat 3x5 |
| Bench Press 3x5 | Overhead Press 3x5 |
| Deadlift 1x5 | Deadlift 1x5 |

**Schedule**: 3 days per week, alternating A-B-A, B-A-B

## All The Sheets

| Sheet | What It Does |
|-------|--------------|
| README | Quick start guide |
| Settings | Configure weights, increments, thresholds |
| Assistance Exercises | Add/remove optional exercises |
| Workout Log | Main tracking - this is where you log your lifts |
| Warm-Up Calculator | Figure out your warm-up sets |
| Body Weight Log | Track weekly weigh-ins |
| Progress Summary | See your weekly bests |
| Progress Chart | Visualize your gains |
| Program Phase | Know when to move to intermediate |

## Smart Exercise Introduction

The tracker is smart enough to introduce exercises at the right time:

- **Light Squat**: Shows up when any lift hits the stall threshold (you're transitioning to intermediate)
  - Uses 80% of your Squat working weight
  - Helps keep squat frequency up during the transition

- **Chin Ups**: Shows up after 2 weeks of training
  - Gives you time to get the main lifts dialed in first

---

## For Developers (or Curious Vibe Coders 🤓)

Want to poke around the code or run the tests? Here's some extra stuff:

### Run the Test Suite

This project includes **169 automated tests** that verify everything works correctly. To run them:

```
node test-tracker.js
```

You'll see a bunch of green checkmarks (hopefully!) and a summary at the end:
```
Total Tests: 169
Passed: 169
Failed: 0
Success Rate: 100.0%
```

**What gets tested:**
- All 9 sheets exist and have correct structure
- All formulas are in place and reference the right cells
- Default settings are correct
- Dropdown menus work
- Stall detection logic
- Exercise introduction timing
- And much more...

### Project Structure

```
barbell-strength-tracker/
├── create-tracker.js    # Main script that generates the Excel file
├── test-tracker.js      # 169 automated tests
├── package.json         # Project config and dependencies
└── README.md            # You're reading it!
```

### How This Was Built

This entire project was built through conversations with **Claude Code CLI** - an AI coding assistant. Here's what that looked like:

1. **Started with a simple idea**: "I want an Excel tracker for Starting Strength"
2. **Iteratively built features**: Each conversation added new capabilities
3. **Refactored and improved**: Claude helped reorganize code and add polish
4. **Added comprehensive tests**: 169 tests to make sure nothing breaks
5. **Prepared for GitHub**: README, .gitignore, even this description!

If you want to modify this project, you can use Claude Code CLI too. Just open the project folder and start chatting:

```
claude
```

Then describe what you want to change. It's like pair programming with an AI! 🤖

### Want to Contribute?

Feel free to:
- Open an issue if you find a bug
- Submit a pull request with improvements
- Fork it and make it your own

---

## Troubleshooting

**"node is not recognized"**
- You need to install Node.js (Step 1) and restart Command Prompt

**"npm is not recognized"**
- Same as above - install Node.js and restart Command Prompt

**"Cannot find module 'exceljs'"**
- You forgot to run `npm install` (Step 4)

**The Excel file won't open**
- Make sure you have Microsoft Excel or a compatible app installed
- Try uploading to OneDrive and opening in Excel Online

**Formulas show errors on mobile**
- Make sure you're using the official Microsoft Excel app
- Google Sheets may not support all formulas

---

## Disclaimer

This is an unofficial community tool. Not affiliated with or endorsed by Starting Strength, Inc. or Aasgaard Company. The program is based on publicly available information from "Starting Strength: Basic Barbell Training" by Mark Rippetoe.

## Direct Download

Don't want to generate the file yourself? Download the latest pre-built Excel file:

**[Barbell_Strength_Tracker_2026-01-31.xlsx](./Barbell_Strength_Tracker_2026-01-31.xlsx)**

Just download, open in Excel, and start tracking!

---

## Changelog

### 2026-01-31
**Bug Fixes & Optimizations**
- Fixed Progress Summary formulas - now uses DMAX with helper columns for reliable date filtering
- Fixed Progress Chart PR formulas - uses simple MAX on helper columns
- Added hidden helper columns (P-T) in Workout Log for each exercise weight
- Resolved #VALUE! errors that occurred with various Excel versions
- Code optimization: removed unused functions, added reusable style helpers
- Reduced code by 55 lines while maintaining full functionality

### Previous (v2.5)
- Smart Exercise Introduction: Light Squat and Chin Ups appear based on training progress
- Session/Week tracking auto-calculated from Workout Log
- Exercise list moved to Assistance Exercises sheet column E
- Target Weight formulas use SUMPRODUCT for Excel 2013+ compatibility
- 169 automated tests ensure reliability

---

## License

MIT License - use it, modify it, share it, whatever you want.

## Built With

- JavaScript + [ExcelJS](https://github.com/exceljs/exceljs)
- A lot of conversations with [Claude Code CLI](https://claude.ai/code) 🤖
- Zero prior coding experience required

---

*Get stronger. Add weight. Repeat.* 💪
