# How to Upload This Project to GitHub

A step-by-step guide for beginners who have never used Git or GitHub.

---

## What You're About to Do

You will:
1. Create a free GitHub account (if you don't have one)
2. Install Git on your computer
3. Create a new repository (a "folder" on GitHub for your project)
4. Upload your code to that repository
5. Make it public so others can use it

---

## Part 1: Create a GitHub Account

**What is GitHub?**
GitHub is a website where developers store and share code. Think of it like Google Drive, but specifically designed for code projects.

### Steps:

1. Open your web browser and go to: **https://github.com**

2. Click the **"Sign up"** button in the top right corner

3. Follow the prompts to create an account:
   - Enter your email address
   - Create a password
   - Choose a username (this will be visible to others, e.g., `johndoe123`)
   - Complete the verification puzzle
   - Click **"Create account"**

4. Verify your email address by clicking the link GitHub sends you

5. You now have a GitHub account!

---

## Part 2: Install Git on Your Computer

**What is Git?**
Git is a program that runs on your computer and tracks changes to your code. It's what allows you to upload code to GitHub.

### Steps for Windows:

1. Go to: **https://git-scm.com/download/win**

2. The download should start automatically. If not, click the link for your system (likely "64-bit Git for Windows Setup")

3. Run the downloaded installer (`Git-2.xx.x-64-bit.exe`)

4. During installation, use these settings (just click "Next" for most screens):
   - **Select Components**: Keep defaults, click Next
   - **Default editor**: Keep "Use Vim" or select "Use Notepad" if you prefer
   - **Initial branch name**: Select "Override" and type `main`
   - **PATH environment**: Select "Git from the command line and also from 3rd-party software"
   - **HTTPS transport**: Keep "Use the OpenSSL library"
   - **Line ending**: Keep "Checkout Windows-style, commit Unix-style"
   - **Terminal emulator**: Keep "Use MinTTY"
   - **Default behavior of git pull**: Keep "Fast-forward or merge"
   - **Credential helper**: Keep "Git Credential Manager"
   - **Extra options**: Keep defaults
   - Click **Install**

5. Click **Finish** when done

6. **Verify installation**: Open a new Command Prompt (search "cmd" in Windows) and type:
   ```
   git --version
   ```
   You should see something like: `git version 2.43.0.windows.1`

---

## Part 3: Configure Git with Your Identity

**Why?**
Git needs to know who you are so it can label your contributions.

### Steps:

1. Open **Command Prompt** (search "cmd" in Windows Start menu)

2. Type this command, replacing with YOUR name (keep the quotes):
   ```
   git config --global user.name "Your Name"
   ```
   Example: `git config --global user.name "John Doe"`

3. Press **Enter**

4. Type this command, replacing with YOUR email (the one you used for GitHub):
   ```
   git config --global user.email "your.email@example.com"
   ```
   Example: `git config --global user.email "johndoe@gmail.com"`

5. Press **Enter**

---

## Part 4: Create a Repository on GitHub

**What is a repository?**
A repository (or "repo") is like a project folder on GitHub. It will contain all your code files.

### Steps:

1. Go to **https://github.com** and log in

2. Click the **"+"** icon in the top right corner, then click **"New repository"**

3. Fill in the details:
   - **Repository name**: `barbell-strength-tracker`
   - **Description**: `Excel-based workout tracker for novice linear progression barbell training. Based on Mark Rippetoe's Starting Strength program.`
   - **Public/Private**: Select **Public** (so others can see and use it)
   - **Add a README file**: Leave this **UNCHECKED** (we'll add our own)
   - **Add .gitignore**: Leave as **None**
   - **Choose a license**: Select **MIT License** (this allows others to freely use your code)

4. Click the green **"Create repository"** button

5. You'll see a page with setup instructions. **Keep this page open** - you'll need the URL shown (something like `https://github.com/yourusername/barbell-strength-tracker.git`)

---

## Part 5: Prepare Your Project Files

Before uploading, we need to create a few helper files.

### Step 5.1: Create a README.md file

This is the "welcome page" people see when they visit your repository.

1. Open **Notepad** (or any text editor)

2. Copy and paste this content:

```markdown
# Barbell Strength Tracker

A comprehensive Excel-based workout tracker for novice linear progression barbell training.

## About

This tracker is based on Mark Rippetoe's novice linear progression program from "Starting Strength: Basic Barbell Training". It helps you:

- Track your workouts with auto-calculated target weights
- Detect stalls automatically (OK/STALL status)
- Monitor your transition from novice to intermediate
- Log body weight and visualize progress

## Features

- **9 sheets**: README, Settings, Assistance Exercises, Workout Log, Warm-Up Calculator, Body Weight Log, Progress Summary, Progress Chart, Program Phase
- **Auto-progression**: Target weights calculated from your last successful lift
- **Smart exercise introduction**: Light Squat and Chin Ups appear based on training progress
- **Mobile compatible**: Works with Excel on iOS/Android via OneDrive

## How to Generate the Tracker

### Prerequisites

- [Node.js](https://nodejs.org/) (version 14 or higher)

### Steps

1. Clone or download this repository
2. Open a terminal/command prompt in the project folder
3. Install dependencies:
   ```
   npm install
   ```
4. Generate the Excel file:
   ```
   node create-tracker.js
   ```
5. Open `Barbell_Strength_Tracker.xlsx` in Excel

## Customization

Edit the **Settings** sheet to customize:
- Starting weights for each lift
- Weight increments
- Stall threshold
- Chin Ups introduction week
- Light Squat percentage

## Program Overview

**Workout A**: Squat 3x5, Bench Press 3x5, Deadlift 1x5
**Workout B**: Squat 3x5, Overhead Press 3x5, Deadlift 1x5

Schedule: 3 days per week, alternating A-B-A, B-A-B

## Disclaimer

This is an unofficial community tool, not affiliated with or endorsed by Starting Strength, Inc. or Aasgaard Company. The program structure is based on publicly available information from "Starting Strength: Basic Barbell Training" by Mark Rippetoe.

## License

MIT License - feel free to use, modify, and share.
```

3. Save the file as `README.md` in your project folder:
   `C:\Users\munge\Claude Projects\Starting Strength Tracker\README.md`

   **Important**: Make sure to save as "All Files" type, not "Text Documents"

### Step 5.2: Create a .gitignore file

This tells Git which files NOT to upload (like temporary files).

1. Open **Notepad**

2. Copy and paste this content:

```
# Dependencies
node_modules/

# Generated Excel files (users should generate their own)
*.xlsx

# OS files
.DS_Store
Thumbs.db

# Editor files
*.swp
*.swo
.idea/
.vscode/

# Logs
*.log
npm-debug.log*
```

3. Save as `.gitignore` in your project folder

   **Important**: The filename starts with a dot. Save as "All Files" type.

---

## Part 6: Upload Your Code to GitHub

Now we'll use Git commands to upload everything.

### Steps:

1. Open **Command Prompt**

2. Navigate to your project folder by typing:
   ```
   cd "C:\Users\munge\Claude Projects\Starting Strength Tracker"
   ```
   Press **Enter**

3. **Initialize Git** (tells Git to start tracking this folder):
   ```
   git init
   ```
   *What this does: Creates a hidden `.git` folder that Git uses to track changes*

4. **Set the branch name to "main"**:
   ```
   git branch -M main
   ```
   *What this does: Names your main branch "main" (the standard name)*

5. **Connect to your GitHub repository** (replace `yourusername` with your actual GitHub username):
   ```
   git remote add origin https://github.com/yourusername/barbell-strength-tracker.git
   ```
   *What this does: Links your local folder to the GitHub repository you created*

6. **Stage all files for upload**:
   ```
   git add .
   ```
   *What this does: Tells Git "I want to include all these files in my next upload"*

7. **Create a commit** (a snapshot of your files with a description):
   ```
   git commit -m "Initial release: Barbell Strength Tracker v2.5"
   ```
   *What this does: Packages your files with a message describing what you're uploading*

8. **Push (upload) to GitHub**:
   ```
   git push -u origin main
   ```
   *What this does: Uploads your committed files to GitHub*

9. **If prompted for credentials**:
   - A browser window may open asking you to log in to GitHub
   - Or you may be asked for username/password in the command prompt
   - Enter your GitHub username and password (or personal access token)

---

## Part 7: Verify Your Upload

1. Go to your repository page:
   `https://github.com/yourusername/barbell-strength-tracker`

2. You should see:
   - Your files listed (create-tracker.js, test-tracker.js, package.json, etc.)
   - Your README.md displayed beautifully below the file list
   - A green "MIT License" badge

3. **Congratulations!** Your project is now on GitHub!

---

## Sharing Your Project

To share with others, give them this URL:
```
https://github.com/yourusername/barbell-strength-tracker
```

They can:
- **View the code** directly on GitHub
- **Download it** by clicking the green "Code" button → "Download ZIP"
- **Clone it** using Git (for developers)

---

## Quick Reference: Commands Used

| Command | What It Does |
|---------|--------------|
| `git init` | Start tracking a folder with Git |
| `git add .` | Stage all files for commit |
| `git commit -m "message"` | Create a snapshot with a description |
| `git push` | Upload to GitHub |
| `git status` | See what files have changed |

---

## Troubleshooting

### "git is not recognized"
- Restart Command Prompt after installing Git
- Or reinstall Git and make sure to select "Git from the command line"

### "Permission denied" or authentication error
- Go to GitHub → Settings → Developer settings → Personal access tokens
- Generate a new token with "repo" permissions
- Use this token as your password when prompted

### "Repository not found"
- Double-check the URL matches your repository exactly
- Make sure you created the repository on GitHub first

---

## Need Help?

- GitHub's official guide: https://docs.github.com/en/get-started
- Git basics: https://git-scm.com/book/en/v2/Getting-Started-Git-Basics

