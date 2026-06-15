==========================================
  AIRTABLE DOCUMENT AGENT - INSTALL GUIDE
==========================================

Follow these 7 steps. Do NOT skip any. Stop and ask
for help if anything looks wrong.

------------------------------------------
STEP 1 - Allow PowerShell scripts (one time)
------------------------------------------

  1. Click Start, type "PowerShell"
  2. Right-click "Windows PowerShell" -> "Run as Administrator"
  3. Click YES on the security prompt
  4. Paste this command and press Enter:

       Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force

  5. Close that PowerShell window.

------------------------------------------
STEP 2 - Install the agent
------------------------------------------

  1. Right-click install.ps1 (in this folder)
  2. Click "Run with PowerShell"
  3. Wait until it says "Setup complete"
  4. Press Enter to close

------------------------------------------
STEP 3 - Add your API keys to .env
------------------------------------------

  1. Right-click .env -> Open with -> Notepad
     (If .env is hidden: in File Explorer click View ->
      Show -> File name extensions, then look again.)

  2. Find the line:
       ANTHROPIC_API_KEY=REPLACE_WITH_ANTHROPIC_KEY
     Replace it with the Claude key you were given.

  3. Find the line:
       AIRTABLE_API_KEY=REPLACE_WITH_AIRTABLE_PAT
     Replace it with the Airtable token you were given.

  4. Save (Ctrl+S) and close Notepad.

  Note: TEMPLATES_FOLDER and OUTPUT_FOLDER are already
  set to the right paths for this PC. Don't touch them.

------------------------------------------
STEP 4 - Start the agent
------------------------------------------

  1. Right-click start.ps1 -> "Run with PowerShell"
  2. A black window opens. Leave it open.
     It should say "Ready. Open http://localhost:3000"
  3. If you see warnings about the Z: drive or folders,
     stop and call for help.

------------------------------------------
STEP 5 - One-time Airtable setup
------------------------------------------

  1. Open your web browser
  2. Go to:    http://localhost:3000/setup
  3. The page will:
       - Confirm the Template field exists in Airtable
       - Show 3 buttons you need to add manually
  4. For each button on that page:
       a. Click the "Copy" button
       b. Switch to Airtable
       c. Click the + at the right of your table
       d. Choose "Button"
       e. Name the button (Open Templates, Refresh Templates,
          or Run Agent - matches the page)
       f. Action: "Open URL"
       g. Paste the formula
       h. Click Create
  5. Repeat for all 3 buttons.

------------------------------------------
STEP 6 - Make the agent start automatically
------------------------------------------

  1. Right-click enable-autostart.ps1 -> "Run with PowerShell"
  2. Press Enter to close.

  From now on, the agent boots up by itself when you log in
  to Windows. It runs minimized (you'll see a small PowerShell
  icon in the taskbar).

------------------------------------------
STEP 7 - Try it
------------------------------------------

  1. In Airtable, open any record
  2. Click "Refresh Templates" button (only needed once,
     or whenever you add new template files)
  3. Click "Template" dropdown -> pick a template
  4. Click "Run Agent"
  5. Wait ~30 seconds
  6. The filled doc appears in:
       - C:\Users\harold.freund\Downloads
       - The "Draft Letter and sheet" attachment field

==========================================
  HELP / TROUBLESHOOTING
==========================================

If something fails, send the error to the person
who set this up for you. Don't try to fix it yourself
unless you know what you're doing.

To stop the agent right now:
  - Look for the minimized PowerShell window
  - Close it. The agent stops.

To start it back up:
  - Right-click start.ps1 -> Run with PowerShell

To turn off auto-start at login:
  - Right-click disable-autostart.ps1 -> Run with PowerShell
