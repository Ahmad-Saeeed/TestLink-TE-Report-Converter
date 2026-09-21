TESTLINK TO WE FORMATTER  -  Windows app
========================================

WHAT IT DOES
------------
Takes a TestLink "Test Specification" export (the .doc file that TestLink
produces) and turns it into a WE / TelecomEgypt-branded Word document:

  * removes the TestLink logo and the "Community [configure ...]" line
  * keeps all the content: suites, test cases, steps, expected results, etc.
  * applies the WE logo + telecomegypt header, gray divider and purple footer
  * names the result from the CR in the Test Suite, e.g.
        "Message Broker13 Cloud - CR 6894 RestartOnFTTH"
        ->  CR 6894-RestartOnFTTH.docx

The output is identical to what the "testlink-wecom-format" Claude skill makes.


USING THE APP  (for whoever receives the finished app)
-------------------------------------------------------
1. Open the TestLinkWEFormatter folder and double-click TestLinkWEFormatter.exe
2. Click the first "Browse..." and choose the TestLink export.
3. (Optional) Change where the Word file is saved. By default it goes in the
   same folder as the export.
4. Click "Convert to WE format". When it says Done, click "Show in folder".

Nothing needs to be installed. Keep the .exe together with the other files in
its folder - the whole folder is the app.


HOW TO BUILD THE APP  (recommended: GitHub Actions - no Windows PC or Python needed)
--------------------------------------------------------------------------------------
This project can't be turned into an .exe on a Mac/Linux machine or inside
Claude - an .exe has to be built on Windows. GitHub can do that for you for
free on one of its own Windows machines:

1. Create a new repository on github.com (private is fine).
2. Upload the CONTENTS of this folder to the repository root (app.spec,
   main.py, requirements.txt, resources, etc. - not the zip file itself).
3. Add the build instructions. Browsers and file explorers often skip folders
   that start with a dot, so the hidden ".github" folder may not upload. To be
   safe, create the file directly on GitHub:
     - In the repo click  Add file -> Create new file
     - In the name box type exactly:  .github/workflows/build-windows-exe.yml
       (typing each "/" creates the folders)
     - Open .github/workflows/build-windows-exe.yml from this project in a
       text editor, copy everything, and paste it into the big box
     - Click "Commit changes"
   Committing that file also starts the first build.
4. Open the repo's "Actions" tab. You should see "Build Windows EXE". Every
   push starts a build; you can also start one yourself with
   "Build Windows EXE" -> "Run workflow".
   If the Actions tab only says "Get started with GitHub Actions", the
   workflow file is not in the repo yet - redo step 3.
   If it says workflows are disabled, go to Settings -> Actions -> General and
   choose "Allow all actions and reusable workflows".
5. Wait a few minutes for the green check mark, click the finished run, and
   under "Artifacts" download "TestLinkWEFormatter-Windows".
6. Unzip it. Inside is the app folder - that is what you hand out.

Notes:
  * Public repos build for free. Private repos get a monthly free allowance.
  * Downloaded artifacts expire after 90 days - just re-run the workflow.
  * If a build fails, click the failed step on the Actions page to read the
    log. A "ModuleNotFoundError" in the log or when starting the app means a
    module name needs adding to `hiddenimports` in app.spec.


ALTERNATIVE: BUILD ON YOUR OWN WINDOWS PC
------------------------------------------
Only if you would rather not use GitHub:
  1. Install Python 3.10 or newer from python.org (tick "Add Python to PATH").
     Python is needed once, to build. People who receive the finished app
     need nothing installed.
  2. Double-click build.bat.
  3. The app appears in  dist\TestLinkWEFormatter\


IF THE TEMPLATE CHANGES
-----------------------
The WE header, footer, logos and styles live in  resources\skeleton\.
If the branding is ever updated, replace those files (from a new reference
.docx, leaving word\document.xml out) and rebuild.


ABOUT ANTIVIRUS WARNINGS
------------------------
Newly built, unsigned apps made with PyInstaller are sometimes flagged by
Windows Defender / SmartScreen or by scanners such as VirusTotal. This is a
false positive: the scanners use heuristics that treat "bundled Python app"
as slightly suspicious because malware also uses PyInstaller.

This project already reduces the chance of it:
  * it builds a folder (not a single self-extracting .exe), which scans much
    cleaner;
  * it does not use UPX compression, which malware often uses;
  * the .exe carries proper version information (company, product name).

If Windows SmartScreen still shows "Windows protected your PC", click
"More info" -> "Run anyway". If a specific antivirus flags it, submit the file
as a false positive to that vendor (Microsoft has a web form for Defender);
this usually clears within days. The only way to remove these warnings
completely is to buy a code-signing certificate and sign the .exe - a
deliberate choice you can make later.
