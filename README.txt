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
   (Optional) Repeat with  .github/workflows/build-windows-exe-custom-bootloader.yml
   for the second, manual-only workflow described under "About antivirus warnings".
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
Newly built, unsigned apps made with PyInstaller are sometimes flagged by a
few antivirus engines on scanners such as VirusTotal (usually 1-3 out of ~70,
with generic names like "BehavesLike...Backdoor" or "Trojan...!ml"). These are
guesses by heuristic/AI engines, not matches to a real malware family: malware
authors also use PyInstaller, and every PyInstaller app shares the same small
launcher stub. Nobody can promise a scan of 0 detections, but this project
already lowers the odds:

  * it builds a folder (not a single self-extracting .exe), which scans cleaner;
  * no UPX compression, which malware often uses;
  * the .exe carries version information (company, product name) and an icon;
  * plain Tkinter only - no extra GUI packages;
  * the code makes no network connections and launches no other programs.

If a scan still shows detections, in this order:

  1. Run the second workflow, "Build Windows EXE (custom launcher)"
     (Actions tab -> Build Windows EXE (custom launcher) -> Run workflow).
     It compiles PyInstaller's launcher from source first, so the .exe no
     longer carries the stock launcher fingerprint. Scan that artifact and
     compare it with the normal one. (If this workflow fails, the log shows
     why; the normal workflow still works on its own.)
  2. Report each remaining detection as a false positive to the vendor that
     flagged it, for the FINAL build only. Every rebuild changes the file, so
     submit after you have settled on the version you will distribute.
     Microsoft has a web form for Defender; other vendors have their own.
  3. For a company tool, ask IT to sign the .exe with the company code-signing
     certificate or to allowlist it by file hash. Signing is the only way to
     remove this category of warning reliably.
  4. Before handing it out widely, ask IT/security to scan it with the company
     antivirus, and try the whole folder on a Windows PC or in Windows Sandbox.

If Windows SmartScreen shows "Windows protected your PC", click "More info"
-> "Run anyway".

To use your own icon, replace icon.ico (a multi-size .ico file) and rebuild.
