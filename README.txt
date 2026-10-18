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
4. Click "Convert to WE format". When it says Done, click "Open output folder".

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
   (Optional) Repeat with  .github/workflows/build-windows-exe-stock-launcher.yml
   for the fallback workflow described under "About antivirus warnings" below.
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
  * This workflow compiles PyInstaller's own launcher from source before
    building (see "About antivirus warnings"), so it takes a few minutes
    longer than a bare PyInstaller build.


ALTERNATIVE: BUILD ON YOUR OWN WINDOWS PC
------------------------------------------
Only if you would rather not use GitHub:
  1. Install Python 3.10 or newer from python.org (tick "Add Python to PATH").
     Python is needed once, to build. People who receive the finished app
     need nothing installed.
  2. Double-click build.bat.
  3. The app appears in  dist\TestLinkWEFormatter\
This uses PyInstaller's stock launcher (see "About antivirus warnings") since
compiling the launcher from source needs a C compiler most Windows PCs don't
have set up. Prefer GitHub Actions if you can.


IF THE TEMPLATE CHANGES
-----------------------
The WE header, footer, logos and styles live in  resources\skeleton\.
If the branding is ever updated, replace those files (from a new reference
.docx, leaving word\document.xml out), then run  python make_manifest.py  to
refresh resources\skeleton_manifest.json, and rebuild.

Two special parts are stored under plain names so browsers and GitHub uploads
cannot lose them: Content_Types.xml (really "[Content_Types].xml") and
_rels\rels_root.xml (really "_rels\.rels"). The app restores the real names
inside every generated Word file. Do not rename them back.

Before replacing any header/footer image or shape, prefer a plain embedded
picture (like the telecomegypt wordmark and the address footer) over a vector
shape or drawing pasted from another document. A pasted shape can silently
carry legacy VML markup and the originating document's own metadata into the
template - see version 1.0.4 below for what that looked like in practice.


ABOUT ANTIVIRUS WARNINGS
------------------------
Newly built, unsigned apps made with PyInstaller are sometimes flagged by a
few antivirus engines on scanners such as VirusTotal (usually 1-3 out of ~70,
with generic names like "BehavesLike...Backdoor" or "Trojan...!ml"). These are
guesses by heuristic/AI engines, not matches to a real malware family: malware
authors also use PyInstaller, and every PyInstaller app shares the same small
launcher stub. Nobody can promise a scan of 0 detections, but this project
already lowers the odds:

  * the normal build workflow compiles PyInstaller's launcher from source,
    so the .exe doesn't carry the stock launcher fingerprint that some
    antivirus engines have learned to distrust (this made the biggest
    difference in testing - it took a 3/69 result down to 1/69, Microsoft
    Defender included, with no other change);
  * it builds a folder (not a single self-extracting .exe), which scans cleaner;
  * no UPX compression, which malware often uses;
  * the .exe carries version information (company, product name) and an icon;
  * plain Tkinter only - no extra GUI packages;
  * the code makes no network connections and launches no other programs.

If a scan still shows detections, in this order:

  1. Make sure you used "Build Windows EXE" (the from-source-launcher
     workflow), not the stock-launcher fallback. If you already did and a
     detection remains, try the stock-launcher workflow too and compare -
     occasionally one happens to score better than the other.
  2. Report each remaining detection as a false positive to the vendor that
     flagged it, for the FINAL build only. Every rebuild changes the file, so
     submit after you have settled on the version you will distribute.
     Microsoft has a web form for Defender (microsoft.com/wdsi/filesubmission);
     other vendors have their own (e.g. SecureAge/SecureAPlus has a form on
     their knowledge base site).
  3. For a company tool, ask IT to sign the .exe with the company code-signing
     certificate or to allowlist it by file hash. Signing is the only way to
     remove this category of warning reliably.
  4. Before handing it out widely, ask IT/security to scan it with the company
     antivirus, and try the whole folder on a Windows PC or in Windows Sandbox.

If Windows SmartScreen shows "Windows protected your PC", click "More info"
-> "Run anyway".

To use your own icon, replace icon.ico (a multi-size .ico file) and rebuild.


IF WORD SAYS IT "CAN'T OPEN" THE CONVERTED FILE
-----------------------------------------------
The app checks the template files before it builds, and checks the finished
Word file afterwards (structure, XML, links between parts). If anything is
wrong it shows a clear message instead of giving you a broken file - please
send that message to whoever maintains this tool.

If a file still won't open in Word:
  1. In the Word error box click "Details >>>" and note the "Location" line
     (part name, line and column).
  2. Keep the original TestLink export that produced it and the .docx itself -
     with those the problem can be reproduced and fixed quickly.
  3. Check the app version: right-click TestLinkWEFormatter.exe -> Properties ->
     Details -> File version.

Fixed in 1.0.1: control characters pasted into TestLink text, right-to-left
lines at the end of a step cell.
Fixed in 1.0.2: a template file missing from the GitHub upload (files whose
names start with a dot are often skipped) silently produced a Word file that
Word reports as having "a problem with its contents".
1.0.3: adds a small "\u00a9 Ahmed Saeed" credit line in the bottom-right corner of
the window. No change to the conversion itself.
1.0.4: the bundled template's header logo was a legacy vector shape (with a
VML fallback) left over from whoever built the original reference document,
and docProps/custom.xml carried that document's own leftover metadata
(an Adobe InDesign reference, a Grammarly document ID, a SharePoint DocId).
Neither affects how the file looks, but some strict upload validators reject
documents with VML or unusual package metadata. The header logo is now a
plain embedded picture (pixel-identical to before) with no VML, custom.xml is
removed, and the creator/last-modified metadata is generic instead of
personal. Every generated document's created/modified timestamp now reflects
when it was actually built rather than a frozen template date.
