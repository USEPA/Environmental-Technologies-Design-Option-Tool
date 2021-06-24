
StEPP README.TXT file.
======================

VERSION HISTORY.
================

---------------- Version 1.0.0 (23-Oct-1998) -------------------------------
- First version compiled with Microsoft Visual Basic 5.0.

---------------- Version 1.0.1 (23-Oct-1998) -------------------------------
- This is the first version where every program function works properly.

---------------- Version 1.0.2 (10-Nov-1998) -------------------------------
- First version compiled with Microsoft Visual Basic 6.0.

---------------- Version 1.0.3 (08-Feb-1999) -------------------------------
(08-Feb-1998)
- The STEPP.DLL file was updated to correct a problem with the
  Polson method for estimating liquid diffusivity.  The revised 
  STEPP.DLL file is dated 08-Feb-1999 and replaces the previous
  version of the file dated 22-Oct-1998.
      As a numerical check, call up Trichloroethylene (CAS#=79016)
  for a temperature of 10 degC.  The correct value of the Polson
  estimation for liquid diffusivity is 5.39e-10 m^2/s; the incorrect
  value previously reported was 5.39e-6 m^2/s.

---------------- Version 1.0.4 (22-Feb-1999) -------------------------------
(22-Feb-1998)
- Updated names on startup window.

---------------- Version 1.0.5 (05-Mar-1999) -------------------------------
(05-Mar-1999)
- Updated MTU logo on startup window.

---------------- Version 01.00.06 (25-Mar-1999) ------------------------------------
(25-Mar-1999)
- Previously, an error in the printing routine would cause the software to
  enter an infinite loop where the only solution was to hit Alt-Ctrl-Delete
  and force the program to end.  To replicate in the older version, select
  the compound "56553  BENZ(A)ANTHRACENE", select File--Print from the
  menu, select Text File, All Properties, All of them, Print Full
  Description of Properties, and Print Values in SI Units.  Then click
  on the button marked Print.  The program then enters the infinite loop.
      The new version does not suffer from this problem.

---------------- Version 01.00.07 (12-Apr-1999)-------------------------------------
(12-Apr-1999)
- Rephrased and moved disclaimer to the install software.

---------------- Version 01.00.08 (14-Apr-1999)-------------------------------------
(14-Apr-1999)
- Fixed bug where disclaimer appeared; now it does not appear.

---------------- Version 01.00.09 (11-May-1999) ------------------------------------
(11-May-1999)
- Previously, the FORTRAN compiler used for some calculations was Microsoft
  FORTRAN PowerStation.  The new FORTRAN compiler used is DIGITAL Visual Fortran,
  which is self-described as "Year 2000 Ready" (refer to the following web address
  for more details: http://www.digital.com/fortran/y2k.html).

---------------- Version 01.00.10 (21-May-1999) ------------------------------------
(21-May-1999)
- Some minor changes were made to the "splash" and Help--About windows.

---------------- Version 01.00.11 (03-Jun-1999) ------------------------------------
(03-Jun-1999)
- Previously, the Synonyms window would sometimes drop behind the main window.
  This problem no longer occurs.
- Some minor changes were made to the "splash" and Help--About windows.
- The licensing system now handles both Academic and Commercial licenses.

---------------- Version 01.00.12 (07-Jun-1999) ------------------------------------------
(07-Jun-1999)
- Previously, the last-few-files list and other user preferences were stored to
  an unstable location in the Windows directory which sometimes caused these values
  to be lost.  Now, the last-few-files list is stored to $(ETDOT_Path)\DBASE\STEPP.INI,
  for example c:\etdot10\dbase\stepp.ini.
- Previously, a bug in the last-few-files list was present.  This bug would occur 
  when one of the model windows was first opened (e.g. Packed Tower Aeration Design 
  Mode), and the user selected File-Open, and then hit the Cancel button.  In this
  case, all of the files on the last-few-files list would be shifted down one.
  This problem has been fixed.

---------------- Version 01.00.13 (08-Jun-1999) ------------------------------------------
(08-Jun-1999)
- Some minor changes were made to the Help menu (removal of the following
  menu options: "Online Help" and "Technical Assistance Provided By ...".
- The .HLP files were removed

---------------- Version 01.00.14 (24-Jun-1999) ------------------------------------------
(24-Jun-1999)
- An additional command was added to the Help menu: "Manual Printing Instructions".
  This command displays instructions for printing the Adobe Acrobat format or Microsoft
  Word 97 format manual for the software.

---------------- Version 01.00.15 (29-Jun-1999) ------------------------------------------
(29-Jun-1999)
- The Help menu command "Online Manual" now displays the MSWord97 format .DOC manual.
(14-Jul-1999)
- The Superfund Henry's constant database was found to contain a unit error.
  To correct the error, all previous values in the Superfund Henry's constant
  database were multiplied by 0.000018 to put them properly into units of atm-m^3/gmol
  (which are then converted by the StEPP program into dimensionless form).
  For the example of formaldehyde, previously the StEPP program reported a Superfund
  value of 2.24 (dimensionless) for Henry's constant, and now the StEPP program 
  reports a Superfund value of 4.03e-5 for Henry's constant (dimensionless).
      The internal versions of the before- and after-conversion Superfund tables
  for the two database versions are located in the following directory,
  X:\etdot10\code\stepp\comm\misc\990714_superfund_bug_fix, in the following files:
      complete db (all 1800 chemicals).mdb
      release db (400 chemicals).mdb

---------------- Version 01.00.16 (27-Jul-1999) ------------------------------------------
(27-Jul-1999)
- Added button named "Go to web site" on the Help About window.
- Optimized the shell-to-manual command slightly.

---------------- Version 01.00.17 (26-Oct-1999) ----------------------
(26-Oct-1999)
- For all future versions with "Beta" serial numbers, the program
  functionality is reduced so that the user cannot save files, open
  files, or modify the data within the file. The user may only open 
  the Carbon Tetrachloride chemical and view the results.

---------------- Version 01.00.18 (12-Apr-2000) ----------------------
(12-Apr-2000)
- Fixed an error on the Synonyns window.
