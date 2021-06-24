
ASAP README.TXT file.
=====================

VERSION HISTORY.
================

---------------- Version 01.00.00 (21-Oct-1998) -----------------------------------------
First version compiled with Microsoft Visual Basic 5.0.

---------------- Version 01.00.01 (23-Oct-1998) -----------------------------------------
Added link to ASAP.HLP file from Help menu ("Online Help").
Added link to ASAP.PDF file from Help menu ("Online Manual").

---------------- Version 01.00.02 (10-Nov-1998) -----------------------------------------
First version compiled with Microsoft Visual Basic 6.0.

---------------- Version 01.00.03 (05-Mar-1999) -----------------------------------------
Updated MTU logo on startup window.

---------------- Version 01.00.04 (31-Mar-1999) -----------------------------------------
When the user selects "CW O2 Transfer Test Data" on the Bubble Aeration
Design Mode or Rating Mode window, a window appears titled "Find
KLa, O2 from Clean Water Oxygen Transfer Test Data".  Previously, this
window was empty of data due to a mistake during the upgrade process to
version 01.00.00 in October of 1998.  Now, the user interface and calculations
are performed properly.

---------------- Version 01.00.05 (12-Apr-1999) -----------------------------------------
(12-Apr-1999)
Rephrased and moved disclaimer to the install software.

---------------- Version 01.00.06 (15-Apr-1999) -----------------------------------------
(15-Apr-1999)
- Previously, a number of obscure unit-related problems occurred on the
  Packed Tower Aeration Design Mode window.  The affected values were the
  Flows and Loadings settings of "Water Flow Rate" and "Air Flow Rate"
  which appeared with an improper number of significant figures and/or
  in the improper units.  Now, this problem has been fixed.
- Previously, when a print out was generated from the Packed Tower Aeration
  windows (Design Mode or Rating Mode), many of the units were missing from 
  the properties.  Similar problems occurred for print outs of the 
  Surface Aeration window (Rating Mode).  These problems have been fixed.
- ASAP bug fixed related to the "Optimize with All Contaminants"
  button. Problem: If "Optimize ..." button found the previously selected
  Design Contaminant as Remaining the Design Contaminant after the "Optimize
  ..." button was clicked, the program did not generate a click event on the
  Design Contaminant Combo box as required. The solution was to set the list
  index for the combo box to -1 before setting it to the value for the
  optimal contaminant to ensure the click event will occur.

    Old Code:
       frmPTADScreen1!cboSelectCompo.ListIndex = scr1.ID_OptimalDesignContaminant - 1
       frmPTADScreen1!cboSelectCompo.SetFocus

    Corrected Code:
       frmPTADScreen1!cboSelectCompo.ListIndex = -1
       frmPTADScreen1!cboSelectCompo.ListIndex = scr1.ID_OptimalDesignContaminant - 1
       frmPTADScreen1!cboSelectCompo.SetFocus

---------------- Version 01.00.07 (16-Apr-1999) -----------------------------------------
(16-Apr-1999)
- Previously, when a data file was loaded into the Packed Tower Design Mode window,
  the design contaminant name was improperly loaded.  Now, the design contaminant name
  is properly loaded.

---------------- Version 01.00.08 (12-May-1999) -----------------------------------------
(12-May-1999)
- Previously, the FORTRAN compiler used for some calculations was Microsoft
  FORTRAN PowerStation.  The new FORTRAN compiler used is DIGITAL Visual Fortran,
  which is self-described as "Year 2000 Ready" (refer to the following web address
  for more details: http://www.digital.com/fortran/y2k.html).

---------------- Version 01.00.09 (21-May-1999) -----------------------------------------
(21-May-1999)
- Some minor changes were made to the "splash" and Help--About windows.

---------------- Version 01.00.10 (03-Jun-1999) ------------------------------------------
(03-Jun-1999)
- Some minor changes were made to the "splash" and Help--About windows.
- The licensing system now handles both Academic and Commercial licenses.

---------------- Version 01.00.11 (07-Jun-1999) ------------------------------------------
(07-Jun-1999)
- In Packed Tower Aeration Design Mode, a minor bug was fixed in the Select Packing
  window.  Previously, the user was unable to enter leading or terminal space 
  characters in the Name field of the User-Modified Database; this made it
  difficult (but not impossible) to enter any text string containing a space 
  character.  Now, the user can enter these space characters.  Before the packing 
  record is stored to the database, all leading and terminal space characters 
  are removed.

---------------- Version 01.00.12 (07-Jun-1999) ------------------------------------------
(07-Jun-1999)
- Previously, the last-few-files list and other user preferences were stored to
  an unstable location in the Windows directory which sometimes caused these values
  to be lost.  Now, the last-few-files list is stored to $(ETDOT_Path)\DBASE\ASAP.INI,
  for example c:\etdot10\dbase\asap.ini.
- Previously, a bug in the last-few-files list was present.  This bug would occur 
  when one of the model windows was first opened (e.g. Packed Tower Aeration Design 
  Mode), and the user selected File-Open, and then hit the Cancel button.  In this
  case, all of the files on the last-few-files list would be shifted down one.
  This problem has been fixed.

---------------- Version 01.00.13 (08-Jun-1999) ------------------------------------------
(08-Jun-1999)
- Previously, on the Surface Aeration window, if the value in the textbox named 
  "Water Flow Rate" was changed, the value named "Power Required per Tank" was
  not redisplayed properly (unless the user selected a new design contaminant,
  and then reselected the original design contaminant).  This problem has been fixed.
- Some minor changes were made to the Help menu (removal of the following
  menu options: "Online Help" and "Technical Assistance Provided By ...".
- The same Help menu is now accessible from all model windows
- The .HLP files were removed

---------------- Version 01.00.14 (14-Jun-1999) ------------------------------------------
(14-Jun-1999)
- The default data files (dbase\def*.*) for Design and Rating Modes for all three 
  models were replaced.
- Previously, if a default data file (dbase\def*.*) was deleted, the program would
  crash upon entering the applicable model, e.g. a crash would occur on entry to
  the Packed Tower Aeration Design Mode window if dbase\default.des was deleted.
  Now, an error message is displayed that recommends the user re-install the software.
- Previously, if a file was saved while in Surface Aeration Rating Mode, the window
  caption would improperly be changed to "Surface Aeration - Design Mode".
  This has been fixed.

---------------- Version 01.00.15 (24-Jun-1999) ------------------------------------------
(24-Jun-1999)
- An additional command was added to the Help menu: "Manual Printing Instructions".
  This command displays instructions for printing the Adobe Acrobat format or Microsoft
  Word 97 format manual for the software.

---------------- Version 01.00.16 (29-Jun-1999) ------------------------------------------
(29-Jun-1999)
- The Help menu command "Online Manual" now displays the MSWord97 format .DOC manual.

---------------- Version 01.00.17 (27-Jul-1999) ------------------------------------------
(27-Jul-1999)
- Added button named "Go to web site" on the Help About window.
- Optimized the shell-to-manual command slightly.

---------------- Version 01.00.18 (26-Oct-1999) ----------------------
(26-Oct-1999)
- For all future versions with "Beta" serial numbers, the program
  functionality is reduced so that the user cannot save files or open
  files; they can only modify the default data.


