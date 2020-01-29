
AdOx README.TXT file.
======================================================

VERSION HISTORY.
================

---------------- Version 01.00.01 (??) -------------------------------------
(18-Dec-1998)


---------------- Version 01.00.02 (17-May-1999) -------------------------------------
(17-May-1999)
- Compiled application in VB 6.0.  No problems.

---------------- Version 1.0.7 (07-September-1999) -------------------------------------
(26-July-1999)
- A complete rewrite of the program was performed using the Generic App template, which,
  	among other things, changed the data storage from text files to Access files.
- Previously the main window and photochemical window had fixed units.  Now, these units 
	can be reset by the user during data entry.
- There was a command button added for editing and calculating Dye Study data.  It was added next
	to the Number of Tanks field.  
	-When this button is clicked, a window is opened and the user can either manually enter 
	time and concentration data for up to 400 rows, or paste data from the clipboard.  
	-When the data entry is completed, the user clicks on a Calculate button which 
	runs a Fortran program called pec.exe that creates a text file with the results.  
	-The data entered and the output text file are both stored in the Access file when the
	application is saved.
- Two of the screen titles were changed, "PhotoChemical Properties" to "Photochemical Reactor Properties" 
	and "Numerical Simulation Parameters" to "Time Parameters".


