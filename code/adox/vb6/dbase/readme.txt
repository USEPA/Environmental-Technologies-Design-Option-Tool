
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
(07-September-1999)
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

---------------- Version 1.0.10 (26-October-1999) -------------------------------------
(26-October-1999)
- The Dye Study command button was changed to call a stand-alone application.  All dye study
	data and results will be saved within the dye study application rather than in the AdOx
	application.

---------------- Version 1.0.11 (22-May-2000) -------------------------------------
(22-May-2000)
- Added a database for Target Components containing CAS Number, Molecular Weight, and
	Second Rate Constant information.
- On the PhotoChemical Reactor Properties window, the Light Specification Method wasn't 
	adjusting the values in the grid correctly.
---------------- Version 1.0 (30-Aiugust-2000) -------------------------------------
(30-August-2000)
- Included online documentation in pdf format.

---------------- Version 1.0.1 (21-Apr-2002) ---------------------------------------
(21-Apr-2002)
- Corrected mistake in writing to input file for extinction coefficient and quantum yield.
- Added form to display tabulated extinction coeff. vs. wavelength for hydrogen peroxide.

(12-Aug-2002)
- Updated Fortran code (adoxfor.exe) to fix calculation of TIC from Alkalinity.