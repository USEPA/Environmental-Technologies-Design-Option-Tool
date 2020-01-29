
FaVOr README.TXT file.
======================

VERSION HISTORY.
================

12345678901234567890123456789012345678901234567890123456789012345678901234567890

---------------- Version 1.0.0 (01/05/1999) ---------------------------------
(01/05/1999)
First version compiled with Microsoft Visual Basic 6.0 (MSVB6).

---------------- Version 1.0.1 (01/07/1999) ---------------------------------
(01/07/1999)
This is the first MSVB6 version of FaVOr that appears to contain all of the
features present in the previous version of the software.  The modification
list from MSVB5 to MSVB6 is as follows:

- The entire graphical user interface (GUI) was rewritten in MSVB6.
- No modifications were made to the FORTRAN module.
- Data entry in different types of units is now allowed.
- Several modifications were made to the Physico-Chemical Properties window:
  - The window is now broken up into two tabs.
  - It is now clear which parameters are temperature dependent and which
    are not.
  - An import-from-StEPP mechanism has been added (via the clipboard).
  - Internal correlations were added for water density, water viscosity,
    water vapor pressure, air density, and air viscosity.
  - The user may switch between the following possible sources for each
    parameter: user entry, StEPP, and internal correlations.
- A more intelligent data changed message is now displayed.
  - If any item is changed on a given window, the text "Data Changed" is 
    displayed in red at the bottom of the window.
  - With sub-windows, if the user hits OK, their changes are saved;
    if the user hits Cancel, their changes are abandoned.
  - Within the main window, if any data has changed, and if the user 
    attempts to load a different file or to exit the program, they
    are prompted to save their current file.
- A last-few-files list was added.
- Version tracking (including a version history file) were added.
- Many other minor bug fixes and improvements were made.

---------------- Version 1.0.2 (01/08/1999) ---------------------------------
(01/08/1999)
Temporarily removed the serial number interface.

---------------- Version 1.0.3 (01/12/1999) ---------------------------------
(01/12/1999)
Added correlations for the following oxygen properties:
- Saturation Concentration
- Henry's Constant
- Diffusivity in Water

---------------- Version 1.0.4 (01/12/1999) ---------------------------------
(01/12/1999)
Fixed a bug on the "Define CSTR Parameters" window.  Previously, the average
biomass concentration was improperly calculated as the sum of the biomass
concentrations in the individual CSTRs.  Now, it is properly calculated
as the volume-weighted average of the biomass concentrations in the
individual CSTRs.

---------------- Version 01.00.05 (19-Jan-1999) -----------------------------
(19-Jan-1999)
Modified the labels on the Physico-Chemical Properties window to
indicate that the correlation values for the oxygen properties are 
automatically updated if temperature changes.  The language for the
various asterisks ("*") has been clarified a bit.

---------------- Version 01.00.06 (08-May-1999) -----------------------------
(08-May-1999)
- Added textbox "Percentage Removal" to the Primary Clarifier window.
- Added textbox "Effluent Solids Concentration" to the Secondary 
  Clarifier window.
- For the two new inputs listed above, a new version of the FORTRAN
  calculation module, f32voc.exe, has been compiled.  This new
  version is dated 08-May-1999 at 3:37 PM.

---------------- Version 01.00.07 (17-Jul-1999) -----------------------------
(16-Jul-1999)
- Updated names on startup window.
- Previously, a value was improperly communicated to the FORTRAN module,
  f32voc.exe: the percentage removal of the primary clarifier. Previously,
  the value was written to the FORTRAN input file as a percentage. Now,
  the value is written as a fraction (ranging 0 to 1).
- Corrected unit conversion problem that affected ft^3/min.
- On the [Grit Chamber] window, the units for [Gas Flow Rate] have been
  changed to include m³/min and ft³/min instead of m³/m-h and ft³/m-h.
- For the units in all windows, the following unit clarifications were made:
  - What was previously "m³/m-h" has been replaced with "m³/(m-h)".
  - What was previously "ft³/ft-h" has been replaced with "ft³/(ft-h)".
- Previously, the units on the [Secondary Clarifier] window for
  [Effluent solids concentration] could not be stored. This problem
  has been fixed.
- Previously, the [Effluent solids concentration] value was located
  on the [Secondary Clarifier] window. Now, this value is located on
  the [Aeration Basin] window. This value was moved because it is
  an input to the biomass concentration calculation.
- The behavior of the [Uniform] checkbox under [Biomass Concentration] 
  on the [Define CSTR Parameters] window has been changed. The new
  behavior is as follows:
  - If [Step Feed] is turned on, then 
    [Biomass Concentration]:[Uniform] is automatically forced into the off setting
  - If [Volume]:[Uniform] is turned off, then 
    [Biomass Concentration]:[Uniform] is automatically forced into the off setting
  - If [Gas Flowrate]:[Uniform] is turned off, then 
    [Biomass Concentration]:[Uniform] is automatically forced into the off setting
(17-Jul-1999)
- Previously, the biomass calculation was based on an algorithm presented
  on Page 352 of Davis & Cornwell, "Introduction to Environmental Engineering",
  McGraw-Hill, 1991. Now, the biomass calculation is based on a program
  developed by Hebi Li (BIOCALC.EXE).

---------------- Version 01.00.08 (26-Jul-1999) -----------------------------
(23-Jul-1999)
- A new version of F32VOC.EXE is now distributed. The new file is stamped
  with the date/time of 7/23/99 12:02 PM.
- A new version of BIOCALC.EXE is now distributed. The new file is stamped
  with the date/time of 7/23/99 11:34 AM.
- Previously, on the [Physico-Chemical Properties] window, on the 
  [Oxygen, Water, and Air] tab, the correlation for [Sat'n Conc., O2]
  was incorrect. The old (incorrect) equation was as follows:

      C_O2 = 0.21*(1 atm)/(R*T)*(32 g/gmol)*(1000 mg/g)*(H)

  The new (correct) equation is as follows:

      C_O2 = 0.21*(P)/(R*T)*(32 g/gmol)*(1000 mg/g)/(H)

  where C_O2 is the saturation concentration of oxygen in water, P is
  the atmospheric pressure in atm, R is the gas-law constant in L-atm/(gmol-K),
  T is the temperature in K, and H is the Henry's constant of oxygen
  in dimensionless form. In summary, the old expression was incorrect
  because the Henry's constant was improperly placed, and because it 
  assumed the pressure was always 1.00 atm; the new expression properly
  places the Henry's constant and properly makes use of the pressure
  input by the user as [Barometric Pressure] on the [Environment
  and Contaminant] tab of the [Physico-Chemical Properties] window.
        As a quick test, the program previously reported (in versions
  before 01.00.08) that the correlation value for [Sat'n Conc., O2]
  was 8738 mg/L for a temperature of 20.0 degC and a pressure of 1.00 atm. 
  The new version reports a value of 8.93 mg/L for this temperature
  and pressure.
        As another test, the program now reports that the correlation 
  value for [Sat'n Conc., O2] is 8.51 mg/L for a temperature of 
  20.0 degC and a pressure of 0.952 atm; if the pressure is changed
  to 0.850 atm, the new value of [Sat'n Conc., O2] is 7.59 mg/L.

(26-Jul-1999)
- Previously, the value on the [Grit Chamber] window named
  [Gas Flow Rate] was improperly written to the FORTRAN input file.
  Now, this value is properly written in units of L/min.

- Previously, only one example file was distributed with the 
  software, named sample.fvr; this file is no longer distributed.
  Now, four different example files are distributed (each file
  is located in the examples sub-directory):

  - "sample1.fvr" : 
    includes all the unit operations; the aerated grit chamber, 
    the primary clarifier, the aeration basin and the secondary 
    clarifier are covered; the number of the CSTRs is 1; the 
    biomass concentration is detrmined by the program.

  - "sample2.fvr" : 
    includes all the unit operations; the aerated grit chamber,
    the primary clarifier, the aeration basin and the secondary 
    clarifier are covered; the number of the CSTRs is 3; the 
    biomass concentration is detrmined by the program.

  - "sample3.fvr" : 
    includes all the unit operations; the aerated grit chamber,
    the primary clarifier, the aeration basin and the secondary 
    clarifier are uncovered; the number of the CSTRs is 3; the 
    biomass concentration is detrmined by the program.

  - "sample4.fvr" :
    does not include the aerated grit chamber, the primary 
    clarifier, the aeration basin and the secondary clarifier; 
    the number of the CSTRs is 3; the biomass concentration is 
    determined by the program.

---------------- Version 01.00.09 (17-Aug-1999) -----------------------------
(17-Aug-1999)
- Turned on the licensing system.

---------------- Version 01.00.10 (23-Sep-1999) -----------------------------
(22-Sep-1999)
- On the main window result grid:
  - Removed the double-row that was labeled as ["Volatilization"] and 
    ["% of all"] immediately underneath it; now, the double-row marked 
    ["Stripping"] contains the sum of what was contained in the previous two
    double-rows marked ["Volatilization"] and ["Stripping"]
  - Changed the colors somewhat
  - Split the previous single-row labelled ["Effluent liq conc"] into
    a triple-row with the following labels:
    - "Dissolved Effluent Liquid Concentration"
    - "Sorbed Effluent Liquid Concentration"
    - "Effluent Solids Concentration"
  - A button is provided that switches all main-window values into English
    units; and another button switches to S.I. units.
- A new version of F32VOC.EXE is now distributed. The new file is stamped
  with the date/time of 9/22/99 4:45 PM.
(23-Sep-1999)
- On the main window result grid:
  - Changed label "Effluent Solids Concentration" to "Effluent Volatile
    Solids Concentration" 
  - Added two new menu items:
    - View -- Change to English Units
    - View -- Change to S.I. Units
---------------- Version 01.00.11 (03-Jan-2000) -----------------------------
(01-03-2000)
- On the main window result grid:
  - Calculates and fills data fields when previously saved file is opened
- Menu changes:
  - Added File, Print... menu option 
    - Preprints results to Excel which allows user to print to a printer
    - Added validation to Print option for user to recalulate results

   if data has been changed