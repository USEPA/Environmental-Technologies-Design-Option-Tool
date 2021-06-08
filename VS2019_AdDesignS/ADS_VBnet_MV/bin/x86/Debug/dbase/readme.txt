
AdDesignS/IAFM README.TXT file.
===============================

VERSION HISTORY.
================

---------------- Version 1.0.0 ----------------------------------------------------------
    First version compiled with Microsoft Visual Basic 5.0.

---------------- Version 1.0.7 ----------------------------------------------------------
    Converted file format to DAO 3.5 MDB format (Microsoft
    Access 97 compatible).  All past file formats can be
    opened in AdDesignS, but only the DAO 3.5 MDB format
    can be saved.

---------------- Version 1.0.8 ----------------------------------------------------------
    New version 1.1 serial numbers now supported (CA*).

---------------- Version 1.0.10 (10/21/98) ----------------------------------------------
    Fixed a bug with the ECM that caused an "Input Past
    End of File" error to occur.  The problem was due to
    an out-of-date ECM5.EXE module.  The correct module
    is dated 9/20/98.
    NOTE: The eleven CPAS CDs sent out on 10/21/98 contain the
    AdDesignS program with the bug; if anyone requests, the solution 
    is simply to update ECM5.EXE (no need to update ADS.EXE).

---------------- Version 1.0.11 (10/23/98) ----------------------------------------------
    Added link to ADS.PDF file from Help menu ("Online Manual").

---------------- Version 1.0.12 (10/26/98) ----------------------------------------------
(10/26/98)
    Added PSDM-in-Room model for eventual general distribution.
(11/9/98) 
    First version compiled with Microsoft Visual Basic 6.0 (MSVB6).  
    Microsoft has stated that the MSVB6 compiler is Y2K-compliant.

---------------- Version 1.0.14 (11/22/98) ----------------------------------------------
(11/22/98)
    Got PSDM-in-Room model functioning.  This model is only available to 
    those users who have purchased it.

---------------- Version 1.0.15 (12/01/98) ----------------------------------------------
(12/01/98)
    Added time-variable influent capability to PSDM-in-Room model.
(12/01/98)
    Fixed bug where saving AdDesignS files (*.dat) to network drives resulted in an
    error.  Now, saving to network drives works just as well as saving to non-network
    drives.

---------------- Version 01.00.16 (06-Jan-1999) ------------------------------------------
(06-Jan-1999)
    Previously, the air density correlation was off by a factor of 1000.  For example,
    at 20 degC and 0.952 atm, the air density was reported by AdDesignS as 
    1.15e-6 g/cm^3.  Now, the proper value of 1.15e-3 g/cm^3 is reported.

---------------- Version 01.00.17 (03-Feb-1999) ------------------------------------------
(03-Feb-1999)
    - Corrected unit conversion problem that affected ft^3/min.

---------------- Version 01.00.18 (15-Feb-1999) ------------------------------------------
(15-Feb-1999)
    - Previously, the user was allowed to specify more than one axial element (AE) for
      the PSDM-in-Room model.  At this time the PSDM-in-Room model
      is not set up for AE>1.  Now, the AdDesignS program does not allow the user
      to run the PSDM-in-Room model if they have set AE>1 (only AE=1 is allowed).

---------------- Version 01.00.19 (22-Feb-1999) ------------------------------------------
(22-Feb-1999)
    - Previously, the isotherm parameter estimation module titled "D-R Based on Spreading
      Pressure Evaluation" reported zero values for the lower correlation limit,
      upper correlation limit, regression r-squared value, and root mean square
      error value.  These values are not calculated by the estimation module.
      Now, these values are instead hidden on the output window for this calculation.
      In the future, the correlation will be modified to properly output these values.

---------------- Version 01.00.20 (02-Mar-1999) ------------------------------------------
(02-Mar-1999)
    - Previously, when the user clicked on the "Use These Adsorbent Specifications"
      button in the Adsorbent Database window, the name of the adsorbent was not copied
      onto the main window.  Now, this copy does occur.
    - A new opening window is displayed, depending on whether or not the PSDM-in-Room 
      model is present.
    - Previously, when the PSDM-in-Room model was active, the ECM, CPHSDM, and PSDM
      models could still be ran (using the Run and Results menus).  Now, the menu
      entries for the ECM, CPHSDM, and PSDM models are invisible when the PSDM-in-Room 
      model is active.
    - Now, the software consists of two separate versions:
      - AdDesignS (Adsorption Design Software)
        - Contains the ECM, CPHSDM, and PSDM models
      - IAFM (Indoor Air Filtration Model)
        - Contains the PSDM-in-Room Models

---------------- Version 01.00.21 (12-Apr-1999) ------------------------------------------
(12-Apr-1999)
    - Rephrased and moved disclaimer to the install software.

---------------- Version 01.00.22 (13-Apr-1999) ------------------------------------------
(13-Apr-1999)
    - Previously, when the user edited an adsorbent, the particle radius showed up
      properly in units of centimeters on the "Adsorbent Database" window, but improperly
      in units of millimeters on the "Editing an Adsorbent" window.  This problem
      has been corrected.

---------------- Version 01.00.23 (27-Apr-1999) ------------------------------------------
(27-Apr-1999)
    - Previously, the variant of AdDesignS that contained the PSDM-in-Room model
      was known as the Indoor Air Filtration Model (IAFM).  It is now known as the
      Indoor Air Adsorption Design Software (IndoorAirAdDesignS).

---------------- Version 01.00.24 (12-May-1999) ------------------------------------------
(12-May-1999)
    - Previously, the Print button on the "Results for the PSDM" window did not
      print a copy of the C/Co plot.  During the conversion from Visual Basic 3.0
      to Visual Basic 5.0 and 6.0, this printing capability was lost.  Now,
      the plot is printed properly.  This command should only be used for "quick
      and dirty" plots.  To allow for better formatting, the user should copy
      the plotted results to Excel (or some other spreadsheet program) and
      then plot them manually.
    - Similarly, the Print button on the "Results for the Constant Pattern 
      Model (CPHSDM)" window did not print the C/Co plot in previous versions.
      This has been fixed.
    - Previously, the Print button on the "Results for the PSDM" window created
      a printout that had some minor formatting problems.  One problem was that
      the section "PSDM Module Input Variables" printed numbers at random 
      horizontal locations.  Another problem was that several pages with very 
      little text were printed.  Both problems have been fixed.
(12-May-1999)
    - Previously, the FORTRAN compiler used for some calculations was Microsoft
      FORTRAN PowerStation.  The new FORTRAN compiler used is DIGITAL Visual Fortran,
      which is self-described as "Year 2000 Ready" (refer to the following web address
      for more details: http://www.digital.com/fortran/y2k.html).

---------------- Version 01.00.25 (14-May-1999) ------------------------------------------
(14-May-1999)
- Previously, two of the Isotherm Parameter Estimation (I.P.E.) modules were slightly
  incorrect.  The affected routines were ADLIQ() (3-Parameter Polanyi Isotherm 
  Correlation) and HOFMAN() (D-R Uniform Adsorbate).  In each routine, the
  molar volume at the operating temperature was determined from the liquid density
  at the operating temperature; this is the incorrect molar volume to use.
  To correct the routine, the molar volume at the normal boiling point is
  now input to each routine.  (The reason why the MV at the NBP must be used is that
  the fits to determine Polanyi parameters make use of MV at the NBP.)
      While the old version deviated from "hand" calculations by 10-30%, the
  new version matches to within 1%.  The new file IPES4.EXE is required to
  run the new versions of the I.P.E. modules.

---------------- Version 01.00.26 (20-May-1999) ------------------------------------------
(20-May-1999)
- The corrections made in Version 1.0.25 were rolled back.  Now, both ADLIQ()
  and HOFMAN() both make use of the molar volume at the operating temperature.
  The file IPES4.EXE was updated.
- As a future concern, a new field must be added to each record in the adsorbent 
  database that describes whether the Polanyi parameters were normalized using
  the molar volume @ the normal boiling point, or the molar volume @ the operating
  temperature.  In addition, this switch will be displayed (and be modifiable)
  on the Polanyi parameters window in AdDesigns.  And finally, this switch will need
  to be communicated to the ADLIQ() and HOFMAN() routines to vary the calculations
  accordingly.

---------------- Version 01.00.27 (21-May-1999) ------------------------------------------
(21-May-1999)
- Some minor changes were made to the "splash" and Help--About windows.

---------------- Version 01.00.28 (03-Jun-1999) ------------------------------------------
(03-Jun-1999)
- Some minor changes were made to the "splash" and Help--About windows.
- The licensing system now handles both Academic and Commercial licenses.

---------------- Version 01.00.29 (08-Jun-1999) ------------------------------------------
(08-Jun-1999)
- Some minor changes were made to the "splash" window.
- Some minor changes were made to the Help menu (removal of the following
  menu options: "Online Help" and "Technical Assistance Provided By ...".

---------------- Version 01.00.30 (17-Jun-1999) ------------------------------------------
(17-Jun-1999)
- Previously, the isotherm database contained records with the same CAS number but
  differing chemical names. All of the isotherms for each affected record have been
  merged together to avoid confusion. (Refer to the following internal directory for more 
  information: X:\etdot10\code\ads\comm\990617_conversion_of_isotherm_database.)

---------------- Version 01.00.31 (24-Jun-1999) ------------------------------------------
(24-Jun-1999)
- An additional command was added to the Help menu: "Manual Printing Instructions".
  This command displays instructions for printing the Adobe Acrobat format or Microsoft
  Word 97 format manual for the software.

---------------- Version 01.00.32 (29-Jun-1999) ------------------------------------------
(29-Jun-1999)
- The Help menu command "Online Manual" now displays the MSWord97 format .DOC manual.

---------------- Version 01.00.33 (12-Jul-1999) ------------------------------------------
(12-Jul-1999)
- Previously, some strange formatting errors existed when the PSDM or CPHSDM model
  output was saved to an Excel (.XLS) file. Typically, when the saved .XLS file
  was opened in Microsoft Excel, some values would be displayed in units of US 
  currency (dollars, $). This problem has been fixed.
- Previously, the isotherm database contained two isotherms for chemical name
  "1,1,2-Trichloro-ethane" and CAS of 79005. This was a typo, and the chemical
  named has been replaced with "1,1,1-Trichloroethane" and CAS of 71556.
  The records affected were those with ID of 828 and 829.

---------------- Version 01.00.34 (27-Jul-1999) ----------------------
(27-Jul-1999)
- Added button named "Go to web site" on the Help About window.
- Optimized the shell-to-manual command slightly.

---------------- Version 01.00.35 (07-Sep-1999) ----------------------
(07-Sep-1999)
- Revised PSDM-in-Room model. Replaced proom10c.exe with proom11.exe.

---------------- Version 01.00.36 (10-Sep-1999) ----------------------
(10-Sep-1999)
- Previously, the program would fail to run if any of these files
  were not present:
  - dbase\misc1.dat
  - dbase\template.dat
- Now, the following additional files are also checked for:
  - dbase\beds1.txt
  - dbase\beds2.txt
  - dbase\carbon.mdb
  - dbase\corr_com.txt
  - dbase\isotherm.mdb
  - dbase\water_co.txt

---------------- Version 01.00.37 (21-Oct-1999) ----------------------
(21-Oct-1999)
- Previously, the plot on the window [Results for the PSDM] was only
  obtainable for a y-axis of C/Co (or Cr/Cr,ss in the case of the
  PSDM-in-Room model). Now, the following y-axis settings are
  selectable by the user:
  - For the PSDM:
    - C/Co
    - ug/L
    - mg/L
    - g/L
    - ppb
    - ppm
  - For the PSDM-in-Room model:
    - Cr/Cr,ss (only selectable if the value of Cr,ss is non-zero 
      for all compounds)
    - ug/L
    - mg/L
    - g/L
    - ppb
    - ppm
  - NOTE: There is a bug in the ppm and ppb calculations for liquid
    phase that must be fixed in a future version.
- An additional PSDM model variation has been added:
  - Previously, the two possible PSDM variations were:
    - PSDM
    - PSDM-in-Room
  - Now, the three possible PSDM variations are:
    - PSDM: Pore and Surface Diffusion Model, without reaction, 
      one fixed bed only
    - PSDMR-in-Room: PSDM, with reaction, one fixed bed inside
      a room treated as a continuously stirred tank reactor (CSTR)
    - PSDMR Alone: PSDM, with reaction, one fixed bed only
  - All three PSDM variations are accessed from the Run menu on
    the main window.
  - A new PROOM11.EXE is distributed, dated 1999-Oct-21.

---------------- Version 01.00.38 (27-Oct-1999) ----------------------
(26-Oct-1999)
- For all future versions with "Beta" serial numbers, the program
  functionality is reduced so that the user cannot save files or
  modify the data within the file. The user may only open the
  GAS.DAT and LIQUID.DAT files in the EXAMPLES subdirectory.
(27-Oct-1999)
- The "Beta" version was changed slightly to more strictly require
  the user to only open GAS.DAT or LIQUID.DAT.

---------------- Version 01.00.39 (27-Oct-1999) ----------------------
(27-Oct-1999)
- Previously (since 21-Oct-1999), for liquid-phase results, the
  results displayed on the [Results for the PSDM] window were
  incorrect for a Y Axis Type of ppm or ppb. This problem has been
  fixed.

---------------- Version 01.00.40 (18-Nov-1999) ----------------------
(18-Nov-1999)
- Previously, for the PSDMR in Room model, the following values were
  only specifiable as constants (not time-variable):
  - Concentration in Influent Stream to Room
  - Mass Emission Rate Within Room
  Now, these variables are specifiable as time-variable values from
  within the [Parameters for PSDMR Model] window.
  - A new PROOM12.EXE is distributed, dated 18-Nov-1999.
(22-Nov-1999)
- Fixed a bug introduced on 18-Nov-1999 where a division-by-zero
  error would result when the [Component Properties] window was
  opened.

---------------- Version 01.00.41 (05-Jan-2000) ----------------------
(05-Jan-2000)
- Previously, fouling was not permitted for gas-phase PSDMR 
  calculations. Fouling is now permitted for gas-phase PSDMR 
  calculations.
- Previously, the "Changes" flag (dirty flag) was never set when
  the user clicked OK from the Fouling window. Now, the "Changes"
  flag is set.

---------------- Version 01.00.42 (21-Jan-2000) ----------------------
(17-Jan-2000)
- Previously, for the PSDMR in Room model, the following values were
  only specifiable as constants (not time-variable):
  - Freundlich K
  Now, these variables are specifiable as time-variable values from
  within the [Parameters for PSDMR Model] window.
  - A new PROOM12.EXE is distributed, dated 17-Jan-2000.
(21-Jan-2000)
- Previously, the minimum value for Vapor Pressure that could be
  entered on the Component Properties window was 1e-2 Pa. Now, the
  minimum value is 1e-10 Pa.

---------------- Version 01.00.43 (28-Dec-2000) ----------------------
- Added a command button to most screens to enable the user to print the screen 
  to printer.


