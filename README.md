# ETDOT

The **Environmental Technologies Design Option Tool** (ETDOT) was developed by National Center for Clean Industrial and Treatment Technologies (CenCITT) at Michigan Technological University (MTU).

Version 1.0: Copyright 1994–2005

* David R. Hokanson
* David W. Hand
* John C. Crittenden
* Tony N. Rogers
* Eric J. Oman

Version 1.0.50 - Updated AdDesignS
* Michael Verma
* Feng Shang

This GitHub repository includes FORTRAN and VisualBasic Code associated with the suite of programs distributed within ETDOT.

Software that is included:
* Adsorption Design Software for Windows (AdDesignS) Version 1.0
* Advanced Oxidation Process Software (AdOx) Version 1.0.2
* Aeration System Analysis Program (ASAP) Version 1.0
* Biofilter Design Software Version 1.0.27
* Continuous Flow Pore Surface Diffusion Model for Modeling Powdered Activated Carbon
* Dye Study Program (DyeStudy) Version 1.0.0
* Predictive Software for the Fate of Volatile Organics in Municipal Wastewater Treatment Plants (FaVOr) Version 1.0.11
* Ion Exchange Design Software (IonExDesignS) Version 1.0.0 [See notes below]
* Software to Estimate Physical Properties (StEPP) Version 1.0


In 2019, MTU signed a software transfer agreement with the United States Environmental Protection Agency (EPA) granting the EPA the rights to maintain and non-commercially and publicly distribute the ETDOT suite of packages. 


*Disclaimer*: 

The United States Environmental Protection Agency (EPA) GitHub project code is provided on an "as is" basis and the user assumes responsibility for its use. EPA has relinquished control of the information and no longer has responsibility to protect the integrity, confidentiality, or availability of the information. Any reference to specific commercial products, processes, or services by service mark, trademark, manufacturer, or otherwise, does not constitute or imply their endorsement, recommendation or favoring by EPA. The EPA seal and logo shall not be used in any manner to imply endorsement of any commercial product or activity by EPA or the United States Government. 


# Installation Instructions
## AdDesignS - Updated Version
The updated version of AdDesignS can be run without installation and is compatible with Windows 10.

1. Download zip file ('AdDesignS_updated.zip') under [Releases](https://github.com/USEPA/Environmental-Technologies-Design-Option-Tool/releases/tag/1.0.50).
2. Unzip/Unpack zip file to desired location. Select an easily accesible location, as this will be the location where executables will be accessed when needed. 
3. Navigate to unzipped folder. 
4. Select "ads.exe" to run AdDesignS
5. If desired, users can select "Create Shortcut" by right-clicking on 'ads.exe'. A shortcut file will be created in that folder, and users can move this to the Desktop or into another location for more convenient access.

## ETDOT Suite

This software requires *Administrator Rights* to a computer to install and to run. Files are installed directly to a folder X:\ETDOT10\... where X is the system main drive. [To ensure that a user can continue to use ETDOT if a user does not have "Administrator Rights", the user must be provided with "Full Control Access" (read/write) rights to the folder X:\ETDOT10.]

1. Download zip file ('etdot_1-0.zip') under [Releases](https://github.com/USEPA/Environmental-Technologies-Design-Option-Tool/releases/tag/1.0.50).
2. Unzip/Unpack zip file
3. Run *setup.exe* and follow prompts
4. When prompted enter license key: CAADV0-R74JM-QXCNP-7EER9-1AT72
5. To run each module in Windows 7 or newer: Edit *properties* of the program to be run and select Compatibility Tab and  "run in compatibility mode". Select Windows 98/Me from the Compatibility Mode dropdown menu.
6. For AdDesignS: If you get "Error #435...index" or an error about "Biot number out of range", please use English as your language in Windows OS. The comma marker for a decimal used by some non-English languages cannot be passed to the solvers in the program and produce the listed errors.

Available Users manuals will be located in the modules subfolder within the help folder.

# Notes on current software

The **ETDOT** suite of software packages consists of a FORTRAN engine with a Visual Basic (version 6) graphical user interface. The VB6 portion of the code relies on ActiveX control files which are located in the repository, however, these are an older coding standard and no longer supported with current versions of Visual Studio 20##. Precompiled engine files are included.

IonExDesignS within **ETDOT** is currently non-functional with no current plans to update. Please refer to the [Ion Exchange Model](https://github.com/USEPA/Water_Treatment_Models/tree/master/ShinyApp) for Ion Exchange modeling needs. The Ion Exchange Model is a R Shiny app that uses an updated ion exchange model for both gel-type and macroporous ion exchange resins. The corresponding Python version of the same model can be found along with an updated PSDM model for modeling granular activated carbon (GAC) at [Water Treatment Models](https://github.com/USEPA/Water_Treatment_Models).
