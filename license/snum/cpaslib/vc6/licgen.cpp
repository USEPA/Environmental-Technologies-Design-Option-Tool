///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    SOURCE CODE FILE:
//        LICGEN.CPP
//
//    PURPOSE:
//        HOLDS THE CPAS.LIC GENERATION ROUTINES.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//
//  REDISTRIBUTION NOTES:
//  =====================
//
//  As of today (July 2, 1999) this program (cpaslib.dll) is compiled to 
//  the following directory:
//
//      X:\etdot10\license\snum\cpaslib\vc6\Debug
//
//  It must be copied to the following directories to be fully incorporated
//  into future builds of the install software:
//
//      X:\etdot10\extravb\cpaschk\test\dbase
//      X:\etdot10\programs_vb6\cpaschk\dbase
//      X:\is5.5_installs\ETDOT 981112\Setup Files\Compressed Files\Language Independent\Intel 32
//
//
//  IMPORTANT NOTE:
//  ===============
//
//  Previously, the CPASCHK.EXE program was used to generate the license file.
//  This old code still resides in CPASCHK_11 . Do_Create_File_v11() of the
//  CPASCHK.EXE program for the ETDOT install software (located at the following
//  directory: X:\etdot10\extravb\cpaschk\vb6).
//
//  As of (I can't remember when), this DLL is used to generate the license file.
//  Thus snumCpasLicGenerate() completely replaces the old generator in
//  the CPASCHK.EXE program. The old Do_Create_File_v11() does NOT need to
//  ever be maintained again; in fact, it could be commented out.
//
//  Eric J. Oman
//  2:03 PM 7/2/99
//
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


// COMMON C/C++ FUNCTION PROTOTYPES.
#include <errno.h>
#include <fcntl.h>
#include <io.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <time.h>

// FUNCTION PROTOTYPES.
#include "cpaslib.h"
#include "internal.h"
#include "licgen.h"


// GLOBAL VARIABLES.
// NOTE: STORING THESE VARIABLES AS LOCAL VARIABLES WITHIN THE
// snumCpasLicGenerate() FUNCTION RESULTED IN A STACK OVERFLOW.
char ThisSerialNumber[100];
char ForceAll_Z_VERSIONCODE[100];
char ForceAll_Z_RELEASETYPE[100];
char ForceAll_Z_EXPIRATIONDATE[100];
char ForceAll_Z_VERSIONTYPE[100];
struct Recognized_Mod_Type rmt[RMT_SIZE];
int rmt_idx_low;
int rmt_idx_high;
struct ProgramKey_Data_Type pkdt_new[RMT_SIZE];
int pkdt_new_idx_low;
int pkdt_new_idx_high;
struct LicFile_Data_Type lfdt;
char fn_CPASLIC[300];
char fn_MTCHKLIC[300];

// GLOBAL VARIABLES -- SHARED WITH LicFile_*() FUNCTIONS.
int LicFile_keyidx[2][256];
#define LICFILE_KEY (192)
#define LICFILE_STRINGSIZE_MAX (90)
// 'NOTE: LICFILE_STRINGSIZE_MAX CAN'T BE INCREASED PAST ABOUT 254 UNLESS
// 'THE ENCRYPT/DECRYPT ROUTINES ARE REVISED.



// PROTOTYPE:
//     extern "C" int __stdcall snumCpasLicGenerate(char *spCpasDir,char *spWinDir,char *spNumber,char *spUserName,char *spUserCompany);
// INPUTS:
//     spCpasDir = DIRECTORY NAME OF CPAS DIR (E.G. "C:\ETDOT10")
//     spWinDir = DIRECTORY NAME OF MAIN WINDOWS DIR (E.G. "C:\WINDOWS" OR "C:\WINNT")
//     spNumber = SERIAL NUMBER
//     spUserName = USER NAME
//     spUserCompany = USER COMPANY
// PURPOSE:
//     GENERATES CPAS.LIC LICENSE CHECKING FILE.
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
// RETURNS:
//     0 = FAILED!
//     1 = SUCCESSFUL.
//
extern "C" int __stdcall snumCpasLicGenerate(
    char *spCpasDir,
    char *spWinDir,
    char *spNumber,
    char *spUserName,
    char *spUserCompany)
{
//
// LOCAL VARIABLES.
//
char fpath_OutputDir[100];
//char *fn_GoodSerialNumber;
//char *fn_BadSerialNumber;
//char *fn_NEWLIC;
FILE *fp;
int iSnumPlain[21];  
int RetVal;
int iExpiresDay;
int iExpiresMonth;
int iExpiresYear;
int num_pkdt_new;
int i;
long filesize_CPASLIC;
char DummyStr[100];
int This_Position;
long offset_This_PK;
long offset_This_Data;
int AnyErrors;
  //
  // START OF CODE.
  //
  if ( DEBUGMODE )
  {
    // OVERWRITE OLD FILE, IF ANY.
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "w");
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "INPUT PARAMETERS TO snumCpasLicGenerate():\n");
    fprintf(fp, "spCpasDir = `%s`\n", spCpasDir);
    fprintf(fp, "spNumber = `%s`\n", spNumber);
    fprintf(fp, "spUserName = `%s`\n", spUserName);
    fprintf(fp, "spUserCompany = `%s`\n", spUserCompany);
    fclose(fp);
  }
  //
  // DETERMINE OUTPUT DIRECTORY.
  //
  strcpy(fpath_OutputDir, spCpasDir);
  strcat(fpath_OutputDir, "\\DBASE");
//  //
//  // DELETE {CPAS}\DBASE\OKNUM.X AND {CPAS}\DBASE\BADNUM.X IF PRESENT.
//  //
//  strcpy(fn_GoodSerialNumber, fpath_OutputDir);
//  strcat(fn_GoodSerialNumber, "\\");
//  strcat(fn_GoodSerialNumber, LICFILE_GoodSerialNumber);
//  strcpy(fn_BadSerialNumber, fpath_OutputDir);
//  strcat(fn_BadSerialNumber, "\\");
//  strcat(fn_BadSerialNumber, LICFILE_BadSerialNumber);
//  if (_access(fn_GoodSerialNumber,0) == -1)
//  {
//    // FILE DOES NOT EXIST; DO NOTHING.
//  }
//  else
//  {
//    // FILE EXISTS; DELETE IT.
//    if (_unlink(fn_GoodSerialNumber) == -1)
//    {
//      // UNABLE TO DELETE FILE!  EXIT WITH ERROR.
//      goto err_snumCpasLicGenerate_General;
//      return(0);
//    }
//  }
//  if (_access(fn_BadSerialNumber,0) == -1)
//  {
//    // FILE DOES NOT EXIST; DO NOTHING.
//  }
//  else
//  {
//    // FILE EXISTS; DELETE IT.
//    if (_unlink(fn_BadSerialNumber) == -1)
//    {
//      // UNABLE TO DELETE FILE!  EXIT WITH ERROR.
//      goto err_snumCpasLicGenerate_General;
//      return(0);
//    }
//  }
//  //
//  // READ IN THE {CPAS}\DBASE\NEWLIC.X FILE.
//  //
//  strcpy(fn_NEWLIC, fpath_OutputDir);
//  strcat(fn_NEWLIC, "\\");
//  strcat(fn_NEWLIC, LICFILE_NewLicInfo);
//  if (_access(fn_GoodSerialNumber,0) == -1)
//  {
//    // FILE DOES NOT EXIST; EXIT WITH AN ERROR.
//    goto err_snumCpasLicGenerate_General;
//  }
//  else
//  {
//    // FILE DOES EXIST; DO NOTHING.
//  }
//  fp = fopen(fn_NEWLIC, "r");
//  fscanf(fp, "%s", lfdt.Z_SERIALNUMBER);
//  fscanf(fp, "%s", lfdt.Z_USERNAME);
//  fscanf(fp, "%s", lfdt.Z_USERCOMPANY);
//  fclose(fp);
  //
  // READ IN THE SERIAL NUMBER, USER NAME, USER COMPANY DATA.
  //
  strcpy(lfdt.Z_SERIALNUMBER, spNumber);
  strcpy(lfdt.Z_USERNAME, spUserName);
  strcpy(lfdt.Z_USERCOMPANY, spUserCompany);
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "INPUT PARAMETERS TO snumCpasLicGenerate():\n");
    fprintf(fp, "fpath_OutputDir = `%s`\n", fpath_OutputDir);
    fprintf(fp, "lfdt.Z_SERIALNUMBER = `%s`\n", lfdt.Z_SERIALNUMBER);
    fprintf(fp, "lfdt.Z_USERNAME = `%s`\n", lfdt.Z_USERNAME);
    fprintf(fp, "lfdt.Z_USERCOMPANY = `%s`\n", lfdt.Z_USERCOMPANY);
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "Test point Z1\n");
    fclose(fp);
  }
  //
  // VERIFY THE SERIAL NUMBER.
  //
  RetVal = Internal_snumDecode(lfdt.Z_SERIALNUMBER, iSnumPlain);
  //RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    if ( DEBUGMODE )
    {
      fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
      fprintf(fp, "Test point Z1a\n");
      fclose(fp);
    }
    // SERIAL NUMBER IS INVALID; EXIT WITH AN ERROR.
    goto err_snumCpasLicGenerate_General;
  }
  else
  {
    if ( DEBUGMODE )
    {
      fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
      fprintf(fp, "Test point Z1b\n");
      fclose(fp);
    }
    // SERIAL NUMBER IS VALID; DO NOTHING.
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "Test point Z2\n");
    fclose(fp);
  }
  // 
  // THIS IS WHERE THE FIELDS Z_PROGRAMKEY, Z_EXPIRATIONDATE,
  // Z_RELEASETYPE, AND Z_VERSIONTYPE ARE SET BASED ON THE Z_SERIALNUMBER FIELD.
  // NOTE THAT Z_VERSIONCODE IS FORCED TO "1.0" BECAUSE IT IS
  // NO LONGER USED.
  // 
  // NOTE THAT NO UPDATES TO CPAS.LIC ARE PERFORMED ANYMORE.
  // NOW, THE FILE IS REBUILT FROM SCRATCH EACH TIME THE -CREATE_FILE
  // OPTION IS RUN.
  // 
  strcpy(ThisSerialNumber, lfdt.Z_SERIALNUMBER);
  strcpy(ForceAll_Z_VERSIONCODE, "1.0");    // FORCE ALL TO "1.0"; THIS FIELD IS NO LONGER USED.
  RetVal = snumGetVersionType(ThisSerialNumber);
  switch (RetVal)
  {
    case 1:   // ALPHA VERSION.
      strcpy(ForceAll_Z_RELEASETYPE, "ALPHA");
      break;
    case 2:   // BETA VERSION.
      strcpy(ForceAll_Z_RELEASETYPE, "BETA");
      break;
    case 3:   // STANDARD VERSION.
      strcpy(ForceAll_Z_RELEASETYPE, "STANDARD");
      break;
    default:
      // SERIAL NUMBER IS INVALID; EXIT WITH AN ERROR.
      goto err_snumCpasLicGenerate_General;
  }
  RetVal = snumIsExpirationPresent(ThisSerialNumber);
  if (RetVal == 0)
  {
    // NO EXPIRATION DATE PRESENT.
    strcpy(ForceAll_Z_EXPIRATIONDATE, "NEVER");
    strcpy(ForceAll_Z_VERSIONTYPE, "VER_WONT_EXPIRE");
  }
  else
  {
    // EXPIRATION DATE IS PRESENT.
    iExpiresDay = snumGetExpirationDay(ThisSerialNumber);
    if (iExpiresDay == 0) 
    {
      // SERIAL NUMBER IS INVALID; EXIT WITH AN ERROR.
      goto err_snumCpasLicGenerate_General;
    }
    iExpiresMonth = snumGetExpirationMonth(ThisSerialNumber);
    if (iExpiresMonth == 0) 
    {
      // SERIAL NUMBER IS INVALID; EXIT WITH AN ERROR.
      goto err_snumCpasLicGenerate_General;
    }
    iExpiresYear = snumGetExpirationYear(ThisSerialNumber);
    sprintf(ForceAll_Z_EXPIRATIONDATE, "%d/%d/%d",
            iExpiresMonth,
            iExpiresDay,
            iExpiresYear);
    //sprintf(ForceAll_Z_EXPIRATIONDATE, "%d,%d,%d",
    //        iExpiresMonth,
    //        iExpiresDay,
    //        iExpiresYear);
    strcpy(ForceAll_Z_VERSIONTYPE, "VER_WILL_EXPIRE");
  }
  //
  // THIS IS WHERE THE LIST OF RECOGNIZED MODULES IS GENERATED.
  //

  rmt[0].idx = 0;
  strcpy(rmt[0].Z_PROGRAMKEY, "ADS");
  
  rmt[1].idx = 1;
  strcpy(rmt[1].Z_PROGRAMKEY, "ASAP");
  
  rmt[2].idx = 2;
  strcpy(rmt[2].Z_PROGRAMKEY, "STEPP");
  
  rmt[3].idx = 3;
  strcpy(rmt[3].Z_PROGRAMKEY, "FAVOR");
  
  rmt[4].idx = 4;
  strcpy(rmt[4].Z_PROGRAMKEY, "ADOX");
  
  rmt[5].idx = 20;
  strcpy(rmt[5].Z_PROGRAMKEY, "MFB");
  
  rmt[6].idx = 5;
  strcpy(rmt[6].Z_PROGRAMKEY, "STEPP2");
  
  rmt[7].idx = 6;
  strcpy(rmt[7].Z_PROGRAMKEY, "SCENE");
  
  rmt[8].idx = 8;
  strcpy(rmt[8].Z_PROGRAMKEY, "CPASFRONT");
  
  rmt[9].idx = 9;
  strcpy(rmt[9].Z_PROGRAMKEY, "CATREAC");

  rmt[10].idx = 10;
  strcpy(rmt[10].Z_PROGRAMKEY, "FAME");

  rmt[11].idx = 11;
  strcpy(rmt[11].Z_PROGRAMKEY, "BIOFILT");

  rmt[12].idx = 12;
  strcpy(rmt[12].Z_PROGRAMKEY, "IONEXDS");

  rmt_idx_low = 0;
  rmt_idx_high = 12;
  
  //
  // THIS IS WHERE THE LIST OF PURCHASED MODULES IS GENERATED.
  //
  num_pkdt_new = 0;
  pkdt_new_idx_low = 0;
  pkdt_new_idx_high = -1;
  for (i=rmt_idx_low; i<=rmt_idx_high; i++)
  {
    if (snumIsModulePurchased(ThisSerialNumber, rmt[i].idx) == 1)
    {
      num_pkdt_new ++;
      pkdt_new_idx_high ++;
      strcpy(pkdt_new[pkdt_new_idx_high].Z_PROGRAMKEY, rmt[i].Z_PROGRAMKEY);
      strcpy(pkdt_new[pkdt_new_idx_high].Z_EXPIRATIONDATE, ForceAll_Z_EXPIRATIONDATE);
      strcpy(pkdt_new[pkdt_new_idx_high].Z_RELEASETYPE, ForceAll_Z_RELEASETYPE);
      strcpy(pkdt_new[pkdt_new_idx_high].Z_VERSIONCODE, ForceAll_Z_VERSIONCODE);
      strcpy(pkdt_new[pkdt_new_idx_high].Z_VERSIONTYPE, ForceAll_Z_VERSIONTYPE);
    }
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "CONTENTS OF pkdt_new[] ARRAY:\n");
    fprintf(fp, "pkdt_new_idx_high = %d\n", pkdt_new_idx_high);
    for (i=pkdt_new_idx_low; i<=pkdt_new_idx_high; i++)
    {
      fprintf(fp, "    i = %d\n", i);
      fprintf(fp, "    pkdt_new[i].Z_PROGRAMKEY = `%s`\n", pkdt_new[i].Z_PROGRAMKEY);
      fprintf(fp, "    pkdt_new[i].Z_EXPIRATIONDATE = `%s`\n", pkdt_new[i].Z_EXPIRATIONDATE);
      fprintf(fp, "    pkdt_new[i].Z_RELEASETYPE = `%s`\n", pkdt_new[i].Z_RELEASETYPE);
      fprintf(fp, "    pkdt_new[i].Z_VERSIONCODE = `%s`\n", pkdt_new[i].Z_VERSIONCODE);
      fprintf(fp, "    pkdt_new[i].Z_VERSIONTYPE = `%s`\n", pkdt_new[i].Z_VERSIONTYPE);
    }
    fclose(fp);
  }
  //
  // GENERATE ENCRYPTION KEY.
  //
  if (LicFile_GenerateEncryptionKey() == 0)
  {
    // COULD NOT GENERATE ENCRYPTION KEY; EXIT WITH AN ERROR.
    goto err_snumCpasLicGenerate_General;
  }
  //
  // DELETE THE LICENSE FILE IF IT CURRENTLY EXISTS.
  //
  strcpy(fn_CPASLIC, fpath_OutputDir);
  strcat(fn_CPASLIC, "\\");
  strcat(fn_CPASLIC, LICFILE_LicName);
  if (_access(fn_CPASLIC,0) == -1)
  {
    // FILE DOES NOT EXIST; DO NOTHING.
  }
  else
  {
    // FILE EXISTS; DELETE IT.
    if (_unlink(fn_CPASLIC) == -1)
    {
      // UNABLE TO DELETE FILE!  EXIT WITH ERROR.
      goto err_snumCpasLicGenerate_General;
    }
  }
  //
  // OUTPUT LICENSE FILE.
  //
  filesize_CPASLIC = 1000 + (num_pkdt_new) * 1000 + 374;
  if (LicFile_Create(fn_CPASLIC, filesize_CPASLIC) == 0)
  {
    // COULD NOT CREATE NEW LICENSE FILE; EXIT WITH AN ERROR.
    goto err_snumCpasLicGenerate_General;
  }
  // IMPORTANT STEP!!  SET DATE/TIME TO "NEVER" STRINGS.
  strcpy(lfdt.ZZ_LASTEXECUTIONDATE, LICFILE_DATE_NEVER);
  strcpy(lfdt.ZZ_LASTEXECUTIONTIME, LICFILE_DATE_NEVER);
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "Test point AA1\n");
    fclose(fp);
  }
  // UPDATE PROGRAM KEY COUNT.
  lfdt.ZZ_NUMPROGRAMKEYS = num_pkdt_new;
  sprintf(DummyStr, "%d", lfdt.ZZ_NUMPROGRAMKEYS);
  offset_This_Data = lfdt_order_ZZ_NUMPROGRAMKEYS * 100;
offset_This_Data --;  ////////////////////////////////////////////////////////////////////////////////////////////////////////////
  if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0)
  {
    // UNABLE TO PUT STRING; EXIT WITH ERROR.
    goto err_snumCpasLicGenerate_General;
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "Test point AA1a\n");
    fclose(fp);
  }
  for (i=pkdt_new_idx_low; i<=pkdt_new_idx_high; i++)
  {
    AnyErrors = 0;
    This_Position = i;
    //offset_This_PK = 1000 + (This_Position - 1) * 1000;
    offset_This_PK = 1000 + (This_Position ) * 1000;
offset_This_PK --;  ////////////////////////////////////////////////////////////////////////////////////////////////////////////
if ( DEBUGMODE )
{
  fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
  fprintf(fp, "Test point AA1b; offset_ThisPK = %ld\n", offset_This_PK);
  fclose(fp);
}
    // Z_PROGRAMKEY.
    offset_This_Data = offset_This_PK + pkdt_order_Z_PROGRAMKEY * 100;
    strcpy(DummyStr,pkdt_new[i].Z_PROGRAMKEY);
    if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
    // Z_EXPIRATIONDATE.
    offset_This_Data = offset_This_PK + pkdt_order_Z_EXPIRATIONDATE * 100;
    strcpy(DummyStr,pkdt_new[i].Z_EXPIRATIONDATE);
    if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
    // Z_RELEASETYPE.
    offset_This_Data = offset_This_PK + pkdt_order_Z_RELEASETYPE * 100;
    strcpy(DummyStr,pkdt_new[i].Z_RELEASETYPE);
    if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
    // Z_VERSIONCODE.
    offset_This_Data = offset_This_PK + pkdt_order_Z_VERSIONCODE * 100;
    strcpy(DummyStr,pkdt_new[i].Z_VERSIONCODE);
    if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
    // Z_VERSIONTYPE.
    offset_This_Data = offset_This_PK + pkdt_order_Z_VERSIONTYPE * 100;
    strcpy(DummyStr,pkdt_new[i].Z_VERSIONTYPE);
    if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
    // CHECK FOR ERRORS.
    if (AnyErrors == 1)
    {
      // UNABLE TO PUT STRING; EXIT WITH ERROR.
      goto err_snumCpasLicGenerate_General;
    }
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "Test point AA2\n");
    fclose(fp);
  }
  //
  // OUTPUT THE HEADER INFO.
  //
  offset_This_PK = 0;
offset_This_PK --;  ////////////////////////////////////////////////////////////////////////////////////////////////////////////
  AnyErrors = 0;
  // Z_SERIALNUMBER.
  offset_This_Data = offset_This_PK + lfdt_order_Z_SERIALNUMBER * 100;
  strcpy(DummyStr,lfdt.Z_SERIALNUMBER);
  if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
  // Z_USERNAME.
  offset_This_Data = offset_This_PK + lfdt_order_Z_USERNAME * 100;
  strcpy(DummyStr,lfdt.Z_USERNAME);
  if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
  // Z_USERCOMPANY.
  offset_This_Data = offset_This_PK + lfdt_order_Z_USERCOMPANY * 100;
  strcpy(DummyStr,lfdt.Z_USERCOMPANY);
  if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
  // ZZ_LASTEXECUTIONDATE.
  offset_This_Data = offset_This_PK + lfdt_order_ZZ_LASTEXECUTIONDATE * 100;
  strcpy(DummyStr,lfdt.ZZ_LASTEXECUTIONDATE);
  if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
  // ZZ_LASTEXECUTIONTIME.
  offset_This_Data = offset_This_PK + lfdt_order_ZZ_LASTEXECUTIONTIME * 100;
  strcpy(DummyStr,lfdt.ZZ_LASTEXECUTIONTIME);
  if (LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, DummyStr) == 0) AnyErrors=1;
  // CHECK FOR ERRORS.
  if (AnyErrors == 1)
  {
    // UNABLE TO PUT STRING; EXIT WITH ERROR.
    goto err_snumCpasLicGenerate_General;
  }
  //
  // ATTEMPT TO CREATE THE FILE {WIN}\MTCHK.LIC;
  // IF UNABLE TO CREATE THE FILE, NO BIG DEAL.
  // 
  strcpy(fn_MTCHKLIC, spWinDir);
  strcat(fn_MTCHKLIC, "\\");
  strcat(fn_MTCHKLIC, LICFILE_ExtraCheckFile);
  AnyErrors = 0;
  if (LicFile_Create(fn_MTCHKLIC, 1277) == 0) 
  {
    AnyErrors = 1;
  }
  else
  {
    //if (LicFile_PutEncryptedString(fn_MTCHKLIC, 671, LICFILE_ExtraCheckFile_Text) == 0)
    if (LicFile_PutEncryptedString(fn_MTCHKLIC, 671 - 1, LICFILE_ExtraCheckFile_Text) == 0)
    {
      AnyErrors = 1;
    }
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumCpasLicGenerate, "a+");
    fprintf(fp, "Test point AA3\n");
    fclose(fp);
  }
  if (AnyErrors == 1) 
  {
    // UNABLE TO CREATE MTCHK.LIC FILE; SO LONG AS THIS FILE REMAINS
    // UN-WRITEABLE IN THE FUTURE, THE CPASCHK.EXE MODULE WILL CONTINUE
    // TO INDICATE THAT THE LICENSING IS VALID.
  }
  else
  {
    // THE MTCHK.LIC FILE WAS GENERATED.  IF THIS FILE REMAINS WRITEABLE
    // IN THE FUTURE, THE CPASCHK.EXE MODULE WILL CONTINUE TO INDICATE
    // THAT THE LICENSING IS VALID.
  }
  //
  // EXIT WITH SUCCESS INDICATOR.
  //
  return(1);
err_snumCpasLicGenerate_General:
  //
  // EXIT WITH FAILURE INDICATOR.
  //
  return(0);
}


// PROTOTYPE:
//     int LicFile_GenerateEncryptionKey()
// INPUTS:
//     NONE
// OUTPUTS:
//     NONE
// PURPOSE:
//     GENERATE ENCRYPTION KEY.
// CALLED FROM:
//     LICGEN.CPP
// RETURNS:
//     0 = FAILED!
//     1 = SUCCEEDED.
//
int LicFile_GenerateEncryptionKey()
{
int i;
int ThisVal;
  for (i=0; i<=255; i++)
  {
    ThisVal = (i * 3) % 256;
    LicFile_keyidx[0][i] = ThisVal;
    LicFile_keyidx[1][ThisVal] = i;
  }
  // 'STRICTLY SPEAKING, THIS ISN'T REALLY ENCRYPTION; IT'S MORE LIKE
  // 'A JOKE.  AS THEY SAY, "LOCKS ONLY KEEP HONEST PEOPLE OUT."
  // 'THE EASIEST WAY TO BREAK THROUGH A SERIAL NUMBER CHECK IS TO MODIFY
  // 'THE PROGRAM SO IT DOESN'T PERFORM THE CHECK.  THAT WOULD MAKE
  // 'ANY ENCRYPTION A MOOT POINT ANYWAY.
  return(1);
}


// PROTOTYPE:
//     int LicFile_Create(char *fn_CPASLIC, long filesize_CPASLIC)
// INPUTS:
//     fn_CPASLIC = NAME OF CPAS.LIC FILE.
//     filesize_CPASLIC = APPROXIMATE INITIAL SIZE OF CREATED FILE (BYTES).
// OUTPUTS:
//     NONE
// PURPOSE:
//     GENERATE ENCRYPTION KEY.
// CALLED FROM:
//     LICGEN.CPP
// RETURNS:
//     0 = FAILED!
//     1 = SUCCEEDED.
//
int LicFile_Create(char *fn_CPASLIC, long filesize_CPASLIC)
{
int fh;
long n;
long i;
unsigned BytesWritten;
int ThisVal;
  n = filesize_CPASLIC / 2;
  n ++;
  n ++;
  //
  // SEED THE RANDOM NUMBER GENERATOR.
  //
  srand( (unsigned)time( NULL ) );
  //
  // OPEN THE FILE FOR OUTPUT.
  //
  fh = _open(fn_CPASLIC, _O_RDWR | _O_CREAT | _O_BINARY, _S_IREAD | _S_IWRITE);
  if (fh == -1)
  {
    // THE _open() FUNCTION FAILED; EXIT WITH ERROR.
    return(0);
  }
  for (i=0; i<n; i++)
  {
    ThisVal = (rand());
    //BytesWritten = _write(fh, &ThisVal, sizeof(ThisVal));
    BytesWritten = _write(fh, &ThisVal, 2);
    if (BytesWritten == -1)
    {
      // THE _write() FUNCTION FAILED; CLOSE FILE, EXIT WITH ERROR.
      _close(fh);
      return(0);
    }
  }
  _close(fh);
  return(1);
}


// PROTOTYPE:
//     int LicFile_PutEncryptedString(char *fn_This,long pos_This,char *str_This)
// INPUTS:
//     fn_This = NAME OF FILE TO MODIFY.
//     pos_This = POSITION WITHIN FILE TO PUT STRING (BYTES FROM START).
//     str_This = STRING TO PUT.
// OUTPUTS:
//     NONE
// PURPOSE:
//     PUT AN ENCRYPTED STRING INTO A FILE.
// CALLED FROM:
//     LICGEN.CPP
// RETURNS:
//     0 = FAILED!
//     1 = SUCCEEDED.
//
int LicFile_PutEncryptedString(char *fn_This,long pos_This,char *str_This)
{
int UseStrSize;
char UseStr[LICFILE_STRINGSIZE_MAX+1];
char UseStr_Enc[LICFILE_STRINGSIZE_MAX+1];
char UseStrSize_String[10];
char UseStrSize_String_Enc[10];
int i;
unsigned BytesWritten;
int fh;
  //
  // PREPARE STRING FOR OUTPUT.
  //
  if (strlen(str_This) <= LICFILE_STRINGSIZE_MAX)
  {
    UseStrSize = strlen(str_This);
  }
  else
  {
    UseStrSize = LICFILE_STRINGSIZE_MAX;
  }
  for (i=0; i<UseStrSize; i++)
  {
    UseStr[i] = str_This[i];
  }
  UseStr[UseStrSize] = '\0';
  sprintf(UseStrSize_String, "%d", UseStrSize);
  if (LicFile_EncryptString(UseStrSize_String, UseStrSize_String_Enc) == 0)
  {
    // THE LicFile_EncryptString() FUNCTION FAILED; EXIT WITH ERROR.
    return(0);
  }
  if (LicFile_EncryptString(UseStr, UseStr_Enc) == 0)
  {
    // THE LicFile_EncryptString() FUNCTION FAILED; EXIT WITH ERROR.
    return(0);
  }
  //
  // OPEN THE FILE FOR MODIFICATION.
  //
  fh = _open(fn_This, _O_RDWR | _O_CREAT | _O_BINARY, _S_IREAD | _S_IWRITE);
  if (fh == -1)
  {
    // THE _open() FUNCTION FAILED; EXIT WITH ERROR.
    return(0);
  }
  if (_lseek(fh, pos_This, SEEK_SET) == -1L)
  {
    // THE _lseek() FUNCTION FAILED; CLOSE FILE, EXIT WITH ERROR.
    _close(fh);
    return(0);
  }
  BytesWritten = _write(fh, &UseStrSize_String_Enc[0], 3);
  if (BytesWritten == -1)
  {
    // THE _write() FUNCTION FAILED; CLOSE FILE, EXIT WITH ERROR.
    _close(fh);
    return(0);
  }
  BytesWritten = _write(fh, &UseStr_Enc[0], UseStrSize);
  if (BytesWritten == -1)
  {
    // THE _write() FUNCTION FAILED; CLOSE FILE, EXIT WITH ERROR.
    _close(fh);
    return(0);
  }
  _close(fh);
  return(1);
}


// PROTOTYPE:
//     int LicFile_EncryptString(char *str_Plain,char *str_Enc)
// INPUTS:
//     str_Plain = STRING CONTAINING PLAIN-TEXT.
// OUTPUTS:
//     str_Enc = STRING CONTAINING ENCRYPTED-TEXT.
// PURPOSE:
//     ENCRYPT A STRING.
// CALLED FROM:
//     LICGEN.CPP
// RETURNS:
//     0 = FAILED!
//     1 = SUCCEEDED.
//
int LicFile_EncryptString(char *str_Plain,char *str_Enc)
{
int i;
int TempVal;
int str_Plain_Len;
  str_Plain_Len = strlen(str_Plain);
  for (i=0; i<str_Plain_Len; i++)
  {
    TempVal = str_Plain[i];
    TempVal = LicFile_keyidx[0][TempVal] ^ (i+1);
    str_Enc[i] = TempVal;
  }
  str_Enc[str_Plain_Len] = '\0';
  return(1);
}



