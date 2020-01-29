///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    PROGRAM:
//        CPASLIB.DLL
//
//    PURPOSE:
//        GENERATE AND VERIFY CPAS (CLEAN PROCESS ADVISORY SYSTEM)
//        SERIAL NUMBERS FOR PURCHASED/BETA COPIES OF THE SOFTWARE.
//
//    PLATFORM:
//        DYNAMICALLY LINKED LIBRARY (DLL) FOR WINDOWS 95 OR WINDOWS NT4.
//
//    IMPORTED ROUTINES:
//        KERNEL32
//
//    EXPORTED ROUTINES:
//        ALL ROUTINES PROTOTYPED WITH THE FOLLOWING KEYWORDS:
//                extern "C" int __stdcall
//
//    BUILD HISTORY:
//        1998.10.13. FIRST VERSION CONSTRUCTED, EJOMAN.
//        1998.11.11. MODIFIED TO HANDLED CPAS.LIC CREATION, EJOMAN.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// COMMON C/C++ FUNCTION PROTOTYPES.
#include <stdio.h>
#include <string.h>

// FUNCTION PROTOTYPES.
#include "cpaslib.h"
#include "internal.h"
#include "licgen.h"





//---------------------------------------------------------------------------------------------------------------------------------
//--------------  BASIC SERIAL NUMBER VERIFICATION  -------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------------------

// PROTOTYPE:
//     extern "C" int __stdcall snumVerify(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     VERIFIES VALIDITY OF SERIAL NUMBER.
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = SERIAL NUMBER IS INVALID!
//     1 = SERIAL NUMBER IS VALID
//
extern "C" int __stdcall snumVerify(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  else
  {
    // SERIAL NUMBER IS VALID.
    return(1);
  }
}


//---------------------------------------------------------------------------------------------------------------------------------
//--------------  SERIAL NUMBER GENERATION  ---------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------------------

// PROTOTYPE:
//     extern "C" int __stdcall snumGenerate(char *spNumber, int iModules[], int iVersionType, int iExpires, 
//                                 int iExpireDay, int iExpireMonth, int iExpireYear, int iCheck)
// INPUTS:
//     iModules[] = ARRAY ELEMENTS EQUAL TO 0 WERE NOT PURCHASED, EQUAL TO 1 WERE PURCHASED
//     iVersionType = VERSION TYPE: 1=ALPHA, 2=BETA, 3=STANDARD
//     iExpires = WHETHER VERSION EXPIRES: 0=NO, 1=YES
//     iExpiresDay = 0=DOES NOT EXPIRE, 1-31=EXPIRATION DAY
//     iExpiresMonth = 0=DOES NOT EXPIRE, 1-12=EXPIRATION MONTH
//     iExpiresYear = 0=DOES NOT EXPIRE, 0-31=EXPIRATION YEAR MINUS 1990, e.g. 0=1990, 31=2021
//     longInternalSnum = INTERNAL SERIAL NUMBER; VALID RANGE = 1-1,048,575
//     iCheck = SIMPLE SECURITY CHECK; IF NOT SET TO 13892, EXIT IMMEDIATELY WITHOUT EXECUTING.
// OUTPUTS:
//     spNumber = SERIAL NUMBER, SEE BELOW FOR FORMAT
//
// FORMAT FOR SERIAL NUMBERS (spNumber):
//
//   The serial number entered by the user is formatted as follows:
//
//       CA****-*****-*****-*****-*****
//       where *=one of the following: "0123456789ABCDEFGHJKLMNPQRTUVWXY"
//       Note that serial numbers are case-insensitive.
//       Note that I,O,S,Z are not included because they could
//       be mistaken for the numeric characters 1,0,5,2.
//
//   Each character (excluding the hyphen characters) codes for 32 
//   possible states because it may hold any of 32 different characters.
//   After the serial number entered by the user is decoded, the decoded
//   bits are used to store the following information:
//
//   BITS     DESCRIPTION
//   =======  ======================================================================
//   0        Whether AdDesignS was purchased (0=no, 1=yes)
//   1        Whether ASAP was purchased (0=no, 1=yes)
//   2        Whether StEPP was purchased (0=no, 1=yes)
//   4-49     Reserved
//   50-51    Version type (0=Reserved, 1=Alpha, 2=Beta, 3=Standard)
//   52       Whether version expires (0=No expiration date, 1=Expiration date exists)
//   53-57    Expiration day (0=None, 1-31=Day)
//   58-61    Expiration month (0=None, 1-12=Month)
//   62-66    Expiration year-1990 (0-31=Year-1990, e.g. 0=1990, 31=2021)
//   67-87    Internal serial number (1-1,048,575 are valid)
//   88-100   Checksum to ensure serial number is valid
//   101-104  Reserved
//
// PURPOSE:
//     GENERATES A NEW SERIAL NUMBER.
// CALLED FROM:
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = ONE OR MORE INPUT PARAMETERS WERE IMPROPERLY SET!
//     1 = SERIAL NUMBER SUCCESSFULLY GENERATED AND RETURNED IN spNumber PARAMETER.
//
extern "C" int __stdcall snumGenerate(
    char *spNumber, 
    int iModules[], 
    int iVersionType, 
    int iExpires, 
    int iExpiresDay, 
    int iExpiresMonth, 
    int iExpiresYear, 
    long longInternalSnum, 
    int iCheck)
{
FILE *fp;
int iSnumPlain[21];
int iSnumXord[21];
int iSnumXlatd[21];
int iSnumFinal[30];
char sSnumFinal[31];    // EXTRA char TO HOLD THE NULL CHARACTER.
int i;
int j;
int k;
int ThisBit;
int ThisVersionType;
int ThisExpires;
int ThisExpiresDay;
int ThisExpiresMonth;
int ThisExpiresYear;
long ThisInternalSnum;
int ThisChecksum;
int Temp1;
int Temp2;
int Temp3;
int Temp4;
int Temp5;
int iXorConstants[21];
int iScatterConstants[21];
int iCheckInteger;
  // VERIFY THAT PROPER iCheck PARAMETER WAS SENT (SECURITY PRECAUTION).
  if ( iCheck != ICHECK_CORRECT_VALUE )
  {
    goto err_snumGenerate_InvalidInput;
  }
  if ( DEBUGMODE )
  {
    // OVERWRITE OLD FILE, IF ANY.
    fp = fopen(FNDEBUG_snumGenerate, "w");
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "INPUT PARAMETERS TO snumGenerate():\n");
    fprintf(fp, "iVersionType = %d\n", iVersionType);
    fprintf(fp, "iExpires = %d\n", iExpires);
    fprintf(fp, "iExpiresDay, = %d\n", iExpiresDay);
    fprintf(fp, "iExpiresMonth = %d\n", iExpiresMonth);
    fprintf(fp, "iExpiresYear = %d\n", iExpiresYear);
    fprintf(fp, "longInternalSnum = %ld\n", longInternalSnum);
    fprintf(fp, "iCheck = %d\n", iCheck);
    for (i=0; i<50; i++)
    {
      fprintf(fp, "iModules[%d] = %d\n", i, iModules[i]);
    }
    fclose(fp);
  }
  //
  // VALIDATE INPUT PARAMETERS.
  //
  switch (iVersionType)
  {
    case VERSIONTYPE_ALPHA: ThisVersionType = iVersionType; break;       
    case VERSIONTYPE_BETA: ThisVersionType = iVersionType; break;       
    case VERSIONTYPE_STANDARD: ThisVersionType = iVersionType; break;
    default: goto err_snumGenerate_InvalidInput;
  }
  switch (iExpires)
  {
    case EXPIRES_NO: ThisExpires = iExpires; break;       
    case EXPIRES_YES: ThisExpires = iExpires; break;       
    default: goto err_snumGenerate_InvalidInput;
  }
  if (iExpires == EXPIRES_YES)
  {
    if ((iExpiresDay>=1) && (iExpiresDay<=31))
      ThisExpiresDay = iExpiresDay;
    else
      goto err_snumGenerate_InvalidInput;
    if ((iExpiresMonth>=1) && (iExpiresMonth<=12))
      ThisExpiresMonth = iExpiresMonth;
    else
      goto err_snumGenerate_InvalidInput;
    if ((iExpiresYear>=1990) && (iExpiresYear<=2021))
      ThisExpiresYear = iExpiresYear;
    else
      goto err_snumGenerate_InvalidInput;
  }
  else
  {
    ThisExpiresDay = 0;
    ThisExpiresMonth = 0;
    ThisExpiresYear = 1990;     // STORED AS A ZERO VALUE.
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- CONTENTS OF ThisExpires*:\n");
    fprintf(fp, "ThisExpiresDay = %d\n", ThisExpiresDay);
    fprintf(fp, "ThisExpiresMonth = %d\n", ThisExpiresMonth);
    fprintf(fp, "ThisExpiresYear = %d\n", ThisExpiresYear);
    fclose(fp);
  }
  if ((longInternalSnum>=1) && (longInternalSnum<=1048575))
    ThisInternalSnum = longInternalSnum;
  else
    goto err_snumGenerate_InvalidInput;

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  GENERATE THE iSnumPlain[] ARRAY  ////////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  //
  // GENERATE THE iSnumPlain[] ARRAY (BITS 0 TO 86).
  //
  // COPY iModules[] TO BITS [0-49].
  for (i=0; i<=20; i++)
  {
    iSnumPlain[i] = 0;
  }
  for (i=0; i<=9; i++)
  {
    ThisBit = 1;
    for (j=0; j<=4; j++)
    {
      k = 5*i + j;
      iSnumPlain[i] += (ThisBit*((iModules[k]==1)?1:0));
      ThisBit = ThisBit << 1;
    }
  }
  // COPY ThisVersionType TO (INT 10 BITS 1-2)[50-51].
  iSnumPlain[10] += ThisVersionType;
  // COPY ThisExpires TO (INT 10 BIT 3)[52].
  iSnumPlain[10] += ThisExpires * 4;
  // COPY ThisExpiresDay SPLIT ACROSS (INT 10 BITS 4-5)[53-54] AND (INT 11 BITS 1-3)[55-57].
  Temp1 = (ThisExpiresDay & (1+2)) << 3;
  Temp2 = (ThisExpiresDay & (4+8+16)) >> 2;
  iSnumPlain[10] += Temp1;
  iSnumPlain[11] += Temp2;
  // COPY ThisExpiresMonth SPLIT ACROSS (INT 11 BITS 4-5)[58-59] AND (INT 12 BITS 1-2)[60-61].
  Temp1 = (ThisExpiresMonth & (1+2)) << 3;
  Temp2 = (ThisExpiresMonth & (4+8)) >> 2;
  iSnumPlain[11] += Temp1;
  iSnumPlain[12] += Temp2;
  // COPY ThisExpiresYear SPLIT ACROSS (INT 12 BITS 3-5)[62-64] AND (INT 13 BITS 1-2)[65-66].
  Temp1 = ((ThisExpiresYear-1990) & (1+2+4)) << 2;
  Temp2 = ((ThisExpiresYear-1990) & (8+16)) >> 3;
  iSnumPlain[12] += Temp1;
  iSnumPlain[13] += Temp2;
  // COPY ThisInternalSnum SPLIT ACROSS (INT 13 BITS 3-5)[67-69] AND 
  // (INT 14 BITS 1-5)[70-74] AND (INT 15 BITS 1-5)[75-79] AND
  // (INT 16 BITS 1-5)[80-84] AND (INT 17 BITS 1-3)[85-87].
  Temp1 = (ThisInternalSnum & (1+2+4)) << 2;
  Temp2 = (ThisInternalSnum & (8+16+32+64+128)) >> 3;
  Temp3 = (ThisInternalSnum & (256+512+1024+2048+4096)) >> (3+5);
  Temp4 = (ThisInternalSnum & (8192+16384+32768+65536+131072)) >> (3+5+5);
  Temp5 = (ThisInternalSnum & (262144+524288+1048576)) >> (3+5+5+5);
  iSnumPlain[13] += Temp1;
  iSnumPlain[14] += Temp2;
  iSnumPlain[15] += Temp3;
  iSnumPlain[16] += Temp4;
  iSnumPlain[17] += Temp5;
  //
  // CALCULATE CHECKSUM (ThisChecksum).
  //
  ThisChecksum = 0;
  for (i=0; i<=17; i++)
  {
    j = 18 - i;
    ThisChecksum += (j)*(iSnumPlain[i]);
  }
  //
  // GENERATE THE iSnumPlain[] ARRAY (BITS 88 TO 100).
  //
  // COPY ThisChecksum SPLIT ACROSS (INT 17 BITS 4-5)[88-89] AND 
  // (INT 18 BITS 1-5)[90-94] AND (INT 19 BITS 1-5)[95-99] AND
  // (INT 20 BIT 1)[100].
  Temp1 = (ThisChecksum & (1+2)) << 3;
  Temp2 = (ThisChecksum & (4+8+16+32+64)) >> 2;
  Temp3 = (ThisChecksum & (128+256+512+1024+2048)) >> (2+5);
  Temp4 = (ThisChecksum & (4096)) >> (2+5+5);
  iSnumPlain[17] += Temp1;
  iSnumPlain[18] += Temp2;
  iSnumPlain[19] += Temp3;
  iSnumPlain[20] += Temp4;
  iSnumPlain[21] += Temp5;
  //
  // NOTE: BITS 101 TO 104 ARE RESERVED.
  //

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  XOR iSnumPlain[] TO GET iSnumXord[]  ////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  Internal_GetXorConstants(iXorConstants);
  for (i=0; i<=20; i++)
  {
    iSnumXord[i] = iSnumPlain[i] ^ iXorConstants[i];
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  TRANSLATE iSnumXord[] TO BASE-32 iSnumXlatd[]  //////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  for (i=0; i<=20; i++)
  {
    iSnumXlatd[i] = Internal_NormalToBase32(iSnumXord[i]);
  }
  
  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  SCATTER iSnumXlatd[] INTO iSnumFinal[]  /////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  // CLEAR iSnumFinal[].
  for (i=0; i<=29; i++)
  {
    iSnumFinal[i] = 0;
  }
  // GET SCATTER CONSTANTS AND SCATTER INTO iSnumFinal[].
  Internal_GetScatterConstants(iScatterConstants);
  for (i=0; i<=20; i++)
  {
    iSnumFinal[iScatterConstants[i]] = iSnumXlatd[i];
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  PERFORM ADDITIONAL STEPS WITH iSnumFinal[]  /////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  // SET FIRST TWO INTEGERS TO "CA".
  iSnumFinal[0] = 67;   // "C"
  iSnumFinal[1] = 65;   // "A"
  // PLACE RANDOM CHARACTER INTO POSITION 26.
  iSnumFinal[26] = 65;    // to do: add random function
  // PLACE HYPHEN CHARACTERS INTO POSITIONS 6,12,18,24.
  iSnumFinal[6] = 45;     // "-"
  iSnumFinal[12] = 45;    // "-"
  iSnumFinal[18] = 45;    // "-"
  iSnumFinal[24] = 45;    // "-"
  // CALCULATE CHECK INTEGER FOR POSITIONS 4 AND 19.
  iCheckInteger = 0;
  for (i=0; i<=20; i++)
  {
    iCheckInteger ^= iSnumFinal[iScatterConstants[i]];
  }
  Temp1 = (iCheckInteger & (1+2+4+8));
  Temp1 += 16;
  Temp2 = (iCheckInteger & (16+32+64+128)) >> 3;
  Temp2 += 1;
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- CONTENTS OF Temp* (POINT A1):\n");
    fprintf(fp, "Temp1 = %d\n", Temp1);
    fprintf(fp, "Temp2 = %d\n", Temp2);
    fclose(fp);
  }
  iSnumFinal[4] = Internal_NormalToBase32(Temp1);
  iSnumFinal[19] = Internal_NormalToBase32(Temp2);

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  CONVERT iSnumFinal[] TO spNumber  ///////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  for (i=0; i<=29; i++)
  {
    sSnumFinal[i] = iSnumFinal[i];
  }
  sSnumFinal[30] = 0;
  for (i=0; i<=30; i++)
  {
    spNumber[i] = sSnumFinal[i];
  }

  //
  // OUTPUT DEBUG MESSAGES IF DEBUGMODE IS ON (1).
  //
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- CONTENTS OF iSnumPlain[]:\n");
    for (i=0; i<=20; i++)
    {
      fprintf(fp, "iSnumPlain[%d] = %d\n", i, iSnumPlain[i]);
    }
    fprintf(fp, "\n---------- CONTENTS OF iSnumXord[]:\n");
    for (i=0; i<=20; i++)
    {
      fprintf(fp, "iSnumXord[%d] = %d\n", i, iSnumXord[i]);
    }
    fprintf(fp, "\n---------- CONTENTS OF iSnumXlatd[]:\n");
    for (i=0; i<=20; i++)
    {
      fprintf(fp, "iSnumXlatd[%d] = %d\n", i, iSnumXlatd[i]);
    }
    fprintf(fp, "\n---------- CONTENTS OF iSnumFinal[]:\n");
    for (i=0; i<=29; i++)
    {
      fprintf(fp, "iSnumFinal[%d] = %d\n", i, iSnumFinal[i]);
    }
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- CONTENTS OF iXorConstants[]:\n");
    for (i=0; i<=20; i++)
    {
      fprintf(fp, "iXorConstants[%d] = %d\n", i, iXorConstants[i]);
    }
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- CONTENTS OF iScatterConstants[]:\n");
    for (i=0; i<=20; i++)
    {
      fprintf(fp, "iScatterConstants[%d] = %d\n", i, iScatterConstants[i]);
    }
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- TEST OF Internal_NormalToBase32():\n");
    for (i=0; i<=31; i++)
    {
      fprintf(fp, "X = Internal_NormalToBase32(%d) = %c; Internal_Base32ToNormal(X) = %d\n", 
              i, 
              Internal_NormalToBase32(i),
              Internal_Base32ToNormal(Internal_NormalToBase32(i)));
    }
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "\n---------- OUTPUT OF iCheckInteger:\n");
    fprintf(fp, "iCheckInteger = `%d`\n", iCheckInteger);
    fprintf(fp, "\n---------- FINAL OUTPUT OF spNumber[]:\n");
    fprintf(fp, "spNumber = `%s`\n", spNumber);
    fclose(fp);
  }

  // RETURN "SUCCESS" MESSAGE.
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "RETURN VALUE = 1\n");
    fclose(fp);
  }
  return(1);
  
err_snumGenerate_InvalidInput:
  // INVALID INPUT.
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumGenerate, "a+");
    fprintf(fp, "RETURN VALUE = 0\n");
    fclose(fp);
  }
  return(0);
}


//---------------------------------------------------------------------------------------------------------------------------------
//--------------  PURCHASING/VERSION-TYPE RELATED  --------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------------------

// PROTOTYPE:
//     extern "C" int __stdcall snumIsModulePurchased(char *spNumber, int iModule)
// INPUTS:
//     spNumber = SERIAL NUMBER
//     iModule = MODULE CODE (E.G. 1=AdDesignS, 2=ASAP, 3=StEPP, etc.)
// PURPOSE:
//     DETERMINES WHETHER OR NOT A GIVEN MODULE WAS PURCHASED.
// RETURNS:
//     0 = NO, THAT MODULE WAS NOT PURCHASED! (OR THE SERIAL NUMBER IS INVALID.)
//     1 = YES, THAT MODULE WAS PURCHASED
//
extern "C" int __stdcall snumIsModulePurchased(char *spNumber, int iModule)
{
int iSnumPlain[21];  
int RetVal;
int ThisBit;
int ThisInteger;
int ThisIntegerPos;
int ThisBitPos;
int ThisBitVal;
FILE *fp;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT MODULE CODE INFORMATION.
  //
  ThisIntegerPos = iModule / 5;
  ThisBitPos = iModule % 5;
  ThisInteger = iSnumPlain[ThisIntegerPos];
  ThisBitVal = 1 << (ThisBitPos);
  ThisBit = ((ThisInteger & ThisBitVal) ? (1) : (0));
  //
  // MISCELLANEOUS DEBUG STUFF.
  //
  if ( DEBUGMODE )
  {
    // OVERWRITE OLD FILE, IF ANY.
    fp = fopen(FNDEBUG_snumIsModulePurchased, "w");
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_snumIsModulePurchased, "a+");
    fprintf(fp, "INPUT PARAMETERS TO FNDEBUG_snumIsModulePurchased():\n");
    fprintf(fp, "spNumber = not output\n");
    fprintf(fp, "iModule = %d\n", iModule);
    fprintf(fp, "-------- OUTPUT VALUES:\n");
    fprintf(fp, "ThisIntegerPos = %d\n", ThisIntegerPos);
    fprintf(fp, "ThisBitPos = %d\n", ThisBitPos);
    fprintf(fp, "ThisInteger = %d\n", ThisInteger);
    fprintf(fp, "ThisBitVal = %d\n", ThisBitVal);
    fprintf(fp, "ThisBit = %d\n", ThisBit);
    fclose(fp);
  }
  //
  // RETURN WHETHER USER PURCHASED THAT MODULE.
  //
  if (ThisBit == 1)
  {
    // YES, THE USER DID PURCHASE THAT MODULE!
    return(1);
  }
  else
  {
    // NO, THE USER DID NOT PURCHASE THAT MODULE!
    return(0);
  }
}


// PROTOTYPE:
//     extern "C" int __stdcall snumGetVersionType(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     GETS TYPE OF VERSION (ALPHA, BETA, STANDARD).
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = INVALID SERIAL NUMBER!
//     1 = THIS IS AN ALPHA VERSION.
//     2 = THIS IS A BETA VERSION.
//     3 = THIS IS A STANDARD VERSION.
//
extern "C" int __stdcall snumGetVersionType(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
//FILE *fp;
int ThisData;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT VERSION TYPE.
  //
  ThisData = iSnumPlain[10] & (1+2);
  //
  // RETURN VERSION TYPE.
  //
  switch (ThisData)
  {
    case 1: 
      // ALPHA VERSION.
      return(1);
      break;
    case 2: 
      // BETA VERSION.
      return(2);
      break;
    case 3: 
      // STANDARD VERSION.
      return(3);
      break;
    default:
      // INVALID VERSION TYPE!
      return(0);
      break;
  }
  ////
  //// MISCELLANEOUS DEBUG STUFF.
  ////
  //if ( DEBUGMODE )
  //{
  //  // OVERWRITE OLD FILE, IF ANY.
  //  fp = fopen(FNDEBUG_snumIsModulePurchased, "w");
  //  fclose(fp);
  //}
}


//---------------------------------------------------------------------------------------------------------------------------------
//--------------  EXPIRATION RELATED  ---------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------------------

// PROTOTYPE:
//     extern "C" int __stdcall snumIsExpirationPresent(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     DETERMINES WHETHER OR NOT THE VERSION WILL EXPIRE.
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = NO, THE VERSION WILL NOT EVER EXPIRE (OR ELSE THE SERIAL NUMBER IS INVALID).
//     1 = YES, THE VERSION WILL EVENTUALLY EXPIRE.
//
extern "C" int __stdcall snumIsExpirationPresent(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
//FILE *fp;
int ThisData;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT EXPIRATION BIT.
  //
  ThisData = (iSnumPlain[10] & (4)) >> 2;
  //
  // RETURN VERSION TYPE.
  //
  switch (ThisData)
  {
    case 1: 
      // VERSION _WILL_ EXPIRE.
      return(1);
      break;
    default:
      // VERSION WILL _NOT_ EXPIRE.
      return(0);
      break;
  }
}


// PROTOTYPE:
//     extern "C" int __stdcall snumGetExpirationDay(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     DETERMINES EXPIRATION DAY.
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = THIS VERSION WILL NOT EVER EXPIRE!
//     NON-0 = THE DAY THE VERSION EXPIRES.
//
extern "C" int __stdcall snumGetExpirationDay(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
//FILE *fp;
int Temp[2];
int ThisData;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT EXPIRATION DAY BITS.
  //
  Temp[0] = (iSnumPlain[10] & (8+16)) >> 3;
  Temp[1] = (iSnumPlain[11] & (1+2+4)) << 2;
  ThisData = Temp[0] + Temp[1];
  //
  // RETURN EXPIRATION DAY.
  //
  return(ThisData);
  ////
  //// MISCELLANEOUS DEBUG STUFF.
  ////
  //if ( DEBUGMODE )
  //{
  //  // OVERWRITE OLD FILE, IF ANY.
  //  fp = fopen(FNDEBUG_snumIsModulePurchased, "w");
  //  fclose(fp);
  //}
}


// PROTOTYPE:
//     extern "C" int __stdcall snumGetExpirationMonth(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     DETERMINES EXPIRATION MONTH.
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = THIS VERSION WILL NOT EVER EXPIRE!
//     NON-0 = THE MONTH THE VERSION EXPIRES.
//
extern "C" int __stdcall snumGetExpirationMonth(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
//FILE *fp;
int Temp[2];
int ThisData;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT EXPIRATION MONTH BITS.
  //
  Temp[0] = (iSnumPlain[11] & (8+16)) >> 3;
  Temp[1] = (iSnumPlain[12] & (1+2)) << 2;
  ThisData = Temp[0] + Temp[1];
  //
  // RETURN EXPIRATION MONTH.
  //
  return(ThisData);
  ////
  //// MISCELLANEOUS DEBUG STUFF.
  ////
  //if ( DEBUGMODE )
  //{
  //  // OVERWRITE OLD FILE, IF ANY.
  //  fp = fopen(FNDEBUG_snumIsModulePurchased, "w");
  //  fclose(fp);
  //}
}


// PROTOTYPE:
//     extern "C" int __stdcall snumGetExpirationYear(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     DETERMINES EXPIRATION YEAR.
// CALLED FROM:
//     INSTALLSHIELD SCRIPT
//     CPASCHK VB5 PROGRAM
//     CPAS_SERIAL VB5 PROGRAM
// RETURNS:
//     0 = THIS VERSION WILL NOT EVER EXPIRE!
//     NON-0 = THE YEAR THE VERSION EXPIRES; e.g. the year 1999 is returned as 1999.
//
extern "C" int __stdcall snumGetExpirationYear(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
//FILE *fp;
int Temp[2];
int ThisData;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT EXPIRATION YEAR BITS.
  //
  Temp[0] = (iSnumPlain[12] & (4+8+16)) >> 2;
  Temp[1] = (iSnumPlain[13] & (1+2)) << 3;
  ThisData = (Temp[0] + Temp[1]) + 1990;
  //
  // RETURN EXPIRATION YEAR.
  //
  return(ThisData);
  ////
  //// MISCELLANEOUS DEBUG STUFF.
  ////
  //if ( DEBUGMODE )
  //{
  //  // OVERWRITE OLD FILE, IF ANY.
  //  fp = fopen(FNDEBUG_snumIsModulePurchased, "w");
  //  fclose(fp);
  //}
}



// PROTOTYPE:
//     extern "C" long __stdcall snumGetInternalSnum(char *spNumber)
// INPUTS:
//     spNumber = SERIAL NUMBER
// PURPOSE:
//     DETERMINES INTENAL SERIAL NUMBER.
// CALLED FROM:
//     CPASCHK VB5 PROGRAM
// RETURNS:
//     0 = INVALID!
//     NON-0 = THE INTERNAL SERIAL NUMBER.
//
extern "C" long __stdcall snumGetInternalSnum(char *spNumber)
{
int iSnumPlain[21];  
int RetVal;
//FILE *fp;
//int Temp[2];
long ThisData;
long Temp[5];
int i;
  RetVal = Internal_snumDecode(spNumber, iSnumPlain);
  if (RetVal == 0)
  {
    // SERIAL NUMBER IS INVALID!
    return(0);
  }
  //
  // EXTRACT EXPIRATION YEAR BITS.
  //
  Temp[0] = (iSnumPlain[13] & (4+8+16)) >> 2;
  Temp[1] = (iSnumPlain[14] & (1+2+4+8+16+32)) << (3);
  Temp[2] = (iSnumPlain[15] & (1+2+4+8+16+32)) << (3+5);
  Temp[3] = (iSnumPlain[16] & (1+2+4+8+16+32)) << (3+5+5);
  Temp[4] = (iSnumPlain[17] & (1+2+4)) << (3+5+5+5);
  ThisData = 0;
  for (i=0; i<=4; i++)
    ThisData += Temp[i];
  

  //// COPY ThisInternalSnum SPLIT ACROSS (INT 13 BITS 3-5)[67-69] AND 
  //// (INT 14 BITS 1-5)[70-74] AND (INT 15 BITS 1-5)[75-79] AND
  //// (INT 16 BITS 1-5)[80-84] AND (INT 17 BITS 1-3)[85-87].
  //Temp1 = (ThisInternalSnum & (1+2+4)) << 2;
  //Temp2 = (ThisInternalSnum & (8+16+32+64+128)) >> 3;
  //Temp3 = (ThisInternalSnum & (256+512+1024+2048+4096)) >> (3+5);
  //Temp4 = (ThisInternalSnum & (8192+16384+32768+65536+131072)) >> (3+5+5);
  //Temp5 = (ThisInternalSnum & (262144+524288+1048576)) >> (3+5+5+5);
  //iSnumPlain[13] += Temp1;
  //iSnumPlain[14] += Temp2;
  //iSnumPlain[15] += Temp3;
  //iSnumPlain[16] += Temp4;
  //iSnumPlain[17] += Temp5;

  
  //Temp[0] = (iSnumPlain[12] & (4+8+16)) >> 2;
  //Temp[1] = (iSnumPlain[13] & (1+2)) << 3;
  //ThisData = (Temp[0] + Temp[1]) + 1990;
  //
  // RETURN INTERNAL SERIAL NUMBER.
  //
  return(ThisData);
}



