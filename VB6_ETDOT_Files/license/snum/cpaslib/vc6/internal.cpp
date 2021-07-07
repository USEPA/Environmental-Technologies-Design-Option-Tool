///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    SOURCE CODE FILE:
//        INTERNAL.CPP
//
//    PURPOSE:
//        VARIOUS INTERNAL ROUTINES FOR CPASLIB.CPP.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// COMMON C/C++ FUNCTION PROTOTYPES.
#include <stdio.h>
#include <string.h>

// FUNCTION PROTOTYPES.
#include "cpaslib.h"
#include "internal.h"
#include "licgen.h"





// PROTOTYPE:
//     void Internal_GetXorConstants(int iXorConstants[])
// OUTPUTS:
//     iXorConstants = ARRAY OF XOR CONSTANTS.
// PURPOSE:
//     RETURNS THE ARRAY OF XOR CONSTANTS.
// CALLED FROM:
//     CPASLIB.CPP
// RETURNS:
//     NONE.
//
void Internal_GetXorConstants(int iXorConstants[])
{
  iXorConstants[0] = 17;
  iXorConstants[1] = 22;
  iXorConstants[2] = 17;
  iXorConstants[3] = 18;
  iXorConstants[4] = 9;
  iXorConstants[5] = 10;
  iXorConstants[6] = 25;
  iXorConstants[7] = 1;
  iXorConstants[8] = 24;
  iXorConstants[9] = 26;
  iXorConstants[10] = 22;
  iXorConstants[11] = 2;
  iXorConstants[12] = 13;
  iXorConstants[13] = 27;
  iXorConstants[14] = 25;
  iXorConstants[15] = 12;
  iXorConstants[16] = 30;
  iXorConstants[17] = 28;
  iXorConstants[18] = 2;
  iXorConstants[19] = 30;
  iXorConstants[20] = 12;
}


// PROTOTYPE:
//     void Internal_GetScatterConstants(int iScatterConstants[])
// OUTPUTS:
//     iScatterConstants = ARRAY OF SCATTER CONSTANTS.
// PURPOSE:
//     RETURNS THE ARRAY OF SCATTER CONSTANTS.
// CALLED FROM:
//     CPASLIB.CPP
// RETURNS:
//     NONE.
//
void Internal_GetScatterConstants(int iScatterConstants[])
{
  iScatterConstants[0] = 21;
  iScatterConstants[1] = 16;
  iScatterConstants[2] = 17;
  iScatterConstants[3] = 10;
  iScatterConstants[4] = 23;
  iScatterConstants[5] = 2;
  iScatterConstants[6] = 22;
  iScatterConstants[7] = 25;
  iScatterConstants[8] = 13;
  iScatterConstants[9] = 27;
  iScatterConstants[10] = 11;
  iScatterConstants[11] = 29;
  iScatterConstants[12] = 3;
  iScatterConstants[13] = 28;
  iScatterConstants[14] = 14;
  iScatterConstants[15] = 20;
  iScatterConstants[16] = 9;
  iScatterConstants[17] = 8;
  iScatterConstants[18] = 5;
  iScatterConstants[19] = 7;
  iScatterConstants[20] = 15;
}


// PROTOTYPE:
//     int Internal_NormalToBase32(int in_Normal)
// INPUTS:
//     in_Normal = NORMAL NUMBER (ACCEPTABLE RANGE = 0 TO 31).
// PURPOSE:
//     RETURNS THE NUMBER IN BASE 32 NOTATION, SEE BELOW.
//
//     TRANSLATION TABLE FOR [NORMAL] <===> [BASE32].
//
//               1111111111222222222233
//     01234567890123456789012345678901    (NORMAL)
//     ================================
//     0123456789ABCDEFGHJKLMNPQRTUVWXY    (BASE32)
//     ================================
//     89012345675678901245678012456789    (ASCII EQUIVALENT OF (BASE32))
//     44555555556666677777777888888888
//
//     NOTE: THE LETTERS I,O,S,Z WERE LEFT OUT DUE TO THE FACT
//     THAT IT'S POSSIBLE FOR THEM TO BE MISTAKEN FOR
//     THE NUMBERS 1,0,5,2.
//
// CALLED FROM:
//     CPASLIB.CPP
// RETURNS:
//     THE NUMBER IN BASE32 NOTATION.
//
int Internal_NormalToBase32(int in_Normal)
{
int iRetVal;
  iRetVal = 0;
  if ((in_Normal>=0) && (in_Normal<=9))
  {
    iRetVal = in_Normal + 48;
  }
  if ((in_Normal>=10) && (in_Normal<=17))
  {
    iRetVal = (in_Normal-10) + 65;
  }
  if ((in_Normal>=18) && (in_Normal<=22))
  {
    iRetVal = (in_Normal-18) + 74;
  }
  if ((in_Normal>=23) && (in_Normal<=25))
  {
    iRetVal = (in_Normal-23) + 80;
  }
  if ((in_Normal>=26) && (in_Normal<=31))
  {
    iRetVal = (in_Normal-26) + 84;
  }
  return(iRetVal);
}


// PROTOTYPE:
//     int Internal_Base32ToNormal(int in_Normal)
// INPUTS:
//     in_Normal = BASE32 NUMBER (ACCEPTABLE RANGES = 48-57,65-72,74-78,80-82,84-89).
// PURPOSE:
//     RETURNS THE NUMBER IN NORMAL NOTATION, SEE BELOW.
//
//     TRANSLATION TABLE FOR [NORMAL] <===> [BASE32].
//
//               1111111111222222222233
//     01234567890123456789012345678901    (NORMAL)
//     ================================
//     0123456789ABCDEFGHJKLMNPQRTUVWXY    (BASE32)
//     ================================
//     89012345675678901245678012456789    (ASCII EQUIVALENT OF (BASE32))
//     44555555556666677777777888888888
//
//     NOTE: THE LETTERS I,O,S,Z WERE LEFT OUT DUE TO THE FACT
//     THAT IT'S POSSIBLE FOR THEM TO BE MISTAKEN FOR
//     THE NUMBERS 1,0,5,2.
//
// CALLED FROM:
//     CPASLIB.CPP
// RETURNS:
//     THE NUMBER IN NORMAL NOTATION.
//
int Internal_Base32ToNormal(int in_Normal)
{
int iRetVal;
  iRetVal = 0;
  if ((in_Normal>=48) && (in_Normal<=57))
  {
    iRetVal = in_Normal - 48;
  }
  if ((in_Normal>=65) && (in_Normal<=72))
  {
    iRetVal = (in_Normal-65) + 10;
  }
  if ((in_Normal>=74) && (in_Normal<=78))
  {
    iRetVal = (in_Normal-74) + 18;
  }
  if ((in_Normal>=80) && (in_Normal<=82))
  {
    iRetVal = (in_Normal-80) + 23;
  }
  if ((in_Normal>=84) && (in_Normal<=89))
  {
    iRetVal = (in_Normal-84) + 26;
  }
  return(iRetVal);
}


// PROTOTYPE:
//     int Internal_snumDecode(char *IN_spNumber, int OUT_iSnumPlain[])
// INPUTS:
//     IN_spNumber = STRING POINTER TO SERIAL NUMBER (NOT NECESSARILY NULL-TERMINATED!).
// OUTPUTS:
//     OUT_iSnumPlain = DECODED SERIAL NUMBER DATA.
// PURPOSE:
//     VERIFIES AND DECODES THE SERIAL NUMBER, RETURNS THE DATA IT CONTAINS.
// CALLED FROM:
//     CPASLIB.CPP
// RETURNS:
//     0 = SERIAL NUMBER IS NOT VALID!
//     1 = SERIAL NUMBER IS VALID.
//
int Internal_snumDecode(char *IN_spNumber, int OUT_iSnumPlain[])
{
FILE *fp;
int iSnumPlain[21];
int iSnumXord[21];
int iSnumXlatd[21];
int iSnumFinal[30];
char sSnumFinal[31];    // EXTRA char TO HOLD THE NULL CHARACTER.
int i;
int j;
//int k;
//int ThisBit;
//int ThisVersionType;
//int ThisExpires;
//int ThisExpiresDay;
//int ThisExpiresMonth;
//int ThisExpiresYear;
//long ThisInternalSnum;
int ThisChecksum;
int Temp1;
int Temp2;
//int Temp3;
//int Temp4;
//int Temp5;
int iXorConstants[21];
int iScatterConstants[21];
int iCheckInteger;
int iErrorCode_Internal_snumDecode;
int ThisChecksum_Extracted;
int ThisInt;
  iErrorCode_Internal_snumDecode = 0;
  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  CONVERT spNumber TO iSnumFinal[]  ///////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  for (i=0; i<=29; i++)
  {
    sSnumFinal[i] = IN_spNumber[i];
  }
  sSnumFinal[30] = '\0';
  for (i=0; i<=29; i++)
  {
    iSnumFinal[i] = sSnumFinal[i];
  }
  if ( DEBUGMODE )
  {
    // OVERWRITE OLD FILE, IF ANY.
    fp = fopen(FNDEBUG_Internal_snumDecode, "w");
    fclose(fp);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "INPUT PARAMETERS TO Internal_snumDecode():\n");
    fprintf(fp, "sSnumFinal = `%s`\n", sSnumFinal);
    fclose(fp);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  PERFORM VERIFICATION STEPS WITH iSnumFinal[]  ///////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  // VERIFY FIRST TWO INTEGERS SET TO "CA".
  if ((iSnumFinal[0]!=67) || (iSnumFinal[1]!=65))
  {
    iErrorCode_Internal_snumDecode = ERROR_Internal_snumDecode_BAD_PREFIX;
    goto err_snumDecode_InvalidSnum;
  }
  // VERIFY CHECK CHARACTER IN POSITION 26.
  if (iSnumFinal[26]!=65)
  {
    iErrorCode_Internal_snumDecode = ERROR_Internal_snumDecode_BAD_CHECK_CHAR;
    goto err_snumDecode_InvalidSnum;
  }
  // PLACE HYPHEN CHARACTERS INTO POSITIONS 6,12,18,24.
  if ((iSnumFinal[6]!=45)||(iSnumFinal[12]!=45)||(iSnumFinal[18]!=45)||(iSnumFinal[24]!=45))
  {
    iErrorCode_Internal_snumDecode = ERROR_Internal_snumDecode_BAD_HYPHENS;
    goto err_snumDecode_InvalidSnum;
  }
  // GET SCATTER CONSTANTS.
  Internal_GetScatterConstants(iScatterConstants);
  // CALCULATE PROPER CHECK INTEGER AND VERIFY IN POSITIONS 4 AND 19.
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
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "\n---------- CONTENTS OF Temp* (POINT A1):\n");
    fprintf(fp, "Temp1 = %d\n", Temp1);
    fprintf(fp, "Temp2 = %d\n", Temp2);
    fclose(fp);
  }
  if ( (Internal_Base32ToNormal(iSnumFinal[4])!=Temp1) ||
       (Internal_Base32ToNormal(iSnumFinal[19])!=Temp2))
  {
    iErrorCode_Internal_snumDecode = ERROR_Internal_snumDecode_BAD_CHECK_INTEGER;
    goto err_snumDecode_InvalidSnum;
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z1\n");
    fclose(fp);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  UNSCATTER iSnumFinal[] INTO iSnumXlatd[]  ///////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  // GET SCATTER CONSTANTS AND UNSCATTER INTO iSnumXlatd[].
  Internal_GetScatterConstants(iScatterConstants);
  for (i=0; i<=20; i++)
  {
    iSnumXlatd[i] = iSnumFinal[iScatterConstants[i]];
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z2\n");
    fclose(fp);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  TRANSLATE BASE-32 iSnumXlatd[] TO NORMAL NOTATION iSnumXord[]  //
  ///////////////////////////////////////////////////////////////////////////////////////////
  for (i=0; i<=20; i++)
  {
    iSnumXord[i] = Internal_Base32ToNormal(iSnumXlatd[i]);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z3\n");
    fclose(fp);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  XOR iSnumXord[] TO GET iSnumPlain[]  ////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  Internal_GetXorConstants(iXorConstants);
  for (i=0; i<=20; i++)
  {
    iSnumPlain[i] = iSnumXord[i] ^ iXorConstants[i];
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z4\n");
    fclose(fp);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  VERIFY THE CHECKSUM IN iSnumPlain[]  ////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  //
  // CALCULATE CHECKSUM (ThisChecksum).
  //
  ThisChecksum = 0;
  for (i=0; i<=17; i++)
  {
    if (i == 17)
    {
      ThisInt = iSnumPlain[i] & (1+2+4);
    }
    else
    {
      ThisInt = iSnumPlain[i];
    }
    j = 18 - i;
    ThisChecksum += (j)*(ThisInt);
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z5\n");
    fclose(fp);
  }
  //
  // EXTRACT ThisChecksum_Extracted FROM THE iSnumPlain[] ARRAY (BITS 88 TO 100).
  //
  // EXTRACT ThisChecksum_Extracted SPLIT ACROSS
  // (INT 17 BITS 4-5)[88-89] AND (INT 18 BITS 1-5)[90-94] AND 
  // (INT 19 BITS 1-5)[95-99] AND (INT 20 BIT 1)[100].
  //
  ThisChecksum_Extracted = 0;
  ThisChecksum_Extracted += (iSnumPlain[17] & (8+16)) >> 3;
  ThisChecksum_Extracted += (iSnumPlain[18] & (1+2+4+8+16)) << 2;
  ThisChecksum_Extracted += (iSnumPlain[19] & (1+2+4+8+16)) << (2+5);
  ThisChecksum_Extracted += (iSnumPlain[20] & (1+2+4+8+16)) << (2+5+5);
  if ( ThisChecksum_Extracted != ThisChecksum ) 
  {
    iErrorCode_Internal_snumDecode = ERROR_Internal_snumDecode_BAD_EXTRACTED_CHECKSUM;
    goto err_snumDecode_InvalidSnum;
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z6\n");
    fclose(fp);
  }
  
  ///////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////  TRANSFER iSnumPlain[] TO OUT_iSnumPlain[]  //////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////
  for (i=0; i<=20; i++)
  {
    OUT_iSnumPlain[i] = iSnumPlain[i];
  }
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "Test point Z7\n");
    fclose(fp);
  }

  // RETURN "SUCCESS" MESSAGE.
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "iErrorCode_Internal_snumDecode VALUE = %d\n", iErrorCode_Internal_snumDecode);
    fprintf(fp, "RETURN VALUE = 1\n");
    fclose(fp);
  }
  return(1);
  
err_snumDecode_InvalidSnum:
  // INVALID INPUT.
  if ( DEBUGMODE )
  {
    fp = fopen(FNDEBUG_Internal_snumDecode, "a+");
    fprintf(fp, "iErrorCode_Internal_snumDecode VALUE = %d\n", iErrorCode_Internal_snumDecode);
    fprintf(fp, "RETURN VALUE = 0\n");
    fclose(fp);
  }
  return(0);
}


