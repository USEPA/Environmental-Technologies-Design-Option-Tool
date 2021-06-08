///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    SOURCE CODE FILE:
//        INTERNAL.H
//
//    PURPOSE:
//        PROTOTYPES/ETC FOR INTERNAL.CPP.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



// NON-EXPORTED FUNCTION PROTOTYPES.
void Internal_GetXorConstants(int iXorConstants[]);
void Internal_GetScatterConstants(int iScatterConstants[]);
int Internal_NormalToBase32(int in_Normal);
int Internal_Base32ToNormal(int in_Normal);
int Internal_snumDecode(char *IN_spNumber, int OUT_iSnumPlain[]);

// MISCELLANEOUS DEFINITIONS.
#define FNDEBUG_Internal_snumDecode ("c:\\cpaslib_snumdecode.txt")
#define ERROR_Internal_snumDecode_BAD_PREFIX                  (101)
#define ERROR_Internal_snumDecode_BAD_CHECK_CHAR              (102)
#define ERROR_Internal_snumDecode_BAD_HYPHENS                 (103)
#define ERROR_Internal_snumDecode_BAD_CHECK_INTEGER           (104)
#define ERROR_Internal_snumDecode_BAD_EXTRACTED_CHECKSUM      (105)
//#define ERROR_Internal_snumDecode_   ()

