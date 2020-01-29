///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    SOURCE CODE FILE:
//        CPASLIB.H
//
//    PURPOSE:
//        PROTOTYPES/ETC FOR CPASLIB.CPP.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



// MISCELLANEOUS DEFINITIONS.
//#define DEBUGMODE (1)
#define DEBUGMODE (0)
#define FNDEBUG_snumVerify                  ("c:\\cpaslib_snumverify.txt")
#define FNDEBUG_snumGenerate                ("c:\\cpaslib_snumgenerate.txt")
#define FNDEBUG_snumIsModulePurchased       ("c:\\cpaslib_snumIsModulePurchased.txt")
#define FNDEBUG_snumGetVersionType          ("c:\\cpaslib_snumGetVersionType.txt")
#define FNDEBUG_snumIsExpirationPresent     ("c:\\cpaslib_snumIsExpirationPresent.txt")
#define FNDEBUG_snumGetExpirationDay        ("c:\\cpaslib_snumGetExpirationDay.txt")
#define FNDEBUG_snumGetExpirationMonth      ("c:\\cpaslib_snumGetExpirationMonth.txt")
#define FNDEBUG_snumGetExpirationYear       ("c:\\cpaslib_snumGetExpirationYear.txt")

// FUNCTION PARAMETER DEFINITIONS.
#define VERSIONTYPE_ALPHA           (1)
#define VERSIONTYPE_BETA            (2)
#define VERSIONTYPE_STANDARD        (3)
#define EXPIRES_NO                  (0)
#define EXPIRES_YES                 (1)
#define ICHECK_CORRECT_VALUE        (13892)

// EXPORTED FUNCTION PROTOTYPES.
extern "C" int __stdcall snumVerify(char *spNumber);
extern "C" int __stdcall snumGenerate(char *spNumber, int iModules[], int iVersionType, int iExpires, 
                           int iExpiresDay, int iExpiresMonth, int iExpiresYear, 
                           long longInternalSnum, int iCheck);
extern "C" int __stdcall snumIsModulePurchased(char *spNumber, int iModule);
extern "C" int __stdcall snumGetVersionType(char *spNumber);
extern "C" int __stdcall snumIsExpirationPresent(char *spNumber);
extern "C" int __stdcall snumGetExpirationDay(char *spNumber);
extern "C" int __stdcall snumGetExpirationMonth(char *spNumber);
extern "C" int __stdcall snumGetExpirationYear(char *spNumber);
extern "C" long __stdcall snumGetInternalSnum(char *spNumber);



