///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    SOURCE CODE FILE:
//        LICGEN.H
//
//    PURPOSE:
//        PROTOTYPES/ETC FOR LICGEN.CPP.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


// EXPORTED FUNCTION PROTOTYPES.
extern "C" int __stdcall snumCpasLicGenerate(
    char *spCpasDir,
    char *spWinDir,
    char *spNumber,
    char *spUserName,
    char *spUserCompany);
//extern "C" int __stdcall snumVerify(char *spNumber);
//extern "C" int __stdcall snumGenerate(char *spNumber, int iModules[], int iVersionType, int iExpires, 
//                           int iExpiresDay, int iExpiresMonth, int iExpiresYear, 
//                           long longInternalSnum, int iCheck);
//extern "C" int __stdcall snumIsModulePurchased(char *spNumber, int iModule);
//extern "C" int __stdcall snumGetVersionType(char *spNumber);
//extern "C" int __stdcall snumIsExpirationPresent(char *spNumber);
//extern "C" int __stdcall snumGetExpirationDay(char *spNumber);
//extern "C" int __stdcall snumGetExpirationMonth(char *spNumber);
//extern "C" int __stdcall snumGetExpirationYear(char *spNumber);

// NON-EXPORTED FUNCTION PROTOTYPES.
int LicFile_GenerateEncryptionKey();
int LicFile_Create(char *fn_CPASLIC, long filesize_CPASLIC);
int LicFile_PutEncryptedString(char *fn_This,long pos_This,char *str_This);
int LicFile_EncryptString(char *str_Plain,char *str_Enc);
//extern "C" int __stdcall snumCpasLicGenerate(char *spNumber);
//void Internal_GetXorConstants(int iXorConstants[]);
//void Internal_GetScatterConstants(int iScatterConstants[]);
//int Internal_NormalToBase32(int in_Normal);
//int Internal_Base32ToNormal(int in_Normal);
//int Internal_snumDecode(char *IN_spNumber, int OUT_iSnumPlain[]);

// MISCELLANEOUS DEFINITIONS.
//#define LICFILE_GoodSerialNumber ("OKNUM.X")
//#define LICFILE_BadSerialNumber ("BADNUM.X")
//#define LICFILE_NewLicInfo ("NEWLIC.X")
#define LICFILE_LicName ("CPAS.LIC")
#define LICFILE_DATE_NEVER ("NEVER")
#define LICFILE_ExtraCheckFile ("MTCHK.LIC")
#define LICFILE_ExtraCheckFile_Text ("IOW2EK4FV832")
//#define FNDEBUG_Internal_snumDecode ("c:\\cpaslib_snumdecode.txt")
//#define ERROR_Internal_snumDecode_BAD_PREFIX                  (101)
//#define ERROR_Internal_snumDecode_BAD_CHECK_CHAR              (102)
//#define ERROR_Internal_snumDecode_BAD_HYPHENS                 (103)
//#define ERROR_Internal_snumDecode_BAD_CHECK_INTEGER           (104)
//#define ERROR_Internal_snumDecode_BAD_EXTRACTED_CHECKSUM      (105)
////#define ERROR_Internal_snumDecode_   ()
#define FNDEBUG_snumCpasLicGenerate ("c:\\cpaslicgen.txt")

#define lfdt_order_Z_SERIALNUMBER (2)
#define lfdt_order_Z_USERNAME (3)
#define lfdt_order_Z_USERCOMPANY (4)
#define lfdt_order_ZZ_LASTEXECUTIONDATE (5)
#define lfdt_order_ZZ_LASTEXECUTIONTIME (6)
#define lfdt_order_ZZ_NUMPROGRAMKEYS (7)

#define pkdt_order_Z_PROGRAMKEY (0)
#define pkdt_order_Z_EXPIRATIONDATE (1)
#define pkdt_order_Z_RELEASETYPE (2)
#define pkdt_order_Z_VERSIONCODE (3)
#define pkdt_order_Z_VERSIONTYPE (4)


// STRUCTURE DEFINITIONS.
struct LicFile_Data_Type
{
  char Z_SERIALNUMBER[100];
  char Z_USERNAME[100];
  char Z_USERCOMPANY[100];
  char ZZ_LASTEXECUTIONDATE[100];
  char ZZ_LASTEXECUTIONTIME[100];
  int ZZ_NUMPROGRAMKEYS;
};
//#define RMT_SIZE (10)
#define RMT_SIZE (30)
struct Recognized_Mod_Type
{
  int idx;
  char Z_PROGRAMKEY[100];
};
struct ProgramKey_Data_Type
{
  char Z_PROGRAMKEY[100];
  char Z_EXPIRATIONDATE[100];
  char Z_RELEASETYPE[100];
  char Z_VERSIONCODE[100];
  char Z_VERSIONTYPE[100];
};

