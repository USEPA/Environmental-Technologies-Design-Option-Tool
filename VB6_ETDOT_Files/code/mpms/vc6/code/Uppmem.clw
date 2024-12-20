; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CMechModelsDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "Uppmem.h"
LastPage=0

ClassCount=21
Class1=CUppmemApp
Class2=CUppmemDoc
Class3=CUppmemView
Class4=CMainFrame

ResourceCount=15
Resource1=IDD_ERROR
Resource2=IDD_ADDTL_PARAMS
Class5=CAboutDlg
Class6=CMechModelsDlg
Resource3=IDR_MAINFRAME
Class7=CEmpModelsDlg
Resource4=IDD_RUN
Resource5=IDD_EMPIRICAL
Class8=CPlantDesignDlg
Class9=CPlantRunDlg
Resource6=IDD_FLUX_RANGE
Resource7=IDD_ABOUTBOX
Resource8=IDD_PREDEF_MEMB
Resource9=IDD_PARAM_RANGE
Class10=CAddParamDlg
Class11=CPartDistribDlg
Class12=CParamRangeDlg
Resource10=IDD_EMP_PICK_MODEL
Class13=CPreDefMemDlg
Class14=CErrBox
Resource11=IDD_EMP_DATA_ENTER
Resource12=IDD_DESIGN
Class15=CEnterDataDlg
Class16=CEmpPickModel
Class17=CEmpPickModelDlg
Resource13=IDD_DISTRIB
Class18=CFluxRangeDlg
Class19=CEmpModel
Class20=CEmpData
Resource14=IDD_MECHANISTIC
Class21=CPermDistribDlg
Resource15=IDD_PERM_PART_DISTRIB_DLG

[CLS:CUppmemApp]
Type=0
HeaderFile=Uppmem.h
ImplementationFile=Uppmem.cpp
Filter=N
BaseClass=CWinApp
VirtualFilter=AC

[CLS:CUppmemDoc]
Type=0
HeaderFile=UppmemDoc.h
ImplementationFile=UppmemDoc.cpp
Filter=N
BaseClass=CDocument
VirtualFilter=DC

[CLS:CUppmemView]
Type=0
HeaderFile=UppmemView.h
ImplementationFile=UppmemView.cpp
Filter=C
LastObject=CUppmemView

[CLS:CMainFrame]
Type=0
HeaderFile=MainFrm.h
ImplementationFile=MainFrm.cpp
Filter=T



[CLS:CAboutDlg]
Type=0
HeaderFile=Uppmem.cpp
ImplementationFile=Uppmem.cpp
Filter=D
LastObject=CAboutDlg

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutDlg
ControlCount=6
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDOK,button,1342373889
Control4=IDC_STATIC,static,1342308352
Control5=IDC_STATIC,static,1342308352
Control6=IDC_STATIC,static,1342308352

[MNU:IDR_MAINFRAME]
Type=1
Class=CMainFrame
Command1=ID_FILE_NEW
Command2=ID_FILE_OPEN
Command3=ID_FILE_SAVE
Command4=ID_FILE_SAVE_AS
Command5=ID_FILE_PRINT
Command6=ID_FILE_PRINT_PREVIEW
Command7=ID_FILE_PRINT_SETUP
Command8=ID_FILE_MRU_FILE1
Command9=ID_APP_EXIT
Command10=ID_EDIT_UNDO
Command11=ID_EDIT_CUT
Command12=ID_EDIT_COPY
Command13=ID_EDIT_PASTE
Command14=ID_VIEW_TOOLBAR
Command15=ID_VIEW_STATUS_BAR
Command16=ID_MODEL_MECH
Command17=ID_MODEL_EMP
Command18=ID_MODEL_DESIGN
Command19=ID_MODEL_RUN
Command20=ID_HELP_FINDER
Command21=ID_APP_ABOUT
CommandCount=21

[ACL:IDR_MAINFRAME]
Type=1
Class=CMainFrame
Command1=ID_FILE_NEW
Command2=ID_FILE_OPEN
Command3=ID_FILE_SAVE
Command4=ID_FILE_PRINT
Command5=ID_EDIT_UNDO
Command6=ID_EDIT_CUT
Command7=ID_EDIT_COPY
Command8=ID_EDIT_PASTE
Command9=ID_EDIT_UNDO
Command10=ID_EDIT_CUT
Command11=ID_EDIT_COPY
Command12=ID_EDIT_PASTE
Command13=ID_NEXT_PANE
Command14=ID_PREV_PANE
Command15=ID_CONTEXT_HELP
Command16=ID_HELP
CommandCount=16

[TB:IDR_MAINFRAME]
Type=1
Class=?
Command1=ID_FILE_NEW
Command2=ID_FILE_OPEN
Command3=ID_FILE_SAVE
Command4=ID_FILE_PRINT
Command5=ID_MODEL_MECH
Command6=ID_MODEL_EMP
Command7=ID_MODEL_DESIGN
Command8=ID_MODEL_RUN
Command9=ID_APP_ABOUT
Command10=ID_CONTEXT_HELP
CommandCount=10

[DLG:IDD_MECHANISTIC]
Type=1
Class=CMechModelsDlg
ControlCount=47
Control1=IDC_MECH_PARAM_RANGE,button,1342242816
Control2=IDC_MECH_PRESSURE,edit,1350631552
Control3=IDC_MECH_TEMP,edit,1350631552
Control4=IDC_MECH_INFLUENT_FLOW,edit,1350631552
Control5=IDC_MECH_VISCOSITY,edit,1350631552
Control6=IDC_MECH_MEMB_SELECT,button,1342242816
Control7=IDC_MECH_PORE_RADIUS,edit,1350631552
Control8=IDC_MECH_MEMB_RESISTANCE,edit,1350631552
Control9=IDC_MECH_MEMB_CHANNEL_RADIUS,edit,1350631552
Control10=IDC_MECH_MEMB_LENGTH,edit,1350631552
Control11=IDC_MECH_MEMB_AREA,edit,1350631552
Control12=IDC_MECH_RECIRC,edit,1350631552
Control13=IDC_MECH_PART_DISTRIBUTION,button,1342242816
Control14=IDC_MECH_CALC_REJECT,button,1342242819
Control15=IDC_MECH_AVE_PART_RADIUS,edit,1350631552
Control16=IDC_MECH_AVE_PART_CONC,edit,1350631552
Control17=IDC_MECH_AVE_PART_DENSITY,edit,1350631552
Control18=IDC_MECH_ADDTL_MODEL_PARAMS,button,1342242816
Control19=IDC_MECH_MEMSYS_RADIO,button,1342177289
Control20=IDC_MECH_SE_RADIO,button,1342177289
Control21=IDC_MECH_RESISTANCE_RADIO,button,1342177289
Control22=IDC_MECH_GEL_RADIO,button,1342177289
Control23=IDC_MECH_CALC_FLUX,button,1342254849
Control24=ID_MECH_SAVE,button,1342242816
Control25=ID_HELP,button,1342242816
Control26=IDCANCEL,button,1342242816
Control27=IDC_STATIC,button,1342177287
Control28=IDC_STATIC,static,1342308352
Control29=IDC_STATIC,static,1342308352
Control30=IDC_STATIC,static,1342308352
Control31=IDC_STATIC,static,1342308352
Control32=IDC_STATIC,button,1342177287
Control33=IDC_STATIC,static,1342308352
Control34=IDC_STATIC,static,1342308352
Control35=IDC_STATIC,button,1342177287
Control36=IDC_STATIC,button,1342177287
Control37=IDC_STATIC,static,1342308352
Control38=IDC_STATIC,static,1342308352
Control39=IDC_STATIC,static,1342308352
Control40=IDC_STATIC,static,1342308352
Control41=IDC_STATIC,static,1342308352
Control42=IDC_STATIC,static,1342308352
Control43=IDC_STATIC,static,1342308352
Control44=IDC_MECH_FLUX_MS,edit,1342179458
Control45=IDC_MECH_FLUX_LH,edit,1342179458
Control46=IDC_STATIC,static,1342308864
Control47=IDC_STATIC,static,1342308864

[CLS:CMechModelsDlg]
Type=0
HeaderFile=MechModelsDlg.h
ImplementationFile=MechModelsDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CMechModelsDlg

[DLG:IDD_EMPIRICAL]
Type=1
Class=CEmpModelsDlg
ControlCount=37
Control1=IDC_EMP_PARAM_RANGE,button,1342242816
Control2=IDC_EMP_AVE_PART_CONC,edit,1350631552
Control3=IDC_EMP_PRESSURE,edit,1350631552
Control4=IDC_EMP_TEMP,edit,1350631552
Control5=IDC_EMP_INFLUENT_FLOW,edit,1350631552
Control6=IDC_EMP_VISCOSITY,edit,1484849280
Control7=IDC_EMP_EXP_DATA,button,1476460544
Control8=IDC_EMP_PERM_TIME,edit,1350631552
Control9=IDC_EMP_CLEAN_TIME,edit,1350631552
Control10=IDC_EMP_CALC_FLUX,button,1342254849
Control11=ID_EMP_SAVE,button,1342242816
Control12=ID_HELP,button,1342242816
Control13=IDCANCEL,button,1342242816
Control14=IDC_STATIC,button,1342177287
Control15=IDC_STATIC,button,1342177287
Control16=IDC_STATIC,static,1342308352
Control17=IDC_STATIC,static,1342308352
Control18=IDC_STATIC,static,1342308352
Control19=IDC_STATIC,static,1476526080
Control20=IDC_STATIC,static,1342308352
Control21=IDC_EMP_FLUX_MS,edit,1342179458
Control22=IDC_EMP_FLUX_LH,edit,1342179458
Control23=IDC_STATIC,static,1342308864
Control24=IDC_STATIC,static,1342308864
Control25=IDC_STATIC,static,1342308352
Control26=IDC_STATIC,static,1342308352
Control27=IDC_EMP_FLUX_TIME,edit,1342179456
Control28=IDC_STATIC,static,1342308352
Control29=IDC_EMP_CONC_UNITS,edit,1342179456
Control30=IDC_STATIC,button,1342177287
Control31=IDC_EMP_PARAM_A_NAME,edit,1342244994
Control32=IDC_EMP_PARAM_A_VAL,edit,1350631552
Control33=IDC_EMP_PARAM_C_NAME,edit,1342244994
Control34=IDC_EMP_PARAM_B_NAME,edit,1342244994
Control35=IDC_EMP_PARAM_B_VAL,edit,1350631552
Control36=IDC_EMP_PARAM_C_VAL,edit,1350631552
Control37=IDC_EMP_CUST_MODEL_VAL,button,1342242819

[CLS:CEmpModelsDlg]
Type=0
HeaderFile=EmpModelsDlg.h
ImplementationFile=EmpModelsDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CEmpModelsDlg

[DLG:IDD_DESIGN]
Type=1
Class=CPlantDesignDlg
ControlCount=4
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342308352
Control4=IDC_STATIC,static,1342177283

[DLG:IDD_RUN]
Type=1
Class=CPlantRunDlg
ControlCount=4
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342308352
Control4=IDC_STATIC,static,1342177283

[CLS:CPlantDesignDlg]
Type=0
HeaderFile=PlantDesignDlg.h
ImplementationFile=PlantDesignDlg.cpp
BaseClass=CDialog
Filter=D

[CLS:CPlantRunDlg]
Type=0
HeaderFile=PlantRunDlg.h
ImplementationFile=PlantRunDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=CPlantRunDlg

[DLG:IDD_PARAM_RANGE]
Type=1
Class=CParamRangeDlg
ControlCount=20
Control1=IDC_RANGE_PRESS1,edit,1350631552
Control2=IDC_RANGE_PRESS2,edit,1350631552
Control3=IDC_RANGE_TEMP1,edit,1350631552
Control4=IDC_RANGE_TEMP2,edit,1350631552
Control5=IDC_RANGE_FLOW1,edit,1350631552
Control6=IDC_RANGE_FLOW2,edit,1350631552
Control7=IDC_RANGE_VISC1,edit,1350631552
Control8=IDC_RANGE_VISC2,edit,1350631552
Control9=IDC_RANGE_CONC1,edit,1350631552
Control10=IDC_RANGE_CONC2,edit,1350631552
Control11=IDC_RANGE_NUM_STEPS,edit,1350631552
Control12=IDOK,button,1342242817
Control13=IDCANCEL,button,1342242816
Control14=IDC_STATIC,static,1342308352
Control15=IDC_RANGE_PRESS_RADIO,button,1342177289
Control16=IDC_RANGE_TEMP_RADIO,button,1342177289
Control17=IDC_RANGE_FLOW_RADIO,button,1342177289
Control18=IDC_RANGE_VISC_RADIO,button,1342177289
Control19=IDC_STATIC,static,1342308352
Control20=IDC_RANGE_CONC_RADIO,button,1342177289

[DLG:IDD_ADDTL_PARAMS]
Type=1
Class=CAddParamDlg
ControlCount=17
Control1=IDC_MECH_AMP_OP_REST,edit,1350631552
Control2=IDC_MECH_AMP_IRREV_REST,edit,1350631552
Control3=IDCANCEL,button,1342242816
Control4=IDOK,button,1342242817
Control5=IDC_STATIC,static,1342308352
Control6=IDC_STATIC,button,1342177287
Control7=IDC_MECH_AMP_MTC_ESTIMATE_RADIO,button,1342308361
Control8=IDC_MECH_AMP_MTC_ENTER_RADIO,button,1342308361
Control9=IDC_STATIC,button,1342177287
Control10=IDC_STATIC,static,1342308352
Control11=IDC_STATIC,static,1342308352
Control12=IDC_MECH_AMP_MGL_RADIO,button,1342308361
Control13=IDC_MECH_AMP_VOL_RADIO,button,1342308361
Control14=ID_HELP,button,1342242816
Control15=IDC_MECH_AMP_MGL_CGEL,edit,1350631552
Control16=IDC_MECH_AMP_VOLFR_CGEL,edit,1350631552
Control17=IDC_MECH_AMP_MTC,edit,1350631552

[DLG:IDD_PREDEF_MEMB]
Type=1
Class=CPreDefMemDlg
ControlCount=22
Control1=IDC_MEMB_NAME,edit,1350631552
Control2=IDC_MEMB_RESISTANCE,edit,1350631552
Control3=IDC_MEMB_CHANNEL_RADIUS,edit,1350631552
Control4=IDC_MEMB_MANFC,edit,1350631552
Control5=IDC_MEMB_LENGTH,edit,1350631552
Control6=IDC_MEMB_PORE_RADIUS,edit,1350631552
Control7=IDC_MEMB_AREA,edit,1350631552
Control8=IDC_MEMB_VIEW,button,1342242816
Control9=IDC_MEMB_ENTER,button,1342242816
Control10=IDC_MEMB_REMOVE,button,1342242816
Control11=ID_HELP,button,1342242816
Control12=IDCANCEL,button,1342242816
Control13=IDOK,button,1342242817
Control14=IDC_STATIC,static,1342308353
Control15=IDC_MEMB_LIST,listbox,1352728963
Control16=IDC_STATIC,static,1342308352
Control17=IDC_STATIC,static,1342308352
Control18=IDC_STATIC,static,1342308352
Control19=IDC_STATIC,static,1342308352
Control20=IDC_STATIC,static,1342308352
Control21=IDC_STATIC,static,1342308352
Control22=IDC_STATIC,static,1342308352

[DLG:IDD_DISTRIB]
Type=1
Class=CPartDistribDlg
ControlCount=12
Control1=IDC_PARTICLE_LIST,listbox,1353777283
Control2=IDC_DISTRIB_SIZE,edit,1350631552
Control3=IDC_DISTRIB_CONC_MASS,edit,1350631552
Control4=IDC_DISTRIB_ENTER,button,1342242816
Control5=IDC_DISTRIB_REMOVE,button,1342242816
Control6=IDC_DISTRIB_VIEW,button,1342242816
Control7=ID_HELP,button,1342242816
Control8=IDCANCEL,button,1342242816
Control9=IDOK,button,1342242817
Control10=IDC_STATIC,static,1342308353
Control11=IDC_STATIC,static,1342308352
Control12=IDC_STATIC,static,1342308352

[CLS:CAddParamDlg]
Type=0
HeaderFile=AddParamDlg.h
ImplementationFile=AddParamDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=CAddParamDlg
VirtualFilter=dWC

[CLS:CPartDistribDlg]
Type=0
HeaderFile=PartDistribDlg.h
ImplementationFile=PartDistribDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CPartDistribDlg

[CLS:CParamRangeDlg]
Type=0
HeaderFile=ParamRangeDlg.h
ImplementationFile=ParamRangeDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CParamRangeDlg

[CLS:CPreDefMemDlg]
Type=0
HeaderFile=PreDefMemDlg.h
ImplementationFile=PreDefMemDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC

[DLG:IDD_ERROR]
Type=1
Class=CErrBox
ControlCount=6
Control1=IDOK,button,1342242817
Control2=IDC_ERR_MSG3,edit,1342244992
Control3=IDC_ERR_MSG,edit,1342244992
Control4=IDC_ERR_MSG2,edit,1342244992
Control5=IDC_STATIC,static,1342177283
Control6=IDCANCEL,button,1342242816

[CLS:CErrBox]
Type=0
HeaderFile=ErrBox.h
ImplementationFile=ErrBox.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=IDC_ERR_MSG

[DLG:IDD_EMP_DATA_ENTER]
Type=1
Class=CEnterDataDlg
ControlCount=21
Control1=IDC_EMP_ENT_FLUX_DATA,edit,1350635588
Control2=IDC_EMP_ENT_PARAM2_DATA,edit,1350635588
Control3=IDC_EMP_ENT_CONC,edit,1350631552
Control4=IDC_EMP_ENT_PRES,edit,1350631552
Control5=IDC_EMP_ENT_VLOS,edit,1350631552
Control6=IDC_EMP_ENT_TEMP,edit,1350631552
Control7=IDOK,button,1342242817
Control8=ID_HELP,button,1342242816
Control9=IDCANCEL,button,1342242816
Control10=IDC_STATIC,static,1342308352
Control11=IDC_STATIC,button,1342178055
Control12=IDC_EMP_ENT_PARAM2_NAME,edit,1342179328
Control13=IDC_STATIC,button,1342177287
Control14=IDC_EMP_ENT_TIME_RADIO,button,1342308361
Control15=IDC_EMP_ENT_CONC_RADIO,button,1342308361
Control16=IDC_EMP_ENT_PRES_RADIO,button,1342308361
Control17=IDC_EMP_ENT_VLOS_RADIO,button,1342308361
Control18=IDC_EMP_ENT_TEMP_RADIO,button,1342308361
Control19=IDC_EMP_ENT_MGL_RADIO,button,1342308361
Control20=IDC_EMP_ENT_VOL_RADIO,button,1342308361
Control21=IDC_EMP_INFO,edit,1342179329

[CLS:CEnterDataDlg]
Type=0
HeaderFile=EnterDataDlg.h
ImplementationFile=EnterDataDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=ID_HELP
VirtualFilter=dWC

[CLS:CEmpPickModel]
Type=0
HeaderFile=EmpPickModel.h
ImplementationFile=EmpPickModel.cpp
BaseClass=CDialog
Filter=D

[DLG:IDD_EMP_PICK_MODEL]
Type=1
Class=CEmpPickModelDlg
ControlCount=12
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342308353
Control4=ID_HELP,button,1342242816
Control5=IDC_EMP_PICK_FOUL_RADIO,button,1342177289
Control6=IDC_EMP_PICK_REST_RADIO,button,1342177289
Control7=IDC_EMP_PICK_SURF_RADIO,button,1342177289
Control8=IDC_EMP_PICK_GEL_RADIO,button,1342177289
Control9=IDC_STATIC,static,1342177283
Control10=IDC_STATIC,static,1342177283
Control11=IDC_STATIC,static,1342177283
Control12=IDC_STATIC,static,1342177283

[CLS:CEmpPickModelDlg]
Type=0
HeaderFile=EmpPickModelDlg.h
ImplementationFile=EmpPickModelDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=ID_HELP
VirtualFilter=dWC

[DLG:IDD_FLUX_RANGE]
Type=1
Class=CFluxRangeDlg
ControlCount=5
Control1=IDOK,button,1342242817
Control2=IDC_RANGE_PARAM_NAME,edit,1342244992
Control3=IDC_RANGE_PARAM_EDIT,edit,1350567940
Control4=IDC_RANGE_FLUX_EDIT,edit,1350567940
Control5=IDC_STATIC,static,1342308352

[CLS:CFluxRangeDlg]
Type=0
HeaderFile=FluxRangeDlg.h
ImplementationFile=FluxRangeDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CFluxRangeDlg

[CLS:CEmpModel]
Type=0
HeaderFile=EmpModel.h
ImplementationFile=EmpModel.cpp
BaseClass=generic CWnd
Filter=W

[CLS:CEmpData]
Type=0
HeaderFile=EmpData.h
ImplementationFile=EmpData.cpp
BaseClass=generic CWnd
Filter=W
LastObject=CEmpData

[DLG:IDD_PERM_PART_DISTRIB_DLG]
Type=1
Class=CPermDistribDlg
ControlCount=5
Control1=IDOK,button,1342242817
Control2=IDC_PERM_PART_LIST,listbox,1353777539
Control3=IDC_RETEN_PART_LIST,listbox,1352728963
Control4=IDC_STATIC,static,1342308353
Control5=IDC_STATIC,static,1342308353

[CLS:CPermDistribDlg]
Type=0
HeaderFile=PermDistribDlg.h
ImplementationFile=PermDistribDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=CPermDistribDlg
VirtualFilter=dWC

