// AddParamDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "AddParamDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddParamDlg dialog


CAddParamDlg::CAddParamDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddParamDlg::IDD, pParent)
{

	//{{AFX_DATA_INIT(CAddParamDlg)
	m_mgl_radio = -1;
	m_est_radio = -1;
	m_vol_radio = -1;
	m_ent_radio = -1;
	m_ir_res = 1e11;
	m_op_res = 1e11;
	m_mgl_conc = 2e4;
	m_vol_conc = 2e2;
	m_K_str = _T("1e-006");
	//}}AFX_DATA_INIT

	
}


void CAddParamDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddParamDlg)
	DDX_Radio(pDX, IDC_MECH_AMP_MGL_RADIO, m_mgl_radio);
	DDX_Radio(pDX, IDC_MECH_AMP_MTC_ESTIMATE_RADIO, m_est_radio);
	DDX_Radio(pDX, IDC_MECH_AMP_VOL_RADIO, m_vol_radio);
	DDX_Radio(pDX, IDC_MECH_AMP_MTC_ENTER_RADIO, m_ent_radio);
	DDX_Text(pDX, IDC_MECH_AMP_IRREV_REST, m_ir_res);
	DDV_MinMaxDouble(pDX, m_ir_res, 0., 1.e+020);
	DDX_Text(pDX, IDC_MECH_AMP_OP_REST, m_op_res);
	DDV_MinMaxDouble(pDX, m_op_res, 0., 1.e+020);
	DDX_Text(pDX, IDC_MECH_AMP_MGL_CGEL, m_mgl_conc);
	DDV_MinMaxDouble(pDX, m_mgl_conc, 0., 1.e+006);
	DDX_Text(pDX, IDC_MECH_AMP_VOLFR_CGEL, m_vol_conc);
	DDV_MinMaxDouble(pDX, m_vol_conc, 0., 1.e+002);
	DDX_Text(pDX, IDC_MECH_AMP_MTC, m_K_str);
	DDV_MaxChars(pDX, m_K_str, 15);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAddParamDlg, CDialog)
	//{{AFX_MSG_MAP(CAddParamDlg)
	ON_BN_CLICKED(IDC_MECH_AMP_MTC_ENTER_RADIO, OnMechEnterK)
	ON_BN_CLICKED(IDC_MECH_AMP_MTC_ESTIMATE_RADIO, OnMechEstimateK)
	ON_BN_CLICKED(IDC_MECH_AMP_MGL_RADIO, OnMechMglConc)
	ON_BN_CLICKED(IDC_MECH_AMP_VOL_RADIO, OnMechVolConc)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()



/////////////////////////////////////////////////////////////////////////////
// CAddParamDlg message handlers

void CAddParamDlg::OnMechEnterK() 
{
	// User selects button to enter mass transfer coefficient
	UpdateData(true);
	m_est_radio = -1;
	m_ent_radio = 0;
	UpdateData(false);
}

void CAddParamDlg::OnMechEstimateK() 
{
	// User selects button to estimate mass transfer coefficient
	UpdateData(true);
	m_est_radio = 0;
	m_ent_radio = -1;
	UpdateData(false);
}


void CAddParamDlg::OnMechMglConc() 
{
	// User selects button to enter concentration in mg/l
	UpdateData(true);
	m_mgl_radio = 0;
	m_vol_radio = -1;
	UpdateData(false);
}

void CAddParamDlg::OnMechVolConc() 
{
	// User selects button to enter concentration as volume fraction
	UpdateData(true);
	m_mgl_radio = -1;
	m_vol_radio = 0;
	UpdateData(false);
}
