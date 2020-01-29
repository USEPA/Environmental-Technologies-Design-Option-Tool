// ParamRangeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "ParamRangeDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CParamRangeDlg dialog


CParamRangeDlg::CParamRangeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CParamRangeDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CParamRangeDlg)
	m_steps = 1;
	m_hiPress = 100.0;
	m_loPress = 100.0;
	m_loTemp = 20.0;
	m_hiTemp = 20.0;
	m_loVisc = 0.001;
	m_hiVisc = 0.001;
	m_loConc = 2000.0;
	m_hiConc = 2000.0;
	m_loQin = 1800.0;
	m_hiQin = 1800.0;
	//}}AFX_DATA_INIT

	param = -1;

	// low and high values for the varied parameter
	range[0] = range[1] = 0;
}


void CParamRangeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CParamRangeDlg)
	DDX_Text(pDX, IDC_RANGE_NUM_STEPS, m_steps);
	DDV_MinMaxInt(pDX, m_steps, 1, 20);
	DDX_Text(pDX, IDC_RANGE_PRESS2, m_hiPress);
	DDV_MinMaxDouble(pDX, m_hiPress, 0., 1000000.);
	DDX_Text(pDX, IDC_RANGE_PRESS1, m_loPress);
	DDV_MinMaxDouble(pDX, m_loPress, 0., 1000000.);
	DDX_Text(pDX, IDC_RANGE_TEMP1, m_loTemp);
	DDV_MinMaxDouble(pDX, m_loTemp, 0., 100.);
	DDX_Text(pDX, IDC_RANGE_TEMP2, m_hiTemp);
	DDV_MinMaxDouble(pDX, m_hiTemp, 0., 100.);
	DDX_Text(pDX, IDC_RANGE_VISC1, m_loVisc);
	DDV_MinMaxDouble(pDX, m_loVisc, 0., 1.);
	DDX_Text(pDX, IDC_RANGE_VISC2, m_hiVisc);
	DDV_MinMaxDouble(pDX, m_hiVisc, 0., 1.);
	DDX_Text(pDX, IDC_RANGE_CONC1, m_loConc);
	DDV_MinMaxDouble(pDX, m_loConc, 0., 1000000.);
	DDX_Text(pDX, IDC_RANGE_CONC2, m_hiConc);
	DDV_MinMaxDouble(pDX, m_hiConc, 0., 1000000.);
	DDX_Text(pDX, IDC_RANGE_FLOW1, m_loQin);
	DDV_MinMaxDouble(pDX, m_loQin, 0., 1.e+020);
	DDX_Text(pDX, IDC_RANGE_FLOW2, m_hiQin);
	DDV_MinMaxDouble(pDX, m_hiQin, 0., 1.e+020);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CParamRangeDlg, CDialog)
	//{{AFX_MSG_MAP(CParamRangeDlg)
	ON_BN_CLICKED(IDC_RANGE_PRESS_RADIO, OnRangePressRadio)
	ON_BN_CLICKED(IDC_RANGE_TEMP_RADIO, OnRangeTempRadio)
	ON_BN_CLICKED(IDC_RANGE_VISC_RADIO, OnRangeViscRadio)
	ON_BN_CLICKED(IDC_RANGE_CONC_RADIO, OnRangeConcRadio)
	ON_BN_CLICKED(IDC_RANGE_FLOW_RADIO, OnRangeFlowRadio)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CParamRangeDlg message handlers

void CParamRangeDlg::OnRangePressRadio() 
{
	param = 0;	
}

void CParamRangeDlg::OnRangeTempRadio() 
{
	param = 1;	
}

void CParamRangeDlg::OnRangeFlowRadio() 
{
	param = 2;
}

void CParamRangeDlg::OnRangeViscRadio() 
{
	param = 3;	
}

void CParamRangeDlg::OnRangeConcRadio() 
{
	param = 4;	
}

