// FluxRangeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "FluxRangeDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CFluxRangeDlg dialog


CFluxRangeDlg::CFluxRangeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CFluxRangeDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CFluxRangeDlg)
	m_FluxEdit = _T("");
	m_ParamEdit = _T("");
	m_ParamName = _T("");
	//}}AFX_DATA_INIT
}


void CFluxRangeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFluxRangeDlg)
	DDX_Text(pDX, IDC_RANGE_FLUX_EDIT, m_FluxEdit);
	DDV_MaxChars(pDX, m_FluxEdit, 1000);
	DDX_Text(pDX, IDC_RANGE_PARAM_EDIT, m_ParamEdit);
	DDV_MaxChars(pDX, m_ParamEdit, 1000);
	DDX_Text(pDX, IDC_RANGE_PARAM_NAME, m_ParamName);
	DDV_MaxChars(pDX, m_ParamName, 20);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CFluxRangeDlg, CDialog)
	//{{AFX_MSG_MAP(CFluxRangeDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFluxRangeDlg message handlers

// create the list to print out
void CFluxRangeDlg::set_param(int param, int num, double range[][20])
{
	// initialize the values for m_param, m_num, & m_range
	m_param = param;
	m_num   = num;
	for (int i=0; i<num; i++)
	{
		m_range[0][i] = range[0][i];
		m_range[1][i] = range[1][i];
	}
}


// initialization of the dialogue box
BOOL CFluxRangeDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	int index = 0;

	switch(m_param)
	{
	case 0:		m_ParamName = "Pressure (kPa):";
	break;
	case 1:		m_ParamName = "Temperature (C):";
	break;
	case 2:		m_ParamName = "Flow (L/h):";
	break;
	case 3:		m_ParamName = "Viscosity (kg/m*s):";
	break;
	case 4:		m_ParamName = "Concentration (mg/L):";
	}

	char buffer[50];

	for (int i=0; i<m_num; i++)
	{
		index++;

		// convert the number to a character string, and insert
		//   it into the respective edit box
		_gcvt((m_range[0][i]), 5, buffer);
		m_FluxEdit += (CString)buffer;
		m_FluxEdit += '\t';

		_gcvt((m_range[1][i]), 5, buffer);
		m_ParamEdit += (CString)buffer;
		m_ParamEdit += '\t';
		m_ParamEdit += '\t';
	}
	
	UpdateData(false);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
