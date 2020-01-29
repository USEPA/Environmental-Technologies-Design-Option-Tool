// EmpPickModelDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "EmpPickModelDlg.h"
#include "ErrBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEmpPickModelDlg dialog


CEmpPickModelDlg::CEmpPickModelDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CEmpPickModelDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CEmpPickModelDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT

	ModelType = 0;
}


void CEmpPickModelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEmpPickModelDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CEmpPickModelDlg, CDialog)
	//{{AFX_MSG_MAP(CEmpPickModelDlg)
	ON_BN_CLICKED(IDC_EMP_PICK_FOUL_RADIO, OnEmpPickFoulRadio)
	ON_BN_CLICKED(IDC_EMP_PICK_GEL_RADIO, OnEmpPickGelRadio)
	ON_BN_CLICKED(IDC_EMP_PICK_REST_RADIO, OnEmpPickRestRadio)
	ON_BN_CLICKED(IDC_EMP_PICK_SURF_RADIO, OnEmpPickSurfRadio)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEmpPickModelDlg message handlers

void CEmpPickModelDlg::OnEmpPickSurfRadio() 
{
	ModelType = 1;
}

void CEmpPickModelDlg::OnEmpPickFoulRadio() 
{
	ModelType = 2;
}

void CEmpPickModelDlg::OnEmpPickRestRadio() 
{
	ModelType = 3;
}

void CEmpPickModelDlg::OnEmpPickGelRadio() 
{
	ModelType = 4;
}


void CEmpPickModelDlg::OnOK() 
{
	if (ModelType == 0)
	{
		CErrBox box;		// error has occurred,
		box.SetErr(100);	// no model selected
		box.DoModal();		// show the error
		return;				// do not continue
	}

	// else
	CDialog::OnOK();
}
