// PlantRunDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "PlantRunDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPlantRunDlg dialog


CPlantRunDlg::CPlantRunDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CPlantRunDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPlantRunDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CPlantRunDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPlantRunDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPlantRunDlg, CDialog)
	//{{AFX_MSG_MAP(CPlantRunDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPlantRunDlg message handlers
