// PlantDesignDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "PlantDesignDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPlantDesignDlg dialog


CPlantDesignDlg::CPlantDesignDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CPlantDesignDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPlantDesignDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CPlantDesignDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPlantDesignDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPlantDesignDlg, CDialog)
	//{{AFX_MSG_MAP(CPlantDesignDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPlantDesignDlg message handlers
