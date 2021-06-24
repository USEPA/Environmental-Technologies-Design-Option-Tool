#if !defined(AFX_PLANTDESIGNDLG_H__15F79AE3_B288_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_PLANTDESIGNDLG_H__15F79AE3_B288_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// PlantDesignDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPlantDesignDlg dialog

class CPlantDesignDlg : public CDialog
{
// Construction
public:
	CPlantDesignDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CPlantDesignDlg)
	enum { IDD = IDD_DESIGN };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPlantDesignDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CPlantDesignDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PLANTDESIGNDLG_H__15F79AE3_B288_11D1_9A00_0020AFD5753F__INCLUDED_)
