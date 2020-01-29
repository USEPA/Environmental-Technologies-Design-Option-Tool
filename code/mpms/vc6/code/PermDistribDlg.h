#if !defined(AFX_PERMDISTRIBDLG_H__77CA5A03_61C4_11D2_9A00_0020AFD5753F__INCLUDED_)
#define AFX_PERMDISTRIBDLG_H__77CA5A03_61C4_11D2_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "MemCond.h"

// PermDistribDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPermDistribDlg dialog

class CPermDistribDlg : public CDialog
{
// Construction
public:
	CPermDistribDlg(CWnd* pParent = NULL);   // standard constructor
	void set_list(CMemCond mem);
	double permeate[2][20];		// size (microns), conc (#/ml)
	double retentate[2][20];
	int part_num;

// Dialog Data
	//{{AFX_DATA(CPermDistribDlg)
	enum { IDD = IDD_PERM_PART_DISTRIB_DLG };
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPermDistribDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	CListBox m_permList;
	CListBox m_retenList;
	
	// Generated message map functions
	//{{AFX_MSG(CPermDistribDlg)
		// NOTE: the ClassWizard will add member functions here
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PERMDISTRIBDLG_H__77CA5A03_61C4_11D2_9A00_0020AFD5753F__INCLUDED_)
