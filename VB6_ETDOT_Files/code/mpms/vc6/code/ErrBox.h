#if !defined(AFX_ERRBOX_H__40C31F21_CECA_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_ERRBOX_H__40C31F21_CECA_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include <string>

// ErrBox.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CErrBox dialog

class CErrBox : public CDialog
{

// Construction
public:
	void SetErr(int val);  // sets the errors to be displayed
	CErrBox(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CErrBox)
	enum { IDD = IDD_ERROR };
	CString	m_ErrMsg;
	CString	m_ErrMsgB;
	CString	m_ErrMsgC;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CErrBox)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CErrBox)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};




//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ERRBOX_H__40C31F21_CECA_11D1_9A00_0020AFD5753F__INCLUDED_)
