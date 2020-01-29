// Uppmem.h : main header file for the UPPMEM application
//

#if !defined(AFX_UPPMEM_H__E1139ACD_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_UPPMEM_H__E1139ACD_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CUppmemApp:
// See Uppmem.cpp for the implementation of this class
//

class CUppmemApp : public CWinApp
{
public:
	CUppmemApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CUppmemApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CUppmemApp)
	afx_msg void OnAppAbout();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_UPPMEM_H__E1139ACD_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_)
