#if !defined(AFX_PREDEFMEMDLG_H__6CA116C4_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_PREDEFMEMDLG_H__6CA116C4_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// PreDefMemDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPreDefMemDlg dialog

class CPreDefMemDlg : public CDialog
{
// Construction
public:
	CPreDefMemDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CPreDefMemDlg)
	enum { IDD = IDD_PREDEF_MEMB };
	CString	m_name;
	double	m_resist;
	double	m_crad;
	double	m_length;
	double	m_prad;
	double	m_area;
	CString	m_maker;
	CString	m_membrane;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPreDefMemDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CListBox  m_membList;
	int mem_num;

	// Generated message map functions
	//{{AFX_MSG(CPreDefMemDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnMembEnter();
	afx_msg void OnMembRemove();
	afx_msg void OnMembView();
	virtual void OnOK();
	afx_msg void OnDblclkMembList();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PREDEFMEMDLG_H__6CA116C4_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
