#if !defined(AFX_EMPPICKMODELDLG_H__D3391D03_F9FE_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_EMPPICKMODELDLG_H__D3391D03_F9FE_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// EmpPickModelDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CEmpPickModelDlg dialog

class CEmpPickModelDlg : public CDialog
{
// Construction
public:
	CEmpPickModelDlg(CWnd* pParent = NULL);   // standard constructor

	int ModelType;	// 1=surface renewal,	2=fouling resistance,
					// 3=resistance,		4=gel polarization

// Dialog Data
	//{{AFX_DATA(CEmpPickModelDlg)
	enum { IDD = IDD_EMP_PICK_MODEL };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEmpPickModelDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CEmpPickModelDlg)
	afx_msg void OnEmpPickFoulRadio();
	afx_msg void OnEmpPickGelRadio();
	afx_msg void OnEmpPickRestRadio();
	afx_msg void OnEmpPickSurfRadio();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EMPPICKMODELDLG_H__D3391D03_F9FE_11D1_9A00_0020AFD5753F__INCLUDED_)
