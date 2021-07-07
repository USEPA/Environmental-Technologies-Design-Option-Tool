#if !defined(AFX_ADDPARAMDLG_H__6CA116C1_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_ADDPARAMDLG_H__6CA116C1_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// AddParamDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddParamDlg dialog

class CAddParamDlg : public CDialog
{
// Construction
public:
	CAddParamDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddParamDlg)
	enum { IDD = IDD_ADDTL_PARAMS };
	int		m_mgl_radio;
	int		m_est_radio;
	int		m_vol_radio;
	int		m_ent_radio;
	double	m_ir_res;
	double	m_op_res;
	double	m_mgl_conc;
	double	m_vol_conc;
	CString	m_K_str;
	//}}AFX_DATA

/*	Variable values:
	m_K_str:	mass transfer coefficient in m/s
	m_mgl_conc:	gel concentration in mg/L
	m_vol_conc:	gel concentration in %
	m_ir_res:	irreversable fouling resistances
	m_op_res:	pressure-dependant resistances
*/



// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddParamDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddParamDlg)
	afx_msg void OnMechEnterK();
	afx_msg void OnMechEstimateK();
	afx_msg void OnMechMglConc();
	afx_msg void OnMechVolConc();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDPARAMDLG_H__6CA116C1_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
