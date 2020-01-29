#if !defined(AFX_PARAMRANGEDLG_H__6CA116C3_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_PARAMRANGEDLG_H__6CA116C3_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// ParamRangeDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CParamRangeDlg dialog

class CParamRangeDlg : public CDialog
{
// Construction
public:
	CParamRangeDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CParamRangeDlg)
	enum { IDD = IDD_PARAM_RANGE };
	short	m_steps;
	double	m_hiPress;
	double	m_loPress;
	double	m_loTemp;
	double	m_hiTemp;
	double	m_loVisc;
	double	m_hiVisc;
	double	m_loConc;
	double	m_hiConc;
	double	m_loQin;
	double	m_hiQin;
	//}}AFX_DATA

	// 0=pressure, 1=temp, 2=velocity, 3=viscosity, 4=conc
	int param;

	// low and high values for the varied parameter
	double range[2];

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CParamRangeDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CParamRangeDlg)
	afx_msg void OnRangePressRadio();
	afx_msg void OnRangeTempRadio();
	afx_msg void OnRangeViscRadio();
	afx_msg void OnRangeConcRadio();
	afx_msg void OnRangeFlowRadio();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PARAMRANGEDLG_H__6CA116C3_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
