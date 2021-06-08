#if !defined(AFX_FLUXRANGEDLG_H__8D210761_FEDB_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_FLUXRANGEDLG_H__8D210761_FEDB_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// FluxRangeDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CFluxRangeDlg dialog

class CFluxRangeDlg : public CDialog
{
// Construction
public:
	void set_param(int param, int num, double[][20]);
	CFluxRangeDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CFluxRangeDlg)
	enum { IDD = IDD_FLUX_RANGE };
	CString	m_FluxEdit;
	CString	m_ParamEdit;
	CString	m_ParamName;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFluxRangeDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	int    m_param;			// which parameter has the range of values
	int    m_num;			// how many steps
	double m_range[2][20];	// flux vs param for up to 20 steps

	// Generated message map functions
	//{{AFX_MSG(CFluxRangeDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FLUXRANGEDLG_H__8D210761_FEDB_11D1_9A00_0020AFD5753F__INCLUDED_)
