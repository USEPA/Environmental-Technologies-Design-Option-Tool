#if !defined(AFX_EMPMODELSDLG_H__15F79AE2_B288_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_EMPMODELSDLG_H__15F79AE2_B288_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// EmpModelsDlg.h : header file
//

#include "EmpModel.h"


/////////////////////////////////////////////////////////////////////////////
// CEmpModelsDlg dialog

class CEmpModelsDlg : public CDialog
{
// Construction
public:
	CEmpModelsDlg(CWnd* pParent = NULL);   // standard constructor
	
	// this object contains all information about the membrane setup
	CEmpModel module;		
	bool mgl;				// are conc units mgl?


	// variables needed to test a range of parameters
	bool   iterate;
	bool   custom;
	int    param;			// 0=Press, 1=Temp, 2=Velocity, 3=Visc., 4=conc
	int    ModelType;
	short  steps;
	double increase;
	double range[2];		// high, low values for parameter
	double flux[2][20];		// flux vs parameter

	// sets the units to mg/l or %vol, and the default params
	void set_units(bool mgl, double C, double P, double T, double V);

	// determine flux if the user inputs the values
	int get_flux(double &flux);

	// round all off the values updated by the computer
	void RoundAll();

// Dialog Data
	//{{AFX_DATA(CEmpModelsDlg)
	enum { IDD = IDD_EMPIRICAL };
	double	m_conc;
	CString	m_conc_units;
	double	m_pres;
	double	m_temp;
	double	m_visc;
	double	m_lhFlux;
	double	m_msFlux;
	double	m_tFlux;
	double	m_cTime;
	double	m_pTime;
	CString	m_a_name;
	CString	m_b_name;
	CString	m_c_name;
	double	m_a_val;
	double	m_b_val;
	double	m_c_val;
	double	m_Qin;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEmpModelsDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CEmpModelsDlg)
	afx_msg void OnEmpExpData();
	afx_msg void OnEmpParamRange();
	afx_msg void OnEmpCalcFlux();
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnUseCustomVal();
	afx_msg void OnEmpSave();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EMPMODELSDLG_H__15F79AE2_B288_11D1_9A00_0020AFD5753F__INCLUDED_)
