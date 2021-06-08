#if !defined(AFX_MECHMODELSDLG_H__15F79AE1_B288_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_MECHMODELSDLG_H__15F79AE1_B288_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// MechModelsDlg.h : header file
//

#include "MemCond.h"
#include "MechModel.h"


/////////////////////////////////////////////////////////////////////////////
// CMechModelsDlg dialog

class CMechModelsDlg : public CDialog
{
// Construction
public:
	CMechModelsDlg(CWnd* pParent = NULL);   // standard constructor

////////////////////////////////////////////////////////////////////
// Class Wizard deletes comments between "{{AFX_DATA" comments when
//   it updates the variables, so the documentation for the grey 
//   code below must be here.  The variables are, in order: 
// Membrane operating conditions: pressure (Kpa), temperature (C),
//   feed water velocity (m/s), and water viscosity (kg/m*s)
// Average feed water conditions: concentration (mg/L),
//   density (g/cm^3), particle radius (microns)
// Membrane parameters: channel radius (mm), channel length (m),
//   membrane resistance (1/m), membrane area (m^2),
//   average pore size (microns)
////////////////////////////////////////////////////////////////////

// Dialog Data
	//{{AFX_DATA(CMechModelsDlg)
	enum { IDD = IDD_MECHANISTIC };
	double	m_press;
	double	m_temp;
	double	m_Cav;
	double	m_Dav;
	double	m_Rav;
	double	m_Area;
	double	m_CRad;
	double	m_CLen;
	double	m_MRes;
	double	m_PRad;
	double	m_visc;
	CString	m_MSFlux;
	CString	m_LHFlux;
	double	m_Circ;
	BOOL	m_find_reject;
	double	m_Qin;
	//}}AFX_DATA



// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMechModelsDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// For flux prediction, etc.
	CMemCond mem;
	CMechModel module;
	int model_type;
	int error_num;

	// Convert to the strings m_MSFlux and m_LHFlux
	double MSflux, LHflux;

	// Additional Model Parameters Dialogue Box variables
	bool amp_Cmgl;			// is concentration in mg/l?
	bool amp_estimate;		// does user wish to estimate k?
	double gel_conc;		// gel concentration in mg/l or volume fraction
	double user_k;			// user input value of	k
	double mgl_conc;		//						mg/l conc
	double vol_conc;		//						% conc
	double op_res;			// operational resistance
	double ir_res;			// irreversable (fouling) resistance

	// Particle Distribution Dialogue Box parameters
	void   CalcReject();
	double particles[2][20];
	double permeate[2][20];
	double retentate[2][20];
	int    part_num;


	// Parameter Range Dialogue Box Parameters
	bool   iterate;
	bool   save_iterate;
	int    param;			// 0=Press, 1=Temp, 2=Velocity, 3=Visc.
	short  steps;
	double increase;
	double range[2];		// high, low values for parameter
	double flux[2][20];		// flux vs parameter
	void   do_calcs();


	// Generated message map functions
	//{{AFX_MSG(CMechModelsDlg)
	afx_msg void OnMechPartDistribution();
	afx_msg void OnMechCalcFlux();
	afx_msg void OnMechAddtlModelParams();
	afx_msg void OnMechParamRange();
	afx_msg void OnMechMembSelect();
	afx_msg void OnMechMemsysRadio();
	afx_msg void OnMechGelRadio();
	afx_msg void OnMechResistanceRadio();
	afx_msg void OnMechSERadio();
	afx_msg void OnMechSave();
//	afx_msg void OnHelp();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
	bool CalcFlux();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MECHMODELSDLG_H__15F79AE1_B288_11D1_9A00_0020AFD5753F__INCLUDED_)
