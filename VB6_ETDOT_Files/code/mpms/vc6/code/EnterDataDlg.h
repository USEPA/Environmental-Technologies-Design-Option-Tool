#if !defined(AFX_ENTERDATADLG_H__D3391D01_F9FE_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_ENTERDATADLG_H__D3391D01_F9FE_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// EnterDataDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CEnterDataDlg dialog

class CEnterDataDlg : public CDialog
{
// Construction
public:
	CEnterDataDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CEnterDataDlg)
	enum { IDD = IDD_EMP_DATA_ENTER };
	double	m_conc;
	CString	m_flux;
	CString	m_param;
	CString	m_param_name;
	double	m_pres;
	double	m_temp;
	double	m_vlos;
	int		m_conc_radio;
	int		m_pres_radio;
	int		m_temp_radio;
	int		m_time_radio;
	int		m_vlos_radio;
	CString	m_info;
	int		m_mgl_radio;
	int		m_vol_radio;
	//}}AFX_DATA

// Dialogue box objects
	CEdit conc_edit, pres_edit, vlos_edit, temp_edit;


	int ModelType;		// 1=surf, 2=foul, 3=res, 4=gel
	int param;			// 1=conc, 2=pres, 3=vlos, 4=temp, 5=time
	bool primary;		// primary data set?
	bool mgl;			// are conc units in mgl?  
	bool alter;			// is user allowed to alter parameters in any way?

	// set the Model Type, and primary value
	void set_data(int type, bool prime = false);

	// set the parameters if this is a second or third data set
	void set_param(double C, double P, double V, double T, bool M);

	// check to see if parameter and model type match
	bool check(int prm);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEnterDataDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CEnterDataDlg)
	afx_msg void OnEmpEntConcRadio();
	afx_msg void OnEmpEntMglRadio();
	afx_msg void OnEmpEntPresRadio();
	afx_msg void OnEmpEntTempRadio();
	afx_msg void OnEmpEntVlosRadio();
	afx_msg void OnEmpEntVolRadio();
	afx_msg void OnEmpEntTimeRadio();
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ENTERDATADLG_H__D3391D01_F9FE_11D1_9A00_0020AFD5753F__INCLUDED_)
