#if !defined(AFX_PARTDISTRIBDLG_H__6CA116C2_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_PARTDISTRIBDLG_H__6CA116C2_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// PartDistribDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPartDistribDlg dialog

class CPartDistribDlg : public CDialog
{
// Construction
public:
	void SetList(double particles[2][20], int num, double den);
	CPartDistribDlg(CWnd* pParent = NULL);   // constructor

	// the list of particle sizes, and number of particles
	double m_particles[2][20];		// radius in m, concentration in #/ml
	double m_density;					// from MechModelDlg
	int    m_partNum;
	int    conc;	// 0 = #/ml, 1 = volume fraction, 2 = mg/L

	// average values
	double ave_rad;		// in microns
	double ave_conc;	// in mg/L

// Dialog Data
	//{{AFX_DATA(CPartDistribDlg)
	enum { IDD = IDD_DISTRIB };
	double	m_massConc;
	double	m_partRad;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPartDistribDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// the list box in the dialogue box
	CListBox  m_partList;

	// Generated message map functions
	//{{AFX_MSG(CPartDistribDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnDistribView();
	afx_msg void OnDistribRemove();
	afx_msg void OnDistribEnter();
	afx_msg void OnDblclkParticleList();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PARTDISTRIBDLG_H__6CA116C2_CA3C_11D1_9A00_0020AFD5753F__INCLUDED_)
