#if !defined(AFX_EMPMODEL_H__402E5262_0763_11D2_9A00_0020AFD5753F__INCLUDED_)
#define AFX_EMPMODEL_H__402E5262_0763_11D2_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// EmpModel.h : header file
//

#include "MemCond.h"
#include "EmpData.h"
#include "EnterDataDlg.h"
#include <fstream.h>

/////////////////////////////////////////////////////////////////////////////
// CEmpModel window

class CEmpModel : public CWnd
{

public:
	CMemCond mem;
	CEmpModel();
	void set_Emp_model(int type);	// 1=surf ren, 2=fouling, 3=resistance, 4=gel
	void get_Emp_model(int &type);
	void insert_tail(CEmpData data);	// insert a data set
	double param_a, param_b, param_c;	// dummy model params


private:
	int	ModelType;	// 1=surf ren, 2=fouling, 3=resistance, 4=gel

	// a queue of CEmpData objects, each with a range of validity,
	//   a list of data points, and model parameters
	CList<CEmpData, CEmpData&> DataList;

	CList<double, double&> xList, yList;		// used by curve fitting algorithms, x vs. y

	bool set;		// were parameters set by the user?

	// used to find parameters for viable CEmpData objects
	int surf_fit(CEmpData &dead_end, CEmpData &x_flow);
	int foul_fit(CEmpData &ad_foul, CEmpData &pp_foul, CEmpData &cp_foul);
	int  res_fit(CEmpData &data);
	int  gel_fit(CEmpData &data);

	// the 4 curve-fitting algorithms
	int xList_ave(double &ave);
	int line(double &m, double &b);
	int resistance(double &R, double &theta);
	int surf_renewal(const CEmpData data, double &A, double &Js);

	// used to merge a non-viable with a viable CEmpData object
	//   to obtain C,V,T,P-dependance on model coefficients
	int surf_merge();
	int foul_merge();
	int res_merge();
	int gel_merge();



public:

	// always make a new CEmpData object for new data sets,
	//   if primary parameter is not used, make CEmpData object
	//   non-viable.  Called when additional data is entered
	void add_data(CEnterDataDlg data);

	// first, build all primary models for objects in CList 
	// second, see if any other data sets could help each model.
	//   If so, create a NEW CEmpData object with NO DATA, just
	//   model parameters, and a range.
	int build_model();


	// getting the flux from given data sets
	int get_flux(double &flux, int &compat, double C, double P, 
			 double V, double T, double visc, double TIME);

	// set the parameters so user can predict flux
	int set_param(double p_a, double p_b, double p_c);

	// get the gel polarization flux
	int get_gel_flux(double &flux, double k, double Cg, double Cb);

	// get the surface renewal flux
	int get_surf_flux(double &flux, double Js, double A, double s, double t);

	// overload the "<<" operator, so you can easily print it out
	friend fstream& operator<< (fstream& os, const CEmpModel& output_data);


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEmpModel)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CEmpModel();

	// Generated message map functions
protected:
	//{{AFX_MSG(CEmpModel)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EMPMODEL_H__402E5262_0763_11D2_9A00_0020AFD5753F__INCLUDED_)
