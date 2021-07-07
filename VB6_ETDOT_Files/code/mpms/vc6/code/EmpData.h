#if !defined(AFX_EMPDATA_H__B86703C1_0B6F_11D2_9A00_0020AFD5753F__INCLUDED_)
#define AFX_EMPDATA_H__B86703C1_0B6F_11D2_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// EmpData.h : header file
//


#include "ModelVar.h"
#include "EnterDataDlg.h"
#include <afxtempl.h>

/////////////////////////////////////////////////////////////////////////////
// Point class, needed for CEmpData objects

class Point
{
public:
	double	flux,	// in m/s 
			conc,   // in mg/L OR %
			pres,	// in Pa
			vlos,	// in m/s
			temp,	// in C
			time;	// in s

	// default constructor
	Point()		
	{
		flux = conc = pres = vlos = temp = time = 0;
	}

	// regular constructor
	Point(double F, double C, double P,
		  double V, double T, double TIME)
	{
		flux = F;
		conc = C;
		pres = P;
		vlos = V;
		temp = T;
		time = TIME;
	}

};




/////////////////////////////////////////////////////////////////////////////
// CEmpData window
class CEmpData : public CWnd
{

public:
	int	ModelType;	// 1=surf ren, 2=fouling, 3=resistance, 4=gel

	// the range of values for which data is valid, or -1 in first
	// position if its valid for only one parameter value
	double	conc[2], pres[2], vlos[2], temp[2], time[2];
	// units	(%)		 (kPa)    (m/s)	   (C)		(s)		

	CList<Point, const Point&> Points;	// a CList of data points
	bool viable;			// is this object just a collection of data, or a viable model?
	double tolerance;		// 2% tolerance:  100 = 102-98 with 2% tolerance

	CEmpData();							// Constructor
	CEmpData(const CEmpData &rhs);		// Copy constructor
	void set_type(int type);			// Set model type


	bool primary;	// is this data set primary with respect to ModelType?
	int  param;		// varied parameter: 1=conc, 2=pres, 3=vlos,
					//					 4=temp, 5=time

// Member functions
	// inserts a new point into the list
	bool insert(int var, double F, double C, double P,
		  double V, double T, double TIME);

	// Extract data from dialoge box object, set ranges
	bool set_data(const CEnterDataDlg& data);	

	// the range compatibility test
	int in_range(double C, double P, double V, 
				double T, double TIME);	

	// equality operator:
	const CEmpData& operator=(const CEmpData &rhs);



// Variables for the necessary models
	// Surface Renewal, #1
	CModelVar	A,			// rate of flux decline (1/s)
				Js,			// minimum dead-end flux (m/s)
				s;			// rate of surface renewal (1/s)

	// Fouling Resistance Model, #2
	CModelVar	Ra,			// adsorption fouling resistance (1/m)
				Rpp,		// pore plugging resistance (1/m)
				Rcp;		// conc. polarization resistance (1/m)

	// Resistance Model, #3
	CModelVar	Rf,			// fouling layer resistance (1/m)
				Rop;		// pressure dependant resistances (1/Pa)

	// Gel Model, #4
	CModelVar	k,			// mass transfer coefficient (m/s)
				Cg;			// gel concentration (% or mg/L)



// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEmpData)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CEmpData();

	// Generated message map functions
protected:
	//{{AFX_MSG(CEmpData)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EMPDATA_H__B86703C1_0B6F_11D2_9A00_0020AFD5753F__INCLUDED_)
