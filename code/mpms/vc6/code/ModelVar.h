// ModelVar.h: interface for the CModelVar class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MODELVAR_H__402E5263_0763_11D2_9A00_0020AFD5753F__INCLUDED_)
#define AFX_MODELVAR_H__402E5263_0763_11D2_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class CModelVar  
{
public:
	CModelVar();
	virtual ~CModelVar();

	void set_val(double val);		// set the value

	// obtain the value for a given conc, press, velocity, and temp
	double get_val(double C, double P, double V, double T);

	// obtain the original value
	double get_val();

	// overloaded - set either range or constants
	//   type:  1=C, 2=P, 3=V, 4=T
	void set_val(int type, double k, double x);		

private:
	// the value it is first set to
	double value;

	// the way to alter 'value' if it has a range,
	// value' = value * (kC*C^xC) * (kP*P^xP) * (kV*V^xV) * (kT*T^xT)
	double kC, kP, kV, kT;		// scalar multiples
	double xC, xP, xV, xT;		// exponents


};

#endif // !defined(AFX_MODELVAR_H__402E5263_0763_11D2_9A00_0020AFD5753F__INCLUDED_)
