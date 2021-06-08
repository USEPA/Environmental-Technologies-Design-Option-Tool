// ModelVar.cpp: implementation of the CModelVar class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Uppmem.h"
#include "ModelVar.h"
#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CModelVar::CModelVar()
{
	kC = kP = kV = kT = 1;
	xC = xP = xV = xT = 0;
}

CModelVar::~CModelVar()
{

}


double CModelVar::get_val()
{
	return value;
}


double CModelVar::get_val(double C, double P, double V, double T)
{
	double val;
	val = value * kC * pow(C, xC) * kP * pow(P, xP) * kV * pow(V, xV) * kT * pow(T, xT);
	return val;
}



void CModelVar::set_val(double val)
{
	value = val;	
}



void CModelVar::set_val(int type, double k, double x)
{
	switch(type)
	{
	case 0:		kC = k;
				xC = x;
	break;

	case 1:		kP = k;
				xP = x;
	break;

	case 2:		kV = k;
				xV = x;
	break;

	case 3:		kT = k;
				xT = x;
	break;
	}
}



