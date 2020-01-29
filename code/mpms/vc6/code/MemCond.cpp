// MemCond.cpp: implementation of the CMemCond class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Uppmem.h"
#include "MemCond.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#ifndef CMEMCOND_CPP
#define CMEMCOND_CPP

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMemCond::CMemCond()
{
	flux = 0;
}



CMemCond::~CMemCond()
{
	// nothing to do here...
	//   class contains no dynamic memory
}



//////////////////////////////////////////////////////////////////////
// Member Variable Initialization
//////////////////////////////////////////////////////////////////////
	

// Sets the particle distribution for influent and
//   effluent streams based on the flux
bool CMemCond::SetDstb(void)
{
	// must determine rejection of each particle one at a 
	//   time.  Then, using water rejection as a parameter,
	//   determine the new concentration distribution for
	//   retentate and permeate streams

	// more applicable to the virtual plant design, skip for now
	return true;
}



// Determines the average particle size and concentration based
//   on the distribution.  Called after object is fully built.
//   A simple weighted average based on concentration
bool CMemCond::SetAves(void)
{
	return true;
}


// Sets the average concentration and particle size if the user does not
//   use a distribution.  Do type conversion before calling this function
bool CMemCond::SetAves(double rad, double conc, double den)
{
	if (rad>0.0 && conc>0.0 && den>0.0)
	{
		Rav = rad;
		Cav = conc;
		Dav = den;
		return true;
	}

	else
		return false;
}



// Sets up the system parameters.  Should do type conversions
//   before calling this function
bool CMemCond::SetParam(double p, double t, double vl, double vc)
{
	// check for valid data
	if (p>0.0 && t>0.0 && vl>0.0 && vc >0.0 )
	{
		press  = p;
		temp   = t;		
		veloc  = vl;	
		visc   = vc;	
		return true;
	}

	else
		return false;
}



// Sets up the influent particle distribution
bool CMemCond::SetDstb(double influent[2][20], double effluent[2][20],
					   double retentate[2][20], int num)
{
	bool valid = true;
	int i, j;

	NumDstb = num;

	// check for valid data; if valid, enter it.
	//   if not, set return flag to false.
	//   Zero all values not used.
	for (i = 0; i<2; i++)
	{
		for (j = 0; j<num; j++)
		{
			if (influent[i][j] < 0 || effluent[i][j] < 0 || retentate[i][j] < 0)
				valid = false;
			else
			{
				InfDstb[i][j] = influent[i][j];
				EffDstb[i][j] = effluent[i][j];
				RtnDstb[i][j] = retentate[i][j];
			}
		}		// end for

		for (j = num; j < 20; j++)
		{
			InfDstb[i][j] = 0;
			EffDstb[i][j] = 0;
			RtnDstb[i][j] = 0;
		}
	}			// end for

	return valid;

}



// Sets up the membrane size parameters.  Do type conversions before
//   calling this function
bool CMemCond::SetMem(double PR, double RS, double CR, 
					  double CL, double MA, double RC)
{
	// check for valid data
	if (PR>0 && RS>0 && CR>0 && CL>0 && MA>0 && RC>=0)
	{
		PRad	= PR;
		MRes	= RS;
		CRad	= CR;
		CLen	= CL;
		Area    = MA;
		Circ    = RC;
		return true;
	}

	else
		return false;
}



// Assignes a value to the Flux
bool CMemCond::SetFlux(double F)
{
	// establish validity
	if (F>=0)
	{
		flux = F;
		return true;
	}

	else
		return false;
}



// Returns the flux
double CMemCond::ShowFlux() const
{
	return flux;
}



#endif			// CMEMCOND_CPP