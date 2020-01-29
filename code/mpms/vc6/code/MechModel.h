// MechModel.h: interface for the CMechModel class.
//
//////////////////////////////////////////////////////////////////////


#if !defined(AFX_MECHMODELS_H__59A621E3_B8E0_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_MECHMODELS_H__59A621E3_B8E0_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


#ifndef CMECHMODEL_H
#define CMECHMODEL_H

#include "MemCond.h"	// include the files about membrane conditions
#include <Math.h>		// for pow and fabs functions



class CMechModel : public CMemCond
{

private:
	CMemCond mem;		// Access to variables and a place to
						//   return flux when done
	int ModelType;		// which model to use when
						//   determining flux

	// Global Variables:
	double Kb;		// Boltzman's constant, J/K
	double pi;		// pi


public:
	// items from Additional Model Parameters
	double Blt;		// Resistance:	boundary layer thickness

	double Cg;		// Gel polar.:	gel concentration in mg/L or volume fraction
	double Kg;		//				mass transfer coefficient in m/s
    bool   Cmgl;	//				concentration in mg/L?

	double op_res;	// pressure-dependant operational resistance
	double ir_res;	// constant irreversable resistance


	CMechModel();						// default constructor
	virtual ~CMechModel();				// destructor

	// initializes the type and CMemCond object
	void MechModelInit(int type, CMemCond mc);

	// sets the private member variables for gel polarization
	bool SetRes(double op, double ir);

	// sets the private member variables for gel polarization
	bool SetGel(double gel, bool mgl, double mtc);

	// calculates the mass transfer coefficient
	void SetMTC(double veloc);

	// "switching station" for flux models, the number returned
	//   represents the error, if any, occurred:  1 = okay
	//   -1 = negative flux		0 = unknown flux model			
	//   Model specific errors are listed in source code,
	//   but 1x errors are memsys errors,
	//       2x errors are SE errors,
	//       3x errors are resistance errors,
	//       4x errors are gel polarization errors.
	int FindFlux();

	// Returns the flux value
	double GetFlux();

	// Determine flux with MEMSYS, #1
	int MEMSYSFlux(double &ms_flux) const;

	// Determine flux with Song and Elimelech model, #2
	int SEFlux(double &se_flux) const;		// main routine
	int do_SEFlux(double &se_flux, double conc, double vlos) const;		// looped routine

	// Determine flux with the resistance model, #3
	int ResFlux(double &rs_flux) const;

	// Determine flux with gel polarization model, #4
	int GelFlux(double &gl_flux);

};


#endif		// CMECHMODEL_H

#endif // !defined(AFX_MECHMODELS_H__59A621E3_B8E0_11D1_9A00_0020AFD5753F__INCLUDED_)
