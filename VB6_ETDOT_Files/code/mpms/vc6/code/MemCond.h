// MemCond.h: interface for the CMemCond class.
//
//////////////////////////////////////////////////////////////////////


#if !defined(AFX_MEMCOND_H__59A621E1_B8E0_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_MEMCOND_H__59A621E1_B8E0_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000



#ifndef CMEMCOND_H
#define CMEMCOND_H


class CMemCond  
{

private:
	double flux;		// flux in m/s

public:
	CMemCond();				// constructor
	virtual ~CMemCond();	// destructor: not needed

	// Sets the particle distribution for influent and
	//   effluent streams based on the flux
	bool SetDstb(void);		

	// Determines the average particle size and concentration
	//   based on the distribution
	bool SetAves(void);

	// Sets the averages if the user chooses not to go with
	//   a particle distribution
	bool SetAves(double rad, double conc, double den);

	// Sets up the system parameters
	bool SetParam(double p, double t, double vl, double vc);

	// Sets up the influent particle distribution
	bool SetDstb(double influent[2][20], double effluent[2][20],
				 double retentate[2][20], int num);

	// Sets up the membrane size parameters
	bool SetMem(double PR, double RS, double CR, 
				double CL, double MA, double RC);

	// Sets up the flux after a model determines it
	bool SetFlux(double F);

	// Shows the flux value to any client
	double ShowFlux() const;



// protected:
public:

	// system parameters
	double press;		// pressure in Pa
	double temp;		// temperature in K
	double veloc;		// velocity in m/s
	double visc;		// absolute viscosity in kg/m*s


	// feed water parameters: the influent distrinution
	//   is user-defined and is used to calculate Rav and
	//   Cav.  Eff & RtnDstb are water rejection dependant

	double Rav;				// average particle radius in m
	double Cav;				// average concentration in #/ml
	double Dav;				// average density in mg/L
	double InfDstb[2][20];	// size, conc distribution: influent
	double EffDstb[2][20];	//							effluent
	double RtnDstb[2][20];  //							retentate
	int    NumDstb;			// number of particles in the distribution

	// membrane selection parameters
	double PRad;			// average pore radius in m
	double MRes;			// membrane resistance in 1/m
	double CRad;			//			channel radius in m 
	double CLen;			//			channel length in m
	double Area;			//			area in m^2
	double Circ;			// water recirculation (Qin/Qr), fraction
	

};

#endif		// CMEMCOND_H

#endif // !defined(AFX_MEMCOND_H__59A621E1_B8E0_11D1_9A00_0020AFD5753F__INCLUDED_)
