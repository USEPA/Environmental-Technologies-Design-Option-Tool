// MechModel.cpp: implementation of the CMechModel class.
//
//////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "Uppmem.h"
#include "MechModel.h"
#include <iostream.h>
#include <fstream.h>
#include <iomanip.h>
#include <math.h> 
#include <stdlib.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#ifndef CMECHMODEL_CPP
#define CMECHMODEL_CPP


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMechModel::CMechModel()
{
	ModelType = 0;
	Cmgl = false;
	Kb = 1.38E-23;		// J/K
	pi = 3.1416;
}



CMechModel::~CMechModel()
{
	//  nothing happens here!
}



void CMechModel::MechModelInit(int type, CMemCond mc)
{
	if (type > 0  &&  type < 5)
		ModelType = type;
	else
		ModelType = 0;
	mem = mc;
}



// calls the appropriate flux model based on which was chosen
//   in the CMechModelsDlg dialogue box
int CMechModel::FindFlux()
{
	int flag = 0;			// another check for poor data entered
	double temp_flux;		// used to set the flux


	switch (ModelType)
	{
		case 1: flag += MEMSYSFlux(temp_flux);
				break;

		case 2: flag += SEFlux(temp_flux);
				break;

		case 3: flag += ResFlux(temp_flux);
				break;

		case 4: flag += GelFlux(temp_flux);
				break;

		default:return flag;	// unknown model choice
				
	}   // end switch


	if (mem.SetFlux(temp_flux))		// if it sets the flux w/out error
		return flag;				// return the other warnings
	else
		return (-1);		//  flux is negative
}


// Returns the flux
double CMechModel::GetFlux()
{
	double f;
	f = mem.ShowFlux();   // ShowFlux from the MemCond object
	return f;
}


//////////////////////////////////////////////////////////////////////
// Determine flux with MEMSYS
int CMechModel::MEMSYSFlux(double &ms_flux) const
{
/*	Values for errors, warnings passed by Memsys:
		flag = 1		No error
		flag = 10		Memsys didn't load	
		flag = 11		User didn't save flux values to default file */

	int flag = 1;

// To run Memsys with the front end, we will first use file i/o
//   to alter the default files that Memsys loads at startup, 
//   making the parameters the same as the ones the user input.  
// Then we open Memsys, allow the user to run it with the default 
//   data, and do anything else with it they wish.  They should
//   save the flux values to a file, and close Memsys
// Finally, this program will scan the data file for flux, and 
//   set it to ms_flux

// Step 1:  Alter default Memsys files 
	// convert necessarry variables for particle list
	double size = mem.Rav * 1000000;	// m to microns
	double conc = mem.Cav;				// no conversion
	ofstream fout;
	fout.open("MS-PART.DAT");			// open default memsys file to overwrite present values
	fout << setw(12) << 1 << endl;		// write formatted numbers to the file
	fout << setw(25) << size;
	fout << setw(25) << conc << endl << endl;
	fout.close();						// close the file

// convert variables for membrane conditions
	double diam = mem.CRad * 2 * 1000;	// radius from m to diam in mm
	double len  = mem.CLen * 100;		// length from m to cm
	double pres = mem.press / 100000;	// pressure from Pa to Bars
	double vlos = mem.veloc * 100;		// speed from m/s to cm/s
	double temp = mem.temp - 273;		// temp from K to C
	double porD = mem.PRad * 1000000;	// pore diameter from m to microns
	double rest = mem.MRes / 100;		// membrane resistance from 1/m to 1/cm
	double dens = 2;				// density - Uppmem does all necessarry conversions for Memsys, this value is not needed, so arbitrarilly set	
	double volf = .58;				// max volume fraction, not input by user
	double rec  = 95;				// water recovery

// open global values default file, overwrite with these numbers:
	fout.open("MS-GLOB.DAT");
	fout << setw(25) << diam << setw(25) << len << setw(25) << pres << endl;
	fout << setw(25) << vlos << setw(25) << rec << setw(25) << temp << endl;
	fout << setw(25) << porD << setw(25) << rest<< setw(25) << dens << endl;
	fout << setw(25) << volf << endl << endl;
	fout.close();						// close the file


// Step 2:  Open and run Memsys
	// pass by reference values needed to initialize CreateProcess
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	ZeroMemory( &si, sizeof(si) );
	si.cb = sizeof(si);
	
	// open the memsys window 
	if( !CreateProcess( NULL,	// Name for 32 bit app
			"memsys.exe",		// Command line, or 16 bit app name
			NULL,		// Process handle not inheritable. 
			NULL,		// Thread handle not inheritable. 
			FALSE,		// Set handle inheritance to FALSE. 
			0,			// No creation flags. 
			NULL,		// Use parent’s environment block. 
			NULL,		// Use parent’s starting directory. 
			&si,		// Pointer to STARTUPINFO structure.
			&pi )		// Pointer to PROCESS_INFORMATION structure.
			) 
	{
		flag = 10;	// Memsys didn't load if (!CreateProcess)
		ms_flux = 0.0;	// set flux to zero
		return flag;
	}

	// Wait until memsys exits.
	WaitForSingleObject( pi.hProcess, INFINITE );

	// Close memsys and thread handles. 
	CloseHandle( pi.hProcess );
	CloseHandle( pi.hThread );


// Step 3:  Read in flux value from data file, set it to
//   ms_flux, and return
	char ch;
	char temp_flux[8];
	int i;
	bool SI_units = false;				// si or english units?
	for (i = 0; i< 8; i++)
		temp_flux[i] = '0';

	ifstream fin;
	fin.open("MS-FLUX.OUT", ios::in | ios::nocreate);	
	if (!fin)			// if it couldn't be opened, the user didn't
	{					//   save data in the default name
		flag = 11;		
		ms_flux = 0.0;
		return flag;
	}

	else				// file exists, so read it!
	{
		for (i = 0; i < 765; i++)	// scan the file to position 764
			fin.get(ch);
		if (ch == 'L')				// do units say L/m^2*h or gfd?
			SI_units = true;
		for (i=765; i < 773; i++)	// keep scanning
			fin.get(ch);
		fin >> temp_flux;				// in L/h*m^2

		ms_flux = atof(temp_flux);		// convert string to a number

		if (!SI_units)					// if file has english units
			ms_flux = ms_flux * 1.698;		// convert gfd to L/h*m^2
		ms_flux = ms_flux / 3600000; 		// convert L/h*m^2 to m/s

		// close and delete the default file
		fin.close();
		system("del MS-FLUX.OUT");
	}

	return flag;		// return error flags, if any
}



// Sets up the SEFlux model to determine flux if recirculation
//   is taken into account
int CMechModel::SEFlux(double &se_flux) const
{
	// iteration variables
	int flag, mem_iter, max_mem_iter = 100;

	// recirculation loop variables
	double Cinf, Cin, Cout, Cr, vlos,
		Qinf, Qr, Qeff, Qeff_new, Qperm;

	// Note: Inf, Eff, and Perm are for total membrane mass balance,
	//   In, Out, and R are for mass balances inside the recirculation loop

	// first guesses for variables
	Cinf = mem.Cav * 1000000;	// #/ml to #/m^3
	Qinf = mem.veloc * (mem.CRad * mem.Area)/(2 * mem.CLen);	// m/s to m^3/s
	Qeff = Qinf / 2;
	Cr = Cout = Cinf * Qinf/Qeff;		// assumes perfect rejection
	Qr = Qinf * mem.Circ;
	mem_iter = 0;
	vlos = (Qinf + Qr) * (2 * mem.CLen) / (mem.CRad * mem.Area);

	// loop until Qeff makes sense
	do		
	{
		// set a new concentration
		Cin = (Cinf * Qinf + Cr * Qr) / (Qinf + Qr);

		// determine flux over entire membrane
		flag = do_SEFlux(se_flux, Cin, vlos);

		// determine volumetric flows from influent flow and flux
		Qperm = se_flux * mem.Area;

		// determine Cr, Qeff_new, and Qeff
		Qeff_new = Qinf - Qperm;
		if (Qeff_new < 0)			// non steady-state, a problem
			Qeff = Qeff / 2;
		else
			Qeff += (Qeff_new - Qeff) / 2;

		// determine Cr with a mass balance, assuming perfect rejection
		//   Cr = Cout = Ceff
		Cr = Qinf * Cinf / (Qinf - Qperm);

//		determine permeation concentration? imperfect rejection?

		mem_iter++;

	} while (mem_iter < max_mem_iter && 
			((Qeff_new-Qeff)*(Qeff_new-Qeff)/(Qeff*Qeff)) > 0.000001);

	if (mem_iter >= max_mem_iter)
		flag = 23;		// recirculation rate is too high

	return flag;
}




// Determine flux with Song and Elimelech model:
//   J. Chem. Soc. Faraday Trans., 1995, 91(19), 3389-3398
// Steady-state, average flux in the membrane module
int CMechModel::do_SEFlux(double &se_flux, double conc, 
						  double vlos) const
{

/*	Values for errors/warnings passed by 'flag':
	1  = no warnings, everything went correctly
	20 = solution did not converge
	21 = flux is membrane controlled
	22 = 20 & 21						*/

	int flag = 1;	


/*	Local variables for flux prediction:	
	Kb = Boltzman's constant	 pi = pi, of course!
	D = diffusion coefficient    L  = lambda = MemRes * visc
	Nf = Filtration number       Nc = cake thickness factor
	Amx = As(THETA max)          Asmx = As(THETA* max)
	As = As(THETA*) 
	Pp = drop in pressure over conc. polarization region
	Pm = drop in pressure over membrane
	Ps = pressure* for cake filtration flux
	Beta = Beta value for CP flux determination
	Betas = Beta* value for cake filtration 
	SEflux = calculated flux in m/s
	flux = temporary flux in iterative solution
	error = the abs % difference between Flux and SEFlux
	CWflux = flux of pure water across the membrane			*/


	// Declare and assign initial values for local variables
	int count;
	double D, Nf, Amx, Asmx, As, Pp, Pm, Ps, Beta, Betas, L,
		SEflux, flux, CWflux, error;

	D = Kb * mem.temp / (6 * pi * mem.visc * mem.Rav);
	CWflux = mem.press / (mem.MRes * mem.visc);
	L = mem.MRes * mem.visc;
	flux = 0;
	Pm = 0;
	count = 0;

	// perform an iteration at least once
	do	
	{
		//  The value of Nf will determine the dominant mechanism
		Nf = 4 * pi * (pow (mem.Rav, 3)) * (mem.press - Pm)
							/ (3 * Kb * mem.temp);


// Calculate flux based on the dominant mechanism.  Note: since
//   both flux equations are large, the flux will be determined
//   in several easier-to-follow steps

		if (Nf >= 15)		// Cake Filtration
		{
			double partA, partB, partC; 

			// determine needed coefficients
			Asmx = 23.56;
			Amx  = 123.22;

			Betas = Kb * mem.temp * conc * Amx / 
				(D * D * vlos / mem.CRad);

			Pp = 45 * Kb * mem.temp / (4 * pi * pow(mem.Rav, 3));

			Ps = mem.press - Pp * ( 1 - Amx / Asmx);

			// calculate flux piece-by-piece, eq 4.18 & 3.25
			partA = pow( (pow( (1+ L*L*L /(3 * Betas * Ps * Ps * 
				mem.CLen)), 0.5) + 1 ), 0.33333);

			partB = pow( (pow( (1+ L*L*L /(3 * Betas * Ps * Ps *
				mem.CLen)), 0.5) - 1 ), 0.33333);

			partC = pow( (Ps / (3 * Betas * mem.CLen)), 0.33333) * 
				(partA - partB);

			SEflux = mem.press * (1 - L * partC / mem.press) /
				(mem.CLen * Betas * partC * partC);
			
		}			// end if


		else if (Nf > 0)		// Concentration Polarization
		{
			double partA, partB, partC, partD;

			// Determine As and Beta for Flux prediction
			if (Nf > 2.64)
				As = 2.544 * pow(Nf, 0.804) + 1;
			else
				As = 3.024 * pow(Nf, 0.626) + 1;

			Beta = (Kb * mem.temp * conc * As) /
				(pow (D, 2) * vlos / mem.CRad);


			// Determine the flux piece by piece, eq 3.23 & 3.25
			partA = mem.press / pow((L*L*L + 6 * Beta *
				mem.CLen * pow(mem.press,2)) , 0.33333);

			partB = pow( (pow( (L*L*L / (L*L*L + 6 * Beta * 
				pow(mem.press,2) * mem.CLen)), .5 ) + 2 ) , 0.33333);

			partC = pow( (pow( (L*L*L / (L*L*L + 6 * Beta * 
				pow(mem.press,2) * mem.CLen)), .5 ) - 2 ) , 0.33333);

			partD = partA * (partB - partC);

			SEflux = mem.press * (1 - L * partD / mem.press) /
				(mem.CLen * Beta * partD * partD);
					
		}			// end else if


		else	// Nf is negative - something went wrong with
		{		//   the variable collection.  Shouldn't ever happen.
			flag = -1;
			return flag;
		}


		// reasonable flux check, div 0 error?
		if (SEflux > CWflux) 
		{
			SEflux = 0.99 * CWflux;		// avoid negative pressures in iterations
			flag = 20;					// warning - membrane controlled flux
		}


		// Determine error
		error = fabs ((SEflux - flux) / SEflux);


		// reset Flux and Pm
		Pm =  (SEflux * L + Pm) / 2;
		flux = (SEflux + flux) / 2;
		count ++;


    // continue until change is small, and while the 
	//   solution is converging rapidly enough
	} while ( (error > 0.01) && (count < 100) );
	
	
	if (flag == 20 && count >= 100)	// warning - membrane controlled
		flag = 22;					//   and did not converge

	else if (count >= 100)	// warning - did not converge
		flag = 21;			


	se_flux = SEflux;		// sets the pass-by-ref value seflux
							//   to the flux determined

	// determine permeate and influent flow in m^3/s
	double Qperm, Qin;
	Qperm = se_flux * mem.Area;
	Qin = mem.veloc * (mem.CRad * mem.Area)/(2 * mem.CLen);

	// if permeate is greater than the influent flow
	if (Qperm > Qin)
		flag = 24;

	return flag;				// return the success/error number

}		// end SEFlux function



// Set the mass transfer coefficient of Kr if passed true, Kg if false
void CMechModel::SetMTC(double veloc = -1)
{
	
	double  Re,		// Reynolds Number 
		    v,		// Kinematic Viscosity   m^2/s
		    p,		// Fluid (water) density    kg/m^3		
		    D,		// Diffusion    m^2/s
		    Sc,		// Schmidt Number
			K;		// temporary mtc in m/s

	// use Stokes Einstein equation for diffusivity
	D = Kb * mem.temp / (6 * pi * mem.visc * mem.Rav);
	p = 1000 + 0.0163 * mem.temp - 0.0059 * mem.temp * mem.temp + 0.00002 * mem.temp * mem.temp * mem.temp;
	v = mem.visc / p;
	Sc = v / D;

	// if user inputs a velocity, use that, if not, use the mem object's velocity
	if (veloc < 0)
		Re = mem.veloc * (2 * mem.CRad) / v;
	else
		Re = veloc * (2 * mem.CRad) / v;

	// if the flow is laminar, use Leveque:
	if (Re < 10000)
		K = 1.62 * D / (2 * mem.CRad) * pow ((Re * Sc * 2 * mem.CRad / mem.CLen), 0.33);

	// if mildly turbulent, use Chilton-Colburn
	else if (Re < 100000)
		K = 0.34 * D / (2 * mem.CRad) * pow(Re, 0.75) * pow(Sc, 0.33);

	// if fully turbulent, use Dittus-Boelter:
	else
		K = 0.023 * D / (2 * mem.CRad) * pow(Re, 0.8) * pow(Sc, 0.33);

	// set K to resistance or gel mtc
	Kg = K;
}



// Set the parameters needed for flux prediction with
//   the resistance model
bool CMechModel::SetRes(double op, double ir)
{
	op_res = op;
	ir_res = ir;
	return true;
}



// Determine flux with the resistance model
int CMechModel::ResFlux(double &rs_flux) const
{
	int flag = 1;

	// J = P / [u (Rm + Rf + P*Rop)]

	rs_flux = mem.press / (mem.visc * (mem.MRes + ir_res + (op_res * mem.press / 1000)));

	return flag;
}



// Set the parameters needed for flux prediction with
//   gel polarization
bool CMechModel::SetGel(double gel, bool mgl, double mtc)
{
	bool flag = true;

	if (gel > 0 && mtc >= 0)  // if reasonable values were passed
	{
		Cg = gel;
		if (mgl)
			Cmgl = true;
		else
			Cmgl = false;


		if (mtc == 0)		// if no mass transfer coefft passed, set it
			SetMTC();	
		else				// mtc > 0
			Kg = mtc;
	}

	else
		flag = false;

	return flag;
}



// Determine flux with gel polarization model
// cannot be 'const' because we're calling SetMTC()
int CMechModel::GelFlux(double &gl_flux)
{

/*	Values for errors/warnings passed by 'flag':
	1  = no warnings, everything went correctly
	40 = solution didn't converge    */

	int flag = 1;
	int elems = 100;
	int max_iter = 100;
	int max_mem_iter = 100;
	int iter, mem_iter = 0;

//	later, find out a way to encorperate variable rejection
	double rejection = 0.9;	// 90% rejection of solute by membrane

	// values for iterating over entire membrane
	double Cinf, Qinf, Cr, Qr, Ceff, Qeff, Qeff_new;

	// values for iterating each element
	double Cin, Cout, Cb, Cperm, Qin, Qout, Qperm, vlos;
	

	// discretize membrane into 'elems' units.  In each unit,
	//   Cout(i) = Cin (i+1), Qout(i) = Qin(i+1), 
	//   Qperm(i) is determined from Cb(i) = Cout(i)
	// each element must be iterated until a reasonable Cout
	//   is found.

	// covert Cin to whatever units, to Cgel has
	Cinf = mem.Cav * 4.1888 * pow(mem.Rav, 3) * 1000000;	// #/ml to # (volume fraction)
	if (Cmgl)		// if Cgel is in mg/l
		Cinf = Cinf * mem.Dav * 1000000;		// # to mg/L

	// convert velocity in one fiber to total flow through the 
	//   module (m^3/s)
	Qinf = mem.veloc * (mem.CRad * mem.Area)/(2 * mem.CLen);

	// first guesses for Cr, Ceff, Cout, Qr, Qeff, Qout
	Qeff = Qinf / 2;
	Cr = Ceff = Cout = Cinf * Qinf/Qeff;
	Qr = Qinf * mem.Circ;
	Qout = Qeff + Qr;

	// triple loop: find Cout and Qout with zero rejection, then iterate
	//   while including recirculation rate until Qout = Qout_new
	do		
	{
		mem_iter++;
		gl_flux = 0;

		// Qin and Cin for the first element depend on recirculation
		//   flow and concentration, and initial values
		Qin = Qinf + Qr;
		Cin = (Qinf * Cinf + Qr * Cr) / (Qinf + Qr);

		for (int i = 0; i < elems; i++)		// for each element
		{
			iter = 0;		// reset iter

			// determine the mass transfer coefficient for this element
			//   based on the flow through the element
			vlos = Qin * (2 * mem.CLen)/(mem.CRad * mem.Area);
			SetMTC(vlos);

			do			// determine steady state, completely mixed
			{			// value for Cb, and the flux for the unit
				// Determine Cb and Cperm from last guess
				Cb = Cout;
				Cperm = (1-rejection) * Cb;

				// Find flux for this element in m^3/s
				Qperm = Kg * log((Cg-Cperm)/(Cb-Cperm)) * mem.Area / elems;	
				
				// Find Cout from mass balances
				Cout = (Cin * Qin - Cperm * Qperm) / (Qin - Qperm);
				
				iter++;
			}	while (((Cout-Cb)*(Cout-Cb)/(Cout*Cout)) > 0.000001  
		  		    && iter < max_iter);

			if (iter >= max_iter)
				flag = 40;			// solution didn't converge due to concentration

			// BCs for this element
			Qout = Qin - Qperm;

			// New BC's for next element
			Cin = Cout;			// Cin(i) = Cout(i-1)
			Qin = Qout;			// Qin(i) = Qout(i-1)
			gl_flux = gl_flux + Qperm;		// total flow thru membrane
			iter = 0;
		}

		// determine Cr, Qeff_new, and Qeff
		Qeff_new = Qinf - Qperm;	// or   Qeff_new = Qout - Qr
		if (Qeff_new < 0)			// non steady-state, a problem
			Qeff = Qeff / 2;
		else
			Qeff += (Qeff_new - Qeff) / 2;

		Cr = Cout;

		// see if it equals Qout from last iteration, or if too many
		//   iterations have been done		
	} while (mem_iter < max_mem_iter && 
			((Qeff_new-Qeff)*(Qeff_new-Qeff)/(Qeff*Qeff)) > 0.000001);

	if (mem_iter >= max_mem_iter)
		flag = 42;					// solution didn't converge due to recirculation

	gl_flux = gl_flux / mem.Area;	// flux in m/s
	return flag;
}




#endif		// CMECHMODEL_CPP

