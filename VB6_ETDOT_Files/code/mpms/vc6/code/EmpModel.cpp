// EmpModel.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "EmpModel.h"
#include "ErrBox.h"
#include <math.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEmpModel

CEmpModel::CEmpModel()
{
	ModelType = 0;
	param_a = param_b = param_c = 0;
	set = false;
}


CEmpModel::~CEmpModel()
{
	// destructor: remove all items from CLists
	DataList.RemoveAll();
	xList.RemoveAll();
	yList.RemoveAll();
}


void CEmpModel::set_Emp_model(int type)
{
	ModelType = type;
}

void CEmpModel::get_Emp_model(int &type)
{
	type = ModelType;
}


// insert an item at the tail of DataList
void CEmpModel::insert_tail(CEmpData data)
{
	DataList.AddTail(data);
}


// a mathematical test to see if two numbers are within a 
//   certain percentage range, used in surf_renewal
bool within(double new_num, double old_num, double range)
{
	double err_sq = pow( ((new_num-old_num)/old_num), 2);
	if ( err_sq > (range * range) )
		return false;
	else
		return true;
}


/////////////////////////////////////////////////////////////////////////////
// CEmpModel model making functions

// determines the values for A and Js with 'dead_end', then
//   determines s with 'x_flow'.  Assumes that the two data
//   sets have identical ranges.
int CEmpModel::surf_fit(CEmpData &dead_end, CEmpData &x_flow)
{
	int flag=0;
	int count=0;
	double A, Js, flux, time;
	Point pTemp;

	// check for insufficient data...
	if (((dead_end.Points).GetCount()) < 2 ||
		((x_flow.Points).GetCount()) < 1)
		return 13;		// insufficient data error (added to 300 later)

	// use first data set to find A and Js
	flag += surf_renewal(dead_end, A, Js);

	// now, use second data set to find s

	xList.RemoveAll();		// clean up xList
	// get position of first point in Points list
	POSITION index = (x_flow.Points).GetHeadPosition();
	
	// using flux and time from x_flow, determine 's' at this 
	//   velocity for all points, enter them into xList
	double fOld, fpOld, Jo, sOld, s;

	// peek at first 'Point' in the CList 'Points' in the object 'data', 
	// obtain the constant pressure
	double const_pres = (((x_flow.Points).GetHead()).pres);

	// determine clean water flux - 0.001 is typical MKS water viscosity
	Jo = const_pres / (0.001 * mem.MRes);

	while (index != NULL)
	{
		pTemp = (x_flow.Points).GetNext(index);
		flux = pTemp.flux;
		time = pTemp.time;

		// need to do an iterative technique for finding s
		sOld = 0;		// first guesses
		s = 1/time;

		while ( ( ((s-sOld)/s) > 0.001 || ((s-sOld)/s) < -0.001 ) 
			    && count < 1000 )
		{
			count++;

			// Use Newtons method
			fOld  = (Jo - Js) * (s/(s+A)) * (1 - exp(-time*(s+A))) / (1 - exp(-time*s)) + Js - flux;
			fpOld = (Jo - Js) * ( ( (1-exp(-time*(s+A))) / (1-exp(-time*(s+A)))  *  A/((s+A)*(s+A)) )  +  
				                  ( (s*time)/(s+A) * ( (exp(-time*s) - exp(-time*(s+A))) / (pow( (1-exp(-time*s)), 2)) ) ) );
			sOld  = s;
			s = sOld - fOld / fpOld;
		}

		xList.AddHead(s);
	}

	if (count == 1000)
		flag = 12;		// 1000 iterations, no solution for s,
						//   add to 300 in build_model to make 312

	// find average value for s
	xList_ave(s);

	// place A & Js into 'dead_end', place all into 'x_flow'
	(dead_end.A).set_val(A);
	(dead_end.Js).set_val(Js);
	(x_flow.A).set_val(A);
	(x_flow.Js).set_val(Js);
	(x_flow.s).set_val(s);
	x_flow.viable = true;		// make 'x_flow' a viable model

	return flag;
}



// fit data for the 
int CEmpModel::foul_fit(CEmpData &ad_foul, 
						CEmpData &cp_foul, CEmpData &pp_foul)
{
	int flag = 0;
	double Ra, Rpp, Rcp;
	double Rm = mem.MRes;
	POSITION index;
	Point pTemp;

	// check for insufficient data, pressure ratios...
	//

	// problem: viscosity is not set anywhere!  must assume
	//   a viscosity, do not enable user to alter it?
	double visc = 0.001;

	// first data set:  Ja  = P / (visc * (Rm + Ra))
	// find value for Ra for each data point, place it in xList,
	//   and find the average value for Ra.
	index = (ad_foul.Points).GetHeadPosition();
	xList.RemoveAll();
	while (index != NULL)
	{
		pTemp = (ad_foul.Points).GetNext(index);
		Ra = (pTemp.pres) / ((pTemp.flux)*(visc)) - Rm;
		xList.AddHead(Ra);
	}

	xList_ave(Ra);

	
	// Note:  the user enters cp_foul before pp_foul by the nature
	//   of the experimental setup.
	// second set:		Jpp = P / (visc * (Rm + Ra + Rpp))
	index = (pp_foul.Points).GetHeadPosition();
	xList.RemoveAll();
	while (index != NULL)
	{
		pTemp = (pp_foul.Points).GetNext(index);
		Rpp = (pTemp.pres) / ((pTemp.flux)*(visc)) - Rm - Ra;
		xList.AddHead(Rpp);
	}

	xList_ave(Rpp);


	// third set:		Jcp = P / (visc * (Rm + Ra + Rpp + Rcp))
	index = (cp_foul.Points).GetHeadPosition();
	xList.RemoveAll();
	while (index != NULL)
	{
		pTemp = (cp_foul.Points).GetNext(index);
		Rcp = (pTemp.pres) / ((pTemp.flux)*(visc)) - Rm - Ra - Rpp;
		xList.AddHead(Rcp);
	}

	xList_ave(Rcp);

	// now all values for resistances are set, update their values
	(ad_foul.Ra).set_val(Ra);
	(pp_foul.Ra).set_val(Ra);
	(pp_foul.Rpp).set_val(Rpp);
	(cp_foul.Ra).set_val(Ra);
	(cp_foul.Rpp).set_val(Rpp);
	(cp_foul.Rcp).set_val(Rcp);
	cp_foul.viable = true;

	// future errors - negative values for resistances?

	return flag;
}


int CEmpModel::res_fit(CEmpData &data)
{
	int flag = 0;

	// if variable parameter is not pressure, 
	//   therefore not primary, then return
	if (!(data.primary))
		return flag;

	// else, fit pressure vs. flux data

	// variable declarations, setup for iteration
	double R, T, Rf, Rop, visc;
	POSITION index = (data.Points).GetHeadPosition();
	Point pTemp;
	yList.RemoveAll();
	xList.RemoveAll();

	// place flux in yList, pressure in xList, for each point
	while (index != NULL)
	{
		pTemp = (data.Points).GetNext(index);
		yList.AddHead((pTemp.flux));
		xList.AddHead((pTemp.pres));
	}

	// call curve-fitting algorithm to find R (visc*(Rm + Rfoul)), and T (visc*theta)
	resistance(R, T);

	// determine Rf and theta from R, T
//		cannot yet do viscosity independant flux, assume viscosity
//		of water: 0.001
	visc = 0.001;
	Rf  = (R / visc) - mem.MRes;	// units 1/m
	Rop = T / (visc);				// units 1/m*Pa

	(data.Rf).set_val(Rf);
	(data.Rop).set_val(Rop);
	data.viable = true;

	return flag;
}


int CEmpModel::gel_fit(CEmpData &data)
{
	int flag = 0;

	// flux = k*ln(Cg/Cb) = kln(Cg) - kln(Cb)

	// if variable parameter is not concentration, 
	//   therefore not primary, simply return
	if (!(data.primary))
		return flag;

	// else, fit conc vs. flux data

	// variable declarations, setup for iteration
	double k, Cg, log_conc, slope, intercept;
	POSITION index = (data.Points).GetHeadPosition();
	Point pTemp;
	yList.RemoveAll();
	xList.RemoveAll();

	// place flux in yList, ln(conc) in xList, for each point
	while (index != NULL)
	{
		pTemp = (data.Points).GetNext(index);
		yList.AddHead((pTemp.flux));
		log_conc = log((pTemp.conc));
		xList.AddHead(log_conc);
	}

	// call curve-fitting algorithm to find 
	line(slope, intercept);

	// extract Cg and k from 'slope' and 'intercept'
	k  = -slope;
	Cg = exp(intercept / k);

	// if negative values are reported, alter flag value
	if (k < 0 || Cg < 0)
		flag = 1;

	(data.k).set_val(k);
	(data.Cg).set_val(Cg);
	data.viable = true;

	return flag;
}



/////////////////////////////////////////////////////////////////////////////
//	CEmpModel curve fitting algorithms

// finds the average value for the series of values in xList
int CEmpModel::xList_ave(double &ave)
{
	int flag = 0;
	int count = xList.GetCount();
	double sum = 0;
	POSITION index = xList.GetHeadPosition();
	
	// iterate, finding the num and number of entries in xList
	while (index != NULL)
		sum += (xList.GetNext(index));

	ave = sum / count;

	return flag;
}


int CEmpModel::line(double &m, double &b)
{
	int flag=0;

	// sums required for curve fitting analysis:
	double	x,		// sum of (x)
			y,		// sum of (y)
			xSQ,	// sum of (x^2)
			xy,		// sum of (x*y)
			num,	// number of entries
			ynum,	// the current y-number
			xnum;	// the current x-number
	y = x = xSQ = xy = num = 0;

	// scan xList and yList to calculate sums
	POSITION yIndex, xIndex;
	yIndex = yList.GetHeadPosition();
	xIndex = xList.GetHeadPosition();

	while (yIndex != NULL && xIndex != NULL)
	{
		ynum = (yList.GetNext(yIndex));
		xnum = (xList.GetNext(xIndex));
		y += ynum;
		x += xnum;
		xSQ += (xnum * xnum);
		xy += (xnum * ynum);
		num++;
	}

	// determine slope (m) and intercept (b) from sums:

	m = (xy - (x * y) / num) / (xSQ - (x * x) / num);
	b = (y - (m * x)) / num;

	return flag;
}


int CEmpModel::resistance(double &R, double &T)
{
	int flag=0;

	// sums required for curve fitting analysis:
	double	y,		// sum of (y)
			ySQ,	// sum of (y^2)
			yOVx,	// sum of (y/x)
			ysqOVx,	// sum of (y^2/x)
			yOVxSQ,	// sum of ((y/x)^2)
			ynum,	// the current y-number
			xnum;	// the current x-number
	y = ySQ = yOVx = ysqOVx = yOVxSQ = 0;

	// scan xList and yList to calculate sums
	POSITION yIndex, xIndex;
	yIndex = yList.GetHeadPosition();
	xIndex = xList.GetHeadPosition();

	while (yIndex != NULL && xIndex != NULL)
	{
		ynum = (yList.GetNext(yIndex));
		xnum = (xList.GetNext(xIndex));
		y += ynum;
		ySQ += (ynum * ynum);
		yOVx += (ynum / xnum);
		ysqOVx += ((ynum*ynum) / xnum);
		yOVxSQ += ((ynum*ynum) / (xnum*xnum));
	}

	// determine R and T from sums
	R = (yOVx - ((y * ysqOVx)/ySQ)) / (yOVxSQ - ((ysqOVx * ysqOVx)/y));

	T = (y - R * ysqOVx) / ySQ;

	return flag;
}



int CEmpModel::surf_renewal(const CEmpData data, double &A, double &Js)
{
	int flag=0;

	// fit the data, flux versus time, to determine
	//   the coefficients A, and Js

	// variable declarations
	POSITION index, head, tail;
	tail = (data.Points).GetTailPosition();
	head = (data.Points).GetHeadPosition();
	double Jold, Jo, Jtemp, R, Rold, flux, time;
	int count = 0;
	Point pTemp;

	// first guess: Js is the last value for flux, 
	//              Jold is the one before Js
	//		since the data is entered as a stack, the last value
	//		in is the first value out: therefore, the head is the
	//		last value the user entered, and is most likely the
	//		data point with the lowest flux, and greatest time

	Js   = ((data.Points).GetNext(head)).flux;
	Jold = ((data.Points).GetNext(head)).flux;

	// peek at first 'Point' in the CList 'Points' in the object 'data', 
	// obtain the pressure, which is constant for all points
	double const_pres = (((data.Points).GetHead()).pres);

	// determine clean water flux - 0.001 is typical MKS water viscosity
	Jo = const_pres / (0.001 * mem.MRes);

	do 
	{
		// set up 'x' to solve for A using an average, work backwards
		index = tail;
		xList.RemoveAll();

		while (index != NULL)
		{
			// get flux and time from next Point object from data.Points
			pTemp = ((data.Points).GetPrev(index));
			flux = pTemp.flux;
			time = pTemp.time;

			// if we can determine A, do so:
			if (flux != Js)
			{
				// determine this value for A
				A = (-1/time) * log((flux - Js)/(Jo - Js));

				// insert it at head of xList
				xList.AddHead(A);
			}

		}

		// find the average of xList, set it to A
		xList_ave(A);

		// find the residual - the sum of the squares of the differences
		index = tail;
		R = 0;			// reset R
		while (index != NULL)
		{
			pTemp = ((data.Points).GetPrev(index));
			flux = pTemp.flux;
			time = pTemp.time;

			// sum of the squares of the differences
			R += pow( (flux - ((Jo-Js)*exp(-A*time) + Js)), 2);	
		}	

		// approximation for Rold, but only for the first iteration
		if (count == 0)
			Rold = 2 * R;

		// find new value for Js with secant method, set Jold to Js,
		//   use Jtemp to store old Js (which is NOT Jold)
		Jtemp = Js;
		Js = Jold - R * (Js-Jold)/(R-Rold);
		Jold = Jtemp;

		Rold = R;
		count++;

		// if Js ~ Jold, we're done, if not, keep looping
	}	while ( ( ((Js-Jold)/Js) > 0.01 || 
				  ((Js-Jold)/Js) < (-0.01) ) &&
				count < 100);

	if (count == 100)
		flag = 11;		// 100 iterations, no solution for A or J
						//   add to 300

	return flag;
}



/////////////////////////////////////////////////////////////////////////////
// CEmpModel merging functions... may not get around to them until
//   the fall, if at all.  These functions will scan all viable
//   models and see if the ModelVar objects in the models can be
//   improved upon with other viable or non-viable models
int CEmpModel::surf_merge()
{
	int flag=0;

	// Js should be constant,
	// A varies as P*kC^x
	// s varies as kV^x

	return flag;
}


int CEmpModel::foul_merge()
{
	int flag=0;

	// Ra depends on concentration, and temp
	// Rpp depends on pressure,
	// Rcp depends on velocity

	return flag;
}


int CEmpModel::res_merge()
{
	int flag=0;
	
	// Rf depends on temp and conc
	// Rop depends on press and velocity

	return flag;
}


int CEmpModel::gel_merge()
{
	int flag=0;

	// Cg should never change
	// k changes with everything

	return flag;
}

	
/////////////////////////////////////////////////////////////////////////////
// CEmpModel data handling functions
void CEmpModel::add_data(CEnterDataDlg data)
{
	CEmpData NewData;

	NewData.set_type(ModelType);
}


// first, build each CEmpData object's model params, then merge
//   viable with non-viable data sets
int CEmpModel::build_model()
{
	int flag = 300;			// initially set

	if (DataList.IsEmpty())
		return flag;		// data list is empty!

	CEmpData data, dataB, dataC;

	// builds the model in 2 steps: first, use 'xxx_fit' to fit some
	//   of the data to model coefficients.  Then use 'xxx_merge'
	//   to complete, or expand the model coefficients.  Each acts
	//   differently, so check comments for each.

	int i, data_sets = DataList.GetCount();

	switch (ModelType)		// 1=surf ren, 2=res foul, 3=res, 4=gel
	{
	case 1:		for (i = 0; i<(data_sets/2); i++)
				{	// modify 2 data sets in the queue
					data  = DataList.RemoveHead();	// dead end
					dataB = DataList.RemoveHead();	// cross flow
					
					// determine and set model coefficients
					flag  += surf_fit(data, dataB);

					// place them in the back of the queue
					DataList.AddTail(data);
					DataList.AddTail(dataB);
				} 
				flag += surf_merge();
				break;

	case 2:		for (i = 0; i<(data_sets/3); i++)
				{	// modify 3 data sets in the queue
					data  = DataList.RemoveHead();	// static adsorption
					dataB = DataList.RemoveHead();	// full operation
					dataC = DataList.RemoveHead();	// pore plugging
					
					flag  += foul_fit(data, dataB, dataC);

					DataList.AddTail(data);
					DataList.AddTail(dataB);
					DataList.AddTail(dataC);
				} 
				flag += foul_merge();
				break;

	case 3:		for (i = 0; i<data_sets; i++)
				{	// modify 1 data set in the queue
					data = DataList.RemoveHead();
					flag += res_fit(data);
					DataList.AddTail(data);
				}
				flag += res_merge();
				break;

	case 4:		for (i = 0; i<data_sets; i++)
				{	// modify 1 data set in the queue
					data = (DataList.RemoveHead());
					flag += gel_fit(data);
					DataList.AddTail(data);
				}
				flag += gel_merge();
				break;
	}

	// each subroutine returns '0' if they ran properly, then if 
	//   flag = 300, no problems occured
	if (flag == 300)	
		flag = 1;		// no problems occurred at all


	return flag;		// if not 1, then problems occurred
}



// Using the passed parameters, find the data set who range is 
//   closest, and calculate flux with that model
int CEmpModel::get_flux(double &flux, int &compat, double C, double P, 
						double V, double T, double visc, double TIME=0)
{
	int flag = 1;

	// units of input variables:
	// C = conc(% or mg/L)
	// P = pres(Pa)
	// V = vlos(m/s)
	// T = temp(C)
	// visc(kg/m*s)
	// TIME(s)

	// set the parameter values in mem
	mem.SetParam(P, (T+273), V, visc);

	// cannot set conc or time, our conc has varying units, 
	//   and there isn't a variable for permeation time, etc.

	// starting with the head of the CList 'DataList', scan the
	//   CEmpData objects for maximum range compatibility

	POSITION max_index, cur_index, next_index;
	int max_compat;		// range compatibility, higher is better
	bool set_max = true;

	// intially set next_index, and max_compat
	next_index = DataList.GetHeadPosition();
	max_compat = 0;

	// loop until all items have been seen, note that GetNext
	//   increments 'next_index', thus the need for 'cur_index'

	while (next_index != NULL)
	{
		cur_index = next_index;

		// only test viable models
		if ((DataList.GetNext(next_index)).viable)
		{
			// see if current CEmpData object is in range
			compat = (DataList.GetAt(cur_index)).in_range(C, P, V, T, TIME);
			if (max_compat < compat || set_max)
			{
				max_compat = compat;
				max_index = cur_index;
				set_max = false;		
			}	// end if
		}		// end if

	}


	// set return value for compat to max_compat, add 350 to make
	//   sure compat corresponds to the appropriate error msg
	compat = max_compat + 350;


	// max_index is the index of the CEmpData object which is most
	//   compatible with the ranges.  Obtain the flux coefficients
	//   from that model, and calculate flux:
	switch(ModelType)
	{
	case 1:		// Surface Renewal
		if (!set)		// if user has not preset the parameters
		{	param_a = ((DataList.GetAt(max_index)).Js).get_val();
			param_b = ((DataList.GetAt(max_index)).A).get_val();
			param_c = ((DataList.GetAt(max_index)).s).get_val();
		}
		flag = get_surf_flux(flux, param_a, param_b, param_c, TIME);
		break;

	case 2:		// Resistance Fouling
		if (!set)
		{	param_a = ((DataList.GetAt(max_index)).Ra).get_val();
			param_b = ((DataList.GetAt(max_index)).Rpp).get_val();
			param_c = ((DataList.GetAt(max_index)).Rcp).get_val();
		}
		flux = P / (visc * ((mem.MRes) + param_a + param_b + param_c));
		break;

	case 3:		// Resistance
		if (!set)
		{	param_a = ((DataList.GetAt(max_index)).Rf).get_val();
			param_b = ((DataList.GetAt(max_index)).Rop).get_val();
			param_c = 0;
		}
		flux = P / (visc * ((mem.MRes) + param_a + P * param_b));
		break;

	case 4:		// Gel Polarization
		if (!set)
		{	param_a = ((DataList.GetAt(max_index)).k).get_val();
			param_b = ((DataList.GetAt(max_index)).Cg).get_val();
			param_c = 0;
		}
		flag = get_gel_flux(flux, param_a, param_b, C);
		break;
	}

	// if the data is not extrapolated or interpolated, do the
	//   estimation here 

//	needed for debug test for some reason...
	double test_a, test_b, test_c;
	test_a = param_a;
	test_b = param_b;
	test_c = param_c;

	// another flag
	if (flux < 0.0)
		flag += 1000;	// flux is negative

	set = false;
	return flag;
}


int CEmpModel::set_param(double p_a, double p_b, double p_c)
{
	param_a = p_a;
	param_b = p_b;
	param_c = p_c;
	set = true;
	return 1;
}


int CEmpModel::get_gel_flux(double &flux, double k, double Cg, double Cb)
{
	int flag;

	// use this simplistic gel model:

	flux = k * log (Cg / Cb);

	if (Cb > Cg)
	{
		flag = 340;		// invalid bulk concentration
		flux = 0;
	}

	if (Cg <= 0 || k <= 0)
		flag = 301;		// invalid Cg or k value

	return 1;
}


int CEmpModel::get_surf_flux(double &flux, double Js, double A, double s, double tp)
{
	double Jo = (mem.press) / (mem.visc * mem.MRes);	// clean water flux
	double Jj, Ji; 				// fluxes
	double tj, ti; 				// times
	int num_steps = 100;		// the number of steps

	// check for valid parameters first
	if (Js <= 0 || A <= 0 || s <= 0)
		return 301;

	// initial conditions
	Ji = Jo;
	ti = 0;
	flux = 0;

	// determine time averaged flux numerically
	for (int i=0; i<num_steps; i++)
	{
		// calculate tj, the NEXT iteration's time value
		tj = tp * (i+1)/num_steps;

		// calculate Jj, the NEXT iteration's flux value
		Jj = Js + (Jo - Js) * (s/(s+A)) * ( (1-exp(-tj*(s+A))) / (1-exp(-tj*s)) );

		// increment flux with trapezoid rule
		flux += (Jj + Ji)/2;

		// set up BC's for next iteration
		Ji = Jj;
		ti = tj;
	}

	// determine average flux - this equation works only because
	//   the change in t is constant from zero to tp
	flux = flux / num_steps;

	if (flux < 0)
		return 302;

	return 1;

}






/////////////////////////////////////////////////////////////////////////////
// Printing out, or printing the data to a file

// overload the "<<" operator, so you can easily print it out
fstream& operator<< (fstream& fout, const CEmpModel& output_data)
{
	int data_count = 1;
	POSITION point_index, data_index;
	data_index = (output_data.DataList).GetHeadPosition();
	CEmpData data;
	Point data_point;

	while (data_index != NULL)
	{
		data = (output_data.DataList).GetNext(data_index);

		// for each object, number it
		fout << "Data set #" << data_count << endl;
		data_count++;

		// report its viability; if viable, then its values
		if (data.viable)
		{
			fout << "Valid model coefficients:" << endl;
			switch(data.ModelType)
			{
			case 1:	fout << "Min Flux:\t\t\t" << ((data.Js).get_val()) << " m/s" << endl;
					fout << "Flux Decline:\t\t\t" << ((data.A).get_val()) << " 1/s" << endl;
					fout << "Surface Renewal:\t\t"<< ((data.s).get_val()) << " 1/s" << endl;
			break;

			case 2:	fout << "Ra:\t\t\t" << ((data.Ra).get_val()) << " 1/m" << endl;
					fout << "Rpp:\t\t\t" << ((data.Rpp).get_val()) << " 1/m" << endl;
					fout << "Rcp:\t\t\t" << ((data.Rcp).get_val()) << " 1/m" << endl;
			break;

			case 3:	fout << "Rf:\t\t\t" << ((data.Rf).get_val()) << " 1/m" << endl;
					fout << "Rop:\t\t\t" << ((data.Rop).get_val()) << " 1/m" << endl;
			break;

			case 4:	fout << "k:\t\t\t" << ((data.k).get_val()) << " m/s" << endl;
					fout << "Cg:\t\t\t" << ((data.Cg).get_val()) << endl;
			break;
			}
		}			// end if




		// next, the parameters
		fout << endl << "Parameters for the entered data set" << endl;
		fout << "Concentration:\t" << (data.conc[1]) << endl;
		fout << "Pressure:     \t" << (data.pres[1]) << endl;
		fout << "Velocity:     \t" << (data.vlos[1]) << endl;
		fout << "Temperature:  \t" << (data.temp[1]) << endl;
		fout << "Time:         \t" << (data.time[1]) << endl;



		// last, the flux vs. parameter value
		fout << endl << "Flux (m/s):\t\t\t";
		switch (data.param)
		{
		case 1:	fout << "Concentration:" << endl;
		break;

		case 2:	fout << "Pressure (kPa):" << endl;
		break;

		case 3: fout << "Velocity (m/s):" << endl;
		break;

		case 4: fout << "Temperature (C):" << endl;
		break;

		case 5:	fout << "Time (s):" << endl;
		break;
		}

		point_index = (data.Points).GetHeadPosition();
		
		while (point_index != NULL)
		{
			data_point = (data.Points).GetNext(point_index);

			fout << (data_point.flux) << "\t\t\t";
			switch (data.param)
			{
			case 1:	fout << (data_point.conc) << endl;
			break;
			case 2:	fout << (data_point.pres) << endl;
			break;
			case 3: fout << (data_point.vlos) << endl;
			break;
			case 4: fout << (data_point.temp) << endl;
			break;
			case 5: fout << (data_point.time) << endl;
			break;
			}
		}

		fout << endl << endl;

	}	// end while
		
	return fout;
}




BEGIN_MESSAGE_MAP(CEmpModel, CWnd)
	//{{AFX_MSG_MAP(CEmpModel)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CEmpModel message handlers
