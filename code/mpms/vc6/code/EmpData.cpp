// EmpData.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "EmpData.h"
#include <afxtempl.h>
#include <afx.h>

#ifndef EMPDATA_CPP
#define EMPDATA_CPP

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEmpData

// Constructor
CEmpData::CEmpData()
{
	conc[0] = conc[1] = -1;
	pres[0] = pres[1] = -1;
	vlos[0] = vlos[1] = -1;
	temp[0] = temp[1] = -1;
	time[0] = time[1] = -1;
	tolerance = 0.02;
	viable = false;
}



// copy constructor
CEmpData::CEmpData(const CEmpData &rhs)
{
	int i,j;

	// copy model parameters
	ModelType = rhs.ModelType;	
	viable = rhs.viable;
	tolerance = rhs.tolerance;
	primary = rhs.primary;
	param = rhs.param;

	// copy model coefficients
	A = rhs.A;
	Js = rhs.Js;
	s = rhs.s;
	Ra = rhs.Ra;
	Rpp = rhs.Rpp;
	Rcp = rhs.Rcp;
	Rf = rhs.Rf;
	Rop = rhs.Rop;
	k = rhs.k;
	Cg = rhs.Cg;

	// copy ranges
	for (i=0; i<2; i++)
	{
		conc[i] = rhs.conc[i];
		pres[i] = rhs.pres[i];
		vlos[i] = rhs.vlos[i];
		temp[i] = rhs.temp[i];
		time[i] = rhs.time[i];
	}

	// copy data list, one data item at a time
	j = rhs.Points.GetCount();
	if (j == 0)		// if CList is empty, do nothing
		return;

	Point temp_pt;

	// get the memory location of the head of the list
	POSITION index = rhs.Points.GetHeadPosition();

	for (i=0; i<j; i++)
	{
		// make a copy of the item at 'index', which is automatically incremented
		temp_pt = rhs.Points.GetNext(index);

		// add the new item to left-hand-side's list of data points
		Points.AddTail(temp_pt);
	}
}



// equality operator: exact same as copy constructor, 
//   except for return values
const CEmpData& CEmpData::operator=(const CEmpData &rhs)
{
	int i,num_points;

	// copy model parameters
	ModelType = rhs.ModelType;	
	viable = rhs.viable;
	tolerance = rhs.tolerance;
	primary = rhs.primary;
	param = rhs.param;

	// copy model coefficients
	A = rhs.A;
	Js = rhs.Js;
	s = rhs.s;
	Ra = rhs.Ra;
	Rpp = rhs.Rpp;
	Rcp = rhs.Rcp;
	Rf = rhs.Rf;
	Rop = rhs.Rop;
	k = rhs.k;
	Cg = rhs.Cg;

	// copy ranges
	for (i=0; i<2; i++)
	{
		conc[i] = rhs.conc[i];
		pres[i] = rhs.pres[i];
		vlos[i] = rhs.vlos[i];
		temp[i] = rhs.temp[i];
		time[i] = rhs.time[i];
	}

	// copy data list, one data item at a time
	num_points = rhs.Points.GetCount();
	if (num_points == 0)		// if CList is empty, do nothing
		return *this;

	Point temp_pt;

	// get the memory location of the head of the list
	POSITION index = rhs.Points.GetHeadPosition();

	for (i=0; i<num_points; i++)
	{
		// make a copy of the item at 'index', which is automatically incremented
		temp_pt = rhs.Points.GetNext(index);

		// add the new item to left-hand-side's list of data points
		Points.AddTail(temp_pt);
	}

	// return the newly created CEmpData object
	return *this;
}



// set the data type
void CEmpData::set_type(int type)
{
	ModelType = type;
}


// needed by set_data for string parsing
bool parse_data(CString &string, int &front, double &num)
{
	int back = string.Find('\n');	// first instance of a newline
	if (back == -1)					// no newline found
		return false;
	CString item = string.Mid(front, back);	// extract a chunk of the string
	front = back + 1;						// front of next item 
	string.SetAt(back, ' ');					// replace newline with a space, so it wont be 'found' twice 
	num = atof(item);		// convert item to a number 
	return true;
}				


// assemble the data points
bool CEmpData::set_data(const CEnterDataDlg& data)
{
	// if strings are empty (no flux or param entries), return false
	if ( ((data.m_flux).IsEmpty())  ||  ((data.m_param).IsEmpty()))
		return false;

	// set the varied parameter, and primary value
	// 1=conc, 2=pres, 3=vlos, 4=temp, 5=time
	param = data.param;
	primary = data.primary;

	// set the Model Type, 1=surf, 2=foul, 3=res, 4=gel
	ModelType = data.ModelType;

	// viable is set to TRUE only after the model is built
	viable = false;


	// Point object variables
	double t_flux, t_conc, t_pres, t_vlos, t_temp, t_time;
	t_flux = t_time = 0;
	t_conc = data.m_conc;
	t_pres = data.m_pres * 1000;	// convert kPa to Pa
	t_vlos = data.m_vlos;
	t_temp = data.m_temp;

	// construct pTemp (varied param is overwritten later)
	Point pTemp(t_flux, t_conc, t_pres, t_vlos, t_temp, t_time);

	double param_lo, param_hi, param_new;
	int flux_front, param_front, low_flag=0;
	flux_front = param_front = 0;
	CString flux_data = data.m_flux;
	CString param_data = data.m_param;

	// parse the strings for the numbers between newlines '\n',
	//   copy the string, and turn it into a number
	while (parse_data(flux_data, flux_front, t_flux))
	{
		low_flag = param_front;
		if (!(parse_data(param_data, param_front, param_new)))
			return false;		// error!  flux num doesn't correspond to a param num

		// find the smallest and largest param value for setting ranges
		if (low_flag == 0)		// if not yet set
			param_lo = param_hi = param_new;
		else			// determine if param_new is lower than
		{				// param_lo, or higher than param_hi
			if (param_lo > param_new)
				param_lo = param_new;
			else if (param_hi < param_new)
				param_hi = param_new;
		}

		// reset values for flux and the param in pTemp,
		//   the other values should remain the same
		switch(param)
		{
		case 1 :	pTemp.conc = param_new;
			break;
		case 2 :	pTemp.pres = param_new * 1000;	// kPa to Pa
			break;
		case 3 :	pTemp.vlos = param_new;
			break;
		case 4 :	pTemp.temp = param_new;
			break;
		case 5 :	pTemp.time = param_new;
			break;
		}
		
		pTemp.flux = t_flux;

		// Add pTemp to the list of Points
		Points.AddHead(pTemp);
	}


	// set the ranges
	conc[0] = pres[0] = vlos[0] = temp[0] = time[0] = -1;
	conc[1] = t_conc;
	pres[1] = t_pres;	// Pa
	vlos[1] = t_vlos;
	temp[1] = t_temp;
	time[1] = t_time;

	// reset the low and high range of the varied param
	switch(param)
	{
	case 1 :	conc[0] = param_lo;
				conc[1] = param_hi;
		break;
	case 2 :	pres[0] = param_lo * 1000;	// kPa to Pa
				pres[1] = param_hi * 1000;
		break;
	case 3 :	vlos[0] = param_lo;
				vlos[1] = param_hi;
		break;
	case 4 :	temp[0] = param_lo;
				temp[1] = param_hi;
		break;
	case 5 :	time[0] = param_lo;
				time[1] = param_hi;
		break;
	}

	return true;
}


// destructor
CEmpData::~CEmpData()
{
	Points.RemoveAll();
}


// not needed after I finish with 'set_data'
bool CEmpData::insert(int var, double F, double C, double P,
					  double V, double T, double TIME)
{
	// if CList is empty, set ranges, make a new head, return true
	if (Points.IsEmpty())
	{
		// set the second item in array for "range" of this point
		conc[1] = C;
		pres[1] = P;
		vlos[1] = V;
		temp[1] = T;
		time[1] = TIME;
		Point pTemp(F, C, P, V, T, TIME);
		Points.AddHead(pTemp);
		return true;
	}

	// else, make a new Point, and insert it at the tail of CList
	Point pTemp(F, C, P, V, T, TIME);
	Points.InsertAfter((Points.GetTailPosition()), pTemp);

	// update the range values later
	
	return true;
}



// tests for range compatibility, can tell how many parameters
//   are in range, returns a validity flag
int CEmpData::in_range(double C, double P, double V, 
					   double T, double TIME)
{
	int num = 0;		// # of times any param is in range
	int out_param;		// parameter that is out of range
	int return_val;		/*	0 = not enough params in range,	(3/5 or 2/4)
							1 = enough to estimate			(4/5 or 3/4, main param is in range)
							2 = enough to extrapolate		(4/5 or 3/4, main param is out of range)
							3 = enough to interpolate		(5/5 or 4/4)								*/

	double Cdiff, Pdiff, Vdiff, Tdiff, TIMEdiff;


	// Concentration
	if (conc[0] <= -1)		// if conc is simply a value,
	{						//   see if it is within tolerance
		Cdiff= (C-conc[1])/(conc[1]);
		if ( (Cdiff*Cdiff) < (tolerance*tolerance) ) 
			num++;
		else	out_param = 1;
	}
	else					// else, check conc range
	{	if (C >= conc[0] && C <= conc[1])
			num++;
		else	out_param = 1;
	}


	// Pressure
	if (pres[0] <= -1)
	{
		Pdiff = (P-pres[1])/(pres[1]);
		if ( (Pdiff*Pdiff) < (tolerance*tolerance) )
			num++;
		else	out_param = 2;
	}
	else
	{
		if (P >= pres[0] && P <= pres[1])
			num++;
		else	out_param = 2;
	}


	// Velocity
	if (vlos[0] <= -1)
	{
		Vdiff = (V-vlos[1])/(vlos[1]);
		if ( (Vdiff*Vdiff) < (tolerance*tolerance) )
			num++;
		else	out_param = 3;
	}
	else
	{
		if (V >= vlos[0] && V <= vlos[1])
			num++;
		else	out_param = 3;
	}


	// Temperature
	if (temp[0] <= -1)
	{
		Tdiff = (T-temp[1])/(temp[1]);
		if ( (Tdiff*Tdiff) < (tolerance*tolerance) )
			num++;
		else	out_param = 4;
	}
	else
	{
		if (T >= temp[0] && T <= temp[1])
			num++;
		else	out_param = 4;
	}


	// if the model type is not surface renewal, or otherwise time
	//   dependant, we can stop here, check the values, set a 
	//   return value, and return it
	
	if (ModelType != 1)
	{
		if (num < 3)
			return_val = 0;		// can't do anything
		else if (num == 3 && out_param != param)
			return_val = 1;		// can estimate
		else if (num == 3)
			return_val = 2;		// can extrapolate
		else if (num == 4)
			return_val = 3;		// can interpolate

		return return_val;
	}


	// Time
	if (time[0] <= -1)
	{
		TIMEdiff = (TIME-time[1])/(time[1]);
		if ( (TIMEdiff*TIMEdiff) < (tolerance*tolerance) )
			num++;
		else	out_param = 5;
	}
	else
	{
		if (TIME >= time[0] && TIME <= time[1])
			num++;
		else	out_param = 5;
	}



	if (num < 4)
		return_val = 0;		// can't do anything
	else if (num == 4 && out_param != param)
		return_val = 1;		// can estimate
	else if (num == 4)
		return_val = 2;		// can extrapolate
	else if (num == 5)
		return_val = 3;		// can interpolate

	return return_val;

}									





BEGIN_MESSAGE_MAP(CEmpData, CWnd)
	//{{AFX_MSG_MAP(CEmpData)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CEmpData message handlers








#endif		// EMPDATA_CPP