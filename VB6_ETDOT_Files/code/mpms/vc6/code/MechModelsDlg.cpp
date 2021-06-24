// MechModelsDlg.cpp : implementation file
//

#include "stdafx.h"			// standard inclusions
#include "Uppmem.h"
#include <stdlib.h>
#include <afxwin.h>

#include <afxdlgs.h>		// for saving files
#include <fstream.h>

#include "MechModelsDlg.h"		// for all necessarry 
#include "AddParamDlg.h"		// predefined objects
#include "ParamRangeDlg.h"
#include "PartDistribDlg.h"
#include "PreDefMemDlg.h"
#include "ErrBox.h"
#include "FluxRangeDlg.h"
#include "PermDistribDlg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



///////////////////////////////////////////////////////////////
// Non-member function, but used to open a CErrBox with a 
//   certain code, and display it
int show_err(int num)
{
	CErrBox box;		// initialize the box
	box.SetErr(num);	// set error number
	box.DoModal();		// show the box
	return 0;			// a hack to avoid linking errors
}



/////////////////////////////////////////////////////////////////////////////
// CMechModelsDlg dialog


CMechModelsDlg::CMechModelsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMechModelsDlg::IDD, pParent)
{
	model_type = 0;			// uninitialized model type

	// Additional Model Parameters Dialogue Box variables 
	amp_Cmgl = true;		// AMP conc is in mg/l
	amp_estimate = true;	// 
	user_k = 1e-6;			// user input value of k
	mgl_conc = 20000;		//					   mg/l conc
	vol_conc = 20;			//					   % conc

	//{{AFX_DATA_INIT(CMechModelsDlg)
	m_press = 100.0;
	m_temp = 20.0;
	m_Cav = 2000.0;
	m_Dav = 2.0;
	m_Rav = 0.025;
	m_Area = 1.0;
	m_CRad = 1.0;
	m_CLen = 1.0;
	m_MRes = 1e+11;
	m_PRad = 0.05;
	m_visc = 0.001;
	m_MSFlux = _T(" ");
	m_LHFlux = _T(" ");
	m_Circ = 50.0;
	m_find_reject = FALSE;
	m_Qin = 1800.0;
	//}}AFX_DATA_INIT

	// Particle Distribution Dialogue box parameters
	particles[0][0] = m_Rav / 1000000;	// size in meters
	particles[1][0] = m_Cav * 1000000 / (m_Dav * 4.1888 * pow(m_Rav,3));	// mg/L to #/ml
	part_num = 1;

	// Parameter range dialogue box values
	iterate = save_iterate = false;
	steps = increase = 0;
	param = -1;
	range[0] = range[1] = 0;

	// default additional model parameters
	gel_conc = 20000;
	amp_Cmgl = true;
	amp_estimate = true;
	user_k = 0.000001;
	op_res = 1e+11;
	ir_res = 1e+11;
}


void CMechModelsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMechModelsDlg)
	DDX_Text(pDX, IDC_MECH_PRESSURE, m_press);
	DDV_MinMaxDouble(pDX, m_press, 1.e-020, 100000.);
	DDX_Text(pDX, IDC_MECH_TEMP, m_temp);
	DDV_MinMaxDouble(pDX, m_temp, 1.e-020, 100.);
	DDX_Text(pDX, IDC_MECH_AVE_PART_CONC, m_Cav);
	DDV_MinMaxDouble(pDX, m_Cav, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MECH_AVE_PART_DENSITY, m_Dav);
	DDV_MinMaxDouble(pDX, m_Dav, 1.e-020, 1000.);
	DDX_Text(pDX, IDC_MECH_AVE_PART_RADIUS, m_Rav);
	DDV_MinMaxDouble(pDX, m_Rav, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MECH_MEMB_AREA, m_Area);
	DDV_MinMaxDouble(pDX, m_Area, 1.e-020, 10000.);
	DDX_Text(pDX, IDC_MECH_MEMB_CHANNEL_RADIUS, m_CRad);
	DDV_MinMaxDouble(pDX, m_CRad, 1.e-020, 1000.);
	DDX_Text(pDX, IDC_MECH_MEMB_LENGTH, m_CLen);
	DDV_MinMaxDouble(pDX, m_CLen, 1.e-020, 1000.);
	DDX_Text(pDX, IDC_MECH_MEMB_RESISTANCE, m_MRes);
	DDV_MinMaxDouble(pDX, m_MRes, 1.e-020, 1.e+020);
	DDX_Text(pDX, IDC_MECH_PORE_RADIUS, m_PRad);
	DDV_MinMaxDouble(pDX, m_PRad, 1.e-020, 1000000.);
	DDX_Text(pDX, IDC_MECH_VISCOSITY, m_visc);
	DDV_MinMaxDouble(pDX, m_visc, 1.e-020, 1.);
	DDX_Text(pDX, IDC_MECH_FLUX_MS, m_MSFlux);
	DDV_MaxChars(pDX, m_MSFlux, 50);
	DDX_Text(pDX, IDC_MECH_FLUX_LH, m_LHFlux);
	DDV_MaxChars(pDX, m_LHFlux, 50);
	DDX_Text(pDX, IDC_MECH_RECIRC, m_Circ);
	DDV_MinMaxDouble(pDX, m_Circ, 0., 1.e+020);
	DDX_Check(pDX, IDC_MECH_CALC_REJECT, m_find_reject);
	DDX_Text(pDX, IDC_MECH_INFLUENT_FLOW, m_Qin);
	DDV_MinMaxDouble(pDX, m_Qin, 0., 1.e+020);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CMechModelsDlg, CDialog)
	//{{AFX_MSG_MAP(CMechModelsDlg)
	ON_BN_CLICKED(IDC_MECH_PART_DISTRIBUTION, OnMechPartDistribution)
	ON_BN_CLICKED(IDC_MECH_CALC_FLUX, OnMechCalcFlux)
	ON_BN_CLICKED(IDC_MECH_ADDTL_MODEL_PARAMS, OnMechAddtlModelParams)
	ON_BN_CLICKED(IDC_MECH_PARAM_RANGE, OnMechParamRange)
	ON_BN_CLICKED(IDC_MECH_MEMB_SELECT, OnMechMembSelect)
	ON_BN_CLICKED(IDC_MECH_MEMSYS_RADIO, OnMechMemsysRadio)
	ON_BN_CLICKED(IDC_MECH_GEL_RADIO, OnMechGelRadio)
	ON_BN_CLICKED(IDC_MECH_RESISTANCE_RADIO, OnMechResistanceRadio)
	ON_BN_CLICKED(IDC_MECH_SE_RADIO, OnMechSERadio)
	ON_BN_CLICKED(ID_MECH_SAVE, OnMechSave)
//	ON_BN_CLICKED(ID_HELP, OnHelp)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CMechModelsDlg member functions:
void CMechModelsDlg::do_calcs()
{
	for (int i=0; i<steps; i++)
	{
		// predict flux with current values
		CalcFlux();
		flux[0][i] = atof((const char *)m_MSFlux);

		// record value for flux, and alter the desired parameter
		switch (param)
		{
		case 0:		flux[1][i] = m_press;
					m_press += increase;
		break;

		case 1:		flux[1][i] = m_temp;
					m_temp += increase;
		break;

		case 2:		flux[1][i] = m_Qin;
					m_Qin += increase;
		break;

		case 3:		flux[1][i] = m_visc;
					m_visc += increase;
		break;

		case 4:		flux[1][i] = m_Cav;
					m_Cav += increase;
		break;
		}

		// write new value to screen, and repeat
		UpdateData(false);
	}

	// display values on the screen with new dialogue box
	CFluxRangeDlg dlg;
	dlg.set_param(param, steps, flux);
	dlg.DoModal();

	iterate = false;
	save_iterate = true;
}



/////////////////////////////////////////////////////////////////////////////
// CMechModelsDlg message handlers

// Read in the values for membrane conditions, and the model 
//   type.  Calculate flux, and return it to the dialogue box.
void CMechModelsDlg::OnMechCalcFlux() 
{

	BeginWaitCursor(); // display the hourglass cursor

 	// double-flag system to enable calculating flux multiple times
	if (iterate == false)
	{		// user doesn't want to iterate
		save_iterate = false;

		// do calculations, see if user wants to find rejection
		if (CalcFlux() && m_find_reject)	
			CalcReject();
	}

	else	// user wants to iterate
		do_calcs();



	EndWaitCursor(); // remove the hourglass cursor
}


bool CMechModelsDlg::CalcFlux()
{
	double VolRatio, press, temp, Rav, Cav, 
			Dav, PRad, CRad, Circ, vlos;

	UpdateData(true);		// read data from dlg box

	// if they have not selected a flux model, tell them to do so
	if (model_type == 0)
	{
		show_err(100);		// no model slected
		return false;		// do not go further with routine
	}

	// If Memsys was picked, show info box
	if (model_type == 1)
	{
		CErrBox box;
		box.SetErr(12);					// info box code
		if (box.DoModal() != IDOK)		// show the box
			return false;				// do nothing if they hit cancel
	}


	// convert units from what they were in the dlg box to
	//   what they should be in the CMemCond object
	press = m_press * 1000;		// kPa to Pa
	temp = m_temp + 273;		// C to K
	Rav = m_Rav / 1000000;		// microns to meters
	VolRatio = m_Cav / (m_Dav * 1000000);				// (mg/L) / (cm^3/g)
	Cav = VolRatio * 3 / (4 * 3.1416 * Rav * Rav * Rav);	// to #/m^3
	Cav = Cav / 1000000;									// to #/ml
	PRad = m_PRad / 1000000;	// microns to meters
	CRad = m_CRad / 1000;		// mm to meters
	Circ = m_Circ / 100;		// % to fraction
	Dav = m_Dav;				// save as g/cm^3

	// flow in l/h to speed in m/s
	// velocity = Q / (Xarea of 1 fiber * # of fibers)
	vlos = (m_Qin / 3600000) * 2 * m_CLen / (CRad * m_Area);

	// set the values in mem
	mem.SetParam(press, temp, vlos, m_visc);
	mem.SetAves(Rav, Cav, Dav);
	mem.SetMem(PRad, m_MRes, CRad, m_CLen, m_Area, Circ);

	// switch flux models according to type selected
	if (model_type > 0 && model_type < 5)
	{
		module.MechModelInit(model_type, mem);

		// set the values in module
		if (amp_estimate)
		{
			module.SetGel(gel_conc, amp_Cmgl, 0);
			user_k = module.Kg;		// update value for user_k
		}
		else 
			module.SetGel(gel_conc, amp_Cmgl, user_k);

		// set resistance values
		module.SetRes(op_res, ir_res);

		error_num = module.FindFlux();  // returns errors, if any

		MSflux = module.GetFlux();			// flux in m/s
		LHflux = module.GetFlux() * m_Area * 3600000;	// flux in l/h
		

		// Deal with all the errors - display them, and make sure that
		//   they will not force the program to crash.
		if (error_num < 0)		// fairly serious error
		{	// do not return undefined data - will cause problems
			MSflux = LHflux = 0;
		}
		
		else if (!(LHflux < 1e20 && LHflux >= 0))
		{	// Flux values are out of range, therefore WAY off.
			MSflux = LHflux = 0;
			error_num += 2000;	// signifies an out of range error, in
		}						//   addition to any other error passed


		if (error_num != 1)
			show_err(error_num);
		
		// convert number values to character strings (for easy rounding purposes)
		char buffer[20];
		_gcvt(MSflux, 5, buffer);
		m_MSFlux = (CString) buffer;
		_gcvt(LHflux, 5, buffer);
		m_LHFlux = (CString) buffer;

		// Write flux to the screen
		UpdateData(false);		

	}

	return true;

}




void CMechModelsDlg::OnMechSave() 
{
	// saving should only be allowed after the user predicts flux,
	//   but if the model choice is Memsys, or problems can occurr
	
	// if last button pushed was to get the flux, save automatically,
	//   otherwise give a warning
	// how to do that?

	// open up the standard "Save As" dialogue box
	static char filterCode[] = "Mechanistic Uppmem Model (*.mec)|*.mec|Data Files (*.dat)|*.dat|Text Files (*.txt)|*.txt|All Files (*.*)|*.*||";
	CFileDialog FileSaveDialog(FALSE, "mec", "module.mec",
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, (LPCTSTR)filterCode);
	if (FileSaveDialog.DoModal() == IDOK)
	{
		CString NewFileName = FileSaveDialog.GetFileName();
		CString NewPathName = FileSaveDialog.GetPathName();
		
		// create a file with a temporary name
		ofstream fout;
		fout.open((const char*)NewFileName);

		// write all data to it in specific format, one tab = 6 spaces
		fout << "Mechanistic Model Flux Prediction, and Operating Parameters" << endl;
		fout << "Model: ";
		switch (model_type)
		{
		case  0:	fout << "None" << endl;
		break;

		case  1:	fout << "Memsys" << endl;
		break;

		case  2:	fout << "Song & Elimelech" << endl;
		break;

		case  3:	fout << "Resistance" << endl;
		break;

		case  4:	fout << "Gel Polarization" << endl;
		break;
		}

		fout << endl << "Operating Conditions:" << endl;
		fout << "Pressure:\t\t\t" << m_press << " kPa" << endl;
		fout << "Temperature:\t\t" << m_temp << " C" << endl;
		fout << "Influent Flow:\t\t\t" << m_Qin << " l/h" << endl;
		fout << "Viscosity:\t\t\t" << m_visc << " kg/m*s" << endl;

		fout << endl << "Particle Distribution:" << endl;
		fout << "Density:\t\t\t" << m_Dav << " g/cm^3" << endl;
		fout << "Number:\t\t\t" << part_num << endl;
		fout << "Size (microns):\t\tConcentration(#/ml)" << endl;
		int i;
		for (i=0; i<part_num; i++)
			fout << (particles[0][i] * 1000000) << "\t\t\t\t" << (particles[1][i]) << endl;

		// if user wanted to analyze rejection, save that too
		if (m_find_reject)
		{
			fout << endl << "Permeate Particle distribution:" << endl;
			for (i=0; i<part_num; i++)
				fout << (permeate[0][i] * 1000000) << "\t\t\t\t" << (particles[1][i]) << endl;

			fout << endl << "Retentate Particle distribution:" << endl;
			for (i=0; i<part_num; i++)
				fout << (retentate[0][i] * 1000000) << "\t\t\t\t" << (particles[1][i]) << endl;
		}

		fout << endl << "Membrane Parameters:" << endl;
		fout << "Pore Radius:\t\t" << m_PRad << " microns" << endl;
		fout << "Resistance:\t\t\t" << m_MRes << " 1/m" << endl;
		fout << "Channel Radius:\t\t" << m_CRad << " mm" << endl;
		fout << "Channel Length:\t\t" << m_CLen << " m" << endl;
		fout << "Area:\t\t\t\t" << m_Area << " m^2" << endl;
		fout << "Recirculation:\t\t\t" << m_Circ << " %" << endl;

		if (model_type > 2)
		{
			fout << endl << "Additional Model Parameters:" << endl;
			fout << "Gel Concentration:\t\t" << gel_conc;
			if (amp_Cmgl)
				fout << " mg/L" << endl;
			else
				fout << " %" << endl;
			fout << "Mass Transfer Coefft:\t\t" << user_k << " m/s" << endl;
		}

		if (!save_iterate)
		{
			fout << endl << "Flux:" << endl;
			fout << m_MSFlux << "\t\tm/s" << endl;
			fout << m_LHFlux << "\tl/h*m^2" << endl;
		}
		else
		{
			if (param == 0)
				fout << endl << "Pressure (kPa):\t\t";
			else if (param == 1)
				fout << endl << "Temp (C):\t\t";
			else if (param == 2)
				fout << endl << "Influent Flow (l/h):\t\t";
			else if (param == 3)
				fout << endl << "Viscosity (kg*m/s):\t\t";
			fout << "Flux (m/s):" << endl;

			for (int i=0; i<steps; i++)
				fout << (flux[1][i]) << "\t\t\t\t" << (flux[0][i]) << endl;
		}

		// close the file
		fout.close();
	}
}





////////////////////////////////////////////////////////////////
//  Sub-dialogue boxes
//  Used optionally for additional parameters
////////////////////////////////////////////////////////////////
void CMechModelsDlg::OnMechParamRange() 
{
	CParamRangeDlg dlg;

	iterate = true;

	// store current values
	UpdateData(true);

	// Invoke the dialog box
	if (dlg.DoModal() == IDOK)
	{
		switch (dlg.param)
		{
		case 0:		m_press  = range[0] = dlg.m_loPress;
					range[1] = dlg.m_hiPress;
		break;

		case 1:		m_temp   = range[0] = dlg.m_loTemp;
					range[1] = dlg.m_hiTemp;
		break;

		case 2:		m_Qin  = range[0] = dlg.m_loQin;
					range[1] = dlg.m_hiQin;
		break;

		case 3:		m_visc   = range[0] = dlg.m_loVisc;
					range[1] = dlg.m_hiVisc;
		break;
		
		case 4:		m_Cav	 = range[0] = dlg.m_loConc;
					range[1] = dlg.m_hiConc;
		break;

		default:	iterate = false;
		break;
		}

		steps = dlg.m_steps;
		param = dlg.param;
		increase = (range[1] - range[0]) / steps;

		// write new membrane parameters to the screen
		UpdateData(false);
	}
}



void CMechModelsDlg::OnMechPartDistribution() 
{

	CPartDistribDlg dlg;

	// read in values for particles, and density
	UpdateData(true);

	// initialize dlg with list of particle sizes,
	//   the number of particles, and density
	dlg.SetList(particles, part_num, m_Dav);

	// Invoke the dialog box
	if (dlg.DoModal() == IDOK)
	{
		part_num = dlg.m_partNum;

    	// update values for particles[2][20]
		for (int i = 0; i < (part_num); i++)
		{
			particles[0][i] = dlg.m_particles[0][i];
			particles[1][i] = dlg.m_particles[1][i];
		}

		// do not set values in mem until user predicts flux!

		// read in current values
		UpdateData(true);

		// retrieve and set average conc and radius (round them?)
		m_Cav = dlg.ave_conc;
		m_Rav = dlg.ave_rad;

		// round the values to 6 digits
		char buffer[50];
		m_Cav  = atof((_gcvt(m_Cav, 6, buffer)));
		m_Rav  = atof((_gcvt(m_Rav, 6, buffer)));

		// display them on the screen
		UpdateData(false);
	}
}



void CMechModelsDlg::OnMechMembSelect() 
{	
	CPreDefMemDlg dlg;

	// store current values
	UpdateData(true);

	// Invoke the dialog box
	if (dlg.DoModal() == IDOK)
	{
		// update the membrane parameters
		m_Area = dlg.m_area;
		m_CRad = dlg.m_crad;
		m_CLen = dlg.m_length;
		m_MRes = dlg.m_resist;
		m_PRad = dlg.m_prad;
//		m_Rec  = dlg.m_rec;
	}

	// write new membrane parameters to the screen
	UpdateData(false);
}




void CMechModelsDlg::OnMechAddtlModelParams() 
{
	CAddParamDlg ampdlg;

	// store current values
	UpdateData(true);

	// Adjust variables in dialogue box to match current values
	if (amp_Cmgl == true)			// if conc is in mg/L
		ampdlg.m_mgl_radio = 0;		// '0' means button is pressed
	else
		ampdlg.m_vol_radio = 0;

	if (amp_estimate == true)		// if user wanted k estimated
		ampdlg.m_est_radio = 0;
	else
		ampdlg.m_ent_radio = 0;

	ampdlg.m_mgl_conc = mgl_conc;
	ampdlg.m_vol_conc = vol_conc;
	ampdlg.m_op_res = op_res;			
	ampdlg.m_ir_res = ir_res;			

	// Convert user_k to a string
	char buffer[20];
	_gcvt(user_k, 5, buffer);
	ampdlg.m_K_str = (CString) buffer;
	

	// Invoke the dialog box
	if (ampdlg.DoModal() == IDOK)
	{
		mgl_conc = ampdlg.m_mgl_conc;
		vol_conc = ampdlg.m_vol_conc;

		if (ampdlg.m_mgl_radio == 0)	// user enters Cg in mg/L
		{
			gel_conc = ampdlg.m_mgl_conc;
			amp_Cmgl = true;
		}
		else			// user enters Cg in %
		{
			gel_conc = ampdlg.m_vol_conc / 100;
			amp_Cmgl = false;
		}

		if (ampdlg.m_est_radio == 0)	// user wants to estimate k
			amp_estimate = true;
		else
			amp_estimate = false;


		// update resistances and K value
		op_res = ampdlg.m_op_res;		
		ir_res = ampdlg.m_ir_res;
		user_k = atof(ampdlg.m_K_str);
	}

}




////////////////////////////////////////////////////////////////
// Radio controls: determines which model to use.
//   Probably not the most efficient use of member 
//   functions, but it will suffice.
////////////////////////////////////////////////////////////////
void CMechModelsDlg::OnMechMemsysRadio() 
{
	model_type = 1;
}

void CMechModelsDlg::OnMechSERadio() 
{
	model_type = 2;	
}

void CMechModelsDlg::OnMechResistanceRadio() 
{
	model_type = 3;
}

void CMechModelsDlg::OnMechGelRadio() 
{
	model_type = 4;	
}



// what to do when they want to calculate rejection
void CMechModelsDlg::CalcReject() 
{
	double rec;		// water recovery: permeate flow / feed flow
	double L;		// lambda: solute size/pore size
	double SC;		// sieve constant
	double msflux;

	// convert m_MSFlux to a number (m/s)
	msflux = atof((const char *)m_MSFlux);

	// determine water recovery (Qperm/Qinf)
	// MemCond object must be set!
	rec = msflux * mem.Area / (m_Qin / 3600000);
//	rec = msflux * 2 * mem.CLen / mem.veloc * mem.CRad; 
	
	// for each particle in "particles", determine sieve
	//   coefficient.  Use that to find eff. and reten. conc.
	for (int i = 0; i < part_num; i++)
	{
		L = particles[0][i] / mem.PRad;

		if (L>1)
			SC = 0.0;
		else
			SC = (1-L)*(1-L) * ( 2 - (1-L)*(1-L)) * exp(-.7472*L*L);

		// set particle sizes
		(retentate[0][i]) = (permeate[0][i]) = (particles[0][i]);

		// set concentrations
		retentate[1][i] = (particles[1][i]) / (1 - ((1-SC) * rec));
		permeate[1][i] = (retentate[1][i]) * SC;
	}

	// set the new values in the MemCond object
	mem.SetDstb(particles, permeate, retentate, part_num);

	// open up a new box and show the values
	CPermDistribDlg pd_dlg;
	pd_dlg.set_list(mem);
	pd_dlg.DoModal();
}


/*
void CMechModelsDlg::OnHelp() 
{
	// Open the main html help page with all the 
	//   mechanistic modelling info
//	int error = system("open help\\mech_main.htm");
	HINSTANCE error =  ShellExecute(NULL, "open", 
		"help\\mech_main.htm", NULL, NULL, SW_SHOWNORMAL);
	HINSTANCE error =  ShellExecute(NULL, "open", 
		"help\\mech_main.htm", NULL, NULL, SW_SHOWNORMAL);

	// if an error occurred, perhaps file cant be found
	if ((const long)error == ERROR_FILE_NOT_FOUND ||
		(const long)error == ERROR_PATH_NOT_FOUND)
	{
		// show an error box stating the problem
		// allow the user to find help folder?
	}
}
*/
