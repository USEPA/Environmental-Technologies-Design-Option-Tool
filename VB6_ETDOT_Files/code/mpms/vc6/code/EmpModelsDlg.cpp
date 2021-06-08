// EmpModelsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "ErrBox.h"
#include "FluxRangeDlg.h"
#include "EmpModelsDlg.h"
#include "AddParamDlg.h"
#include "ParamRangeDlg.h"
#include "PartDistribDlg.h"
#include "PreDefMemDlg.h"
#include <math.h>
#include <fstream.h>
#include <stdlib.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



///////////////////////////////////////////////////////////////
// Non-member function, but used to open a CErrBox with a 
//   certain code, and display it
void show_err(int num)
{
	CErrBox box;		// initialize the box
	box.SetErr(num);	// set error number
	box.DoModal();		// show the box
}



/////////////////////////////////////////////////////////////////////////////
// CEmpModelsDlg dialog


CEmpModelsDlg::CEmpModelsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CEmpModelsDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CEmpModelsDlg)
	m_conc = 0.0;
	m_conc_units = _T("mg/L");
	m_pres = 100.0;
	m_temp = 20.0;
	m_visc = 0.001;
	m_lhFlux = 0.0;
	m_msFlux = 0.0;
	m_tFlux = 0.0;
	m_cTime = 180.0;
	m_pTime = 1800.0;
	m_a_name = _T("");
	m_b_name = _T("");
	m_c_name = _T("");
	m_a_val = 0.0;
	m_b_val = 0.0;
	m_c_val = 0.0;
	m_Qin = 0.0;
	//}}AFX_DATA_INIT
(m_Qin * 3.6) / (3.1416 * (module.mem).CRad * (module.mem).CRad);

	iterate = false;
	custom = false;
}


void CEmpModelsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEmpModelsDlg)
	DDX_Text(pDX, IDC_EMP_AVE_PART_CONC, m_conc);
	DDV_MinMaxDouble(pDX, m_conc, 0., 1000000.);
	DDX_Text(pDX, IDC_EMP_CONC_UNITS, m_conc_units);
	DDV_MaxChars(pDX, m_conc_units, 10);
	DDX_Text(pDX, IDC_EMP_PRESSURE, m_pres);
	DDV_MinMaxDouble(pDX, m_pres, 0., 10000000000.);
	DDX_Text(pDX, IDC_EMP_TEMP, m_temp);
	DDV_MinMaxDouble(pDX, m_temp, 0., 100.);
	DDX_Text(pDX, IDC_EMP_VISCOSITY, m_visc);
	DDV_MinMaxDouble(pDX, m_visc, 0., 1.);
	DDX_Text(pDX, IDC_EMP_FLUX_LH, m_lhFlux);
	DDV_MinMaxDouble(pDX, m_lhFlux, -1.79769e+308, 1.79769e+308);
	DDX_Text(pDX, IDC_EMP_FLUX_MS, m_msFlux);
	DDV_MinMaxDouble(pDX, m_msFlux, -1.79769e+308, 1.79769e+308);
	DDX_Text(pDX, IDC_EMP_FLUX_TIME, m_tFlux);
	DDV_MinMaxDouble(pDX, m_tFlux, -1.79769e+308, 1.79769e+308);
	DDX_Text(pDX, IDC_EMP_CLEAN_TIME, m_cTime);
	DDV_MinMaxDouble(pDX, m_cTime, 0., 1.e+015);
	DDX_Text(pDX, IDC_EMP_PERM_TIME, m_pTime);
	DDV_MinMaxDouble(pDX, m_pTime, 0., 1.e+015);
	DDX_Text(pDX, IDC_EMP_PARAM_A_NAME, m_a_name);
	DDV_MaxChars(pDX, m_a_name, 30);
	DDX_Text(pDX, IDC_EMP_PARAM_B_NAME, m_b_name);
	DDV_MaxChars(pDX, m_b_name, 30);
	DDX_Text(pDX, IDC_EMP_PARAM_C_NAME, m_c_name);
	DDV_MaxChars(pDX, m_c_name, 30);
	DDX_Text(pDX, IDC_EMP_PARAM_A_VAL, m_a_val);
	DDV_MinMaxDouble(pDX, m_a_val, -1.e+300, 1.e+300);
	DDX_Text(pDX, IDC_EMP_PARAM_B_VAL, m_b_val);
	DDV_MinMaxDouble(pDX, m_b_val, -1.e+300, 1.e+300);
	DDX_Text(pDX, IDC_EMP_PARAM_C_VAL, m_c_val);
	DDV_MinMaxDouble(pDX, m_c_val, -1.e+300, 1.e+300);
	DDX_Text(pDX, IDC_EMP_INFLUENT_FLOW, m_Qin);
	DDV_MinMaxDouble(pDX, m_Qin, 0., 1.e+020);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CEmpModelsDlg, CDialog)
	//{{AFX_MSG_MAP(CEmpModelsDlg)
	ON_BN_CLICKED(IDC_EMP_EXP_DATA, OnEmpExpData)
	ON_BN_CLICKED(IDC_EMP_PARAM_RANGE, OnEmpParamRange)
	ON_BN_CLICKED(IDC_EMP_CALC_FLUX, OnEmpCalcFlux)
	ON_BN_CLICKED(IDC_EMP_CUST_MODEL_VAL, OnUseCustomVal)
	ON_BN_CLICKED(ID_EMP_SAVE, OnEmpSave)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEmpModelsDlg message handlers



// standard stuff here



/////////////////////////////////////////////////////////////////////////////
// Deal with other sub-dialogue boxes here
/////////////////////////////////////////////////////////////////////////////

// call this function before opening the dlg box
void CEmpModelsDlg::set_units(bool c_mgl, double C, double P, 
							  double T, double V)
{
	m_conc = C;
	m_pres = P;
	m_temp = T;
	m_Qin  = V * 3600000 * ((module.mem).CRad * (module.mem).Area) 
				/ (2 * (module.mem).CLen);

	// get and set ModelType
	module.get_Emp_model(ModelType);

	switch(ModelType)	// 1=surf ren, 2=fouling, 3=resistance, 4=gel
	{
	// parameter order:		a, pb, pc
	// surf renewal:		Js, A, s
	// resistance fouling:	Ra, Rpp, Rcp
	// resistance:			Rf, Rop
	// gel polarization:	k, Cg
	case 1:		m_a_name = "Min Flux (m/s)";
				m_b_name = "Flux Decline (1/s)";
				m_c_name = "Surface Renewal (1/s)";
	break;

	case 2:		m_a_name = "Ra (1/m)";
				m_b_name = "Rpp (1/m)";
				m_c_name = "Rcp (1/m)";
	break;

	case 3:		m_a_name = "Rf (1/m)";
				m_b_name = "Rop (1/m*Pa)";
				m_c_name = "NA";
	break;

	case 4:		m_a_name = "k (m/s)";
				m_b_name = "Cg ";
				m_c_name = "NA";
				if (mgl)
					m_b_name += "(mg/l)";
				else
					m_b_name += "(%)";
	break;
	}


	// get and set concentration units
	mgl = c_mgl;

	if (mgl)
		m_conc_units = "mg/l";
	else
		m_conc_units = "%";

}


void CEmpModelsDlg::OnEmpExpData() 
{
	// TODO: what to do when experimental data is entered?
	// nothing until I have all that figured out!!!
}


void CEmpModelsDlg::OnEmpParamRange() 
{
	CParamRangeDlg dlg;
	iterate = true;

	// Invoke the dialog box
	if (dlg.DoModal() == IDOK)
	{
		switch (dlg.param)
		{
		case 0:		m_pres  = range[0] = dlg.m_loPress;
					range[1] = dlg.m_hiPress;
		break;

		case 1:		m_temp   = range[0] = dlg.m_loTemp;
					range[1] = dlg.m_hiTemp;
		break;

		case 2:		m_Qin = range[0] = dlg.m_loQin;
					range[1] = dlg.m_hiQin;
		break;

		case 3:		m_visc   = range[0] = dlg.m_loVisc;
					range[1] = dlg.m_hiVisc;
		break;
		
		case 4:		m_conc   = range[0] = dlg.m_loConc;
					range[1] = dlg.m_hiConc;
		break;

		default:	iterate = false;
		break;
		}

		steps = dlg.m_steps;
		param = dlg.param;
		increase = (range[1] - range[0]) / steps;

		// write new parameter values to the screen
		UpdateData(false);
	}
}



void CEmpModelsDlg::OnEmpCalcFlux() 
{
	int error_num, compat;

	// read in the parameter values
	UpdateData(true);

	// convert pressure(kPa) to Pa
	double pres = m_pres * 1000;

	// convert flow (l/h) to a velocity (m/s),
	double vlos = (m_Qin / 3600000) * 2 * (module.mem).CLen /
					( (module.mem).CRad * (module.mem).Area);

	if (!iterate)	// if user isn't testing a range of parameters
	{

		// read in the data values, call get_flux()
		if (custom)
			error_num = get_flux(m_msFlux);
		else
			error_num = module.get_flux(m_msFlux, compat, m_conc, 
							pres, vlos, m_temp, m_visc, m_pTime);

		// convert to liters/hour
		m_lhFlux = m_msFlux * ((module.mem).Area) * 3600000;

		// determine the time averaged flux
		if (m_cTime == 0 && m_pTime == 0)
			m_tFlux = 0;
		else
			m_tFlux = m_lhFlux * (m_pTime / (m_pTime + m_cTime));

	}


	else
	{
		for (int i=0; i < steps; i++)
		{

			if (custom)
				error_num = get_flux(m_msFlux);
			else
				error_num = module.get_flux(m_msFlux, compat, 
						m_conc, m_pres, vlos, m_temp, m_visc, m_pTime);


			// set the values in flux[][]
			flux[0][i] = m_msFlux;
			switch(param)	// 0=Press, 1=Temp, 2=Velocity, 3=Visc., 4=conc
			{
			case 0:	flux[1][i] = m_pres;
					m_pres += increase;
			break;

			case 1:	flux[1][i] = m_temp;
					m_temp += increase;
			break;

			case 2: flux[1][i] = m_Qin;
					m_Qin += increase;
					vlos = (m_Qin*3.6) / (3.1415 * (module.mem).CRad * (module.mem).CRad);
			break;

			case 3:	flux[1][i] = m_visc;
					m_visc += increase;
			break;

			case 4:	flux[1][i] = m_conc;
					m_conc += increase;
			break;
			}

		}		// end for loop

		// display values on the screen with new dialogue box
		CFluxRangeDlg dlg;
		dlg.set_param(param, steps, flux);
		dlg.DoModal();

	}			// end else


	// display errors
	if (error_num >= 1000)		// flux is negative
	{
		show_err(error_num);
		error_num = error_num - 1000;
	}

	if (error_num != 1)			// problem finding flux
		show_err(error_num);

	if (!custom && compat != 353)		// not interpolation
		show_err(compat);



	// Update parameter values if not entered by user,
	//   round them to 6 digits
	char buffer[50];

	if (!custom)
	{
		m_a_val = atof((_gcvt((module.param_a), 6, buffer)));
		m_b_val = atof((_gcvt((module.param_b), 6, buffer)));
		m_c_val = atof((_gcvt((module.param_c), 6, buffer)));
	}

	// round all flux values set by the computer
	m_tFlux  = atof((_gcvt(m_tFlux, 6, buffer)));
	m_lhFlux = atof((_gcvt(m_lhFlux, 6, buffer)));
	m_msFlux = atof((_gcvt(m_msFlux, 6, buffer)));

	// write new data to screen
	UpdateData(false);
}


// determine flux when the user inputs the model params
int CEmpModelsDlg::get_flux(double &flux)
{
	// parameter order:		m_a_val, m_b_val, m_c_val
	// surf renewal:		Js, A, s
	// resistance fouling:	Ra, Rpp, Rcp
	// resistance:			Rf, Rop
	// gel polarization:	k, Cg

	// test for errors: zeros, negatives, etc.
	if (m_a_val <= 0 || m_b_val <=0 || ((ModelType == 1 || ModelType == 2) && m_c_val <=0))
		return 360;		// invalid parameter values

	int flag;

	// switch according to model type
	switch(ModelType)
	{
	case 1:		// Surface Renewal
//		fix this
		flag = module.get_surf_flux(flux, m_a_val, m_b_val, m_c_val, m_pTime);
		break;

	case 2:		// Resistance Fouling (convert kPa to Pa)
		flux = m_pres*1000 / (m_visc * (((module.mem).MRes) + m_a_val + m_b_val + m_c_val));
		break;

	case 3:		// Resistance (connvert kPa to Pa)
		flux = m_pres*1000 / (m_visc * (((module.mem).MRes) + m_a_val + m_pres * 1000 * m_b_val));
		break;

	case 4:		// Gel Polarization
		flux = m_a_val * log (m_b_val / m_conc);
		break;
	}

	return 1;

}


void CEmpModelsDlg::RoundAll()
{
	char buffer[50];
	
	// round all the values updated by the computer to 6 digits
	m_a_val  = atof((_gcvt(m_a_val, 6, buffer)));
	m_b_val  = atof((_gcvt(m_b_val, 6, buffer)));
	m_c_val  = atof((_gcvt(m_c_val, 6, buffer)));
	m_tFlux  = atof((_gcvt(m_tFlux, 6, buffer)));
	m_lhFlux = atof((_gcvt(m_lhFlux, 6, buffer)));
	m_msFlux = atof((_gcvt(m_msFlux, 6, buffer)));
}



// if the user hits the button "Use Custom Values"
void CEmpModelsDlg::OnUseCustomVal() 
{
	if (custom)
		custom = false;
	else
		custom = true;
}


void CEmpModelsDlg::OnEmpSave() 
{
	// open up the standard Save As... dialogue box
	static char filterCode[] = "Empirical Uppmem Model (*.emp)|*.emp|Data Files (*.dat)|*.dat|Text Files (*.txt)|*.txt|All Files (*.*)|*.*||";
	CFileDialog FileSaveDialog(FALSE, "emp", "module.emp",
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, (LPCTSTR)filterCode);
	if (FileSaveDialog.DoModal() == IDOK)
	{
		// get the name and path for the file they want to save
		CString NewFileName = FileSaveDialog.GetFileName();
		CString NewPathName = FileSaveDialog.GetPathName();
		
		// create an output file with a temporary name
		fstream fout;
		fout.open((const char*)NewFileName, ios::out);

		// write all data to it in specific format, one tab = 6 spaces
		fout << "Empirical Model Flux Prediction, and Operating Parameters" << endl;
		fout << "Model: ";

		switch (ModelType)
		{
		case  1:	fout << "Surface Renewal" << endl;
		break;

		case  2:	fout << "Fouling Resistance" << endl;
		break;

		case  3:	fout << "Resistance" << endl;
		break;

		case  4:	fout << "Gel Polarization" << endl;
		break;
		}

		fout << endl << "Operating Conditions:" << endl;
		fout << "Pressure:\t\t" << m_pres << " kPa" << endl;
		fout << "Temperature:\t\t" << m_temp << " C" << endl;
		fout << "Influent Flow:\t\t" << m_Qin << " l/h" << endl;
		fout << "Viscosity:\t\t" << m_visc << " kg/m*s" << endl;
		fout << "Concentration:\t\t" << m_conc;
		if (mgl)
			fout << " mg/l" << endl;
		else
			fout << " %" << endl;

		fout << "Permeation time:\t" << m_pTime << " s" << endl;
		fout << "Cleaning time:\t\t" << m_cTime << " s" << endl;

		fout << endl << "Membrane Parameters:" << endl;
		fout << "Pore Radius:\t\t" << ((module.mem).PRad) << " microns" << endl;
		fout << "Resistance:\t\t" << ((module.mem).MRes) << " 1/m" << endl;
		fout << "Channel Radius:\t\t" << ((module.mem).CRad) << " mm" << endl;
		fout << "Channel Length:\t\t" << ((module.mem).CLen) << " m" << endl;
		fout << "Area:\t\t\t" << ((module.mem).Area) << " m^2" << endl;
		fout << "Recirculation:\t\t" << ((module.mem).Circ) << " %" << endl;

		fout << endl << "User input model coefficients:" << endl;
		switch(ModelType)
		{
		case 1:	fout << "Min Flux:\t\t\t" << m_a_val << " m/s" << endl;
				fout << "Flux Decline:\t\t\t" << m_b_val << " 1/s" << endl;
				fout << "Surface Renewal:\t\t"<< m_c_val << " 1/s" << endl;
		break;

		case 2:	fout << "Ra:\t\t\t" << m_a_val << " 1/m" << endl;
				fout << "Rpp:\t\t\t" << m_b_val << " 1/m" << endl;
				fout << "Rcp:\t\t\t" << m_c_val << " 1/m" << endl;
		break;

		case 3:	fout << "Rf:\t\t\t" << m_a_val << " 1/m" << endl;
				fout << "Rop:\t\t\t" << m_b_val << " 1/m" << endl;
		break;

		case 4:	fout << "k:\t\t\t" << m_a_val << " m/s" << endl;
				fout << "Cg:\t\t\t" << m_b_val;
				if (mgl)
					fout << " mg/l" << endl;
				else
					fout << " %" << endl;
		break;
		}

		// now output all the data points from all the
		//   EmpData objects
		fout << endl << endl << endl << "EXPERIMENTAL DATA:" << endl;

		// '<<' is overloaded in the CEmpModel file to just output the
		//   user-input experimental data
		fout << module << endl;

	}		// end if
			// else? do nothing

}



