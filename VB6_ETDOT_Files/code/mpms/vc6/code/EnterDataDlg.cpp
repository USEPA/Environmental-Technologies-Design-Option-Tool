// EnterDataDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "ErrBox.h"
#include "EnterDataDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEnterDataDlg dialog


CEnterDataDlg::CEnterDataDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CEnterDataDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CEnterDataDlg)
	m_conc = 2000.0;
	m_flux = _T("");
	m_param = _T("");
	m_param_name = _T("");
	m_pres = 100.0;
	m_temp = 20.0;
	m_vlos = 1.0;
	m_conc_radio = -1;
	m_pres_radio = -1;
	m_temp_radio = -1;
	m_time_radio = -1;
	m_vlos_radio = -1;
	m_info = _T("Enter Experimental Flux Data");
	m_mgl_radio =  0;
	m_vol_radio = -1;
	//}}AFX_DATA_INIT

	alter = true;
	mgl = true;			// initial state, mg/L is selected
}


// called when a second or third data set is entered for the same
//   model, ie. for fouling resistance, or surface renewal
void CEnterDataDlg::set_param(double C, double P, 
							  double V, double T, bool M)
{
	// copy the values for each of the parameters
	m_conc = C;
	m_pres = P;
	m_vlos = V;
	m_temp = T;
	mgl = M;

	if (mgl)
	{
		m_mgl_radio = 0;
		m_vol_radio = -1;
	}

	else
	{
		m_mgl_radio = -1;
		m_vol_radio = 0;
	}
}


// set the Model Type, and primary value.
// call this before opening the dlg box
void CEnterDataDlg::set_data(int type, bool prime)
{
	ModelType = type;
	primary = prime;

	switch (type)
	{
	case 1:		// surface renewal, 1st data set
		m_param_name = "Time (s)";
		m_time_radio = 0;
		m_info = "Enter Dead-End Experimental Flux Data";
		param = 5;
		break;

	case 2:		// fouling resistance, 1st data set
		m_param_name = "Pressure (kPa)";
		m_pres_radio = 0;
		param = 2;
		m_info = "Enter Pure Water Flux Data After Static Adsorption Fouling";
		break;

	case 3:		// resistance
		m_param_name = "Pressure (kPa)";
		m_pres_radio = 0;
		param = 2;
		break;

	case 4:		// gel polarization, first set
		m_param_name = "Concentration";
		m_conc_radio = 0;
		param = 1;
		break;

	case 11:	// surface renewal, 2nd data set, do not allow parameters to be adjusted
		m_param_name = "Time (s)";
		m_info = "Enter Cross-Flow Experimental Flux Data";
		m_time_radio = 0;
		param = 5;
		alter = false;
		ModelType = 1;		// reset model type to surf ren
							// disable edit boxes
		break;

	case 12:	// fouling resistance, 2nd data set, do not allow parameters to be adjusted
		m_param_name = "Pressure (kPa)";
		m_info = "Enter Permeate Flux Data";
		m_pres_radio = 0;
		param = 2;
		alter = false;
		ModelType = 2;		// reset model type to fouling res
							// disable edit boxes
		break;

	case 22:	// fouling resistance, 3rd data set, do not allow parameters to be adjusted
		m_param_name = "Pressure (kPa)";
		m_info = "Enter Pure Water Flux Data After Permeation Fouling";
		m_pres_radio = 0;
		param = 2;
		alter = false;
		ModelType = 2;		// reset model type to fouling res
							// disable edit boxes
		break;
	}
}


void CEnterDataDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEnterDataDlg)
	DDX_Text(pDX, IDC_EMP_ENT_CONC, m_conc);
	DDV_MinMaxDouble(pDX, m_conc, 0., 1000000.);
	DDX_Text(pDX, IDC_EMP_ENT_FLUX_DATA, m_flux);
	DDV_MaxChars(pDX, m_flux, 10000);
	DDX_Text(pDX, IDC_EMP_ENT_PARAM2_DATA, m_param);
	DDV_MaxChars(pDX, m_param, 10000);
	DDX_Text(pDX, IDC_EMP_ENT_PARAM2_NAME, m_param_name);
	DDV_MaxChars(pDX, m_param_name, 30);
	DDX_Text(pDX, IDC_EMP_ENT_PRES, m_pres);
	DDV_MinMaxDouble(pDX, m_pres, 0., 1000000.);
	DDX_Text(pDX, IDC_EMP_ENT_TEMP, m_temp);
	DDV_MinMaxDouble(pDX, m_temp, 0., 100.);
	DDX_Text(pDX, IDC_EMP_ENT_VLOS, m_vlos);
	DDV_MinMaxDouble(pDX, m_vlos, 0., 1000.);
	DDX_Radio(pDX, IDC_EMP_ENT_CONC_RADIO, m_conc_radio);
	DDX_Radio(pDX, IDC_EMP_ENT_PRES_RADIO, m_pres_radio);
	DDX_Radio(pDX, IDC_EMP_ENT_TEMP_RADIO, m_temp_radio);
	DDX_Radio(pDX, IDC_EMP_ENT_TIME_RADIO, m_time_radio);
	DDX_Radio(pDX, IDC_EMP_ENT_VLOS_RADIO, m_vlos_radio);
	DDX_Text(pDX, IDC_EMP_INFO, m_info);
	DDV_MaxChars(pDX, m_info, 200);
	DDX_Radio(pDX, IDC_EMP_ENT_MGL_RADIO, m_mgl_radio);
	DDX_Radio(pDX, IDC_EMP_ENT_VOL_RADIO, m_vol_radio);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CEnterDataDlg, CDialog)
	//{{AFX_MSG_MAP(CEnterDataDlg)
	ON_BN_CLICKED(IDC_EMP_ENT_CONC_RADIO, OnEmpEntConcRadio)
	ON_BN_CLICKED(IDC_EMP_ENT_MGL_RADIO, OnEmpEntMglRadio)
	ON_BN_CLICKED(IDC_EMP_ENT_PRES_RADIO, OnEmpEntPresRadio)
	ON_BN_CLICKED(IDC_EMP_ENT_TEMP_RADIO, OnEmpEntTempRadio)
	ON_BN_CLICKED(IDC_EMP_ENT_VLOS_RADIO, OnEmpEntVlosRadio)
	ON_BN_CLICKED(IDC_EMP_ENT_VOL_RADIO, OnEmpEntVolRadio)
	ON_BN_CLICKED(IDC_EMP_ENT_TIME_RADIO, OnEmpEntTimeRadio)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEnterDataDlg message handlers

// this function determines whether or not to issue errors
//   if primary=true, and the user tries to change param 2
bool CEnterDataDlg::check(int prm)
{
	if (!primary)
		return true;	// no error occurred

	int err = 0;

	// surface renewal
	if (ModelType == 1)
	{
		err = 202;			// param must be time
		m_time_radio = 0;	// "push" time button
	}

	// resistance fouling or resistance
	else if (ModelType == 2 || ModelType == 3 && prm != 2)
	{
		err = 201;			// param must be pressure
		m_pres_radio = 0;	// "push" pressure button
	}

	// gel polarization
	else if (ModelType == 4 && prm != 1)
	{
		err = 200;			// param must be concentration
		m_conc_radio = 0;	// "push" concentration button
	}

	// if the user is not allowed to alter ANY parameters
	if (!alter)
		err = 203;


	if (err !=0)
	{
		CErrBox box;
		box.SetErr(err);
		box.DoModal();
		return false;		// an error occurred
	}

	return true;		// error did not occur
}


void CEnterDataDlg::OnEmpEntConcRadio() 
{
	UpdateData(true);

	if (!check(1))
	{
		m_conc_radio = -1;
		UpdateData(false);
		return;
	}

	// set name for param, write it to the screen.
	m_param_name = "Concentration";
	param = 1;
	UpdateData(false);
}

void CEnterDataDlg::OnEmpEntPresRadio() 
{
	UpdateData(true);

	if (!check(2))
	{
		m_pres_radio = -1;
		UpdateData(false);
		return;
	}

	// set name for param, write it to the screen.
	m_param_name = "Pressure (kPa)";
	param = 2;
	UpdateData(false);
}

void CEnterDataDlg::OnEmpEntVlosRadio() 
{
	UpdateData(true);

	if (!check(3))
	{
		m_vlos_radio = -1;
		UpdateData(false);
		return;
	}

	// set name for param, write it to the screen.
	m_param_name = "Velocity (m/s)";	
	param = 3;
	UpdateData(false);
}

void CEnterDataDlg::OnEmpEntTempRadio() 
{
	UpdateData(true);

	if (!check(4))
	{
		m_temp_radio = -1;
		UpdateData(false);
		return;
	}

	// set name for param, write it to the screen.
	m_param_name = "Temperature (C)";
	param = 4;
	UpdateData(false);
}

void CEnterDataDlg::OnEmpEntTimeRadio() 
{
	UpdateData(true);

	if (!check(5))
	{
		m_time_radio = -1;
		UpdateData(false);
		return;
	}

	// set name for param, write it to the screen.
	m_param_name = "Time(s)";
	param = 5;
	UpdateData(false);	
}



void CEnterDataDlg::OnEmpEntVolRadio() 
{
	UpdateData(true);

	if (alter)
	{
		mgl = false;	
		m_mgl_radio = -1;
		m_vol_radio = 0;
		UpdateData(false);		// write new values to screen
	}
}


void CEnterDataDlg::OnEmpEntMglRadio() 
{
	UpdateData(true);

	if (alter)
	{
		mgl = true;	
		m_mgl_radio = 0;
		m_vol_radio = -1;
		UpdateData(false);		// write new values to screen
	}
}


void CEnterDataDlg::OnOK() 
{
	// data for flux and param is turned into point objects
	//   with set_data from CEmpData, in the CUppmemDoc object
	
	CDialog::OnOK();
}


BOOL CEnterDataDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	VERIFY(conc_edit.SubclassDlgItem(IDC_EMP_ENT_CONC, this));
	VERIFY(pres_edit.SubclassDlgItem(IDC_EMP_ENT_PRES, this));
	VERIFY(vlos_edit.SubclassDlgItem(IDC_EMP_ENT_VLOS, this));
	VERIFY(temp_edit.SubclassDlgItem(IDC_EMP_ENT_TEMP, this));

	if (!alter)		// if user is not allowed to alter params,
	{				//   disable the above text boxes
		conc_edit.SetReadOnly();
		pres_edit.SetReadOnly();
		vlos_edit.SetReadOnly();
		temp_edit.SetReadOnly();
	}

	return true;
}
