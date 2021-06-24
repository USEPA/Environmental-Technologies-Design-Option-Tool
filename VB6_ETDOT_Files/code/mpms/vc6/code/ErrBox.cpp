// ErrBox.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "ErrBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CErrBox dialog


CErrBox::CErrBox(CWnd* pParent /*=NULL*/)
	: CDialog(CErrBox::IDD, pParent)
{
	//{{AFX_DATA_INIT(CErrBox)
	m_ErrMsg = _T("");
	m_ErrMsgB = _T("");
	m_ErrMsgC = _T("");
	//}}AFX_DATA_INIT
}


void CErrBox::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CErrBox)
	DDX_Text(pDX, IDC_ERR_MSG, m_ErrMsg);
	DDV_MaxChars(pDX, m_ErrMsg, 100);
	DDX_Text(pDX, IDC_ERR_MSG2, m_ErrMsgB);
	DDV_MaxChars(pDX, m_ErrMsgB, 100);
	DDX_Text(pDX, IDC_ERR_MSG3, m_ErrMsgC);
	DDV_MaxChars(pDX, m_ErrMsgC, 100);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CErrBox, CDialog)
	//{{AFX_MSG_MAP(CErrBox)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CErrBox message handlers

// Sets the data output by the dialogue box, error or warning, and
//   where the error occurred.  A substitute for a non-default 
//   constructor, since the syntax would get complicated.  Declare 
//   the dialogue box, call this function, then invoke the dialogue 
//   box to view it.
void CErrBox::SetErr(int value)
{

// Allow up to 4 spaces for error reporting.  First is for
//   error or warning stuff, second is always there, third is
//   sometimes there, fourth is for out of bounds check.
//   Might as well output the error number too.  Do this later,
//   after you fix the ugly box graphics.

	switch (value)
	{
	// 	Standard errors related to the program
	case -1:	m_ErrMsg = "Negative Flux Determined!";
	break;

	case 0 :	m_ErrMsg = "Error in finding flux algorithm!";
	break;


///////////////////////////////////////////////////////////////
//  Errors from mechanistic flux determination: 
//		1x errors are memsys errors,
	case 10:	m_ErrMsg  = "Memsys: Memsys program did not load. Make";
				m_ErrMsgB = "certain the Memsys program is in the same";
				m_ErrMsgC = "directory as Uppmem.";
	break;

	case 11:	m_ErrMsg  = "Memsys: Could not load Memsys output file.";
				m_ErrMsgB = "Remember to save files using F2.";
	break;

	case 12:	m_ErrMsg  = "This command will launch the Memsys program.";
				m_ErrMsgB = "Remember to save output files with F2 if you";
				m_ErrMsgC = "wish to import them to Uppmem.";
	break;

//		2x errors are SE errors,
	case 20:	m_ErrMsg  = "SE Model: Solution did not converge.";
	break;

	case 21:	m_ErrMsg  = "SE Model: Flux is membrane controlled.";
	break;

	case 22:	m_ErrMsg  = "SE Model: Solution did not converge.";
				m_ErrMsgB = "          Flux is membrane controlled.";
	break;

	case 23:	m_ErrMsg  = "SE Model: Solution did not converge.";
				m_ErrMsgB = "          Recirculation rate is too high.";
	break;
							
	case 24:	m_ErrMsg  = "SE Model: Permeate is predicted to be greater than";
				m_ErrMsgB = "          influent.  Increase pressure or concentration,";
				m_ErrMsgC = "          decrease velocity or recirculation rate.";
	break;

//		3x errors are resistance errors,
	case 30:	m_ErrMsg = "Resistance Model: Model not completed yet.";
	break;

//		4x errors are gel polarization errors
	case 40:	m_ErrMsg  = "Gel Polarization: Solution didn't converge.";
				m_ErrMsgB = "Possible clogging of module - reduce concentration.";
	break;

	case 41:	m_ErrMsg  = "Gel Polarization: You must specify additional";
				m_ErrMsgB = "                  parameters for this model.";
	break;

	case 42:	m_ErrMsg  = "Gel Polarization: Solution didn't converge.";
				m_ErrMsgB = "Recirculation rate is too high for a steady";
				m_ErrMsgC = "state solution.  Reduce Recirculation rate.";
	break;

///////////////////////////////////////////////////////////////
//  Errors from empirical flux determination:


///////////////////////////////////////////////////////////////
//  Errors from sub-dialogue boxes:
//		General
	case 100:	m_ErrMsg = "Please select a flux model.";
	break;

//		Iteration over Parameter range: 11x

//		Particle distribution: 12x
	case 120:	m_ErrMsg  = "Replace current concentration for this";
				m_ErrMsgB = "particle size?";
	break;

	case 121:	m_ErrMsg  = "Remove this item from the list of particles?";
	break;

	case 122:	m_ErrMsg  = "Particle size is zero.  Please enter a positive";
				m_ErrMsgB = "value for particle size.";
	break;

//		Predefined membranes: 13x
	case 130:	m_ErrMsg  = "Another predefined membrane exists with the";
				m_ErrMsgB = "same name.  Press OK to replace the old";
				m_ErrMsgC = "membrane with the new parameters.";
	break;

	case 131:	m_ErrMsg  = "Warning: this will premanently delete this";
				m_ErrMsgB = "entry.  Press OK to delete permanently.";
	break;

//		Additional model parameters: 14x


///////////////////////////////////////////////////////////////
//	Empirical sub-dialogue boxes
//		Entering Empirical Data: 20x

	case 200:	m_ErrMsg  = "Experimental data must be of the form:";
				m_ErrMsgB = "Flux vs. Concentration";
				m_ErrMsgC = "to initially construct this model.";
	break;
	
	case 201:	m_ErrMsg  = "Experimental data must be of the form:";
				m_ErrMsgB = "Flux vs. Pressure";
				m_ErrMsgC = "to initially construct this model.";
	break;

	case 202:	m_ErrMsg  = "Experimental data must be of the form:";
				m_ErrMsgB = "Flux vs. Time";
				m_ErrMsgC = "to initially construct this model.";
	break;

	case 203:	m_ErrMsg  = "For this data model, initial data sets must";
				m_ErrMsgB = "have identical parameters.";
	break;

///////////////////////////////////////////////////////////////
//	Empirical Model errors: 3xx
	case 300:	m_ErrMsg  = "Data list is empty!  No prediction can";
				m_ErrMsgB = "be done with an empty data list.";
	break;

	case 301:	m_ErrMsg  = "One or more model parameters has a negative";
				m_ErrMsgB = "value.  Adjust parameters manually, or enter";
				m_ErrMsgC = "a new data set.";
	break;

	case 302:	m_ErrMsg  = "Negative Flux Determined!";
				m_ErrMsgB = "Check validity of model coefficients, and";
				m_ErrMsgC = "membrane parameters.";
	break;

	case 303:	m_ErrMsg  = "Error in data entry.  Data list is either empty,";
				m_ErrMsgB = "or the data columns are not of equal height.";
				m_ErrMsgC = "Please reenter data.";
	break;

//	Surface Renewal: 31x
	case 311:	m_ErrMsg  = "Could not find a solution for the data set";
				m_ErrMsgB = "given.  Add more data or try a new model.";
	break;

	case 312:	m_ErrMsg  = "Could not determine the value for surface";
				m_ErrMsgB = "renewal with the given data.";
	break;

	case 313:	m_ErrMsg  = "Insufficient data for prediction with Surface";
				m_ErrMsgB = "Renewal Model.  Please add more data.";
	break;

//	Fouling Resistance: 32x

//	Resistance: 33x

//	Gel Polarization: 34x
	case 340:	m_ErrMsg  = "Bulk concentration is greater than Gel concentration.";
				m_ErrMsgB = "Flux model predicts zero flux.";
	break;

//  Range warnings:	35x
	case 350:	m_ErrMsg  = "No data set exists for a valid prediction";
				m_ErrMsgB = "of flux for the given parameters.";
	break;

	case 351:	m_ErrMsg  = "This flux prediction is based on an estimation";
				m_ErrMsgB = "of flux for the given parameters.  Low reliability.";
	break;

	case 352:	m_ErrMsg  = "This flux is based on an extrapolation of";
				m_ErrMsgB = "the given data sets.  Medium reliability.";
	break;

	case 353:	m_ErrMsg  = "This flux prediction is based on an interpolation";
				m_ErrMsgB = "of given data sets.  High reliability.";
	break;

//	User-defined model coefficients errors
	case 360:	m_ErrMsgB = "Invalid model coefficients!";
	break;

///////////////////////////////////////////////////////////////
//  Default:
	default:	m_ErrMsg = "Unknown Error Message!  Update your software.";
	}	// end switch


	if (value >= 1000)
	{
		m_ErrMsg = "Flux value is either negative, or undefined.";
	}

}







