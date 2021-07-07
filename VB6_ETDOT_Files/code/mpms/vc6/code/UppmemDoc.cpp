// UppmemDoc.cpp : implementation of the CUppmemDoc class
//

#include "stdafx.h"
#include "Uppmem.h"
#include "UppmemDoc.h"
#include "MechModel.h"
#include "MemCond.h"
#include "EmpModel.h"
#include "EmpData.h"	// redundant, but here anyway
#include "ErrBox.h"

#include "EmpModelsDlg.h"
#include "MechModelsDlg.h"
#include "PlantDesignDlg.h"
#include "PlantRunDlg.h"
#include "EmpPickModelDlg.h"
#include "EnterDataDlg.h"
#include "PreDefMemDlg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CUppmemDoc

IMPLEMENT_DYNCREATE(CUppmemDoc, CDocument)

BEGIN_MESSAGE_MAP(CUppmemDoc, CDocument)
	//{{AFX_MSG_MAP(CUppmemDoc)
	ON_COMMAND(ID_MODEL_MECH, OnModelMech)
	ON_COMMAND(ID_MODEL_DESIGN, OnModelDesign)
	ON_COMMAND(ID_MODEL_EMP, OnModelEmp)
	ON_COMMAND(ID_MODEL_RUN, OnModelRun)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CUppmemDoc construction/destruction

CUppmemDoc::CUppmemDoc()
{
	// TODO: add one-time construction code here

}


CUppmemDoc::~CUppmemDoc()
{
}


BOOL CUppmemDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CUppmemDoc serialization

void CUppmemDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CUppmemDoc diagnostics

#ifdef _DEBUG
void CUppmemDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CUppmemDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG



/////////////////////////////////////////////////////////////////////////////
// CUppmemDoc commands

// Mechanistic Model dialogue box 
void CUppmemDoc::OnModelMech() 
{
	CMechModelsDlg mec_dlg;

	// run the mechanistic subroutine
	if (mec_dlg.DoModal() != IDOK)
		return;

	// save the data in mec_dlg for virtual plant?
}



void CUppmemDoc::OnModelEmp() 
{
	// variable declaration: the necessarry dialogue boxes
	CEmpPickModelDlg pick_dlg;
	CPreDefMemDlg    mem_dlg;
	CEnterDataDlg    data_dlg, data_dlg_B, data_dlg_C;
	CEmpModelsDlg    emp_dlg;
	int error;	// in case of an error, report it

	// first box, let the user choose a flux model
	if (pick_dlg.DoModal() != IDOK)
		return;

	int model_type = pick_dlg.ModelType;

	// update module in emp_dlg object, and data_dlg objects
	(emp_dlg.module).set_Emp_model((pick_dlg.ModelType));
	data_dlg.set_data((pick_dlg.ModelType), true);

	// update other 2 data dlg boxes, whether or not they are used
	data_dlg_B.set_data((pick_dlg.ModelType)+10, true);
	data_dlg_C.set_data((pick_dlg.ModelType)+20, true);



	// second box, let the user choose a membrane
	if (mem_dlg.DoModal() != IDOK)
		return;

	// update mem in module object, in the emp_dlg object 
	((emp_dlg.module).mem).SetMem((mem_dlg.m_prad), 
				(mem_dlg.m_resist), (mem_dlg.m_crad), 
				(mem_dlg.m_length), (mem_dlg.m_area), 0);



	// new loop: if the user set the data properly, exit the loop.
	//   if not, loop until they do it properly, or until they
	//   hit cancel.

	CEmpData data_a, data_b, data_c;
	bool set_flagA, set_flagB, set_flagC;
	set_flagA = false;
	set_flagB = set_flagC = true;

	// while at least one of them isn't set properly
	while (!set_flagA || !set_flagB || !set_flagC)
	{
		// third box, let the user enter primary data 
		if (data_dlg.DoModal() != IDOK)
			return;
	
		// if needed, make user enter more data sets
		switch (model_type)
		{
		case 1:		// surface renewal, enter 1 more data set
			// make the parameters equal
			data_dlg_B.set_param((data_dlg.m_conc), (data_dlg.m_pres), 
					(data_dlg.m_vlos), (data_dlg.m_temp), (data_dlg.mgl));
			if (data_dlg_B.DoModal() != IDOK)
				return;
			break;

		case 2:		// fouling resistance, enter 2 more data sets
			data_dlg_B.set_param((data_dlg.m_conc), (data_dlg.m_pres), 
					(data_dlg.m_vlos), (data_dlg.m_temp), (data_dlg.mgl));
			data_dlg_C.set_param((data_dlg.m_conc), (data_dlg.m_pres), 
					(data_dlg.m_vlos), (data_dlg.m_temp), (data_dlg.mgl));
			if (data_dlg_B.DoModal() != IDOK)
				return;
			if (data_dlg_C.DoModal() != IDOK)
				return;
			break;
		}


		// extract the new CEmpData objects from the data_dlg objects,
		//   depending on model type, you can extract 1, 2 or 3 
		set_flagA = data_a.set_data(data_dlg);
		if (model_type == 1 || model_type == 2)
			set_flagB = data_b.set_data(data_dlg_B);
		if (model_type == 2)
			set_flagC = data_c.set_data(data_dlg_C);

		// if one of them wasn't set, we must try again
		if (!set_flagA || !set_flagB || !set_flagC)
		{
			CErrBox box_dlg;
			box_dlg.SetErr(303);	// improper data entry
			box_dlg.DoModal();
		}

	}		// end while loop



	// insert them as new tail nodes in the CList of CEmpData 
	//   objects in the CEmpModel object 'module', in the 
	//   CEmpModelsDlg object 'emp_dlg'.  Whew!
	(emp_dlg.module).insert_tail(data_a);
	if (model_type == 1 || model_type == 2)
		(emp_dlg.module).insert_tail(data_b);
	if (model_type == 2)
		(emp_dlg.module).insert_tail(data_c);

	// update the other variables in the emp_dlg object
	emp_dlg.set_units((data_dlg.mgl), (data_dlg.m_conc), (data_dlg.m_pres),
					  (data_dlg.m_temp), (data_dlg.m_vlos));

	// do the curve fitting analyses for the data sets, 
	//   and set the model parameters
	error = (emp_dlg.module).build_model();

	if (error != 1)			// if an error occurs
	{
		CErrBox box_dlg;
		box_dlg.SetErr(error);	// show the error
		if (box_dlg.DoModal() != IDOK)
			return;			// and quit if they hit cancel
	}

	// fourth, let the user alter parameters, 
	//   and enter new data in emp_dlg box.
	if (emp_dlg.DoModal() != IDOK)
		return;

	// save data in emp_dlg for virtual plant?
}



void CUppmemDoc::OnModelDesign() 
{
	CPlantDesignDlg dlg;
	if (dlg.DoModal() == IDOK)
	{
		// what to do when OK is hit
	}	
}



void CUppmemDoc::OnModelRun() 
{
	CPlantRunDlg dlg;
	if (dlg.DoModal() == IDOK)
	{
		// what to do when OK is hit
	}	
}
