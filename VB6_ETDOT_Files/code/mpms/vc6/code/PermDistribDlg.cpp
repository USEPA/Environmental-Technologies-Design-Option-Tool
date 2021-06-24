// PermDistribDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "PermDistribDlg.h"
#include <math.h> 

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPermDistribDlg dialog


CPermDistribDlg::CPermDistribDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CPermDistribDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPermDistribDlg)
	//}}AFX_DATA_INIT
}


void CPermDistribDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPermDistribDlg)
	//}}AFX_DATA_MAP
}


void CPermDistribDlg::set_list(CMemCond mem)
{
	// read in values for retentate and permeate distribs,
	//   set them to the respective arrays in this object

	part_num = mem.NumDstb;

	for (int i=0; i<part_num; i++)
	{
		permeate[0][i] = (mem.EffDstb[0][i]) * 1000000;		// meters to microns
		permeate[1][i] = (mem.EffDstb[1][i]) * 4.1888e12 * mem.Dav * pow((mem.EffDstb[0][i]),3);	// #/ml to mg/L
		retentate[0][i] = (mem.RtnDstb[0][i]) * 1000000;	// meters to microns
		retentate[1][i] = (mem.RtnDstb[1][i]) * 4.1888e12 * mem.Dav * pow((mem.RtnDstb[0][i]),3);	// #/ml to mg/L
	}
}


// Turns a number into a doubly tab suffixed string
//   used in 'OnInitDialogue'
CString perm_num_stuff(double num)
{
	char buffer[50];
	_gcvt(num, 5, buffer);
	CString str = (CString)buffer;
	str += '\t';
	str += '\t';
	
	// add a third tab if string is too small
	if (str.GetLength() < 11)
		str += '\t';

	return str;
}


BOOL CPermDistribDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	// transfer control of the listboxes from the main 
	//   program to this dialogue box object
	VERIFY(m_permList.SubclassDlgItem(IDC_PERM_PART_LIST, this));
	VERIFY(m_retenList.SubclassDlgItem(IDC_RETEN_PART_LIST, this));

	CString perm_entry;		// permeate particle entry item
	CString reten_entry;	// retentate particle entry item

	// First item in the listboxes
	perm_entry = " Size (microns)\t\tConc(mg/l)";
	reten_entry = perm_entry;

	// place strings into list boxes
	m_permList.AddString((LPCSTR)perm_entry);
	m_retenList.AddString((LPCSTR)reten_entry);


	// add items in arrays to edit boxes
	for (int i=0; i<part_num; i++)
	{
		// Convert doubles to strings, concact entries
		perm_entry = perm_num_stuff((permeate[0][i]));
		perm_entry += perm_num_stuff((permeate[1][i]));

		reten_entry = perm_num_stuff((retentate[0][i]));
		reten_entry += perm_num_stuff((retentate[1][i]));

		// place strings into list boxes
		m_permList.AddString((LPCSTR)perm_entry);
		m_retenList.AddString((LPCSTR)reten_entry);
	}


	UpdateData(false);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


BEGIN_MESSAGE_MAP(CPermDistribDlg, CDialog)
	//{{AFX_MSG_MAP(CPermDistribDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPermDistribDlg message handlers
