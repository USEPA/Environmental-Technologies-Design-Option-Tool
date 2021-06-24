// PreDefMemDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include "PreDefMemDlg.h"
#include <afx.h>
#include "ErrBox.h"
#include <fstream.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPreDefMemDlg dialog


CPreDefMemDlg::CPreDefMemDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CPreDefMemDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPreDefMemDlg)
	m_area = 1.0;
	m_length = 1.0;
	m_membrane = _T("");
	m_maker = _T("Membranes-R-Us");
	m_name = _T("MRU10");
	m_prad = 0.05;
	m_resist = 9.98e+10;
	m_crad = 1;
	//}}AFX_DATA_INIT

	mem_num = 0;
}


void CPreDefMemDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPreDefMemDlg)
	DDX_Text(pDX, IDC_MEMB_NAME, m_name);
	DDV_MaxChars(pDX, m_name, 100);
	DDX_Text(pDX, IDC_MEMB_RESISTANCE, m_resist);
	DDV_MinMaxDouble(pDX, m_resist, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MEMB_CHANNEL_RADIUS, m_crad);
	DDV_MinMaxDouble(pDX, m_crad, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MEMB_LENGTH, m_length);
	DDV_MinMaxDouble(pDX, m_length, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MEMB_PORE_RADIUS, m_prad);
	DDV_MinMaxDouble(pDX, m_prad, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MEMB_AREA, m_area);
	DDV_MinMaxDouble(pDX, m_area, 1.e-300, 1.e+300);
	DDX_Text(pDX, IDC_MEMB_MANFC, m_maker);
	DDV_MaxChars(pDX, m_maker, 100);
	DDX_LBString(pDX, IDC_MEMB_LIST, m_membrane);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPreDefMemDlg, CDialog)
	//{{AFX_MSG_MAP(CPreDefMemDlg)
	ON_BN_CLICKED(IDC_MEMB_ENTER, OnMembEnter)
	ON_BN_CLICKED(IDC_MEMB_REMOVE, OnMembRemove)
	ON_BN_CLICKED(IDC_MEMB_VIEW, OnMembView)
	ON_LBN_DBLCLK(IDC_MEMB_LIST, OnDblclkMembList)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Non-member functions needed by member functions
/////////////////////////////////////////////////////////////////////////////

// Turns a number into a tab suffixed string
CString stuff_num(double num)
{
	char buffer[50];
	CString str;
	_gcvt(num, 3, buffer);
	str  = (CString)buffer;
	str += "\t";
	return str;
}


// parse the passed item for the tab character.  Replace
//   '\t' with ' ' so tab won't be counted twice
CString parse_membrane(CString &mem, int &front)
{
	int back = mem.Find('\t');				// first instance of a tab
	CString item = mem.Mid(front, back);	// extract a chunk of the string
	front = back + 1;						// front of next item 
	mem.SetAt(back, ' ');					// replace tab with a space, so tab
											//   wont be 'found' twice 
	return item;
}					




/////////////////////////////////////////////////////////////////////////////
// CPreDefMemDlg message handlers

BOOL CPreDefMemDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// transfer control of the listbox (IDC_MEMB_LIST)
	//   from the main program to this dialogue box object
	VERIFY(m_membList.SubclassDlgItem(IDC_MEMB_LIST, this));


	// set the tab stops to varying lengths
	int stops[6] = {60, 108, 128, 148, 196, 216};
	LPINT tab_stops = &stops[0];
	m_membList.SetTabStops(6, tab_stops);


	// Read in all membranes strings from file custmem.txt, 
	//   and insert each into the list box item
	ifstream fin;	
	fin.open("custmem.txt", ios::in | ios::nocreate);

	// if the file does not exist, make a new one
	if (fin.fail())
	{
		fin.close();
		ifstream new_fin;
		new_fin.open("custmem.txt");
		new_fin.close();
	}

	else
	{
		char entry[200];
		CString membrane;
		fin >> mem_num;
		fin.getline(entry, 200);		// extract extra junk from first line
		for (int i=0; i<mem_num; i++)
		{
			// extract first line, stores all values except '\n'
			//   in the array 'entry'
			fin.getline(entry, 200);

			// convert to CString
			membrane = (CString)entry;

			// add 'membrane' to the list box
			m_membList.AddString((LPCSTR)membrane);
		}

		// close the file
		fin.close();
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}






void CPreDefMemDlg::OnMembEnter() 
{
	// Read in data from screen
	UpdateData(true);

	// Check if another membrane has the same name
	CString membrane, name;
	int index, i = 0;
	while(i<mem_num && name!=m_name)
	{
		index = 0;
		m_membList.GetText(i, membrane);
		name = parse_membrane(membrane, index);
		i++;
	}

	// If same membrane exists, show error message
	if (name == m_name)	
	{
		CErrBox box;
		box.SetErr(130);
		if (box.DoModal() == IDOK)
		{
			// remove item with index i
			m_membList.DeleteString(i-1);
			// decrement the number of items
			mem_num--;
		}
		else
			return;		// do nothing if the user presses cancel
	}

	// Convert numbers to strings, concact them all together,
	//   use tabs to seperate entries
	CString p_entry;

	p_entry = m_name;
	p_entry += '\t';
	p_entry += stuff_num(m_resist);
	p_entry += stuff_num(m_crad);
	p_entry += stuff_num(m_length);
	p_entry += stuff_num(m_prad);
	p_entry += stuff_num(m_area);
	p_entry += m_maker;
	p_entry += '\t';

	// Place the string into the ListBox
	m_membList.AddString((LPCSTR)p_entry);

	// Increment mem_num
	mem_num++;
}



void CPreDefMemDlg::OnMembRemove() 
{
	int item = m_membList.GetCurSel();

	// if no item is selected, do nothing
	if (item == LB_ERR)
		return;

	// display dialogue box, making sure they want to remove
	CErrBox box;
	box.SetErr(131);
	if (box.DoModal() == IDOK)
	{
		m_membList.DeleteString(item);		// remove item with index 'item'
		mem_num--;		// decrement mem_num
	}

	// do nothing if the user presses cancel

}



// when the user presses the 'view' button
void CPreDefMemDlg::OnMembView() 
{
	// read in selected list item
	CString membrane;
	int index = 0;
	int item = m_membList.GetCurSel();

	// if no item is selected, do nothing
	if (item == LB_ERR)
		return;

	m_membList.GetText(item, membrane);

	// parse the data, convert numbers with atof
	m_name   = parse_membrane(membrane, index);
	m_resist = atof(((const char*)parse_membrane(membrane, index)));
	m_crad   = atof(((const char*)parse_membrane(membrane, index)));
	m_length = atof(((const char*)parse_membrane(membrane, index)));
	m_prad   = atof(((const char*)parse_membrane(membrane, index)));
	m_area   = atof(((const char*)parse_membrane(membrane, index)));
	m_maker  = parse_membrane(membrane, index);

	// display it
	UpdateData(false);
}


// if the user clicks 'OK', dump selected list item to the screen
void CPreDefMemDlg::OnOK() 
{
	OnMembView();

	// update the file custmem.txt, overwrite with everything
	//   in the current ListBox object
	ofstream fout;
	fout.open("custmem.txt", ios::out);

	// first line in file: number of items
	fout << mem_num << endl;

	// output each ListBox string on different lines
	CString membrane;
	for (int i=0; i<mem_num; i++)
	{
		m_membList.GetText(i, membrane);
		fout << membrane << endl;
	}

	fout.close();

	CDialog::OnOK();
}


// if the user double-clicks an item on the list, 'View' the item
void CPreDefMemDlg::OnDblclkMembList() 
{
	OnMembView();
}
