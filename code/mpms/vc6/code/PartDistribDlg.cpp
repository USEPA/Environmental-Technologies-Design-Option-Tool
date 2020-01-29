// PartDistribDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Uppmem.h"
#include <math.h> 
#include <stdlib.h>
#include "PartDistribDlg.h"
#include "ErrBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPartDistribDlg dialog


CPartDistribDlg::CPartDistribDlg(CWnd* pParent /*=NULL*/): 
				 CDialog(CPartDistribDlg::IDD, pParent)
{

	//{{AFX_DATA_INIT(CPartDistribDlg)
	m_massConc = 0.0;
	m_partRad = 0.0;
	m_density = 0.0;
	//}}AFX_DATA_INIT
}



void CPartDistribDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPartDistribDlg)
	DDX_Text(pDX, IDC_DISTRIB_CONC_MASS, m_massConc);
	DDV_MinMaxDouble(pDX, m_massConc, 0., 1000000.);
	DDX_Text(pDX, IDC_DISTRIB_SIZE, m_partRad);
	DDV_MinMaxDouble(pDX, m_partRad, 0., 1000000.);
	//}}AFX_DATA_MAP
}



BEGIN_MESSAGE_MAP(CPartDistribDlg, CDialog)
	//{{AFX_MSG_MAP(CPartDistribDlg)
	ON_BN_CLICKED(IDC_DISTRIB_VIEW, OnDistribView)
	ON_BN_CLICKED(IDC_DISTRIB_REMOVE, OnDistribRemove)
	ON_BN_CLICKED(IDC_DISTRIB_ENTER, OnDistribEnter)
	ON_LBN_DBLCLK(IDC_PARTICLE_LIST, OnDblclkParticleList)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()



/////////////////////////////////////////////////////////////////////////////
// CPartDistribDlg member functions

// accepts the current particle list from MechModelDlg, and sets
//   it equal to m_particles
void CPartDistribDlg::SetList(double particles[2][20], int num, double den)
{					
	for (int i = 0; i < num; i++)
	{
		m_particles[0][i] = particles[0][i];
		m_particles[1][i] = particles[1][i];
	}

	m_partNum = num;
	m_density = den;
}



/////////////////////////////////////////////////////////////////////////////
//  Additional functions for CPartDistribDlg functions

// Turns a number into a doublely tab suffixed string
CString num_stuff(double num)
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



// parse the passed item for the next tab character.
//   Replace '\t' with ' ' so it won't be counted twice
double parse_line(CString &line, int &front)
{
	// first time this is called, front should be zero
	int back = line.Find('\t');				// first instance of a tab
	CString item = line.Mid(front, back);	// extract a chunk of the string
	front = back + 1;						// front of next item 
	line.SetAt(back, ' ');					// replace tabs with spaces, so tabs
	line.SetAt((back+1), ' ');				//   wont be 'found' twice 

	// convert item to a double, and return
	double num = atof((const char*)item);
	return num;
}			


/////////////////////////////////////////////////////////////////////////////
// CPartDistribDlg message handlers

// Sets up the ListBox just prior to displaying it on the screen
BOOL CPartDistribDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	conc = 0;

	// transfer control of the listbox (IDC_PARTICLE_LIST)
	//   from the main program to this dialogue box object
	VERIFY(m_partList.SubclassDlgItem(IDC_PARTICLE_LIST, this));

	// local variables, conc is always stored in #/ml
	double rad, num_conc, vol_conc, mgl_conc;
	CString p_entry;

	// enter each particle and its size into the list box
	for (int i=0; i<m_partNum; i++)
	{
		rad = 1000000 * m_particles[0][i];		// meters to microns
		num_conc = m_particles[1][i];			// #/ml
		vol_conc = num_conc * 4.1888 * (pow(rad, 3)) * 1e-12 * 100; // convert #/ml to %
		mgl_conc = vol_conc / 100 * m_density * 1000000; // convert % to mg/l

		// Convert doubles to strings, concact p_entry
		p_entry  = num_stuff(rad);
		p_entry += num_stuff(mgl_conc);
		p_entry += '\t';
		p_entry += '\t';

		// place string into list box
		m_partList.AddString((LPCSTR)p_entry);
	}

	// place last item in the box for editing putposes
	m_partRad  = rad;
	m_massConc = mgl_conc;

	UpdateData(false);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}



void CPartDistribDlg::OnDistribView() 
{
	// read in selected list item
	CString particle;
	int index = 0;
	int item = m_partList.GetCurSel();

	// if no item is selected, do nothing
	if (item == LB_ERR)
		return;

	m_partList.GetText(item, particle);

	// parse the data, convert numbers with atof
	m_partRad  = parse_line(particle, index);
	m_massConc = parse_line(particle, index);

	// display it
	UpdateData(false);	
}



void CPartDistribDlg::OnDistribRemove() 
{
	int item = m_partList.GetCurSel();

	// if no item is selected, do nothing
	if (item == LB_ERR)
		return;

	// display dialogue box, making sure they want to remove
	CErrBox box;
	box.SetErr(121);
	if (box.DoModal() == IDOK)
		m_partList.DeleteString(item);		// remove item with index 'item'
	else
		return;		// do nothing if the user presses cancel	

	// decrement the number of particles in the list
	m_partNum--;

	// no need to erase values in the matrix particles[][]
}



void CPartDistribDlg::OnDistribEnter() 
{
	// read in the new values
	UpdateData(true);

	// Check if another list item has the same particle size
	CString particle;
	double size;
	int index, i = 0;
	while(i<m_partNum && size!=m_partRad)
	{
		index = 0;
		m_partList.GetText(i, particle);
		size = parse_line(particle, index);
		i++;

		// If same particle exists, show error message
		if (size == m_partRad)	
		{
			CErrBox box;
			box.SetErr(120);
			if (box.DoModal() == IDOK)
			{
				// remove item with index i
				m_partList.DeleteString(i-1);
				// decrement the number of items
				m_partNum--;
				// enter new particle data below
			}

			else
				return;		// do nothing if the user presses cancel
		}
	}

	// second, check to see if particle size is zero
	if (m_partRad == 0)
	{
		CErrBox box;
		box.SetErr(122);
		box.DoModal();
		return;				// allow user to enter a new value
	}

	// Convert numbers to strings, concact them all together,
	//   use tabs to seperate entries
	CString p_entry;
	p_entry  = num_stuff(m_partRad);
	p_entry += num_stuff(m_massConc);
	p_entry += '\t';
	p_entry += '\t';

	// add item to the list box
	m_partList.AddString((LPCSTR)p_entry);

	// add item to the particle array
	double num_conc;
	m_particles[0][m_partNum] = m_partRad / 1000000;	// microns to meters

	// convert mg/L to #/ml
	num_conc = m_massConc / (4.1888 * pow(m_partRad, 3) * m_density * 1e-6);
	m_particles[1][m_partNum] = num_conc;				// #/ml

	// increment the number of particles
	m_partNum++;
}





/////////////////////////////////////////////////////////////////////////////
//	Other Controls
/////////////////////////////////////////////////////////////////////////////

void CPartDistribDlg::OnDblclkParticleList() 
{
	// if the user double-clicks on a list box item,
	//   the screen will be updated
	OnDistribView();
}

void CPartDistribDlg::OnOK() 
{
	// determine average particle size and conc 

	// subscripts: r=radius, v=vol, c=conc
	double r, v, c, sum_rvc, sum_vc, sum_c, sum_mass_c;
	sum_rvc = sum_vc = sum_c = sum_mass_c = 0;

	for (int i=0; i<m_partNum; i++)
	{
		r = m_particles[0][i];		// radius in m
		v = 4.1888 * pow (r, 3);	// volume in m^3
		c = m_particles[1][i];		// conc in #/ml
		sum_rvc += r*v*c;
		sum_vc  += v*c;
		sum_c   += c;
		sum_mass_c += c * (4.1888 * pow(r, 3) * m_density * 1e12);
	}

	// average radius in microns
	ave_rad = 1000000 * sum_rvc / sum_vc;

	// 'average' concentration is just the sum of the
	//   concentrations.  In mg/L
	ave_conc = sum_mass_c; 


	CDialog::OnOK();
}
