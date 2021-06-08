/////////////////////////////////////////////////////////////////////////////
//	This is a way to convert numbers to strings, and to stuff
//    them with tab characters if necessarry.

#include <math.h> 
#include <stdlib.h>

// Turns a number into a tab prefixed string
CString stuff_num(double num, char buffer[], bool tab)
{
	CString str;
	_gcvt(num, 2, buffer);

	if(tab)
	{
		str = '\t';
		str += (CString)buffer;
	}

	else
		str = (CString)buffer;

	return str;
}


// parse the passed item for the next 5 spaces,
// stuff a '5' into 5 spaces so it won't be counted twice
CString parse_membrane(CString &mem, int &front)
{
	int back = mem.Find('\t');				// first instance of a tab
	CString item = mem.Mid(front, back);	// extract a chunk of the string
	front = back + 1;						// front of next item 
	mem.SetAt(back, ' ');					// replace tab with a space, so tab
											//   wont be 'found' twice 
	return item;
}