Attribute VB_Name = "Module1"
'Declare Sub SubIsomorph Lib "c:\msvc\mosdap16\mosdap.dll" (ByVal WordA As String, ByVal WordB As String, ByVal IntVar1 As Integer, ByRef IntVar2 As Integer, ByRef OneArray As Integer, ByRef TwoArray As Integer)
Declare Sub MOSDAP Lib "MOSDAP2.dll" (ByVal Query As String, _
ByVal QueryType As Byte, ByVal Subfile As String, ByVal Outfile As String, _
ByVal SearchType As Byte, ByRef SearchResult As Byte, ByRef SF_ID As Long, _
ByRef SF_Quant As Long, ByRef MF_ID As Long, ByRef MF_Quant As Long)
'The arrays intSF_ID[] and intSF_Quant[] are dimensioned to 100, and the arrays intMF_ID[] and intMF_Quant[] are dimensioned to 21.  To pass them to the .dll, you pass the first element in the array (e.e., intSF_ID[0] which is passed by reference).



'query string input of smile string
'Querytype 0 for string 1 for file
'Subfile is the name of substructure (ie unifac)
'Outfile is were to output file if querytype is type 1 also is delimiated ie tab
'Searchtype 0 none truncating, sequential 1 truncating sequencial
        '2 truncating combinatorial
'Searchresult flag pass fail or patrial falure
'sf_id array.. subfragment id .. multiple groups seperated by -1 intialized to 0
'sf_quant arry.. subfragment quantity .. multiple groups seperated by -1 intialized to 0
'mf_id arry.. molecular feture ... intialized to 0
'mf_quant arry.. molecular feture ... intialized to 0
