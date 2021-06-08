Attribute VB_Name = "DB_Access_Mod"
'
' declare the stuff to let us access the database
'

Global db As database       'Paradox database
''''Global tb As Table          'Paradox database

'
' this will store the record number of the first
'   entry of each unique chemical in the database
'
Global db_index() As Long

' the number of unique chemicals in the database
Global db_num_entries As Integer

Global cbulk_array(15)
Global omag_array(15)
Global nl_array(15)
Global ns_array(15)
Global nv_array(15)
Global rg_array(15)
Global mx_array(15)

''''Global Selection As dynaset
Global Selection As Recordset

'
'DATABASE PASSWORD LISTED BELOW.  AS A DIFFERENCE FROM
'THE WAY StEPP WORKED UNDER MS VB 3.0,
'NO USER NAME IS REQUIRED, ONLY A PASSWORD.
'
'NOTE: FOR SOME #$(*&@#$*# ANNOYING REASON, ONCE A
'PASSWORD IS SUPPLIED FOR A DATABASE TO BE ACCESSED
'FROM VISUAL BASIC 5 (IN THE WAY THIS PROGRAM ACCESSES
'THE DATABASES), IT CANNOT BE DIRECTLY ACCESS FROM
'MICROSOFT ACCESS USING THE SAME PASSWORD IF THE PASSWORD
'WAS 14 CHARACTERS LONG (OR MAYBE JUST THIS EXAMPLE).
'USING THE NEW 9 CHARACTER PASSWORD WORKS EQUALLY
'WELL IN VB 5.0 AND MS ACCESS 8.0.
'              Eric J. Oman
'              10/23/98
'
''Global Const User_Name = "victor t. hart"
''Global Const User_Password = "frieda4wisc836"
'Global Const Encrypted_User_Name = "qJ8k\e%kO%G2ek"
'Global Const Encrypted_User_Password = "AeJ>;2atJh8m^g"
'
'Global Const User_Password = "frieda836"
Global Const Encrypted_User_Password = "AeJ>;2m^g"







