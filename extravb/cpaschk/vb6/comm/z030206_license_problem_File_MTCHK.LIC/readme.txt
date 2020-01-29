If the user at MTU experiences this problem when I perform the install:

"Unable to read licensing data.  You may need to reinstall the software."

The solution is simple:

Hokanson (or administrator) login to appropriate PC and locate the
following file in the windows folder (e.g., c:\winnt)

   MTCHK.LIC

Right click this file, choose Properties.
Click the "Security" tab.

Go down to the choice that says "Users"
Select "Full Control" so all the choices are checked in this box under Allow.
When done, the Permission for "Users" will read as follows:


Permissions            Allow                Deny

Full Control           Checked              Not Checked
Modify                 Checked              Not Checked
Read & Execute         Checked              Not Checked
Read                   Checked              Not Checked
Write                  Checked              Not Checked
