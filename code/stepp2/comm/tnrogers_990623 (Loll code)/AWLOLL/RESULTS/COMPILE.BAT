@echo off
del *.obj
del *.out
rem Fl /Od /Gt /MW lintest.for fit.for simp_new.for linpac.for /link corewin /SE:512
rem Fl /Od /MW /Gt fit.for simp_new.for linpac.for /link corewin /SE:4096
Fl /Od /MW /Gt fitwfgct.for>a.err simp_dp.for linpac.for /link corewin /SE:512

