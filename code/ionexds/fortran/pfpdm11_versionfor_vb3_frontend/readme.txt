PFPDM11

Fixes problem with Number of Axial Elements described below
under PFPDM08.  The problem actually was not related to the
axial elements, but was a problem in DIFFUN for the calculation
of CTAVG used in the CPORE calculation for the case of 
variable influent.  This error was fixed with a single line
of code and corrected the problems of different results for
increased number of axial elements.  Now the program produces
identical results as expected when the only change is to
number of axial elements.

A previous change was an incorrect attempt to fix this problem.
This previous change involved breaking out of the program when
all ions reached within 1% of their influent conc.  This 
incorrect code modification from the past was eliminated in
PFPDM11. 

D. Hokanson
01/12/01

PFPDM10

Prints out some additional output files as described in
comment block at beginning of PFPDM10.for.

D. Hokanson
8/00

PFPDM09

Modifies PFPDM08 to accept EPS and DH0 as input parameters.
D. Hokanson
8/11/00


PFPDM08

Note:  There is still a problem with this version.  Running
       the model with 1 Tank produces a different result than
       running the program with 4 tanks in series (i.e. the
       spreading of the curves changes).  The results should
       be exactly the same so this problem needs to be
       investigated in the near future.

D. Hokanson
9/11/95

The test for exiting the program was changed in this
version.  The test is now that each component reach
1 % of its influent conc. CBO without regard to 
variable influent data.

D. Hokanson
9/6/95 

The capability to feed a presaturant condition where
more than one ion presaturates the adsorbent was added
to this version.

D. Hokanson
5/10/95



