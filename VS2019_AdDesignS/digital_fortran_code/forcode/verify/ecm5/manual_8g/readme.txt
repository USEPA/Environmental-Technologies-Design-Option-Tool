
NOTE ABOUT CPHSDM2 VERIFICATIONS.
=================================

The VB3 version of AdDesignS (pre-dating 9/3/98) includes a
link to a DLL that contains the ECM() subroutine.  The VB5
version of AdDesignS (as of 9/3/98) includes a link to
an EXE that contains the ECM() subroutine.  (Apparently
EXEs are more stable, but they are definitely easier to
debug.)

The files in this directory are:

    ECM_VB3.DAT   -   Test file in AdDesignS-VB3 format
    ECM_VB5.DAT   -   Test file in AdDesignS-VB5 format
    ECM_VB3.TXT   -   ECM results from AdDesignS-VB3
    ECM_VB5.TXT   -   ECM results from AdDesignS-VB5

The VB5 version produces better ECM results than the VB3 version.
This is due to a bug in the VB3 version where the order of
components was improperly displayed on-screen and in the
ECM results output file (no other data was impacted).

Note that all three components were used in each of
the ECM model runs.

Eric J. Oman
2:43 PM 9/3/98
