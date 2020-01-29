C *****************************************************************************
C *                                                                           *
C *                   ERROR AND WARNING HANDLING SUBROUTINE                   *
C *                                                                           *
C *****************************************************************************

      SUBROUTINE ERROR (ERRMAT,ERRNUM,CODE)
      IMPLICIT NONE
      INTEGER*2 ERRNUM,ERRMAT(30),CODE 
      
      ERRNUM = ERRNUM + 1

      ERRMAT(ERRNUM) = CODE

      END

