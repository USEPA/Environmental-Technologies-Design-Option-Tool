!CC *****************************************************************************
!CC *                                                                           *
!CC *                   ERROR AND WARNING HANDLING SUBROUTINE                   *
!CC *                                                                           *
!CC *****************************************************************************

      SUBROUTINE ERROR (ERRMAT,ERRNUM,CODE)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ERROR
!MS$ ATTRIBUTES ALIAS:'_ERROR@12':: ERROR
!MS$ ATTRIBUTES REFERENCE::ERRMAT,ERRNUM,CODE

      DOUBLE PRECISION ERRNUM,ERRMAT 

      INTEGER CODE

      DIMENSION ERRMAT(30)

      ERRNUM = ERRNUM + 1

      ERRMAT(ERRNUM) = CODE

      END


