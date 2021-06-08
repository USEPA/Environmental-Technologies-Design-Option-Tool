CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE VBOUTPUT
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

      OPEN (UNIT=20, FILE='voc.out')

C PERCENT REMOVAL BY STRIPPING
      WRITE (20,3001) PEPSTP                   
C PERCENT REMOVAL BY VOLATILIZATION
      WRITE (20,3001) PEPVOL
C PERCENT REMOVAL BY SOLID WASTAGE
      WRITE (20,3001) PERSP            
C PERCENT REMOVAL BY LIQUID WASTAGE
      WRITE (20,3001) PERSPW            
C PERCENT REMOVAL BY BIODEGRADATION
      WRITE (20,3001) PERBIO           
C TOTAL PERCENT REMOVED (by concentrations)
      WRITE (20,3001) TPER
C TOTAL PERCENT REMOVED (by removals)
      WRITE (20,3001) TPERX

C TOTAL AMOUNT STRIPPED (kg/day)
      WRITE (20,3001) TRSTRP
C TOTAL AMOUNT VOLATILIZED (kg/day)
      WRITE (20,3001) TRSVOL
C TOTAL AMOUNT WASTED AS SOLID (kg/day)
      WRITE (20,3001) TRORPS
C TOTAL AMOUNT WASTED AS LIQUID (kg/day)
      WRITE (20,3001) TRORPW
C TOTAL AMOUNT BIODEGRADED (kg/day)
      WRITE (20,3001) TRBIOS
C TOTAL AMOUNT IN INFLUENT (kg/day)
      WRITE (20,3001) TIN
C TOTAL AMOUNT IN EFFLUENT (kg/day)
      WRITE (20,3001) TEFF


C Influent Liquid Concentrations (ug/L)
      WRITE (20,3001) CO1

C ****************** INFLUENT WEIR

      IF (W1.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CA
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RVOLW1
          WRITE (20,3001) PERW1
      END IF

C ****************** AERATED GRIT CHAMBER

      IF (GC.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CB
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RSTRPG
          WRITE (20,3001) PERAGC
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) RVOLAG
          WRITE (20,3001) PRVLGC
      END IF

C ****************** PRIMARY CLARIFIER

C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CC
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) RVOLPRI
          WRITE (20,3001) PERVLP
C         kg/day REMOVED BY SOLID WASTE | % OF TOTAL
          WRITE (20,3001) RSORPI
          WRITE (20,3001) PERSPA
C         kg/day REMOVED BY LIQ. WASTE | % OF TOTAL
          WRITE (20,3001) RSORPW
          WRITE (20,3001) PERSPC

C ****************** PRIMARY CLARIFIER WEIR

      IF (W2.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CD
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RVOLW2
          WRITE (20,3001) PERW2
      END IF

C ****************** AERATION BASIN

C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CE2(NTK)
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) TRSTAB 
          WRITE (20,3001) PESTAB
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) TRVLSA
          WRITE (20,3001) PERVLA
C         kg/day REMOVED BY BIODEG | % OF TOTAL
          WRITE (20,3001) TRBIOS
          WRITE (20,3001) PERBIO

C ****************** SECONDARY CLARIFIER

C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CE2(NTKP1)
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) TRVLSB
          WRITE (20,3001) PERVLB
C         kg/day REMOVED BY SOLID WASTE | % OF TOTAL
          WRITE (20,3001) RSOPSC
          WRITE (20,3001) PERSPB
C         kg/day REMOVED BY LIQ. WASTE | % OF TOTAL
          WRITE (20,3001) RSOPSW
          WRITE (20,3001) PERSPD

C ****************** SECONDARY CLARIFIER WEIR

      IF (W3.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) TCE2
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RVOLW3
          WRITE (20,3001) PERW3
      END IF

 	CLOSE(UNIT=20)
      
 3001 FORMAT(' ',1P, E20.12)

        RETURN
        END



