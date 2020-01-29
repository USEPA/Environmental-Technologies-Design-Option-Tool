	SUBROUTINE ODEQUATN(N,T,YO,YDOT,RUVG,RUVH,XKE,NPHOT)
      IMPLICIT NONE
C
      INTEGER MAXIRREV,MAXPHOT,MAXREV,MAXODE,MAXWVLEN,MAXCOMP,MAXNTANK
	INTEGER MAXNTARGET,MAXSTEPS,MAXEQUATN
      PARAMETER (MAXIRREV=100)
      PARAMETER (MAXPHOT=20)
      PARAMETER (MAXREV=20)
      PARAMETER (MAXODE=30)
      PARAMETER (MAXWVLEN=100)
      PARAMETER (MAXCOMP=50)
	PARAMETER (MAXNTARGET=10)
	PARAMETER (MAXNTANK=25)
	PARAMETER (MAXSTEPS=2000)
      PARAMETER (MAXEQUATN=MAXODE*MAXNTANK)
C
      INTEGER COMPA(MAXIRREV),COMPB(MAXIRREV),
     +          COMPC(MAXIRREV),COMPD(MAXIRREV),
     +          COMPE(MAXREV),COMPF(MAXREV),
     +          COMPG(MAXPHOT),COMPH(MAXPHOT)
C
	INTEGER IDREACT,NODE,NIRREV,NPHOT,NTANK
      INTEGER I,J,N,NT,NITANK
C
      DOUBLE PRECISION RTHMO(MAXEQUATN,MAXIRREV)      
      DOUBLE PRECISION CONCINI(MAXCOMP),
	+                 CONC(MAXEQUATN+MAXNTANK,0:MAXSTEPS)
	DOUBLE PRECISION RUVG(MAXCOMP,MAXNTANK),RUVH(MAXCOMP,MAXNTANK)
C
      DOUBLE PRECISION TAU,HYG(MAXNTANK)
C
      DOUBLE PRECISION YDOT(MAXEQUATN),YO(MAXEQUATN)
      DOUBLE PRECISION XK(MAXIRREV),XKE(MAXREV)
      DOUBLE PRECISION T
C
      COMMON /DATA1/ IDREACT,TAU,NTANK
      COMMON /DATA3/ CONCINI,CONC,HYG,XK,NODE,NIRREV,
     +               COMPA,COMPB,COMPC,COMPD,COMPE,COMPF
	COMMON /PHOTO2/ COMPG,COMPH
C
C-----formation of ordinary differential equations
C
C-----equations for tank 1
C
      DO I = 1,NODE
        YDOT(I)=0
        DO J = 1, NIRREV
          IF ((COMPA(J).GT.NODE).AND.(COMPB(J).GT.NODE)) THEN
            IF (((COMPA(J)-NODE).EQ.I).OR.((COMPB(J)-NODE).EQ.I)) THEN
              RTHMO(I,J)=-XK(J)
     +                *(XKE(COMPA(J)-NODE)/HYG(1))
     +                *(XKE(COMPB(J)-NODE)/HYG(1))
     +                *YO(COMPA(J)-NODE)*YO(COMPB(J)-NODE)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j)-Node,')*yo(',COMPB(j)-Node,')'
            ENDIF
          ELSEIF ((COMPA(J).GT.NODE).AND.(COMPB(J).LE.NODE)) THEN
            IF (((COMPA(J)-NODE).EQ.I).OR.(COMPB(J).EQ.I)) THEN
              RTHMO(I,J)=-XK(J)*(XKE(COMPA(J)-NODE)/HYG(1))
     +                *YO(COMPA(J)-NODE)*YO(COMPB(J))
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j)-Node,')*yo(',COMPB(j),')'
            ENDIF
          ELSEIF ((COMPA(J).LE.NODE).AND.(COMPB(J).GT.NODE)) THEN
            IF ((COMPA(J).EQ.I).OR.((COMPB(J)-NODE).EQ.I)) THEN
              RTHMO(I,J)=-XK(J)*(XKE(COMPB(J)-NODE)/HYG(1))
     +                *YO(COMPA(J))*YO(COMPB(J)-NODE)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j),')*yo(',COMPB(j)-Node,')'
            ENDIF  
          ELSEIF ((COMPA(J).EQ.I).OR.(COMPB(J).EQ.I)) THEN
            RTHMO(I,J)=-XK(J)*YO(COMPA(J))*YO(COMPB(J))
            YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j),')*yo(',COMPB(j),')'
          ENDIF
C
          IF ((COMPC(J).EQ.I).OR.(COMPD(J).EQ.I)) THEN
            IF ((COMPA(J).GT.NODE).AND.(COMPB(J).GT.NODE)) THEN
              RTHMO(I,J)=XK(J)
     +                *(XKE(COMPA(J)-NODE)/HYG(1))
     +                *(XKE(COMPB(J)-NODE)/HYG(1))
     +                *YO(COMPA(J)-NODE)*YO(COMPB(J)-NODE)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node,')*yo(',COMPB(j)-Node,')'
            ELSEIF ((COMPA(J).GT.NODE).AND.(COMPB(J).LE.NODE)) THEN
              RTHMO(I,J)=XK(J)*(XKE(COMPA(J)-NODE)/HYG(1))
     +                *YO(COMPA(J)-NODE)*YO(COMPB(J))
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node,')*yo(',COMPB(j),')'
            ELSEIF ((COMPA(J).LE.NODE).AND.(COMPB(J).GT.NODE)) THEN
              RTHMO(I,J)=XK(J)*(XKE(COMPB(J)-NODE)/HYG(1))
     +                *YO(COMPA(J))*YO(COMPB(J)-NODE)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j),')*yo(',COMPB(j)-Node,')'
            ELSE
              RTHMO(I,J)=XK(J)*YO(COMPA(J))*YO(COMPB(J))
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j),')*yo(',COMPB(j),')'
            ENDIF
          ELSEIF (((COMPC(J)-NODE).EQ.I).OR.((COMPD(J)-NODE).EQ.I)) THEN
            IF ((COMPA(J).GT.NODE).AND.(COMPB(J).GT.NODE)) THEN
              RTHMO(I,J)=XK(J)
     +                *(XKE(COMPA(J)-NODE)/HYG(1))
     +                *(XKE(COMPB(J)-NODE)/HYG(1))
     +                *YO(COMPA(J)-NODE)*YO(COMPB(J)-NODE)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node,')*yo(',COMPB(j)-Node,')'
            ELSEIF ((COMPA(J).GT.NODE).AND.(COMPB(J).LE.NODE)) THEN
              RTHMO(I,J)=XK(J)
     +                *(XKE(COMPA(J)-NODE)/HYG(1))
     +                *YO(COMPA(J)-NODE)*YO(COMPB(J))   
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node,')*yo(',COMPB(j),')'
            ELSEIF ((COMPA(J).LE.NODE).AND.(COMPB(J).GT.NODE)) THEN
              RTHMO(I,J)=XK(J)
     +                *(XKE(COMPB(J)-NODE)/HYG(1))
     +                *YO(COMPA(J))*YO(COMPB(J)-NODE)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j),')*yo(',COMPB(j)-Node,')'
            ELSE
              RTHMO(I,J)=XK(J)
     +                *YO(COMPA(J))*YO(COMPB(J))  
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j),')*yo(',COMPB(j),')'  
            ENDIF
          ENDIF     
        ENDDO
      ENDDO
C 
C-----end of equations for tank 1
C
C-----equations for tanks 2 to NTANK
C
	DO NT = 2, NTANK
	  NITANK = (NT-1)*NODE
        DO I = NITANK+1, NITANK+NODE
          YDOT(I)=0
          DO J = 1, NIRREV
            IF ((COMPA(J).GT.NODE).AND.(COMPB(J).GT.NODE)) THEN
              IF (((COMPA(J)-NODE).EQ.(I-NITANK)).OR.
	+		   ((COMPB(J)-NODE).EQ.(I-NITANK))) THEN
                RTHMO(I,J)=-XK(J)
     +                *(XKE(COMPA(J)-NODE)/HYG(NT))
     +                *(XKE(COMPB(J)-NODE)/HYG(NT))
     +                *YO(COMPA(J)-NODE+NITANK)*YO(COMPB(J)-NODE+NITANK)
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j)-Node+nitank,')*yo(',COMPB(j)-Node+nitank,')'
              ENDIF
            ELSEIF ((COMPA(J).GT.NODE).AND.(COMPB(J).LE.NODE)) THEN
              IF (((COMPA(J)-NODE).EQ.(I-NITANK)).OR.
	+		   (COMPB(J).EQ.(I-NITANK))) THEN
                RTHMO(I,J)=-XK(J)*(XKE(COMPA(J)-NODE)/HYG(NT))
     +                *YO(COMPA(J)-NODE+NITANK)*YO(COMPB(J)+NITANK)
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j)-Node+nitank,')*yo(',COMPB(j)+nitank,')'
              ENDIF
            ELSEIF ((COMPA(J).LE.NODE).AND.(COMPB(J).GT.NODE)) THEN
              IF ((COMPA(J).EQ.(I-NITANK)).OR.
	+           ((COMPB(J)-NODE).EQ.(I-NITANK))) THEN
                RTHMO(I,J)=-XK(J)*(XKE(COMPB(J)-NODE)/HYG(NT))
     +                *YO(COMPA(J)+NITANK)*YO(COMPB(J)-NODE+NITANK)
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j)+nitank,')*yo(',COMPB(j)-Node+nitank,')'
              ENDIF  
            ELSEIF ((COMPA(J).EQ.(I-NITANK)).OR.
     +             (COMPB(J).EQ.(I-NITANK))) THEN
              RTHMO(I,J)=-XK(J)*YO(COMPA(J)+NITANK)*YO(COMPB(J)+NITANK)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=-k(',j,')*yo(',COMPA(j)+nitank,')*yo(',COMPB(j)+nitank,')'
            ENDIF
C
          IF ((COMPC(J).EQ.(I-NITANK)).OR.
	+       (COMPD(J).EQ.(I-NITANK))) THEN
            IF ((COMPA(J).GT.NODE).AND.(COMPB(J).GT.NODE)) THEN
              RTHMO(I,J)=XK(J)
     +              *(XKE(COMPA(J)-NODE)/HYG(NT))
     +              *(XKE(COMPB(J)-NODE)/HYG(NT))
     +              *YO(COMPA(J)-NODE+NITANK)*YO(COMPB(J)-NODE+NITANK)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node+nitank,')*yo(',COMPB(j)-Node+nitank,')'
            ELSEIF ((COMPA(J).GT.NODE).AND.(COMPB(J).LE.NODE)) THEN
              RTHMO(I,J)=XK(J)*(XKE(COMPA(J)-NODE)/HYG(NT))
     +              *YO(COMPA(J)-NODE+NITANK)*YO(COMPB(J)+NITANK)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node+nitank,')*yo(',COMPB(j)+nitank,')'
            ELSEIF ((COMPA(J).LE.NODE).AND.(COMPB(J).GT.NODE)) THEN
              RTHMO(I,J)=XK(J)*(XKE(COMPB(J)-NODE)/HYG(NT))
     +              *YO(COMPA(J)+NITANK)*YO(COMPB(J)-NODE+NITANK)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)+nitank,')*yo(',COMPB(j)-Node+nitank,')'
            ELSE
              RTHMO(I,J)=XK(J)*YO(COMPA(J)+NITANK)*YO(COMPB(J)+NITANK)
              YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)+nitank,')*yo(',COMPB(j)+nitank,')'
            ENDIF
            ELSEIF (((COMPC(J)-NODE).EQ.(I-NITANK)).OR.
	+           ((COMPD(J)-NODE).EQ.(I-NITANK))) THEN
              IF ((COMPA(J).GT.NODE).AND.(COMPB(J).GT.NODE)) THEN
                RTHMO(I,J)=XK(J)
     +              *(XKE(COMPA(J)-NODE)/HYG(NT))
     +              *(XKE(COMPB(J)-NODE)/HYG(NT))
     +              *YO(COMPA(J)-NODE+NITANK)*YO(COMPB(J)-NODE+NITANK)
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node+nitank,')*yo(',COMPB(j)-Node+nitank,')'
              ELSEIF ((COMPA(J).GT.NODE).AND.(COMPB(J).LE.NODE)) THEN
                RTHMO(I,J)=XK(J)
     +              *(XKE(COMPA(J)-NODE)/HYG(NT))
     +              *YO(COMPA(J)-NODE+NITANK)*YO(COMPB(J)+NITANK)   
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)-Node+nitank,')*yo(',COMPB(j)+nitank,')'
              ELSEIF ((COMPA(J).LE.NODE).AND.(COMPB(J).GT.NODE)) THEN
                RTHMO(I,J)=XK(J)
     +              *(XKE(COMPB(J)-NODE)/HYG(NT))
     +              *YO(COMPA(J)+NITANK)*YO(COMPB(J)-NODE+NITANK)
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)+nitank,')*yo(',COMPB(j)-Node+nitank,')'
              ELSE
                RTHMO(I,J)=XK(J)
     +              *YO(COMPA(J)+NITANK)*YO(COMPB(J)+NITANK)  
                YDOT(I)=YDOT(I)+RTHMO(I,J)
c      write(*,*)I,'--r=k(',j,')*yo(',COMPA(j)+nitank,')*yo(',COMPB(j)+nitank,')'  
              ENDIF
            ENDIF     
          ENDDO
        ENDDO
	ENDDO
C
C-----end of equations for tanks 2 to NTANK
C
C-----add in the photolysis rates for the compounds undergoing photolysis
C
	DO NT = 1, NTANK
	  NITANK = (NT-1)*NODE
	  DO I = NITANK+1, NITANK+NODE
          DO J = 1, NPHOT
            IF (COMPG(J).EQ.(I-NITANK)) THEN
              YDOT(I)=YDOT(I)+RUVG(COMPG(J),NT)
            ENDIF
            IF (COMPH(J).EQ.(I-NITANK)) THEN
              YDOT(I)=YDOT(I)+RUVH(COMPH(J),NT)
            ENDIF
		ENDDO
        ENDDO
	ENDDO
C
C-----consider the reactor hydrodynamics
C
	DO I = 1, NODE
	  YDOT(I)=YDOT(I)+IDREACT*NTANK*(CONCINI(I)-YO(I))/TAU
	ENDDO
	DO NT = 2, NTANK 
	  NITANK = (NT-1)*NODE
	  DO I = NITANK+1, NITANK+NODE
          YDOT(I)=YDOT(I)+IDREACT*NTANK*(YO(I-NODE)-YO(I))/TAU
	  ENDDO
      ENDDO
C
C-----end of ordinary differential equations
C
      RETURN
      END
