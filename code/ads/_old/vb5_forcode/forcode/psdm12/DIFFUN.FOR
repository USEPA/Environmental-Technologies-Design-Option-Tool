C                                                                       
C **********************************************************************                                                                       
		  SUBROUTINE DIFFUN (N,T,Y0,YDOT)                       
C **********************************************************************                                                                       
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)


ccccccccccccccccccccccccccccc get rid of implicit double precision!
ccccccccccccccccccccccccccccc replace with implicit none
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
Cc---- Constants
C      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
C
Cc**** Change Hokanson 2/8/97
Cc      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=6,MAXPTS=400,MAXDE=750)
C      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
Cc**** End Change Hokanson 2/8/97

c---- Subroutine Parameters
      INTEGER*4 N
      DOUBLE PRECISION T
      DOUBLE PRECISION Y0(MAXDE),YDOT(MAXDE)

c---- Local variables
      DOUBLE PRECISION WW(MAXMC),AAU(MAXMC),BB(MAXNC,MAXMC),
     &                 Z(MXCOMP),Q0(MXCOMP),CBS(MXCOMP,MAXMC)
      DOUBLE PRECISION CPORE(MAXDE),XKTIME(MXCOMP),
     &                 RT(MXCOMP),FAC(MXCOMP)
      INTEGER I,II,III,IIII,J,JJ,K,KK,M

c---- Common block variables
      DOUBLE PRECISION DG(MXCOMP),ST(MXCOMP),EDS(MXCOMP),EDP(MXCOMP),
     &                 BR(MAXNC,MAXNC),D(MXCOMP)
      DOUBLE PRECISION YM(MXCOMP),XNI(MXCOMP),XN(MXCOMP),WR(MAXNC),
     &                 AZ(MAXMC,MAXMC)
      INTEGER*2 MC,NC,NCOMP,N1
      DOUBLE PRECISION DGT
      INTEGER*2 NIN
      DOUBLE PRECISION STD(MXCOMP),BEDS(MXCOMP,MAXNC,MAXNC),
     &                 BEDP(MXCOMP,MAXNC,MAXNC),DGI(MXCOMP)
      INTEGER*2 MND,ND,MD
      DOUBLE PRECISION TOR(MXCOMP),PART(MXCOMP),TCONV,TORTU(MXCOMP)
      DOUBLE PRECISION RK1(MXCOMP),RK2(MXCOMP),RK3(MXCOMP),RK4(MXCOMP),
     &                 XK(MXCOMP)
      DOUBLE PRECISION CBO(MXCOMP)

c---- Common blocks
      COMMON /BLOCKA/ DG,ST,EDS,EDP,BR,D
      COMMON /BLOCKB/ YM,XNI,XN,WR,AZ
      COMMON /BLOCKC/ MC,NC,NCOMP,N1,DGT,NIN
      COMMON /BLOCKE/ STD,BEDS,BEDP,DGI,MND,ND,MD
      COMMON /BLOCKF/ TOR,PART,TCONV,TORTU
      COMMON /BLOCKG/ RK1,RK2,RK3,RK4,XK
      COMMON /BLOCKJ/ CBO

c---- Debug variables
      INTEGER*2 DEBUGM
      DOUBLE PRECISION LAST_T
      COMMON /DEBUG/ LAST_T, DEBUGM
      DOUBLE PRECISION O_RT(10)

C                                                                       
C---- Determine liquid phase concentrations at each radial and
C.... axial position within adsorbent particle using Ideal
C.... Adsorbed Solution Theory
C 
      DO 2 I = 1,NCOMP
	 XKTIME(I) = XK(I)*(RK1(I)+RK2(I)*(T/TCONV) +
     &               RK3(I)*DEXP(RK4(I)*(T/TCONV)))
c      XKTIME(I) = 0.01D0*XK(I)*(RK1(I) - RK2(I)*(T/TCONV) +
c     $            RK3(I)*DEXP(-RK4(I)*(T/TCONV)))
      IF (XKTIME(I) .LE. (XK(I)/1.0D+03)) THEN
	  XKTIME(I) = XK(I)/1.0D+03
      ENDIF
 2    CONTINUE
      DO 3 I = 1,NCOMP
	IF (TOR(I) .LT. 1.0D0) THEN
	  IF ((T/TCONV) .GT. 1.008D5) THEN
	     TORTU(I) = TOR(I) + PART(I)*(T/TCONV)
	     RT(I) = TORTU(I)/1.0D0
	     FAC(I) = ((1.0D0/RT(I)) - D(I))/(1.0D0 - D(I))
	  ELSE
	     FAC(I) = 1.0D0
	  ENDIF
	ELSE
	  FAC(I) = 1.0D0
	ENDIF
 3    CONTINUE                                                                        

      II = 0                                                            
      JJ = 0                                                            
      DO 15 K = 1,MC                                                    
	 DO 8 M = 1,NC                                                  
	    QTE = 0.0D0                                                 
	    YT0 = 0.0D0                                                 
	    DO 5 I = 1,NCOMP                                            
	       II = II + 1                                              
	       Z(I) = YM(I)*Y0(II)                                      
	       QTE = QTE + Z(I)                                         
	       YT0 = YT0 + XNI(I)*Z(I)                                  
	       II = II + N1 - 1                                         
    5       CONTINUE                                                    
	    DO 6 I = 1,NCOMP                                            
	       JJ = JJ + 1                                              
	       IF ( QTE .LE. 0.0D0 .OR. YT0 .LE. 0.0D0 ) THEN           
		  CPORE(JJ) = 0.0D0                                     
	       ELSE                                                     
		  Z(I) = Z(I)/QTE                                       
		  Q0(I) = YT0*XN(I)/YM(I)                               
		  IF ( XNI(I)*LOG10(Q0(I)) .LT. -20.0D0 ) THEN          
		     CPORE(JJ) = 0.0D0                                  
		  ELSE                                                  
	       CPORE(JJ) =(Z(I)*Q0(I)**XNI(I))*(XK(I)/XKTIME(I))**XNI(I)                   
		  ENDIF                                                 
	       ENDIF                                                    
	       JJ = JJ + N1 - 1                                         
    6       CONTINUE                                                    
	    IF ( M .LT. NC - 1 ) THEN                                   
	       II = (K - 1)*ND + M                                      
	       JJ = (K - 1)*ND + M                                      
	    ELSE                                                        
	       II = (K - 1) + MND                                       
	       JJ = (K - 1) + MND                                       
	    ENDIF                                                       
    8    CONTINUE                                                       
	 II = ND*K                                                      
	 JJ = ND*K                                                      
   15 CONTINUE                                                          
      DO 60 I = 1,NCOMP                                                 
	 II = (I-1)*N1                                                  
	 III = II + MND                                                 
	 IIII = III + MD                                                

	 IF ( NIN .EQ. 0 ) THEN
	    CINFL = 1.0D0
	 ELSE

CCCCCc---Modified by ejo on 8/10/96
CCCCC            CINFL = CINF(I,T/TCONV) / CBO(I)
CCCCCc---Original code follows:
CCCCCc            CINFL = CINF(I,T)
CCCCCc---End of modification comments

	    CINFL = CINF(I,T)
	 ENDIF

	 DO 20 K = 2,MC
	 IF ( CPORE(III + K) .LE. 0.0D0 ) THEN                          
	    CBS(I,K) = STD(I)*Y0(IIII + K)                              
	 ELSE                                                           
	    CBS(I,K) = STD(I)*(Y0(IIII + K) - CPORE(III + K))           
	 ENDIF                                                          
   20    CONTINUE                                                       
	 DO 40 K = 1,MC                                                 
	    WW(K) = 0.0D0                                               
	    AAU(K) = 0.0D0                                              
	    KK = II + (K-1)*ND                                          
	    DO 30 J = 1,ND                                              
	      BB(J,K) = 0.0D0                                           
	       DO 25 M = 1,ND                                           
		  BB(J,K) = BB(J,K) + BEDS(I,J,M)*Y0(KK + M)            
     +                           + BEDP(I,J,M)*FAC(I)*CPORE(KK + M)            
   25          CONTINUE                                                 
	       BB(J,K) = BB(J,K) + BEDS(I,J,NC)*Y0(III + K)             
     +                        + BEDP(I,J,NC)*FAC(I)*CPORE(III + K)             
   30       CONTINUE                                                    
	    DO 35 J = 1,ND                                              
	       JJ = KK + J                                              
C                                                                       
C---- Intraparticle Phase Mass Balance (excluding boundary)
C                                                                       
	       YDOT(JJ) = BB(J,K)                                       

	       WW(K) = WW(K) + WR(J)*YDOT(JJ)                           
   35       CONTINUE                                                    
   40    CONTINUE                                                       
C                                                                       
C---- Liquid-Solid Boundary Layer Mass Balance at column entrance
C                                                                       
	 YDOT(III+1) = (STD(I)*DGI(I)*(CINFL - CPORE(III + 1))          
     +                 - WW(1)) / WR(NC)                                
C                                                                       
	 DO 55 K = 2,MC                                                 
C                                                                       
C---- Liquid-Solid Boundary Layer Mass Balance within column
C                                                                       
	    YDOT(III+K) = (CBS(I,K)*DGI(I) - WW(K)) / WR(NC)            
C                                                                       
	    DO 50 M = 2,MC                                              
	       AAU(K) = AAU(K) + AZ(K,M)*Y0(IIII+M)                     
   50       CONTINUE                                                    
C                                                                       
C---- Liquid Phase Mass Balance
C                                                                       
	    YDOT(IIII+K) = -DGT*(AZ(K,1)*CINFL + AAU(K))                
     +                     - 3.0D0*CBS(I,K)                             
C                                                                       
   55    CONTINUE                                                       
   60 CONTINUE                                                          

      IF (DEBUGM .EQ. 1) THEN
c       Ensure that each time is output only once.
c       Note: The DIFFUN() routine is called several times for
c       each instant in time, and this debug output includes only
c       a snapshot of the first call to DIFFUN() for each time.
c       If this is a problem, just comment out the LAST_T code.
	TEST1 = T/TCONV
	TEST2 = LAST_T
	IF ( TEST1 .EQ. 0D0 ) THEN
	  GOTO 62
	END IF
	IF ( DABS((TEST1-TEST2)/TEST1) .LE. 1.0D-5 ) THEN
c         Do nothing--same time as the last output line.
	ELSE
C          WRITE(8,*) ((T/TCONV)/1.440D3),
C     &               (T/TCONV),
C     &               (O_RT(I), I=1, NCOMP)
c     &               CINFL,
c     &               CINFL * CBO(1)
c     &               (XKTIME(I), I=1, NCOMP),
c     &               (D(I), I=1, NCOMP),
c     &               (FAC(I), I=1, NCOMP),
c     &               (Y0(I), I=1, (((NC+1)*MC)-1)*NCOMP)
c     &               (YDOT(I), I=1, (((NC+1)*MC)-1)*NCOMP),
c     &               (CPORE(I), I=1, (((NC+1)*MC)-1)*NCOMP)
	  WRITE(8,*) ' '
	  LAST_T = T/TCONV
	END IF
      END IF

62    RETURN
      END                                                               

