C
C        ********EQUILIBRIUM THEORY PROGRAM********
C
C    THIS PROGRAM CALCULATES MULTICOMPONENT BREAKTHROUGH 
C      FOR FIXED BED ADSORBERS. THE PROGRAM ASSUMES 
C       NO MASS TRANSFER RESISTANCE.IDEAL ADSORBED 
C     SOLUTION THEORY IS USED TO PREDICT COMPETITION.
C
C    PROGRAM WRITTEN BY:  THOMAS FRANCIS SPETH, MICHIGAN TECH. UNIV.
C    PROGRAM REVISED BY:  BETH ERLANSON, UNIV. OF TEXAS AT AUSTIN
C
C    NOTES:
C      BETWEEN ECM_MOD2.FOR AND ECM_MOD3.FOR, THE FOLLOWING CHANGES:
C        - CO(X,X)  ===>  COK(X)
C        - BVF      ===>  VOID
C
C
C                  VARIABLE DEFINITIONS
C
C  BVF   = BED VOLUMES FED
C  C     = LIQUID PHASE CONCENTRATION (ug/L)
C  CH    = WORKING CHARACTER
C  CHAR  = NAME OF THE COMPONENTS (TEN LETTERS)
C  COK   = INITIAL CONCENTRATIONS (ug/L)
C  DEN   = BULK DENSITY OF ADSORBENT (g/cm**3)
C  DGX   = DIMENSIONLESS GROUP X:  USED TO FIND STRONGEST COMPONENT
C  DGY   = DIMENSIONLESS GROUP Y:  USED TO FIND STRONGEST COMPONENT
C  FCN   = SUBROUTINE THAT SETS UP THE NON-LINEAR EQUATIONS
C  FCS   = C/COK
C  FLRT  = FLOW RATE (GPM/FT**2)
C  FNORM = OUTPUT:  SUM OF THE RESIDUALS
C  I     = COUNTER
C  IAST  = SUBROUTINE TO ACCOUNT FOR COMPETITIVE EFFECTS
C  IER   = OUTPUT:  ERROR PARAMETER
C  ITMAX = MAXIMUM NUMBER OF ITERATIONS
C  IX    = USED TO KEEP TRACK OF STRONGEST COMPONENT
C  J     = COUNTER
C  K     = COUNTER
C  L     = COUNTER FOR ERROR FIXING
C  M     = COUNTER
C  MW    = MOLECULAR WEIGHT
C  N     = NUMBER OF COMPONENTS TOTAL
C  NN    = NUMBER OF COMPONENTS IN A ZONE
C  NS    = NSIG INPUT
C  NSIG  = NUMBER OF DIGITS OF ACCURACY DESIRED IN THE COMPUTED ROOT
C  PAR   = PARAMETER SET
C  PAR(1 to N)= FREUNDLICH K VALUES
C  PAR(10 to 10+N)= FRUENDLICH N VALUES
C  PAR(20 to 20+N)= INITIAL CONCENTRATIONS
C  PAR(30)= VELOCITY OF THE WAVE: VW (cm/s)
C  PAR(35)= VELOCITY OF FLOW: VF (cm/s)
C  PAR(40 to 40+I)= CALCULATED LIQUID CONCENTRATIONS
C  PAR(60 to 60+I)= Q's OF THE PREVIOUS WAVE
C  PAR(80 to 80+I)= C's OF THE PREVIOUS WAVE
C  Q     = SOLID PHASE CONCENTRATION (ug/g)
C  QAVE  = AVERAGE Q IN ZONE
C  SSTC  = SINGLE SOLUTE TREATMENT CAPACITY (mg C/L WATER)
C  SUM   = USED TO CALCULATE VW AND BVF
C  VF    = VELOCITY OF FLOW (cm/s)
C  VOID   = BED VOID FRACTION
C  VW    = VELOCITY OF WAVE (cm/s)
C  WK    = WORK VECTOR:  LENGTH=N*(3*N+15)/2
C  X     = ONE DIMENSIONAL SOLID-PHASE CONCENTRATION (um/g)
C  XGUESS= INITIAL ESTIMATE OF ROOT (EMPTY)
C  XK    = FREUNDLICH K's (um/g)*((L/um)**1/N
C  XN    = FRUENDLICH 1/n 's
C  ZSQ   = COMMON BLOCK
C  ZZ    = VARIABLE USED TO CALCULATE INITIAL Q's
C  ZZZ   = DIMENSIONLESS BED LENGTH
C
C
C                        SAMPLE INPUT
C
C           N,VOID,DEN,FLRT
C           CHAR(I),XK(I),XN(I),COK(I),MW(I)
C           ....    ..    ..    ..    ..
C           CHAR(J),XK(J),XN(J),COK(J),MW(J)
C           NS
C
C                         DIMENSIONS
C
      subroutine ecm(nx, void_i, den_i, flrt_i, index_i,
     &               xk_i, xn_i, c0_io, xmw_i, nflagb,
     &               c_o, dgy_o, fcs_o, oats_o, q_o,
     &               qave_o, sstc_o, vw_o, zzz_o)

c-----// General stuff.
      implicit none
      integer nmax
      parameter (nmax = 20)

c-----// Parameters passed to the DLL.
      INTEGER*2 NX
      DOUBLE PRECISION VOID_I, DEN_I, FLRT_I
      INTEGER*2 INDEX_I(1:NMAX)
      DOUBLE PRECISION XK_I(1:NMAX)
      DOUBLE PRECISION XN_I(1:NMAX)
      DOUBLE PRECISION C0_IO(1:NMAX)
      DOUBLE PRECISION XMW_I(1:NMAX)
      INTEGER*2 NFLAGB
      DOUBLE PRECISION C_O(1:NMAX,1:NMAX)
      DOUBLE PRECISION DGY_O(1:NMAX,1:NMAX)
      DOUBLE PRECISION FCS_O(1:NMAX,1:NMAX)
      DOUBLE PRECISION OATS_O(1:NMAX)
      DOUBLE PRECISION Q_O(1:NMAX,1:NMAX)
      DOUBLE PRECISION QAVE_O(1:NMAX,1:NMAX)
      DOUBLE PRECISION SSTC_O(1:NMAX)
      DOUBLE PRECISION VW_O(1:NMAX)
      DOUBLE PRECISION ZZZ_O(1:NMAX)
C      integer*2 nx
C      double precision void_i, den_i, flrt_i
C      integer*2 index_i(1:nx)
C      double precision xk_i(1:nx)
C      double precision xn_i(1:nx)
C      double precision c0_i(1:nx)
C      double precision xmw_i(1:nx)
C      integer*2 nflagb
C      double precision c_o(1:nx,1:nx)
C      double precision dgy_o(1:nx,1:nx)
C      double precision fcs_o(1:nx,1:nx)
C      double precision oats_o(1:nx)
C      double precision q_o(1:nx,1:nx)
C      double precision qave_o(1:nx,1:nx)
C      double precision sstc_o(1:nx)
C      double precision vw_o(1:nx)
C      double precision zzz_o(1:nx)
c
c-----// Declaration of external subroutines/functions and common blocks.
      external dneqnf
      external e1init
      external fcn
      common /zsq/ void, den, m, par
      double precision par(100)

c-----// Local Variables.
      integer*2 index(20)

      integer nflag
      integer n
      double precision void
      double precision den
      double precision flrt
      double precision xk(20)
      double precision xn(20)
      double precision cok(20)
      double precision mw(20)
      double precision vf
      integer i
      double precision vw(20)
      integer j, k, l, m, nn
      double precision zz
      double precision errrel
      double precision sum
      double precision xguess(20)
      double precision x(20)
      integer itmax
      double precision fnorm
      double precision q(20, 20)
      double precision c(20, 20)
      double precision dg
      double precision dgx
      integer ix
      double precision tempd
      integer*2 tempi
      double precision cisum

c-----// Output data to be returned from this subroutine.
      double precision bvf(20)
      double precision oats(20)
      double precision qave(20, 20)
      double precision zzz(20)
      double precision sstc(20)
      double precision dgy(20, 20)
      double precision fcs(20, 20)
      double precision cio(20, 20)
c
c   GET VARIABLES PASSED FROM DLL.
c




c           goto 9191

cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc

c      open(unit=9, file='c:\debug.ecm', share='DENYNONE',
c     a   status='UNKNOWN')
c      write(9,*) 'Test file:'
c      write(9,*) '===================='

cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc






      nflag = nflagb

      n = nx
      void = void_i
      den = den_i
      flrt = flrt_i

c      print *, 'Point 1', nflagb
      do 10 i=1, n
        mw(i) = xmw_i(i)
        xk(i) = xk_i(i)
        xn(i) = xn_i(i)
        xn(i) = 1.0D0 / xn(i)
        cok(i) = c0_io(i) / mw(i)
        index(i) = index_i(i)
c        co(i,1) = c0_i(i)
c        co(i,1) = c0_i(i) / mw(i)
10    continue




c      write(9,*) 'Number of components =      ', nx
c      write(9,*) 'Bed void fraction =         ', void_i
c      write(9,*) 'Bulk density of adsorbent = ', den_i
c      write(9,*) 'Flowrate =                  ', flrt_i
c      write(9,*) 'Number of components =      ', n
c      write(9,*) 'Bed void fraction =         ', void
c      write(9,*) 'Bulk density of adsorbent = ', den
c      write(9,*) 'Flowrate =                  ', flrt
c      write(9,*) '.'
c      write(9,*) 'Other inputs (xk(i), xn(i), cok(i), mw(i)):'
c      do 8910 i=1,n
c        write(9,*) i, xk(i), xn(i), cok(i), mw(i)
c8910  continue




c
c   CHANGE UNITS.
c
      vf = flrt * 0.067910D0 / void
      den = den * 1000.0D0
c      print *, 'Point a', nflagb
c
c   SET ZONE ONE CONCENTRATIONS TO ZERO.
c
      do 17 i=1, n
        vw(i) = 0.0D0
        par(20 + i) = cok(i)
        par(60 + i) = 0.0D0
        par(80 + i) = 0.0D0
17    continue
c
c   SOLVE FOR EACH ZONE SEPARATELY.
c



      do 100 j=1, n


c      write(9,*) 'SOLVING FOR ZONE ',J

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      L=0
      M=J
      NN=N+1-J
      ZZ=1.0D0
      ZZ=ZZ/1
      ERRREL=.0000000000001
      SUM=0.0D0
C
C               CALCULATE INITIAL GUESSES OF Q's
C
21      DO 22 I=1,N
         Q(I,J)=ZZ*XK(I)*COK(I)**(1.0D0/XN(I))
22        CONTINUE
C
C               PUT Q INTO ONE-DIMENSIONAL FORM
C
       DO 24 I=1,N
        XGUESS(I)=Q(I,J)
C        X(I)=XGUESS(I)
24     CONTINUE
C
C                    SET IAST PARAMETERS
C
       DO 26 I=1,NN
        XGUESS(I)=XGUESS(M-1+I)
        PAR(I)=XK(M-1+I)
        PAR(10+I)=XN(M-1+I)
        PAR(60+I)=PAR(60+M-1+I)
        PAR(80+I)=PAR(80+M-1+I)
26     CONTINUE
       PAR(30)=VW(J-1)
       PAR(35)=VF
       ITMAX=100
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC



c        l = 0
c        m = j
c        nn = n + 1 - j
c        zz = 1.0D0
c        zz = zz / 1
c        errrel = .0000000000001
cc        errrel = .0001
c        sum = 0.0D0
cc      print *, 'Point b', nflagb
cc
cc   CALCULATE INITIAL GUESSES OF Q'S.
cc
c21      do 22 i=1, n
c          q(i, j) = zz * xk(i) * cok(i)**(1.0D0 / xn(i))
c22      continue
cc
cc   PUT Q INTO ONE-DIMENSIONAL FORM.
cc
c        do 24 i=1, n
c          xguess(i) = q(i, j)
c24      continue
cc
cc   SET IAST PARAMETERS.
cc
c        do 26 i=1, nn
c          xguess(i) = xguess(m - 1 + i)
c          par(i) = xk(m - 1 + i)
c          par(10 + i) = xn(m - 1 + i)
cCCC PAR(20+I)=CO(M-1+I,J)
c          par(60 + i) = par(60 + m - 1 + i)
c          par(80 + i) = par(80 + m - 1 + i)
c26      continue
c
c        par(30) = vw(j - 1)
c        par(35) = vf
c        itmax = 100



c      print *, 'Point c', nflagb

c        write(9,*) '-- Entering dneqnf.'
        call dneqnf(fcn, errrel, nn, itmax,
     &              xguess, x, fnorm, nflag)
c        write(9,*) '-- Exited dneqnf.'



CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C
C...   SET X TO TWO-DIMENSIONAL OUTPUT FOR PRINT OUT
C
       IF (M .GT. 1) THEN
        DO 30 I=1,M-1
         Q(I,J)=0.0D0
30      CONTINUE
       ENDIF
       DO 31 I=1,NN
        Q(I+M-1,J)=X(I)
31     CONTINUE
C
C              CALCULATE THE LIQUID CONCENTRATIONS
C
       IF (M .GT. 1) THEN
        DO 33 I=1,M-1
         C(I,J)=0.0D0
33      CONTINUE
       ENDIF
       DO 34 I=1,NN
        C(I+M-1,J)=PAR(40+I)
34     CONTINUE
C
C           DETERMINE THE STRONGEST COMPONENT IN ZONE J
C
       DGX=0.0D0
       DO 35 I=M,N
        DG=DEN*Q(I,J)/(C(I,J)*VOID)
        IF (DG .GT. DGX) THEN
         DGX=DG
         IX=I
        ENDIF
35     CONTINUE
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC




cc
cc   SET X TO TWO-DIMENSIONAL OUTPUT.
cc
c        if (m .gt. 1) then
c          do 30 i=1, m-1
c            q(i, j) = 0.0D0
c30        continue
c        endif
c        do 31 i=1, nn
c          q(i + m - 1, j) = x(i)
c31      continue
cc
cc   CALCULATE THE LIQUID CONCENTRATIONS.
cc
c        if (m .gt. 1) then
c          do 33 i=1, m-1
c            c(i, j) = 0.0D0
c33        continue
c        endif
c        do 34 i=1, nn
c          c(i + m - 1, j) = par(40 + i)
c34      continue
cc      print *, 'Point d', nflagb
cc
cc   DETERMINE THE STRONGEST COMPONENT IN ZONE J.
cc
c        dgx = 0.0D0
c        do 35 i=m, n
cc      print *, 'Point d-II; i = ', i, nflagb
cc      print *, 'den = ', den
cc      print *, 'void = ', void
cc      print *, 'den * q(i,j) = ', den*q(i,j)
cc      print *, '(c(i, j)*void)', (c(i, j)*void)
c          dg = 0.0D0
c          if (c(i,j) .ne. 0) then
c            dg = den*q(i, j) / (c(i, j)*void)
c            if (dg .gt. dgx) then
cc      print *, 'Point d-I'
c              dgx = dg
c              ix = i
c            endif
c          endif
c35      continue




c      print *, 'Point d2'
c
c   SET STRONGEST COMPONENT TO ZONE J.
c
ccc       CH=CHAR(IX)
ccc       CHAR(IX)=CHAR(J)
ccc       CHAR(J)=CH



ccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc
c      write(9,*) 'Swapping ', ix, 'and', j, '...'
ccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc



        tempd = mw(ix)
        mw(ix) = mw(j)
        mw(j) = tempd

        tempd = xk(ix)
        xk(ix) = xk(j)
        xk(j) = tempd

        tempd = xn(ix)
        xn(ix) = xn(j)
        xn(j) = tempd

        tempd = cok(ix)
        cok(ix) = cok(j)
        cok(j) = tempd

        tempi = index(ix)
        index(ix) = index(j)
        index(j) = tempi

c      print *, 'Point e'
        do 36 k=1, j
CCC          tempd = CO(IX,K)
CCC          CO(IX,K) = CO(J,K)
CCC          CO(J,K) = tempd

          tempd = c(ix, k)
          c(ix, k) = c(j, k)
          c(j, k) = tempd

          tempd = q(ix, k)
          q(ix, k) = q(j, k)
          q(j, k) = tempd
36      continue




CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C
C                 SET C's AND Q's FOR NEXT ZONE
C
       DO 38 I=1,N
        PAR(60+I)=Q(I,J)
        PAR(80+I)=C(I,J)
38     CONTINUE
C
C           CALCULATE VELOCITY OF THE WAVE FOR ZONE J
C
       IF (J .EQ. 1) THEN
        VW(J)=VF*VOID*COK(1)/(Q(1,J)*DEN+C(1,J)*VOID)
       ENDIF
       IF (J .GE. 2) THEN
        SUM=(Q(J,1)*DEN+VOID*C(J,1))*VW(1)
       ENDIF
       IF (J .GT. 2) THEN
        DO 40 K=2,J-1
         SUM=SUM+((Q(J,K)*DEN+VOID*C(J,K))*(VW(K)-VW(K-1)))
40      CONTINUE
       ENDIF
       IF (J .GE. 2) THEN
        VW(J)=(VOID*VF*COK(J)-SUM+(Q(J,J)*DEN+VOID*C(J,J))*VW(J-1)
     $)/(Q(J,J)*DEN+VOID*C(J,J))
       ENDIF
C
C                     SET Cok FOR NEXT ZONE
C
       DO 50 I=J+1,N
        C(I,J+1)=C(I,J)
50     CONTINUE
       DO 60 I=1,J
        C(I,J+1)=0.0D0
60     CONTINUE
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC



cc
cc   SET C'S AND Q'S FOR NEXT ZONE.
cc
c        do 38 i=1, n
c          par(60 + i) = q(i, j)
c          par(80 + i) = c(i, j)
c38      continue
cc      print *, 'Point f'
cc
cc   CALCULATE VELOCITY OF THE WAVE FOR ZONE J.
cc
c        if (j .eq. 1) then
cCCC       VW(J)=VF*BVF*CO(1,J)/(Q(1,J)*DEN+C(1,J)*BVF)
c          vw(j) = vf*void*cok(1) /
c     &            (q(1, j)*den + c(1, j)*void)
c        endif
c        if (j .eq. 2) then
c          sum = (q(j, 1)*den + void*c(j,1))*vw(1)
c        endif
c        if (j .gt. 2) then
c          do 40 k=2, j-1
c            sum = sum + ((q(j, k)*den +
c     &            void*c(j, k))*(vw(k) - vw(k - 1)))
c40        continue
c        endif
c        if (j .ge. 2) then
cCCC       VW(J)=(BVF*VF*CO(J,J)-SUM+(Q(J,J)*DEN+BVF*C(J,J))*VW(J-1)) /
cCCC  &          (Q(J,J)*DEN+BVF*C(J,J))
c          vw(j) = (void*vf*cok(j) - sum +
c     &            (q(j, j)*den + void*c(j, j)) * vw(j-1)) /
c     &            (q(j, j)*den + void*c(j, j))
c        endif
cc      print *, 'Point g'
cc
cc   SET COK FOR NEXT ZONE.
cc
cCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC        if (j+1 .le. n) then
c        do 50 i=j+1, n
c          c(i, j+1) = c(i, j)
c50      continue
cCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC        endif
c        do 60 i=1, j
c          c(i, j+1) = 0.0D0
c60      continue




c       write(9,*) 'Values of i, C(i,j), Q(i,j), j=', j, ':'
c       do 8930 i=1,n
c         write(9,*) i, c(i,j), q(i,j)
c8930   continue







100   continue


c
c   CALCULATE BED VOLUMES FED.
c
c      print *, 'Point h'
      do 110 i=1, n
        bvf(i) = vf * void / vw(i)

CCCCCCCCCCCCCCCC ** Note: the following 7 lines are not in
CCCCCCCCCCCCCCCC    their program!
        SUM=(Q(I,1)*DEN+C(I,J)*void)*VW(1)
        IF (I .GE. 2) THEN
          DO 105 K=2,I
            SUM=SUM+(Q(I,K)*DEN+C(I,K)*void)*(VW(K)-VW(K-1))
105       CONTINUE
        ENDIF
        OATS(I)=SUM/(COK(I)*VW(I))

110   continue
c
c   CALCULATE Q TOTAL AVERAGE FOR EACH ZONE.
c
      do 120 i=1, n
        qave(i, 1) = q(i, 1)
120   continue
      do 140 i=1, n
        do 130 j=2, n
          qave(i, j) = (qave(i, j-1)*vw(j-1) + q(i, j) *
     &                 (vw(j) - vw(j - 1))) / vw(j)
130     continue
140   continue
c
c   CALCULATE DIMENSIONLESS BED LENGTH.
c
c      print *, 'Point i'
      do 150 i=1, n
        if (i .eq. 1) then
          zzz(1) = 0.0D0
        else
          zzz(i) = vw(i - 1) / vw(n)
        endif
        zzz(i) = vw(i) / vw(n) - zzz(i)
150   continue
c
c   CALCULATE SINGLE SOLUTE TREATMENT CAPACITY.
c
      do 160 i=1, n
        sstc(i) = 1000D0 * cok(i)**(1.0D0 - 1.0D0/xn(i)) / xk(i)
160   continue
c
c   CALCULATE DG'S.
c
      do 180 j=1, n
        do 170 i=1, n
CCC       IF (C(I,J) .EQ. 0.0) DGY(I,J)=0.0
CCC       if (c(i,j) .ne. 0.0) DGY(I,J)=DEN*Q(I,J)/(C(I,J)*BVF)
          if (c(i, j) .eq. 0.0) dgy(i, j) = 0.0
          if (c(i, j) .gt. 0.0) then
            dgy(i, j) = den * q(i, j) / (c(i, j) * void)
          endif
170     continue
180   continue
c      print *, 'Point j'
c
c   CALCULATE C/COK.
c
CCCccc      do 195 j=1, n
CCCccc        do 190 i=1, n
CCCcccCCC       IF (CO(I,J) .EQ. 0.0) FCS(I,J)=0.0
CCCcccCCC       if (co(i,j) .ne. 0.0) FCS(I,J)=C(I,J)/CO(I,J)
CCCccc          if (c(i, j-1) .eq. 0.0) fcs(i, j) = 0.0
CCCccc          if (c(i, j-1) .gt. 0.0) then
CCCccc            fcs(i, j) = c(i, j) / cok(i)
CCCccc          endif
CCCccc190     continue
CCCccc195   continue

      do 195 j=1, n
        do 190 i=1, n
CCC       IF (CO(I,J) .EQ. 0.0) FCS(I,J)=0.0
CCC       if (co(i,j) .ne. 0.0) FCS(I,J)=C(I,J)/CO(I,J)
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
          if (j .gt. 1) then
            if (c(i, j-1) .eq. 0.0) fcs(i, j) = 0.0
            if (c(i, j-1) .gt. 0.0) then
              fcs(i, j) = c(i, j) / cok(i)
            endif
          endif
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
190     continue
195   continue
c
c   CALCULATE CIO.
c
      do 250 j=1, n
        cisum = 0.0D0
        do 230 i=1, n
          cisum = cisum + xn(i)*q(i,j)
230     continue
        do 240 i=1, n
          cio(i, j) = (cisum / xn(i) / xk(i))**xn(i)
240     continue
250   continue
c      print *, 'Point k'
c
c   PUT INTO UG/L UNITS.
c
      do 210 j=1, n
        cok(j) = cok(j) * mw(j)
        do 200 i=1, n
CCC       CO(I,J)=CO(I,J)*MW(I)
          cio(i, j) = cio(i, j) * mw(i)
          q(i, j) = q(i, j) * mw(i)
          qave(i, j) = qave(i, j) * mw(i)
          c(i, j) = c(i, j) * mw(i)
200     continue
210   continue
c
c   SET VALUES FOR OUTPUT.
c
c      print *, 'Point m'
      do 9002 i=1, n
        do 9001 j=1, n
          c_o(i, j) = c(i, j)
          dgy_o(i, j) = dgy(i, j)
          fcs_o(i, j) = fcs(i, j)
          q_o(i, j) = q(i, j)
C          print *, 'q_o(i,j) = ', q_o(i,j)
          qave_o(i, j) = qave(i, j)
9001    continue
        oats_o(i) = oats(i)
        sstc_o(i) = sstc(i)
        vw_o(i) = vw(i)
        zzz_o(i) = zzz(i)
        index_i(i) = index(i)
        c0_io(i) = cok(i)
9002  continue

C      pause 'testing ...'
C      DO I=1,N
C        DO J=1,N
C          Q_O(I,J) = Q(I,J)
C          print *, 'q_o(i,j) = ', q_o(i,j)
C        ENDDO      
C      ENDDO

c      do 9910 i=1,n
c        write(9,*) i, ', bvf(i)=', bvf(i), ', vw(i)=', vw(i)
c9910  continue



      nflagb = nflag



cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc

c      close(unit=9)

cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc

c9191      nflagb = 1010


      return
      end
c
c
c
c
c   SUBROUTINE FCN
c
c   THIS SUBROUTINE WILL SET UP THE EQUATIONS THAT WILL
c   BE USED IN THE DNEQNF SUBROUTINE.
c
c
c
c
      subroutine fcn(x, f, nn)
c
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DIMENSION X(NN),F(NN)
      common /zsq/ void, den, m, par
      double precision par(100)

      QT=0.0D0
      QNQ=0.0D0
      DO 1010 I=1,NN
        QT=QT+X(I)
        QNQ=QNQ+PAR(10+I)*X(I)
1010  CONTINUE
C
C                       CALCULATE F(I)
C
      IF (M .EQ. 1) THEN
        DO 1020 I=1,NN
          F(I)=-PAR(20+I)+X(I)/QT*(QNQ/PAR(10+I)/PAR(I))**PAR(10+I)
1020    CONTINUE
      ENDIF
      IF (M .GT. 1) THEN
        DO 1030 I=1,NN
          F(I)=-X(I)/QT*(QNQ/PAR(10+I)/PAR(I))**PAR(10+I)+
     &         ((X(I)-PAR(60+I))*
     &         DEN*PAR(30))/((PAR(35)-PAR(30))*VOID)+PAR(80+I)
1030    CONTINUE
      ENDIF
C
C                 CALCULATE LIQUID CONCENTRATION
C
      IF (M .EQ. 1) THEN 
        DO 1040 I=1,NN
          PAR(40+I)=X(I)*(QNQ/(PAR(10+I)*PAR(I)))**PAR(10+I)/QT
1040    CONTINUE
      ENDIF
      IF (M .GT. 1) THEN
        DO 1050 I=1,NN
          PAR(40+I)=((X(I)-PAR(60+I))*DEN*PAR(30))/
     &              ((PAR(35)-PAR(30))*VOID)+
     &              PAR(80+I)
1050    CONTINUE
      ENDIF

c      do 2019 i=1, nn
c        write(9,*) i, 'f(i) = ', f(i)
c2019  continue

      RETURN
      END
C********************************************************************
C
C                      MASS_BALANCE
C 
C Description:  This routine will do the mass balance on the output
C               from the ECM program for each component and tell
C               the percent error on the mass balance.
C
C Input Variables:
C    N =        Number of Components
C    VW =       Array of Wave Velocities for each zone (1 to N)
C               (cm/s)
C    C =        Array of Liquid Phase Concentrations for Each
C               Component in Each Zone : C(Component,Zone) -
C               N x N two-dimensional array (ug/L)
C    Q =        Array of Gas Phase Concentrations for Each 
C               Component in Each Zone : q(Component,Zone) -
C               N x N two-dimensional array (ug/g) 
C    EBED =     Void Fraction of Bed (-)
C    DEN =      Bulk Density of Adsorbent (g/cm3)
C    FLRT =     Flowrate (gpm/ft2)
C    COK =      Array of Liquid Phase Influent Concentrations 
C               (1 to N) (ug/L)
C
C Output Variables:
C    C0_e_Vf =  Left-hand side of mass balance (ug/cm2/s). Array
C               from 1 to N.
C    TERM_SUM = Right-hand side of mass balance (ug/cm2/s).
C               Array from 1 to N.
C    PERCENT_ERR = Percent difference between C0_e_Vf and
C                  TERM_SUM (%). Array from 1 to N.
C
C Variables internal to this Subroutine:
C    VF =       Interstitial fluid velocity (L/cm2/s)
C
C********************************************************************

      SUBROUTINE ECM_MASSBAL (N,VW,C,Q,EBED,DEN,FLRT,COK,C0_e_Vf,
     &                         TERM_SUM,PERCENT_ERR,VF)

      IMPLICIT NONE
      INTEGER N,I,J,K
      DOUBLE PRECISION VW(N),C(N,N),Q(N,N),EBED,DEN,FLRT,COK(N)
      DOUBLE PRECISION TERM,C0_e_Vf(N),TERM_SUM(N),
     &                 PERCENT_ERR(N),VF

      VF = FLRT * 1000.0D0 / 60.0D0 / (30.48D0**2) / 264.17D0 / EBED

      DO 10, I = 1,N
         C0_e_Vf(I) = COK(I) * VF * EBED    
 10   CONTINUE        

C**** Note:  I = Number of Component, J = Number of Zone

      DO 20, I = 1,N
         TERM_SUM(I) = 0.0D0
         DO 30, J = 1,N
            IF (J.EQ.1) THEN
               TERM = VW(J) * (Q(I,J)*DEN+C(I,J)*EBED/1000.0D0)
            ELSE
               TERM = (VW(J)-VW(J-1)) * 
     &                (Q(I,J)*DEN+C(I,J)*EBED/1000.0D0)           
            END IF
            TERM_SUM(I) = TERM_SUM(I) + TERM            
 30      CONTINUE
         PERCENT_ERR(I) = DABS(((C0_e_Vf(I)-TERM_SUM(I))/C0_e_Vf(I)))
     &                      * 100.0D0
 20   CONTINUE      

      END

C********************************************************************




