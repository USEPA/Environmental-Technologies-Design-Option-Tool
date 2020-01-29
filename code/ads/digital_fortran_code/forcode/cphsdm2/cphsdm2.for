C  USER ORIENTED SOLUTIONS TO THE HSDM FOR THE DESIGN OF FIXED-BED
C  ADSORPTION SYSTEMS     
C                                                                     
C  PROGRAM DEVELOPED BY : RANDY D. CORTRIGHT
C                         DAVID W. HAND
C                         JOHN C. CRITTENDEN
C                        
C  DATE: APRIL 14, 1986                                           
C                                                                     
C  THIS PROGRAM USES THE PROCEDURE DESCRIBED IN THE FOLLOWING PAPER   
C  TO CALCULATE THE EMPTY BED CONTACT TIME OF THE MASS TRANSFER ZONE  
C  WITHIN ADSORBER AND CALCULATE THE MASS TRANSFER PROFILE:                   
C                                                                     
C  HAND, DAVID W.,JOHN C. CRITTENDEN, AND WILLIAM E. THACKER,         
C  'SIMPLIFIED MODELS FOR DESIGN OF FIXED-BED ADSORPTION SYSTEMS',    
C  JOURNAL OF ENVIRONMENTAL ENGINEERING, VOL 110, No. 2, April, 1984. 
C                                                                     
C--------------------------------------------------------------------
C Slightly modified by F. Gobin to be compiled as a DLL and used from
C V.B. 3.0
C Last Modified 10/26/94
C ----------------- Input/Ouput Parameters --------------------------
C
C   Bed   : (1) -> DIA  PArticle diameter         cm     Input
C           (2) -> RHOB apparent bed density      g/cm3 
C           (3) -> RHOP apparent particle density g/cm3
C           (4) -> EBCT                           min  
C           (5) -> VS superficial velocity        cm/s
C           (6) -> EPOR Porosity of particle      -
C   Compo : (1) -> MW  Molecular Weight           g/cm3
C           (2) -> CBO inlet concentration        ug/l   
C           (3) -> K Freundlich K parameter       (*) 
C           (4) -> N Freundlich 1/n parameter      -
C   Kine  : (1) -> KF                 cm/s               I
C           (2) -> DS                 cm2/s              I
C   TACT  :  Time for C/C0			  days	                 Output
C   CC    :  C/C0 - size 210			  -	                   Output
C   PARAM :  Various parameters - size 7                 Output
C	      (1) -> minimum Stanton number         -
C           (2) ->
C           (3) -> 	
C	      (4) -> time for 95 % of MTZ            days
C           (5) -> time for 5 % of MTZ             days
C           (6) -> 	
C           (7) -> 	
C--------------- Error Flag -----------------------------------------
C   ER_FLAG : = 40,41,41,43,44                           Output
C             see listing for corresponding message
C--------------------------------------------------------------------
C Note :  (*) K in (umol/g)*(l/umol)^(1/n)

      SUBROUTINE CPM (Bed,Compo,Kine,
     &                TACT,CC,PARAM,ER_FLAG)
      IMPLICIT NONE
      INTEGER*2 NFLAG
C
C------ PARAMETERS PASSED TO SUBROUTINE.
C
      DOUBLE PRECISION Bed(1:6)
      DOUBLE PRECISION Compo(1:4)
      DOUBLE PRECISION Kine(1:2)
      DOUBLE PRECISION TACT(1:210)
      DOUBLE PRECISION CC(1:210)
      DOUBLE PRECISION PARAM(1:7)
      INTEGER*2 ER_FLAG
C
C------ LOCAL VARIABLES.
C
      DOUBLE PRECISION RHOB,DIA,RHOP,EPOR,VS,EBCT,K,N,CBO,MW
      DOUBLE PRECISION KF                                                
      DOUBLE PRECISION T(210),TMIN(210),
     $                 BVFACT(210),USAGE(210)                                
      DOUBLE PRECISION Q,SF,EBED,DG 
      DOUBLE PRECISION RAD,DS,BI,STM,STMIN,ETMIN,EMLEN,
     &                 TAUMIN,TAUACT,A0,A1,A2,A3,A4,T95,
     &                 T05,ETMTZ,EMTZL 
      INTEGER I 
     
      COMMON /FLAG/ NFLAG

C----- Initialization of variables  ---------------------------------

      DIA  = Bed(1)
      RHOB = Bed(2)   
      RHOP = Bed(3) 
      EBCT = Bed(4)    
      VS   = Bed(5) 
      EPOR = Bed(6) 

      MW = Compo(1)
      CBO= Compo(2)
      K  = Compo(3)
      N  = Compo(4)

      KF = KINE(1)
      DS = KINE(2) 
C                                                                     
C  CALC. THE EQUILBRIUM CONCENTRATION ON THE CARBON AND IN THE LIQUID 
C                                                                     
C                                                                     
      CBO = CBO/MW                                                    
      Q = K*CBO**N                                                    
      EBED=1-RHOB/RHOP

      SF = VS*14.724                                                  
C                                                                     
C  CALCULATE THE REYNOLDS AND SCHMIDT NUMBERS, AND THE FILM TRANSFER C
C                                                                     
C      RE = (DIA*VS*DW)/(VW*EBED)                                      
C      SC = (VW/(DW*DIFL))                                             
      DG = RHOP*Q*(1-EBED)*1000.0/(EBED*CBO)                          
      RAD = DIA/2.0                                                   

C  CALCULATE THE PARTITION COEFFICIENT                                
C                                                                     
      DG = (RHOP * Q * (1.0 - EBED) * 1000.0) / (EBED * CBO)          
C                                                                     
C  CALCULATE THE BIOT NUMBER                                          
C                                                                     
      BI = (KF * DIA/2.0 * (1.0 - EBED))/(DG * DS * EBED)             

C     
C  CALCULATE THE MINIMUN STANTON NUMBER AND THE EBCT MIN              
C                                                                     
      STM = STMIN(BI,N)                                               
      ETMIN =  (STM * DIA/2.0) / (KF * (1.0 - EBED))                  
      EMLEN = VS * ETMIN                                              
C  CALCULATE THE THROUGHPUT FOR 5 AND 95 PERCENT BREAKTHRU AND        
C  FIND THE EBCT FOR THE MASS TRANSFER ZONE                           
C                                                                     
      TAUMIN = ETMIN*EBED/60.0                                        
      TAUACT = EBED*EBCT                                              
      CALL TPUT(N,BI,A0,A1,A2,A3,A4)                                  
      T95 = A0 + A1 * (0.95**A2) + A3/(1.01 - 0.95**A4)               
      T05 = A0 + A1 * (0.05**A2) + A3/(1.01 - 0.05**A4)               
      ETMTZ = ETMIN * (T95 - T05)                                     
      EMTZL = ETMTZ * VS                                              
      CBO = CBO*MW                                                    
C                                                                      
C  CALCULATE THE PROFILE FOR THE SINGLE SOLUTE SYSTEM                 
C             
                                                      
      CC(1) = 0.01                                                    
      DO 65 I = 1, 200
         T(I) = A0 + A1 * (CC(I)**A2) + A3/(1.01 - CC(I)**A4)         
         CC(I+1) = CC(I) + 0.01                                       
         TMIN(I) = TAUMIN*(DG+1.)*T(I)                                
         TACT(I) = TMIN(I) + (TAUACT-TAUMIN)*(DG+1.)                  
         TMIN(I) = TMIN(I)/1440.0                                     
         BVFACT(I) = TACT(I)*(EBED)/(TAUACT)                                 
         TACT(I) = TACT(I)/1440.0                                     
         USAGE(I) = BVFACT(I)/(RHOB*1000.0)
                         
65    CONTINUE                                                        
     
      PARAM(1)= STM
      PARAM(2)= ETMIN/60.0
      PARAM(3)= EMLEN
      PARAM(4)= T95
      PARAM(5)= T05
      PARAM(6)= ETMTZ/60.0
      PARAM(7)= EMTZL

  999 ER_FLAG=NFLAG
      RETURN
      END                                                             
C                                                                     
C                                                                     
C  FUNCTION STMIN FOR FINDING THE MINIMUN STANTON NUMBER REQUIRED     
C  FOR CONSTANT PATTERN                                               
C                                                                     
C                                                                     
      DOUBLE PRECISION FUNCTION STMIN(BI,N)                                       
      IMPLICIT NONE
      INTEGER*2 NFLAG 
      INTEGER I,J,M
      DOUBLE PRECISION N,BI,A0,A1
      DOUBLE PRECISION FN(10),A01(10),A11(10),A02(10)                        

      COMMON /FLAG/ NFLAG

      DATA (FN(I), I = 1,10)/0.05,0.10,0.20,0.30,0.40,0.50,0.60,0.70, 
     $                       0.80,0.90/                               
      DATA (A01(I),I = 1,10)/2.10526E-2,2.10526E-2,4.21053E-2,        
     $                       1.05263E-1,2.31579E-1,5.26316E-1,        
     $                       1.15789,1.78947,3.68421,6.31579/         
      DATA (A11(I),I = 1,10)/1.98947,2.18947,2.37895,2.54737,2.68421, 
     $                       2.73684,3.42105,7.10526,13.1579,56.8421/ 
      DATA (A02(I),I = 1,10)/0.22,0.24,0.28,0.36,0.50,0.80,1.50,2.50, 
     $                       5.00,12.00/                              
C                                                                     
C                                                                     
C                                                                     
      M = 10                                                          
      IF((BI .GE. 0.5) .AND. (BI .LE. 10.0)) THEN                     
         J = 1                                                        
10       IF(J .LE. M) THEN                                            
            IF ((N .GE. FN(J)) .AND. (N .LT. FN(J+1))) THEN           
               A0 = A01(J) + (A01(J+1)-A01(J)) * ((N - FN(J))/        
     $                                         (FN(J+1) - FN(J)))     
               A1 = A11(J) + (A11(J+1)-A11(J)) * ((N - FN(J))/        
     $                                         (FN(J+1) - FN(J)))     
               STMIN = A0 * BI + A1                                   
               GO TO 30                                               
            ELSE                                                      
               J = J + 1                                              
               GO TO 10                                               
            END IF                                                    
         ELSE                                                         
C            PRINT*, ' THE VALUE OF 1/N IS OUT OF RANGE FOR STMIN'     
           NFLAG = 40
         ENDIF                                                        
      ELSEIF (BI .GT. 10.0) THEN                                      
         J = 1                                                        
20       IF(J .LE. M) THEN                                            
            IF ((N .GE. FN(J)) .AND. (N .LT. FN(J+1))) THEN           
               A0 = A02(J) + (A02(J+1)-A02(J)) * ((N - FN(J))/        
     $                                         (FN(J+1) - FN(J)))     
               STMIN = A0 * BI                                        
               GO TO 30                                               
            ELSE                                                      
               J = J + 1                                              
               GO TO 20                                               
            END IF                                                    
         ELSE                                                         
C            PRINT*, ' THE VALUE OF 1/N IS OUT OF RANGE FOR STMIN'     
             NFLAG = 41
         ENDIF                                                        
      ELSE                                                            
C         PRINT*, ' THE VALUE OF THE BIOT NUMBER IS OUT OF RANGE'      
C         PRINT*, ' BIOT NUMBER = ',BI                                 
        NFLAG = 42
      ENDIF                                                           
30    RETURN                                                          
      END                                                             

C                                                                     
C  SUBROUTINE TPUT TO FIND THE CONSTANTS TO FIND EBCTMIN              
C                                                                     
C                                                                     
      SUBROUTINE TPUT(N,BI,A0,A1,A2,A3,A4)                            
      IMPLICIT NONE
      DOUBLE PRECISION N,BI,A0,A1,A2,A3,A4
      INTEGER*2 NFLAG                                                    
      COMMON /FLAG/ NFLAG

      IF (N .LT. 0.075) THEN                                          
         CALL T1(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.075) .AND. (N .LT. 0.15)) THEN                 
         CALL T2(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.15) .AND. (N .LT. 0.25)) THEN                  
         CALL T3(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.25) .AND. (N .LT. 0.35)) THEN                  
         CALL T4(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.35) .AND. (N .LT. 0.45)) THEN                  
         CALL T5(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.45) .AND. (N .LT. 0.55)) THEN                  
         CALL T6(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.55) .AND. (N .LT. 0.65)) THEN                  
         CALL T7(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.65) .AND. (N .LT. 0.75)) THEN                  
         CALL T8(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.75) .AND. (N .LT. 0.85)) THEN                  
         CALL T9(BI,A0,A1,A2,A3,A4)                                   
      ELSEIF((N .GE. 0.85) .AND. (N .LT. 1.00)) THEN                  
         CALL T10(BI,A0,A1,A2,A3,A4)                                  
      ELSE                                                            
C         PRINT*, ' THE VALUE OF 1/N IS OUT OF RANGE'                  
C         PRINT*, ' 1/N = ',N                                          
       NFLAG = 44
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T1(BI,A0,A1,A2,A3,A4)
      IMPLICIT NONE
      DOUBLE PRECISION BI,A0,A1,A2,A3,A4                               
      IF(BI .LT. 1.25) THEN                                           
         A0 = -5.447214                                               
         A1 = 6.598598                                                
         A2 = 0.026569                                                
         A3 = 0.019384                                                
         A4 = 20.45047                                                
      ELSEIF((BI .GE. 1.25) .AND. (BI .LT. 3.0)) THEN                 
         A0 = -5.465811                                               
         A1 = 6.592484                                                
         A2 = 0.025290                                                
         A3 = 0.004988                                                
         A4 = 0.503250                                                
      ELSEIF((BI .GE. 3.0) .AND. (BI .LT. 5.0)) THEN                  
         A0 = -5.531155                                               
         A1 = 6.584935                                                
         A2 = 0.023580                                                
         A3 = 0.009019                                                
         A4 = 0.273076                                                
      ELSEIF((BI .GE. 5.0) .AND. (BI .LT. 7.0)) THEN                  
         A0 = -5.606508                                               
         A1 = 6.582188                                                
         A2 = 0.022088                                                
         A3 = 0.013126                                                
         A4 = 0.214246                                                
      ELSEIF((BI .GE. 7.0) .AND. (BI .LT. 9.0)) THEN                  
         A0 = -5.606500                                               
         A1 = 6.504701                                                
         A2 = 0.020872                                                
         A3 = 0.017083                                                
         A4 = 0.189537                                                
      ELSEIF((BI .GE. 9.0) .AND. (BI .LT. 12.0)) THEN                 
         A0 = -5.664173                                               
         A1 = 6.456597                                                
         A2 = 0.018157                                                
         A3 = 0.019935                                                
         A4 = 0.149314                                                
      ELSEIF((BI .GE. 12.0) .AND. (BI .LT. 19.5)) THEN                
         A0 = -0.662780                                               
         A1 = 1.411252                                                
         A2 = 0.060709                                                
         A3 = 0.020229                                                
         A4 = 0.143293                                                
      ELSEIF((BI .GE. 19.5) .AND. (BI .LT. 62.5)) THEN                
         A0 = -0.662783                                               
         A1 = 1.350940                                                
         A2 = 0.031070                                                
         A3 = 0.020350                                                
         A4 = 0.129998                                                
      ELSEIF(BI .LT. 62.5) THEN                                       
         A0 = 0.665879                                                
         A1 = 0.711310                                                
         A2 = 2.987309                                                
         A3 = 0.016783                                                
         A4 = 0.361023                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T2(BI,A0,A1,A2,A3,A4)
      IMPLICIT NONE
      DOUBLE PRECISION BI,A0,A1,A2,A3,A4                                
      IF(BI .LT. 1.25) THEN                                           
         A0 = -1.919873                                               
         A1 = 3.055368                                                
         A2 = 0.055488                                                
         A3 = 0.024284                                                
         A4 = 15.311766                                               
      ELSEIF((BI .GE. 1.25) .AND. (BI .LT. 3.0)) THEN                 
         A0 = -2.278950                                               
         A1 = 3.393925                                                
         A2 = 0.046838                                                
         A3 = 0.004751                                                
         A4 = 0.384675                                                
      ELSEIF((BI .GE. 3.0) .AND. (BI .LT. 5.0)) THEN                  
         A0 = -2.337178                                               
         A1 = 3.379926                                                
         A2 = 0.043994                                                
         A3 = 0.008650                                                
         A4 = 0.243412                                                
      ELSEIF((BI .GE. 5.0) .AND. (BI .LT. 7.0)) THEN                  
         A0 = -2.407407                                               
         A1 = 3.374131                                                
         A2 = 0.041322                                                
         A3 = 0.012552                                                
         A4 = 0.196565                                                
      ELSEIF((BI .GE. 7.0) .AND. (BI .LT. 9.0)) THEN                  
         A0 = -2.477819                                               
         A1 = 3.370954                                                
         A2 = 0.038993                                                
         A3 = 0.016275                                                
         A4 = 0.176437                                                
      ELSEIF((BI .GE. 9.0) .AND. (BI .LT. 13.0)) THEN                 
         A0 = -2.566414                                               
         A1 = 3.370950                                                
         A2 = 0.035003                                                
         A3 = 0.019386                                                
         A4 = 0.150788                                                
      ELSEIF((BI .GE. 13.0) .AND. (BI .LT. 23.0)) THEN                
         A0 = -2.567201                                               
         A1 = 3.306341                                                
         A2 = 0.020940                                                
         A3 = 0.019483                                                
         A4 = 0.136813                                                
      ELSEIF((BI .GE. 23.0) .AND. (BI .LT. 65.0)) THEN                
         A0 = -2.568618                                               
         A1 = 3.241783                                                
         A2 = 0.009595                                                
         A3 = 0.019610                                                
         A4 = 0.121746                                                
      ELSEIF(BI .GE. 65.0) THEN                                       
         A0 = -2.568360                                               
         A1 = 3.191482                                                
         A2 = 0.001555                                                
         A3 = 0.019682                                                
         A4 = 0.110113                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T3(BI,A0,A1,A2,A3,A4)
      IMPLICIT NONE
      DOUBLE PRECISION BI,A0,A1,A2,A3,A4                               
      IF (BI .LT. 1.25) THEN                                          
         A0 = -1.441000                                               
         A1 = 2.569000                                                
         A2 = 0.060920                                                
         A3 = 0.002333                                                
         A4 = 0.371100                                                
      ELSEIF((BI .GE. 1.25) .AND. (BI .LT. 3.0)) THEN                 
         A0 = -1.474313                                               
         A1 = 2.558300                                                
         A2 = 0.058480                                                
         A3 = 0.005026                                                
         A4 = 0.241265                                                
      ELSEIF((BI .GE. 3.0) .AND. (BI .LT. 5.0)) THEN                  
         A0 = -1.506696                                               
         A1 = 2.519259                                                
         A2 = 0.055525                                                
         A3 = 0.008797                                                
         A4 = 0.187510                                                
      ELSEIF((BI .GE. 5.0) .AND. (BI .LT. 7.0)) THEN                  
         A0 = -1.035395                                               
         A1 = 1.983018                                                
         A2 = 0.069283                                                
         A3 = 0.012302                                                
         A4 = 0.167924                                                
      ELSEIF((BI .GE. 7.0) .AND. (BI .LT. 9.0)) THEN                  
         A0 = -0.169192                                               
         A1 = 1.077521                                                
         A2 = 0.144879                                                
         A3 = 0.015500                                                
         A4 = 0.168083                                                
      ELSEIF((BI .GE. 9.0) .AND. (BI .LT. 11.5)) THEN                 
         A0 = -1.402932                                               
         A1 = 2.188339                                                
         A2 = 0.052191                                                
         A3 = 0.018422                                                
         A4 = 0.133574                                                
      ELSEIF((BI .GE. 11.5) .AND. (BI .LT. 19.0)) THEN                
         A0 = -1.369220                                               
         A1 = 2.118545                                                
         A2 = 0.039492                                                
         A3 = 0.018453                                                
         A4 = 0.127565                                                
      ELSEIF((BI .GE. 19.0) .AND. (BI .LT. 62.5)) THEN                
         A0 = -1.514159                                               
         A1 = 2.209450                                                
         A2 = 0.017937                                                
         A3 = 0.018510                                                
         A4 = 0.118517                                                
      ELSEIF(BI .GE. 62.5) THEN                                       
         A0 = 0.680346                                                
         A1 = 0.649006                                                
         A2 = 2.570086                                                
         A3 = 0.014947                                                
         A4 = 0.369818                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T4(BI,A0,A1,A2,A3,A4) 
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
      IF (BI .LT. 1.25) THEN                                          
         A0 = -1.758696                                               
         A1 = 2.846576                                                
         A2 = 0.049530                                                
         A3 = 0.003022                                                
         A4 = 0.156816                                                
      ELSEIF ((BI .GE. 1.25) .AND. (BI .LT. 3.0)) THEN                
         A0 = -1.657862                                               
         A1 = 2.688895                                                
         A2 = 0.048409                                                
         A3 = 0.005612                                                
         A4 = 0.140937                                                
      ELSEIF ((BI .GE. 3.0) .AND. (BI .LT. 5.0)) THEN                 
         A0 = -0.565664                                               
         A1 = 1.537833                                                
         A2 = 0.084451                                                
         A3 = 0.008808                                                
         A4 = 0.139086                                                
      ELSEIF ((BI .GE. 5.0) .AND. (BI .LT. 7.0)) THEN                 
         A0 = -0.197077                                               
         A1 = 1.118564                                                
         A2 = 0.117894                                                
         A3 = 0.011527                                                
         A4 = 0.135874                                                
      ELSEIF ((BI .GE. 7.0) .AND. (BI .LT. 9.0)) THEN                 
         A0 = -0.197070                                               
         A1 = 1.069216                                                
         A2 = 0.119760                                                
         A3 = 0.013925                                                
         A4 = 0.132691                                                
      ELSEIF ((BI .GE. 9.0) .AND. (BI .LT. 12.5)) THEN                
         A0 = -0.173358                                               
         A1 = 1.00000                                                 
         A2 = 0.120311                                                
         A3 = 0.015940                                                
         A4 = 0.133973                                                
      ELSEIF ((BI .GE. 12.5) .AND. (BI .LT. 25.0)) THEN               
         A0 = -0.173350                                               
         A1 = 0.919411                                                
         A2 = 0.071768                                                
         A3 = 0.014156                                                
         A4 = 0.086270                                                
      ELSEIF ((BI .GE. 25.0) .AND. (BI .LT. 67.5)) THEN               
         A0 = 0.666471                                                
         A1 = 0.484570                                                
         A2 = 1.719440                                                
         A3 = 0.013444                                                
         A4 = 0.259545                                                
      ELSEIF (BI .GE. 67.5) THEN                                      
         A0 = 0.696161                                                
         A1 = 0.516951                                                
         A2 = 2.054587                                                
         A3 = 0.012961                                                
         A4 = 0.303218                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T5(BI,A0,A1,A2,A3,A4)
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
                                
      IF (BI .LT. 1.25) THEN                                          
         A0 = -0.534251                                               
         A1 = 1.603834                                                
         A2 = 0.094055                                                
         A3 = 0.004141                                                
         A4 = 0.137797                                                
      ELSEIF((BI .GE. 1.25) .AND. (BI .LT. 3.0)) THEN                 
         A0 = -0.166270                                               
         A1 = 1.190897                                                
         A2 = 0.122280                                                
         A3 = 0.006261                                                
         A4 = 0.134278                                                
      ELSEIF((BI .GE. 3.0) .AND. (BI .LT. 5.0)) THEN                  
         A0 = -0.166270                                               
         A1 = 1.131946                                                
         A2 = 0.115513                                                
         A3 = 0.008634                                                
         A4 = 0.126813                                                
      ELSEIF((BI .GE. 5.0) .AND. (BI .LT. 7.5)) THEN                  
         A0 = -0.166270                                               
         A1 = 1.089789                                                
         A2 = 0.112284                                                
         A3 = 0.010463                                                
         A4 = 0.124307                                                
      ELSEIF((BI .GE. 7.5) .AND. (BI .LT. 10.5)) THEN                 
         A0 = 0.491912                                                
         A1 = 0.491833                                                
         A2 = 0.487414                                                
         A3 = 0.011371                                                
         A4 = 0.147747                                                
      ELSEIF((BI .GE. 10.5) .AND. (BI .LT. 13.5)) THEN                
         A0 = 0.564119                                                
         A1 = 0.419196                                                
         A2 = 0.639819                                                
         A3 = 0.011543                                                
         A4 = 0.149005                                                
      ELSEIF((BI .GE. 13.5) .AND. (BI .LT. 20.0)) THEN                
         A0 = 0.640669                                                
         A1 = 0.432466                                                
         A2 = 1.048056                                                
         A3 = 0.011616                                                
         A4 = 0.212726                                                
      ELSEIF((BI .GE. 20.0) .AND. (BI .LT. 62.5)) THEN                
         A0 = 0.672353                                                
         A1 = 0.397007                                                
         A2 = 1.153169                                                
         A3 = 0.011280                                                
         A4 = 0.216883                                                
      ELSEIF(BI .GE. 62.5) THEN                                       
         A0 = 0.741435                                                
         A1 = 0.448054                                                
         A2 = 1.929879                                                
         A3 = 0.010152                                                
         A4 = 0.306448                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T6(BI,A0,A1,A2,A3,A4)
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
                                
      IF(BI .LT. 2.25) THEN                                           
         A0 = -0.040800                                               
         A1 = 1.099652                                                
         A2 = 0.158995                                                
         A3 = 0.005467                                                
         A4 = 0.139116                                                
      ELSEIF((BI .GE. 2.25) .AND. (BI .LT. 7.00)) THEN                
         A0 = -0.040800                                               
         A1 = 0.982757                                                
         A2 = 0.111618                                                
         A3 = 0.008072                                                
         A4 = 0.111404                                                
      ELSEIF((BI .GE. 7.0) .AND. (BI .LT. 12.0)) THEN                 
         A0 = 0.094602                                                
         A1 = 0.754878                                                
         A2 = 0.092069                                                
         A3 = 0.009877                                                
         A4 = 0.090763                                                
      ELSEIF((BI .GE. 12.0) .AND. (BI .LT. 19.5)) THEN                
         A0 = 0.023000                                                
         A1 = 0.802068                                                
         A2 = 0.057545                                                
         A3 = 0.009662                                                
         A4 = 0.084532                                                
      ELSEIF((BI .GE. 19.5) .AND. (BI .LT. 62.5)) THEN                
         A0 = 0.02300                                                 
         A1 = 0.793673                                                
         A2 = 0.039324                                                
         A3 = 0.009326                                                
         A4 = 0.082751                                                
      ELSEIF(BI .GE. 62.5) THEN                                       
         A0 = 0.529213                                                
         A1 = 0.291801                                                
         A2 = 0.082428                                                
         A3 = 0.008317                                                
         A4 = 0.075461                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T7(BI,A0,A1,A2,A3,A4)
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
                                
      IF (BI .LT. 1.25) THEN                                          
         A0 = 0.352536                                                
         A1 = 0.692114                                                
         A2 = 0.263134                                                
         A3 = 0.005482                                                
         A4 = 0.121775                                                
      ELSEIF((BI .GE. 1.25) .AND. (BI .LT. 4.0)) THEN                 
         A0 = 0.521979                                                
         A1 = 0.504220                                                
         A2 = 0.327290                                                
         A3 = 0.005612                                                
         A4 = 0.128679                                                
      ELSEIF((BI .GE. 4.0) .AND. (BI .LT. 10.0)) THEN                 
         A0 = 0.676253                                                
         A1 = 0.334583                                                
         A2 = 0.482297                                                
         A3 = 0.005898                                                
         A4 = 0.138946                                                
      ELSEIF((BI .GE.10.0) .AND. (BI .LT. 32.0)) THEN                 
         A0 = 0.769531                                                
         A1 = 0.259497                                                
         A2 = 0.774068                                                
         A3 = 0.005600                                                
         A4 = 0.165513                                                
      ELSEIF((BI .GE. 32.0) .AND. (BI .LT. 75.0)) THEN                
         A0 = 0.849057                                                
         A1 = 0.215799                                                
         A2 = 1.343183                                                
         A3 = 0.004725                                                
         A4 = 0.223759                                                
      ELSEIF(BI .GE. 75.0) THEN                                       
         A0 = 0.831231                                                
         A1 = 0.227304                                                
         A2 = 1.174756                                                
         A3 = 0.004961                                                
         A4 = 0.212109                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T8(BI,A0,A1,A2,A3,A4) 
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
                               
      IF(BI .LT. 2.25) THEN                                           
         A0 = 0.575024                                                
         A1 = 0.449062                                                
         A2 = 0.278452                                                
         A3 = 0.004122                                                
         A4 = 0.121682                                                
      ELSEIF((BI .GE. 2.25) .AND. (BI .LT. 8.0)) THEN                 
         A0 = 0.715269                                                
         A1 = 0.307172                                                
         A2 = 0.442104                                                
         A3 = 0.004371                                                
         A4 = 0.138351                                                
      ELSEIF((BI .GE. 8.0) .AND. (BI .LT. 18.5)) THEN                 
         A0 = 0.787940                                                
         A1 = 0.243548                                                
         A2 = 0.661599                                                
         A3 = 0.004403                                                
         A4 = 0.162595                                                
      ELSEIF((BI .GE. 18.5) .AND. (BI .LT. 62.5)) THEN                
         A0 = 0.829492                                                
         A1 = 0.204078                                                
         A2 = 0.784529                                                
         A3 = 0.004050                                                
         A4 = 0.179003                                                
      ELSEIF(BI .GE. 62.5) THEN                                       
         A0 = 0.847012                                                
         A1 = 0.190678                                                
         A2 = 0.931686                                                
         A3 = 0.003849                                                
         A4 = 0.183239                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T9(BI,A0,A1,A2,A3,A4) 
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
                               
      IF(BI .LT. 2.25) THEN                                           
         A0 = 0.708905                                                
         A1 = 0.314101                                                
         A2 = 0.357499                                                
         A3 = 0.003276                                                
         A4 = 0.119300                                                
      ELSEIF((BI .GE. 2.25) .AND. (BI .LT. 9.0)) THEN                 
         A0 = 0.784576                                                
         A1 = 0.239663                                                
         A2 = 0.484422                                                
         A3 = 0.003206                                                
         A4 = 0.134987                                                
      ELSEIF((BI .GE. 9.0) .AND. (BI .LT. 57.0)) THEN                 
         A0 = 0.839439                                                
         A1 = 0.188966                                                
         A2 = 0.648124                                                
         A3 = 0.003006                                                
         A4 = 0.157697                                                
      ELSEIF( BI .GE. 57.0) THEN                                      
         A0 = 0.882747                                                
         A1 = 0.146229                                                
         A2 = 0.807987                                                
         A3 = 0.002537                                                
         A4 = 0.174543                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             
C                                                                     
C                                                                     
C                                                                     
      SUBROUTINE T10(BI,A0,A1,A2,A3,A4) 
      IMPLICIT NONE
      DOUBLE PRECISION  BI,A0,A1,A2,A3,A4                              
                              
      IF(BI .LT. 2.25) THEN                                           
         A0 = 0.865453                                                
         A1 = 0.157618                                                
         A2 = 0.444973                                                
         A3 = 0.001650                                                
         A4 = 0.148084                                                
      ELSEIF((BI .GE. 2.25) .AND. (BI .LT. 10.0)) THEN                
         A0 = 0.854768                                                
         A1 = 0.171434                                                
         A2 = 0.495042                                                
         A3 = 0.001910                                                
         A4 = 0.142251                                                
      ELSEIF((BI .GE. 10.0) .AND. (BI .LT. 58.0)) THEN                
         A0 = 0.866180                                                
         A1 = 0.163992                                                
         A2 = 0.573946                                                
         A3 = 0.001987                                                
         A4 = 0.157594                                                
      ELSEIF(BI .GE. 58.0) THEN                                       
         A0 = 0.893192                                                
         A1 = 0.133039                                                
         A2 = 0.624100                                                
         A3 = 0.001740                                                
         A4 = 0.164248                                                
      ENDIF                                                           
      RETURN                                                          
      END                                                             

