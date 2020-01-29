      MODULE SOLVE
	
	IMPLICIT NONE

	CONTAINS

      SUBROUTINE FUNCD(X,F,DF,FLAG)
      
      INTEGER FLAG
      REAL*8 X,VAR,F,DF,ALPHA
      COMMON/NEWT/VAR,ALPHA

        F = VAR-(2/X)+(2.0/(X**2.0))*(1-DEXP(-X))
        DF = 2/X**2+2*(DEXP(-X))/X**2-4*(1-DEXP(-X))/X**3

              
      END SUBROUTINE FUNCD


      FUNCTION rtsafe(funcd,x1,x2,xacc,p)
      INTEGER MAXIT
      REAL*8 rtsafe,x1,x2,xacc
      EXTERNAL funcd
      PARAMETER (MAXIT=1000)
      INTEGER j,p
      REAL*8 df,dx,dxold,f,fh,fl,temp,xh,xl
      call funcd(x1,fl,df,p)
      call funcd(x2,fh,df,p)
      if((fl.gt.0..and.fh.gt.0.).or.(fl.lt.0..and.fh.lt.0.))
	pause 'root must be bracketed in rtsafe'
      if(fl.eq.0.)then
        rtsafe=x1
        return
      else if(fh.eq.0.)then
        rtsafe=x2
        return
      else if(fl.lt.0.)then
        xl=x1
        xh=x2
      else
        xh=x1
        xl=x2
      endif
      rt=.5*(x1+x2)
      dxold=abs(x2-x1)
      dx=dxold
      call funcd(rt,f,df,p)
      do 11 j=1,MAXIT
        if(((rtsafe-xh)*df-f)*((rtsafe-xl)*df-f).ge.0..or. abs(2.*
     *f).gt.abs(dxold*df) ) then
          dxold=dx
          dx=0.5*(xh-xl)
          rtsafe=xl+dx
          if(xl.eq.rtsafe)return
        else
          dxold=dx
          dx=f/df
          temp=rtsafe
          rtsafe=rtsafe-dx
          if(temp.eq.rtsafe)return
        endif
        if(abs(dx).lt.xacc) return
        call funcd(rtsafe,f,df,p)
        if(f.lt.0.) then
          xl=rtsafe
        else
          xh=rtsafe
        endif
11    continue
      pause 'rtsafe exceeding maximum iterations'
      return
      END
C  (C) Copr. 1986-92 Numerical Recipes Software 41.921'L3.
	
      END MODULE SOLVE