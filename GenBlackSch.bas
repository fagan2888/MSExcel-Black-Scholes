Attribute VB_Name = "GenBlackSch"
'
'
'   Copyright (C) 2005  Prof. Jayanth R. Varma, jrvarma@iimahd.ernet.in,
'   Indian Institute of Management, Ahmedabad 380 015, INDIA
'
'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program (see file COPYING); if not, write to the
'   Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'   Boston, MA  02111-1307  USA
'
'This defines a number of functions related to the Black-Scholes
'option pricing formula. This includes Black-Scholes call and put
'prices, Black-Scholes call and put implied volatilities and the various
'option greeks - delta, gamma, vega, theta and rho
'
'With minor changes, these functions could also be used in a
'Standalone Visual Basic application outside of Excel.
'For the normal integral (cumulative distribution function)
'we use the function NormalCDF from the module CumNormal
'This uses Application.NormSDist which is available only in Excel
'For standalone visual basic applications the function NormalCDF
'in the module CumNormal must be changed to use one of the other
'normal cdf functions available in that module
'
'The functions defined here can be used within Excel formulas
'These functions are also used by the BlackSchForm which
'provides an interactive facility for Black-Scholes option
'valuation, implied volatilities and option greeks.
'This module includes some functions that are intended
'only for use in this form
'
'
'The constants below are used in intermediate calculations to represent quantities
'that arise from division by zero or log of zero
Const vHigh = 100
Const vvHigh = 1E+30
'The following type is used by the routines for Implied Volatility
Type DoubleWithStatusString
    Value As Double
    Status As String
End Type
'
'
Sub GBlackScholes()
'This is for use within Excel as part of the form BlackSchForm
GBlackSchForm.Show
End Sub
'
'
Private Function GSafeD1(s As Double, x As Double, r As Double, _
                        Sigma As Double, t As Double, q As Double) As Double
'This computes the BlackScholes quantity d1 safely i.e.
'no division by zero and no log of zero
If (Sigma = 0 Or t = 0) Then
    S0 = s * Exp((r - q + Sigma ^ 2 / 2) * t)
    If (S0 > x) Then GSafeD1 = vHigh
    If (S0 < x) Then GSafeD1 = -vHigh
    If (S0 = x) Then GSafeD1 = 0
Else
    If (x = 0) Then
        GSafeD1 = vHigh
    Else
        'Below is the BlackScholes formula for d1
        GSafeD1 = (Log(s / x) + (r - q + Sigma ^ 2 / 2) * t) / (Sigma * Sqr(t))
    End If
End If
End Function
'
'
Function GBSCall(s As Double, x As Double, r As Double, _
                Sigma As Double, t As Double, q As Double) As Double
'Black Scholes call price
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
d2 = d1 - Sigma * Sqr(t)
GBSCall = s * Exp(-q * t) * NormalCDF(d1) _
         - x * Exp(-r * t) * NormalCDF(d2)
End Function
'
'
Function GBSPut(s As Double, x As Double, r As Double, _
                Sigma As Double, t As Double, q As Double) As Double
'Black Scholes put price
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
d2 = d1 - Sigma * Sqr(t)
GBSPut = -s * Exp(-q * t) * NormalCDF(-d1) _
        + x * Exp(-r * t) * NormalCDF(-d2)
End Function
'
'
Function GBSCallDelta(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes call delta
Dim d1 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
GBSCallDelta = NormalCDF(d1) * Exp(-q * t)
End Function
'
'
Function GBSPutDelta(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes put delta
Dim d1 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
GBSPutDelta = (NormalCDF(d1) - 1) * Exp(-q * t)
End Function
'
'
Function GBSCallTheta(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes call theta
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
d2 = d1 - Sigma * Sqr(t)
If (t = 0) Then
    If (Abs(d1) = vHigh Or Sigma = 0) Then
        a = 0
    Else
        a = -vvHigh
    End If
Else
    a = -s * Exp(-q * t) * NormOrdinate(d1) * Sigma / (2 * Sqr(t))
End If
b = r * x * Exp(-r * t) * NormalCDF(d2)
c = q * s * Exp(-q * t) * NormalCDF(d1)
GBSCallTheta = a - b + c
End Function
'
'
Function GBSPutTheta(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes put theta
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
d2 = d1 - Sigma * Sqr(t)
If (t = 0) Then
    If (Abs(d1) = vHigh Or Sigma = 0) Then
        a = 0
    Else
        a = -vvHigh
    End If
Else
    a = -s * Exp(-q * t) * NormOrdinate(d1) * Sigma / (2 * Sqr(t))
End If
b = r * x * Exp(-r * t) * NormalCDF(-d2)
c = q * s * Exp(-q * t) * NormalCDF(-d1)
GBSPutTheta = a + b - c
End Function
'
'
Function GBSCallGamma(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes call gamma
Dim d1 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
If (Sigma = 0 Or t = 0) Then
    If (Abs(d1) = vHigh) Then
        GBSCallGamma = 0
    Else
        GBSCallGamma = vvHigh
    End If
Else
    GBSCallGamma = NormOrdinate(d1) * Exp(-q * t) / (s * Sigma * Sqr(t))
End If
End Function
'
'
Function GBSPutGamma(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes put gamma (same as call gamma)
GBSPutGamma = GBSCallGamma(s, x, r, Sigma, t, q)
End Function
'
'
Function GBSCallVega(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes call vega
Dim d1 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
GBSCallVega = s * Exp(-q * t) * Sqr(t) * NormOrdinate(d1)
End Function
'
'
Function GBSPutVega(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'put and call vegas are the same
GBSPutVega = GBSCallVega(s, x, r, Sigma, t, q)
End Function
'
'
Function GBSCallRho(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double, _
                    Optional OnFutures As Boolean = False) As Double
'Black Scholes call rho
Dim d1 As Double
Dim d2 As Double
If OnFutures Then
    GBSCallRho = -t * GBSCall(s, x, r, Sigma, t, q)
Else
    d1 = GSafeD1(s, x, r, Sigma, t, q)
    d2 = d1 - Sigma * Sqr(t)
    GBSCallRho = x * t * Exp(-r * t) * NormalCDF(d2)
End If
End Function
'
'
Function GBSPutRho(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double, _
                    Optional OnFutures As Boolean = False) As Double
'Black Scholes put rho
Dim d1 As Double
Dim d2 As Double
If OnFutures Then
    GBSPutRho = -t * GBSPut(s, x, r, Sigma, t, q)
Else
    d1 = GSafeD1(s, x, r, Sigma, t, q)
    d2 = d1 - Sigma * Sqr(t)
    GBSPutRho = -x * t * Exp(-r * t) * NormalCDF(-d2)
End If
End Function
'
'
Function GBSCallRhoForeign(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes call rho foreign
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
GBSCallRhoForeign = -s * t * Exp(-q * t) * NormalCDF(d1)
End Function
'
'
Function GBSPutRhoForeign(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes put rho foreign
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
GBSPutRhoForeign = s * t * Exp(-q * t) * NormalCDF(-d1)
End Function
'
'
Function GBSCallProb(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes risk neutral probability of put exercise
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
d2 = d1 - Sigma * Sqr(t)
GBSCallProb = NormalCDF(d2)
End Function
'
'
Function GBSPutProb(s As Double, x As Double, r As Double, _
                    Sigma As Double, t As Double, q As Double) As Double
'Black Scholes risk neutral probability of put exercise
Dim d1 As Double
Dim d2 As Double
d1 = GSafeD1(s, x, r, Sigma, t, q)
d2 = d1 - Sigma * Sqr(t)
GBSPutProb = NormalCDF(-d2)
End Function
'
'
Function GBSCallImplied(s As Double, x As Double, r As Double, _
                    price As Double, t As Double, q As Double, _
                    Optional ErrType As String = "brief")
'Black Scholes Implied Volatility from Call Price
GBSCallImplied = GBSImplied(s, x, r, price, t, q, False, ErrType)
End Function
'
'
Function GBSPutImplied(s As Double, x As Double, r As Double, _
                    price As Double, t As Double, q As Double, _
                    Optional ErrType As String = "brief")
'Black Scholes Implied Volatility from Put Price
GBSPutImplied = GBSImplied(s, x, r, price, t, q, True, ErrType)
End Function
'
'
Function GBSImplied(s As Double, x As Double, r As Double, _
                   price As Double, t As Double, q As Double, PutOpt As Boolean, _
                   Optional ErrType As String = "brief")
'Black Scholes Implied Volatility from Call or Put Price
Dim SplitLines As Boolean, Temp As DoubleWithStatusString
Temp = GBSImplied_0(s, x, r, price, t, q, PutOpt)
If (Temp.Status = "success") Then
    GBSImplied = Temp.Value
End If
If (Temp.Status = "undefined") Then
    GBSImplied = Temp.Status
End If
If (Temp.Status = "error") Then
    If (ErrType = "brief") Then
        GBSImplied = Temp.Status
    End If
    SplitLines = (ErrType = "detailed multi")
    If (ErrType = "detailed" Or ErrType = "detailed multi") Then
        GBSImplied = GBSImpliedError(s, x, r, Temp.Value, t, q, _
                    price, PutOpt, SplitLines)
    End If
End If
End Function
'
'
Function GBSCallImplied_0(s As Double, x As Double, r As Double, _
                    price As Double, t As Double, q As Double _
                    ) As DoubleWithStatusString
'interface routine to GBSImplied_0
GBSCallImplied_0 = GBSImplied_0(s, x, r, price, t, q, False)
End Function
'
'
Function GBSPutImplied_0(s As Double, x As Double, r As Double, _
                    price As Double, t As Double, q As Double _
                    ) As DoubleWithStatusString
'interface routine to GBSImplied_0
GBSPutImplied_0 = GBSImplied_0(s, x, r, price, t, q, True)
End Function
'
'
Function GBSImplied_0(S0 As Double, X0 As Double, r As Double, _
                    price0 As Double, t As Double, q As Double, _
                    PutOpt As Boolean _
                    ) As DoubleWithStatusString
'Computes implied volatility from call price (if PutOpt is false)
'or from put price (if PutOpt is true)
'
'This function returns a DoubleWithStatusString in which
'Value contains the estimated implied volatility
'Status contains the status of the estimate:
'"undefined" if the implied is undefined
'"success" if the iterations converges
'"error" if iterations do not converge
'If Status is not "success",  then caller must
'use a function like GBSCallImpliedError to report a
'more meaningful error status
Dim Pi As Double, Root2Pi As Double, s As Double, _
    SplusX As Double, SminusX As Double, Radical As Double, _
    SigmaRootT As Double, AbsErr As Double, RelErr As Double, _
    Iter As Integer, Vega As Double, Step As Double, _
    Sigma As Double, OldErr As Double, Factor As Double, _
    TrySigma As Double, price As Double, LineCount As Integer, _
    x As Double, PredImpr As Double, H As Double, Temp As Double
Const Eps = 0.000001, Eta = 0.000001, MaxIter = 20, _
      zero As Double = 0, MaxSigma As Double = 100
On Error GoTo ErrHndlr
'discount the exercise price to eliminate r
x = X0 * Exp(-r * t)
'similarly eliminate q
s = S0 * Exp(-q * t)
'use put call parity to convert put option into call option
If PutOpt Then
    price = price0 + s - x
Else
    price = price0
End If
SminusX = s - x
SplusX = s + x
'price must be at least intrinsic value (Max(SminusX,0)) and cannot exceed S
If price < SminusX Or price < 0 Or price > s Then
    GBSImplied_0.Value = 0
    GBSImplied_0.Status = "undefined"
    Exit Function
End If
If price = SminusX Or price = 0 Then
'if price equals intrinsic value, volatility is zero
    GBSImplied_0.Value = zero
    GBSImplied_0.Status = "success"
    Exit Function
End If
If x = 0 Then ' and price <> S is implicit here
'if x is 0, option is same as stock
    GBSImplied_0.Value = 0
    GBSImplied_0.Status = "undefined"
    Exit Function
End If
'We use an approximate value of sigma to start the
'Newton-Raphson iterations
Sigma = GBSImpliedApprox(s, X0, r, price0, t, q, PutOpt)
If GBSCallVega(s, x, 0, Sigma, t, 0) = 0 Then
'Newton-Raphson iterations cannot proceed if vega is zero
'So we choose a starting point where vega is likely to be high
    Sigma = Sqr(2 * Abs(Log(s / x))) / Sqr(t)
    ' the point of maximum vega is where d1 is close to 0
    ' d1 = a/S + S/2 where a = ln(S*exp(rt)/x)
    ' if a < 0 then the above value of S sets d1 to 0
    ' else it sets it to its min value of root a/2
End If
If GBSCallVega(s, x, 0, Sigma, t, 0) = 0 And s > x And price > SminusX Then
    GBSImplied_0.Value = Sigma
    GBSImplied_0.Status = "error"
    Exit Function
End If
'Start Newton-Raphson Iterations
Iter = 0
LineCount = 0
OldErr = price - GBSCall(s, x, 0, Sigma, t, 0)
'The first "step" is a zero step
Step = 0
Factor = 0
PredImpr = 0
Do
    Do
    'In this loop we reduce the step size if necessary to ensure
    'that the actual change in the price is not too different from
    'what is predicted by the vega
        LineCount = LineCount + 1
        TrySigma = Sigma + Step * Factor
        AbsErr = price - GBSCall(s, x, 0, TrySigma, t, 0)
        Vega = GBSCallVega(s, x, 0, TrySigma, t, 0)
        Factor = Factor / 2
        If LineCount > 10 Then
            GBSImplied_0.Value = Sigma
            GBSImplied_0.Status = "error"
            Exit Function
        End If
    Loop While Vega = 0 _
                Or Abs(AbsErr) - Abs(OldErr) > 0.5 * PredImpr * Factor
    Sigma = TrySigma
    RelErr = AbsErr / price0
    Iter = Iter + 1
    ' do not permit a step exceeding MaxSigma
    If Abs(AbsErr) > MaxSigma * Abs(Vega) Then
        Step = Sgn(AbsErr) * MaxSigma / Sgn(Vega)
    Else
        'This is the Newton step
        Step = AbsErr / Vega
    End If
    'do not permit sigma to go negative
    Step = Application.max(Step, -0.99 * Sigma)
    OldErr = AbsErr
    PredImpr = Abs(Step * Vega)
    Factor = 1
    'the termination condition is
    'a low absolute error
    'or a low relative error
    'or non convergence within MaxIter iterations
Loop While (Abs(AbsErr) > Eps And Abs(RelErr) > Eta And Iter < MaxIter)
GBSImplied_0.Value = Sigma
If Abs(AbsErr) > Eps And Abs(RelErr) > Eta Then
    GBSImplied_0.Status = "error"
Else
    GBSImplied_0.Status = "success"
End If
Exit Function
ErrHndlr:
    GBSImplied_0.Value = Sigma
    GBSImplied_0.Status = "error"
End Function
'
'
Function GBSImpliedApprox(S0 As Double, X0 As Double, r As Double, _
                    price0 As Double, t As Double, q As Double, _
                    PutOpt As Boolean) As Double
'Compute an approximate implied volatility
'This approximation can be useful in its own right for
'at or near money options
'In this module, it is used as the starting point for
'Newton Raphson iterations in function GBSImplied_0
Dim Pi As Double, Root2Pi As Double, _
    SplusX As Double, SminusX As Double, Radical As Double, _
    SigmaRootT As Double, price As Double, _
    x As Double, H As Double, Temp As Double
'discount the exercise price to eliminate r
x = X0 * Exp(-r * t)
'use put call parity to convert put option into call option
'similarly eliminate q
s = S0 * Exp(-q * t)
If PutOpt Then
    price = price0 + s - x
Else
    price = price0
End If
SminusX = s - x
SplusX = s + x
'We use a Taylor series approximation similar to
'Brenner, H. and Subrahmanyam, M. G. (1994), "A simple approximation to option valuation
'and hedging in the Black Scholes model", Financial Analysts Journal, Mar-Apr 1994, 25-28
'Basically we eliminate r by discounting the exercise price as above and then
'approximate log(S/x) as 2(S-x)/(s+x) = SminusX/H
'where SminusX = S-x and H=(s+x)/2.
'We then approximate the normal integral by a Taylor series
'If price is the call price, this gives the approximation
'price/H = SigmaRootT(1/Root2Pi +SminusX/(2*H*SigmaRootT)
'                           +SminusX^2/(2*H^2*SigmaRootT^2*Root2Pi)
'Letting z=SigmaRootT*H/Root2Pi, this yields a quadratic equation for z:
'z^2 - (price-SminusX/2)*z +SminusX^2/(4*Pi) = 0
'The solution is
'z = (price - SminusX/2)/2 + Radical
'where Radical is the square root of (price-SminusX/2)^2 - SminusX^2/Pi
'The linear approximation is obtained by dropping the constant to give the linear equation:
'z = (price-SminusX/2)
'
Pi = Application.Pi
'pi = 3.141592653589793
Root2Pi = Sqr(2 * Pi)
H = 0.5 * SplusX
Temp = (price - 0.5 * SminusX)
Radical = Temp ^ 2 - SminusX ^ 2 / Pi
If Radical < 0 Then
    'Try Linear Approximation
    SigmaRootT = (Root2Pi / H) * Temp
Else
    'Try Quadratic Approximation
    Radical = Sqr(Radical)
    SigmaRootT = (Root2Pi / H) * (Temp / 2 + Radical)
End If
GBSImpliedApprox = SigmaRootT / Sqr(t)
End Function
'
'
Function GBSImpliedError(s As Double, x As Double, _
                    r As Double, Sigma As Double, t As Double, q As Double, _
                    price As Double, PutOpt As Boolean, _
                    Optional SplitLines As Boolean = False _
                    ) As String
'Interface to Call/Put Implied Error
'Given a possibly incorrect implied volatility, return the
'magnitude of the error as a string
If (PutOpt) Then
    GBSImpliedError = GBSPutImpliedError(s, x, r, Sigma, t, _
                     price, q, SplitLines)
Else
    GBSImpliedError = GBSCallImpliedError(s, x, r, Sigma, t, _
                     price, q, SplitLines)
End If
End Function
'
'
Function GBSCallImpliedError(s As Double, x As Double, _
                    r As Double, Sigma As Double, t As Double, _
                    price As Double, q As Double, _
                    Optional SplitLines As Boolean = False _
                    ) As String
'Given a possibly incorrect implied volatility, return the
'magnitude of the error as a string
Dim aprice As Double, perror As Double, Separator As String
aprice = GBSCall(s, x, r, Sigma, t, q)
perror = aprice - price
If (SplitLines) Then
    Separator = Chr$(10) & Chr$(10)
Else
    Separator = " "
End If
GBSCallImpliedError = MultiStrings(Separator, False, _
                     "Implied of", Sigma * 100, _
                     "gives price of", aprice, _
                     "error =", perror)
End Function
'
'
Function GBSPutImpliedError(s As Double, x As Double, _
                    r As Double, Sigma As Double, t As Double, _
                    price As Double, q As Double, _
                    Optional SplitLines As Boolean = False _
                    ) As String
'Given a possibly incorrect implied volatility, return the
'magnitude of the error as a string
Dim aprice As Double, perror As Double, Separator As String
aprice = GBSPut(s, x, 0, Sigma, t, q)
perror = aprice - price
If (SplitLines) Then
    Separator = Chr$(10) & Chr$(10)
Else
    Separator = " "
End If
GBSPutImpliedError = MultiStrings(Separator, False, _
                    "Implied of ", Sigma * 100, _
                    "gives price of", aprice, _
                    "error =", perror)
End Function
'
'
Private Function NormOrdinate(z As Double) As Double
'The normal ordinate (probability density function)
NormOrdinate = Exp(-0.5 * z * z) / Sqr(2 * Application.Pi)
End Function
Function GBSD1(s As Double, x As Double, r As Double, _
                        Sigma As Double, t As Double, q As Double) As Double
GBSD1 = GSafeD1(s, x, r, Sigma, t, q)
End Function
Function GBSD2(s As Double, x As Double, r As Double, _
                        Sigma As Double, t As Double, q As Double) As Double
GBSD2 = GSafeD1(s, x, r, Sigma, t, q) - Sigma * Sqr(t)
End Function




