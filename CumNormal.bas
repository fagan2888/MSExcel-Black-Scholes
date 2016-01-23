Attribute VB_Name = "CumNormal"
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
'
'This computes the normal integral (cumulative distribution function
'of the standard normal distribution

Function NormalCDF(x As Double) As Double
'Within Excel, the NormSDist function is available, and so we
'can simply call it
NormalCDF = Application.NormSDist(x)
'Otherwise comment out the above line and uncomment one of the
'following two lines
'NormalCDF = NormalCDF_Marsaglia(x)
'NormalCDF = NormalCDF_Fike(x)
End Function
'
'
Function NormalCDF_Marsaglia(x As Double) As Double
' Converted into  Visual Basic from C code in
' George Marsaglia, Evaluating the Normal Distribution,
' Journal of Statistical Software, July 2004, Volume 11, Issue 4.
' http://www.jstatsoft.org/
Dim Sum As Double, q As Double, i As Double, _
    summand As Double, tolerance As Double, half_ln_twopi As Double
Sum = x
summand = x
q = x * x
i = 1
tolerance = 0.0000000001
half_ln_twopi = 0.918938533204673
If Abs(x) < 15 Then
' this test is not there in Marsaglia, but is needed
' to prevent overflow when x is 100 or so
    Do
        i = i + 2
        summand = summand * q / i
        Sum = Sum + summand
    Loop Until summand < tolerance
    NormalCDF_Marsaglia = 0.5 + Sum * Exp(-0.5 * q - half_ln_twopi)
Else
    If x < 0 Then
        NormalCDF_Marsaglia = 0
    Else
        NormalCDF_Marsaglia = 1
    End If
End If
End Function
'
'
Function NormalCDF_Fike(y As Double) As Double
'
'Computes normal integral by using a rational polynomial approximation
'This approximation is from Example 9.7.3 of
'Fike, C.T. (1968), Computer Evaluation of Mathematical Functions
'Englewood Cliffs, N.J., Prentice Hall
'Let P(x) be the integral of the normal density from 0 to x. Then,
'the best minimax approximation R(x) to P(x) in the range [0,infinity)
'among the class of rational functions V5,5[0,infinity) satisfying
'R(0) = P(0) = 0, and
'lim x tends to infinity R(x) = lim x tends to infinity P(x) = 0.5
'is the function:
'      a1 + a2*x + a3*x^2 + a4*x^3  + a5*x^4  + a6*x^5
'    ----------------------------------------------------
'      b1 + b2*x + b3*x^2 + b4*x^3  + b5*x^4  + b6*x^5
'where the constants a1, a2, ..., a6 and b1, b2, ..., b6
'are as defined below
'The maximum absolute error of this approximation is 0.46x10^-4
'i.e. 0.000046. Therefore, this approximation has the same accuracy
'as the 4 place tables commonly found in statistics test books
'
Const a1 = 0
Const a2 = 9.050508
Const a3 = 0.767742
Const a4 = 1.666902
Const a5 = -0.624298
Const a6 = 0.5
Const b1 = 22.601228
Const B2 = 2.776898
Const b3 = 5.148169
Const b4 = 2.995582
Const b5 = -1.238661
Const b6 = 1
'
'We now compute R(abs(y)) as an approximation to P(abs(y))
'
x = Abs(y)
'It may be better to test for large x and set temp = 0.5
' for x > 7 or so
temp1 = ((((a6 * x + a5) * x + a4) * x + a3) * x + a2) * x + a1
temp2 = ((((b6 * x + b5) * x + b4) * x + b3) * x + B2) * x + b1
Temp = temp1 / temp2
'
'We now compute the normal integral N(y) from P(abs(y))
'
If (y < 0) Then
    NormalCDF_Fike = 0.5 - Temp
Else
    NormalCDF_Fike = 0.5 + Temp
End If
End Function


