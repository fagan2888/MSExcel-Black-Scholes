VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoldenSearchForm 
   Caption         =   "Golden Section Search to Find Max/Min of a Function"
   ClientHeight    =   4170
   ClientLeft      =   1050
   ClientTop       =   330
   ClientWidth     =   8010
   OleObjectBlob   =   "GoldenSearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoldenSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'
'   Copyright (C) 2001  Prof. Jayanth R. Varma, jrvarma@iimahd.ernet.in,
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
'This uses the Golden Section method to find the min or max of a function
'
'
'Global Variables
'xCell and fCell are the cells containing x and f(x)
Dim xCell As Range, fCell As Range
'a and b define the range of x ( a <= x <= b)
'Eta and Eps are tolerance parameters
Dim a As Double, b As Double, Eta As Double, Eps As Double
'min is True for Minimisation, false for Maximisation
Dim min As Boolean
'background color of the labels
Dim LabelColor
'
'
Private Function f(x As Double) As Double
'This converts the spreadsheet into a function f(x)
'The input value is entered into the spreadsheet (xCell)
'and the value f(x) is read from the spreadsheet (fCell)
'Internally, everything is a minimisation problem
'We maximise f(x) by minimising -f(x)
On Error GoTo function_error
xCell.Value = x
If min Then f = fCell.Value Else f = -fCell.Value
Exit Function
function_error:
MsgBox "Error while trying to evaluate function for x = " & _
       Str$(x)
Dim errno As Integer
errno = Err.Number
Err.Raise errno
'Unload Me
End Function
'
'
Private Sub solve()
Dim fa As Double, fb As Double, fm1 As Double, fm2 As Double
Dim m1 As Double, m2 As Double, z As Double, fz As Double
Dim r As Double, phi As Double, rphi As Double
Dim currmin As Double, currmax As Double, a0 As Double, b0 As Double
Static xLabel As Object, fxLabel As Object
xaL.Caption = ""
x1L.Caption = ""
x2L.Caption = ""
xbL.Caption = ""
fxaL.Caption = ""
fx1L.Caption = ""
fx2L.Caption = ""
fxbL.Caption = ""
xBoundL.Caption = ""
fxBoundL.Caption = ""
'xLabel and fxLabel are the cells containing the minimum values
'at the last invocation of this form. These labels are
'shaded yellow
If (xLabel Is Nothing) Then
'Store the background color
    LabelColor = xaL.BackColor
Else
'If this is not the first invocation, we must change the yellow
'shade back to the normal background colour
    xLabel.BackColor = LabelColor
    fxLabel.BackColor = LabelColor
End If
'Golden search will change a and b
'We store the original values as a0 and b0
a0 = a
b0 = b
r = b - a
phi = (Sqr(5) - 1) / 2
rphi = r * phi
m1 = b - rphi
m2 = a + rphi
'These are the four points in golden search: a, m2, m1, b
'Find f(x) at these points
fa = f(a)
fb = f(b)
fm1 = f(m1)
fm2 = f(m2)
'Find current error bound on f(x)
currmin = Application.min(fa, fb, fm1, fm2)
currmax = Application.max(fa, fb, fm1, fm2)
'The termination condition is that the error bounds on either x or f(x)
'is below the tolerance
While ((currmax - currmin) > Eta And Abs(a - b) > Eps)
    'display current error bounds
    xBoundL.Caption = myFormat(Abs(a - b))
    fxBoundL.Caption = myFormat(currmax - currmin)
    Me.Repaint
    'find new a, m2, m1, b
    rphi = rphi * phi
    If (fm1 < fm2) Then
        b = m2
        m2 = m1
        m1 = b - rphi
        fb = fm2
        fm2 = fm1
        fm1 = f(m1)
    Else
        a = m1
        m1 = m2
        m2 = a + rphi
        fa = fm1
        fm1 = fm2
        fm2 = f(m2)
    End If
    'find new error bound on f(x)
    currmin = Application.min(fa, fb, fm1, fm2)
    currmax = Application.max(fa, fb, fm1, fm2)
Wend
'The minimum value from fa, fb, fm1, fm2 is now determined
'and the appropriate x value and the labels are found
z = a
fz = fa
Set xLabel = xaL
Set fxLabel = fxaL
If (fm1 < fz) Then
    z = m1
    fz = fm1
    Set xLabel = x1L
    Set fxLabel = fx1L
End If
If (fm2 < fz) Then
    z = m2
    fz = fm2
    Set xLabel = x2L
    Set fxLabel = fx2L
End If
If (fb < fz) Then
    z = b
    fz = fb
    Set xLabel = xbL
    Set fxLabel = fxbL
End If
fz = f(z)
'display the x values
xaL.Caption = myFormat(a)
x1L.Caption = myFormat(m1)
x2L.Caption = myFormat(m2)
xbL.Caption = myFormat(b)
'display the f(x) values
If min Then
    fxaL.Caption = myFormat(fa)
    fx1L.Caption = myFormat(fm1)
    fx2L.Caption = myFormat(fm2)
    fxbL.Caption = myFormat(fb)
Else
    fxaL.Caption = myFormat(-fa)
    fx1L.Caption = myFormat(-fm1)
    fx2L.Caption = myFormat(-fm2)
    fxbL.Caption = myFormat(-fb)
End If
'display the bounds
xBoundL.Caption = myFormat(Abs(a - b))
fxBoundL.Caption = myFormat(currmax - currmin)
'shade the solution yellow
xLabel.BackColor = RGB(255, 255, 0)
fxLabel.BackColor = RGB(255, 255, 0)
a = a0
b = b0
End Sub
'
'
Private Sub CloseButton_Click()
Unload Me
End Sub
'
'
Private Sub SolveButton_Click()
Set xCell = Range(xCellRE.Value)
Set fCell = Range(fCellRE.Value)
a = LowTB.Value
b = UpTB.Value
Eta = Application.max(Abs(etaTB.Value), 0.00000001)
Eps = Application.max(Abs(epsTB.Value), 0.00000001)
min = MinOptButt.Value
etaTB.Value = Eta
epsTB.Value = Eps
'We store the current values of all user parameters as the
'new defaults for this Excel session
'This function is in the LineSearch module to allow static
'variables to be used
Call GoldenSearchDefaults(xCell, fCell, a, b, Eta, Eps, min, True)
'Solve for min/max
On Error GoTo myerror
Call solve
Exit Sub
myerror:
MsgBox "There was an error"
Unload Me
End Sub
'
'
Private Sub UserForm_Initialize()
'Read default values. In case this form has been invoked earlier
'in this Excel session, the values are remembered from that invocation
Call GoldenSearchDefaults(xCell, fCell, a, b, Eta, _
                        Eps, min, False)
'xCell and fCell are also remembered from earlier
'invocation if any. Else these are all empty
Call ValidateRange(xCell)
Call ValidateRange(fCell)
xCellRE.Value = RangeAddress(xCell, "")
fCellRE.Value = RangeAddress(fCell, "")
LowTB.Value = a
UpTB.Value = b
etaTB.Value = Eta
epsTB.Value = Eps
MinOptButt.Value = min
MaxOptButt.Value = Not min
End Sub
