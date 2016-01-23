VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GBlackSchForm 
   Caption         =   "Generalized Black Scholes Option Price, Volatility, Greeks"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   OleObjectBlob   =   "GBlackSchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GBlackSchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



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
'This form provides the user interface to the BlackSch module
'which defines a number of functions related to the Black-Scholes
'option pricing formula. This includes Black-Scholes call and put
'prices, Black-Scholes call and put implied volatilities and the various
'option greeks - delta, gamma, vega, theta and rho
'
'
'Global Variables
's, x, r, sigma, t are the parameters of the Black Scholes formula
'namely, the tock price, the exercise price, the risk free interest rate
'the volatility and the time to maturity respectively
'If implied is true,we calculate implied volatility from price
'else we calculate price from the given volatility
'SigmaoOrPrice is the input volatility or price
Dim s As Double, x As Double, r0 As Double, q0 As Double, _
    Sigma As Double, t As Double, Implied As Boolean, _
    SigmaOrPrice As Double, Futures As Boolean, _
    FX As Boolean, Simple As Boolean, Initialized As Boolean, _
    FirstRow As String, LastRow As String
'
'
Private Function val(x) As Double
'Non numeric input is converted into -1 which is an invalid value
'for all the Black Scholes parameters. r is the only parameter
'that can be negative and even that must exceed -1 (-100%)
If (IsNumeric(x)) Then
    val = x
Else
    If (x = "") Then
        val = 0
    Else
        val = -1
    End If
End If
End Function
'
'
Private Sub Compute()
'Compute and display put and call values and option greeks
'This sub is called whenever the user changes any value
Dim CallResult, PutResult, _
    CallSigma As Double, PutSigma As Double, CallFlag As String, _
    PutFlag As String, q As Double, r As Double
'We check for valid input. All parameters must be non negative
'except the risk free interest rate which can be negative but
'must exceed -1 (-100%). The routine val(x) converts non numeric
'input into -1. So the test below covers all cases of invalid input
If s > 0 And x >= 0 And SigmaOrPrice >= 0 And t >= 0 _
And r0 > -1 And q0 > -1 Then
    If Continuous.Value = True Then
            q = q0 / 100
            r = r0 / 100
    ElseIf Annual.Value = True Then
            q = Log(1 + q0 / 100)
            r = Log(1 + r0 / 100)
    ElseIf SemiAnnual.Value = True Then
            q = 2 * Log(1 + q0 / 200)
            r = 2 * Log(1 + r0 / 200)
    End If
    If Futures Then q = r
    If Simple Then q = 0
    If Implied Then
        'Calculated implied volatility from the given prices
        CallResult = GBSCallImplied(s, x, r, SigmaOrPrice, _
                     t, q, "detailed multi")
        PutResult = GBSPutImplied(s, x, r, SigmaOrPrice, _
                     t, q, "detailed multi")
        If IsNumeric(CallResult) Then
            'Set sigma so that option greeks can be computed
            CallSigma = CallResult
            'Multiply by 100 to display as %
            CallResult = CallResult * 100
        End If
        If IsNumeric(PutResult) Then
            'Set sigma so that option greeks can be computed
            PutSigma = PutResult
            'Multiply by 100 to display as %
            PutResult = PutResult * 100
        End If
    Else ' i.e. not implied
        'determine Call and Put prices from given volatility
        CallResult = GBSCall(s, x, r, SigmaOrPrice, t, q)
        PutResult = GBSPut(s, x, r, SigmaOrPrice, t, q)
        'set sigma to given value to compute option greeks
        CallSigma = SigmaOrPrice
        PutSigma = SigmaOrPrice
    End If
    If IsNumeric(CallResult) Then
        'display result (price or implied) and option greeks
        If FX Then
            CallTB.Text = MultiLineZ(CallResult, _
                            GBSCallDelta(s, x, r, CallSigma, t, q), _
                            GBSCallGamma(s, x, r, CallSigma, t, q), _
                            GBSCallVega(s, x, r, CallSigma, t, q), _
                            GBSCallRho(s, x, r, CallSigma, t, q, Futures), _
                            GBSCallTheta(s, x, r, CallSigma, t, q), _
                            GBSCallRhoForeign(s, x, r, CallSigma, t, q))
        Else
            CallTB.Text = MultiLineZ(CallResult, _
                            GBSCallDelta(s, x, r, CallSigma, t, q), _
                            GBSCallGamma(s, x, r, CallSigma, t, q), _
                            GBSCallVega(s, x, r, CallSigma, t, q), _
                            GBSCallRho(s, x, r, CallSigma, t, q, Futures), _
                            GBSCallTheta(s, x, r, CallSigma, t, q))
        End If
    Else
        'volatility is "undefined" or error
        CallTB.Text = CallResult
    End If
    If IsNumeric(PutResult) Then
        'display result (price or implied) and option greeks
        If FX Then
            PutTB.Text = MultiLineZ(PutResult, _
                            GBSPutDelta(s, x, r, PutSigma, t, q), _
                            GBSPutGamma(s, x, r, PutSigma, t, q), _
                            GBSPutVega(s, x, r, PutSigma, t, q), _
                            GBSPutRho(s, x, r, PutSigma, t, q, Futures), _
                            GBSPutTheta(s, x, r, PutSigma, t, q), _
                            GBSPutRhoForeign(s, x, r, PutSigma, t, q))
        Else
            PutTB.Text = MultiLineZ(PutResult, _
                            GBSPutDelta(s, x, r, PutSigma, t, q), _
                            GBSPutGamma(s, x, r, PutSigma, t, q), _
                            GBSPutVega(s, x, r, PutSigma, t, q), _
                            GBSPutRho(s, x, r, PutSigma, t, q, Futures), _
                            GBSPutTheta(s, x, r, PutSigma, t, q))
        End If
    Else
        'volatility is "undefined" or error
        PutTB.Text = PutResult
    End If
Else
    'non numeric input or
    's is non positive or
    'x, SigmaOrPrice, or t is negative or
    'r or q are below their bound of -1 (-100%)
    CallTB.Text = MultiLine("No", "Valid", "Input")
    PutTB.Text = MultiLine("No", "Valid", "Input")
End If
End Sub

Private Sub Annual_Click()
If Initialized Then Compute
End Sub

'
'
Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub Continuous_Click()
If Initialized Then Compute
End Sub

'
'
Private Sub ImpliedToggle_Click()
Implied = Not Implied
If Implied Then
    ImpliedToggle.Caption = "Switch to Finding the Price"
    SigmaOrPriceL.Caption = "Price of Option"
    FirstRow = "Implied (%)"
Else
    FirstRow = "Price (%)"
    ImpliedToggle.Caption = "Switch to Finding the Implied"
    SigmaOrPriceL.Caption = "Volatility (sigma) %"
End If
LegendTB.Text = MultiLine(FirstRow, "Delta", "Gamma", _
                          "Vega", "Rho", "Theta", LastRow)
SigmaOrPriceTB.Value = ""
SigmaOrPrice = 0
Compute
End Sub
'
'
Private Sub qTB_Change()
q0 = val(qTB.Value)
'recompute everything if q changes
Compute
End Sub

Private Sub SemiAnnual_Click()
If Initialized Then Compute
End Sub

'
'
Private Sub sTB_Change()
s = val(sTB.Value)
'recompute everything if s changes
Compute
End Sub
'
'
Private Sub xTB_Change()
x = val(xTB.Value)
'recompute everything if x changes
Compute
End Sub
'
'
Private Sub rTB_Change()
r0 = val(rTB.Value)
'recompute everything if r changes
Compute
End Sub
'
'
Private Sub SigmaOrPriceTB_Change()
SigmaOrPrice = val(SigmaOrPriceTB.Value)
If Not Implied Then SigmaOrPrice = SigmaOrPrice / 100
'recompute everything if SigmaOrPrice changes
Compute
End Sub
'
'
Private Sub tTB_Change()
t = val(tTB.Value)
Compute
'recompute everything if t changes
End Sub
'
'
Private Sub UserForm_Initialize()
Dim TypeList
TypeList = Array("Stock that does not pay dividends", _
               "Asset paying constant dividend yield", _
               "Foreign currency", _
               "Futures contract on any asset")
OptionType.List = TypeList
OptionType.Value = "Stock that does not pay dividends"
FirstRow = "Price"
LastRow = ""
LegendTB.Text = MultiLine(FirstRow, "Delta", "Gamma", _
                          "Vega", "Rho", "Theta", LastRow)
Initialized = True
Compute
End Sub
'
'
Private Sub OptionType_Click()
Futures = (OptionType.Value = "Futures contract on any asset")
FX = (OptionType.Value = "Foreign currency")
Simple = (OptionType.Value = "Stock that does not pay dividends")
If FX Then
    LastRow = "Foreign Rho"
Else
    LastRow = ""
End If
LegendTB.Text = MultiLine(FirstRow, "Delta", "Gamma", _
                          "Vega", "Rho", "Theta", LastRow)
If (Simple Or Futures) Then
    qTB.Text = ""
    qTB.Enabled = False
    qTB.Visible = False
    qLabel.Visible = False
Else
    qTB.Enabled = True
    qTB.Visible = True
    If (FX) Then
        qLabel.Caption = "Foreign Interest Rate %"
    Else
        qLabel.Caption = "Dividend Yield %"
    End If
    qLabel.Visible = True
End If
If Initialized Then Compute
End Sub

