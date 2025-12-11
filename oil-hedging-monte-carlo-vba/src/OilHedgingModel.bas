Attribute VB_Name = "OilHedgingModel"
Option Explicit
' -----------------------------------------------------------------------
' Project : Oil Hedging Model
' Author  : Jae Yeon Park
' Desc    : Monte Carlo Simulation for Oil Price Hedging with Barrier Options
' Course  : Introduction to Financial Engineering (2023 Fall)
' -----------------------------------------------------------------------

Function MonteCarloSimulation2(Spot As Double, Strike1 As Double, Strike2 As Double, _
                               Barrier As Double, Volatility As Double, IR As Double, _
                               Maturity1 As Integer, Maturity2 As Integer, Maturity3 As Integer, _
                               Maturity4 As Integer, Maturity5 As Integer, Maturity6 As Integer, _
                               Maturity7 As Integer, Maturity8 As Integer, Maturity9 As Integer, _
                               Maturity10 As Integer, Maturity11 As Integer, Maturity12 As Integer, _
                               NPath As Long, AsofDate As Date) As Double

    Dim i As Long, j As Long
    Dim S As Double
    Dim Payoff() As Double
    Dim Drift As Double, vSqrt As Double
    Dim Payoff1 As Double, Payoff2 As Double, Payoff3 As Double
    Dim Payoff4 As Double, Payoff5 As Double, Payoff6 As Double
    Dim Payoff7 As Double, Payoff8 As Double, Payoff9 As Double
    Dim Payoff10 As Double, Payoff11 As Double, Payoff12 As Double

    ReDim Payoff(1 To NPath)

    ' Geometric Brownian Motion parameters (dt = 1/365)
    Drift = (IR - 0.5 * Volatility ^ 2) / 365
    vSqrt = Volatility * Sqr(1 / 365)

    Randomize ' initialize RNG

    For i = 1 To NPath
        ' reset per path
        S = Spot
        Payoff(i) = 0
        Payoff1 = 0: Payoff2 = 0: Payoff3 = 0
        Payoff4 = 0: Payoff5 = 0: Payoff6 = 0
        Payoff7 = 0: Payoff8 = 0: Payoff9 = 0
        Payoff10 = 0: Payoff11 = 0: Payoff12 = 0

        ' simulate daily to final maturity (dates are treated as serial numbers)
        For j = AsofDate + 1 To Maturity12
            ' Geometric Brownian Motion step
            S = S * Exp(Drift + vSqrt * Application.WorksheetFunction.Norm_S_Inv(Rnd()))

            ' --- Initial 3 months (fixed rate) ---
            If j = Maturity1 Then Payoff1 = Exp(-IR * (Maturity1 - AsofDate) / 365) * ((S - Strike1) * 10)
            If j = Maturity2 Then Payoff2 = Exp(-IR * (Maturity2 - AsofDate) / 365) * ((S - Strike1) * 10)
            If j = Maturity3 Then Payoff3 = Exp(-IR * (Maturity3 - AsofDate) / 365) * ((S - Strike1) * 10)

            Payoff(i) = Payoff1 + Payoff2 + Payoff3

            ' --- Barrier-linked months 4–12 ---

            ' Month 4
            If j > Maturity3 And j <= Maturity4 Then
                If S > Barrier Then Exit For
                If j = Maturity4 Then
                    If S > Strike1 Then
                        Payoff4 = Exp(-IR * (Maturity4 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff4 = Exp(-IR * (Maturity4 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 5
            If j > Maturity4 And j <= Maturity5 Then
                If S > Barrier Then Exit For
                If j = Maturity5 Then
                    If S > Strike1 Then
                        Payoff5 = Exp(-IR * (Maturity5 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff5 = Exp(-IR * (Maturity5 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 6
            If j > Maturity5 And j <= Maturity6 Then
                If S > Barrier Then Exit For
                If j = Maturity6 Then
                    If S > Strike1 Then
                        Payoff6 = Exp(-IR * (Maturity6 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff6 = Exp(-IR * (Maturity6 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 7
            If j > Maturity6 And j <= Maturity7 Then
                If S > Barrier Then Exit For
                If j = Maturity7 Then
                    If S > Strike1 Then
                        Payoff7 = Exp(-IR * (Maturity7 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff7 = Exp(-IR * (Maturity7 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 8
            If j > Maturity7 And j <= Maturity8 Then
                If S > Barrier Then Exit For
                If j = Maturity8 Then
                    If S > Strike1 Then
                        Payoff8 = Exp(-IR * (Maturity8 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff8 = Exp(-IR * (Maturity8 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 9
            If j > Maturity8 And j <= Maturity9 Then
                If S > Barrier Then Exit For
                If j = Maturity9 Then
                    If S > Strike1 Then
                        Payoff9 = Exp(-IR * (Maturity9 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff9 = Exp(-IR * (Maturity9 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 10
            If j > Maturity9 And j <= Maturity10 Then
                If S > Barrier Then Exit For
                If j = Maturity10 Then
                    If S > Strike1 Then
                        Payoff10 = Exp(-IR * (Maturity10 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff10 = Exp(-IR * (Maturity10 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 11
            If j > Maturity10 And j <= Maturity11 Then
                If S > Barrier Then Exit For
                If j = Maturity11 Then
                    If S > Strike1 Then
                        Payoff11 = Exp(-IR * (Maturity11 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff11 = Exp(-IR * (Maturity11 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' Month 12 (final)
            If j > Maturity11 And j <= Maturity12 Then
                If S > Barrier Then Exit For
                If j = Maturity12 Then
                    If S > Strike1 Then
                        Payoff12 = Exp(-IR * (Maturity12 - AsofDate) / 365) * ((S - Strike1) * 10)
                    Else
                        Payoff12 = Exp(-IR * (Maturity12 - AsofDate) / 365) * ((S - Strike2) * 20)
                    End If
                End If
            End If

            ' accumulate if we reached a payoff date
            Payoff(i) = Payoff1 + Payoff2 + Payoff3 + Payoff4 + Payoff5 + Payoff6 _
                        + Payoff7 + Payoff8 + Payoff9 + Payoff10 + Payoff11 + Payoff12

        Next j
    Next i

    MonteCarloSimulation2 = Application.Average(Payoff)

End Function
