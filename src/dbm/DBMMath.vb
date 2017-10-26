Option Explicit
Option Strict


' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
'
' This file is part of DBM.
'
' DBM is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' DBM is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with DBM.  If not, see <http://www.gnu.org/licenses/>.


Imports System.Double
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMMath


    ' Contains mathematical and statistical functions.


    Private Shared Random As New Random


    Public Shared Function NormSInv(p As Double) As Double

      ' Returns the inverse of the standard normal cumulative distribution.
      ' The distribution has a mean of zero and a standard deviation of one.
      ' Approximation of inverse standard normal CDF developed by
      ' Peter J. Acklam

      Const a1 As Double = -39.6968302866538, a2 As Double = 220.946098424521
      Const a3 As Double = -275.928510446969, a4 As Double = 138.357751867269
      Const a5 As Double = -30.6647980661472, a6 As Double = 2.50662827745924
      Const b1 As Double = -54.4760987982241, b2 As Double = 161.585836858041
      Const b3 As Double = -155.698979859887, b4 As Double = 66.8013118877197
      Const b5 As Double = -13.2806815528857
      Const c1 As Double = -7.78489400243029E-03
      Const c2 As Double = -0.322396458041136
      Const c3 As Double = -2.40075827716184, c4 As Double = -2.54973253934373
      Const c5 As Double = 4.37466414146497, c6 As Double = 2.93816398269878
      Const d1 As Double = 7.78469570904146E-03, d2 As Double = 0.32246712907004
      Const d3 As Double = 2.445134137143, d4 As Double = 3.75440866190742
      Const p_low As Double = 0.02425, p_high As Double = 1-p_low

      Dim q, r As Double

      If p < p_low Then ' Left tail
        q = Sqrt(-2*Log(p))
        Return (((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6)/ _
          ((((d1*q+d2)*q+d3)*q+d4)*q+1)
      ElseIf p <= p_high Then
        q = p-0.5
        r = q*q
        Return (((((a1*r+a2)*r+a3)*r+a4)*r+a5)*r+a6)*q/ _
          (((((b1*r+b2)*r+b3)*r+b4)*r+b5)*r+1)
      Else ' Right tail
        q = Sqrt(-2*Log(1-p))
        Return -(((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6)/ _
          ((((d1*q+d2)*q+d3)*q+d4)*q+1)
      End If

    End Function


    Public Shared Function TInv2T(p As Double, dof As Integer) As Double

      ' Returns the two-tailed inverse of the Student's t-distribution.
      ' Hill's approx. inverse t-dist.: Comm. of A.C.M Vol.13 No.10 1970 pg 620

      Dim a, b, c, d, x, y As Double

      If dof = 1 Then
        p *= PI/2
        Return Cos(p)/Sin(p)
      ElseIf dof = 2 Then
        Return Sqrt(2/(p*(2-p))-2)
      Else
        a = 1/(dof-0.5)
        b = 48/(a^2)
        c = ((20700*a/b-98)*a-16)*a+96.36
        d = ((94.5/(b+c)-3)/b+1)*Sqrt(a*PI/2)*dof
        x = d*p
        y = x^(2/dof)
        If y > a+0.05 Then
          x = NormSInv(p/2)
          y = x^2
          If dof < 5 Then
            c += 0.3*(dof-4.5)*(x+0.6)
          End If
          c = (((d/2*x-0.5)*x-7)*x-2)*x+b+c
          y = (((((0.4*y+6.3)*y+36)*y+94.5)/c-y-3)/b+1)*x
          y = Exp(a*y^2)-1
        Else
          y = _
            ((1/(((dof+6)/(dof*y)-0.089*d-0.822)*(dof+2)*3)+0.5/(dof+4))*y-1)* _
            (dof+1)/(dof+2)+1/y
        End If
        Return Sqrt(dof*y)
      End If

    End Function


    Public Shared Function TInv(p As Double, dof As Integer) As Double

      ' Returns the left-tailed inverse of the Student's t-distribution.

      Return Sign(p-0.5)*TInv2T(1-Abs(p-0.5)*2, dof)

    End Function


    Public Shared Function MeanAbsoluteDeviationScaleFactor As Double

      ' Estimator; scale factor k
      ' For normally distributed data, multiply MAD by scale factor k to
      ' obtain an estimate of the normal scale parameter sigma.
      ' R.C. Geary. The Ratio of the Mean Deviation to the Standard Deviation
      '  as a Test of Normality. Biometrika, 1935. Cited on page 8.

      Return Sqrt(PI/2)

    End Function


    Public Shared Function MedianAbsoluteDeviationScaleFactor _
      (n As Integer) As Double ' Estimator; scale factor k

      ' k is a constant scale factor, which depends on the distribution.
      ' For a symmetric distribution with zero mean, the population MAD is the
      ' 75th percentile of the distribution.
      ' Huber, P. J. (1981). Robust statistics. New York: John Wiley.

      If n < 30 Then
        Return 1/TInv(0.75, n) ' n<30 Student's t-distribution
      Else
        Return 1/NormSInv(0.75) ' n>=30 Standard normal distribution
      End If

    End Function


    Public Shared Function ControlLimitRejectionCriterion(p As Double, _
      n As Integer) As Double

      ' Return two-sided critical z-values for confidence interval p.
      ' Student's t-distribution approaches the normal z distribution at
      ' 30 samples.
      ' Student. 1908. Probable error of a correlation
      '  coefficient. Biometrika 6, 2-3, 302–310.
      ' Hogg and Tanis' Probability and Statistical Inference (7e).

      If n < 30 Then
        Return TInv((p+1)/2, n) ' n<30 Student's t-distribution
      Else
        Return NormSInv((p+1)/2) ' n>=30 Standard normal distribution
      End If

    End Function


    Public Shared Function Mean(Values() As Double) As Double

      ' Returns the arithmetic mean; the sum of the sampled values divided
      ' by the number of items in the sample.

      Dim NonNaNCount As Integer = 0
      Dim Value As Double
      Dim Sum As Double = 0

      For Each Value In Values
        If Not IsNan(Value) Then
          Sum += Value
          NonNaNCount += 1
        End If
      Next

      Return Sum / NonNaNCount

    End Function


    Public Shared Function Median(Values() As Double) As Double

      ' The median is the value separating the higher half of a data sample,
      ' a population, or a probability distribution, from the lower half. In
      ' simple terms, it may be thought of as the "middle" value of a data set.

      Dim NonNaNCount As Integer = Values.Count(Function(v) Not IsNaN(v))
      Dim MedianValues(NonNaNCount-1) As Double

      ' Only use non NaN values
      NonNaNCount = 0
      For i as Integer = 0 to Values.Length-1
        If Not IsNan(Values(i)) Then
          MedianValues(NonNaNCount) = Values(i)
          NonNaNCount += 1
        End If
      Next i

      Array.Sort(MedianValues)

      If MedianValues.Length Mod 2 = 0 Then
        Return (MedianValues(MedianValues.Length\2)+ _
          MedianValues(MedianValues.Length\2-1))/2
      Else
        Return MedianValues(MedianValues.Length\2)
      End If

    End Function


    Public Shared Function AbsoluteDeviation(Values() As Double, _
      From As Double) As Double()

      ' Returns an array which contains the absolute values of the input
      ' array from which the central tendency has been subtracted.

      Dim i As Integer
      Dim AbsDev(Values.Length-1) As Double

      For i = 0 to Values.Length-1
        AbsDev(i) = Abs(Values(i)-From)
      Next i

      Return AbsDev

    End Function


    Public Shared Function MeanAbsoluteDeviation(Values() As Double) As Double

      ' The mean absolute deviation (MAD) of a set of data
      ' is the average distance between each data value and the mean.

      Return Mean(AbsoluteDeviation(Values, Mean(Values)))

    End Function


    Public Shared Function MedianAbsoluteDeviation(Values() As Double) As Double

      ' The median absolute deviation (MAD) is a robust measure of the
      ' variability of a univariate sample of quantitative data.

      Return Median(AbsoluteDeviation(Values, Median(Values)))

    End Function


    Public Shared Function RemoveOutliers(Values() As Double) As Double()

      ' Returns an array which contains the input data from which outliers
      ' are removed (NaN) using either the mean or median absolute deviation
      ' function.

      Dim CentralTendency, MAD, ControlLimit As Double
      Dim i As Integer

      Dim NonNaNCount As Integer = Values.Count(Function(v) Not IsNaN(v))

      CentralTendency = Median(Values)
      MAD = Median(AbsoluteDeviation(Values, CentralTendency))
      ControlLimit = MAD*MedianAbsoluteDeviationScaleFactor(NonNaNCount)* _
        ControlLimitRejectionCriterion(ConfidenceInterval, NonNaNCount)

      If ControlLimit = 0 Then ' This only happens when MAD equals 0.
        ' Use Mean Absolute Deviation instead of Median Absolute Deviation
        ' to detect outliers. Median absolute deviation has a 50% breakdown
        ' point.
        CentralTendency = Mean(Values)
        MAD = Mean(AbsoluteDeviation(Values, CentralTendency))
        ControlLimit = MAD*MeanAbsoluteDeviationScaleFactor* _
          ControlLimitRejectionCriterion(ConfidenceInterval, NonNaNCount)
      End If

      For i = 0 to Values.Length-1
        If Abs(Values(i)-CentralTendency) > ControlLimit Then ' Limit exceeded?
          Values(i) = NaN ' Exclude outlier by setting to NaN.
        End If
      Next i

      Return Values

    End Function


    Public Shared Function ExponentialMovingAverage _
      (Values() As Double) As Double

      ' Filter high frequency variation
      ' An exponential moving average (EMA), is a type of infinite impulse
      ' response filter that applies weighting factors which decrease
      ' exponentially.

      Dim Weight, TotalWeight, Value As Double

      ExponentialMovingAverage = 0
      Weight = 1 ' Initial weight
      TotalWeight = 0
      For Each Value In Values ' Least significant value first
        ExponentialMovingAverage += Value*Weight
        TotalWeight += Weight
        Weight /= 1-2/(Values.Length+1) ' Increase weight for more recent values
      Next
      ExponentialMovingAverage /= TotalWeight

      Return ExponentialMovingAverage

    End Function


    Public Shared Function SlopeToAngle(Slope As Double) As Double

      ' Returns angle in degrees for Slope.

      Return Atan(Slope)/(2*PI)*360

    End Function


    Public Shared Function RandomNumber(Min As Integer, _
      Max As Integer) As Integer

      ' Returns a random number between Min (inclusive) and Max (inclusive).

      Return Random.Next(Min, Max+1)

    End Function


  End Class


End Namespace
