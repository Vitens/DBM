Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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


Imports System
Imports System.Double
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMMath


    ' Contains mathematical and statistical functions.


    Public Shared Function NormSInv(p As Double) As Double

      ' Returns the inverse of the standard normal cumulative distribution.
      ' The distribution has a mean of zero and a standard deviation of one.
      ' Approximation of inverse standard normal CDF developed by
      ' Peter J. Acklam

      Dim q, r As Double
      Dim f As Integer

      If p < 0.02425 Then ' Left tail
        q = Sqrt(-2*Log(p))
        f = 1
      ElseIf p <= 0.97575 Then
        q = p-0.5
        r = q*q
        Return (((((-3.96968302866538E+01*r+2.20946098424521E+02)*r-
          2.75928510446969E+02)*r+1.38357751867269E+02)*r-3.06647980661472E+01)*
          r+2.50662827745924)*q/(((((-5.44760987982241E+01*r+
          1.61585836858041E+02)*r-1.55698979859887E+02)*r+6.68013118877197E+01)*
          r-1.32806815528857E+01)*r+1)
      Else ' Right tail
        q = Sqrt(-2*Log(1-p))
        f = -1
      End If

      Return f*(((((-7.78489400243029E-03*q-3.22396458041136E-01)*q-
        2.40075827716184)*q-2.54973253934373)*q+4.37466414146497)*q+
        2.93816398269878)/((((7.78469570904146E-03*q+3.2246712907004E-01)*q+
        2.445134137143)*q+3.75440866190742)*q+1)

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
          y = ((1/(((dof+6)/(dof*y)-0.089*d-0.822)*(dof+2)*3)+0.5/(dof+4))*y-1)*
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


    Public Shared Function MedianAbsoluteDeviationScaleFactor(
      n As Integer) As Double ' Estimator; scale factor k

      ' k is a constant scale factor, which depends on the distribution.
      ' For a symmetric distribution with zero mean, the population MAD is the
      ' 75th percentile of the distribution.
      ' Huber, P. J. (1981). Robust statistics. New York: John Wiley.

      If n < 30 Then
        Return 1/TInv(0.75, n) ' n < 30 Student's t-distribution
      Else
        Return 1/NormSInv(0.75) ' n >= 30 Standard normal distribution
      End If

    End Function


    Public Shared Function ControlLimitRejectionCriterion(p As Double,
      n As Integer) As Double

      ' Return two-sided critical z-values for confidence interval p.
      ' Student's t-distribution approaches the normal z distribution at
      ' 30 samples.
      ' Student. 1908. Probable error of a correlation
      '  coefficient. Biometrika 6, 2-3, 302–310.
      ' Hogg and Tanis' Probability and Statistical Inference (7e).

      If n < 30 Then
        Return TInv((p+1)/2, n) ' n < 30 Student's t-distribution
      Else
        Return NormSInv((p+1)/2) ' n >= 30 Standard normal distribution
      End If

    End Function


    Public Shared Function NonNaNCount(values() As Double) As Integer

      ' Returns number of values in the array excluding NaNs.

      Dim value As Double
      Dim count As Integer

      For Each value In values
        If Not IsNaN(value) Then
          count += 1
        End If
      Next

      Return count

    End Function


    Public Shared Function Mean(values() As Double) As Double

      ' Returns the arithmetic mean; the sum of the sampled values divided
      ' by the number of items in the sample. NaNs are excluded.

      Dim value, sum As Double
      Dim count As Integer

      For Each value In values
        If Not IsNaN(value) Then
          sum += value
          count += 1
        End If
      Next

      Return sum/count

    End Function


    Public Shared Function Median(values() As Double) As Double

      ' The median is the value separating the higher half of a data sample,
      ' a population, or a probability distribution, from the lower half. In
      ' simple terms, it may be thought of as the "middle" value of a data set.
      ' NaNs are excluded.

      Dim medianValues(NonNaNCount(values)-1), Value As Double
      Dim count As Integer

      If medianValues.Length = 0 Then Return NaN ' No non-NaN values.

      For Each Value In values
        If Not IsNaN(Value) Then
          medianValues(count) = Value
          count += 1
        End If
      Next

      Array.Sort(medianValues)

      If medianValues.Length Mod 2 = 0 Then
        Return (medianValues(medianValues.Length\2)+
          medianValues(medianValues.Length\2-1))/2
      Else
        Return medianValues(medianValues.Length\2)
      End If

    End Function


    Public Shared Function StDevP(values() As Double) As Double

      ' In statistics, the standard deviation is a measure of the amount of
      ' variation or dispersion of a set of values. A low standard deviation
      ' indicates that the values tend to be close to the mean (also called the
      ' expected value) of the set, while a high standard deviation indicates
      ' that the values are spread out over a wider range.

      Dim meanValues, value As Double
      Dim count As Integer

      meanValues = Mean(values)

      For Each value In values
        If Not IsNaN(value) Then
          StDevP += (value-meanValues)^2
          count += 1
        End If
      Next
      StDevP /= count
      StDevP = Sqrt(StDevP)

      Return StDevP

    End Function


    Public Shared Function AbsoluteDeviation(values() As Double,
      from As Double) As Double()

      ' Returns an array which contains the absolute values of the input
      ' array from which the central tendency has been subtracted.

      Dim i As Integer
      Dim absDev(values.Length-1) As Double

      For i = 0 to values.Length-1
        absDev(i) = Abs(values(i)-from)
      Next i

      Return absDev

    End Function


    Public Shared Function MeanAbsoluteDeviation(values() As Double) As Double

      ' The mean absolute deviation (MAD) of a set of data
      ' is the average distance between each data value and the mean.

      Return Mean(AbsoluteDeviation(values, Mean(values)))

    End Function


    Public Shared Function MedianAbsoluteDeviation(values() As Double) As Double

      ' The median absolute deviation (MAD) is a robust measure of the
      ' variability of a univariate sample of quantitative data.

      Return Median(AbsoluteDeviation(values, Median(values)))

    End Function


    Public Shared Function UseMeanAbsoluteDeviation(
      values() As Double) As Boolean

      ' Returns true if the Mean Absolute Deviation has to be used instead of
      ' the Median Absolute Deviation to detect outliers. Median absolute
      ' deviation has a 50% breakdown point.

      Return MedianAbsoluteDeviation(values) = 0

    End Function


    Public Shared Function CentralTendency(values() As Double) As Double

      ' Returns the central tendency for the data series, based on either the
      ' mean or median absolute deviation. In statistics, a central tendency
      ' (or measure of central tendency) is a central or typical value for a
      ' probability distribution.

      If UseMeanAbsoluteDeviation(values) Then
        Return Mean(values)
      Else
        Return Median(values)
      End If

    End Function


    Public Shared Function ControlLimit(values() As Double,
      p As Double) As Double

      ' Returns the control limits for the data series, based on either the
      ' mean or median absolute deviation, scale factor and rejection criterion.
      ' Control limits are used to detect signals in process data that indicate
      ' that a process is not in control and, therefore, not operating
      ' predictably.

      Dim Count As Integer = NonNaNCount(values)

      If UseMeanAbsoluteDeviation(values) Then
        Return MeanAbsoluteDeviation(values)*
          MeanAbsoluteDeviationScaleFactor*
          ControlLimitRejectionCriterion(p, Count-1)
      Else
        Return MedianAbsoluteDeviation(values)*
          MedianAbsoluteDeviationScaleFactor(Count-1)*
          ControlLimitRejectionCriterion(p, Count-1)
      End If

    End Function


    Public Shared Function RemoveOutliers(values() As Double) As Double()

      ' Returns an array which contains the input data from which outliers
      ' are filtered (replaced with NaNs) using either the mean or median
      ' absolute deviation function.

      Dim valuesCentralTendency, valuesControlLimit,
        filteredValues(values.Length-1) As Double
      Dim i As Integer

      valuesCentralTendency = CentralTendency(values)
      valuesControlLimit = ControlLimit(values, OutlierCI)

      For i = 0 to values.Length-1
        If Abs(values(i)-valuesCentralTendency) > valuesControlLimit Then
          filteredValues(i) = NaN ' Filter outlier
        Else
          filteredValues(i) = values(i) ' Keep inlier
        End If
      Next i

      Return filteredValues

    End Function


    Public Shared Function ExponentialWeights(count As Integer) As Double()

      Dim alpha, weight, weights(count-1), totalWeight As Double
      Dim i As Integer

      alpha = 2/(count+1) ' Smoothing factor
      weight = 1 ' Initial weight

      For i = 0 To count-1
        weights(i) = weight
        totalWeight += weight
        weight /= 1-alpha ' Increase weight
      Next i

      For i = 0 To count-1
        weights(i) /= totalWeight ' Normalise weights
      Next i

      Return weights

    End Function


    Public Shared Function ExponentialMovingAverage(
      values() As Double) As Double

      ' Filter high frequency variation
      ' An exponential moving average (EMA), is a type of infinite impulse
      ' response filter that applies weighting factors which increase
      ' exponentially.

      Dim i As Integer
      Dim weights() As Double = ExponentialWeights(values.Length)
      Dim totalWeight As Double

      For i = 0 To values.Length-1
        If Not IsNaN(values(i)) Then ' Exclude NaN values.
          ExponentialMovingAverage += values(i)*weights(i)
          totalWeight += weights(i) ' Used to correct for NaN values.
        End If
      Next i

      ' Return NaN if there are no non-NaN values, or if the most recent value
      ' is NaN.
      If totalWeight = 0 Or IsNaN(values(values.Length-1)) Then Return NaN

      ExponentialMovingAverage /= totalWeight

      Return ExponentialMovingAverage

    End Function


    Public Shared Function SlopeToAngle(slope As Double) As Double

      ' Returns angle in degrees for Slope.

      Return Atan(slope)/(2*PI)*360

    End Function


    Public Shared Function Lerp(v0 As Double, v1 As Double,
      t As Double) As Double

      ' In mathematics, linear interpolation is a method of curve fitting using
      ' linear polynomials to construct new data points within the range of a
      ' discrete set of known data points.

      ' Imprecise method, which does not guarantee v = v1 when t = 1, due to
      ' floating-point arithmetic error. This method is monotonic. This form may
      ' be used when the hardware has a native fused multiply-add instruction.
      'Return v0+t*(v1-v0)

      ' Precise method, which guarantees v = v1 when t = 1. This method is
      ' monotonic only when v0 * v1 < 0. Lerping between same values might not
      ' produce the same value
      Return (1-t)*v0+t*v1

    End Function


  End Class


End Namespace
