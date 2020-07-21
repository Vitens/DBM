Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
        Return (((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6)/
          ((((d1*q+d2)*q+d3)*q+d4)*q+1)
      ElseIf p <= p_high Then
        q = p-0.5
        r = q*q
        Return (((((a1*r+a2)*r+a3)*r+a4)*r+a5)*r+a6)*q/
          (((((b1*r+b2)*r+b3)*r+b4)*r+b5)*r+1)
      Else ' Right tail
        q = Sqrt(-2*Log(1-p))
        Return -(((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6)/
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

      Return 1.2533141373155 ' Precomputed result for Sqrt(PI/2)

    End Function


    Public Shared Function MedianAbsoluteDeviationScaleFactor(
      n As Integer) As Double ' Estimator; scale factor k

      ' k is a constant scale factor, which depends on the distribution.
      ' For a symmetric distribution with zero mean, the population MAD is the
      ' 75th percentile of the distribution.
      ' Huber, P. J. (1981). Robust statistics. New York: John Wiley.

      ' Precomputed results
      Return {1, 1.22474487139159, 1.30737355162931, 1.35007956889414,
        1.37610857899973, 1.39361518673972, 1.4061893574828, 1.41565536789524,
        1.42303754709167, 1.42895507991759, 1.43380414275432, 1.43784994510845,
        1.44127668562794, 1.44421627361238, 1.44676564546013, 1.44899762977097,
        1.45096800028155, 1.45272019040865, 1.45428852628216, 1.45570049423867,
        1.45697836315859, 1.45814036605892, 1.45920157455578, 1.46017455544875,
        1.46106987022385, 1.46189645963019, 1.46266194303554, 1.46337285380522,
        1.46403482610491, 1.4826022185056}(Min(Max(1, n), 30)-1)

      If n < 30 Then
        Return 1/TInv(0.75, n) ' n<30 Student's t-distribution
      Else
        Return 1/NormSInv(0.75) ' n>=30 Standard normal distribution
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

      ' Precomputed results
      If p = 0.9 Then Return {6.31375151467503, 2.91998558035372,
        2.35336343480182, 2.13184678632665, 2.01504837333302, 1.9431802805153,
        1.89457860509001, 1.8595480375309, 1.83311293265624, 1.81246112281168,
        1.79588481870404, 1.78228755564932, 1.77093339598687, 1.76131013577489,
        1.75305035569257, 1.74588367627625, 1.73960672607507, 1.73406360661754,
        1.72913281152137, 1.72471824292079, 1.72074290281188, 1.71714437438024,
        1.71387152774705, 1.71088207990943, 1.7081407612519, 1.70561791975927,
        1.70328844572213, 1.70113093426593, 1.6991270265335, 1.64485362695147
        }(Min(Max(1, n), 30)-1)
      If p = 0.95 Then Return {12.7062047361747, 4.30265272974947,
        3.18244630528372, 2.7764451051978, 2.57058183563632, 2.44691185114497,
        2.36462425159279, 2.30600413520417, 2.26215716279821, 2.22813885198628,
        2.20098516009164, 2.17881282966723, 2.16036865646279, 2.14478668791781,
        2.13144954555978, 2.11990529922126, 2.10981557783332, 2.10092204024104,
        2.09302405440831, 2.08596344726587, 2.07961384472768, 2.07387306790403,
        2.06865761041905, 2.06389856162803, 2.0595385527533, 2.05552943864287,
        2.05183051648029, 2.04840714179525, 2.04522964213271, 1.95996398454006
        }(Min(Max(1, n), 30)-1)
      If p = 0.98 Then Return {31.8205159537738, 6.96455673428326,
        4.54070285856814, 3.7469473879792, 3.36492999890722, 3.14266840329098,
        2.99795156686853, 2.89645944770962, 2.82143792502581, 2.76376945811269,
        2.71807918381386, 2.68099799312091, 2.65030883791219, 2.62449406759005,
        2.60248029501112, 2.58348718527599, 2.56693398372472, 2.55237963018225,
        2.53948319062396, 2.52797700274157, 2.51764801604474, 2.50832455289908,
        2.49986673949467, 2.49215947315776, 2.48510717541076, 2.47862982359124,
        2.47265991195601, 2.46714009796747, 2.46202136015041, 2.32634787404084
        }(Min(Max(1, n), 30)-1)
      If p = 0.99 Then Return {63.656741162872, 9.92484320091832,
        5.84090930973337, 4.60409487135, 4.03214298355522, 3.70742802132477,
        3.4994832973505, 3.3553873313334, 3.24983554159213, 3.16927267261695,
        3.10580651553928, 3.0545395893929, 3.01227583871658, 2.97684273437083,
        2.94671288347524, 2.9207816224251, 2.89823051967742, 2.87844047273861,
        2.86093460646498, 2.8453397097861, 2.83135955802305, 2.81875606060014,
        2.80733568377, 2.79693950477446, 2.78743581367697, 2.77871453332968,
        2.77068295712221, 2.76326245546145, 2.75638590367061, 2.5758293035489
        }(Min(Max(1, n), 30)-1)
      If p = 0.9998 Then Return {3183.09875711899, 70.7000710749527,
        22.2037422732016, 13.0336717208984, 9.67756630088233, 8.02479277311574,
        7.06343282815736, 6.4419998210207, 6.0101321290816, 5.69382010145722,
        5.45276208882251, 5.26327300782635, 5.11057889762638, 4.98501315771973,
        4.87999828842442, 4.79091010310366, 4.71440651651603, 4.64801415510266,
        4.58986459675513, 4.53852085379374, 4.49286013134789, 4.45199272195483,
        4.41520471297036, 4.3819167519296, 4.35165386640522, 4.32402303690056,
        4.29869629974542, 4.27539785534343, 4.25389411843212, 3.71901648545571
        }(Min(Max(1, n), 30)-1)
      If p = 0.9999 Then Return {6366.19767131461, 99.9924998437688,
        28.0001300109434, 15.5441005815456, 11.1777100702764, 9.082346327295,
        7.88458426241549, 7.12000388273514, 6.5936825839441, 6.21105089129067,
        5.92119416247283, 5.69446579327097, 5.51251504959642, 5.36341304115694,
        5.23908821175432, 5.13389351754554, 5.04376497663826, 4.96570628529171,
        4.89746158886254, 4.8373011529094, 4.78387711643608, 4.73612406096515,
        4.6931890010752, 4.65438114688262, 4.61913523493401, 4.58698434762756,
        4.55753948237116, 4.53047399738543, 4.50551163121124, 3.89059188641312
        }(Min(Max(1, n), 30)-1)

      If n < 30 Then
        Return TInv((p+1)/2, n) ' n<30 Student's t-distribution
      Else
        Return NormSInv((p+1)/2) ' n>=30 Standard normal distribution
      End If

    End Function


    Public Shared Function NonNaNCount(Values() As Double) As Integer

      ' Returns number of values in the array excluding NaNs.

      Dim Value As Double
      Dim Count As Integer

      For Each Value In Values
        If Not IsNaN(Value) Then
          Count += 1
        End If
      Next

      Return Count

    End Function


    Public Shared Function Mean(Values() As Double) As Double

      ' Returns the arithmetic mean; the sum of the sampled values divided
      ' by the number of items in the sample. NaNs are excluded.

      Dim Value, Sum As Double
      Dim Count As Integer

      For Each Value In Values
        If Not IsNaN(Value) Then
          Sum += Value
          Count += 1
        End If
      Next

      Return Sum/Count

    End Function


    Public Shared Function Median(Values() As Double) As Double

      ' The median is the value separating the higher half of a data sample,
      ' a population, or a probability distribution, from the lower half. In
      ' simple terms, it may be thought of as the "middle" value of a data set.
      ' NaNs are excluded.

      Dim MedianValues(NonNaNCount(Values)-1), Value As Double
      Dim Count As Integer

      If MedianValues.Length = 0 Then Return NaN ' No non-NaN values.

      For Each Value In Values
        If Not IsNaN(Value) Then
          MedianValues(Count) = Value
          Count += 1
        End If
      Next

      Array.Sort(MedianValues)

      If MedianValues.Length Mod 2 = 0 Then
        Return (MedianValues(MedianValues.Length\2)+
          MedianValues(MedianValues.Length\2-1))/2
      Else
        Return MedianValues(MedianValues.Length\2)
      End If

    End Function


    Public Shared Function AbsoluteDeviation(Values() As Double,
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


    Public Shared Function UseMeanAbsoluteDeviation(
      Values() As Double) As Boolean

      ' Returns true if the Mean Absolute Deviation has to be used instead of
      ' the Median Absolute Deviation to detect outliers. Median absolute
      ' deviation has a 50% breakdown point.

      Return MedianAbsoluteDeviation(Values) = 0

    End Function


    Public Shared Function CentralTendency(Values() As Double) As Double

      ' Returns the central tendency for the data series, based on either the
      ' mean or median absolute deviation. In statistics, a central tendency
      ' (or measure of central tendency) is a central or typical value for a
      ' probability distribution.

      If UseMeanAbsoluteDeviation(Values) Then
        Return Mean(Values)
      Else
        Return Median(Values)
      End If

    End Function


    Public Shared Function ControlLimit(Values() As Double,
      p As Double) As Double

      ' Returns the control limits for the data series, based on either the
      ' mean or median absolute deviation, scale factor and rejection criterion.
      ' Control limits are used to detect signals in process data that indicate
      ' that a process is not in control and, therefore, not operating
      ' predictably.

      Dim Count As Integer = NonNaNCount(Values)

      If UseMeanAbsoluteDeviation(Values) Then
        Return MeanAbsoluteDeviation(Values)*
          MeanAbsoluteDeviationScaleFactor*
          ControlLimitRejectionCriterion(p, Count-1)
      Else
        Return MedianAbsoluteDeviation(Values)*
          MedianAbsoluteDeviationScaleFactor(Count-1)*
          ControlLimitRejectionCriterion(p, Count-1)
      End If

    End Function


    Public Shared Function RemoveOutliers(Values() As Double) As Double()

      ' Returns an array which contains the input data from which outliers
      ' are filtered (replaced with NaNs) using either the mean or median
      ' absolute deviation function.

      Dim ValuesCentralTendency, ValuesControlLimit,
        FilteredValues(Values.Length-1) As Double
      Dim i As Integer

      ValuesCentralTendency = CentralTendency(Values)
      ValuesControlLimit = ControlLimit(Values, OutlierCI)

      For i = 0 to Values.Length-1
        If Abs(Values(i)-ValuesCentralTendency) > ValuesControlLimit Then
          FilteredValues(i) = NaN ' Filter outlier
        Else
          FilteredValues(i) = Values(i) ' Keep inlier
        End If
      Next i

      Return FilteredValues

    End Function


    Public Shared Function ExponentialMovingAverage(
      Values() As Double) As Double

      ' Filter high frequency variation
      ' An exponential moving average (EMA), is a type of infinite impulse
      ' response filter that applies weighting factors which decrease
      ' exponentially.

      Dim Weight, TotalWeight, Value As Double

      ExponentialMovingAverage = 0
      Weight = 1 ' Initial weight
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

      ' Precomputed result
      If Slope = 1 Then Return 45

      Return Atan(Slope)/(2*PI)*360

    End Function


    Public Shared Function RandomNumber(Min As Integer,
      Max As Integer) As Integer

      ' Returns a random number between Min (inclusive) and Max (inclusive).

      Return Random.Next(Min, Max+1)

    End Function


    Public Shared Function AlignPreviousInterval(Value As Double,
      Interval As Double) As Double

      ' Align a value to the previous interval. Use a negative interval for
      ' aligning to the next interval.

      If Interval = 0 Then Return NaN

      If Value > 0 And Interval < 0 Then
        Value = Value - Interval
        Interval = Abs(Interval)
      End If

      If Value < 0 And Interval > 0 Then
        Value = Value - Interval
      End If

      Return (Value-Value Mod Interval)

    End Function


  End Class


End Namespace
