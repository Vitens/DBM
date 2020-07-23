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

      ' Precomputed lookup table. Return Sqrt(PI/2).
      Return 1.2533141373155

    End Function


    Public Shared Function MedianAbsoluteDeviationScaleFactor(
      n As Integer) As Double ' Estimator; scale factor k

      ' k is a constant scale factor, which depends on the distribution.
      ' For a symmetric distribution with zero mean, the population MAD is the
      ' 75th percentile of the distribution.
      ' Huber, P. J. (1981). Robust statistics. New York: John Wiley.

      ' Precomputed lookup table. For n < 30 use Student's t-distribution and
      ' return 1/TInv(0.75, n). For n >= 30 use Standard normal distribution and
      ' return 1/NormSInv(0.75).
      Return {1, 1.22474487139159, 1.30736889697561, 1.35007883495785,
        1.3761084297607, 1.39361514055844, 1.40618933948217, 1.41565535988511,
        1.42303754318443, 1.42895507786675, 1.43380414160659, 1.43784994442648,
        1.44127668519821, 1.44421627332509, 1.44676564525625, 1.44899762961753,
        1.45096800015947, 1.45272019030657, 1.45428852619312, 1.45570049415833,
        1.45697836308413, 1.45814036598851, 1.45920157448821, 1.46017455538319,
        1.46106987015974, 1.46189645956711, 1.46266194297323, 1.46337285374349,
        1.46403482604361, 1.48260221844551}(Min(Max(1, n), 30)-1)

    End Function


    Public Shared Function ControlLimitRejectionCriterion(p As Double,
      n As Integer) As Double

      ' Return two-sided critical z-values for confidence interval p.
      ' Student's t-distribution approaches the normal z distribution at
      ' 30 samples.
      ' Student. 1908. Probable error of a correlation
      '  coefficient. Biometrika 6, 2-3, 302–310.
      ' Hogg and Tanis' Probability and Statistical Inference (7e).

      ' Precomputed lookup table. For n < 30 use Student's t-distribution and
      ' return TInv((p+1)/2, n). For n >= 30 use Standard normal distribution
      ' and return NormSInv((p+1)/2).
      If p = 0.9 Then Return {6.31375151467504, 2.91998558035372,
        2.35337987284673, 2.13187860590667, 2.01505450968809, 1.94318189391582,
        1.89457912775647, 1.85954823258324, 1.833113012866, 1.81246115789112,
        1.79588483442826, 1.78228756248624, 1.77093339851784, 1.76131013613017,
        1.75305035491304, 1.74588367489161, 1.73960672436397, 1.7340636047305,
        1.72913280954136, 1.72471824089417, 1.72074290076473, 1.7171443723272,
        1.7138715256962, 1.71088207786507, 1.70814075921615, 1.705617917733,
        1.70328844370546, 1.70113093225862, 1.69912702453507, 1.64485362513316
        }(Min(Max(1, n), 30)-1)
      If p = 0.95 Then Return {12.7062047361747, 4.30265272974946,
        3.18244911629114, 2.77653956803265, 2.57059886484523, 2.44691604736583,
        2.36462552580135, 2.3060045809483, 2.26215733496121, 2.22813892331464,
        2.20098519115986, 2.17881284372084, 2.16036866305888, 2.14478669119537,
        2.13144954737931, 2.11990530043824, 2.10981557884346, 2.10092204122422,
        2.0930240554416, 2.08596344837733, 2.07961384592246, 2.07387306917712,
        2.06865761176154, 2.06389856303009, 2.05953855420554, 2.05552944013689,
        2.05183051800882, 2.04840714335215, 2.04522964371284, 1.95996398611898
        }(Min(Max(1, n), 30)-1)
      If p = 0.98 Then Return {31.8205159537739, 6.96455673428327,
        4.54070318934434, 3.74695426898794, 3.36497848297365, 3.14267908752268,
        2.99795439823355, 2.896460270221, 2.82143816179733, 2.76376951169122,
        2.718079179983, 2.68099797348541, 2.65030881635261, 2.6244940484058,
        2.60248027914904, 2.5834871725602, 2.56693397366672, 2.55237962226773,
        2.53948318440368, 2.52797699784989, 2.51764801219327, 2.50832454986352,
        2.49986673710149, 2.4921594712728, 2.48510717393004, 2.47862982243387,
        2.47265991105876, 2.46714009728068, 2.46202135963482, 2.32634787438802
        }(Min(Max(1, n), 30)-1)
      If p = 0.99 Then Return {63.6567411628715, 9.92484320091829,
        5.84090941254748, 4.60409673759695, 4.03215882548672, 3.70744533800376,
        3.49948706845407, 3.35538805213598, 3.24983553572569, 3.16927251686568,
        3.10580635664082, 3.05453945960833, 3.01227573967676, 2.97684266060759,
        2.94671282897674, 2.92078158222468, 2.89823049001191, 2.87844045084628,
        2.86093459034021, 2.84533969797327, 2.83135954946119, 2.81875605451074,
        2.80733567957723, 2.79693950204879, 2.78743581209524, 2.77871453264666,
        2.7706829571503, 2.76326245605596, 2.75638590471912, 2.57582930644392
        }(Min(Max(1, n), 30)-1)
      If p = 0.9998 Then Return {3183.0987571185, 70.7000710749682,
        22.2037424431453, 13.0336716769753, 9.67756637001068, 8.02479327295924,
        7.06343541230695, 6.44200892035063, 6.0100631085745, 5.6937817765884,
        5.45273966480189, 5.26325928448697, 5.11057016818058, 4.985007416193,
        4.87999440055711, 4.79090740230589, 4.71440459747704, 4.64801276383061,
        4.58986356974984, 4.53852008324276, 4.49285954460467, 4.4519922691007,
        4.41520435909057, 4.38191647220402, 4.35165364292672, 4.32402285657295,
        4.29869615286965, 4.27539773465322, 4.25389401842494, 3.71901648212512
        }(Min(Max(1, n), 30)-1)
      If p = 0.9999 Then Return {6366.19767131664, 99.99249984375,
        28.0001302248633, 15.5441005279491, 11.1777101101092, 9.08234654646279,
        7.88458551007244, 7.12000862809636, 6.59355806730753, 6.21098309688037,
        5.9211550860432, 5.6944421625499, 5.51250016488032, 5.36340333198812,
        5.2390816840846, 5.13388901146142, 5.04376179296378, 4.96570398914251,
        4.89745990210705, 4.83729989319778, 4.78387616152046, 4.73612332722897,
        4.69318843027343, 4.65438069776784, 4.61913487784923, 4.58698406095122,
        4.55753925013763, 4.53047380766543, 4.50551147499372, 3.89059188426729
        }(Min(Max(1, n), 30)-1)

      If n < 30 Then
        Return TInv((p+1)/2, n) ' n < 30 Student's t-distribution
      Else
        Return NormSInv((p+1)/2) ' n >= 30 Standard normal distribution
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
