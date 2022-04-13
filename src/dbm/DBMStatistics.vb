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
Imports System.Collections.Generic
Imports System.Double
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStatistics


    ' Contains statistical functions which are returned all at once on
    ' calling the Statistics method.


    Public Overloads Shared Function Statistics(dependent() As Double,
      Optional independent() As Double = Nothing) As DBMStatisticsItem

      ' Performs calculation of several statistics functions on the input
      ' data. If no values for the independent variable are passed, a linear
      ' scale starting at 0 is assumed and exponential weighting is used.
      ' The result of the calculation is returned as a new object.

      Dim i As Integer
      Dim weights(), totalWeight, sumX, sumY, sumXX, sumYY, sumXY As Double

      Statistics = New DBMStatisticsItem

      With Statistics

        If independent Is Nothing Then ' No independent var, assume linear scale
          ReDim independent(dependent.Length-1)
          For i = 0 To dependent.Length-1
            independent(i) = i
          Next i
          weights = ExponentialWeights(dependent.Length) ' Get weights.
        Else
          ReDim weights(dependent.Length-1)
          For i = 0 To dependent.Length-1
            weights(i) = 1 ' Linear weight.
          Next i
        End If

        ' Calculate sums
        If dependent.Length > 0 And dependent.Length = independent.Length Then

          ' Iteration 1: Count number of values and calculate total weight.
          For i = 0 To dependent.Length-1
            If Not IsNaN(dependent(i)) And Not IsNaN(independent(i)) Then
              .Count += 1
              totalWeight += weights(i)
            End If
          Next i

          ' Iteration 2: Calculate weighted statistics.
          For i = 0 To dependent.Length-1
            If Not IsNaN(dependent(i)) And Not IsNaN(independent(i)) Then
              weights(i) = weights(i)/totalWeight*.Count
              .NMBE += weights(i)*(dependent(i)-independent(i))
              .RMSD += weights(i)*(dependent(i)-independent(i))^2
              sumX += weights(i)*independent(i)
              sumY += weights(i)*dependent(i)
              sumXX += weights(i)*independent(i)^2
              sumYY += weights(i)*dependent(i)^2
              sumXY += weights(i)*independent(i)*dependent(i)
            End If
          Next i

        End If

        If .Count = 0 Then Return Statistics ' Empty, non-eq, or no non-NaN pair

        .Mean = sumX/.Count ' Average

        ' MBE (Mean Bias Error), as its name indicates, is the average of the
        ' errors of a sample space. Generally, it is a good indicator of the
        ' overall behavior of the simulated data with regards to the regression
        ' line of the sample. NMBE (Normalized Mean Bias Error) is a
        ' normalization of the MBE index that is used to scale the results of
        ' MBE, making them comparable. It quantifies the MBE index by dividing
        ' it by the mean of measured values, giving the global difference
        ' between the real values and the predicted ones.
        .NMBE = .NMBE/.Count/.Mean

        ' The root-mean-square deviation (RMSD) or root-mean-square error (RMSE)
        ' is a frequently used measure of the differences between values (sample
        ' or population values) predicted by a model or an estimator and the
        ' values observed. The RMSD represents the square root of the second
        ' sample moment of the differences between predicted values and observed
        ' values or the quadratic mean of these differences. These deviations
        ' are called residuals when the calculations are performed over the data
        ' sample that was used for estimation and are called errors (or
        ' prediction errors) when computed out-of-sample. The RMSD serves to
        ' aggregate the magnitudes of the errors in predictions for various data
        ' points into a single measure of predictive power.
        .RMSD = Sqrt(.RMSD/.Count)

        ' Normalizing the RMSD facilitates the comparison between datasets or
        ' models with different scales. Though there is no consistent means of
        ' normalization in the literature, common choices are the mean or the
        ' range (defined as the maximum value minus the minimum value) of the
        ' measured data. This value is commonly referred to as the normalized
        ' root-mean-square deviation or error (NRMSD or NRMSE), and often
        ' expressed as a percentage, where lower values indicate less residual
        ' variance. In many cases, especially for smaller samples, the sample
        ' range is likely to be affected by the size of sample which would
        ' hamper comparisons. When normalizing by the mean value of the
        ' measurements, the term coefficient of variation of the RMSD, CV(RMSD)
        ' may be used to avoid ambiguity. This is analogous to the coefficient
        ' of variation with the RMSD taking the place of the standard deviation.
        .CVRMSD = .RMSD/.Mean

        .Slope = (.Count*sumXY-sumX*sumY)/(.Count*sumXX-sumX^2) ' Lin.regression
        .OriginSlope = sumXY/sumXX ' Lin.regression through the origin (alpha=0)
        .Angle = SlopeToAngle(.Slope) ' Angle in degrees
        .OriginAngle = SlopeToAngle(.OriginSlope) ' Angle in degrees
        .Intercept = (sumX*sumXY-sumY*sumXX)/(sumX^2-.Count*sumXX)

        ' Standard error of the predicted y-value for each x in the regression.
        ' The standard error is a measure of the amount of error in the
        ' prediction of y for an individual x.
        For i = 0 to dependent.Length-1
          If Not IsNaN(dependent(i)) And Not IsNaN(independent(i)) Then
            .StandardError +=
              weights(i)*(dependent(i)-independent(i)*.Slope-.Intercept)^2
          End If
        Next i
        ' n-2 is used because two parameters (slope and intercept) were
        ' estimated in order to estimate the sum of squares.
        .StandardError = Sqrt(.StandardError/Max(0, .Count-2))

        ' A number that quantifies some type of correlation and dependence,
        ' meaning statistical relationships between two or more random
        ' variables or observed data values.
        .Correlation = (.Count*sumXY-sumX*sumY)/
          Sqrt((.Count*sumXX-sumX^2)*(.Count*sumYY-sumY^2))

        ' Average is not removed in modified correlation as the expected average
        ' is zero, assuming the calculated forecasts are correct.
        .ModifiedCorrelation = sumXY/Sqrt(sumXX*sumYY)

        ' A number that indicates the proportion of the variance in the
        ' dependent variable that is predictable from the independent variable.
        .Determination = .Correlation^2

      End With

      Return Statistics

    End Function


    Public Overloads Shared Function Statistics(
      results As List(Of DBMResult)) As DBMStatisticsItem

      Dim i As Integer
      Dim forecasts(results.Count-1), measurements(results.Count-1) As Double

      For i = 0 To results.Count-1
        With results.Item(i).ForecastItem
          forecasts(i) = .Forecast
          measurements(i) = .Measurement
        End With
      Next i

      Return Statistics(forecasts, measurements)

    End Function


  End Class


End Namespace
