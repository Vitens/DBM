Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStatistics


    ' Contains statistical functions which are returned all at once on
    ' calling the Statistics method.


    Public Overloads Shared Function Statistics(Dependent() As Double,
      Optional Independent() As Double = Nothing) As DBMStatisticsItem

      ' Performs calculation of several statistics functions on the input
      ' data. If no values for the independent variable are passed, a linear
      ' scale starting at 0 is assumed and exponential weighting is used.
      ' Using exponential weighting with a growth rate of 10^(1/(n-1))≈1.233
      ' improves the total error of the model by 3.2% (SD by 9.4%) and the
      ' determination by 0.14% (SD by 7.4%). This results in an overall 1.3%
      ' forecast improvement.
      ' The result of the calculation is returned as a new object.

      Dim i As Integer
      Dim ExponentialWeighting As Boolean
      Dim Weight, SumX, SumY, SumXX, SumYY, SumXY As Double

      Statistics = New DBMStatisticsItem

      With Statistics

        If Independent Is Nothing Then ' No independent var, assume linear scale
          ReDim Independent(Dependent.Length-1)
          For i = 0 To Dependent.Length-1
            Independent(i) = i
          Next i
          ExponentialWeighting = True
        Else
          Weight = 1 ' Linear weight.
        End If

        ' Calculate sums
        If Dependent.Length > 0 And Dependent.Length = Independent.Length Then
          For i = 0 To Dependent.Length-1
            If Not IsNaN(Dependent(i)) And Not IsNaN(Independent(i)) Then
              If ExponentialWeighting Then Weight = ExponentialGrowthRate^i
              .Count += Weight
              .NMBE += Dependent(i)-Independent(i)
              .RMSD += (Dependent(i)-Independent(i))^2
              SumX += Weight*Independent(i)
              SumY += Weight*Dependent(i)
              SumXX += Weight*Independent(i)^2
              SumYY += Weight*Dependent(i)^2
              SumXY += Weight*Independent(i)*Dependent(i)
            End If
          Next i
        End If

        If .Count = 0 Then Return Statistics ' Empty, non-eq, or no non-NaN pair

        .Mean = SumX/.Count ' Average

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

        .Slope = (.Count*SumXY-SumX*SumY)/(.Count*SumXX-SumX^2) ' Lin.regression
        .OriginSlope = SumXY/SumXX ' Lin.regression through the origin (alpha=0)
        .Angle = SlopeToAngle(.Slope) ' Angle in degrees
        .OriginAngle = SlopeToAngle(.OriginSlope) ' Angle in degrees
        .Intercept = (SumX*SumXY-SumY*SumXX)/(SumX^2-.Count*SumXX)

        ' Standard error of the predicted y-value for each x in the regression.
        ' The standard error is a measure of the amount of error in the
        ' prediction of y for an individual x.
        For i = 0 to Dependent.Length-1
          If Not IsNaN(Dependent(i)) And Not IsNaN(Independent(i)) Then
            If ExponentialWeighting Then Weight = ExponentialGrowthRate^i
            .StandardError +=
              Weight*(Dependent(i)-Independent(i)*.Slope-.Intercept)^2
          End If
        Next i
        ' n-2 is used because two parameters (slope and intercept) were
        ' estimated in order to estimate the sum of squares.
        .StandardError = Sqrt(.StandardError/Max(0, .Count-2))

        ' A number that quantifies some type of correlation and dependence,
        ' meaning statistical relationships between two or more random
        ' variables or observed data values.
        .Correlation = (.Count*SumXY-SumX*SumY)/
          Sqrt((.Count*SumXX-SumX^2)*(.Count*SumYY-SumY^2))

        ' Average is not removed in modified correlation as the expected average
        ' is zero, assuming the calculated forecasts are correct.
        .ModifiedCorrelation = SumXY/Sqrt(SumXX*SumYY)

        ' A number that indicates the proportion of the variance in the
        ' dependent variable that is predictable from the independent variable.
        .Determination = .Correlation^2

      End With

      Return Statistics

    End Function


    Public Overloads Shared Function Statistics(
      Results As List(Of DBMResult)) As DBMStatisticsItem

      Dim i As Integer
      Dim Forecasts(Results.Count-1), Measurements(Results.Count-1) As Double

      For i = 0 To Results.Count-1
        With Results.Item(i).ForecastItem
          Forecasts(i) = .Forecast
          Measurements(i) = .Measurement
        End With
      Next i

      Return Statistics(Forecasts, Measurements)

    End Function


  End Class


End Namespace
