Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014, 2015, 2016, 2017, 2018  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMForecast


    Public Class DBMForecastData


      Public Measurement, ForecastValue, LowerControlLimit,
        UpperControlLimit As Double


      Public Function Range(p As Double) As Double

        ' Returns the range between the control limits and forecast value scaled
        ' for the requested confidence interval. This function requires that the
        ' results have been calculated for the default range using the Forecast
        ' function before. Always use the standard normal distribution here as
        ' we perform this calculation on an EMA-smoothed series of data and not
        ' on a single forecast result.

        Return (UpperControlLimit-ForecastValue)/
          NormSInv(BandwidthCI)*NormSInv(p)

      End Function


    End Class


    Public Shared Function Forecast(Values() As Double) As DBMForecastData

      ' Calculates and stores forecast and control limits by removing
      ' outliers from the Values array and extrapolating the regression
      ' line by one interval.
      ' The result of the calculation is returned as a new object.

      Dim StatisticsData As New DBMStatisticsData
      Dim Range As Double

      Forecast = New DBMForecastData

      With Forecast

        .Measurement = Values(Values.Length-1)

        ' Calculate statistics for data after removing outliers. Exclude the
        ' last item in the array as this is the current measured value for
        ' which we need to calculate a forecast and control limits.
        Array.Resize(Values, Values.Length-1)
        StatisticsData = Statistics(RemoveOutliers(Values))

        ' Extrapolate regression by one interval and use this result as a
        ' forecast.
        .ForecastValue =
          ComparePatterns*StatisticsData.Slope+StatisticsData.Intercept

        ' Control limits are determined by using measures of process variation
        ' and are based on the concepts surrounding hypothesis testing and
        ' interval estimation. They are used to detect signals in process data
        ' that indicate that a process is not in control and, therefore, not
        ' operating predictably.
        Range = ControlLimitRejectionCriterion(BandwidthCI,
          StatisticsData.Count-1)*StatisticsData.StandardError

        ' Set upper and lower control limits based on forecast, rejection
        ' criterion and standard error of the regression.
        .LowerControlLimit = .ForecastValue-Range
        .UpperControlLimit = .ForecastValue+Range

      End With

      Return Forecast

    End Function


  End Class


End Namespace
