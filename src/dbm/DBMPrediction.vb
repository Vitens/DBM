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


Imports System.Linq
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPrediction


    Public Shared Function Calculate(Values() As Double) As DBMPredictionData

      ' Calculates and stores prediction and control limits by removing
      ' outliers from the Values array and extrapolating the regression
      ' line by one interval.
      ' The result of the calculation is returned as a new object.

      Dim StatisticsData As New DBMStatisticsData
      Dim ControlLimit As Double

      Calculate = New DBMPredictionData

      ' Calculate statistics for data after removing outliers. Exclude the
      ' last sample in the array as this is the current measured value for
      ' which we need to calculate a prediction and control limits.
      StatisticsData = DBMStatistics.Calculate _
        (RemoveOutliers(Values.Take(Values.Length-1).ToArray))

      With Calculate

        .MeasuredValue = Values(ComparePatterns)

        ' Extrapolate regression by one interval and use this result as a
        ' prediction.
        .PredictedValue = _
          ComparePatterns*StatisticsData.Slope+StatisticsData.Intercept

        ' Control limits are determined by using measures of process variation
        ' and are based on the concepts surrounding hypothesis testing and
        ' interval estimation. They are used to detect signals in process data
        ' that indicate that a process is not in control and, therefore, not
        ' operating predictably.
        ControlLimit = ControlLimitRejectionCriterion(ConfidenceInterval, _
          StatisticsData.Count-1)*StatisticsData.StandardError

        ' Set upper and lower control limits based on prediction, rejection
        ' criterion and standard error of the regression.
        .LowerControlLimit = .PredictedValue-ControlLimit
        .UpperControlLimit = .PredictedValue+ControlLimit

      End With

      Return Calculate

    End Function


  End Class


End Namespace
