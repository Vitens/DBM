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

Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMPrediction

    Public MeasuredValue, PredictedValue, LowerControlLimit, _
      UpperControlLimit As Double

    Public Function ShallowCopy As DBMPrediction
      ' The MemberwiseClone method creates a shallow copy by creating a new
      ' object, and then copying the nonstatic fields of the current object to
      ' the new object. If a field is a value type, a bit-by-bit copy of the
      ' field is performed. If a field is a reference type, the reference is
      ' copied but the referred object is not; therefore, the original object
      ' and its clone refer to the same object.
      Return DirectCast(Me.MemberwiseClone, DBMPrediction)
    End Function

    Public Sub New(Optional MeasuredValue As Double = 0, _
      Optional PredictedValue As Double = 0, _
      Optional LowerControlLimit As Double = 0, _
      Optional UpperControlLimit As Double = 0)
      Me.MeasuredValue = MeasuredValue
      Me.PredictedValue = PredictedValue
      Me.LowerControlLimit = LowerControlLimit
      Me.UpperControlLimit = UpperControlLimit
    End Sub

    Public Sub Calculate(Values() As Double)
      ' Calculates and stores prediction and control limits
      Dim Statistics As New DBMStatistics
      Dim ControlLimit As Double
      With Statistics
        ' Calculate statistics for data after removing outliers
        .Calculate(RemoveOutliers(Values.Take(Values.Length-1).ToArray))
        MeasuredValue = Values(ComparePatterns)
        ' Extrapolate regression one interval
        PredictedValue = ComparePatterns*.Slope+.Intercept
        ControlLimit = ControlLimitRejectionCriterion(ConfidenceInterval, _
          .Count-1)*.StandardError
        ' Set upper and lower control limits, based on prediction, rejection
        ' criterion and standard error of the regression.
        LowerControlLimit = PredictedValue-ControlLimit
        UpperControlLimit = PredictedValue+ControlLimit
      End With
    End Sub

  End Class

End Namespace
