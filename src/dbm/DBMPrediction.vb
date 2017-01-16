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

        Public MeasuredValue, PredictedValue, LowerControlLimit, UpperControlLimit As Double

        Public Function ShallowCopy As DBMPrediction
            Return DirectCast(Me.MemberwiseClone, DBMPrediction)
        End Function

        Public Sub New(Optional MeasuredValue As Double = 0, Optional PredictedValue As Double = 0, Optional LowerControlLimit As Double = 0, Optional UpperControlLimit As Double = 0)
            Me.MeasuredValue = MeasuredValue
            Me.PredictedValue = PredictedValue
            Me.LowerControlLimit = LowerControlLimit
            Me.UpperControlLimit = UpperControlLimit
        End Sub

        Public Sub Calculate(Values() As Double) ' Calculates and stores prediction and control limits
            Dim Statistics As New DBMStatistics
            Statistics.Calculate(RemoveOutliers(Values.Take(Values.Length-1).ToArray)) ' Calculate statistics for data after removing outliers
            MeasuredValue = Values(ComparePatterns)
            PredictedValue = ComparePatterns*Statistics.Slope+Statistics.Intercept ' Extrapolate regression one interval
            LowerControlLimit = PredictedValue-ControlLimitRejectionCriterion(ConfidenceInterval, Statistics.Count-1)*Statistics.StandardError
            UpperControlLimit = PredictedValue+ControlLimitRejectionCriterion(ConfidenceInterval, Statistics.Count-1)*Statistics.StandardError
        End Sub

    End Class

End Namespace
