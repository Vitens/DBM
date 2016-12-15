Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

Namespace DBM

    Public Class DBMPrediction

        Public MeasValue,PredValue,LowContrLimit,UppContrLimit As Double

        Public Function ShallowCopy As DBMPrediction
            Return DirectCast(Me.MemberwiseClone,DBMPrediction)
        End Function

        Public Sub New(Optional MeasValue As Double=0,Optional PredValue As Double=0,Optional LowContrLimit As Double=0,Optional UppContrLimit As Double=0)
            Me.MeasValue=MeasValue
            Me.PredValue=PredValue
            Me.LowContrLimit=LowContrLimit
            Me.UppContrLimit=UppContrLimit
        End Sub

        Public Sub Calculate(Data() As Double)
            Dim DBMStatistics As New DBMStatistics
            DBMStatistics.Calculate(DBMMath.RemoveOutliers(Data.Take(Data.Length-1).ToArray)) ' Calculate statistics for data after removing outliers
            MeasValue=Data(DBMParameters.ComparePatterns)
            PredValue=DBMParameters.ComparePatterns*DBMStatistics.Slope+DBMStatistics.Intercept ' Extrapolate regression one interval
            LowContrLimit=PredValue-DBMMath.ControlLimitRejectionCriterion(DBMParameters.ConfidenceInterval,DBMStatistics.Count-1)*DBMStatistics.StandardError
            UppContrLimit=PredValue+DBMMath.ControlLimitRejectionCriterion(DBMParameters.ConfidenceInterval,DBMStatistics.Count-1)*DBMStatistics.StandardError
        End Sub

    End Class

End Namespace
