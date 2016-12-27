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

Namespace DBM

    Public Class DBMPointDriver

        Public Point As Object

        Public Sub New(Point As Object)
            Me.Point=Point
        End Sub

        Public Function GetData(StartTimestamp As DateTime,EndTimestamp As DateTime) As Double
            Return CDbl(DirectCast(Point,PISDK.PIPoint).Data.Summary(StartTimestamp,EndTimestamp,PISDK.ArchiveSummaryTypeConstants.astAverage,PISDK.CalculationBasisConstants.cbTimeWeighted).Value)
        End Function

    End Class

End Namespace
