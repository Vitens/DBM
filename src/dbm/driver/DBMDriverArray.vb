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

    Public Class DBMDriver

        Public Sub New(Optional Data() As Object=Nothing)
        End Sub

    End Class

    Public Class DBMPointDriver

        Public Point As Object

        Public Sub New(Point As Object)
            Me.Point=Point
        End Sub

        Public Function GetData(StartTimestamp As DateTime,EndTimestamp As DateTime) As Double
            GetData=DBMUnitTests.Data(DBMUnitTests.DataIndex) ' Return item from data array
            DBMUnitTests.DataIndex=(DBMUnitTests.DataIndex+1) Mod DBMUnitTests.Data.Length ' Increase index
            Return GetData
        End Function

    End Class

End Namespace
