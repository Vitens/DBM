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

Public Class DBMDriver

    Public Shared Data(-1) As Double
    Public Shared DataIndex As Integer

    Public Sub New(Optional Data() As Object=Nothing)
        Dim i As Integer
        ReDim DBMDriver.Data(Data.Length-1)
        For i=0 To Data.Length-1
            DBMDriver.Data(i)=CDbl(Data(i))
        Next i
        DBMDriver.DataIndex=0
    End Sub

End Class

Public Class DBMPointDriver

    Public Point As String

    Public Sub New(Point As String)
        Me.Point=Point
    End Sub

    Public Function GetData(StartTimestamp As DateTime,EndTimestamp As DateTime) As Double
        GetData=DBMDriver.Data(DBMDriver.DataIndex) ' Return item from data array passed to DBMDriver
        DBMDriver.DataIndex=(DBMDriver.DataIndex+1) Mod DBMDriver.Data.Length ' Increase index
        Return GetData
    End Function

End Class
