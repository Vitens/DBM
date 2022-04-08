Option Explicit
Option Strict


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
Imports System.IO


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMPointDriverCSV")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPointDriver
    Inherits DBMPointDriverAbstract


    ' Description: Driver for CSV files (timestamp,value).
    ' Identifier (Point): String (CSV filename)
    ' Remarks: Data interval must be the same as the CalculationInterval
    '          parameter.


    Private _splitChars As Char() = {","c, "	"c} ' Comma and tab


    Public Sub New(point As Object)

      MyBase.New(point)

    End Sub


    Public Overrides Function ToString As String

      Return "CSV driver " & DirectCast(Point, String)

    End Function


    Public Overrides Sub PrepareData(startTimestamp As DateTime,
      endTimestamp As DateTime)

      ' Retrieves information from a CSV file and stores this in memory. Passed
      ' timestamps are ignored and all data in the CSV is loaded into memory.

      Dim substrings() As String
      Dim timestamp As DateTime
      Dim value As Double

      If File.Exists(DirectCast(Point, String)) Then
        Using StreamReader As New StreamReader(DirectCast(Point, String))
          Do While Not StreamReader.EndOfStream
            ' Comma and tab delimiters; split timestamp and value
            substrings = StreamReader.ReadLine.Split(_splitChars, 2)
            If substrings.Length = 2 AndAlso
              DateTime.TryParse(substrings(0), timestamp) AndAlso
              timestamp >= startTimestamp And timestamp < endTimestamp And
              Double.TryParse(substrings(1), value) Then
              DataStore.AddData(timestamp, value)
            End If
          Loop
        End Using ' Close CSV file
      Else
        ' If Point does not represent a valid, existing file then throw a
        ' File Not Found Exception.
        Throw New FileNotFoundException(DirectCast(Point, String))
      End If

    End Sub


  End Class


End Namespace
