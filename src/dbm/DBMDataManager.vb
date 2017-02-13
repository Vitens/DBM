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

Imports System.Collections.Generic
Imports System.Double
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMTests

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMDataManager

    Public PointDriver As DBMPointDriver
    Private Values As New Dictionary(Of DateTime, Double)

    Public Sub New(PointDriver As DBMPointDriver)
      Me.PointDriver = PointDriver
    End Sub

    Public Function Value(Timestamp As DateTime) As Double
      ' Value returns a specific value using the PointDriver object. Values are
      ' cached in memory, and, if available for the requested timestamp, have
      ' preference over using the PointDriver object.
      ' Returns value at timestamp, either from cache or using driver
      If Values.ContainsKey(Timestamp) Then ' In cache
        Value = Values.Item(Timestamp) ' Return value from cache
      Else
        ' Do not use point driver when running tests
        If TestsRunning Then
          ' Return item from internal test data array in DBMTests
          Value = TestData(TestDataIndex)
          TestDataIndex = (TestDataIndex+1) Mod TestData.Length ' Increase index
        Else
          Try
            ' Get data using driver
            Value = PointDriver.GetData _
              (Timestamp, Timestamp.AddSeconds(CalculationInterval))
          Catch
            Value = NaN ' Error, return Not a Number
          End Try
        End If
        Do While Values.Count >= MaxDataManagerValues ' Limit cache size
          ' Remove random cached value
          Values.Remove(Values.ElementAt(RandomNumber(0, Values.Count-1)).Key)
        Loop
        Values.Add(Timestamp, Value) ' Add to cache
      End If
      Return Value
    End Function

  End Class

End Namespace
