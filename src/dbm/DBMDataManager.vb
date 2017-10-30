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


Imports System
Imports System.Collections.Generic
Imports System.Double
Imports System.Linq
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMTests


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMDataManager


    ' The DBMDataManager is responsible for retrieving and caching input data.
    ' It stores and uses a DBMPointDriverAbstract object, which has a GetData
    ' method used for retrieving data.


    Public Shared UseCache As Boolean = True
    Public PointDriver As DBMPointDriverAbstract
    Private Values As New Dictionary(Of DateTime, Double)


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver

    End Sub


    Public Function Value(Timestamp As DateTime) As Double

      ' Value returns a specific value using the PointDriver object. Values are
      ' cached in memory, and, if available for the requested timestamp, have
      ' preference over using the PointDriver object.
      ' Returns value at timestamp, either from cache or using driver

      If Not UseCache Then
        Return PointDriver.GetData(Timestamp, Timestamp.AddSeconds(CalculationInterval))
      End If

      If Not Values.TryGetValue(Timestamp, Value) Then ' Not in cache
        Try
          ' Get data using GetData method in driver.
          Value = PointDriver.GetData _
            (Timestamp, Timestamp.AddSeconds(CalculationInterval))
        Catch
          Value = NaN ' Error getting data, return Not a Number
        End Try
        Do While Values.Count >= MaxDataManagerValues ' Limit cache size
          ' For the implementation of a random cache eviction policy, just
          ' taking the first element of the dictionary is random enough as
          ' a Dictionary in .NET is implemented as a hashtable.
          Values.Remove(Values.ElementAt(0).Key)
        Loop
        Values.Add(Timestamp, Value) ' Add to cache
      End If

      Return Value

    End Function


  End Class


End Namespace
