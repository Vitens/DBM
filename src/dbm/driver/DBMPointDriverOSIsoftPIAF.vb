Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014, 2015, 2016, 2017, 2018  J.H. Fitié, Vitens N.V.
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
Imports System.DateTime
Imports System.Double
Imports System.Math
Imports System.Threading
Imports System.Threading.Thread
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Data.AFCalculationBasis
Imports OSIsoft.AF.Data.AFSummaryTypes
Imports OSIsoft.AF.Data.AFTimestampCalculation
Imports OSIsoft.AF.PI
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMPointDriverOSIsoftPIAF")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPointDriver
    Inherits DBMPointDriverAbstract


    ' Description: Driver for OSIsoft PI Asset Framework.
    ' Identifier (Point): OSIsoft.AF.PI.PIPoint (PI tag)


    Private Values As New Dictionary(Of AFTime, Object)
    Private CacheInvalidationThread As New Thread(AddressOf InvalidateCache)
    Private LastCacheAccess As DateTime


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Private Sub InvalidateCache

      ' Invalidates the cache after it has not been accessed for at least the
      ' duration of one calculation interval (5 minutes by default). This is
      ' needed to prevent all available memory from filling up f.ex. when using
      ' a PI client application like ProcessBook to visualise large amounts of
      ' DBM results for many PI points using the PI AF data reference.

      Do While Now < LastCacheAccess.AddSeconds(CalculationInterval)
        Sleep(60*1000) ' Sleep for one minute
      Loop
      Values = New Dictionary(Of AFTime, Object)

    End Sub


    Public Overrides Sub PrepareData(StartTimestamp As DateTime, _
      EndTimestamp As DateTime)

      ' Retrieves an average value for each interval in the time range from
      ' OSIsoft PI AF and stores this in the Values dictionary. The (aligned)
      ' end time itself is excluded.

      If Not Values.ContainsKey(New AFTime(StartTimestamp)) Or _
        Not Values.ContainsKey(New AFTime(EndTimestamp.AddSeconds _
        (-CalculationInterval))) Then ' No data yet
        LastCacheAccess = Now ' Cache accessed
        If Not CacheInvalidationThread.IsAlive Then _
          CacheInvalidationThread.Start() ' Start cache invalidation thread
        Values = DirectCast(Point, PIPoint).Summaries(New AFTimeRange(New _
          AFTime(StartTimestamp), New AFTime(EndTimestamp)), New AFTimeSpan(0, _
          0, 0, 0, 0, CalculationInterval, 0), Average, TimeWeighted, _
          EarliestTime).Item(Average).ToDictionary(Function(k) k.Timestamp, _
          Function(v) v.Value) ' Store averages in dictionary
      End If

    End Sub


    Public Overrides Function GetData(Timestamp As DateTime) As Double

      ' GetData retrieves data from the Values dictionary. Non existing
      ' timestamps are retrieved from OSIsoft PI AF directly.

      Dim Value As Object = Nothing

      ' Look up data from memory
      If Values.TryGetValue(New AFTime(Timestamp), Value) Then ' In cache
        LastCacheAccess = Now ' Cache accessed
        Return DirectCast(Value, Double) ' Return value from cache
      Else
        Return DirectCast(DirectCast(Point, PIPoint).Summary(New AFTimeRange _
          (New AFTime(Timestamp), New AFTime(Timestamp.AddSeconds _
          (CalculationInterval))), Average, TimeWeighted, EarliestTime). _
          Item(Average).Value, Double) ' Get average
      End If

    End Function


  End Class


End Namespace
