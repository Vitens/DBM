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
Imports System.DateTime
Imports System.DateTimeKind
Imports System.Double
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Data.AFCalculationBasis
Imports OSIsoft.AF.Data.AFSummaryTypes
Imports OSIsoft.AF.Data.AFTimestampCalculation
Imports OSIsoft.AF.PI
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMPointDriverOSIsoftPIAF")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPointDriver
    Inherits DBMPointDriverAbstract


    ' Description: Driver for OSIsoft PI Asset Framework.
    ' Identifier (Point): OSIsoft.AF.PI.PIPoint (PI tag)


    Private Values As New Dictionary(Of AFTime, Object)


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Public Overrides Sub PrepareData(StartTimestamp As DateTime, _
      EndTimestamp As DateTime)

      EndTimestamp = EndTimestamp.AddSeconds(CalculationInterval)
      If Not Values.ContainsKey(New AFTime(StartTimestamp)) Or _
        Not Values.ContainsKey(New AFTime(EndTimestamp)) Then
        Values = DirectCast(Point, PIPoint).Summaries(New AFTimeRange(New _
          AFTime(SpecifyKind(StartTimestamp, Local)), New AFTime(SpecifyKind _
          (EndTimestamp, Local))), New AFTimeSpan(0, 0, 0, 0, 0, _
          CalculationInterval, 0), Average, TimeWeighted, EarliestTime).Item _
          (Average).ToDictionary(Function(k) k.Timestamp, Function(v) v.Value)
      End If

    End Sub


    Public Overrides Function GetData(Timestamp As DateTime) As Double

      ' GetData retrieves data from the Values dictionary. Non existing
      ' timestamps return NaN.

      Dim Value As Object = Nothing

      ' Look up data from memory
      If Values.TryGetValue(New AFTime(Timestamp), Value) Then
        Return DirectCast(Value, Double) ' Return value from cache
      Else
        Return DirectCast(DirectCast(Point, PIPoint).Summary(New AFTimeRange _
          (New AFTime(SpecifyKind(Timestamp, Local)), New AFTime(SpecifyKind _
          (Timestamp.AddSeconds(CalculationInterval), Local))), Average, _
          TimeWeighted, EarliestTime).Item(Average).Value, Double)
      End If

    End Function


  End Class


End Namespace
