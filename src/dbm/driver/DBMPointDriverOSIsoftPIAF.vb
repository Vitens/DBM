Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
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
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data.AFCalculationBasis
Imports OSIsoft.AF.Data.AFSummaryTypes
Imports OSIsoft.AF.Data.AFTimestampCalculation
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMPointDriverOSIsoftPIAF")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPointDriver
    Inherits DBMPointDriverAbstract


    ' Description: Driver for OSIsoft PI Asset Framework.
    ' Identifier (Point): OSIsoft.AF.Asset.AFAttribute (PI AF attribute)


    Private Values As New Dictionary(Of DateTime, Double)


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Public Overrides Sub PrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)

      ' Retrieves an average value for each interval in the time range from
      ' OSIsoft PI AF and stores this in the Values dictionary. The (aligned)
      ' end time itself is excluded.

      Dim PIValues As AFValues
      Dim Value As AFValue

      PIValues = DirectCast(Point, AFAttribute).Data.Summaries(
        New AFTimeRange(New AFTime(StartTimestamp),
        New AFTime(EndTimestamp)), New AFTimeSpan(0, 0, 0, 0, 0,
        CalculationInterval, 0), Average, TimeWeighted, EarliestTime).
        Item(Average) ' Get averages from PI AF

      Values.Clear
      For Each Value In PIValues ' Store averages in Values dictionary
        If Not Values.ContainsKey(Value.Timestamp.LocalTime) Then ' DST dupes
          If TypeOf Value.Value Is Double Then
            Values.Add(Value.Timestamp.LocalTime,
              DirectCast(Value.Value, Double))
          Else
            Values.Add(Value.Timestamp.LocalTime, NaN)
          End If
        End If
      Next

    End Sub


    Public Overrides Function GetData(Timestamp As DateTime) As Double

      ' GetData retrieves data from the Values dictionary. Non existing
      ' timestamps are retrieved from OSIsoft PI AF directly.

      Dim Value As Double

      ' Look up data from memory
      If Values.TryGetValue(Timestamp, Value) Then ' In cache
        Return Value ' Return value from cache
      Else
        Return DirectCast(DirectCast(Point, AFAttribute).Data.Summary(
          New AFTimeRange(New AFTime(Timestamp), New AFTime(
          Timestamp.AddSeconds(CalculationInterval))), Average, TimeWeighted,
          EarliestTime).Item(Average).Value, Double) ' Get average from PI AF
      End If

    End Function


  End Class


End Namespace
