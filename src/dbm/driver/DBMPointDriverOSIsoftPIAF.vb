Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
Imports System.Math
Imports OSIsoft.AF.Asset
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
    ' Identifier (Point): OSIsoft.AF.Asset.AFAttribute (PI AF attribute)


    Private Values As New Dictionary(Of DateTime, Double)


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Public Overrides Sub PrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)

      ' Retrieves a value for each interval in the time range from OSIsoft PI AF
      ' and stores this in the Values dictionary. The (aligned) end time itself
      ' is excluded.

      Dim Snapshot As DateTime = New AFTime(AlignPreviousInterval(
        DirectCast(Point, AFAttribute).GetValue.Timestamp.UtcSeconds,
        -CalculationInterval)).LocalTime ' Interval after snapshot timestamp
      Dim PIValues As AFValues
      Dim Value As AFValue

      ' Never retrieve values beyond the snapshot time aligned to the previous
      ' interval. Exit this sub if there is no data to retrieve.
      StartTimestamp = New DateTime(Min(StartTimestamp.Ticks, Snapshot.Ticks))
      EndTimestamp = New DateTime(Min(EndTimestamp.Ticks, Snapshot.Ticks))
      If StartTimestamp >= EndTimestamp Then Exit Sub

      ' If we already have the data for the time range, we do not need to do
      ' anything and this sub can be exited.
      If Values.ContainsKey(StartTimestamp) And
        Values.ContainsKey(EndTimestamp.AddSeconds(-CalculationInterval)) Then
        Exit Sub
      End If

      ' Check if we are in the next interval after the previous one.
      If Values.ContainsKey(StartTimestamp.AddSeconds(-CalculationInterval)) And
        Values.ContainsKey(StartTimestamp) And
        Values.ContainsKey(EndTimestamp.AddSeconds(-2*CalculationInterval)) And
        Not Values.ContainsKey(EndTimestamp.
        AddSeconds(-CalculationInterval)) Then
        ' Remove the single value prior to the start timestamp, and only read
        ' the single value at the end timestamp and append this to the
        ' dictionary.
        Values.Remove(StartTimestamp.AddSeconds(-CalculationInterval))
        StartTimestamp = EndTimestamp.AddSeconds(-CalculationInterval)
      Else
        ' We need data for another time range, so just clear the dictionary.
        Values.Clear
      End If

      ' Retrieve interpolated values for each interval in the time range. Since
      ' we want to exclude the end timestamp itself, one interval is subtracted
      ' from the end timestamp in the call.
      PIValues = DirectCast(Point, AFAttribute).Data.InterpolatedValues(
        New AFTimeRange(New AFTime(StartTimestamp),
        New AFTime(EndTimestamp.AddSeconds(-CalculationInterval))),
        New AFTimeSpan(0, 0, 0, 0, 0, CalculationInterval, 0),
        Nothing, Nothing, True)

      For Each Value In PIValues ' Store values in Values dictionary
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

      ' GetData retrieves data from the Values dictionary. If there is no data
      ' for the timestamp, return Not a Number.

      GetData = Nothing
      If Values.TryGetValue(Timestamp, GetData) Then ' In cache
        Return GetData ' Return value from cache
      Else
        Return NaN ' No data in memory for timestamp, return Not a Number.
      End If

    End Function


  End Class


End Namespace
