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
    Private LastValueTimestamp As DateTime


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
      Dim i As Integer
      Dim PIValues As AFValues
      Dim Value As AFValue

      ' Never retrieve values beyond the snapshot time aligned to the next
      ' interval. Exit this sub if there is no data to retrieve.
      StartTimestamp = New DateTime(Min(StartTimestamp.Ticks, Snapshot.Ticks))
      EndTimestamp = New DateTime(Min(EndTimestamp.Ticks, Snapshot.Ticks))
      If StartTimestamp >= EndTimestamp Then Exit Sub

      ' If we already have the data for the time range, we do not need to do
      ' anything and this sub can be exited. This is needed to prevent
      ' retrieving data when zooming in on a larger historical time range for
      ' example. To determine if we have data for the time range in memory, just
      ' check if the first and last interval values are stored.
      If Values.ContainsKey(StartTimestamp) And
        Values.ContainsKey(EndTimestamp.AddSeconds(-CalculationInterval)) Then
        Exit Sub
      End If

      ' Check if we are in a next interval after the previous one, so that
      ' already retrieved data can be reused. This might seem quite complex, but
      ' basically we try to add only the data we do not have to the data we
      ' have already retrieved. Old values that we have stored that are probably
      ' not useful anymore are then removed so that the number of values in
      ' memory remains the same for the new time range. Note that this method
      ' currently only works when moving forward in time; moving backward will
      ' not reuse any stored data and will just retrieve all required data from
      ' the PI system. Take care that when we are moving forward, we do not also
      ' move the start timestamp backward as this would be impossible to do with
      ' the current implementation (this would require adding data to both sides
      ' of the stored data since the duration was changed). So if the start
      ' timestamp is not already stored, clear data and get the whole time
      ' range.
      If StartTimestamp <= LastValueTimestamp And
        EndTimestamp > LastValueTimestamp And
        Values.ContainsKey(StartTimestamp) Then ' Check if we can reuse data.
        If EndTimestamp >
          LastValueTimestamp.AddSeconds(CalculationInterval) Then
          For i = 1 To CInt(((EndTimestamp-LastValueTimestamp).TotalSeconds)/
            CalculationInterval)-1 ' Iterate over old values to remove.
            StartTimestamp = StartTimestamp.AddSeconds(-CalculationInterval)
            If Values.ContainsKey(StartTimestamp) Then
              Values.Remove(StartTimestamp) ' Remove old, out-of-scope values.
            End If
          Next i
        Else
          Exit Sub ' Exit, since we already have all data required in memory.
        End If
        StartTimestamp = LastValueTimestamp.AddSeconds(CalculationInterval)
      Else
        Values.Clear ' There is no data we can reuse, so clear it.
      End If
      LastValueTimestamp = EndTimestamp.AddSeconds(-CalculationInterval)

      If TypeOf DirectCast(Point, AFAttribute).PIPoint Is PIPoint Then

        ' If the source is a PI Point, use time weighted averages for each
        ' interval in the time range. Calculating interval averages is very fast
        ' with the PI Point data reference.

        PIValues = DirectCast(Point, AFAttribute).Data.Summaries(
          New AFTimeRange(New AFTime(StartTimestamp),
          New AFTime(EndTimestamp)), New AFTimeSpan(0, 0, 0, 0, 0,
          CalculationInterval, 0), Average, TimeWeighted, EarliestTime).
          Item(Average)

      Else

        ' If the source is not a PI Point, use interpolated values for each
        ' interval in the time range. This is because calculating time weighted
        ' averages for each interval for non-PI Point data references might be
        ' very costly in terms of performance. Note that the InterpolatedValues
        ' method returns one extra interval compared to the Summaries method,
        ' so one interval is subtracted from the end timestamp in the call.

        PIValues = DirectCast(Point, AFAttribute).Data.InterpolatedValues(
          New AFTimeRange(New AFTime(StartTimestamp),
          New AFTime(EndTimestamp.AddSeconds(-CalculationInterval))),
          New AFTimeSpan(0, 0, 0, 0, 0, CalculationInterval, 0),
          Nothing, Nothing, True)

      End If

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
