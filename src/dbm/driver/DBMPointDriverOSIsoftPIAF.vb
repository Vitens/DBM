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
    Private PreviousStartTimestamp As DateTime = DateTime.MaxValue
    Private PreviousEndTimestamp As DateTime = DateTime.MinValue


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Public Overrides Sub PrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)

      ' Retrieves a value for each interval in the time range from OSIsoft PI AF
      ' and stores this in the Values dictionary. The (aligned) end time itself
      ' is excluded.

      Dim Snapshot As DateTime = New AFTime(AlignNextInterval(
        DirectCast(Point, AFAttribute).GetValue.Timestamp.UtcSeconds,
        CalculationInterval)).LocalTime ' Interval after snapshot timestamp
      Dim Value As AFValue

      ' Never retrieve values beyond the snapshot time aligned to the next
      ' interval. Preserve the Kind property when limiting the variables if
      ' constructing by ticks.
      StartTimestamp = New DateTime(
        Min(StartTimestamp.Ticks, Snapshot.Ticks), StartTimestamp.Kind)
      EndTimestamp = New DateTime(
        Min(EndTimestamp.Ticks, Snapshot.Ticks), EndTimestamp.Kind)

      ' Exit this sub if there is no data to retrieve or when the start
      ' timestamp is not before the end timestamp.
      If Not StartTimestamp < EndTimestamp Then Exit Sub

      ' Determine what data stored in memory can be reused, what needs to be
      ' removed, and what needs to be retrieved from the PI System and stored
      ' in memory. Here is a simplified overview of how the different cases
      ' (S(tart)-E(nd)) are handled, compared to the time range stored
      ' previously (PS-PE):
      '             PS*==========*PE  S...PS E...PE Action
      '   Case 1:     S==========E      =      =    Do nothing
      '   Case 2:  S++|==========E      <      =    Add backward
      '   Case 3:     ---S=======E      >      =    Do nothing
      '   Case 4:     S=======E---      =      <    Do nothing
      '   Case 5:  S++|=======E---      <      <    Remove forward, add backward
      '   Case 6:     ---S====E---      >      <    Do nothing
      '   Case 7:     S==========|++E   =      >    Add forward
      '   Case 8:  S++|==========|++E   <      >    Clear all
      '   Case 9:     ---S=======|++E   >      >    Remove backward, add forward
      '   Case 10: S==E a) or b) S==E   E<=PS S>=PE Clear all
      If (StartTimestamp < PreviousStartTimestamp And
        EndTimestamp > PreviousEndTimestamp) Or
        EndTimestamp <= PreviousStartTimestamp Or
        StartTimestamp >= PreviousEndTimestamp Then ' Cases 8, 10a), 10b)
        Values.Clear ' Clear all
        PreviousStartTimestamp = StartTimestamp
        PreviousEndTimestamp = EndTimestamp
      Else If StartTimestamp >= PreviousStartTimestamp And
        EndTimestamp <= PreviousEndTimestamp Then ' Cases 1, 3, 4, 6
        Exit Sub ' Do nothing
      Else If StartTimestamp < PreviousStartTimestamp Then ' Cases 2, 5
        If EndTimestamp < PreviousEndTimestamp Then ' Case 5
          Do While EndTimestamp < PreviousEndTimestamp ' Remove forward
            PreviousEndTimestamp = PreviousEndTimestamp.
              AddSeconds(-CalculationInterval)
            Values.Remove(PreviousEndTimestamp)
          Loop
        End If
        EndTimestamp = PreviousStartTimestamp ' Add backward
        PreviousStartTimestamp = StartTimestamp
      Else If EndTimestamp > PreviousEndTimestamp Then ' Cases 7, 9
        If StartTimestamp > PreviousStartTimestamp Then ' Case 9
          Do While PreviousStartTimestamp < StartTimestamp ' Remove backward
            Values.Remove(PreviousStartTimestamp)
            PreviousStartTimestamp = PreviousStartTimestamp.
              AddSeconds(CalculationInterval)
          Loop
        End If
        StartTimestamp = PreviousEndTimestamp ' Add forward
        PreviousEndTimestamp = EndTimestamp
      End If

      For Each Value In DirectCast(Point, AFAttribute).Data.InterpolatedValues(
        New AFTimeRange(New AFTime(StartTimestamp),
        New AFTime(EndTimestamp.AddSeconds(-CalculationInterval))),
        New AFTimeSpan(0, 0, 0, 0, 0, CalculationInterval, 0),
        Nothing, Nothing, True) ' Get interpolated values for time range.

        ' Make sure that the retrieved data type is a Double and also that the
        ' timestamp is not already stored in memory (could happen because of
        ' DST time overlap). Be sure the AF Attribute is configured accordingly.
        If TypeOf Value.Value Is Double AndAlso
          Not Values.ContainsKey(Value.Timestamp.LocalTime) Then
          Values.Add(Value.Timestamp.LocalTime, DirectCast(Value.Value, Double))
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
