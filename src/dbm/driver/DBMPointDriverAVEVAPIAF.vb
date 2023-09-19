' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2023  J.H. Fitié, Vitens N.V.
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
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMPointDriverAVEVAPIAF")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPointDriver
    Inherits DBMPointDriverAbstract


    ' Description: Driver for AVEVA PI Asset Framework.
    ' Identifier (Point): OSIsoft.AF.Asset.AFAttribute (PI AF attribute)


    Public Sub New(point As Object)

      MyBase.New(point)

    End Sub


    Public Overrides Function ToString As String

      Return "AVEVA PI AF driver " & DirectCast(Point, AFAttribute).GetPath

    End Function


    Public Overrides Function SnapshotTimestamp As DateTime

      ' Returns the snapshot timestamp from AVEVA PI AF.

      Return DirectCast(Point, AFAttribute).GetValue.Timestamp.LocalTime

    End Function


    Public Overrides Sub PrepareData(startTimestamp As DateTime,
      endTimestamp As DateTime)

      ' Retrieves a value for each interval in the time range from AVEVA PI AF
      ' and stores this in memory. The (aligned) end time itself is excluded.

      Dim values As AFValues
      Dim value As AFValue

      DBM.Logger.LogDebug(
        "startTimestamp " & startTimestamp.ToString("s") & "; " &
        "endTimestamp " & endTimestamp.ToString("s"), Me.ToString)

      values = DirectCast(Point, AFAttribute).Data.InterpolatedValues(
        New AFTimeRange(New AFTime(startTimestamp),
        New AFTime(endTimestamp.AddSeconds(-CalculationInterval))),
        New AFTimeSpan(0, 0, 0, 0, 0, CalculationInterval, 0),
        Nothing, Nothing, True) ' Get interpolated values for time range.

      DBM.Logger.LogTrace(
        "Retrieved " & values.Count.ToString & " values", Me.ToString)

      For Each value In values
        DataStore.AddData(value.Timestamp.LocalTime, value.Value)
      Next

    End Sub


  End Class


End Namespace
