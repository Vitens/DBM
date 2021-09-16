Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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
Imports System.Globalization
Imports System.Math
Imports System.Threading
Imports System.Threading.Thread
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMInfo
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBM")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBM


    ' Water company Vitens has created a demonstration site called the Vitens
    ' Innovation Playground (VIP), in which new technologies and methodologies
    ' are developed, tested, and demonstrated. The projects conducted in the
    ' demonstration site can be categorized into one of four themes: energy
    ' optimization, real-time leak detection, online water quality monitoring,
    ' and customer interaction. In the real-time leak detection theme, a method
    ' for leak detection based on statistical demand forecasting was developed.

    ' Using historical demand patterns and statistical methods - such as median
    ' absolute deviation, linear regression, sample variance, and exponential
    ' moving averages - real-time values can be compared to a forecast demand
    ' pattern and checked to be within calculated bandwidths. The method was
    ' implemented in Vitens' realtime data historian, continuously comparing
    ' measured demand values to be within operational bounds.

    ' One of the advantages of this method is that it doesn't require manual
    ' configuration or training sets. Next to leak detection, unmeasured supply
    ' between areas and unscheduled plant shutdowns were also detected. The
    ' method was found to be such a success within the company, that it was
    ' implemented in an operational dashboard and is now used in day-to-day
    ' operations.

    ' Vitens is the largest drinking water company in The Netherlands. We
    ' deliver top quality drinking water to 5.6 million people and companies in
    ' the provinces Flevoland, Fryslan, Gelderland, Utrecht and Overijssel and
    ' some municipalities in Drenthe and Noord-Holland. Annually we deliver 350
    ' million m3 water with 1,400 employees, 100 water treatment works and
    ' 49,000 kilometres of water mains.

    ' One of our main focus points is using advanced water quality, quantity and
    ' hydraulics models to further improve and optimize our treatment and
    ' distribution processes.


    Public Shared Logger As DBMLoggerAbstract = New DBMLoggerConsole
    Private Points As New Dictionary(Of Object, DBMPoint)


    Public Sub New(Optional Logger As DBMLoggerAbstract = Nothing)

      If Logger IsNot Nothing Then DBM.Logger = Logger
      DBM.Logger.LogDebug(Application)

    End Sub


    Private Function Point(PointDriver As DBMPointDriverAbstract) As DBMPoint

      ' Returns DBMPoint object from the dictionary. If dictionary does not yet
      ' contain object, it is added.

      Monitor.Enter(Points) ' Lock
      Try

        If Not Points.ContainsKey(PointDriver.Point) Then
          Points.Add(PointDriver.Point, New DBMPoint(PointDriver))
        End If

        Return Points.Item(PointDriver.Point)

      Finally
        Monitor.Exit(Points)
      End Try

    End Function


    Public Shared Function HasCorrelation(RelErrCorr As Double,
      RelErrAngle As Double) As Boolean

      ' If correlation with measurement and (relative) forecast errors are
      ' about the same size.

      Return RelErrCorr > CorrelationThreshold And
        Abs(RelErrAngle-SlopeToAngle(1)) <= RegressionAngleRange

    End Function


    Public Shared Function HasAnticorrelation(AbsErrCorr As Double,
      AbsErrAngle As Double, SubtractSelf As Boolean) As Boolean

      ' If anticorrelation with adjacent measurement and (absolute) forecast
      ' errors are about the same size.

      Return AbsErrCorr < -CorrelationThreshold And
        Abs(AbsErrAngle+SlopeToAngle(1)) <= RegressionAngleRange And
        Not SubtractSelf

    End Function


    Public Shared Function Suppress(Factor As Double,
      AbsErrCorr As Double, AbsErrAngle As Double,
      RelErrCorr As Double, RelErrAngle As Double,
      SubtractSelf As Boolean) As Double

      ' Events can be suppressed when a strong correlation is found in the
      ' relative forecast errors of a containing area, or if a strong
      ' anti-correlation is found in the absolute forecast errors of an
      ' adjacent area. In both cases, the direction (regression through origin)
      ' of the error point cloud has to be around -45 or +45 degrees to indicate
      ' that both errors are about the same (absolute) size.

      If HasCorrelation(RelErrCorr, RelErrAngle) Or
        HasAnticorrelation(AbsErrCorr, AbsErrAngle, SubtractSelf) Then
        Return 0
      Else
        Return Factor
      End If

    End Function


    Public Function GetResult(InputPointDriver As DBMPointDriverAbstract,
      CorrelationPoints As List(Of DBMCorrelationPoint), Timestamp As DateTime,
      Optional Culture As CultureInfo = Nothing) As DBMResult

      ' This is the main function to call to calculate a result for a specific
      ' timestamp. If a list of DBMCorrelationPoints is passed, events can be
      ' suppressed if a strong correlation is found.

      Return GetResults(InputPointDriver, CorrelationPoints,
        Timestamp, NextInterval(Timestamp), 1, Culture)(0)

    End Function


    Public Function GetResults(InputPointDriver As DBMPointDriverAbstract,
      CorrelationPoints As List(Of DBMCorrelationPoint),
      StartTimestamp As DateTime, EndTimestamp As DateTime,
      Optional NumberOfValues As Integer = 0,
      Optional Culture As CultureInfo = Nothing) As List(Of DBMResult)

      ' This is the main function to call to calculate results for a time range.
      ' The end timestamp is exclusive. If a list of DBMCorrelationPoints is
      ' passed, events can be suppressed if a strong correlation is found.

      Dim Offset As Integer = EMATimeOffset(EMAPreviousPeriods+1)
      Dim Result As DBMResult
      Dim CorrelationPoint As DBMCorrelationPoint
      Dim CorrelationResult As DBMResult
      Dim AbsoluteErrorStatsItem,
        RelativeErrorStatsItem As New DBMStatisticsItem

      If CorrelationPoints Is Nothing Then ' Empty list if Nothing was passed.
        CorrelationPoints = New List(Of DBMCorrelationPoint)
      End If

      ' Shift the start and end timestamps into the future using the negative
      ' EMA time offset. Then later on, shift the resulting timestamps back to
      ' the original value.
      StartTimestamp = StartTimestamp.AddSeconds(-Offset)
      EndTimestamp = EndTimestamp.AddSeconds(-Offset)

      ' Use culture used by the current thread if no culture was passed.
      If Culture Is Nothing Then Culture = CurrentThread.CurrentCulture

      ' Calculate results for input point.
      GetResults = Point(InputPointDriver).GetResults(StartTimestamp,
        EndTimestamp, NumberOfValues, True, CorrelationPoints.Count > 0,
        Nothing, Culture)

      If CorrelationPoints.Count > 0 Then ' If correlation points are available.

        For Each Result In GetResults ' Iterate over results for time range.

          With Result

            If Abs(.Factor) > 0 Then ' If there is an event for this result.

              For Each CorrelationPoint In CorrelationPoints ' Iterate over pts.

                If Abs(.Factor) > 0 Then ' Keep going while event not suppressed

                  ' Calculate result for correlation point. We call the
                  ' GetResults method for the entire remaining time range (but
                  ' only request a single result), so that we already have all
                  ' required data for any next intervals we might need this for
                  ' (Case 3 instead of Case 9).
                  If CorrelationPoint.SubtractSelf Then
                    CorrelationResult = Point(CorrelationPoint.PointDriver).
                      GetResults(.Timestamp, EndTimestamp, 1, False, True,
                      Point(InputPointDriver), Culture)(0) ' Subtract input.
                  Else
                    CorrelationResult = Point(CorrelationPoint.PointDriver).
                      GetResults(.Timestamp, EndTimestamp, 1, False, True,
                      Nothing, Culture)(0)
                  End If

                  ' Calculate statistics of errors compared to forecast.
                  AbsoluteErrorStatsItem = Statistics(
                    CorrelationResult.AbsoluteErrors, .AbsoluteErrors)
                  RelativeErrorStatsItem = Statistics(
                    CorrelationResult.RelativeErrors, .RelativeErrors)

                  ' Suppress if not local event.
                  .Factor = Suppress(.Factor,
                    AbsoluteErrorStatsItem.ModifiedCorrelation,
                    AbsoluteErrorStatsItem.OriginAngle,
                    RelativeErrorStatsItem.ModifiedCorrelation,
                    RelativeErrorStatsItem.OriginAngle,
                    CorrelationPoint.SubtractSelf)

                End If

              Next CorrelationPoint

            End If

          End With

        Next Result

      End If

      ' Shift the resulting timestamps back to the original value since they
      ' were moved into the future before to compensate for exponential moving
      ' average (EMA) time shifting.
      For Each Result In GetResults
        Result.Timestamp = Result.Timestamp.AddSeconds(Offset)
      Next Result

      Return GetResults

    End Function


  End Class


End Namespace
