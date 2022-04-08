Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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


    Private Shared _logger As DBMLoggerAbstract = New DBMLoggerConsole
    Private _lock As New Object ' Object for exclusive lock on critical section.
    Private _points As New Dictionary(Of Object, DBMPoint)


    Public Shared Property Logger As DBMLoggerAbstract
      Get
        Return _logger
      End Get
      Set(value As DBMLoggerAbstract)
        _logger = value
      End Set
    End Property


    Public Sub New(Optional logger As DBMLoggerAbstract = Nothing)

      If logger IsNot Nothing Then DBM.Logger = logger
      DBM.Logger.LogDebug(Product)

    End Sub


    Private Function Point(pointDriver As DBMPointDriverAbstract) As DBMPoint

      ' Returns DBMPoint object from the dictionary. If dictionary does not yet
      ' contain object, it is added.

      Monitor.Enter(_lock) ' Block
      Try

        If Not _points.ContainsKey(pointDriver.Point) Then
          _points.Add(pointDriver.Point, New DBMPoint(pointDriver))
        End If

        Return _points.Item(pointDriver.Point)

      Finally
        Monitor.Exit(_lock) ' Unblock
      End Try

    End Function


    Public Shared Function HasCorrelation(relErrCorr As Double,
      relErrAngle As Double) As Boolean

      ' If correlation with measurement and (relative) forecast errors are
      ' about the same size.

      Return relErrCorr > CorrelationThreshold And
        Abs(relErrAngle-SlopeToAngle(1)) <= RegressionAngleRange

    End Function


    Public Shared Function HasAnticorrelation(absErrCorr As Double,
      absErrAngle As Double, subtractSelf As Boolean) As Boolean

      ' If anticorrelation with adjacent measurement and (absolute) forecast
      ' errors are about the same size.

      Return absErrCorr < -CorrelationThreshold And
        Abs(absErrAngle+SlopeToAngle(1)) <= RegressionAngleRange And
        Not subtractSelf

    End Function


    Public Shared Function Suppress(factor As Double,
      absErrCorr As Double, absErrAngle As Double,
      relErrCorr As Double, relErrAngle As Double,
      subtractSelf As Boolean) As Double

      ' Events can be suppressed when a strong correlation is found in the
      ' relative forecast errors of a containing area, or if a strong
      ' anti-correlation is found in the absolute forecast errors of an
      ' adjacent area. In both cases, the direction (regression through origin)
      ' of the error point cloud has to be around -45 or +45 degrees to indicate
      ' that both errors are about the same (absolute) size.

      If HasCorrelation(relErrCorr, relErrAngle) Or
        HasAnticorrelation(absErrCorr, absErrAngle, subtractSelf) Then
        Return 0
      Else
        Return factor
      End If

    End Function


    Public Function GetResult(inputPointDriver As DBMPointDriverAbstract,
      correlationPoints As List(Of DBMCorrelationPoint), timestamp As DateTime,
      Optional culture As CultureInfo = Nothing) As DBMResult

      ' This is the main function to call to calculate a result for a specific
      ' timestamp. If a list of DBMCorrelationPoints is passed, events can be
      ' suppressed if a strong correlation is found.

      Return GetResults(inputPointDriver, correlationPoints,
        timestamp, NextInterval(timestamp), 1, culture)(0)

    End Function


    Public Function GetResults(inputPointDriver As DBMPointDriverAbstract,
      correlationPoints As List(Of DBMCorrelationPoint),
      startTimestamp As DateTime, endTimestamp As DateTime,
      Optional numberOfValues As Integer = 0,
      Optional culture As CultureInfo = Nothing) As List(Of DBMResult)

      ' This is the main function to call to calculate results for a time range.
      ' The end timestamp is exclusive. If a list of DBMCorrelationPoints is
      ' passed, events can be suppressed if a strong correlation is found.

      Dim offset As Integer = EMATimeOffset(EMAPreviousPeriods+1)
      Dim result As DBMResult
      Dim correlationPoint As DBMCorrelationPoint
      Dim correlationResult As DBMResult
      Dim absoluteErrorStatsItem,
        relativeErrorStatsItem As New DBMStatisticsItem

      If correlationPoints Is Nothing Then ' Empty list if Nothing was passed.
        correlationPoints = New List(Of DBMCorrelationPoint)
      End If

      ' Shift the start and end timestamps into the future using the negative
      ' EMA time offset. Then later on, shift the resulting timestamps back to
      ' the original value.
      startTimestamp = startTimestamp.AddSeconds(-offset)
      endTimestamp = endTimestamp.AddSeconds(-offset)

      ' Use culture used by the current thread if no culture was passed.
      If culture Is Nothing Then culture = CurrentThread.CurrentCulture

      ' Calculate results for input point.
      GetResults = Point(inputPointDriver).GetResults(startTimestamp,
        endTimestamp, numberOfValues, True, correlationPoints.Count > 0,
        Nothing, culture)

      If correlationPoints.Count > 0 Then ' If correlation points are available.

        For Each result In GetResults ' Iterate over results for time range.

          With result

            If Abs(.Factor) > 0 Then ' If there is an event for this result.

              For Each correlationPoint In correlationPoints ' Iterate over pts.

                If Abs(.Factor) > 0 Then ' Keep going while event not suppressed

                  ' Calculate result for correlation point. We call the
                  ' GetResults method for the entire remaining time range (but
                  ' only request a single result), so that we already have all
                  ' required data for any next intervals we might need this for
                  ' (Case 3 instead of Case 9).
                  If correlationPoint.SubtractSelf Then
                    correlationResult = Point(correlationPoint.PointDriver).
                      GetResults(.Timestamp, endTimestamp, 1, False, True,
                      Point(inputPointDriver), culture)(0) ' Subtract input.
                  Else
                    correlationResult = Point(correlationPoint.PointDriver).
                      GetResults(.Timestamp, endTimestamp, 1, False, True,
                      Nothing, culture)(0)
                  End If

                  ' Calculate statistics of errors compared to forecast.
                  absoluteErrorStatsItem = Statistics(
                    correlationResult.GetAbsoluteErrors, .GetAbsoluteErrors)
                  relativeErrorStatsItem = Statistics(
                    correlationResult.GetRelativeErrors, .GetRelativeErrors)

                  ' Suppress if not local event.
                  .Factor = Suppress(.Factor,
                    absoluteErrorStatsItem.ModifiedCorrelation,
                    absoluteErrorStatsItem.OriginAngle,
                    relativeErrorStatsItem.ModifiedCorrelation,
                    relativeErrorStatsItem.OriginAngle,
                    correlationPoint.SubtractSelf)

                End If

              Next correlationPoint

            End If

          End With

        Next result

      End If

      ' Shift the resulting timestamps back to the original value since they
      ' were moved into the future before to compensate for exponential moving
      ' average (EMA) time shifting.
      For Each result In GetResults
        result.Timestamp = result.Timestamp.AddSeconds(offset)
      Next result

      Return GetResults

    End Function


  End Class


End Namespace
