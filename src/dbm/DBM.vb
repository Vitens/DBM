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
Imports System.Globalization
Imports System.Math
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


    Private PointsCache As New DBMCache(CInt((4^CacheSizeFactor)/4)) ' 1024 itms


    Private Function Point(PointDriver As DBMPointDriverAbstract) As DBMPoint

      ' Returns DBMPoint object from the cache. If cache does not yet contain
      ' object, it is added.

      If Not PointsCache.HasItem(PointDriver.Point) Then
        PointsCache.AddItem(PointDriver.Point, New DBMPoint(PointDriver))
      End If

      Return DirectCast(PointsCache.GetItem(PointDriver.Point), DBMPoint)

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


    Public Sub PrepareData(InputPointDriver As DBMPointDriverAbstract,
      CorrelationPoints As List(Of DBMCorrelationPoint),
      StartTimestamp As DateTime, EndTimestamp As DateTime)

      ' Will pass start and end timestamps to TryPrepareData method for input
      ' and correlation PointDrivers. The driver can then prepare the dataset
      ' for which calculations are required in the next step. The (aligned) end
      ' time itself is excluded. Useful for retrieving in bulk and caching in
      ' memory.

      Dim CorrelationPoint As DBMCorrelationPoint

      StartTimestamp = NextInterval(StartTimestamp,
        -EMAPreviousPeriods-CorrelationPreviousPeriods).
        AddDays(ComparePatterns*-7)
      If UseSundayForHolidays Then StartTimestamp =
        PreviousSunday(StartTimestamp)
      EndTimestamp = AlignTimestamp(EndTimestamp, CalculationInterval)

      Point(InputPointDriver).PointDriver.
        TryPrepareData(StartTimestamp, EndTimestamp)
      If CorrelationPoints IsNot Nothing Then
        For Each CorrelationPoint In CorrelationPoints
          Point(CorrelationPoint.PointDriver).PointDriver.
            TryPrepareData(StartTimestamp, EndTimestamp)
        Next
      End If

    End Sub


    Public Function Result(InputPointDriver As DBMPointDriverAbstract,
      CorrelationPoints As List(Of DBMCorrelationPoint), Timestamp As DateTime,
      Optional Culture As CultureInfo = Nothing) As DBMResult

      ' This is the main function to call to retrieve results for a specific
      ' timestamp. If a list of DBMCorrelationPoints is passed, events can be
      ' suppressed if a strong correlation is found.

      Dim CorrelationPoint As DBMCorrelationPoint
      Dim CorrelationResult As DBMResult
      Dim AbsoluteErrorStatsItem,
        RelativeErrorStatsItem As New DBMStatisticsItem

      If CorrelationPoints Is Nothing Then ' Empty list if Nothing was passed.
        CorrelationPoints = New List(Of DBMCorrelationPoint)
      End If

      ' Calculate for input point
      Result = Point(InputPointDriver).Result(
        Timestamp, True, CorrelationPoints.Count > 0, Nothing, Culture)

      ' If an event is found and a correlation point is available
      If CorrelationPoints.Count > 0 Then
        For Each CorrelationPoint In CorrelationPoints
          If Result.Factor <> 0 Then

            ' If pattern of correlation point contains input point
            If CorrelationPoint.SubtractSelf Then
              ' Calculate result for correlation point, subtract input point
              CorrelationResult = Point(CorrelationPoint.PointDriver).Result(
                Timestamp, False, True, Point(InputPointDriver), Culture)
            Else
              ' Calculate result for correlation point
              CorrelationResult = Point(CorrelationPoint.PointDriver).Result(
                Timestamp, False, True, Nothing, Culture)
            End If

            ' Calculate statistics of error compared to forecast
            AbsoluteErrorStatsItem = Statistics(
              CorrelationResult.AbsoluteErrors, Result.AbsoluteErrors)
            RelativeErrorStatsItem = Statistics(
              CorrelationResult.RelativeErrors, Result.RelativeErrors)

            Result.Factor = Suppress(Result.Factor,
              AbsoluteErrorStatsItem.ModifiedCorrelation,
              AbsoluteErrorStatsItem.OriginAngle,
              RelativeErrorStatsItem.ModifiedCorrelation,
              RelativeErrorStatsItem.OriginAngle,
              CorrelationPoint.SubtractSelf) ' Suppress if not a local event.

          End If
        Next
      End If

      Return Result

    End Function


  End Class


End Namespace
