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
Imports System.Diagnostics
Imports System.Environment
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics
Imports Vitens.DynamicBandwidthMonitor.DBMTests


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
    ' moving averages - real-time values can be compared to a predicted demand
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


    Public Points As New Dictionary(Of Object, DBMPoint)


    Public Shared Function Version As String

      ' Returns a string containing version, copyright and license information.

      Const GITHASH As String = "" ' Updated automatically by the build script.

      With FileVersionInfo.GetVersionInfo(System.Reflection.Assembly. _
        GetExecutingAssembly.Location)
        Return .ProductName & " v" & .FileVersion & _
          If(GITHASH = "", "", "+" & GITHASH) & NewLine & _
          .Comments & NewLine & _
          .LegalCopyright & NewLine & _
          NewLine & _
          "This program is free software: you can redistribute it and/or " & _
          "modify it under the terms of the GNU General Public License as " & _
          "published by the Free Software Foundation, either version 3 of " & _
          "the License, or (at your option) any later version." & NewLine & _
          NewLine & _
          "This program is distributed in the hope that it will be useful, " & _
          "but WITHOUT ANY WARRANTY; without even the implied warranty of " & _
          "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the " & _
          "GNU General Public License for more details." & NewLine & _
          NewLine & _
          "You should have received a copy of the GNU General Public " & _
          "License along with this program.  " & _
          "If not, see <http://www.gnu.org/licenses/>." & NewLine
      End With

    End Function


    Public Shared Function TestResults As String

      ' Returns a string containing test results and performance index.

      Return _
        " * Unit tests " & _
        If(UnitTestsPassed, "PASSED", "FAILED") & "." & NewLine & _
        " * Integration tests " & _
        If(IntegrationTestsPassed, "PASSED", "FAILED") & "." & NewLine & _
        " * Performance index " & _
        Round(PerformanceIndex, 1).ToString & "." & NewLine

    End Function


    Private Sub RemoveStalePoints

      ' Stale items are removed so that used resources can be freed to prevent
      ' all available memory from filling up.

      Dim Pair As KeyValuePair(Of Object, DBMPoint)
      Dim StalePoints As New List(Of Object)
      Dim StalePoint As Object

      For Each Pair In Points
        If Pair.Value.IsStale Then ' Find stale points
          StalePoints.Add(Pair.Key)
        End If
      Next

      For Each StalePoint In StalePoints
        Points.Remove(StalePoint) ' Remove stale points
      Next

    End Sub


    Private Function Point(PointDriver As DBMPointDriverAbstract) As DBMPoint

      ' Returns DBMPoint object from Points dictionary. If dictionary does not
      ' yet contain object, it is added.

      If Not Points.ContainsKey(PointDriver.Point) Then
        Points.Add(PointDriver.Point, New DBMPoint(PointDriver)) ' Add new point
      End If

      Return Points.Item(PointDriver.Point)

    End Function


    Public Shared Function Suppress(Factor As Double, _
      AbsErrCorr As Double, AbsErrAngle As Double, _
      RelErrCorr As Double, RelErrAngle As Double, _
      SubtractSelf As Boolean) As Double

      ' Events can be suppressed when a strong correlation is found in the
      ' relative prediction errors of a containing area, or if a strong
      ' anti-correlation is found in the absolute prediction errors of an
      ' adjacent area. In both cases, the direction (regression through origin)
      ' of the error point cloud has to be around -45 or +45 degrees to indicate
      ' that both errors are about the same (absolute) size.

      ' If anticorrelation with adjacent measurement and
      ' (absolute) prediction errors are about the same size.
      If Not SubtractSelf And AbsErrCorr < -CorrelationThreshold And _
        Abs(AbsErrAngle+SlopeToAngle(1)) <= RegressionAngleRange Then
        ' If already suppressed due to anticorrelation
        If Factor < -CorrelationThreshold And Factor >= -1 Then
          ' Keep lowest value (strongest anticorrelation)
          Return Min(AbsErrCorr, Factor)
        Else ' Not already suppressed due to anticorrelation
          Return AbsErrCorr ' Suppress
        End If
      ' If correlation with measurement and
      ' (relative) prediction errors are about the same size.
      ElseIf RelErrCorr > CorrelationThreshold And _
        Abs(RelErrAngle-SlopeToAngle(1)) <= RegressionAngleRange Then
        ' If not already suppressed due to anticorrelation
        If Not (Factor < -CorrelationThreshold And Factor >= -1) Then
          ' If already suppressed due to correlation
          If Factor > CorrelationThreshold And Factor <= 1 Then
            ' Keep highest value (strongest correlation)
            Return Max(RelErrCorr, Factor)
          Else ' Not already suppressed due to correlation
            Return RelErrCorr ' Suppress
          End If
        End If
      End If

      Return Factor

    End Function


    Public Sub PrepareData(InputPointDriver As DBMPointDriverAbstract, _
      CorrelationPoints As List(Of DBMCorrelationPoint), _
      StartTimestamp As DateTime, EndTimestamp As DateTime)

      ' Will pass start and end timestamps to PrepareData method for input and
      ' correlation points. The driver can then prepare the dataset for which
      ' calculations are required in the next step. The (aligned) end time
      ' itself is excluded. Useful for retrieving in bulk and caching in memory.

      Dim CorrelationPoint As DBMCorrelationPoint

      StartTimestamp = AlignTimestamp(StartTimestamp, CalculationInterval). _
        AddSeconds((EMAPreviousPeriods+CorrelationPreviousPeriods)* _
        -CalculationInterval).AddDays(ComparePatterns*-7)
      EndTimestamp = AlignTimestamp(EndTimestamp, CalculationInterval)

      InputPointDriver.PrepareData(StartTimestamp, EndTimestamp)
      If CorrelationPoints IsNot Nothing Then
        For Each CorrelationPoint In CorrelationPoints
          CorrelationPoint.PointDriver.PrepareData(StartTimestamp, EndTimestamp)
        Next
      End If

    End Sub


    Public Function Result(InputPointDriver As DBMPointDriverAbstract, _
      CorrelationPoints As List(Of DBMCorrelationPoint), _
      Timestamp As DateTime) As DBMResult

      ' This is the main function to call to retrieve results for a specific
      ' timestamp. If a list of DBMCorrelationPoints is passed, events can be
      ' suppressed if a strong correlation is found.

      Dim CorrelationPoint As DBMCorrelationPoint
      Dim CorrelationResult As DBMResult
      Dim AbsoluteErrorStatsData, _
        RelativeErrorStatsData As New DBMStatisticsData
      Dim Factor As Double

      RemoveStalePoints

      If CorrelationPoints Is Nothing Then ' Empty list if Nothing was passed.
        CorrelationPoints = New List(Of DBMCorrelationPoint)
      End If

      ' Calculate for input point
      Result = Point(InputPointDriver).Result _
        (Timestamp, True, CorrelationPoints.Count > 0)

      ' If an event is found and a correlation point is available
      If Result.Factor <> 0 And CorrelationPoints.Count > 0 Then
        For Each CorrelationPoint In CorrelationPoints
          ' If pattern of correlation point contains input point
          If CorrelationPoint.SubtractSelf Then
            ' Calculate result for correlation point, subtract input point
            CorrelationResult = Point(CorrelationPoint.PointDriver).Result _
              (Timestamp, False, True, Point(InputPointDriver))
          Else
            ' Calculate result for correlation point
            CorrelationResult = Point(CorrelationPoint.PointDriver).Result _
              (Timestamp, False, True)
          End If
          ' Calculate statistics of absolute error compared to prediction
          AbsoluteErrorStatsData = Statistics _
            (CorrelationResult.AbsoluteErrors, Result.AbsoluteErrors)
          ' Calculate statistics of relative error compared to prediction
          RelativeErrorStatsData = Statistics _
            (CorrelationResult.RelativeErrors, Result.RelativeErrors)
          Factor = Suppress(Result.Factor, _
            AbsoluteErrorStatsData.ModifiedCorrelation, _
            AbsoluteErrorStatsData.OriginAngle, _
            RelativeErrorStatsData.ModifiedCorrelation, _
            RelativeErrorStatsData.OriginAngle, _
            CorrelationPoint.SubtractSelf) ' Suppress if not a local event.
          If Factor <> Result.Factor Then ' Has event been suppressed
            Result.Factor = Factor ' Store correlation coefficient
            With CorrelationResult ' Store prediction errors for corr. point
              Array.Copy(.AbsoluteErrors, Result.CorrelationAbsoluteErrors, _
                .AbsoluteErrors.Length)
              Array.Copy(.RelativeErrors, Result.CorrelationRelativeErrors, _
                .RelativeErrors.Length)
            End With
            Result.AbsoluteErrorStatsData = AbsoluteErrorStatsData
            Result.RelativeErrorStatsData = RelativeErrorStatsData
            Result.SuppressedBy = CorrelationPoint.PointDriver ' Suppressed by
          End If
        Next
      End If

      Return Result

    End Function


  End Class


End Namespace
