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
Imports System.Diagnostics
Imports System.Environment
Imports System.Math
Imports System.Text.RegularExpressions
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMTests


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBM")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBM


    Public Points As New Dictionary(Of Object, DBMPoint)


    Public Shared Function Version _
      (Optional SkipTests As Boolean = False) As String

      ' Returns a string containing version, copyright and license information.
      ' Also outputs results of unit and integration tests (unless skipped).

      Const GITHASH As String = ""

      With FileVersionInfo.GetVersionInfo(System.Reflection.Assembly. _
        GetExecutingAssembly.Location)
        Return .FileDescription & " " & _
          "v" & RegEx.Split(.FileVersion, "^(.+\..+\..+)\..+$")(1) & _
          "+" & GITHASH & NewLine & _
          .ProductName & NewLine & _
          .Comments & NewLine & _
          NewLine & _
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
          "If not, see <http://www.gnu.org/licenses/>." & NewLine & _
          If(SkipTests, "", NewLine & TestResults)
      End With

    End Function


    Private Function Point(PointDriver As DBMPointDriverAbstract) As DBMPoint

      ' Returns DBMPoint object from Points dictionary.
      ' If dictionary does not yet contain object, it is added first.

      If Not Points.ContainsKey(PointDriver.Point) Then
        ' Add to dictionary
        Points.Add(PointDriver.Point, New DBMPoint(PointDriver))
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


    Public Function Result(InputPointDriver As DBMPointDriverAbstract, _
      CorrelationPoints As List(Of DBMCorrelationPoint), _
      Timestamp As DateTime) As DBMResult

      ' This is the main function to call to retrieve results for a specific
      ' timestamp. If a list of DBMCorrelationPoints is passed, events can be
      ' suppressed if a strong correlation is found.

      Dim CorrelationPoint As DBMCorrelationPoint
      Dim CorrelationResult As DBMResult
      Dim AbsoluteErrorStats, RelativeErrorStats As New DBMStatistics
      Dim Factor As Double

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
          AbsoluteErrorStats.Calculate _
            (CorrelationResult.AbsoluteErrors, Result.AbsoluteErrors)
          ' Calculate statistics of relative error compared to prediction
          RelativeErrorStats.Calculate _
            (CorrelationResult.RelativeErrors, Result.RelativeErrors)
          Factor = Suppress(Result.Factor, _
            AbsoluteErrorStats.ModifiedCorrelation, _
            AbsoluteErrorStats.OriginAngle, _
            RelativeErrorStats.ModifiedCorrelation, _
            RelativeErrorStats.OriginAngle, _
            CorrelationPoint.SubtractSelf) ' Suppress if not a local event.
          If Factor <> Result.Factor Then ' Has event been suppressed
            Result.Factor = Factor ' Store correlation coefficient
            With CorrelationResult ' Store prediction errors for corr. point
              Array.Copy(.AbsoluteErrors, Result.CorrelationAbsoluteErrors, _
                .AbsoluteErrors.Length)
              Array.Copy(.RelativeErrors, Result.CorrelationRelativeErrors, _
                .RelativeErrors.Length)
            End With
            Result.AbsoluteErrorStats = AbsoluteErrorStats.ShallowCopy
            Result.RelativeErrorStats = RelativeErrorStats.ShallowCopy
            Result.SuppressedBy = CorrelationPoint.PointDriver ' Suppressed by
          End If
        Next
      End If

      Return Result

    End Function


  End Class


End Namespace
