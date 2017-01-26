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

Imports System.Collections.Generic
Imports System.DateTime
Imports System.Diagnostics
Imports System.Environment
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMTests

<assembly:System.Reflection.AssemblyTitle("DBM")>

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBM

    Public Points As New Dictionary(Of Object, DBMPoint)

    Public Shared Function Version _
      (Optional SkipTests As Boolean = False) As String
      Dim Ticks As Int64 = Now.Ticks
      With FileVersionInfo.GetVersionInfo(System.Reflection.Assembly. _
        GetExecutingAssembly.Location)
        Return .FileDescription & " v" & .FileVersion & NewLine & _
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
          If(SkipTests, "", NewLine & " Unit tests " & _
          If(UnitTestResults, "PASSED", "FAILED") & NewLine & _
          " Integration tests " & _
          If(IntegrationTestResults, "PASSED", "FAILED") & NewLine & _
          " " & Round((Now.Ticks-Ticks)/10000).ToString & "ms" & NewLine)
      End With
    End Function

    Private Function Point(PointDriver As DBMPointDriver) As DBMPoint
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
      ' If anticorrelation with adjacent measurement and
      ' (absolute) prediction errors are about the same size
      If Not SubtractSelf And AbsErrCorr < -CorrelationThreshold And _
        Abs(AbsErrAngle+45) <= RegressionAngleRange Then
        ' If already suppressed due to anticorrelation
        If Factor < -CorrelationThreshold And Factor >= -1 Then
          ' Keep lowest value (strongest anticorrelation)
          If AbsErrCorr < Factor Then
            Return AbsErrCorr ' Suppress
          End If
        Else ' Not already suppressed due to anticorrelation
          Return AbsErrCorr ' Suppress
        End If
      ' If correlation with measurement and
      ' (relative) prediction errors are about the same size
      ElseIf RelErrCorr > CorrelationThreshold And _
        Abs(RelErrAngle-45) <= RegressionAngleRange Then
        ' If not already suppressed due to anticorrelation
        If Not (Factor < -CorrelationThreshold And Factor >= -1) Then
          ' If already suppressed due to correlation
          If Factor > CorrelationThreshold And Factor <= 1 Then
            ' Keep highest value (strongest correlation)
            If RelErrCorr > Factor Then
              Return RelErrCorr ' Suppress
            End If
          Else ' Not already suppressed due to correlation
            Return RelErrCorr ' Suppress
          End If
        End If
      End If
      Return Factor
    End Function

    Public Function Result(InputPointDriver As DBMPointDriver, _
      CorrelationPoints As List(Of DBMCorrelationPoint), _
      Timestamp As DateTime) As DBMResult
      Dim CorrelationResult As DBMResult
      Dim AbsoluteErrorStats, RelativeErrorStats As New DBMStatistics
      Dim Factor As Double
      If CorrelationPoints Is Nothing Then
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
            ' Calculate for correlation point, subtract input point
            CorrelationResult = Point(CorrelationPoint.PointDriver).Result _
              (Timestamp, False, True, Point(InputPointDriver))
          Else
            ' Calculate for correlation point
            CorrelationResult = Point(CorrelationPoint.PointDriver).Result _
              (Timestamp, False, True)
          End If
          ' Absolute error compared to prediction
          AbsoluteErrorStats.Calculate _
            (CorrelationResult.AbsoluteErrors, Result.AbsoluteErrors)
          ' Relative error compared to prediction
          RelativeErrorStats.Calculate _
            (CorrelationResult.RelativeErrors, Result.RelativeErrors)
          Factor = Suppress(Result.Factor, _
            AbsoluteErrorStats.ModifiedCorrelation, _
            AbsoluteErrorStats.OriginAngle, _
            RelativeErrorStats.ModifiedCorrelation, _
            RelativeErrorStats.OriginAngle, _
            CorrelationPoint.SubtractSelf)
          If Factor <> Result.Factor Then ' Has event been suppressed
            Result.Factor = Factor
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
