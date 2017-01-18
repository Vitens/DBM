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
Imports System.Environment
Imports System.Globalization.CultureInfo
Imports System.Text.RegularExpressions
Imports Vitens.DynamicBandwidthMonitor.DBMParameters

<assembly:System.Reflection.AssemblyTitle("DBMTester")>

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMTester

    Private Shared InternationalFormat As Boolean = False

    Private Shared Function FormatDateTime(Timestamp As DateTime) As String
      If InternationalFormat Then
        Return Timestamp.ToUniversalTime.ToString("s") & "Z" ' ISO 8601 UTC
      Else
        Return Timestamp.ToString("s") ' ISO 8601
      End If
    End Function

    Private Shared Function Separator As String
      If InternationalFormat Then
        Return ","
      Else
        Return "	" ' Tab character
      End If
    End Function

    Private Shared Function FormatNumber(Value As Double) As String
      If InternationalFormat Then
        Return Value.ToString("0.####", InvariantCulture)
      Else
        Return Value.ToString("0.####")
      End If
    End Function

    Public Shared Sub Main
      Dim Substrings(), Parameter, Value As String
      Dim InputPointDriver As DBMPointDriver = Nothing
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim StartTimestamp, EndTimestamp As DateTime
      Dim AlwaysOutputCorrelationData As Boolean
      Dim Result As DBMResult
      Dim _DBM As New DBM
      ' Parse command line arguments
      For Each CommandLineArg In GetCommandLineArgs
        ' Parameter=Value
        If Regex.IsMatch(CommandLineArg, "^[-/](.+)=(.+)$") Then
          Substrings = CommandLineArg.Split(New Char(){"="c}, 2)
          Parameter = Substrings(0).Substring(1)
          Value = Substrings(1)
          Try
            If Parameter.ToLower.Equals("i") Then
              InputPointDriver = New DBMPointDriver(Value)
            ElseIf Parameter.ToLower.Equals("c") Then
              CorrelationPoints.Add _
                (New DBMCorrelationPoint(New DBMPointDriver(Value), False))
            ElseIf Parameter.ToLower.Equals("cs") Then
              CorrelationPoints.Add _
                (New DBMCorrelationPoint(New DBMPointDriver(Value), True))
            ElseIf Parameter.ToLower.Equals("iv") Then
              CalculationInterval = Convert.ToInt32(Value)
            ElseIf Parameter.ToLower.Equals("p") Then
              ComparePatterns = Convert.ToInt32(Value)
            ElseIf Parameter.ToLower.Equals("ep") Then
              EMAPreviousPeriods = Convert.ToInt32(Value)
            ElseIf Parameter.ToLower.Equals("ci") Then
              ConfidenceInterval = Convert.ToDouble(Value)
            ElseIf Parameter.ToLower.Equals("cp") Then
              CorrelationPreviousPeriods = Convert.ToInt32(Value)
            ElseIf Parameter.ToLower.Equals("ct") Then
              CorrelationThreshold = Convert.ToDouble(Value)
            ElseIf Parameter.ToLower.Equals("st") Then
              StartTimestamp = Convert.ToDateTime(Value)
            ElseIf Parameter.ToLower.Equals("et") Then
              EndTimestamp = Convert.ToDateTime(Value)
            ElseIf Parameter.ToLower.Equals("f") Then
              If Value.ToLower.Equals("local") Then
                InternationalFormat = False
              ElseIf Value.ToLower.Equals("intl") Then
                InternationalFormat = True
              End If
            ElseIf Parameter.ToLower.Equals("oc") Then
              AlwaysOutputCorrelationData = Convert.ToBoolean(Value)
            End If
          Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Exit Sub
          End Try
        End If
      Next
      If InputPointDriver Is Nothing Or StartTimestamp = DateTime.MinValue Then
        ' Perform unit tests
        Console.Write(DBM.Version)
      Else
        If EndTimestamp = DateTime.MinValue Then
          ' No end timestamp, set to start timestamp
          EndTimestamp = StartTimestamp
        Else
          ' Remove one interval from end timestamp
          EndTimestamp = EndTimestamp.AddSeconds(-CalculationInterval)
        End If
        Do While StartTimestamp <= EndTimestamp
          Console.Write(FormatDateTime(StartTimestamp) & Separator)
          Result = _DBM.Result _
            (InputPointDriver, CorrelationPoints, StartTimestamp)
          With Result
            Console.Write(FormatNumber(.Factor) & Separator & _
              FormatNumber(.Prediction.MeasuredValue) & Separator & _
              FormatNumber(.Prediction.PredictedValue) & Separator & _
              FormatNumber(.Prediction.LowerControlLimit) & Separator & _
              FormatNumber(.Prediction.UpperControlLimit))
          End With
          ' If an event is found and a correlation point is available
          If CorrelationPoints.Count > 0 And _
            (Result.Factor <> 0 Or AlwaysOutputCorrelationData) Then
            With Result.AbsoluteErrorStats
              Console.Write(Separator & _
                FormatNumber(.Count) & Separator & _
                FormatNumber(.Slope) & Separator & _
                FormatNumber(.Angle) & Separator & _
                FormatNumber(.Intercept) & Separator & _
                FormatNumber(.StandardError) & Separator & _
                FormatNumber(.Correlation) & Separator & _
                FormatNumber(.ModifiedCorrelation) & Separator & _
                FormatNumber(.Determination))
            End With
            With Result.RelativeErrorStats
              Console.Write(Separator & _
                FormatNumber(.Count) & Separator & _
                FormatNumber(.Slope) & Separator & _
                FormatNumber(.Angle) & Separator & _
                FormatNumber(.Intercept) & Separator & _
                FormatNumber(.StandardError) & Separator & _
                FormatNumber(.Correlation) & Separator & _
                FormatNumber(.ModifiedCorrelation) & Separator & _
                FormatNumber(.Determination))
            End With
            For Each CorrelationPoint In CorrelationPoints
              Result = _DBM.Result _
                (CorrelationPoint.PointDriver, Nothing, StartTimestamp)
              With Result.Prediction
                Console.Write(Separator & _
                  FormatNumber(.MeasuredValue) & Separator & _
                  FormatNumber(.PredictedValue) & Separator & _
                  FormatNumber(.LowerControlLimit) & Separator & _
                  FormatNumber(.UpperControlLimit))
              End With
            Next
          End If
          Console.WriteLine
          ' Next interval
          StartTimestamp = StartTimestamp.AddSeconds(CalculationInterval)
        Loop
      End If
    End Sub

  End Class

End Namespace
