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
Imports System.ComponentModel
Imports System.Math
Imports System.Runtime.InteropServices
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.PI
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMDataRef")>


Namespace DBMDataRef ' TODO Name?


  <Description("DBMDataRef;Dynamic Bandwidth Monitor")> _
  <Guid("e092c5ed-888b-4dbd-89eb-45206d77db9a")> _
  Public Class DBMDataRef


  Inherits AFDataReference


    Const ATTNAMEFACTOR As String = "Factor"
    Const ATTNAMEMEASUREDVALUE As String = "Meetwaarde"
    Const ATTNAMEPREDICTEDVALUE As String = "Voorspelling"
    Const ATTNAMELOWERCONTROLLIMIT As String = "Ondergrens"
    Const ATTNAMEUPPERCONTROLLIMIT As String = "Bovengrens"


    Private CurrentAttribute As AFAttribute
    Private InputPointDriver As DBMPointDriver
    Private CorrelationPoints As New List(Of DBMCorrelationPoint) ' TODO This has to be filled with parents (containing) and siblings (surrounding)
    Private _DBM As New DBM


    Public Overrides Property Attribute As AFAttribute
      Get
        InputPointDriver = New DBMPointDriver(PIPoint.FindPIPoints(PIServer.FindPIServer(CurrentAttribute.Parent.ConfigString.Split(New Char() {"\"c})(2).Split(New Char() {"?"c})(0)), CurrentAttribute.Parent.ConfigString.Split(New Char() {"\"c})(3).Split(New Char() {"?"c})(0))(0))
        Return CurrentAttribute
      End Get
      Set(SetAttribute As AFAttribute)
        CurrentAttribute = SetAttribute
      End Set
    End Property


    Public Overrides Readonly Property SupportedMethods As AFDataReferenceMethod
      Get
        Return AFDataReferenceMethod.GetValue Or AFDataReferenceMethod.GetValues
      End Get
    End Property


    Public Overrides Readonly Property SupportedDataMethods As AFDataMethods
      Get
        Return AFDataMethods.RecordedValues Or AFDataMethods.PlotValues
      End Get
    End Property


    Public Overrides Readonly Property SupportedContexts As AFDataReferenceContext
      Get
        Return AFDataReferenceContext.Time
      End Get
    End Property


    Public Overrides Property ConfigString As String
      Get
        Return CurrentAttribute.Name & NewLine & _
          CurrentAttribute.Parent.Name & ": \\" & DirectCast(InputPointDriver.Point, PIPoint).Server.Name & "\" & DirectCast(InputPointDriver.Point, PIPoint).Name & NewLine & _
          NewLine & NewLine & DBM.Version
      End Get
      Set
      End Set
    End Property


    Public Overrides Function GetValue(context As Object, timeContext As Object, inputAttributes As AFAttributeList, inputValues As AFValues) As AFValue
      Dim Timestamp As DateTime
      Dim Result As DBMResult
      Dim Value As Double
      If timeContext Is Nothing Then
        Timestamp = DirectCast(InputPointDriver.Point, PIPoint).CurrentValue.Timestamp.LocalTime
        Timestamp = Timestamp.Subtract(New TimeSpan(0, 0, CInt((Timestamp-DateTime.MinValue).TotalSeconds Mod CalculationInterval+CalculationInterval)))
      Else
        Timestamp = DirectCast(timeContext, AFTime).LocalTime
      End If
      ' TODO Timestamp has millisecond resolution so not exactly aligned on intervals
      Result = _DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)
      Select Case CurrentAttribute.Name
        Case ATTNAMEFACTOR
          Value = Result.Factor
        Case ATTNAMEMEASUREDVALUE
          Value = Result.PredictionData.MeasuredValue
        Case ATTNAMEPREDICTEDVALUE
          Value = Result.PredictionData.PredictedValue
        Case ATTNAMELOWERCONTROLLIMIT
          Value = Result.PredictionData.LowerControlLimit
        Case ATTNAMEUPPERCONTROLLIMIT
          Value = Result.PredictionData.UpperControlLimit
      End Select
      Return New AFValue(Value, New AFTime(Timestamp))
    End Function


    Public Overrides Function GetValues(context As Object, timeContext As AFTimeRange, numberOfValues As Integer, inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues
      Dim Intervals As Integer
      Dim IntervalStep, Interval As Double
      Dim Values As New AFValues
      timeContext.StartTime = New AFTime(timeContext.StartTime.LocalTime.Subtract(New TimeSpan(0, 0, CInt((timeContext.StartTime.LocalTime-DateTime.MinValue).TotalSeconds Mod CalculationInterval))))
      timeContext.EndTime = New AFTime(timeContext.EndTime.LocalTime.Subtract(New TimeSpan(0, 0, CInt((timeContext.EndTime.LocalTime-DateTime.MinValue).TotalSeconds Mod CalculationInterval-CalculationInterval))))
      ' TODO Create function for aligning timestamps
      Intervals = CInt(timeContext.EndTime.LocalTime.Subtract(timeContext.StartTime.LocalTime).TotalSeconds/CalculationInterval)
      If numberOfValues = 0 Then numberOfValues = Intervals
      numberOfValues = Min(numberOfValues, Intervals)
      IntervalStep = Intervals/numberOfValues
      Do While Interval < Intervals
        Values.Add(GetValue(Nothing, New AFTime(timeContext.StartTime.LocalTime.AddSeconds(Floor(Interval)*CalculationInterval)), Nothing, Nothing))
        Interval += IntervalStep
      Loop
      Return Values
    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange, boundaryType As AFBoundaryType, filterExpression As String, includeFilteredValues As Boolean, inputAttributes As AFAttributeList, inputValues As AFValues(), inputTimes As List(Of AFTime), Optional maxCount As Integer = 0) As AFValues
      Return GetValues(Nothing, timeRange, maxCount, Nothing, Nothing)
    End Function


    Public Overrides Function PlotValues(timeRange As AFTimeRange, intervals As Integer, inputAttributes As AFAttributeList, inputValues As AFValues(), inputTimes As List(Of AFTime)) As AFValues
      Return GetValues(Nothing, timeRange, intervals, Nothing, Nothing)
    End Function


  End Class


End Namespace
