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
Imports System.Environment
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


Namespace Vitens.DynamicBandwidthMonitor


  <Description("DBMDataRef;Dynamic Bandwidth Monitor")> _
    <Guid("e092c5ed-888b-4dbd-89eb-45206d77db9a")> _
    Public Class DBMDataRef
    Inherits AFDataReference


    Const AttributeNameFactor As String = "Factor"
    Const AttributeNameMeasuredValue As String = "Meetwaarde"
    Const AttributeNamePredictedValue As String = "Voorspelling"
    Const AttributeNameLowerControlLimit As String = "Ondergrens"
    Const AttributeNameUpperControlLimit As String = "Bovengrens"


    Private CurrentAttribute As AFAttribute
    Private _DBM As New DBM
    Private InputPointDriver As DBMPointDriver
    Private CorrelationPoints As List(Of DBMCorrelationPoint)


    Private Function StringToPIPoint(Point As String) As PIPoint

      Return PIPoint.FindPIPoints(PIServer.FindPIServer(Point.Split _
        (New Char() {"\"c})(2).Split(New Char() {"?"c})(0)), Point.Split _
        (New Char() {"\"c})(3).Split(New Char() {"?"c})(0))(0)

    End Function


    Private Function PIPointToString(Point As PIPoint) As String

      Return "\\" & Point.Server.Name & "\" & Point.Name

    End Function


    Private Function AlignTime(Timestamp As DateTime, Interval As Integer) _
      As DateTime

      Return Timestamp.AddSeconds(-((Timestamp.Minute*60+Timestamp.Second) Mod _
        Interval+Timestamp.Millisecond/1000))

    End Function


    Private Function GetIntervals(TimeRange As AFTimeRange, _
      Interval As Integer) As Integer

      Return CInt(TimeRange.EndTime.LocalTime.Subtract _
        (TimeRange.StartTime.LocalTime).TotalSeconds/Interval)

    End Function


    Public Overrides Property Attribute As AFAttribute

      Get
        Dim Element, SiblingElement, ParentElement As AFElement
        Element = DirectCast(CurrentAttribute.Element, AFElement)
        InputPointDriver = New DBMPointDriver(StringToPIPoint _
          (CurrentAttribute.Parent.ConfigString))
        CorrelationPoints = New List(Of DBMCorrelationPoint)
        For Each SiblingElement in Element.Parent.Elements
          If Not SiblingElement.UniqueID.Equals(Element.UniqueID) Then
            CorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver _
              (StringToPIPoint(SiblingElement.Attributes  _
              (CurrentAttribute.Parent.Name).ConfigString)), False))
          End If
        Next
        ParentElement = Element
        Do While ParentElement.Parent IsNot Nothing
          ParentElement = ParentElement.Parent
          CorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver _
            (StringToPIPoint(ParentElement.Attributes _
            (CurrentAttribute.Parent.Name).ConfigString)), True))
        Loop
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
        Return AFDataMethods.RecordedValues Or AFDataMethods.PlotValues Or _
          AFDataMethods.Summary Or AFDataMethods.Summaries
      End Get

    End Property


    Public Overrides Readonly Property SupportedContexts _
      As AFDataReferenceContext

      Get
        Return AFDataReferenceContext.Time
      End Get

    End Property


    Public Overrides Property ConfigString As String

      Get
        Dim CorrelationPoint As DBMCorrelationPoint
        ConfigString = CurrentAttribute.Name & NewLine & _
          CurrentAttribute.Parent.Name & ": " & PIPointToString(DirectCast _
            (InputPointDriver.Point, PIPoint)) & NewLine
        For Each CorrelationPoint In CorrelationPoints
          ConfigString &= If(CorrelationPoint.SubtractSelf, "P", "S") & ": " & _
            PIPointToString(DirectCast(CorrelationPoint.PointDriver.Point, _
            PIPoint)) & NewLine
        Next
        ConfigString &= NewLine & NewLine & DBM.Version
        Return ConfigString
      End Get

      Set
      End Set

    End Property


    Public Overrides Function GetValue(context As Object, _
      timeContext As Object, inputAttributes As AFAttributeList, _
      inputValues As AFValues) As AFValue

      Dim Timestamp As DateTime
      Dim Result As DBMResult
      Dim Value As Double

      If timeContext Is Nothing Then
        Timestamp = AlignTime(DirectCast(InputPointDriver.Point, PIPoint). _
          CurrentValue.Timestamp.LocalTime, CalculationInterval). _
          AddSeconds(-CalculationInterval)
      Else
        Timestamp = DirectCast(timeContext, AFTime).LocalTime
      End If
      Result = _DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)
      If CurrentAttribute.Name.Equals(AttributeNameFactor) Then
        Value = Result.Factor
      ElseIf CurrentAttribute.Name.Equals(AttributeNameMeasuredValue) Then
        Value = Result.PredictionData.MeasuredValue
      ElseIf CurrentAttribute.Name.Equals(AttributeNamePredictedValue) Then
        Value = Result.PredictionData.PredictedValue
      ElseIf CurrentAttribute.Name.Equals(AttributeNameLowerControlLimit) Then
        Value = Result.PredictionData.LowerControlLimit
      ElseIf CurrentAttribute.Name.Equals(AttributeNameUpperControlLimit) Then
        Value = Result.PredictionData.UpperControlLimit
      End If

      Return New AFValue(Value, New AFTime(Timestamp))

    End Function


    Public Overrides Function GetValues(context As Object, _
      timeContext As AFTimeRange, numberOfValues As Integer, _
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      Dim Intervals As Integer
      Dim IntervalStep, Interval As Double
      Dim Values As New AFValues

      timeContext.StartTime = New AFTime(AlignTime _
        (timeContext.StartTime.LocalTime, CalculationInterval))
      timeContext.EndTime = New AFTime(AlignTime _
        (timeContext.EndTime.LocalTime, CalculationInterval). _
        AddSeconds(CalculationInterval))
      Intervals = GetIntervals(timeContext, CalculationInterval)
      If numberOfValues = 0 Then numberOfValues = Intervals
      numberOfValues = Min(numberOfValues, Intervals)
      IntervalStep = Intervals/numberOfValues
      Do While Interval < Intervals
        Values.Add(GetValue(Nothing, New AFTime _
          (timeContext.StartTime.LocalTime.AddSeconds _
          (Floor(Interval)*CalculationInterval)), Nothing, Nothing))
        Interval += IntervalStep
      Loop

      Return Values

    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange, _
      boundaryType As AFBoundaryType, filterExpression As String, _
      includeFilteredValues As Boolean, inputAttributes As AFAttributeList, _
      inputValues As AFValues(), inputTimes As List(Of AFTime), _
      Optional maxCount As Integer = 0) As AFValues

      Return GetValues(Nothing, timeRange, maxCount, Nothing, Nothing)

    End Function


    Public Overrides Function PlotValues(timeRange As AFTimeRange, _
      intervals As Integer, inputAttributes As AFAttributeList, _
      inputValues As AFValues(), inputTimes As List(Of AFTime)) As AFValues

      Return GetValues(Nothing, timeRange, intervals, Nothing, Nothing)

    End Function


    Public Overrides Function Summary(timeRange As AFTimeRange, _
      summaryType As AFSummaryTypes, calcBasis As AFCalculationBasis, _
      timeType As AFTimestampCalculation) As _
      IDictionary(Of AFSummaryTypes, AFValue)

      Dim returnValue As New Dictionary(Of AFSummaryTypes, AFValue)

      returnValue.Add(AFSummaryTypes.Count, New AFValue(GetIntervals _
        (timeRange, CalculationInterval), New AFTime(DateTime.Now)))

      Return returnValue

    End Function


    Public Overrides Function Summaries(timeRange As AFTimeRange, _
      summaryDuration As AFTimeSpan, summaryType As AFSummaryTypes, _
      calcBasis As AFCalculationBasis, timeType As AFTimestampCalculation) _
      As IDictionary(Of AFSummaryTypes, AFValues)

      Dim Values As New AFValues
      Dim returnValues As New Dictionary(Of AFSummaryTypes, AFValues)

      Values.Add(Summary(timeRange, summaryType, calcBasis, timeType) _
        (AFSummaryTypes.Count))
      returnValues.Add(AFSummaryTypes.Count, Values)

      Return returnValues

    End Function


  End Class


End Namespace
