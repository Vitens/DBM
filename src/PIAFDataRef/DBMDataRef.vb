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
Imports System.DateTime
Imports System.DateTimeKind
Imports System.Environment
Imports System.Math
Imports System.Runtime.InteropServices
Imports System.TimeSpan
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
    Const AttributeNameMeasuredValue As String = "MeasuredValue"
    Const AttributeNamePredictedValue As String = "PredictedValue"
    Const AttributeNameLowerControlLimit As String = "LowerControlLimit"
    Const AttributeNameUpperControlLimit As String = "UpperControlLimit"

    Private _DBM As New DBM
    Private InputPointDriver As DBMPointDriver
    Private CorrelationPoints As List(Of DBMCorrelationPoint)


    Public Overrides Readonly Property SupportedMethods As AFDataReferenceMethod

      ' This read-only property specifies which of the data reference
      ' methods are supported.

      Get

        Return AFDataReferenceMethod.GetValue Or AFDataReferenceMethod.GetValues

      End Get

    End Property


    Public Overrides Readonly Property SupportedDataMethods As AFDataMethods

      ' This read-only property specifies which of the data methods are
      ' supported by the data reference.

      Get

        Return AFDataMethods.RecordedValue Or AFDataMethods.RecordedValues Or _
          AFDataMethods.PlotValues Or AFDataMethods.Summary Or _
          AFDataMethods.Summaries

      End Get

    End Property


    Public Overrides Readonly Property SupportedContexts _
      As AFDataReferenceContext

      ' This read-only property specifies which of the data reference
      ' contexts are supported when getting and/or setting values.

      Get

        Return AFDataReferenceContext.Time

      End Get

    End Property


    Public Overrides Property ConfigString As String

      ' This property returns the current configuration of the attribute's
      ' data reference as a string suitable for displaying to an end-user.

      Get

        Return DBM.Version

      End Get

      Set
      End Set

    End Property


    Private Function StringToPIPoint(Point As String) As PIPoint

      Dim SplitItems As Char() = {"\"c}
      Dim SplitGUID As Char() = {"?"c}

      Return PIPoint.FindPIPoints(PIServer.FindPIServer( _
        Point.Split(SplitItems)(2).Split(SplitGUID)(0)), _
        Point.Split(SplitItems)(3).Split(SplitGUID)(0))(0)

    End Function


    Private Sub GetInputAndCorrelationPoints

      Dim Element, ParentElement, SiblingElement As AFElement

      If Attribute IsNot Nothing Then ' If owned by an attribute

        Element = DirectCast(Attribute.Element, AFElement)
        ParentElement = Element.Parent
        InputPointDriver = New DBMPointDriver(StringToPIPoint _
          (Attribute.Parent.ConfigString)) ' Parent attribute
        CorrelationPoints = New List(Of DBMCorrelationPoint)

        ' Find siblings
        If ParentElement IsNot Nothing Then
          For Each SiblingElement In ParentElement.Elements
            If Not SiblingElement.UniqueID.Equals(Element.UniqueID) And _
              SiblingElement.Template.UniqueID.Equals _
              (Element.Template.UniqueID) Then ' Same template, skip self
              CorrelationPoints.Add(New DBMCorrelationPoint _
                (New DBMPointDriver(StringToPIPoint(SiblingElement.
                Attributes(Attribute.Parent.Name).ConfigString)), False))
            End If
          Next
        End If

        ' Find parents recursively
        Do While ParentElement IsNot Nothing
          If ParentElement.Template.UniqueID.Equals _
            (Element.Template.UniqueID) Then ' Same template
            CorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver _
              (StringToPIPoint(ParentElement.Attributes _
              (Attribute.Parent.Name).ConfigString)), True))
          End If
          ParentElement = ParentElement.Parent
        Loop

      End If

    End Sub


    Private Function CurrentTimestamp(PointDriver As DBMPointDriverAbstract) _
      As DateTime

      Return DirectCast(PointDriver.Point, PIPoint).CurrentValue. _
        Timestamp.LocalTime

    End Function


    Private Function AlignTime(Timestamp As DateTime) As DateTime

      Return Timestamp.AddSeconds(-Timestamp.Ticks/TicksPerSecond _
        Mod CalculationInterval)

    End Function


    Public Overrides Function GetValue(context As Object, _
      timeContext As Object, inputAttributes As AFAttributeList, _
      inputValues As AFValues) As AFValue

      ' This method gets the value based upon the data reference
      ' configuration within the specified context.

      Dim Timestamp As DateTime
      Dim CorrelationPoint As DBMCorrelationPoint
      Dim Result As DBMResult
      Dim Value As Double

      If InputPointDriver Is Nothing Then GetInputAndCorrelationPoints

      If timeContext Is Nothing Then
        ' No time was specified. Use the latest possible timestamp based on the
        ' current timestamp of the input and correlation points.
        Timestamp = CurrentTimestamp(InputPointDriver)
        For Each CorrelationPoint In CorrelationPoints ' Find earliest
          Timestamp = SpecifyKind(New DateTime(Min(Timestamp.Ticks, _
            CurrentTimestamp(CorrelationPoint.PointDriver).Ticks)), Local)
        Next
        ' Align timestamp to previous interval and subtract one interval
        Timestamp = AlignTime(Timestamp).AddSeconds(-CalculationInterval)
      Else
        Timestamp = AlignTime(DirectCast(timeContext, AFTime).LocalTime)
      End If

      Result = _DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)

      If Attribute.Name.Equals(AttributeNameFactor) Then
        Value = Result.Factor
      ElseIf Attribute.Name.Equals(AttributeNameMeasuredValue) Then
        Value = Result.PredictionData.MeasuredValue
      ElseIf Attribute.Name.Equals(AttributeNamePredictedValue) Then
        Value = Result.PredictionData.PredictedValue
      ElseIf Attribute.Name.Equals(AttributeNameLowerControlLimit) Then
        Value = Result.PredictionData.LowerControlLimit
      ElseIf Attribute.Name.Equals(AttributeNameUpperControlLimit) Then
        Value = Result.PredictionData.UpperControlLimit
      End If

      Return New AFValue(Value, New AFTime(Timestamp))

    End Function


    Public Overrides Function GetValues(context As Object, _
      timeContext As AFTimeRange, numberOfValues As Integer, _
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      ' This method gets a collection of AFValue objects for an attribute
      ' based upon the data reference configuration within the specified
      ' AFTimeRange context.

      Dim Intervals As Integer
      Dim IntervalStep, Interval As Double
      Dim Values As New AFValues

      timeContext.StartTime = New AFTime(AlignTime _
        (timeContext.StartTime.LocalTime)) ' Align timestamp
      Intervals = Max(1, CInt(AlignTime(timeContext.EndTime.LocalTime). _
        AddSeconds(CalculationInterval).Subtract(AlignTime _
        (timeContext.StartTime.LocalTime)).TotalSeconds/CalculationInterval-1))
      If numberOfValues = 0 Then numberOfValues = Intervals
      numberOfValues = Min(numberOfValues, Intervals)
      IntervalStep = Intervals/numberOfValues

      Do While Interval < Intervals ' Loop through intervals
        Values.Add(GetValue(Nothing, New AFTime _
          (timeContext.StartTime.LocalTime.AddSeconds _
          (CInt(Interval)*CalculationInterval)), Nothing, Nothing))
        Interval += IntervalStep
      Loop

      Return Values

    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange, _
      boundaryType As AFBoundaryType, filterExpression As String, _
      includeFilteredValues As Boolean, inputAttributes As AFAttributeList, _
      inputValues As AFValues(), inputTimes As List(Of AFTime), _
      Optional maxCount As Integer = 0) As AFValues

      ' Returns a list of compressed values for the requested time range from
      ' the source provider.

      Return GetValues(Nothing, timeRange, maxCount, Nothing, Nothing)

    End Function


  End Class


End Namespace
