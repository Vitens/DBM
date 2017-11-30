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
Imports System.Runtime.InteropServices
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.PI
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMDataRef")>


Namespace Vitens.DynamicBandwidthMonitor


  <Description("DBMDataRef;Dynamic Bandwidth Monitor")> _
    <Guid("e092c5ed-888b-4dbd-89eb-45206d77db9a")> _
    Public Class DBMDataRef
    Inherits AFDataReference


    Private _DBM As New DBM
    Private InputPointDriver As DBMPointDriver
    Private CorrelationPoints As List(Of DBMCorrelationPoint)


    Public Overrides Readonly Property SupportedContexts _
      As AFDataReferenceContext

      Get
        Return AFDataReferenceContext.Time
      End Get

    End Property


    Public Overrides Readonly Property SupportedDataMethods As AFDataMethods

      Get
        Return AFDataMethods.RecordedValue Or AFDataMethods.RecordedValues Or _
          AFDataMethods.PlotValues Or AFDataMethods.Summary Or _
          AFDataMethods.Summaries
      End Get

    End Property


    Public Overrides Readonly Property SupportedMethods As AFDataReferenceMethod

      Get
        Return AFDataReferenceMethod.GetValue Or AFDataReferenceMethod.GetValues
      End Get

    End Property


    Public Overrides Property ConfigString As String

      Get
        Return DBM.Version
      End Get

      Set
      End Set

    End Property


    Private Function StringToPIPoint(Point As String) As PIPoint

      Return PIPoint.FindPIPoints(PIServer.FindPIServer(Point.Split("\"c)(2). _
        Split("?"c)(0)), Point.Split("\"c)(3).Split("?"c)(0))(0)

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
              (StringToPIPoint(ParentElement.Attributes(Attribute.Parent. _
              Name).ConfigString)), True))
          End If
          ParentElement = ParentElement.Parent
        Loop

      End If

    End Sub


    Public Overrides Function GetValue(context As Object, _
      timeContext As Object, inputAttributes As AFAttributeList, _
      inputValues As AFValues) As AFValue

      Dim Timestamp As AFTime
      Dim Result As DBMResult
      Dim Value As Double

      If InputPointDriver Is Nothing Then GetInputAndCorrelationPoints

      If timeContext Is Nothing Then
        Timestamp = DirectCast(InputPointDriver.Point, PIPoint). _
          CurrentValue.Timestamp
      Else
        Timestamp = DirectCast(timeContext, AFTime)
      End If

      Result = _DBM.Result(InputPointDriver, CorrelationPoints, _
        Timestamp.LocalTime)

      If Attribute.Name.Equals("Factor") Then
        Value = Result.Factor
      ElseIf Attribute.Name.Equals("MeasuredValue") Then
        Value = Result.PredictionData.MeasuredValue
      ElseIf Attribute.Name.Equals("PredictedValue") Then
        Value = Result.PredictionData.PredictedValue
      ElseIf Attribute.Name.Equals("LowerControlLimit") Then
        Value = Result.PredictionData.LowerControlLimit
      ElseIf Attribute.Name.Equals("UpperControlLimit") Then
        Value = Result.PredictionData.UpperControlLimit
      End If

      Return New AFValue(Value, Result.Timestamp)

    End Function


    Public Overrides Function GetValues(context As Object, _
      timeContext As AFTimeRange, numberOfValues As Integer, _
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      If InputPointDriver Is Nothing Then GetInputAndCorrelationPoints

      _DBM.PrepareData(InputPointDriver, CorrelationPoints, _
        timeContext.StartTime.LocalTime, timeContext.EndTime.LocalTime)

      GetValues = New AFValues
      Do While timeContext.EndTime > timeContext.StartTime
        GetValues.Add(GetValue(Nothing, timeContext.StartTime, _
          Nothing, Nothing))
        timeContext.StartTime = New AFTime _
          (timeContext.StartTime.UtcSeconds+CalculationInterval)
      Loop

      Return GetValues

    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange, _
      boundaryType As AFBoundaryType, filterExpression As String, _
      includeFilteredValues As Boolean, inputAttributes As AFAttributeList, _
      inputValues As AFValues(), inputTimes As List(Of AFTime), _
      Optional maxCount As Integer = 0) As AFValues

      Return GetValues(Nothing, timeRange, Nothing, Nothing, Nothing)

    End Function


  End Class


End Namespace
