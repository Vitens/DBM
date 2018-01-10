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
Imports System.ComponentModel
Imports System.DateTime
Imports System.Math
Imports System.Runtime.InteropServices
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.PI
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMDataRef")>


Namespace Vitens.DynamicBandwidthMonitor


  <Description("DBMDataRef;Dynamic Bandwidth Monitor")> _
    <Guid("e092c5ed-888b-4dbd-89eb-45206d77db9a")> _
    Public Class DBMDataRef
    Inherits AFDataReference


    ' DBMDataRef is a custom OSIsoft PI Asset Framework data reference which
    ' integrates DBM with PI AF. The build script automatically registers the
    ' data reference and support assemblies when run on the PI AF server.
    ' The data reference uses the PI tag from the parent attribute as input and
    ' automatically uses PI tags from sibling and parent elements based on the
    ' same template for correlation calculations, unless the NoCorrelation
    ' category is applied to the attribute. The value returned from the DBM
    ' calculation is determined by the applied category (Factor, MeasuredValue,
    ' PredictedValue, LowerControlLimit or UpperControlLimit).


    Const CategoryNoCorrelation As String = "NoCorrelation"
    Const CategoryReturnFactor As String = "Factor"
    Const CategoryReturnMeasuredValue As String = "MeasuredValue"
    Const CategoryReturnPredictedValue As String = "PredictedValue"
    Const CategoryReturnLowerControlLimit As String = "LowerControlLimit"
    Const CategoryReturnUpperControlLimit As String = "UpperControlLimit"


    Private Shared _DBM As New DBM
    Private LastGetPointsTime As DateTime
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
        Return DBM.LicenseNotice
      End Get

      Set
      End Set

    End Property


    Private Sub GetInputAndCorrelationPoints

      ' Retrieve input and correlation PI points from AF hierarchy. Recheck for
      ' changes after every calculation interval and only if owned by an
      ' attribute (element is an instance of an element template) and attribute
      ' has a parent attribute referring to an input PI point.

      Dim Element, ParentElement, SiblingElement As AFElement

      If Now >= AlignTimestamp(LastGetPointsTime, CalculationInterval). _
        AddSeconds(CalculationInterval) And _
        Attribute IsNot Nothing And Attribute.Parent IsNot Nothing Then

        LastGetPointsTime = Now
        Element = DirectCast(Attribute.Element, AFElement)
        ParentElement = Element.Parent
        InputPointDriver = New DBMPointDriver(Attribute.Parent.PIPoint) ' Parent
        CorrelationPoints = New List(Of DBMCorrelationPoint)

        ' Retrieve correlation points only when calculating the DBM factor value
        ' and if the correlation calculations are not disabled using categories.
        If Attribute.CategoriesString.Contains(CategoryReturnFactor) And _
          Not Attribute.CategoriesString.Contains(CategoryNoCorrelation) Then

          ' Find siblings
          If ParentElement IsNot Nothing Then
            For Each SiblingElement In ParentElement.Elements
              If Not SiblingElement.UniqueID.Equals(Element.UniqueID) And _
                SiblingElement.Template IsNot Nothing AndAlso _
                SiblingElement.Template.UniqueID.Equals _
                (Element.Template.UniqueID) Then ' Same template, skip self
                Try ' Catch exception if PIPoint does not exist
                  CorrelationPoints.Add(New DBMCorrelationPoint(New _
                    DBMPointDriver(SiblingElement.Attributes(Attribute.Parent. _
                    Name).PIPoint), False))
                Catch
                End Try
              End If
            Next
          End If

          ' Find parents recursively
          Do While ParentElement IsNot Nothing
            If ParentElement.Template IsNot Nothing AndAlso _
              ParentElement.Template.UniqueID.Equals _
              (Element.Template.UniqueID) Then ' Same template
              Try ' Catch exception if PIPoint does not exist
                CorrelationPoints.Add(New DBMCorrelationPoint _
                  (New DBMPointDriver(ParentElement.Attributes _
                  (Attribute.Parent.Name).PIPoint), True))
              Catch
              End Try
            End If
            ParentElement = ParentElement.Parent
          Loop

        End If

      End If

    End Sub


    Public Overrides Function GetValue(context As Object, _
      timeContext As Object, inputAttributes As AFAttributeList, _
      inputValues As AFValues) As AFValue

      Dim Timestamp As AFTime
      Dim Result As DBMResult
      Dim Value As New AFValue

      GetInputAndCorrelationPoints

      If timeContext Is Nothing Then
        Timestamp = DirectCast(InputPointDriver.Point, PIPoint). _
          CurrentValue.Timestamp
      Else
        Timestamp = DirectCast(timeContext, AFTime)
      End If

      Result = _DBM.Result(InputPointDriver, CorrelationPoints, _
        Timestamp.LocalTime)

      ' Return value based on applied category.
      With Result.PredictionData
        If Attribute.CategoriesString. _
          Contains(CategoryReturnFactor) Then
          Value = New AFValue(Result.Factor, Result.Timestamp)
          Value.Questionable = Abs(Result.Factor) > 1 ' Unsuppressed exception
          Value.Substituted = Abs(Result.Factor) > 0 And _
            Abs(Result.Factor) <= 1 ' Suppressed exception
        ElseIf Attribute.CategoriesString. _
          Contains(CategoryReturnMeasuredValue) Then
          Value = New AFValue(.MeasuredValue, Result.Timestamp)
        ElseIf Attribute.CategoriesString. _
          Contains(CategoryReturnPredictedValue) Then
          Value = New AFValue(.PredictedValue, Result.Timestamp)
        ElseIf Attribute.CategoriesString. _
          Contains(CategoryReturnLowerControlLimit) Then
          Value = New AFValue(.LowerControlLimit, Result.Timestamp)
        ElseIf Attribute.CategoriesString. _
          Contains(CategoryReturnUpperControlLimit) Then
          Value = New AFValue(.UpperControlLimit, Result.Timestamp)
        End If
      End With

      Return Value

    End Function


    Public Overrides Function GetValues(context As Object, _
      timeContext As AFTimeRange, numberOfValues As Integer, _
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      ' Returns values for each interval in a time range. The (aligned) end time
      ' itself is excluded.

      Dim IntervalSeconds As Double

      GetInputAndCorrelationPoints

      _DBM.PrepareData(InputPointDriver, CorrelationPoints, _
        timeContext.StartTime.LocalTime, timeContext.EndTime.LocalTime)

      GetValues = New AFValues
      IntervalSeconds = Max(1, ((timeContext.EndTime.UtcSeconds-timeContext. _
        StartTime.UtcSeconds)/CalculationInterval-1)/(numberOfValues-1))* _
        CalculationInterval ' Required interval, first and last interv inclusive
      Do While timeContext.EndTime > timeContext.StartTime
        GetValues.Add(GetValue(Nothing, timeContext.StartTime, _
          Nothing, Nothing))
        timeContext.StartTime = New AFTime _
          (timeContext.StartTime.UtcSeconds+IntervalSeconds)
      Loop

      Return GetValues

    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange, _
      boundaryType As AFBoundaryType, filterExpression As String, _
      includeFilteredValues As Boolean, inputAttributes As AFAttributeList, _
      inputValues As AFValues(), inputTimes As List(Of AFTime), _
      Optional maxCount As Integer = 0) As AFValues

      Return GetValues(Nothing, timeRange, maxCount, Nothing, Nothing)

    End Function


  End Class


End Namespace
