Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
Imports System.Double
Imports System.Math
Imports System.Runtime.InteropServices
Imports System.Threading
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Asset.AFAttributeTrait
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMInfo
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMMisc
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMDataRef")>


Namespace Vitens.DynamicBandwidthMonitor


  <Description("DBMDataRef;Dynamic Bandwidth Monitor")>
    <Guid("e092c5ed-888b-4dbd-89eb-45206d77db9a")>
    Public Class DBMDataRef
    Inherits AFDataReference


    ' DBMDataRef is a custom OSIsoft PI Asset Framework data reference which
    ' integrates DBM with PI AF. The build script automatically registers the
    ' data reference and support assemblies when run on the PI AF server.
    ' The data reference uses the parent attribute as input and automatically
    ' uses attributes from sibling and parent elements based on the same
    ' template containing good data for correlation calculations, unless the
    ' NoCorrelation category is applied to the output attribute. The value
    ' returned from the DBM calculation is determined by the applied
    ' property/trait:
    '   None      Factor
    '   Target    Measured value
    '   Forecast  Forecast value
    '   Minimum   Lower control limit (p = 0.9999)
    '   LoLo      Lower control limit (default)
    '   Lo        Lower control limit (p = 0.95)
    '   Hi        Upper control limit (p = 0.95)
    '   HiHi      Upper control limit (default)
    '   Maximum   Upper control limit (p = 0.9999)


    Const CategoryNoCorrelation As String = "NoCorrelation"
    Const pValueLoHi As Double = 0.95 ' Confidence interval for Lo and Hi
    Const pValueMinMax As Double = 0.9999 ' CI for Minimum and Maximum


    Private Shared DBMShared As New DBM ' Use shared DBM object, if available.
    Private DBMNonShared As New DBM ' Fall back to nonshared object.


    Public Overrides Readonly Property SupportedContexts _
      As AFDataReferenceContext

      ' This read-only property specifies which of the data reference contexts
      ' are supported when getting and/or setting values.
      ' The context normally applies to the GetValue, SetValue, and GetValues
      ' methods. However, the TimeRange context applies only to GetValue and
      ' indicates the data reference can return a single value for the entire
      ' range, such as a totalization.

      Get
        ' Specifies which of the data reference contexts are supported. Its
        ' value can be one or more of the AFDataReferenceContext enumeration
        ' values logically ORed together.
        Return AFDataReferenceContext.Time
      End Get

    End Property


    Public Overrides Readonly Property SupportedDataMethods As AFDataMethods

      ' This read-only property specifies which of the data methods are
      ' supported by the data reference.
      ' If a data reference implementation supports any data methods, then it
      ' should override this property and indicate which data methods are
      ' supported. If an implementation indicates that a data method is
      ' supported but does not override the data method, then the base
      ' implementation will be used.

      Get
        ' Specifies which of the data methods are supported by the data
        ' reference. Its value can be one or more of the AFDataMethods
        ' enumeration values logically ORed together. The default value is None.
        Return AFDataMethods.RecordedValue Or AFDataMethods.RecordedValues Or
          AFDataMethods.PlotValues Or AFDataMethods.Summary Or
          AFDataMethods.Summaries
      End Get

    End Property


    Public Overrides Readonly Property SupportedMethods As AFDataReferenceMethod

      ' This read-only property specifies which of the data reference methods
      ' are supported.
      ' If a data reference is intended to be read-only, it should not support
      ' the SetValue method. Normally, to support GetValues, the data reference
      ' must support the context of Time, or be based on another attribute which
      ' supports the GetValues method. Data References which specify the
      ' CreateConfig option will have the user interface make available a
      ' command to inform the data reference to update the foreign database's
      ' configuration. For example, the PI Point Data Reference can be made to
      ' update the tag configuration of the specified point.

      Get
        ' Specifies which of the data reference methods are supported. Its value
        ' can be one or more of the AFDataReferenceMethod enumeration values
        ' logically ORed together.
        Return AFDataReferenceMethod.GetValue Or AFDataReferenceMethod.GetValues
      End Get

    End Property


    Public Overrides Readonly Property [Step] As Boolean

      ' This property returns True if the value returned for this
      ' AFDataReference is stepped.
      ' The step attribute defines how the values are to be interpolated. When
      ' set to False, the values are treated as a continuous signal, and
      ' adjacent values are linearly interpolated. When set to True, the values
      ' are treated discretely and adjacent values are not interpolated.

      Get
        ' Returns True if the value for this AFDataReference is stepped.
        Return True
      End Get

    End Property


    Public Overrides Property ConfigString As String

      ' This property represents the current configuration of the Data Reference
      ' as a string suitable for displaying to an end-user and/or setting its
      ' configuration.
      ' Setting this property to Nothing or an empty string resets the
      ' configuration to its default values.

      Get
        ' Returns the current configuration of the Data Reference as a string
        ' suitable for displaying to an end-user. It must fully represent all
        ' the configuration information in a format that can also be used to set
        ' the configuration using this property.
        Return LicenseNotice
      End Get

      Set
      End Set

    End Property


    Public Overrides Function GetValue(context As Object,
      timeContext As Object, inputAttributes As AFAttributeList,
      inputValues As AFValues) As AFValue

      ' Returns a value for a single timestamp. Calls the GetValues method with
      ' aligned timestamps for results. If no result is available, NoSample is
      ' returned.

      ' This method gets the value based upon the data reference configuration
      ' within the specified context.
      ' timeContext
      '   The time context. If the context is Nothing, then the most recent
      '   value is returned. If the context is an AFTime object, then the value
      '   at the specified time is returned. If the context is an AFTimeRange
      '   object, then the value for the specified time range is returned. For
      '   convenience, AF SDK will convert a time context of AFCase into a time
      '   range, time at case's end time, or Nothing as appropriate.
      '   The AF SDK will adjust the time context to send only contexts that are
      '   specified in the SupportedContexts property. If Time Range context is
      '   passed but not supported, then context will be changed to Time context
      '   at the end time of the range. If Time context is passed but not
      '   supported, then context will be changed to no context.
      ' The data reference configuration specifies how to read the value of the
      ' attribute. The UOM property specifies the actual UOM of the external
      ' data value referenced by this attribute. This method should not be
      ' invoked directly by the user. Instead, the user should call one of the
      ' AFAttribute.GetValue Overload methods which will in-turn, invoke this
      ' method.

      Dim Timestamp As AFTime = Now

      If timeContext IsNot Nothing Then
        Timestamp = DirectCast(timeContext, AFTime)
      End If
      Timestamp = New AFTime(AlignPreviousInterval(Timestamp.UtcSeconds,
        CalculationInterval)) ' Align

      ' Returns the value for the attribute.
      Return GetValues(Nothing, New AFTimeRange(Timestamp,
        New AFTime(Timestamp.UtcSeconds+CalculationInterval)), 1,
        Nothing, Nothing)(0) ' Request a single value

    End Function


    Public Overrides Function GetValues(context As Object,
      timeContext As AFTimeRange, numberOfValues As Integer,
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      ' Returns values for each interval in a time range. The (aligned) end time
      ' itself is excluded. Make sure a value for every timestamp in the time
      ' range is returned by appending NoSample digital state values if
      ' required. A call from GetValue will always result in an IntervalSeconds
      ' of 300, DBM.PrepareData will never be used, and a single value will be
      ' returned.

      ' This method gets a collection of AFValue objects for an attribute based
      ' upon the data reference configuration within the specified AFTimeRange
      ' context.
      ' timeContext
      '   The time-range to be used when getting the attribute's collection of
      '   values.
      ' numberOfValues
      '   The number of values desired. The behavior of GetValues varies based
      '   on the value of this parameter:
      '   If 0, then all recorded values within the timeRange will be returned
      '   with an interpolated value at the start and end time, if possible.
      '   If the number of values requested is less than zero, then if
      '   supported, the Data Reference will return evenly spaced interpolated
      '   values across the timeRange, with a value returned at both end points
      '   of the time range. For example, specifying -25 over a 24 hour period
      '   will produce an hourly value.
      '   If the number of values requested is greater than zero, the method
      '   will behave like the PlotValues method. This method is designed to
      '   return a set of values over a time period that will produce the most
      '   accurate plot while minimizing the amount of data returned. The number
      '   of intervals specifies the number of pixels that need to be
      '   represented in the time period. For each interval, the data available
      '   is examined and significant values are returned. Each interval can
      '   produce up to 5 values if they are unique, the first value in the
      '   interval, the last value, the highest value, the lowest value and at
      '   most one exceptional point (bad value and/or annotated value). If no
      '   data is available in an interval, then no value is generated for that
      '   interval. As a result, the method may return more or fewer values than
      '   the number of values requested, depending on the distribution of
      '   recorded values over the time range. For Data References which use
      '   multiple PI Point sources, the list of values returned in each
      '   interval may be greater than 5 depending upon the data reference
      '   implementation.
      ' The returned collection of AFValue objects represent the values of the
      ' attribute within the specified time-range context. This method should
      ' not be invoked directly by the user. Instead, the user should call the
      ' AFAttribute.GetValues method which will in-turn, invoke this method.

      Dim IntervalSeconds As Double
      Dim Element, ParentElement, SiblingElement As AFElement
      Dim InputPointDriver As DBMPointDriver
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)

      GetValues = New AFValues
      timeContext.StartTime = New AFTime(AlignPreviousInterval(
        timeContext.StartTime.UtcSeconds, CalculationInterval)) ' Align
      IntervalSeconds = PIAFIntervalSeconds(numberOfValues,
        timeContext.EndTime.UtcSeconds-timeContext.StartTime.UtcSeconds)

      ' Retrieve correlation PI points from AF hierarchy if owned by an
      ' attribute (element is an instance of an element template) and attribute
      ' has a parent attribute.
      If Attribute IsNot Nothing And Attribute.Parent IsNot Nothing Then

        Element = DirectCast(Attribute.Element, AFElement)
        InputPointDriver = New DBMPointDriver(Attribute.Parent) ' Parent attrib.

        ' Retrieve correlation points for non-root elements only when
        ' calculating the DBM factor value and if correlation calculations are
        ' not disabled using categories.
        If Not Element.IsRoot And Attribute.Trait Is Nothing And
          Not Attribute.CategoriesString.Contains(CategoryNoCorrelation) Then

          ParentElement = Element.Parent

          ' Find siblings
          If ParentElement IsNot Nothing Then
            For Each SiblingElement In ParentElement.Elements
              If Not SiblingElement.UniqueID.Equals(Element.UniqueID) And
                SiblingElement.Template IsNot Nothing AndAlso
                SiblingElement.Template.UniqueID.Equals(
                Element.Template.UniqueID) Then ' Same template, skip self
                If SiblingElement.Attributes(Attribute.Parent.Name).
                  GetValue.IsGood Then ' Add only if has good data
                  CorrelationPoints.Add(New DBMCorrelationPoint(
                    New DBMPointDriver(SiblingElement.Attributes(
                    Attribute.Parent.Name)), False))
                End If
              End If
            Next
          End If

          ' Find parents recursively
          Do While ParentElement IsNot Nothing
            If ParentElement.Template IsNot Nothing AndAlso
              ParentElement.Template.UniqueID.Equals(
              Element.Template.UniqueID) Then ' Same template
              If ParentElement.Attributes(Attribute.Parent.Name).
                GetValue.IsGood Then ' Add only if has good data
                CorrelationPoints.Add(New DBMCorrelationPoint(
                  New DBMPointDriver(ParentElement.Attributes(
                  Attribute.Parent.Name)), True))
              End If
            End If
            ParentElement = ParentElement.Parent
          Loop

        End If

        For Each DBM In {DBMShared, DBMNonShared, New DBM} ' Prefer shared obj.

          If Monitor.TryEnter(DBM, TimeSpan.FromSeconds(
            Sqrt(CalculationInterval)/2*(1+RandomNumber(0, 1000)/1000))) Then
            Try

              If PIAFShouldPrepareData(timeContext.EndTime.UtcSeconds-
                timeContext.StartTime.UtcSeconds) Then DBM.PrepareData(
                InputPointDriver, CorrelationPoints,
                timeContext.StartTime.LocalTime, timeContext.EndTime.LocalTime)

              Do While timeContext.EndTime > timeContext.StartTime

                With DBM.Result(InputPointDriver, CorrelationPoints,
                  timeContext.StartTime.LocalTime)

                  If Attribute.Trait Is LimitTarget Then
                    GetValues.Add(New AFValue(.ForecastItem.Measurement,
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is Forecast Then
                    GetValues.Add(New AFValue(.ForecastItem.ForecastValue,
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is LimitMinimum Then
                    GetValues.Add(New AFValue(.ForecastItem.ForecastValue-
                      .ForecastItem.Range(pValueMinMax),
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is LimitLoLo Then
                    GetValues.Add(New AFValue(.ForecastItem.LowerControlLimit,
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is LimitLo Then
                    GetValues.Add(New AFValue(.ForecastItem.ForecastValue-
                      .ForecastItem.Range(pValueLoHi),
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is LimitHi Then
                    GetValues.Add(New AFValue(.ForecastItem.ForecastValue+
                      .ForecastItem.Range(pValueLoHi),
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is LimitHiHi Then
                    GetValues.Add(New AFValue(.ForecastItem.UpperControlLimit,
                      timeContext.StartTime.LocalTime))
                  ElseIf Attribute.Trait Is LimitMaximum Then
                    GetValues.Add(New AFValue(.ForecastItem.ForecastValue+
                      .ForecastItem.Range(pValueMinMax),
                      timeContext.StartTime.LocalTime))
                  Else
                    GetValues.Add(New AFValue(.Factor,
                      timeContext.StartTime.LocalTime))
                  End If

                End With

                timeContext.StartTime = New AFTime(
                  timeContext.StartTime.UtcSeconds+IntervalSeconds) ' Next intv.

              Loop

              Exit For

            Finally
              Monitor.Exit(DBM) ' Ensure that the lock is released.
            End Try
          End If

        Next DBM

      End If

      Do While timeContext.EndTime > timeContext.StartTime ' Missing results
        GetValues.Add(AFValue.CreateSystemStateValue(AFSystemStateCode.NoSample,
          timeContext.StartTime.LocalTime)) ' Return NoSample
        timeContext.StartTime = New AFTime(
          timeContext.StartTime.UtcSeconds+IntervalSeconds) ' Next interval
      Loop

      ' Returns the collection of values for the attribute sorted in increasing
      ' time order.
      Return GetValues

    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange,
      boundaryType As AFBoundaryType, filterExpression As String,
      includeFilteredValues As Boolean, inputAttributes As AFAttributeList,
      inputValues As AFValues(), inputTimes As List(Of AFTime),
      Optional maxCount As Integer = 0) As AFValues

      ' Returns a list of compressed values for the requested time range from
      ' the source provider.
      ' timeRange
      '   The bounding time range for the recorded values request. If the
      '   StartTime is earlier than the EndTime, the resulting values will be in
      '   time-ascending order, otherwise they will be in time-descending order.
      ' maxCount
      '   The maximum number of values to be returned. If zero, then all of the
      '   events within the requested time range will be returned.
      ' Returned times are affected by the specified boundaryType. If no values
      ' are found for the time range and conditions specified then the method
      ' will return success and an empty AFValues collection.
      ' When specifying True for the includeFilteredValues parameter,
      ' consecutive filtered events are not returned. The first value that would
      ' be filtered out is returned with its time and the enumeration value
      ' "Filtered". The next value in the collection will be the next compressed
      ' value in the specified direction that passes the filter criteria - if
      ' any.
      ' When both boundaryType and a filterExpression are specified, the events
      ' returned for the boundary condition specified are passed through the
      ' filter. If the includeFilteredValues parameter is True, the boundary
      ' values will be reported at the proper timestamps with the enumeration
      ' value "Filtered" when the filter conditions are not met at the boundary
      ' time. If the includeFilteredValues parameter is False for this case, no
      ' event is returned for the boundary time.

      ' Returns an AFValues collection with the recorded values.
      Return GetValues(Nothing, timeRange, maxCount, Nothing, Nothing)

    End Function


    Public Overrides Function PlotValues(timeRange As AFTimeRange,
      intervals As Integer, inputAttributes As AFAttributeList,
      inputValues As AFValues(), inputTimes As List(Of AFTime)) As AFValues

      ' Returns a single AFValue whose value is interpolated at the passed time.
      ' timeRange
      '   The bounding time range for the plot values request.
      ' intervals
      '   The number of intervals to plot over. Typically, this would be the
      '   number of horizontal pixels in the trend.
      ' To retrieve recorded values for multiple attributes, use the
      ' AFListData.InterpolatedValue to achieve better performance.

      ' Returns the AFValue for the attribute. If AFValue.IsGood returned by the
      ' data reference is True, then the returned value is converted to the
      ' proper UOM and Type. The timestamp of the value will be at the requested
      ' time.
      Return GetValues(Nothing, timeRange, intervals, Nothing, Nothing)

    End Function


  End Class


End Namespace
