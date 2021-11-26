Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Asset.AFAttributeTrait
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMInfo
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics
Imports Vitens.DynamicBandwidthMonitor.DBMStrings
Imports Vitens.DynamicBandwidthMonitor.DBMTimeSeries


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
    '   Target    Measured value, forecast if not available
    '             Indicates the aimed-for measurement value or process output.
    '   Forecast  Forecast value
    '   Minimum   Lower control limit (p = 0.9999)
    '             Indicates the very lowest possible measurement value or
    '             process output.
    '   LoLo      Lower control limit (default)
    '             Indicates a very low measurement value or process output,
    '             typically an abnormal one that initiates an alarm.
    '   Lo        Lower control limit (p = 0.95)
    '             Indicates a low measurement value or process output, typically
    '             one that initiates a warning.
    '   Hi        Upper control limit (p = 0.95)
    '             Indicates a high measurement value or process output,
    '             typically one that initiates a warning.
    '   HiHi      Upper control limit (default)
    '             Indicates a very high measurement value or process output,
    '             typically an abnormal one that initiates an alarm.
    '   Maximum   Upper control limit (p = 0.9999)
    '             Indicates the very highest possible measurement value or
    '             process output.


    Const CategoryNoCorrelation As String = "NoCorrelation"
    Const pValueLoHi As Double = 0.95 ' Confidence interval for Lo and Hi
    Const pValueMinMax As Double = 0.9999 ' CI for Minimum and Maximum


    Private Annotations As New Dictionary(Of AFTime, Object)
    Private Shared DBM As New DBM(New DBMLoggerAFTrace)


    Public Shared Function CreateDataPipe As Object

      ' This Data Reference is a System of Record, so expose the CreateDataPipe
      ' method and return a new instance of the event source class derived from
      ' AFEventSource.

      Return New DBMEventSource

    End Function


    Public Overrides Readonly Property SupportedContexts _
      As AFDataReferenceContext

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/P_OSIsoft_A
      ' F_Asset_AFDataReference_SupportedContexts.htm
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


    Private Function SupportsFutureData As Boolean

      ' Returns a boolean indicating if this attribute supports future data,
      ' f.ex. when it returns a forecast. We should not enable the data pipe for
      ' these attributes as services like the Analyses Service will not be able
      ' to use future data because the snapshot timestamp is seen as the last
      ' available timestamp.

      Return Attribute.Trait Is LimitTarget Or Attribute.Trait Is Forecast Or
        Attribute.Trait Is LimitMinimum Or Attribute.Trait Is LimitLoLo Or
        Attribute.Trait Is LimitLo Or Attribute.Trait Is LimitHi Or
        Attribute.Trait Is LimitHiHi Or Attribute.Trait Is LimitMaximum

    End Function


    Public Overrides Readonly Property SupportedDataMethods As AFDataMethods

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/P_OSIsoft_A
      ' F_Asset_AFDataReference_SupportedDataMethods.htm
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
        SupportedDataMethods = AFDataMethods.InterpolatedValue Or
          AFDataMethods.InterpolatedValues Or
          AFDataMethods.InterpolatedValuesAtTimes Or
          AFDataMethods.RecordedValue Or
          AFDataMethods.RecordedValues Or
          AFDataMethods.RecordedValuesByCount Or
          AFDataMethods.PlotValues Or
          AFDataMethods.Annotations Or
          AFDataMethods.RecordedValuesAtTimes Or
          AFDataMethods.Summary Or
          AFDataMethods.Summaries
        If SupportsFutureData Then
          ' Support future data if available.
          SupportedDataMethods = SupportedDataMethods Or AFDataMethods.Future
        End If
        If Not SupportsFutureData Or Attribute.Trait Is LimitTarget Then
          ' Support data pipe for non-future data, as well as for Target.
          SupportedDataMethods = SupportedDataMethods Or AFDataMethods.DataPipe
        End If
        Return SupportedDataMethods
      End Get

    End Property


    Public Overrides Readonly Property SupportedMethods As AFDataReferenceMethod

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/P_OSIsoft_A
      ' F_Asset_AFDataReference_SupportedMethods.htm
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


    Public Overrides Property ConfigString As String

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/P_OSIsoft_A
      ' F_Asset_AFDataReference_ConfigString.htm
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


    Private Function ConfigurationIsValid As Boolean

      ' Check if this attribute is properly configured. The attribute and it's
      ' parent (input source) need to be an instance of an object, and the data
      ' type for both needs to be a double.

      Return Attribute IsNot Nothing AndAlso
        Attribute.Type Is GetType(Double) AndAlso
        Attribute.Parent IsNot Nothing AndAlso
        Attribute.Parent.Type Is GetType(Double)

    End Function


    Public Overrides Readonly Property [Step] As Boolean

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/P_OSIsoft_A
      ' F_Asset_AFDataReference_Step.htm
      ' This property returns True if the value returned for this
      ' AFDataReference is stepped.
      ' The step attribute defines how the values are to be interpolated. When
      ' set to False, the values are treated as a continuous signal, and
      ' adjacent values are linearly interpolated. When set to True, the values
      ' are treated discretely and adjacent values are not interpolated.

      Get
        ' Returns True if the value for this AFDataReference is stepped. The
        ' value is inherited from the parent attribute when the configuration
        ' is valid, or set to False when the configuration is invalid.
        Return ConfigurationIsValid AndAlso Attribute.Parent.Step
      End Get

    End Property


    Private Sub Annotate(Value As AFValue, Annotation As Object)

      ' Stores the annotation object in a dictionary for the timestamp of the
      ' AFValue.

      If Value IsNot Nothing Then ' Key

        Annotations.Remove(Value.Timestamp) ' Remove existing
        Value.Annotated = False

        If Annotation IsNot Nothing Then ' Value

          DBM.Logger.LogDebug(
            "Timestamp " & Value.Timestamp.LocalTime.ToString("s") & "; " &
            "Annotation " & DirectCast(Annotation, String), Attribute.GetPath)

          Annotations.Add(Value.Timestamp, Annotation) ' Add
          Value.Annotated = True

        End If

      End If

    End Sub


    Public Overrides Function GetAnnotation(value As AFValue) As Object

      ' Gets the annotation associated with a single historical event.

      ' OSIsoft Tech Support: "[...] the reason you can't see it in Processbook,
      '   is because of an old known issue related to annotations with AF2
      '   attributes"

      GetAnnotation = Nothing
      If value IsNot Nothing AndAlso
        Annotations.TryGetValue(value.Timestamp, GetAnnotation) Then
        Annotations.Remove(value.Timestamp) ' Remove after get
        Return GetAnnotation
      Else
        Return String.Empty ' Default
      End If

    End Function


    Public Overrides Function GetValue(context As Object,
      timeContext As Object, inputAttributes As AFAttributeList,
      inputValues As AFValues) As AFValue

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/M_OSIsoft_A
      ' F_Asset_AFDataReference_GetValue_1.htm
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

      Dim SourceTimestamp As DateTime
      Dim Timestamp As DateTime = Now

      ' Check if this attribute is properly configured. If it is not configured
      ' properly, return a Configure system state. This will be done in the
      ' GetValues method.
      If ConfigurationIsValid Then

        ' Retrieve current calculation timestamp.
        If Attribute.Trait Is LimitTarget Then
          SourceTimestamp =
            New DBMPointDriver(Attribute.Parent).SnapshotTimestamp
        Else
          SourceTimestamp =
            New DBMPointDriver(Attribute.Parent).CalculationTimestamp
        End If

        If timeContext Is Nothing Then

          ' No passed timestamp, use current calculation timestamp.
          Timestamp = SourceTimestamp

        Else

          ' Use passed timestamp.
          Timestamp = DirectCast(timeContext, AFTime).LocalTime

          ' For attributes without future data, as well as for Target, return
          ' the snapshot value for timestamps up to 10 minutes into the future.
          ' Attributes not supporting future data will not return any values
          ' beyond 10 minutes past the snapshot timestamp, but return a No Data
          ' system state instead.
          If (Not SupportsFutureData Or Attribute.Trait Is LimitTarget) And
            Timestamp > SourceTimestamp And
            Timestamp < SourceTimestamp.AddMinutes(10) Then
            Timestamp = SourceTimestamp
          End If

        End If

      End If

      ' Returns the single value for the attribute.
      DBM.Logger.LogDebug(
        "Timestamp " & Timestamp.ToString("s"), Attribute.GetPath)
      GetValue = GetValues(Nothing, New AFTimeRange(New AFTime(Timestamp),
        New AFTime(Timestamp.AddSeconds(CalculationInterval))), 1,
        Nothing, Nothing)(0) ' Request a single value
      DBM.Logger.LogTrace("Return value", Attribute.GetPath)
      Return GetValue

    End Function


    Private Function Deflatline(Values As AFValues,
      Results As List(Of DBMResult), Stepped As Boolean) As AFValues

      ' De-flatline detects and removes incorrect flatlines in the data and
      ' replaces them with forecast values, which are then weight adjusted to
      ' the original time range total. Data quality is marked as Substituted for
      ' this data. There are two situations we want to automatically improve
      ' here: whenever data is manually replaced by inserting an average value
      ' over a longer time range, and whenever we receive a flatline containing
      ' all zeroes, followed by a single spike value containing the total of the
      ' previous flatline time range.

      Dim Value As AFValue
      Dim iV, iFL, i As Integer ' Iterators
      Dim MeasurementWeight, ForecastWeight As Double

      Deflatline = New AFValues

      For Each Value In Values

        Deflatline.Add(Value) ' Add original value.

        ' The flatline ends when a value is different from the value that is
        ' currently being monitored. This value is also replaced by the weight
        ' adjusted forecast, as it may contain a spike value which needs to be
        ' redistributed.
        If iV > iFL AndAlso Not Convert.ToDouble(Value.Value) =
          Convert.ToDouble(Values.Item(iFL).Value) Then

          ' The following timestamps are important:
          '  * iFL-1 Last good value before flatline
          '  * iFL   Start of flatline
          '  * iV    End of flatline, spike value
          '  * iV+1  First good value

          ' If all of the following conditions are true, we can redistribute the
          ' total:
          '  1) A flatline can never start at the first value in the time
          '       series, as we don't know if the start was already before this.
          '  2) It also cannot end on the last value in the time series, as we
          '       don't know if the end was after this. We need at least one
          '       extra value after the spike value to determine it's length.
          '  3) We need at least 3 values: two repeating values and one spike
          '       value.
          '  4) The duration between the last and first good values around the
          '       flatline has to be at least 12 calculation intervals. With the
          '       default calculation interval of 5 minutes, this equals one
          '       hour.
          If (iFL > 0 And iV < Values.Count-1 And iV-iFL > 1) AndAlso
            Values.Item(iV+1).Timestamp.LocalTime.Subtract(
            Values.Item(iFL-1).Timestamp.LocalTime).TotalSeconds/
            CalculationInterval >= 12 Then

            DBM.Logger.LogDebug("Found flatline from " &
              Values.Item(iFL).Timestamp.LocalTime.ToString("s") & " to " &
              Values.Item(iV).Timestamp.LocalTime.ToString("s"),
              Attribute.GetPath)

            ' Determining the scaling factor for weight adjustment:
            '  * For stepped values:
            '      z = a(u-t)+qb(v-u)+qc(w-v)+qd(x-w)+qe(y-x)
            '      Solve for q:
            '        q = (z-a(u-t))/(b(v-u)+c(w-v)+d(x-w)+e(y-x))
            '  * For non-stepped values:
            '      z = (a+qb)/2(u-t)+(qb+qc)/2(v-u)+(qc+qd)/2(w-v)+
            '          (qd+qe)/2(x-w)+(qe+f)/2(y-x)
            '      Solve for q:
            '        q = (z-a(u-t)/2-f(y-x)/2)/
            '            (b(v-t)/2+c(w-u)/2+d(x-v)/2+e(y-w)/2)
            ' Where a-f = values (iFL-1, iFL, first forecast,
            '               last forecast, iV, iV+1),
            '       t-y = their timestamps, q = scaling factor, z = weight.

            ' Phase 1: Calculate weight of original from the last good value to
            ' the first good value. These two values are unmodified in the end
            ' result.
            i = -1 ' Start at the last good value.
            MeasurementWeight = 0 ' Initial weight.
            Do While iFL+i < iV+1
              ' Add weight.
              MeasurementWeight += TimeWeightedValue(
                Convert.ToDouble(Values.Item(iFL+i).Value),
                Convert.ToDouble(Values.Item(iFL+i+1).Value),
                Values.Item(iFL+i).Timestamp.LocalTime,
                Values.Item(iFL+i+1).Timestamp.LocalTime, Stepped)
              i += 1 ' Increase iterator.
            Loop ' iFL-1 to iV.

            ' Phase 2: Calculate weight of forecast.
            i = 0
            ForecastWeight = 0 ' Initial weight.
            Do While i < Results.Count-1 ' Exclude last value, we need i+1.
              If Results.Item(i).Timestamp >
                Values.Item(iFL-1).Timestamp.LocalTime Then
                ' All forecasts after the last good value and before the first
                ' good value.
                If Results.Item(i-1).Timestamp <=
                  Values.Item(iFL-1).Timestamp.LocalTime Then
                  If Stepped Then
                    ' Last good value to first forecast for stepped values.
                    MeasurementWeight -= TimeWeightedValue(
                      Convert.ToDouble(Values.Item(iFL-1).Value), Nothing,
                      Values.Item(iFL-1).Timestamp.LocalTime,
                      Results.Item(i).Timestamp) ' -a(u-t)
                    ForecastWeight += TimeWeightedValue(
                      Results.Item(i).ForecastItem.Forecast, Nothing,
                      Results.Item(i).Timestamp,
                      Results.Item(i+1).Timestamp) ' b(v-u)
                  Else
                    ' Last good value to first forecast for non-stepped values.
                    MeasurementWeight -= TimeWeightedValue(
                      Convert.ToDouble(Values.Item(iFL-1).Value), Nothing,
                      Values.Item(iFL-1).Timestamp.LocalTime,
                      Results.Item(i).Timestamp)/2 ' -a(u-t)/2
                    ForecastWeight += TimeWeightedValue(
                      Results.Item(i).ForecastItem.Forecast, Nothing,
                      Values.Item(iFL-1).Timestamp.LocalTime,
                      Results.Item(i+1).Timestamp)/2 ' b(v-t)/2
                  End If ' Stepped.
                ElseIf Results.Item(i+1).Timestamp >=
                  Values.Item(iV+1).Timestamp.LocalTime Then
                  If Stepped Then
                    ' Last forecast to first good value for stepped values.
                    ForecastWeight += TimeWeightedValue(
                      Results.Item(i).ForecastItem.Forecast, Nothing,
                      Results.Item(i).Timestamp,
                      Values.Item(iV+1).Timestamp.LocalTime) ' e(y-x)
                  Else
                    ' Last forecast to first good value for non-stepped values.
                    MeasurementWeight -= TimeWeightedValue(
                      Convert.ToDouble(Values.Item(iV+1).Value), Nothing,
                      Results.Item(i).Timestamp,
                      Values.Item(iV+1).Timestamp.LocalTime)/2 ' -f(y-x)/2
                    ForecastWeight += TimeWeightedValue(
                      Results.Item(i).ForecastItem.Forecast, Nothing,
                      Results.Item(i-1).Timestamp,
                      Values.Item(iV+1).Timestamp.LocalTime)/2 ' e(y-w)/2
                  End If ' Stepped.
                  Exit Do ' Done.
                Else
                  ' All forecasts after the last good value and before the first
                  ' good value, except first and last.
                  If Stepped Then
                    ForecastWeight += TimeWeightedValue(
                      Results.Item(i).ForecastItem.Forecast, Nothing,
                      Results.Item(i).Timestamp,
                      Results.Item(i+1).Timestamp) ' c(w-v), ...
                  Else
                    ' If not stepped, each value has a certain weight in it's
                    ' previous and it's next interval. For both of these
                    ' intervals it accounts for half of the total weight, since
                    ' a non-stepped interval weight is based on the average
                    ' value between the previous and the next value, with the
                    ' other half of the weight coming from it's previous and
                    ' next values. So we can just add half the total weight from
                    ' the previous interval to the next.
                    ForecastWeight += TimeWeightedValue(
                      Results.Item(i).ForecastItem.Forecast, Nothing,
                      Results.Item(i-1).Timestamp,
                      Results.Item(i+1).Timestamp)/2 ' c(w-u)/2, ...
                  End If ' Stepped.
                End If
              End If
              i += 1 ' Increase iterator.
            Loop ' Results.

            ' Log weights
            DBM.Logger.LogDebug(
              "MeasurementWeight " & MeasurementWeight.ToString & "; " &
              "ForecastWeight " & ForecastWeight.ToString, Attribute.GetPath)

            ' Phase 3: Remove all values after the last good value.
            DBM.Logger.LogDebug(
              "Remove " & (iV-iFL+1).ToString & " raw values",
              Attribute.GetPath)
            Do While Deflatline.Item(Deflatline.Count-1).
              Timestamp.LocalTime > Values.Item(iFL-1).Timestamp.LocalTime
              Deflatline.RemoveAt(Deflatline.Count-1)
            Loop

            ' Phase 4: Add weight adjusted forecasts.
            i = 0
            Do While i < Results.Count-1
              If Results.Item(i).Timestamp >
                Values.Item(iFL-1).Timestamp.LocalTime And
                Results.Item(i).TimestampIsValid Then
                ' Add weight adjusted forecasts for valid timestamps after the
                ' last good value and before the first good value.
                Deflatline.Add(New AFValue(
                  Results.Item(i).ForecastItem.Forecast*
                  MeasurementWeight/ForecastWeight,
                  New AFTime(Results.Item(i).Timestamp)))
                Deflatline.Item(Deflatline.Count-1).Substituted = True
                If Results.Item(i-1).Timestamp <=
                  Values.Item(iFL-1).Timestamp.LocalTime Then
                  ' Annotate the first weight adjusted forecast with the factor
                  ' value.
                  Annotate(Deflatline.Item(Deflatline.Count-1),
                    String.Format(sForecastFactorAnnotation,
                    MeasurementWeight/ForecastWeight))
                End If
              End If
              i += 1 ' Increase iterator.
              If Results.Item(i).Timestamp >=
                Values.Item(iV+1).Timestamp.LocalTime Then Exit Do ' Done.
            Loop ' Results.

            ' After deflatlining, set the flatline start index to the value
            ' after the first good value. This ensures that inserted weight
            ' adjusted forecast data will not be overwritten.
            iFL = iV+2

          Else

            ' Values differ, but no flatline.
            iFL = iV ' Set flatline start index to the current value.

          End If ' Check flatline.

        End If ' Check values differ.

        iV += 1 ' Move iterator to next value.

      Next Value

      Return Deflatline

    End Function


    Public Overrides Function GetValues(context As Object,
      timeRange As AFTimeRange, numberOfValues As Integer,
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      ' Returns values for each interval in a time range. The (aligned) end time
      ' itself is excluded. Run the calculation for a maximum duration of one
      ' CalculationInterval, and return ScanTimeout digital state values for
      ' each next value in the time range after exceeding this duration.

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/M_OSIsoft_A
      ' F_Asset_AFDataReference_GetValues.htm
      ' This method gets a collection of AFValue objects for an attribute based
      ' upon the data reference configuration within the specified AFTimeRange
      ' context.
      ' timeRange
      '   The time-range to be used when getting the attribute's collection of
      '   values.
      ' numberOfValues
      '   The number of values desired. The behavior of GetValues varies based
      '   on the value of this parameter:
      '   If 0, then all recorded values within the timeRange will be returned
      '   with an interpolated value at the start and end time, if possible.
      '   If the number of values requested is less than zero, then if
      '   supported, the Data Reference will return evenly spaced interpolated
      '   values across the timeRange, with a value returned at both end
      '   points of the time range. For example, specifying -25 over a 24 hour
      '   period will produce an hourly value.
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

      Dim Element, ParentElement, SiblingElement As AFElement
      Dim InputPointDriver As DBMPointDriver
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim Results As List(Of DBMResult)
      Dim RawSnapshot As DateTime
      Dim RawValues As AFValues = Nothing
      Dim Result As DBMResult
      Dim iR, iD As Integer ' Iterators for raw values and DBM results.

      GetValues = New AFValues

      ' Check if this attribute is properly configured.
      If Not ConfigurationIsValid Then
        ' Attribute or parent attribute is not configured properly, return a
        ' Configure system state. Definition: 'The point configuration has been
        ' rejected as invalid by the data source.'
        DBM.Logger.LogWarning(
          "Invalid configuration, return Configure", Attribute.GetPath)
        GetValues.Add(AFValue.CreateSystemStateValue(
          AFSystemStateCode.Configure, timeRange.StartTime))
        Return GetValues
      End If

      Element = DirectCast(Attribute.Element, AFElement)
      InputPointDriver = New DBMPointDriver(Attribute.Parent) ' Parent attribute

      DBM.Logger.LogDebug(
        "StartTime " & timeRange.StartTime.LocalTime.ToString("s") & "; " &
        "EndTime " & timeRange.EndTime.LocalTime.ToString("s") & "; " &
        "numberOfValues " & numberOfValues.ToString, Attribute.GetPath)

      ' Retrieve correlation points from AF hierarchy for first-level child
      ' attributes in non-root elements only when calculating the DBM factor
      ' value and if correlation calculations are not disabled using
      ' categories.
      If Not Element.IsRoot And Attribute.Parent.Parent Is Nothing And
        Attribute.Trait Is Nothing And
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
          Next SiblingElement
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

      ' Get DBM results for time range.
      Results = DBM.GetResults(InputPointDriver, CorrelationPoints,
        timeRange.StartTime.LocalTime, timeRange.EndTime.LocalTime,
        numberOfValues)

      ' Retrieve raw snapshot timestamp and raw values for Target trait. Only
      ' retrieve values if the start timestamp for this time range is before the
      ' next interval after the snapshot timestamp.
      If Attribute.Trait Is LimitTarget Then
        RawSnapshot = InputPointDriver.SnapshotTimestamp
        If timeRange.StartTime.LocalTime < NextInterval(RawSnapshot) Then
          RawValues = Attribute.Parent.
            GetValues(timeRange, numberOfValues, Nothing)
          DBM.Logger.LogTrace(
            "Retrieved " & RawValues.Count.ToString & " raw values",
            Attribute.GetPath)
          ' If there are no DBM results to iterate over, and there are raw
          ' values for this time range, return the raw values directly.
          If Results.Count = 0 And RawValues.Count > 0 Then
            DBM.Logger.LogTrace(
              "Return " & RawValues.Count.ToString & " raw values",
              Attribute.GetPath)
            Return RawValues
          End If
        Else
          RawValues = New AFValues ' Future data, no raw values.
        End If
      End If

      ' Iterate over DBM results for time range.
      For Each Result In Results

        With Result

          If .TimestampIsValid And
            (SupportsFutureData Or Not .IsFutureData) Then

            If Attribute.Trait Is LimitTarget Then

              ' The Target trait returns the original, valid raw measurements,
              ' augmented with forecast data for invalid and future data. Data
              ' quality is marked as Substituted for invalid data replaced by
              ' forecast values and for future forecast values, or as
              ' Questionable for data exceeding Minimum and Maximum control
              ' limits.

              ' Augment raw values with forecast. This is done if any of four
              ' conditions is true:
              '  1) There are no raw values for the time period. Since there are
              '       no values, the best we can do is return the forecast.
              '  2) For the first raw value, if this value is not good. The
              '       forecast is returned because either this value is before
              '       this result, or there are no values before this value.
              '  3) For all but the first raw value, if the previous raw value
              '       is not good. While this value is not good, the forecast is
              '       returned. Note that the previous raw value can be on the
              '       exact same timestamp as this result.
              '  4) If the timestamp is past the raw snapshot timestamp. This
              '       appends forecast values to the future.
              If RawValues.Count = 0 OrElse
                Not RawValues.Item(Max(0, iR-1)).IsGood OrElse
                .Timestamp > RawSnapshot Then
                If IsNaN(.ForecastItem.Forecast) Then
                  ' If there is no valid forecast result, return an InvalidData
                  ' state. Definition: 'Invalid Data state.'
                  GetValues.Add(AFValue.CreateSystemStateValue(
                    AFSystemStateCode.InvalidData, New AFTime(.Timestamp)))
                Else
                  ' Replace bad values with forecast, append forecast values to
                  ' the future. Mark as substituted.
                  GetValues.Add(New AFValue(
                    .ForecastItem.Forecast, New AFTime(.Timestamp)))
                  GetValues.Item(GetValues.Count-1).Substituted = True
                End If
              End If

              ' Include valid raw values. Raw values are appended while there
              ' are still values available before or on the raw snapshot
              ' timestamp, and if any of two conditions is true:
              '  1) If there is a next DBM result, and the raw value timestamp
              '       is on or before this result. This includes all values
              '       until the timestamp of the next DBM result (inclusive).
              '  2) The raw value timestamp is on or after the last DBM result.
              '       This includes all remaining values after the last result.
              Do While iR < RawValues.Count AndAlso
                RawValues.Item(iR).Timestamp.LocalTime <= RawSnapshot AndAlso
                ((iD+1 < Results.Count AndAlso RawValues.Item(iR).
                Timestamp.LocalTime <= Results.Item(iD+1).Timestamp) OrElse
                RawValues.Item(iR).Timestamp.LocalTime >=
                Results.Item(Results.Count-1).Timestamp)
                If RawValues.Item(iR).IsGood Then ' Only include good values
                  ' Create a copy of the value, so that it's attribute is not
                  ' linked to an attribute that might not support annotations.
                  GetValues.Add(New AFValue(
                    RawValues.Item(iR).Value, RawValues.Item(iR).Timestamp))
                  ' Mark events (exceeding Minimum and Maximum control limits)
                  ' as questionable.
                  GetValues.Item(GetValues.Count-1).Questionable =
                    Abs(.ForecastItem.Measurement-.ForecastItem.Forecast) >
                    .ForecastItem.Range(pValueMinMax)
                End If
                iR += 1 ' Move iterator to next raw value.
              Loop

            ElseIf Attribute.Trait Is Forecast Then
              GetValues.Add(New AFValue(
                .ForecastItem.Forecast, New AFTime(.Timestamp)))

            ElseIf Attribute.Trait Is LimitMinimum Then
              GetValues.Add(New AFValue(
                .ForecastItem.Forecast-.ForecastItem.Range(pValueMinMax),
                New AFTime(.Timestamp)))

            ElseIf Attribute.Trait Is LimitLoLo Then
              GetValues.Add(New AFValue(
                .ForecastItem.LowerControlLimit, New AFTime(.Timestamp)))

            ElseIf Attribute.Trait Is LimitLo Then
              GetValues.Add(New AFValue(
                .ForecastItem.Forecast-.ForecastItem.Range(pValueLoHi),
                New AFTime(.Timestamp)))

            ElseIf Attribute.Trait Is LimitHi Then
              GetValues.Add(New AFValue(
                .ForecastItem.Forecast+.ForecastItem.Range(pValueLoHi),
                New AFTime(.Timestamp)))

            ElseIf Attribute.Trait Is LimitHiHi Then
              GetValues.Add(New AFValue(
                .ForecastItem.UpperControlLimit, New AFTime(.Timestamp)))

            ElseIf Attribute.Trait Is LimitMaximum Then
              GetValues.Add(New AFValue(
                .ForecastItem.Forecast+.ForecastItem.Range(pValueMinMax),
                New AFTime(.Timestamp)))

            Else ' Factor
              GetValues.Add(New AFValue(
                .Factor, New AFTime(.Timestamp)))

            End If

          End If

        End With

        iD += 1 ' Move iterator to next DBM result.

      Next Result

      ' If there are no calculation results, return a NoData state.
      ' Definition: 'Data-retrieval functions use this state for time periods
      ' where no archive values for a tag can exist 10 minutes into the future
      ' or before the oldest mounted archive.'
      If GetValues.Count = 0 Then
        DBM.Logger.LogTrace(
          "No values to return, return NoData", Attribute.GetPath)
        GetValues.Add(AFValue.CreateSystemStateValue(
          AFSystemStateCode.NoData, timeRange.StartTime))
        Return GetValues
      End If

      If GetValues.Count > 1 And
        Not numberOfValues = 1 And
        Attribute.Trait Is LimitTarget Then

        ' De-flatline
        GetValues = Deflatline(GetValues, Results, [Step])

        ' If there is more than one value, annotate the first value in the
        ' Target values with model calibration metrics.
        Annotate(GetValues.Item(0), Statistics(Results).Brief)

      End If

      ' Returns the collection of values for the attribute sorted in increasing
      ' time order.
      DBM.Logger.LogTrace(
        "Return " & GetValues.Count.ToString & " values", Attribute.GetPath)
      Return GetValues

    End Function


    Public Overrides Function RecordedValues(timeRange As AFTimeRange,
      boundaryType As AFBoundaryType, filterExpression As String,
      includeFilteredValues As Boolean, inputAttributes As AFAttributeList,
      inputValues As AFValues(), inputTimes As List(Of AFTime),
      Optional maxCount As Integer = 0) As AFValues

      ' https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/M_OSIsoft_A
      ' F_Asset_AFDataReference_RecordedValues.htm
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


    Public Overrides Function InterpolatedValuesByCount(
      timeRange As AFTimeRange, numberOfValues As Integer,
      filterExpression As String, includeFilteredValues As Boolean,
      inputAttributes As AFAttributeList, inputValues As AFValues()) As AFValues

      ' https://docs.osisoft.com/bundle/af-sdk/page/html/M_OSIsoft_AF_Asset_AFDa
      ' taReference_InterpolatedValuesByCount.htm
      ' Retrieves interpolated values over the specified time range evenly
      ' spaced using the numberOfValues.

      Dim Interval As AFTimeSpan

      DBM.Logger.LogDebug(
        "StartTime " & timeRange.StartTime.LocalTime.ToString("s") & "; " &
        "EndTime " & timeRange.EndTime.LocalTime.ToString("s") & "; " &
        "numberOfValues " & numberOfValues.ToString, Attribute.GetPath)

      InterpolatedValuesByCount = New AFValues
      If timeRange.StartTime = timeRange.EndTime Or numberOfValues = 1 Then
        ' Return a single value.
        InterpolatedValuesByCount.Add(
          GetValue(Nothing, timeRange.StartTime, Nothing, Nothing))
      Else
        ' Return multiple values.
        Interval = New AFTimeSpan(0, 0, 0, 0, 0,
          (timeRange.EndTime.UtcSeconds-timeRange.StartTime.UtcSeconds)/
          (numberOfValues-1), 0)
        InterpolatedValuesByCount = Summaries(
          New AFTimeRange(timeRange.StartTime, timeRange.EndTime+Interval),
          Interval,
          AFSummaryTypes.Average,
          AFCalculationBasis.TimeWeighted,
          AFTimestampCalculation.EarliestTime)(AFSummaryTypes.Average)
      End If

      DBM.Logger.LogTrace(
        "Return " & InterpolatedValuesByCount.Count.ToString & " values",
        Attribute.GetPath)
      Return InterpolatedValuesByCount

    End Function


  End Class


End Namespace
