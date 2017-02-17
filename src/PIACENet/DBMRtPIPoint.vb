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


Imports PISDK
Imports PITimeServer
Imports System.Collections.Generic
Imports System.Math
Imports System.Text.RegularExpressions
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMRtPIPoint


    Private InputPointDriver, OutputPointDriver As DBMPointDriver
    Private CorrelationPoints As New List(Of DBMCorrelationPoint)


    Public Sub New(InputPIPoint As PIPoint, OutputPIPoint As PIPoint)

      Dim ExDesc, SubstringsA(), SubstringsB(), Server, Point As String
      Dim SubtractSelf As Boolean

      ' Set input and output DBMPointDriver objects.
      InputPointDriver = New DBMPointDriver(InputPIPoint)
      OutputPointDriver = New DBMPointDriver(OutputPIPoint)

      ' ExDesc attribute should contain correlation PI point(s)
      ExDesc = DirectCast(OutputPointDriver.Point, PIPoint). _
        PointAttributes("ExDesc").Value.ToString
      If Regex.IsMatch(ExDesc, _
        "^[-]?[\w\.-]+:[^:\?\*&]+(&[-]?[\w\.-]+:[^:\?\*&]+)*$") Then
        ' Split multiple correlation PI points by &
        SubstringsA = ExDesc.Split(New Char(){"&"c})
        For Each SubstringA In SubstringsA
          ' Format: [-]PI server:PI point
          ' If PI server is preceded by a '-', then set SubtractSelf to true,
          ' meaning that the input tag has to be subtracted from the
          ' correlation tag. For example when the correlation tag contains
          ' the input tag instead of it being adjacent to the input tag.
          SubstringsB = SubstringA.Split(New Char(){":"c})
          SubtractSelf = SubstringsB(0).Substring(0, 1).Equals("-")
          Server = SubstringsB(0).Substring(If(SubtractSelf, 1, 0))
          Point = SubstringsB(1)
          Try
            If Not DBMRtCalculator.PISDK.Servers(Server). _
              PIPoints(Point).Name.Equals(String.Empty) Then ' Check input
              CorrelationPoints.Add(New DBMCorrelationPoint _
                (New DBMPointDriver(DBMRtCalculator.PISDK.Servers(Server). _
                PIPoints(Point)), SubtractSelf)) ' Add to correlation points
            End If
          Catch
          End Try
        Next
      End If

    End Sub


    Public Sub Calculate

      ' This method is called for each output PI tag on each PI server. It will
      ' first determine the latest possible timestamp that can be calculated and
      ' then performs the DBM calculation and stores the resulting factor in the
      ' output PI tag.

      Dim InputTimestamp, OutputTimestamp As PITime

      ' Timestamp of latest value in input point
      InputTimestamp = DirectCast _
        (InputPointDriver.Point, PIPoint).Data.Snapshot.TimeStamp

      ' Check timestamp of latest value in correlation points
      For Each CorrelationPoint In CorrelationPoints
        ' Timestamp of correlation point, keep earliest
        InputTimestamp.UTCSeconds = Min(InputTimestamp.UTCSeconds, _
          DirectCast(CorrelationPoint.PointDriver.Point, PIPoint). _
          Data.Snapshot.TimeStamp.UTCSeconds)
      Next

      ' Can calculate output until (inclusive); aligned on interval.
      InputTimestamp.UTCSeconds -= _
        CalculationInterval+InputTimestamp.UTCSeconds Mod CalculationInterval

      ' Timestamp of output point
      OutputTimestamp = _
        DirectCast(OutputPointDriver.Point, PIPoint).Data.Snapshot.TimeStamp

      ' Next calculation timestamp; aligned on interval.
      OutputTimestamp.UTCSeconds += _
        CalculationInterval-OutputTimestamp.UTCSeconds Mod CalculationInterval

      ' If calculation timestamp can be calculated
      If InputTimestamp.UTCSeconds >= OutputTimestamp.UTCSeconds Then
        ' Perform calculations and write resulting factor to output point
        DirectCast(OutputPointDriver.Point, PIPoint).Data.UpdateValue _
          (DBMRtCalculator.DBM.Result(InputPointDriver, CorrelationPoints, _
          InputTimestamp.LocalDate).Factor, InputTimestamp.LocalDate)
      End If

    End Sub


  End Class


End Namespace
