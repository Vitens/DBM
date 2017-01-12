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

Imports Vitens.DynamicBandwidthMonitor

Namespace Vitens.DynamicBandwidthMonitor.RealTime

    Public Class DBMRtPIPoint

        Private InputPointDriver,OutputPointDriver As DBMPointDriver
        Private CorrelationPoints As New Collections.Generic.List(Of DBMCorrelationPoint)

        Public Sub New(InputPIPoint As PISDK.PIPoint,OutputPIPoint As PISDK.PIPoint)
            Dim ExDesc,SubstringsA(),SubstringsB() As String
            InputPointDriver=New DBMPointDriver(InputPIPoint)
            OutputPointDriver=New DBMPointDriver(OutputPIPoint)
            ExDesc=DirectCast(OutputPointDriver.Point,PISDK.PIPoint).PointAttributes("ExDesc").Value.ToString
            If Text.RegularExpressions.Regex.IsMatch(ExDesc,"^[-]?[\w\.-]+:[^:\?\*&]+(&[-]?[\w\.-]+:[^:\?\*&]+)*$") Then ' ExDesc attribute should contain correlation PI point(s)
                SubstringsA=ExDesc.Split(New Char(){"&"c}) ' Split multiple correlation PI points by &
                For Each SubstringA In SubstringsA
                    SubstringsB=SubstringA.Split(New Char(){":"c}) ' Format: [-]PI server:PI point
                    Try
                        If Not DBMRtCalculator.PISDK.Servers(SubstringsB(0).Substring(If(SubstringsB(0).Substring(0,1).Equals("-"),1,0))).PIPoints(SubstringsB(1)).Name.Equals(String.Empty) Then
                            CorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver(DBMRtCalculator.PISDK.Servers(SubstringsB(0).Substring(If(SubstringsB(0).Substring(0,1).Equals("-"),1,0))).PIPoints(SubstringsB(1))),SubstringsB(0).Substring(0,1).Equals("-"))) ' Add to correlation points
                        End If
                    Catch
                    End Try
                Next
            End If
        End Sub

        Public Sub Calculate
            Dim InputTimestamp,OutputTimestamp As PITimeServer.PITime
            InputTimestamp=DirectCast(InputPointDriver.Point,PISDK.PIPoint).Data.Snapshot.TimeStamp ' Timestamp of input point
            For Each CorrelationPoint In CorrelationPoints ' Check timestamp of correlation points
                InputTimestamp.UTCSeconds=Math.Min(InputTimestamp.UTCSeconds,DirectCast(CorrelationPoint.PointDriver.Point,PISDK.PIPoint).Data.Snapshot.TimeStamp.UTCSeconds) ' Timestamp of correlation point, keep earliest
            Next
            InputTimestamp.UTCSeconds-=DBMParameters.CalculationInterval+InputTimestamp.UTCSeconds Mod DBMParameters.CalculationInterval ' Can calculate output until (inclusive)
            OutputTimestamp=DirectCast(OutputPointDriver.Point,PISDK.PIPoint).Data.Snapshot.TimeStamp ' Timestamp of output point
            OutputTimestamp.UTCSeconds+=DBMParameters.CalculationInterval-OutputTimestamp.UTCSeconds Mod DBMParameters.CalculationInterval ' Next calculation timestamp
            If InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds Then ' If calculation timestamp can be calculated
                DirectCast(OutputPointDriver.Point,PISDK.PIPoint).Data.UpdateValue(DBMRtCalculator.DBM.Result(InputPointDriver,CorrelationPoints,InputTimestamp.LocalDate).Factor,InputTimestamp.LocalDate) ' Write calculated factor to output point
            End If
        End Sub

    End Class

End Namespace
