Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

Public Class DBMRtPIPoint

    Private InputDBMPointDriver,OutputDBMPointDriver As DBMPointDriver
    Private DBMCorrelationPoints As Collections.Generic.List(Of DBMCorrelationPoint)

    Public Sub New(InputPIPoint As PISDK.PIPoint,OutputPIPoint As PISDK.PIPoint)
        Dim ExDesc,FieldsA(),FieldsB() As String
        InputDBMPointDriver=New DBMPointDriver(InputPIPoint)
        OutputDBMPointDriver=New DBMPointDriver(OutputPIPoint)
        DBMCorrelationPoints=New Collections.Generic.List(Of DBMCorrelationPoint)
        ExDesc=OutputDBMPointDriver.Point.PointAttributes("ExDesc").Value.ToString
        If Text.RegularExpressions.Regex.IsMatch(ExDesc,"^[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}(&[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}){0,}$") Then
            FieldsA=Split(ExDesc.ToString,"&")
            For Each thisField As String In FieldsA
                FieldsB=Split(thisField,":")
                Try
                    If DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)).Name<>"" Then
                        DBMCorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver(DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1))),Left(FieldsB(0),1)="-"))
                    End If
                Catch
                End Try
            Next
        End If
    End Sub

    Public Sub Calculate
        Dim InputTimestamp,OutputTimestamp As PITimeServer.PITime
        Dim Value As Double
        InputTimestamp=InputDBMPointDriver.Point.Data.Snapshot.TimeStamp ' Timestamp of input point
        For Each thisDBMCorrelationPoint As DBMCorrelationPoint In DBMCorrelationPoints ' Check timestamp of correlation points
            InputTimestamp.UTCSeconds=Math.Min(InputTimestamp.UTCSeconds,thisDBMCorrelationPoint.DBMPointDriver.Point.Data.Snapshot.TimeStamp.UTCSeconds) ' Timestamp of correlation point, keep earliest
        Next
        InputTimestamp.UTCSeconds-=DBMParameters.CalculationInterval+InputTimestamp.UTCSeconds Mod DBMParameters.CalculationInterval ' Can calculate output until (inclusive)
        OutputTimestamp=OutputDBMPointDriver.Point.Data.Snapshot.TimeStamp ' Timestamp of output point
        OutputTimestamp.UTCSeconds+=DBMParameters.CalculationInterval-OutputTimestamp.UTCSeconds Mod DBMParameters.CalculationInterval ' Next calculation timestamp
        If InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds Then ' If calculation timestamp can be calculated
            Try
                Value=DBMRtCalculator.DBM.Calculate(InputDBMPointDriver,DBMCorrelationPoints,InputTimestamp.LocalDate).Factor
            Catch
                Value=Double.NaN ' Calculation error
            End Try
            Try
                OutputDBMPointDriver.Point.Data.UpdateValue(Value,InputTimestamp.LocalDate) ' Write value to output point
            Catch
            End Try
        End If
    End Sub

End Class
