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

Namespace DBMRt

    Public Class DBMRtPIPoint

        Private InputDBMPointDriver,OutputDBMPointDriver As DBM.DBMPointDriver
        Private DBMCorrelationPoints As Collections.Generic.List(Of DBM.DBMCorrelationPoint)

        Public Sub New(InputPIPoint As PISDK.PIPoint,OutputPIPoint As PISDK.PIPoint)
            Dim ExDesc,FieldsA(),FieldsB() As String
            InputDBMPointDriver=New DBM.DBMPointDriver(InputPIPoint)
            OutputDBMPointDriver=New DBM.DBMPointDriver(OutputPIPoint)
            DBMCorrelationPoints=New Collections.Generic.List(Of DBM.DBMCorrelationPoint)
            ExDesc=DirectCast(OutputDBMPointDriver.Point,PISDK.PIPoint).PointAttributes("ExDesc").Value.ToString
            If Text.RegularExpressions.Regex.IsMatch(ExDesc,"^[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}(&[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}){0,}$") Then
                FieldsA=ExDesc.ToString.Split(New Char(){"&"c})
                For Each thisField As String In FieldsA
                    FieldsB=thisField.Split(New Char(){":"c})
                    Try
                        If DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)).Name<>"" Then
                            DBMCorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1))),Left(FieldsB(0),1)="-"))
                        End If
                    Catch
                    End Try
                Next
            End If
        End Sub

        Public Sub Calculate
            Dim InputTimestamp,OutputTimestamp As PITimeServer.PITime
            Dim Value As Double
            InputTimestamp=DirectCast(InputDBMPointDriver.Point,PISDK.PIPoint).Data.Snapshot.TimeStamp ' Timestamp of input point
            For Each thisDBMCorrelationPoint As DBM.DBMCorrelationPoint In DBMCorrelationPoints ' Check timestamp of correlation points
                InputTimestamp.UTCSeconds=Math.Min(InputTimestamp.UTCSeconds,DirectCast(thisDBMCorrelationPoint.DBMPointDriver.Point,PISDK.PIPoint).Data.Snapshot.TimeStamp.UTCSeconds) ' Timestamp of correlation point, keep earliest
            Next
            InputTimestamp.UTCSeconds-=DBM.DBMParameters.CalculationInterval+InputTimestamp.UTCSeconds Mod DBM.DBMParameters.CalculationInterval ' Can calculate output until (inclusive)
            OutputTimestamp=DirectCast(OutputDBMPointDriver.Point,PISDK.PIPoint).Data.Snapshot.TimeStamp ' Timestamp of output point
            OutputTimestamp.UTCSeconds+=DBM.DBMParameters.CalculationInterval-OutputTimestamp.UTCSeconds Mod DBM.DBMParameters.CalculationInterval ' Next calculation timestamp
            If InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds Then ' If calculation timestamp can be calculated
                Try
                    Value=DBMRtCalculator.DBM.Calculate(InputDBMPointDriver,DBMCorrelationPoints,InputTimestamp.LocalDate).Factor
                Catch
                    Value=Double.NaN ' Calculation error
                End Try
                Try
                    DirectCast(OutputDBMPointDriver.Point,PISDK.PIPoint).Data.UpdateValue(Value,InputTimestamp.LocalDate) ' Write value to output point
                Catch
                End Try
            End If
        End Sub

    End Class

End Namespace
