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
    Private DBMRtCorrelationPIPoints As Collections.Generic.List(Of DBMRtCorrelationPIPoint)

    Public Sub New(ByVal InputPIPoint As PISDK.PIPoint,ByVal OutputPIPoint As PISDK.PIPoint)
        Dim ExDesc,FieldsA(),FieldsB() As String
        InputDBMPointDriver=New DBMPointDriver(InputPIPoint)
        OutputDBMPointDriver=New DBMPointDriver(OutputPIPoint)
        DBMRtCorrelationPIPoints=New Collections.Generic.List(Of DBMRtCorrelationPIPoint)
        ExDesc=OutputDBMPointDriver.Point.PointAttributes("ExDesc").Value.ToString
        If Text.RegularExpressions.Regex.IsMatch(ExDesc,"^[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}(&[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}){0,}$") Then
            FieldsA=Split(ExDesc.ToString,"&")
            For Each thisField As String In FieldsA
                FieldsB=Split(thisField,":")
                Try
                    If DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)).Name<>"" Then
                        DBMRtCorrelationPIPoints.Add(New DBMRtCorrelationPIPoint(DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)),Left(FieldsB(0),1)="-"))
                    End If
                Catch
                End Try
            Next
        End If
    End Sub

    Public Sub Calculate
        Dim InputTimestamp,OutputTimestamp As PITimeServer.PITime
        Dim Annotation As String
        Dim PIAnnotations As PISDK.PIAnnotations
        Dim PINamedValues As PISDKCommon.NamedValues
        Dim PIValues As PISDK.PIValues
        Dim Value,NewValue As Double
        InputTimestamp=InputDBMPointDriver.Point.Data.Snapshot.TimeStamp ' Timestamp of input point
        For Each thisDBMRtCorrelationPIPoint As DBMRtCorrelationPIPoint In DBMRtCorrelationPIPoints ' Check timestamp of correlation points
            InputTimestamp.UTCSeconds=Math.Min(InputTimestamp.UTCSeconds,thisDBMRtCorrelationPIPoint.DBMPointDriver.Point.Data.Snapshot.TimeStamp.UTCSeconds) ' Timestamp of correlation point, keep earliest
        Next
        InputTimestamp.UTCSeconds-=DBMConstants.CalculationInterval+InputTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval ' Can calculate output until (inclusive)
        OutputTimestamp=OutputDBMPointDriver.Point.Data.Snapshot.TimeStamp ' Timestamp of output point
        OutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval-OutputTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval ' Next calculation timestamp
        If InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds Then ' If calculation timestamp can be calculated
            Annotation=""
            PIAnnotations=New PISDK.PIAnnotations
            PINamedValues=New PISDKCommon.NamedValues
            PIValues=New PISDK.PIValues
            PIValues.ReadOnly=False
            Try
                If DBMRtCorrelationPIPoints.Count=0 Then
                    Value=DBMRtCalculator.DBM.Calculate(InputDBMPointDriver,Nothing,InputTimestamp.LocalDate).Factor
                Else
                    Value=0
                    For Each thisDBMRtCorrelationPIPoint As DBMRtCorrelationPIPoint In DBMRtCorrelationPIPoints
                        NewValue=DBMRtCalculator.DBM.Calculate(InputDBMPointDriver,thisDBMRtCorrelationPIPoint.DBMPointDriver,InputTimestamp.LocalDate,thisDBMRtCorrelationPIPoint.SubstractSelf).Factor
                        If NewValue=0 Then ' If there is no exception
                            Exit For ' Do not calculate next correlation points
                        End If
                        If (Value=0 And Math.Abs(NewValue)>1) Or (Math.Abs(NewValue)<=1 And ((NewValue<0 And NewValue>=-1 And (NewValue<Value Or Math.Abs(Value)>1)) Or (NewValue>0 And NewValue<=1 And (NewValue>Value Or Math.Abs(Value)>1) And Not(Value<0 And Value>=-1)))) Then
                            Value=NewValue
                            If Math.Abs(Value)>0 And Math.Abs(Value)<=1 Then
                                Annotation="Exception suppressed by \\" & thisDBMRtCorrelationPIPoint.DBMPointDriver.Point.Server.Name & "\" & thisDBMRtCorrelationPIPoint.DBMPointDriver.Point.Name & " due to " & CStr(IIf(Value<0,"anti","")) & "correlation (r=" & Value & ")."
                            End If
                        End If
                    Next
                End If
            Catch
                Value=Double.NaN
                Annotation="Calculation error."
            End Try
            If Annotation<>"" Then
                PIAnnotations.Add("","",Annotation,False,"String")
                PINamedValues.Add("Annotations",CObj(PIAnnotations))
            End If
            PIValues.Add(InputTimestamp.LocalDate,Value,PINamedValues)
            Try
                OutputDBMPointDriver.Point.Data.UpdateValues(PIValues)
            Catch
            End Try
        End If
    End Sub

End Class
