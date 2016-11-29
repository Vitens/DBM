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
    Private DBMRtEventPipeWatchers As Collections.Generic.List(Of DBMRtEventPipeWatcher)
    Private DBMRtCorrelationPIPoints As Collections.Generic.List(Of DBMRtCorrelationPIPoint)
    Private CalcTimestamps As Collections.Generic.List(Of Date)

    Public Sub New(ByVal InputPIPoint As PISDK.PIPoint,ByVal OutputPIPoint As PISDK.PIPoint)
        Dim ExDesc,FieldsA(),FieldsB() As String
        InputDBMPointDriver=New DBMPointDriver(InputPIPoint)
        OutputDBMPointDriver=New DBMPointDriver(OutputPIPoint)
        DBMRtEventPipeWatchers=New Collections.Generic.List(Of DBMRtEventPipeWatcher)
        DBMRtEventPipeWatchers.Add(New DBMRtEventPipeWatcher(InputPIPoint))
        DBMRtCorrelationPIPoints=New Collections.Generic.List(Of DBMRtCorrelationPIPoint)
        ExDesc=OutputDBMPointDriver.Point.PointAttributes("ExDesc").Value.ToString
        If Text.RegularExpressions.Regex.IsMatch(ExDesc,"^[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}(&[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}){0,}$") Then
            FieldsA=Split(ExDesc.ToString,"&")
            For Each thisField As String In FieldsA
                FieldsB=Split(thisField,":")
                Try
                    If DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)).Name<>"" Then
                        AddCorrelationPIPoint(DBMRtCalculator.PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)),Left(FieldsB(0),1)="-")
                    End If
                Catch
                End Try
            Next
        End If
        CalcTimestamps=New Collections.Generic.List(Of Date)
    End Sub

    Private Sub AddCorrelationPIPoint(ByVal PIPoint As PISDK.PIPoint,ByVal SubstractSelf As Boolean)
        DBMRtCorrelationPIPoints.Add(New DBMRtCorrelationPIPoint(PIPoint,SubstractSelf))
        DBMRtEventPipeWatchers.Add(New DBMRtEventPipeWatcher(PIPoint))
    End Sub

    Public Sub GetCalculationTimestamps(ByVal InputTimestamp As PITimeServer.PITime,OutputTimestamp As PITimeServer.PITime)
        Do While InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds
            If Not CalcTimestamps.Contains(OutputTimestamp.LocalDate) Then
                CalcTimestamps.Add(OutputTimestamp.LocalDate)
            End If
            OutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval
        Loop
    End Sub

    Public Sub GetRecalculationTimestamps(OutputTimestamp As PITimeServer.PITime)
        Dim EventObject As PISDK.PIEventObject
        Dim PointValue As PISDK.PointValue
        Dim RecalcBaseTimestamp As PITimeServer.PITime
        Dim DBMRtEventPipeWatcherDBMPointIndex,PatternCounter,OffsetCounter As Integer
        Dim RecalcTimestamp As Date
        For Each thisDBMRtEventPipeWatcher As DBMRtEventPipeWatcher In DBMRtEventPipeWatchers
            Do While thisDBMRtEventPipeWatcher.EventPipe.Count>0 And CalcTimestamps.Count<DBMRtConstants.MaxCalcTimestamps
                EventObject=thisDBMRtEventPipeWatcher.EventPipe.Take
                PointValue=CType(EventObject.EventData,PISDK.PointValue)
                If PointValue.PIValue.TimeStamp.UTCSeconds<OutputTimestamp.UTCSeconds Then
                    ' TODO: Calculate from previous to next value per interval
                    RecalcBaseTimestamp=PointValue.PIValue.TimeStamp
                    RecalcBaseTimestamp.UTCSeconds-=RecalcBaseTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval
                    DBMRtEventPipeWatcherDBMPointIndex=DBMRtCalculator.DBM.DBMPointDriverIndex(New DBMPointDriver(thisDBMRtEventPipeWatcher.PIPoint))
                    DBMRtCalculator.DBM.DBMPoints(DBMRtEventPipeWatcherDBMPointIndex).DBMDataManager.InvalidateCache(RecalcBaseTimestamp.LocalDate)
                    For PatternCounter=0 To DBMConstants.ComparePatterns
                        For OffsetCounter=0 To DBMConstants.EMAPreviousPeriods+DBMConstants.CorrelationPreviousPeriods
                            RecalcTimestamp=DateAdd("d",PatternCounter*7,DateAdd("s",OffsetCounter*DBMConstants.CalculationInterval,RecalcBaseTimestamp.LocalDate))
                            If DateDiff("d",RecalcTimestamp,Now())<=DBMRtConstants.MaxCalculationAge And RecalcTimestamp<OutputTimestamp.LocalDate Then
                                If Not CalcTimestamps.Contains(RecalcTimestamp) Then
                                    CalcTimestamps.Add(RecalcTimestamp)
                                End If
                            End If
                        Next OffsetCounter
                    Next PatternCounter
                End If
            Loop
        Next
    End Sub

    Public Function GetInputTimestamp As PITimeServer.PITime
        GetInputTimestamp=InputDBMPointDriver.Point.Data.Snapshot.TimeStamp
        For Each thisDBMRtCorrelationPIPoint As DBMRtCorrelationPIPoint In DBMRtCorrelationPIPoints
            GetInputTimestamp.UTCSeconds=Math.Min(GetInputTimestamp.UTCSeconds,thisDBMRtCorrelationPIPoint.DBMPointDriver.Point.Data.Snapshot.TimeStamp.UTCSeconds)
        Next
        Return GetInputTimestamp
    End Function

    Public Function GetNextOutputTimestamp As PITimeServer.PITime
        GetNextOutputTimestamp=OutputDBMPointDriver.Point.Data.Snapshot.TimeStamp
        If DateDiff("d",GetNextOutputTimestamp.LocalDate,Now())>DBMRtConstants.MaxCalculationAge Then
            GetNextOutputTimestamp.LocalDate=DateAdd("d",-DBMRtConstants.MaxCalculationAge,Now())
        End If
        GetNextOutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval-GetNextOutputTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval
        Return GetNextOutputTimestamp
    End Function

    Public Sub Calculate
        Dim InputTimestamp,OutputTimestamp As PITimeServer.PITime
        Dim Annotation As String
        Dim PIAnnotations As PISDK.PIAnnotations
        Dim PINamedValues As PISDKCommon.NamedValues
        Dim PIValues As PISDK.PIValues
        Dim Value,NewValue As Double
        InputTimestamp=GetInputTimestamp
        OutputTimestamp=GetNextOutputTimestamp
        GetCalculationTimestamps(InputTimestamp,OutputTimestamp)
        GetRecalculationTimestamps(OutputTimestamp)
        If CalcTimestamps.Count>0 Then
            CalcTimestamps.Sort
            Annotation=""
            PIAnnotations=New PISDK.PIAnnotations
            PINamedValues=New PISDKCommon.NamedValues
            PIValues=New PISDK.PIValues
            PIValues.ReadOnly=False
            Try
                If DBMRtCorrelationPIPoints.Count=0 Then
                    Value=DBMRtCalculator.DBM.Calculate(InputDBMPointDriver,Nothing,CalcTimestamps(CalcTimestamps.Count-1)).Factor
                Else
                    Value=0
                    For Each thisDBMRtCorrelationPIPoint As DBMRtCorrelationPIPoint In DBMRtCorrelationPIPoints
                        NewValue=DBMRtCalculator.DBM.Calculate(InputDBMPointDriver,thisDBMRtCorrelationPIPoint.DBMPointDriver,CalcTimestamps(CalcTimestamps.Count-1),thisDBMRtCorrelationPIPoint.SubstractSelf).Factor
                        If NewValue=0 Then
                            Exit For
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
            PIValues.Add(CalcTimestamps(CalcTimestamps.Count-1),Value,PINamedValues)
            Try
                OutputDBMPointDriver.Point.Data.UpdateValues(PIValues)
            Catch
            End Try
            CalcTimestamps.RemoveAt(CalcTimestamps.Count-1)
        End If
    End Sub

End Class
