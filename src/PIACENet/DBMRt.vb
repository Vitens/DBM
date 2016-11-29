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

Imports System.Reflection
<assembly:AssemblyTitle("DBMRt")>
<assembly:AssemblyVersion("2.6.0.*")>
<assembly:AssemblyProduct("Dynamic Bandwidth Monitor Real-time")>
<assembly:AssemblyDescription("Leak detection method implemented in a real-time data historian")>
<assembly:AssemblyCopyright("Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.")>
<assembly:AssemblyCompany("Vitens N.V.")>

Public Class DBMRt

    Inherits OSIsoft.PI.ACE.PIACENetClassModule

    Private Class DBMRt

        Private Class DBMPIServer

            Private Class DBMPIPoint

                Private Class EventPipeWatcher

                    Public PIPoint As PISDK.PIPoint
                    Public EventPipe As PISDK.IEventPipe2

                    Public Sub New(ByVal PIPoint As PISDK.PIPoint)
                        Dim PIData2 As PISDK.IPIData2
                        Me.PIPoint=PIPoint
                        PIData2=CType(Me.PIPoint.Data,PISDK.IPIData2)
                        EventPipe=CType(PIData2.ArchiveEventPipe,PISDK.IEventPipe2)
                    End Sub

                End Class

                Private Class CorrelationPIPoint

                    Public DBMPointDriver As DBMPointDriver
                    Public SubstractSelf As Boolean

                    Public Sub New(ByVal PIPoint As PISDK.PIPoint,ByVal SubstractSelf As Boolean)
                        DBMPointDriver=New DBMPointDriver(PIPoint)
                        Me.SubstractSelf=SubstractSelf
                    End Sub

                End Class

                Private InputDBMPointDriver,OutputDBMPointDriver As DBMPointDriver
                Private EventPipeWatchers As Collections.Generic.List(Of EventPipeWatcher)
                Private CorrelationPIPoints As Collections.Generic.List(Of CorrelationPIPoint)
                Private CalcTimestamps As Collections.Generic.List(Of Date)

                Public Sub New(ByVal InputPIPoint As PISDK.PIPoint,ByVal OutputPIPoint As PISDK.PIPoint)
                    Dim ExDesc,FieldsA(),FieldsB() As String
                    InputDBMPointDriver=New DBMPointDriver(InputPIPoint)
                    OutputDBMPointDriver=New DBMPointDriver(OutputPIPoint)
                    EventPipeWatchers=New Collections.Generic.List(Of EventPipeWatcher)
                    EventPipeWatchers.Add(New EventPipeWatcher(InputPIPoint))
                    CorrelationPIPoints=New Collections.Generic.List(Of CorrelationPIPoint)
                    ExDesc=OutputDBMPointDriver.Point.PointAttributes("ExDesc").Value.ToString
                    If Text.RegularExpressions.Regex.IsMatch(ExDesc,"^[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}(&[-]{0,1}[a-zA-Z0-9][a-zA-Z0-9_\.-]{0,}:[^:?*&]{1,}){0,}$") Then
                        FieldsA=Split(ExDesc.ToString,"&")
                        For Each thisField As String In FieldsA
                            FieldsB=Split(thisField,":")
                            Try
                                If PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)).Name<>"" Then
                                    AddCorrelationPIPoint(PISDK.Servers(Mid(FieldsB(0),1+CInt(IIf(Left(FieldsB(0),1)="-",1,0)))).PIPoints(FieldsB(1)),Left(FieldsB(0),1)="-")
                                End If
                            Catch
                            End Try
                        Next
                    End If
                    CalcTimestamps=New Collections.Generic.List(Of Date)
                End Sub

                Private Sub AddCorrelationPIPoint(ByVal PIPoint As PISDK.PIPoint,ByVal SubstractSelf As Boolean)
                    CorrelationPIPoints.Add(New CorrelationPIPoint(PIPoint,SubstractSelf))
                    EventPipeWatchers.Add(New EventPipeWatcher(PIPoint))
                End Sub

                Public Sub CalculatePoint
                    Dim InputTimestamp,OutputTimestamp,RecalcBaseTimestamp As PITimeServer.PITime
                    Dim EventObject As PISDK.PIEventObject
                    Dim PointValue As PISDK.PointValue
                    Dim EventPipeWatcherDBMPointIndex,i,j As Integer
                    Dim RecalcTimestamp As Date
                    Dim Value,NewValue As Double
                    Dim Annotation As String
                    Dim PIAnnotations As PISDK.PIAnnotations
                    Dim PINamedValues As PISDKCommon.NamedValues
                    Dim PIValues As PISDK.PIValues
                    InputTimestamp=InputDBMPointDriver.Point.Data.Snapshot.TimeStamp
                    For Each thisCorrelationPIPoint As CorrelationPIPoint In CorrelationPIPoints
                        InputTimestamp.UTCSeconds=Math.Min(InputTimestamp.UTCSeconds,thisCorrelationPIPoint.DBMPointDriver.Point.Data.Snapshot.TimeStamp.UTCSeconds)
                    Next
                    OutputTimestamp=OutputDBMPointDriver.Point.Data.Snapshot.TimeStamp
                    If DateDiff("d",OutputTimestamp.LocalDate,Now())>DBMRtConstants.MaxCalculationAge Then
                        OutputTimestamp.LocalDate=DateAdd("d",-DBMRtConstants.MaxCalculationAge,Now())
                    End If
                    OutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval-OutputTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval
                    For Each thisEventPipeWatcher As EventPipeWatcher In EventPipeWatchers
                        Do While thisEventPipeWatcher.EventPipe.Count>0 ' TODO: Limit amount of events processed per interval
                            EventObject=thisEventPipeWatcher.EventPipe.Take
                            PointValue=CType(EventObject.EventData,PISDK.PointValue)
                            If PointValue.PIValue.TimeStamp.UTCSeconds<OutputTimestamp.UTCSeconds Then
                                ' TODO: Calculate from previous to next value per interval
                                RecalcBaseTimestamp=PointValue.PIValue.TimeStamp
                                RecalcBaseTimestamp.UTCSeconds-=RecalcBaseTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval
                                EventPipeWatcherDBMPointIndex=DBM.DBMPointDriverIndex(New DBMPointDriver(thisEventPipeWatcher.PIPoint))
                                DBM.DBMPoints(EventPipeWatcherDBMPointIndex).DBMDataManager.InvalidateCache(RecalcBaseTimestamp.LocalDate)
                                For i=0 To DBMConstants.ComparePatterns
                                    For j=0 To DBMConstants.EMAPreviousPeriods+DBMConstants.CorrelationPreviousPeriods
                                        RecalcTimestamp=DateAdd("d",i*7,DateAdd("s",j*DBMConstants.CalculationInterval,RecalcBaseTimestamp.LocalDate))
                                        If DateDiff("d",RecalcTimestamp,Now())<=DBMRtConstants.MaxCalculationAge And RecalcTimestamp<OutputTimestamp.LocalDate Then
                                            If Not CalcTimestamps.Contains(RecalcTimestamp) Then
                                                CalcTimestamps.Add(RecalcTimestamp)
                                            End If
                                        End If
                                    Next j
                                Next i
                            End If
                        Loop
                    Next
                    Do While InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds
                        CalcTimestamps.Add(OutputTimestamp.LocalDate)
                        OutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval
                    Loop
                    If CalcTimestamps.Count>0 Then
                        CalcTimestamps.Sort
                        Annotation=""
                        PIAnnotations=New PISDK.PIAnnotations
                        PINamedValues=New PISDKCommon.NamedValues
                        PIValues=New PISDK.PIValues
                        PIValues.ReadOnly=False
                        Try
                            If CorrelationPIPoints.Count=0 Then
                                Value=DBM.Calculate(InputDBMPointDriver,Nothing,CalcTimestamps(CalcTimestamps.Count-1)).Factor
                            Else
                                Value=0
                                For Each thisCorrelationPIPoint As CorrelationPIPoint In CorrelationPIPoints
                                    NewValue=DBM.Calculate(InputDBMPointDriver,thisCorrelationPIPoint.DBMPointDriver,CalcTimestamps(CalcTimestamps.Count-1),thisCorrelationPIPoint.SubstractSelf).Factor
                                    If NewValue=0 Then
                                        Exit For
                                    End If
                                    If (Value=0 And Math.Abs(NewValue)>1) Or (Math.Abs(NewValue)<=1 And ((NewValue<0 And NewValue>=-1 And (NewValue<Value Or Math.Abs(Value)>1)) Or (NewValue>0 And NewValue<=1 And (NewValue>Value Or Math.Abs(Value)>1) And Not(Value<0 And Value>=-1)))) Then
                                        Value=NewValue
                                        If Math.Abs(Value)>0 And Math.Abs(Value)<=1 Then
                                            Annotation="Exception suppressed by \\" & thisCorrelationPIPoint.DBMPointDriver.Point.Server.Name & "\" & thisCorrelationPIPoint.DBMPointDriver.Point.Name & " due to " & CStr(IIf(Value<0,"anti","")) & "correlation (r=" & Value & ")."
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

            Private PIServer As PISDK.Server
            Private DBMPIPoints As Collections.Generic.List(Of DBMPIPoint)

            Public Sub New(ByVal PIServer As PISDK.Server,Optional ByVal TagFilter As String="*")
                Dim InstrTag,Fields() As String
                Me.PIServer=PIServer
                DBMPIPoints=New Collections.Generic.List(Of DBMPIPoint)
                Try
                    For Each thisPIPoint As PISDK.PIPoint In Me.PIServer.GetPointsSQL("PIpoint.Tag='" & TagFilter & "' AND PIpoint.PointSource='dbmrt' AND PIpoint.Scan=1")
                        InstrTag=thisPIPoint.PointAttributes("InstrumentTag").Value.ToString
                        If Text.RegularExpressions.Regex.IsMatch(InstrTag,"^[a-zA-Z0-9_\.-]{1,}:[^:?*&]{1,}$") Then
                            Fields=Split(InstrTag.ToString,":")
                            Try
                                If PISDK.Servers(Fields(0)).PIPoints(Fields(1)).Name<>"" Then
                                    DBMPIPoints.Add(New DBMPIPoint(PISDK.Servers(Fields(0)).PIPoints(Fields(1)),thisPIPoint))
                                End If
                            Catch
                            End Try
                        End If
                    Next
                Catch
                End Try
            End Sub

            Public Sub CalculatePoints
                For Each thisDBMPIPoint As DBMPIPoint In DBMPIPoints
                    Try
                        thisDBMPIPoint.CalculatePoint
                    Catch
                    End Try
                Next
            End Sub

        End Class

        Private Shared PISDK As New PISDK.PISDK
        Private Shared DBM As New DBM
        Private DBMPIServers As Collections.Generic.List(Of DBMRt.DBMPIServer)

        Public Sub New(Optional ByVal DefaultServerOnly As Boolean=False,Optional ByVal TagFilter As String="*")
            DBMPIServers=New Collections.Generic.List(Of DBMRt.DBMPIServer)
            For Each thisServer As PISDK.Server In PISDK.Servers
                If Not DefaultServerOnly Or thisServer Is PISDK.Servers.DefaultServer Then
                    DBMPIServers.Add(New DBMPIServer(thisServer,TagFilter))
                End If
            Next
        End Sub

        Public Sub CalculateServers
            For Each thisDBMPIServer As DBMPIServer In DBMPIServers
                thisDBMPIServer.CalculatePoints
            Next
        End Sub

    End Class

    Public Overrides Sub ACECalculations()
    End Sub

    Protected Overrides Sub InitializePIACEPoints()
    End Sub

    Protected Overrides Sub ModuleDependentInitialization()
        Dim _DBMRt As New DBMRt(True)
        Do While True
            _DBMRt.CalculateServers
            Threading.Thread.Sleep(DBMRtConstants.CalculationDelay*1000)
        Loop
    End Sub

    Protected Overrides Sub ModuleDependentTermination()
    End Sub

End Class
