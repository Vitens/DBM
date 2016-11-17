Option Explicit
Option Strict

Imports System.Reflection
<assembly:AssemblyTitle("DBMRt")>
<assembly:AssemblyCompany("Vitens N.V.")>
<assembly:AssemblyProduct("Dynamic Bandwidth Monitor Real-time")>
<assembly:AssemblyCopyright("J.H. Fitié, Vitens N.V.")>
<assembly:AssemblyVersion("2.4.*")>

Public Class DBMRt

    Inherits OSIsoft.PI.ACE.PIACENetClassModule

    Private Class DBMRt

        Private Class DBMPIServer

            Private Class DBMPIPoint

                Private Class CorrelationPIPoint

                    Public PIPoint As PISDK.PIPoint
                    Public SubstractSelf As Boolean

                    Public Sub New(ByVal PIPoint As PISDK.PIPoint,ByVal SubstractSelf As Boolean)
                        Me.PIPoint=PIPoint
                        Me.SubstractSelf=SubstractSelf
                    End Sub

                End Class

                Private InputPIPoint,OutputPIPoint As PISDK.PIPoint
                Private CorrelationPIPoints As Collections.Generic.List(Of CorrelationPIPoint)
                Private CalcTimestamps As Collections.Generic.List(Of Date)

                Public Sub New(ByVal InputPIPoint As PISDK.PIPoint,ByVal OutputPIPoint As PISDK.PIPoint)
                    Dim ExDesc,FieldsA(),FieldsB() As String
                    Me.InputPIPoint=InputPIPoint
                    Me.OutputPIPoint=OutputPIPoint
                    Me.CorrelationPIPoints=New Collections.Generic.List(Of CorrelationPIPoint)
                    ExDesc=Me.OutputPIPoint.PointAttributes("ExDesc").Value.ToString
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
                    Me.CalcTimestamps=New Collections.Generic.List(Of Date)
                End Sub

                Private Sub AddCorrelationPIPoint(ByVal PIPoint As PISDK.PIPoint,ByVal SubstractSelf As Boolean)
                    Me.CorrelationPIPoints.Add(New CorrelationPIPoint(PIPoint,SubstractSelf))
                End Sub

                Public Sub CalculatePoint
                    Dim InputTimestamp,OutputTimestamp As PITimeServer.PITime
                    Dim Value,NewValue As Double
                    Dim Annotation As String
                    Dim PIAnnotations As PISDK.PIAnnotations
                    Dim PINamedValues As PISDKCommon.NamedValues
                    Dim PIValues As PISDK.PIValues
                    InputTimestamp=Me.InputPIPoint.Data.Snapshot.TimeStamp
                    For Each thisCorrelationPIPoint As CorrelationPIPoint In Me.CorrelationPIPoints
                        InputTimestamp.UTCSeconds=Math.Min(InputTimestamp.UTCSeconds,thisCorrelationPIPoint.PIPoint.Data.Snapshot.TimeStamp.UTCSeconds)
                    Next
                    OutputTimestamp=Me.OutputPIPoint.Data.Snapshot.TimeStamp
                    If DateDiff("d",OutputTimestamp.LocalDate,Now())>DBMRtConstants.MaxCalculationAge Then
                        OutputTimestamp.LocalDate=DateAdd("d",-DBMRtConstants.MaxCalculationAge,Now())
                    End If
                    OutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval-OutputTimestamp.UTCSeconds Mod DBMConstants.CalculationInterval
                    Do While InputTimestamp.UTCSeconds>=OutputTimestamp.UTCSeconds
                        Me.CalcTimestamps.Add(OutputTimestamp.LocalDate)
                        OutputTimestamp.UTCSeconds+=DBMConstants.CalculationInterval
                    Loop
                    If Me.CalcTimestamps.Count>0 Then
                        Me.CalcTimestamps.Sort
                        Annotation=""
                        PIAnnotations=New PISDK.PIAnnotations
                        PINamedValues=New PISDKCommon.NamedValues
                        PIValues=New PISDK.PIValues
                        PIValues.ReadOnly=False
                        Try
                            If Me.CorrelationPIPoints.Count=0 Then
                                Value=DBM.Calculate(Me.InputPIPoint,Nothing,Me.CalcTimestamps(Me.CalcTimestamps.Count-1)).Factor
                            Else
                                Value=0
                                For Each thisCorrelationPIPoint As CorrelationPIPoint In Me.CorrelationPIPoints
                                    NewValue=DBM.Calculate(Me.InputPIPoint,thisCorrelationPIPoint.PIPoint,Me.CalcTimestamps(Me.CalcTimestamps.Count-1),thisCorrelationPIPoint.SubstractSelf).Factor
                                    If NewValue=0 Then
                                        Exit For
                                    End If
                                    If (Value=0 And Math.Abs(NewValue)>1) Or (Math.Abs(NewValue)<=1 And ((NewValue<0 And NewValue>=-1 And (NewValue<Value Or Math.Abs(Value)>1)) Or (NewValue>0 And NewValue<=1 And (NewValue>Value Or Math.Abs(Value)>1) And Not(Value<0 And Value>=-1)))) Then
                                        Value=NewValue
                                        If Math.Abs(Value)>0 And Math.Abs(Value)<=1 Then
                                            Annotation="Exception suppressed by \\" & thisCorrelationPIPoint.PIPoint.Server.Name & "\" & thisCorrelationPIPoint.PIPoint.Name & " due to " & CStr(IIf(Value<0,"anti","")) & "correlation (r=" & Value & ")."
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
                        PIValues.Add(Me.CalcTimestamps(Me.CalcTimestamps.Count-1),Value,PINamedValues)
                        Try
                            Me.OutputPIPoint.Data.UpdateValues(PIValues)
                        Catch
                        End Try
                        Me.CalcTimestamps.RemoveAt(Me.CalcTimestamps.Count-1)
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
