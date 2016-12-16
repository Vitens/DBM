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

<assembly:System.Reflection.AssemblyTitle("DBM")>

Namespace DBM

    Public Class DBM

        Public DBMPoints As New Collections.Generic.Dictionary(Of Object,DBMPoint)

        Public Shared Function DBMVersion As String
            Dim Ticks As Int64=DateTime.Now.Ticks
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileDescription & _
                " v" & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileVersion & vbCrLf & _
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).ProductName & vbCrLf & _
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).Comments & vbCrLf & vbCrLf & _
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).LegalCopyright & vbCrLf & vbCrLf & _
                " * Unit tests OK: " & CStr(DBMUnitTests.TestResults) & ", in " & Math.Round((DateTime.Now.Ticks-Ticks)/10000) & "ms" & vbCrLf
        End Function

        Private Function GetDBMPoint(DBMPointDriver As DBMPointDriver) As DBMPoint
            If Not DBMPoints.ContainsKey(DBMPointDriver.Point) Then
                DBMPoints.Add(DBMPointDriver.Point,New DBMPoint(DBMPointDriver)) ' Add to dictionary
            End If
            Return DBMPoints.Item(DBMPointDriver.Point)
        End Function

        Public Shared Function KeepOrSuppressEvent(Factor As Double,AbsErrModCorr As Double,RelErrModCorr As Double,SubstractSelf As Boolean) As Double
            KeepOrSuppressEvent=Factor
            If Not SubstractSelf And AbsErrModCorr<-DBMParameters.CorrelationThreshold Then ' If anticorrelation with adjacent measurement
                If KeepOrSuppressEvent<-DBMParameters.CorrelationThreshold And KeepOrSuppressEvent>=-1 Then ' If already suppressed due to anticorrelation
                    If AbsErrModCorr<KeepOrSuppressEvent Then ' Keep lowest value (strongest anticorrelation)
                        KeepOrSuppressEvent=AbsErrModCorr ' Suppress
                    End If
                Else ' Not already suppressed due to anticorrelation
                    KeepOrSuppressEvent=AbsErrModCorr ' Suppress
                End If
            ElseIf RelErrModCorr>DBMParameters.CorrelationThreshold Then ' If correlation with measurement
                If Not (KeepOrSuppressEvent<-DBMParameters.CorrelationThreshold And KeepOrSuppressEvent>=-1) Then ' If not already suppressed due to anticorrelation
                    If KeepOrSuppressEvent>DBMParameters.CorrelationThreshold And KeepOrSuppressEvent<=1 Then ' If already suppressed due to correlation
                        If RelErrModCorr>KeepOrSuppressEvent Then ' Keep highest value (strongest correlation)
                            KeepOrSuppressEvent=RelErrModCorr ' Suppress
                        End If
                    Else ' Not already suppressed due to correlation
                        KeepOrSuppressEvent=RelErrModCorr ' Suppress
                    End If
                End If
            End If
            Return KeepOrSuppressEvent
        End Function

        Public Function Calculate(InputDBMPointDriver As DBMPointDriver,DBMCorrelationPoints As Collections.Generic.List(Of DBMCorrelationPoint),Timestamp As DateTime) As DBMResult
            Dim CorrelationDBMResult As DBMResult
            Dim AbsErrorStats,RelErrorStats As New DBMStatistics
            Dim Factor As Double
            If DBMCorrelationPoints Is Nothing Then
                DBMCorrelationPoints=New Collections.Generic.List(Of DBMCorrelationPoint)
            End If
            Calculate=GetDBMPoint(InputDBMPointDriver).Calculate(Timestamp,True,DBMCorrelationPoints.Count>0) ' Calculate for input point
            If Calculate.Factor<>0 And DBMCorrelationPoints.Count>0 Then ' If an event is found and a correlation point is available
                For Each thisDBMCorrelationPoint As DBMCorrelationPoint In DBMCorrelationPoints
                    If thisDBMCorrelationPoint.SubstractSelf Then ' If pattern of correlation point contains input point
                        CorrelationDBMResult=GetDBMPoint(thisDBMCorrelationPoint.DBMPointDriver).Calculate(Timestamp,False,True,GetDBMPoint(InputDBMPointDriver)) ' Calculate for correlation point, substract input point
                    Else
                        CorrelationDBMResult=GetDBMPoint(thisDBMCorrelationPoint.DBMPointDriver).Calculate(Timestamp,False,True) ' Calculate for correlation point
                    End If
                    AbsErrorStats.Calculate(CorrelationDBMResult.AbsoluteErrors,Calculate.AbsoluteErrors) ' Absolute error compared to prediction
                    RelErrorStats.Calculate(CorrelationDBMResult.RelativeErrors,Calculate.RelativeErrors) ' Relative error compared to prediction
                    Factor=KeepOrSuppressEvent(Calculate.Factor,AbsErrorStats.ModifiedCorrelation,RelErrorStats.ModifiedCorrelation,thisDBMCorrelationPoint.SubstractSelf)
                    If Factor<>Calculate.Factor Then ' Has event been suppressed
                        Calculate.Factor=Factor
                        Calculate.AbsErrorStats=AbsErrorStats.ShallowCopy
                        Calculate.RelErrorStats=RelErrorStats.ShallowCopy
                        Calculate.SuppressedBy=thisDBMCorrelationPoint.DBMPointDriver ' Suppressed by
                    End If
                Next
            End If
            Return Calculate
        End Function

    End Class

End Namespace
