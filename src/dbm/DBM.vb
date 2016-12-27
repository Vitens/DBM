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

<assembly:System.Reflection.AssemblyTitle("DBM")>

Namespace DBM

    Public Class DBM

        Public Points As New Collections.Generic.Dictionary(Of Object,DBMPoint)

        Public Shared Function Version(Optional SkipUnitTests As Boolean=False) As String
            Dim Ticks As Int64=DateTime.Now.Ticks
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileDescription & _
                " v" & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileVersion & vbCrLf & _
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).ProductName & vbCrLf & _
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).Comments & vbCrLf & vbCrLf & _
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).LegalCopyright & vbCrLf & _
                If(SkipUnitTests,"",vbCrLf & " * Unit tests " & If(DBMUnitTests.TestResults,"PASSED","FAILED") & " in " & Math.Round((DateTime.Now.Ticks-Ticks)/10000) & "ms." & vbCrLf)
        End Function

        Private Function Point(PointDriver As DBMPointDriver) As DBMPoint
            If Not Points.ContainsKey(PointDriver.Point) Then
                Points.Add(PointDriver.Point,New DBMPoint(PointDriver)) ' Add to dictionary
            End If
            Return Points.Item(PointDriver.Point)
        End Function

        Public Shared Function Suppress(Factor As Double,AbsErrModCorr As Double,RelErrModCorr As Double,SubstractSelf As Boolean) As Double
            If Not SubstractSelf And AbsErrModCorr<-DBMParameters.CorrelationThreshold Then ' If anticorrelation with adjacent measurement
                If Factor<-DBMParameters.CorrelationThreshold And Factor>=-1 Then ' If already suppressed due to anticorrelation
                    If AbsErrModCorr<Factor Then ' Keep lowest value (strongest anticorrelation)
                        Return AbsErrModCorr ' Suppress
                    End If
                Else ' Not already suppressed due to anticorrelation
                    Return AbsErrModCorr ' Suppress
                End If
            ElseIf RelErrModCorr>DBMParameters.CorrelationThreshold Then ' If correlation with measurement
                If Not (Factor<-DBMParameters.CorrelationThreshold And Factor>=-1) Then ' If not already suppressed due to anticorrelation
                    If Factor>DBMParameters.CorrelationThreshold And Factor<=1 Then ' If already suppressed due to correlation
                        If RelErrModCorr>Factor Then ' Keep highest value (strongest correlation)
                            Return RelErrModCorr ' Suppress
                        End If
                    Else ' Not already suppressed due to correlation
                        Return RelErrModCorr ' Suppress
                    End If
                End If
            End If
            Return Factor
        End Function

        Public Function Result(InputPointDriver As DBMPointDriver,CorrelationPoints As Collections.Generic.List(Of DBMCorrelationPoint),Timestamp As DateTime) As DBMResult
            Dim CorrelationResult As DBMResult
            Dim AbsoluteErrorStats,RelativeErrorStats As New DBMStatistics
            Dim Factor As Double
            If CorrelationPoints Is Nothing Then
                CorrelationPoints=New Collections.Generic.List(Of DBMCorrelationPoint)
            End If
            Result=Point(InputPointDriver).Result(Timestamp,True,CorrelationPoints.Count>0) ' Calculate for input point
            If Result.Factor<>0 And CorrelationPoints.Count>0 Then ' If an event is found and a correlation point is available
                For Each CorrelationPoint In CorrelationPoints
                    If CorrelationPoint.SubstractSelf Then ' If pattern of correlation point contains input point
                        CorrelationResult=Point(CorrelationPoint.PointDriver).Result(Timestamp,False,True,Point(InputPointDriver)) ' Calculate for correlation point, substract input point
                    Else
                        CorrelationResult=Point(CorrelationPoint.PointDriver).Result(Timestamp,False,True) ' Calculate for correlation point
                    End If
                    AbsoluteErrorStats.Calculate(CorrelationResult.AbsoluteErrors,Result.AbsoluteErrors) ' Absolute error compared to prediction
                    RelativeErrorStats.Calculate(CorrelationResult.RelativeErrors,Result.RelativeErrors) ' Relative error compared to prediction
                    Factor=Suppress(Result.Factor,AbsoluteErrorStats.ModifiedCorrelation,RelativeErrorStats.ModifiedCorrelation,CorrelationPoint.SubstractSelf)
                    If Factor<>Result.Factor Then ' Has event been suppressed
                        Result.Factor=Factor
                        Result.AbsoluteErrorStats=AbsoluteErrorStats.ShallowCopy
                        Result.RelativeErrorStats=RelativeErrorStats.ShallowCopy
                        Result.SuppressedBy=CorrelationPoint.PointDriver ' Suppressed by
                    End If
                Next
            End If
            Return Result
        End Function

    End Class

End Namespace
