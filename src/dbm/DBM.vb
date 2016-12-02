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

Public Class DBM

    Private DBMDriver As DBMDriver
    Public DBMPoints(-1) As DBMPoint

    Public Sub New(Optional ByVal Data() As Object=Nothing)
        DBMDriver=New DBMDriver(Data)
    End Sub

    Public Function DBMPointDriverIndex(ByVal DBMPointDriver As DBMPointDriver) As Integer
        DBMPointDriverIndex=Array.FindIndex(DBMPoints,Function(FindDBMPoint)FindDBMPoint.DBMDataManager.DBMPointDriver.Point Is DBMPointDriver.Point)
        If DBMPointDriverIndex=-1 Then ' PointDriver not found
            ReDim Preserve DBMPoints(DBMPoints.Length) ' Add to array
            DBMPointDriverIndex=DBMPoints.Length-1
            DBMPoints(DBMPointDriverIndex)=New DBMPoint(DBMPointDriver)
        End If
        Return DBMPointDriverIndex
    End Function

    Public Shared Function KeepOrSuppressEvent(ByVal Factor As Double,ByVal AbsErrModCorr As Double,ByVal RelErrModCorr As Double,ByVal SubstractSelf As Boolean) As Double
        KeepOrSuppressEvent=Factor
        If Not SubstractSelf And AbsErrModCorr<-DBMConstants.CorrelationThreshold Then ' If anticorrelation with adjacent measurement
            If KeepOrSuppressEvent<-DBMConstants.CorrelationThreshold And KeepOrSuppressEvent>=-1 Then ' If already suppressed due to anticorrelation
                If AbsErrModCorr<KeepOrSuppressEvent Then ' Keep lowest value (strongest anticorrelation)
                    KeepOrSuppressEvent=AbsErrModCorr ' Suppress
                End If
            Else ' Not already suppressed due to anticorrelation
                KeepOrSuppressEvent=AbsErrModCorr ' Suppress
            End If
        End If
        If RelErrModCorr>DBMConstants.CorrelationThreshold Then ' If correlation with measurement
            If Not (KeepOrSuppressEvent<-DBMConstants.CorrelationThreshold And KeepOrSuppressEvent>=-1) Then ' If not already suppressed due to anticorrelation
                If KeepOrSuppressEvent>DBMConstants.CorrelationThreshold And KeepOrSuppressEvent<=1 Then ' If already suppressed due to correlation
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

    Public Function Calculate(ByVal InputDBMPointDriver As DBMPointDriver,ByVal DBMCorrelationPoints As Collections.Generic.List(Of DBMCorrelationPoint),ByVal Timestamp As DateTime) As DBMResult
        Dim InputDBMPointDriverIndex,CorrelationDBMPointDriverIndex As Integer
        Dim CorrelationDBMResult As DBMResult
        Dim AbsErrorStats,RelErrorStats As New DBMStatistics
        Dim NewValue As Double
        If IsNothing(DBMCorrelationPoints) Then DBMCorrelationPoints=New Collections.Generic.List(Of DBMCorrelationPoint)
        InputDBMPointDriverIndex=DBMPointDriverIndex(InputDBMPointDriver)
        Calculate=DBMPoints(InputDBMPointDriverIndex).Calculate(Timestamp,True,DBMCorrelationPoints.Count>0) ' Calculate for input point
        If Calculate.Factor<>0 And DBMCorrelationPoints.Count>0 Then ' If an event is found and a correlation point is available
            For Each thisDBMCorrelationPoint As DBMCorrelationPoint In DBMCorrelationPoints
                CorrelationDBMPointDriverIndex=DBMPointDriverIndex(thisDBMCorrelationPoint.DBMPointDriver)
                If thisDBMCorrelationPoint.SubstractSelf Then ' If pattern of correlation point contains input point
                    CorrelationDBMResult=DBMPoints(CorrelationDBMPointDriverIndex).Calculate(Timestamp,False,True,DBMPoints(InputDBMPointDriverIndex)) ' Calculate for correlation point, substract input point
                Else
                    CorrelationDBMResult=DBMPoints(CorrelationDBMPointDriverIndex).Calculate(Timestamp,False,True) ' Calculate for correlation point
                End If
                AbsErrorStats.Calculate(DBMPoints(CorrelationDBMPointDriverIndex).AbsoluteError,DBMPoints(InputDBMPointDriverIndex).AbsoluteError) ' Absolute error compared to prediction
                RelErrorStats.Calculate(DBMPoints(CorrelationDBMPointDriverIndex).RelativeError,DBMPoints(InputDBMPointDriverIndex).RelativeError) ' Relative error compared to prediction
                NewValue=KeepOrSuppressEvent(Calculate.Factor,AbsErrorStats.ModifiedCorrelation,RelErrorStats.ModifiedCorrelation,thisDBMCorrelationPoint.SubstractSelf)
                If NewValue<>Calculate.Factor Then ' Has event been suppressed
                    Calculate.Factor=NewValue
                    Calculate.SuppressedBy=thisDBMCorrelationPoint.DBMPointDriver ' Suppressed by
                End If
            Next
        End If
        Return Calculate
    End Function

End Class
