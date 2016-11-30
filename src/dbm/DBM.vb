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
<assembly:AssemblyTitle("DBM")>
<assembly:AssemblyVersion("1.3.1.*")>
<assembly:AssemblyProduct("Dynamic Bandwidth Monitor")>
<assembly:AssemblyDescription("Leak detection method implemented in a real-time data historian")>
<assembly:AssemblyCopyright("Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.")>
<assembly:AssemblyCompany("Vitens N.V.")>

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

    Public Function Calculate(ByVal InputDBMPointDriver As DBMPointDriver,ByVal CorrelationDBMPointDriver As DBMPointDriver,ByVal Timestamp As DateTime,Optional ByVal SubstractInputPointFromCorrelationPoint As Boolean=False) As DBMResult
        Dim InputDBMPointDriverIndex,CorrelationDBMPointDriverIndex As Integer
        Dim CorrelationDBMResult As DBMResult
        Dim AbsErrorStats,RelErrorStats As New DBMStatistics
        InputDBMPointDriverIndex=DBMPointDriverIndex(InputDBMPointDriver)
        Calculate=DBMPoints(InputDBMPointDriverIndex).Calculate(Timestamp,True,Not IsNothing(CorrelationDBMPointDriver)) ' Calculate for input point
        If Calculate.Factor<>0 And Not IsNothing(CorrelationDBMPointDriver) Then ' If an event is found and a correlation point is available
            CorrelationDBMPointDriverIndex=DBMPointDriverIndex(CorrelationDBMPointDriver)
            If SubstractInputPointFromCorrelationPoint Then ' If pattern of correlation point contains input point
                CorrelationDBMResult=DBMPoints(CorrelationDBMPointDriverIndex).Calculate(Timestamp,False,True,DBMPoints(InputDBMPointDriverIndex)) ' Calculate for correlation point, substract input point
            Else
                CorrelationDBMResult=DBMPoints(CorrelationDBMPointDriverIndex).Calculate(Timestamp,False,True) ' Calculate for correlation point
            End If
            AbsErrorStats.Calculate(DBMPoints(CorrelationDBMPointDriverIndex).AbsoluteError,DBMPoints(InputDBMPointDriverIndex).AbsoluteError) ' Absolute error compared to prediction
            RelErrorStats.Calculate(DBMPoints(CorrelationDBMPointDriverIndex).RelativeError,DBMPoints(InputDBMPointDriverIndex).RelativeError) ' Relative error compared to prediction
            If RelErrorStats.ModifiedCorrelation>DBMConstants.CorrelationThreshold Then ' Suppress event due to correlation of relative error
                Calculate.Factor=RelErrorStats.ModifiedCorrelation
            End If
            If Not SubstractInputPointFromCorrelationPoint And AbsErrorStats.ModifiedCorrelation<-DBMConstants.CorrelationThreshold Then ' Suppress event due to anticorrelation of absolute error (unmeasured supply)
                Calculate.Factor=AbsErrorStats.ModifiedCorrelation
            End If
        End If
        Return Calculate
    End Function

End Class
