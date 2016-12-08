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

Namespace DBM

    Public Class DBMStatistics

        Public Count As Integer
        Public Slope,Intercept,StDevSLinReg,Correlation,ModifiedCorrelation,Determination As Double

        Public Sub Calculate(DataY() As Double,Optional DataX() As Double=Nothing)
            Dim SumX,SumXX,SumY,SumYY,SumXY As Double
            Dim i As Integer
            If DataX Is Nothing Then
                ReDim DataX(DataY.Length-1)
                For i=0 To DataX.Length-1
                    DataX(i)=i
                Next i
            End If
            Count=0
            For i=0 To DataY.Length-1
                If Not Double.IsNaN(DataX(i)) And Not Double.IsNaN(DataY(i)) Then
                    SumX+=DataX(i)
                    SumXX+=DataX(i)^2
                    SumXY+=DataX(i)*DataY(i)
                    SumY+=DataY(i)
                    SumYY+=DataY(i)^2
                    Count+=1
                End If
            Next i
            Slope=(Count*SumXY-SumX*SumY)/(Count*SumXX-SumX^2)
            Intercept=(SumX*SumXY-SumY*SumXX)/(SumX^2-Count*SumXX)
            StDevSLinReg=0
            For i=0 to DataY.Length-1
                If Not Double.IsNaN(DataX(i)) And Not Double.IsNaN(DataY(i)) Then
                    StDevSLinReg+=(DataY(i)-DataX(i)*Slope-Intercept)^2
                End If
            Next i
            StDevSLinReg=Math.Sqrt(StDevSLinReg/(Count-2)) ' n-2 is used because two parameters (slope and intercept) were estimated in order to estimate the sum of squares
            Correlation=(Count*SumXY-SumX*SumY)/Math.Sqrt((Count*SumXX-SumX^2)*(Count*SumYY-SumY^2)) ' Wikipedia: A number that quantifies some type of correlation and dependence, meaning statistical relationships between two or more random variables or observed data values
            ModifiedCorrelation=SumXY/Math.Sqrt(SumXX)/Math.Sqrt(SumYY) ' Average is not removed, as expected average is zero
            Determination=Correlation^2 ' Wikipedia: A number that indicates the proportion of the variance in the dependent variable that is predictable from the independent variable
        End Sub

    End Class

End Namespace
