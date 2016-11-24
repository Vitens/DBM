Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMStatistics

    Public Count As Integer
    Public Slope,Intercept,StDevSLinReg,ModifiedCorrelation As Double

    Public Sub Calculate(ByVal DataY() As Double,Optional ByVal DataX() As Double=Nothing)
        Dim SumX,SumXX,SumY,SumYY,SumXY As Double
        Dim i As Integer
        Me.Count=0
        For i=0 To DataY.Length-1
            If Not Double.IsNaN(DataY(i)) Then
                If DataX Is Nothing Then
                    SumX+=i
                    SumXX+=i^2
                    SumXY+=i*DataY(i)
                Else
                    SumX+=DataX(i)
                    SumXX+=DataX(i)^2
                    SumXY+=DataX(i)*DataY(i)
                End If
                SumY+=DataY(i)
                SumYY+=DataY(i)^2
                Me.Count+=1
            End If
        Next i
        Me.Slope=(Me.Count*SumXY-SumX*SumY)/(Me.Count*SumXX-SumX^2)
        Me.Intercept=(SumX*SumXY-SumY*SumXX)/(SumX^2-Me.Count*SumXX)
        Me.StDevSLinReg=0
        For i=0 to DataY.Length-1
            If Not Double.IsNaN(DataY(i)) Then
                Me.StDevSLinReg+=(DataY(i)-i*Me.Slope-Me.Intercept)^2
            End If
        Next i
        Me.StDevSLinReg=Math.Sqrt(Me.StDevSLinReg/(Me.Count-2)) ' n-2 is used because two parameters (slope and intercept) were estimated in order to estimate the sum of squares
        Me.ModifiedCorrelation=SumXY/Math.Sqrt(SumXX)/Math.Sqrt(SumYY) ' Average is not removed, as expected average is zero
    End Sub

End Class
