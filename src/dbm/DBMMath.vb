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

Namespace DBM

    Public Class DBMMath

        Private Shared Random As New Random

        Public Shared Function NormSInv(p As Double) As Double ' Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.
            ' Approximation of inverse standard normal CDF developed by Peter J. Acklam
            Const a1=-39.6968302866538,a2=220.946098424521,a3=-275.928510446969
            Const a4=138.357751867269,a5=-30.6647980661472,a6=2.50662827745924
            Const b1=-54.4760987982241,b2=161.585836858041,b3=-155.698979859887
            Const b4=66.8013118877197,b5=-13.2806815528857,c1=-7.78489400243029E-03
            Const c2=-0.322396458041136,c3=-2.40075827716184,c4=-2.54973253934373
            Const c5=4.37466414146497,c6=2.93816398269878,d1=7.78469570904146E-03
            Const d2=0.32246712907004,d3=2.445134137143,d4=3.75440866190742
            Const p_low=0.02425,p_high=1-p_low
            Dim q,r As Double
            If p<p_low Then
                q=Math.Sqrt(-2*Math.Log(p))
                Return (((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6)/((((d1*q+d2)*q+d3)*q+d4)*q+1)
            ElseIf p<=p_high Then
                q=p-0.5
                r=q*q
                Return (((((a1*r+a2)*r+a3)*r+a4)*r+a5)*r+a6)*q/(((((b1*r+b2)*r+b3)*r+b4)*r+b5)*r+1)
            Else
                q=Math.Sqrt(-2*Math.Log(1-p))
                Return -(((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6)/((((d1*q+d2)*q+d3)*q+d4)*q+1)
            End If
        End Function

        Public Shared Function TInv2T(p As Double,dof As Integer) As Double ' Returns the two-tailed inverse of the Student's t-distribution.
            ' Hill's approx. inverse t-dist.: Comm. of A.C.M Vol.13 No.10 1970 pg 620
            Dim a,b,c,d,x,y As Double
            If dof=1 Then
                p*=Math.PI/2
                Return Math.Cos(p)/Math.Sin(p)
            ElseIf dof=2 Then
                Return Math.Sqrt(2/(p*(2-p))-2)
            Else
                a=1/(dof-0.5)
                b=48/(a^2)
                c=((20700*a/b-98)*a-16)*a+96.36
                d=((94.5/(b+c)-3)/b+1)*Math.Sqrt(a*Math.PI/2)*dof
                x=d*p
                y=x^(2/dof)
                If y>a+0.05 Then
                    x=NormSInv(p/2)
                    y=x^2
                    If dof<5 Then
                        c+=0.3*(dof-4.5)*(x+0.6)
                    End If
                    c=(((d/2*x-0.5)*x-7)*x-2)*x+b+c
                    y=(((((0.4*y+6.3)*y+36)*y+94.5)/c-y-3)/b+1)*x
                    y=Math.Exp(a*y^2)-1
                Else
                    y=((1/(((dof+6)/(dof*y)-0.089*d-0.822)*(dof+2)*3)+0.5/(dof+4))*y-1)*(dof+1)/(dof+2)+1/y
                End If
                Return Math.Sqrt(dof*y)
            End If
        End Function

        Public Shared Function TInv(p As Double,dof As Integer) As Double ' Returns the left-tailed inverse of the Student's t-distribution.
            Return Math.Sign(p-0.5)*TInv2T(1-Math.Abs(p-0.5)*2,dof)
        End Function

        Public Shared Function MeanAbsoluteDeviationScaleFactor As Double ' Scale factor k
            Return Math.Sqrt(Math.PI/2)
        End Function

        Public Shared Function MedianAbsoluteDeviationScaleFactor(n As Integer) As Double ' Scale factor k
            If n<30 Then
                Return 1/TInv(0.75,n) ' n<30 Student's t-distribution
            Else
                Return 1/NormSInv(0.75) ' n>=30 Standard normal distribution
            End If
        End Function

        Public Shared Function ControlLimitRejectionCriterion(p As Double,n As Integer) As Double
            If n<30 Then
                Return TInv((p+1)/2,n) ' n<30 Student's t-distribution
            Else
                Return NormSInv((p+1)/2) ' n>=30 Standard normal distribution
            End If
        End Function

        Public Shared Function Mean(Values() As Double) As Double
            Mean=0
            For Each Value In Values
                Mean+=Value/Values.Length
            Next
            Return Mean
        End Function

        Public Shared Function Median(Values() As Double) As Double
            Dim MedianValues(Values.Count-1) As Double
            Array.Copy(Values,MedianValues,Values.Count)
            Array.Sort(MedianValues)
            If MedianValues.Length Mod 2=0 Then
                Return (MedianValues(MedianValues.Length\2)+MedianValues(MedianValues.Length\2-1))/2
            Else
                Return MedianValues(MedianValues.Length\2)
            End If
        End Function

        Public Shared Function MeanAbsoluteDeviation(Values() As Double) As Double
            Dim ValuesMean As Double=Mean(Values)
            Dim i As Integer
            Dim MAD(Values.Count-1) As Double
            For i=0 to Values.Length-1
                MAD(i)=Math.Abs(Values(i)-ValuesMean)
            Next i
            Return Mean(MAD)
        End Function

        Public Shared Function MedianAbsoluteDeviation(Values() As Double) As Double
            Dim ValuesMedian As Double=Median(Values)
            Dim i As Integer
            Dim MAD(Values.Count-1) As Double
            For i=0 to Values.Length-1
                MAD(i)=Math.Abs(Values(i)-ValuesMedian)
            Next i
            Return Median(MAD)
        End Function

        Public Shared Function RemoveOutliers(Values() As Double) As Double() ' Returns an array which contains the input data from which outliers are removed (NaN)
            Dim ValuesMedian,ValuesMedianAbsoluteDeviation,ValuesMean,ValuesMeanAbsoluteDeviation As Double
            Dim i As Integer
            ValuesMedian=Median(Values)
            ValuesMedianAbsoluteDeviation=MedianAbsoluteDeviation(Values)
            ValuesMean=Mean(Values)
            ValuesMeanAbsoluteDeviation=MeanAbsoluteDeviation(Values)
            For i=0 to Values.Length-1
                If ValuesMedianAbsoluteDeviation=0 Then ' Use Mean Absolute Deviation instead of Median Absolute Deviation to detect outliers
                    If Math.Abs(Values(i)-ValuesMean)>ValuesMeanAbsoluteDeviation*MeanAbsoluteDeviationScaleFactor*ControlLimitRejectionCriterion(DBMParameters.ConfidenceInterval,Values.Length-1) Then ' If value is an outlier
                        Values(i)=Double.NaN ' Exclude outlier
                    End If
                Else ' Use Median Absolute Deviation to detect outliers
                    If Math.Abs(Values(i)-ValuesMedian)>ValuesMedianAbsoluteDeviation*MedianAbsoluteDeviationScaleFactor(Values.Length-1)*ControlLimitRejectionCriterion(DBMParameters.ConfidenceInterval,Values.Length-1) Then ' If value is an outlier
                        Values(i)=Double.NaN ' Exclude outlier
                    End If
                End If
            Next i
            Return Values
        End Function

        Public Shared Function ExponentialMovingAverage(Values() As Double) As Double ' Filter high frequency variation
            Dim Weight,TotalWeight As Double
            ExponentialMovingAverage=0
            Weight=1 ' Initial weight
            TotalWeight=0
            For Each Value In Values ' Least significant value first
                ExponentialMovingAverage+=Value*Weight
                TotalWeight+=Weight
                Weight/=1-2/(Values.Length+1) ' Increase weight for newer values
            Next
            ExponentialMovingAverage/=TotalWeight
            Return ExponentialMovingAverage
        End Function

        Public Shared Function RandomNumber(Min As Integer,Max As Integer) As Integer ' Returns a random number between Min (inclusive) and Max (inclusive)
            Return Random.Next(Min,Max+1)
        End Function

    End Class

End Namespace
