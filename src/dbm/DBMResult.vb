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

    Public Class DBMResult

        Public Factor,OriginalFactor,AbsoluteErrors(),RelativeErrors() As Double
        Public DBMPrediction As DBMPrediction
        Public AbsErrorStats,RelErrorStats As New DBMStatistics
        Public SuppressedBy As DBMPointDriver

        Public Sub New
            ReDim AbsoluteErrors(DBMParameters.CorrelationPreviousPeriods)
            ReDim RelativeErrors(DBMParameters.CorrelationPreviousPeriods)
        End Sub

        Public Sub Calculate(Index As Integer,MeasValueEMA As Double,PredValueEMA As Double,LowContrLimitEMA As Double,UppContrLimitEMA As Double) ' Calculates and stores prediction errors and initial results
            AbsoluteErrors(Index)=PredValueEMA-MeasValueEMA ' Absolute prediction error (for prediction error correlation calculations)
            RelativeErrors(Index)=PredValueEMA/MeasValueEMA-1 ' Relative prediction error (for prediction error correlation calculations)
            If DBMPrediction Is Nothing Then ' Store initial (no time offset because of prediction error correlation calculations) results
                If MeasValueEMA<LowContrLimitEMA Then ' Lower control limit exceeded
                    Factor=(PredValueEMA-MeasValueEMA)/(LowContrLimitEMA-PredValueEMA)
                ElseIf MeasValueEMA>UppContrLimitEMA Then ' Upper control limit exceeded
                    Factor=(MeasValueEMA-PredValueEMA)/(UppContrLimitEMA-PredValueEMA)
                End If
                OriginalFactor=Factor ' Store original factor before possible suppression
                DBMPrediction=New DBMPrediction(MeasValueEMA,PredValueEMA,LowContrLimitEMA,UppContrLimitEMA)
            End If
        End Sub

    End Class

End Namespace
