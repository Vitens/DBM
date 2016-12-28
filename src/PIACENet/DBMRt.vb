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

<assembly:System.Reflection.AssemblyTitle("DBMRt")>

Namespace DBMRt

    Public Class DBMRt

        Inherits OSIsoft.PI.ACE.PIACENetClassModule

        Private Calculator As DBMRtCalculator=Nothing

        Public Overrides Sub ACECalculations ' Perform calculations
            If Calculator IsNot Nothing Then ' If module and DBMRtCalculator are initialized
                Calculator.Calculate
            End If
        End Sub

        Protected Overrides Sub InitializePIACEPoints
        End Sub

        Protected Overrides Sub ModuleDependentInitialization
            Calculator=New DBMRtCalculator
        End Sub

        Protected Overrides Sub ModuleDependentTermination
            Calculator=Nothing
        End Sub

    End Class

End Namespace
