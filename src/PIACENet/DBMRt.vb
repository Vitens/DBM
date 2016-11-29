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
<assembly:AssemblyTitle("DBMRt")>
<assembly:AssemblyVersion("2.6.0.*")>
<assembly:AssemblyProduct("Dynamic Bandwidth Monitor Real-time")>
<assembly:AssemblyDescription("Leak detection method implemented in a real-time data historian")>
<assembly:AssemblyCopyright("Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.")>
<assembly:AssemblyCompany("Vitens N.V.")>

Public Class DBMRt

    Inherits OSIsoft.PI.ACE.PIACENetClassModule

    Public Overrides Sub ACECalculations()
    End Sub

    Protected Overrides Sub InitializePIACEPoints()
    End Sub

    Protected Overrides Sub ModuleDependentInitialization()
        Dim DBMRtCalculator As New DBMRtCalculator(True)
        Do While True
            DBMRtCalculator.Calculate
            Threading.Thread.Sleep(DBMRtConstants.CalculationDelay*1000)
        Loop
    End Sub

    Protected Overrides Sub ModuleDependentTermination()
    End Sub

End Class
