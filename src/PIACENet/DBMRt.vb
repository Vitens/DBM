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


Imports OSIsoft.PI.ACE


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMRt")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMRt


    ' DBMRt is a real-time calculation module for OSIsoft PI ACE (Advanced
    ' Computing Engine). It is non-standard in that it does not use the
    ' Module Database for configuration, but instead dynamically searches for
    ' relevant PI tags and performs calculations on these when needed. Note
    ' that there is no support for manual or automatic recalculation.
    ' Suggested scheduling: type: clock; period: 60 seconds.


    Inherits PIACENetClassModule ' PI ACE specific


    Private Calculator As DBMRtCalculator


    Public Overrides Sub ACECalculations
      ' This method is called by ACE Scheduler to execute calculations.
      ' Start calculation only if module and DBMRtCalculator are initialized.
      If Calculator IsNot Nothing Then
        Calculator.Calculate
      End If
    End Sub


    Protected Overrides Sub InitializePIACEPoints
      ' Initializes all PIACEPoint objects.
      ' Unused, but must be overridden.
    End Sub


    Protected Overrides Sub ModuleDependentInitialization
      ' Any custom initialization code can be implemented here. Any variable
      ' that does not change should be initialized here.
      ' On module initialization, create a new DBMRtCalculator object.
      ' Creating the object automatically adds tags from the default PI server.
      Calculator = New DBMRtCalculator
    End Sub


    Protected Overrides Sub ModuleDependentTermination
      ' Contains code to dispose any user-defined objects.
      ' On module termination, mark the calculator object for
      ' garbage collection.
      Calculator = Nothing
    End Sub


  End Class


End Namespace
