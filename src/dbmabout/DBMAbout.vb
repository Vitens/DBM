Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
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


Imports System
Imports System.Environment
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMInfo
Imports Vitens.DynamicBandwidthMonitor.DBMTests


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMAbout")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMAbout


    ' DBMAbout is a command line utility that can be used to output some DBM
    ' information.


    Public Shared Sub Main

      Console.Write(LicenseNotice)
      Console.WriteLine(NewLine & "Performance Index " &
        Round(PerformanceIndex, 1).ToString)

    End Sub


  End Class


End Namespace
