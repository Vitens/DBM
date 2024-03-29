' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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
Imports OSIsoft.AF.Diagnostics


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMLoggerAFTrace
    Inherits DBMLoggerAbstract


    Public Overrides Sub Log(level As Level, message As String)

      Select Case level
        Case Level.Error
          If AFTrace.IsTraced(AFTraceSwitchLevel.Error) Then
            AFTrace.TraceError(message)
          End If
        Case Level.Warning
          If AFTrace.IsTraced(AFTraceSwitchLevel.Warning) Then
            AFTrace.TraceWarning(message)
          End If
        Case Level.Information
          If AFTrace.IsTraced(AFTraceSwitchLevel.Information) Then
            AFTrace.TraceInformation(message)
          End If
        Case Level.Debug
          If AFTrace.IsTraced(AFTraceSwitchLevel.Detail) Then
            AFTrace.TraceDetail(message)
          End If
        Case Level.Trace
          If AFTrace.IsTraced(AFTraceSwitchLevel.Data) Then
            AFTrace.TraceData(message)
          End If
      End Select

    End Sub


  End Class


End Namespace
