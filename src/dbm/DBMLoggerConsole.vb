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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMLoggerConsole
    Inherits DBMLoggerAbstract


    Public Overrides Sub Log(level As Level, message As String)

      Select Case level
        Case Level.Error
          Console.ForegroundColor = ConsoleColor.Red
        Case Level.Warning
          Console.ForegroundColor = ConsoleColor.Yellow
        Case Level.Information
          Console.ForegroundColor = ConsoleColor.White
        Case Level.Debug
          Console.ForegroundColor = ConsoleColor.Cyan
        Case Level.Trace
          Console.ForegroundColor = ConsoleColor.DarkGray
      End Select
      Console.WriteLine(message)
      Console.ResetColor

    End Sub


  End Class


End Namespace
