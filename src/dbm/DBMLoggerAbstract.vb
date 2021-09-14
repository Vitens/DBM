Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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


  Public MustInherit Class DBMLoggerAbstract


    Public Enum Level

      [Error]
      Warning
      Information
      Debug
      Trace

    End Enum


    Public MustOverride Sub Log(Level As Level, Message As String)


    Public Sub LogError(Message As String)

      ' Error messages.
      ' For errors and exceptions that cannot be handled. These messages
      ' indicate a failure in the current operation or request, not an app-wide
      ' failure.

      Log(Level.Error, Message)

    End Sub


    Public Sub LogWarning(Message As String)

      ' Warning messages. Encountered a recoverable error.
      ' For abnormal or unexpected events. Typically includes errors or
      ' conditions that don't cause the app to fail.

      Log(Level.Warning, Message)

    End Sub


    Public Sub LogInformation(Message As String)

      ' Informational messages.
      ' Tracks the general flow of the app. May have long-term value.

      Log(Level.Information, Message)

    End Sub


    Public Sub LogDebug(Message As String)

      ' More detailed messages within a method (e.g., Sending email).
      ' For debugging and development.

      Log(Level.Debug, Message)

    End Sub


    Public Sub LogTrace(Message As String)

      ' Data value messages (EmailAddress = john@invalidcompany.com).
      ' Contain the most detailed messages. These messages may contain
      ' sensitive app data.

      Log(Level.Trace, Message)

    End Sub


  End Class


End Namespace
