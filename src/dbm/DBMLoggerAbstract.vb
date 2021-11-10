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
Imports System.DateTime
Imports System.Diagnostics
Imports System.Environment
Imports Vitens.DynamicBandwidthMonitor.DBMInfo


Namespace Vitens.DynamicBandwidthMonitor


  Public MustInherit Class DBMLoggerAbstract


    Private StaticLogInfo As String


    Public Enum Level

      [Error]
      Warning
      Information
      Debug
      Trace

    End Enum


    Private Function EncloseBrackets(Text As String) As String

      If Text.Equals(String.Empty) Or Text.Contains(" ") Then
        Return "[" & Text & "]"
      Else
        Return Text
      End If

    End Function


    Public Sub New

      With Process.GetCurrentProcess
        StaticLogInfo = EncloseBrackets(Environment.MachineName) & " " &
          EncloseBrackets(DBMInfo.ProductName) & " " &
          EncloseBrackets(.ProcessName) & " " &
          .Id.ToString & " " &
          EncloseBrackets(Environment.UserName)
      End With

    End Sub


    Public MustOverride Sub Log(Level As Level, Message As String)


    Private Function FormatLog(Level As Level, Entity As String,
      Message As String) As String

      Dim i As Integer = 1
      Dim FrameCount As Integer = (New StackTrace).FrameCount
      Dim Caller As StackFrame = New StackFrame(0)
      Dim LoggerClass As String =
        Caller.GetMethod.DeclaringType.FullName.ToString

      Do While i < FrameCount
        Caller = New StackFrame(i)
        With Caller.GetMethod
          ' Find the first non-constructor (instance, static) method outside of
          ' this class.
          If Not .DeclaringType.FullName.ToString.Equals(LoggerClass) And
            Not .Name.ToString.Equals(".ctor") And
            Not .Name.ToString.Equals(".cctor") Then Exit Do
        End With
        i+=1
      Loop

      If Message.Contains(NewLine) Then ' Multi-line message
        Message = NewLine & "    " & Message.Replace(NewLine, NewLine & "    ")
      End If

      With Caller.GetMethod
        Return Now.ToString("yyyy-MM-ddTHH:mm:ss.fff") & " " &
          Level.ToString & " " &
          StaticLogInfo & " " &
          .DeclaringType.FullName.ToString & "." & .Name.ToString & " " &
          EncloseBrackets(Entity) & " " &
          Message
      End With

    End Function


    Public Sub LogError(Message As String, Optional Entity As String = "")

      ' Error messages.
      ' For errors and exceptions that cannot be handled. These messages
      ' indicate a failure in the current operation or request, not an app-wide
      ' failure.

      Log(Level.Error, FormatLog(Level.Error, Entity, Message))

    End Sub


    Public Sub LogWarning(Message As String, Optional Entity As String = "")

      ' Warning messages. Encountered a recoverable error.
      ' For abnormal or unexpected events. Typically includes errors or
      ' conditions that don't cause the app to fail.

      Log(Level.Warning, FormatLog(Level.Warning, Entity, Message))

    End Sub


    Public Sub LogInformation(Message As String, Optional Entity As String = "")

      ' Informational messages.
      ' Tracks the general flow of the app. May have long-term value.

      Log(Level.Information, FormatLog(Level.Information, Entity, Message))

    End Sub


    Public Sub LogDebug(Message As String, Optional Entity As String = "")

      ' More detailed messages within a method (e.g., Sending email).
      ' For debugging and development.

      Log(Level.Debug, FormatLog(Level.Debug, Entity, Message))

    End Sub


    Public Sub LogTrace(Message As String, Optional Entity As String = "")

      ' Data value messages (EmailAddress = john@invalidcompany.com).
      ' Contain the most detailed messages. These messages may contain
      ' sensitive app data.

      Log(Level.Trace, FormatLog(Level.Trace, Entity, Message))

    End Sub


  End Class


End Namespace
