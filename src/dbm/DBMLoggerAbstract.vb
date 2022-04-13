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
Imports System.DateTime
Imports System.Diagnostics
Imports System.Environment
Imports Vitens.DynamicBandwidthMonitor.DBMInfo


Namespace Vitens.DynamicBandwidthMonitor


  Public MustInherit Class DBMLoggerAbstract


    Private _staticLogInfo As String


    Public Enum Level

      [Error]
      Warning
      Information
      Debug
      Trace

    End Enum


    Private Function EncloseBrackets(text As String) As String

      If text.Equals(String.Empty) Or text.Contains(" ") Then
        Return "[" & text & "]"
      Else
        Return text
      End If

    End Function


    Public Sub New

      With Process.GetCurrentProcess
        _staticLogInfo = EncloseBrackets(Environment.MachineName) & " " &
          EncloseBrackets(DBMInfo.ProductName) & " " &
          EncloseBrackets(.ProcessName) & " " &
          .Id.ToString & " " &
          EncloseBrackets(Environment.UserName)
      End With

    End Sub


    Public MustOverride Sub Log(level As Level, message As String)


    Private Function FormatLog(level As Level, entity As String,
      message As String) As String

      Dim i As Integer = 1
      Dim frameCount As Integer = (New StackTrace).FrameCount
      Dim caller As StackFrame = New StackFrame(0)
      Dim loggerClass As String =
        caller.GetMethod.DeclaringType.FullName.ToString

      Do While i < frameCount
        caller = New StackFrame(i)
        With caller.GetMethod
          ' Find the first non-constructor (instance, static) method outside of
          ' this class.
          If Not .DeclaringType.FullName.ToString.Equals(loggerClass) And
            Not .Name.ToString.Equals(".ctor") And
            Not .Name.ToString.Equals(".cctor") Then Exit Do
        End With
        i+=1
      Loop

      If message.Contains(NewLine) Then ' Multi-line message
        message = NewLine & "    " & message.Replace(NewLine, NewLine & "    ")
      End If

      With caller.GetMethod
        Return Now.ToString("yyyy-MM-ddTHH:mm:ss.fff") & " " &
          level.ToString & " " &
          _staticLogInfo & " " &
          .DeclaringType.FullName.ToString.Substring(
            .DeclaringType.Namespace.ToString.Length+1) &
          "." & .Name.ToString & " " &
          EncloseBrackets(entity) & " " &
          message
      End With

    End Function


    Public Sub LogError(message As String, Optional entity As String = "")

      ' Error messages.
      ' For errors and exceptions that cannot be handled. These messages
      ' indicate a failure in the current operation or request, not an app-wide
      ' failure.

      Log(Level.Error, FormatLog(Level.Error, entity, message))

    End Sub


    Public Sub LogWarning(message As String, Optional entity As String = "")

      ' Warning messages. Encountered a recoverable error.
      ' For abnormal or unexpected events. Typically includes errors or
      ' conditions that don't cause the app to fail.

      Log(Level.Warning, FormatLog(Level.Warning, entity, message))

    End Sub


    Public Sub LogInformation(message As String, Optional entity As String = "")

      ' Informational messages.
      ' Tracks the general flow of the app. May have long-term value.

      Log(Level.Information, FormatLog(Level.Information, entity, message))

    End Sub


    Public Sub LogDebug(message As String, Optional entity As String = "")

      ' More detailed messages within a method (e.g., Sending email).
      ' For debugging and development.

      Log(Level.Debug, FormatLog(Level.Debug, entity, message))

    End Sub


    Public Sub LogTrace(message As String, Optional entity As String = "")

      ' Data value messages (EmailAddress = john@invalidcompany.com).
      ' Contain the most detailed messages. These messages may contain
      ' sensitive app data.

      Log(Level.Trace, FormatLog(Level.Trace, entity, message))

    End Sub


  End Class


End Namespace
