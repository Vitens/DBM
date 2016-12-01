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

Public Class DBMFunctions

    Public Shared Function DBMVersion As String
        DBMVersion=System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileDescription
        DBMVersion=DBMVersion & " v" & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileVersion & vbCrLf
        DBMVersion=DBMVersion & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).ProductName & vbCrLf
        DBMVersion=DBMVersion & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Comments & vbCrLf & vbCrLf
        DBMVersion=DBMVersion & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).LegalCopyright
        Return DBMVersion
    End Function

    Public Shared Function ArrayRotateLeft(ByVal Data() As Double) As Double() ' Rotate array left, first item becomes last
        Array.Reverse(Data) ' ABCDE -> EDCBA
        Array.Reverse(Data,0,Data.Length-1) ' EDCBA -> BCDEA
        Return Data
    End Function

End Class
