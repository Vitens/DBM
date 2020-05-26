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
Imports System.Diagnostics
Imports System.Environment
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMTests


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMInfo


    Private Shared Function GetFileVersionInfo As FileVersionInfo

      ' Returns FileVersionInfo for assembly.

      Return FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.
        GetExecutingAssembly.Location)

    End Function


    Private Shared Function BuildDate As DateTime

      ' Returns the build date based on FileBuildPart.

      Return New DateTime(2000, 1, 1).AddDays(GetFileVersionInfo.FileBuildPart)

    End Function


    Public Shared Function Version As String

      ' Returns a string containing the full version number including, if set,
      ' Git hash. Use Semantic Versioning Specification (SemVer).

      Const GITHASH As String = "" ' Updated by CI build script (appveyor.yml).

      With GetFileVersionInfo
        Return .FileMajorPart.ToString & "." & .FileMinorPart.ToString & ".0+" &
          BuildDate.ToString("yyMMdd") & If(GITHASH = "", "", "." & GITHASH)
      End With

    End Function


    Public Shared Function LicenseNotice As String

      ' Returns a string containing product name, version number, copyright and
      ' license notice.

      With GetFileVersionInfo
        Return .ProductName & " v" & Version & NewLine &
          .Comments & NewLine &
          .LegalCopyright & NewLine &
          NewLine &
          "This program is free software: you can redistribute it and/or " &
          "modify it under the terms of the GNU General Public License as " &
          "published by the Free Software Foundation, either version 3 of " &
          "the License, or (at your option) any later version." & NewLine &
          NewLine &
          "This program is distributed in the hope that it will be useful, " &
          "but WITHOUT ANY WARRANTY; without even the implied warranty of " &
          "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the " &
          "GNU General Public License for more details." & NewLine &
          NewLine &
          "You should have received a copy of the GNU General Public " &
          "License along with this program.  " &
          "If not, see <http://www.gnu.org/licenses/>." & NewLine
      End With

    End Function


    Public Shared Function TestResults As String

      ' Returns a string containing test results and performance index.

      Return " * Unit tests " &
        If(UnitTestsPassed, "PASSED", "FAILED") & "." & NewLine &
        " * Integration tests " &
        If(IntegrationTestsPassed, "PASSED", "FAILED") & "." & NewLine &
        " * Performance index " &
        Round(PerformanceIndex, 1).ToString & "." & NewLine

    End Function


  End Class


End Namespace
