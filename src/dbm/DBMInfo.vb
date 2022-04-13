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
Imports System.Math
Imports System.Reflection
Imports System.Security.Cryptography.X509Certificates
Imports System.TimeSpan
Imports Vitens.DynamicBandwidthMonitor.DBMStrings
Imports Vitens.DynamicBandwidthMonitor.DBMTests


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMInfo


    Private Shared Function AssemblyLocation As String

      ' Returns assembly.

      Return Assembly.GetExecutingAssembly.Location

    End Function


    Private Shared Function GetFileVersionInfo As FileVersionInfo

      ' Returns FileVersionInfo for assembly.

      Return FileVersionInfo.GetVersionInfo(AssemblyLocation)

    End Function


    Private Shared Function BuildDate As DateTime

      ' Returns the build date based on FileBuildPart.

      Return New DateTime(2000, 1, 1).AddDays(GetFileVersionInfo.FileBuildPart)

    End Function


    Public Shared Function Version As String

      ' Returns a string containing the full version number including, if set,
      ' Git hash. Use Semantic Versioning Specification (SemVer).

      Const GITHASH As String = "#######" ' Updated by CI/CD tools.

      With GetFileVersionInfo
        Return .FileMajorPart.ToString & "." & .FileMinorPart.ToString & ".0+" &
          BuildDate.ToString("yyMMdd") & "." & GITHASH
      End With

    End Function


    Public Shared Function ProductName As String

      ' Returns the name of the product.

      Return GetFileVersionInfo.ProductName

    End Function


    Public Shared Function Application As String

      ' Returns the name and version of the application.

      Return ProductName & " v" & Version

    End Function


    Public Shared Function Product As String

      ' Returns the name of and information about the product.

      With GetFileVersionInfo
        Return Application & NewLine &
          .Comments & NewLine &
          .LegalCopyright
      End With

    End Function


    Public Shared Function LicenseNotice As String

      ' Returns a string containing product name, version number, copyright,
      ' and license notice.

      Return Product & NewLine &
        NewLine &
        "This program is free software: you can redistribute it and/or " &
        "modify it under the terms of the GNU General Public License as " &
        "published by the Free Software Foundation, either version 3 of the " &
        "License, or (at your option) any later version." & NewLine &
        NewLine &
        "This program is distributed in the hope that it will be useful, but " &
        "WITHOUT ANY WARRANTY; without even the implied warranty of " &
        "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU " &
        "General Public License for more details." & NewLine &
        NewLine &
        "You should have received a copy of the GNU General Public License " &
        "along with this program.  " &
        "If not, see <http://www.gnu.org/licenses/>."

    End Function


    Public Shared Function CertificateInfo As String

      ' Returns a string containing the certificate Subject and Issuer, if
      ' available. Else returns 'Unsigned assembly'.

      Try
        With X509Certificate.CreateFromSignedFile(AssemblyLocation)
          Return .Subject & " (" & .Issuer & ")"
        End With
      Catch
        DBM.Logger.LogWarning(UnsignedAssembly & " " & AssemblyLocation)
        Return UnsignedAssembly
      End Try

    End Function


    Public Shared Function TestResults As String

      ' Run unit and integration tests and return test run duration. An
      ' exception occurs if one of the tests fail.

      Dim Timer As DateTime
      Dim UTDurationMs, ITDurationMs, QTDurationMs As Double
      Dim QTResult As String

      Timer = Now
      RunUnitTests
      UTDurationMs = (Now.Ticks-Timer.Ticks)/TicksPerMillisecond

      Timer = Now
      RunIntegrationTests
      ITDurationMs = (Now.Ticks-Timer.Ticks)/TicksPerMillisecond

      Timer = Now
      QTResult = RunQualityTests
      QTDurationMs = (Now.Ticks-Timer.Ticks)/TicksPerMillisecond

      Return "Certificate: " & CertificateInfo & NewLine &
        "Unit tests: " & Round(UTDurationMs).ToString & " ms" & NewLine &
        "Integration tests: " & Round(ITDurationMs).ToString & " ms" & NewLine &
        "Quality tests: " & Round(QTDurationMs).ToString & " ms; " & QTResult

    End Function


  End Class


End Namespace
