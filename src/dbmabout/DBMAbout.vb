Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
Imports System.Security.Cryptography.X509Certificates
Imports Vitens.DynamicBandwidthMonitor.DBMInfo


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMAbout")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMAbout


    ' DBMAbout is a command line utility that can be used to output some DBM
    ' information.


    Public Shared Sub Main

      Console.WriteLine(LicenseNotice)
      Console.WriteLine(NewLine & TestResults)
      Console.WriteLine("Certificate: " & CertificateInfo)

      With X509Certificate2(X509Certificate.CreateFromSignedFile(
        Assembly.GetExecutingAssembly.Location))
        Console.WriteLine("{0}Subject: {1}{0}", Environment.NewLine, .Subject);
        Console.WriteLine("{0}Issuer: {1}{0}", Environment.NewLine, .Issuer);
        Console.WriteLine("{0}Version: {1}{0}", Environment.NewLine, .Version);
        Console.WriteLine("{0}Valid Date: {1}{0}", Environment.NewLine, .NotBefore);
        Console.WriteLine("{0}Expiry Date: {1}{0}", Environment.NewLine, .NotAfter);
        Console.WriteLine("{0}Thumbprint: {1}{0}", Environment.NewLine, .Thumbprint);
        Console.WriteLine("{0}Serial Number: {1}{0}", Environment.NewLine, .SerialNumber);
        Console.WriteLine("{0}Friendly Name: {1}{0}", Environment.NewLine, .PublicKey.Oid.FriendlyName);
        Console.WriteLine("{0}Public Key Format: {1}{0}", Environment.NewLine, .PublicKey.EncodedKeyValue.Format(true));
        Console.WriteLine("{0}Raw Data Length: {1}{0}", Environment.NewLine, .RawData.Length);
        Console.WriteLine("{0}Certificate to string: {1}{0}", Environment.NewLine, .ToString(true));
        Console.WriteLine("{0}Certificate to XML String: {1}{0}", Environment.NewLine, .PublicKey.Key.ToXmlString(false));
      End With

    End Sub


  End Class


End Namespace
