@echo off

rem Dynamic Bandwidth Monitor
rem Leak detection method implemented in a real-time data historian
rem Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
rem
rem This file is part of DBM.
rem
rem DBM is free software: you can redistribute it and/or modify
rem it under the terms of the GNU General Public License as published by
rem the Free Software Foundation, either version 3 of the License, or
rem (at your option) any later version.
rem
rem DBM is distributed in the hope that it will be useful,
rem but WITHOUT ANY WARRANTY; without even the implied warranty of
rem MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
rem GNU General Public License for more details.
rem
rem You should have received a copy of the GNU General Public License
rem along with DBM.  If not, see <http://www.gnu.org/licenses/>.

cd /d %~dp0

rem Sign
powershell -Command "Set-AuthenticodeSignature -Certificate (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('%2', '%3')) -TimestampServer http://timestamp.digicert.com -HashAlgorithm SHA256 -FilePath '%1'" > NUL
powershell -Command "'🔑 ' + (Split-Path '%1' -Leaf) + ' signed by ' + (Get-AuthenticodeSignature -FilePath '%1').SignerCertificate.Subject"
