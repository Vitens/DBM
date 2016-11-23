Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMFunctions

    Public Shared Function DBMVersion As String
        DBMVersion=System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileDescription
        DBMVersion=DBMVersion & " v" & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileVersion
        DBMVersion=DBMVersion & vbCrLf & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).ProductName
        DBMVersion=DBMVersion & vbCrLf & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).LegalCopyright
        Return DBMVersion
    End Function

    Public Shared Function ArrayRotateLeft(ByVal Data() As Double) As Double() ' Rotate array left, first item becomes last
        Array.Reverse(Data) ' ABCDE -> EDCBA
        Array.Reverse(Data,0,Data.Length-1) ' EDCBA -> BCDEA
        Return Data
    End Function

End Class
