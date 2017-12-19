Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014, 2015, 2016, 2017, 2018  J.H. Fitié, Vitens N.V.
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
Imports System.Collections.Generic
Imports System.Double
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMPointDriverCSV")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPointDriver
    Inherits DBMPointDriverAbstract


    ' Description: Driver for CSV files (timestamp,value).
    ' Identifier (Point): String (CSV filename)
    ' Remarks: Data interval must be the same as the CalculationInterval
    '          parameter.


    Private Values As New Dictionary(Of DateTime, Double)
    Private Shared SplitChars As Char() = {","c, "	"c} ' Comma and tab


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Private Function IsEquidistantList(l As IList(Of DateTime)) As Boolean

      ' Checks whether a list of timestamps has the same delta between
      ' consecutive DateTimes. To allow for irregularities during DST
      ' transitions, only the delta of the first and last pair is checked
      ' agains the average delta for the whole list.

      Dim dt_start As Double = l(1).Subtract(l(0)).TotalSeconds
      Dim dt_end As Double = l(l.Count-1).Subtract(l(l.Count-2)).TotalSeconds
      Dim dt_avg As Double = _
        l(l.Count-1).Subtract(l(0)).TotalSeconds/(l.Count-1)

      Return dt_start = dt_end And dt_start = dt_avg

    End Function


    Private Sub LoadBinaryData(SerializedCSVFileName As String)

      ' Load Values dictionary from a serialized Double array for a set of
      ' equidistant timestamps defined by a start timestamp and a delta.

      Dim first_ts, Timestamp As DateTime
      Dim dt, i As Integer
      Dim varray As Double() = Nothing
      Dim formatter As BinaryFormatter = new BinaryFormatter()

      Using fs As Stream = new FileStream(SerializedCSVFileName, _
        FileMode.Open, FileAccess.Read, FileShare.Read)
        first_ts = DirectCast(formatter.Deserialize(fs), DateTime)
        dt = DirectCast(formatter.Deserialize(fs), Integer)
        varray = DirectCast(formatter.Deserialize(fs), Double())
      End Using
      For i = 0 To varray.Length-1
        Timestamp = first_ts.AddSeconds(i*dt)
        If Not Double.IsNaN(varray(i)) And _
          Not Values.ContainsKey(Timestamp) Then
          Values.Add(Timestamp, varray(i))
        End If
      Next

    End Sub


    Private Sub SaveBinaryData(SerializedCSVFileName As String, _
      tslist As IList(Of DateTime))

      ' Convert the Values dictionary into a Double array and
      ' serialize this array.

      Dim first_ts, Timestamp As DateTime
      Dim dt, i As Integer
      Dim varray As Double() = New Double(tslist.Count-1){}
      Dim formatter As BinaryFormatter = new BinaryFormatter()

      first_ts = tslist(0)
      dt = CInt(tslist(1).Subtract(tslist(0)).TotalSeconds)
      For i = 0 To tslist.Count-1
        Timestamp = first_ts.AddSeconds(i*dt)
        If Values.ContainsKey(Timestamp) Then
          varray(i) = Values.Item(Timestamp)
        Else
          varray(i) = NaN
        End If
      Next
      Using fs As Stream = new FileStream(SerializedCSVFileName, _
        FileMode.Create, FileAccess.Write, FileShare.None)
        formatter.Serialize(fs, first_ts)
        formatter.Serialize(fs, dt)
        formatter.Serialize(fs, varray)
      End Using

    End Sub


    Public Overrides Sub PrepareData(StartTimestamp As DateTime, _
      EndTimestamp As DateTime)

      ' Retrieves information from a CSV file and stores this in the Values
      ' dictionary. Passed timestamps are ignored and all data in the
      ' (serialized) CSV is loaded into memory.

      Dim CSVFileName, SerializedCSVFileName, Substrings() As String
      Dim TimestampList As List(Of DateTime)
      Dim Timestamp As DateTime
      Dim Value As Double

      Values.Clear
      CSVFileName = DirectCast(Point, String)
      If File.Exists(CSVFileName) Then
        SerializedCSVFileName = CSVFileName & ".bin"
        If File.Exists(SerializedCSVFileName) And _
          File.GetLastWriteTime(CSVFileName) < _
          File.GetLastWriteTime(SerializedCSVFileName) Then
          LoadBinaryData(SerializedCSVFileName)
        Else
          TimestampList = New List(Of DateTime)
          Using StreamReader As New StreamReader(DirectCast(Point, String))
            Do While Not StreamReader.EndOfStream
              ' Comma and tab delimiters; split timestamp and value
              Substrings = StreamReader.ReadLine.Split(SplitChars, 2)
              If Substrings.Length = 2 Then
                If DateTime.TryParse(Substrings(0), Timestamp) Then
                  If Double.TryParse(Substrings(1), Value) Then
                    TimestampList.Add(Timestamp)
                    If Not Values.ContainsKey(Timestamp) Then
                      Values.Add(Timestamp, Value) ' Add data to dictionary
                    End If
                  End If
                End If
              End If
            Loop
          End Using ' Close CSV file
          If IsEquidistantList(TimestampList) Then
            ' In case of equidistant timestamps: serialize the Values
            ' dictionary to speed up the next run on the same data.
            SaveBinaryData(SerializedCSVFileName, TimestampList)
          End If
        End If
      Else
        ' If Point does not represent a valid, existing file then throw a
        ' File Not Found Exception.
        Throw New FileNotFoundException(DirectCast(Point, String))
      End If

    End Sub


    Public Overrides Function GetData(Timestamp As DateTime) As Double

      ' GetData retrieves data from the Values dictionary. Non existing
      ' timestamps return NaN.

      If Values.Count = 0 Then PrepareData(DateTime.MinValue, DateTime.MaxValue)

      ' Look up data from memory
      GetData = Nothing
      If Values.TryGetValue(Timestamp, GetData) Then ' In cache
        Return GetData ' Return value from cache
      Else
        Return NaN ' No data in memory for timestamp, return Not a Number.
      End If

    End Function


  End Class


End Namespace
