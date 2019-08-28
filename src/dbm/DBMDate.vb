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
Imports System.Globalization
Imports System.TimeSpan
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMDate


    ' Contains date and time functions.


    Public Shared Function AlignTimestamp(Timestamp As DateTime,
      Seconds As Integer) As DateTime

      ' Returns a DateTime for the passed timestamp aligned on the
      ' previous interval (in seconds).

      Return New DateTime(Timestamp.Ticks-Timestamp.Ticks Mod
        Seconds*TicksPerSecond, Timestamp.Kind)

    End Function


    Public Shared Function NextInterval(Timestamp As DateTime,
      Optional Intervals As Integer = 1) As DateTime

      Return AlignTimestamp(Timestamp, CalculationInterval).
        AddSeconds(Intervals*CalculationInterval)

    End Function


    Public Shared Function DaysSinceSunday(Timestamp As DateTime) As Integer

      ' Return the number of days between the passed timestamp and the
      ' preceding Sunday.

      Return Timestamp.DayOfWeek ' 0=Sunday, 6=Saturday

    End Function


    Public Shared Function PreviousSunday(Timestamp As DateTime) As DateTime

      ' Returns a DateTime for the Sunday preceding the timestamp passed.

      Return Timestamp.AddDays(-DaysSinceSunday(Timestamp))

    End Function


    Public Shared Function Computus(Year As Integer) As DateTime

      ' Computus (Latin for computation) is the calculation of the date of
      ' Easter in the Christian calendar. The computus is used to set the time
      ' for ecclesiastical purposes, in particular to calculate the date of
      ' Easter.

      Dim H, i, L As Integer

      H = (Year\100-Year\400-(8*(Year\100)+13)\25+19*(Year Mod 19)+15) Mod 30
      i = H-H\28-(H\28)*(29\H+1)*(21-Year Mod 19)\11
      L = i-(Year+Year\4+i+2-Year\100+Year\400) Mod 7

      Return New DateTime(Year, 3+(L+40)\44, L+28-31*((172+L)\176))

    End Function


    Public Shared Function IsHoliday(Timestamp As DateTime,
      Culture As CultureInfo) As Boolean

      ' Returns True if the passed date is a holiday.

      Dim DaysSinceEaster As Integer =
        Timestamp.Subtract(Computus(Timestamp.Year)).Days

      With Timestamp
        ' For any culture, consider the following days a holiday:
        ' New Year's Day, Easter, Ascension Day, Pentecost,
        ' Christmas Day, and New Year's Eve.
        IsHoliday = (.Month = 1 And .Day = 1) Or
          {0, 39, 49}.Contains(DaysSinceEaster) Or
          (.Month = 12 And {25, 31}.Contains(.Day))
        If Culture.Name.Equals("nl-NL") Then
          ' For the Netherlands, consider the following days a holiday:
          ' 2nd day of Easter, Royal day, Liberation Day (lustrum),
          ' 2nd day of Pentecost, and Boxing Day
          IsHoliday = IsHoliday Or
            {1, 50}.Contains(DaysSinceEaster) Or
            (.Year >= 1980 And .Year < 2014 And .Month = 4 And .Day = 30) Or
            (.Year >= 2014 And .Month = 4 And .Day = 27) Or
            (.Year Mod 5 = 0 And .Month = 5 And .Day = 5) Or
            (.Month = 12 And .Day = 26)
        End If
      End With

      Return IsHoliday

    End Function


    Public Shared Function OffsetDate(Timestamp As DateTime,
      Culture As CultureInfo) As DateTime

      If UseSundayForHolidays And IsHoliday(Timestamp, Culture) Then
        Return PreviousSunday(Timestamp) ' Offset holidays
      Else
        Return Timestamp
      End If

    End Function


  End Class


End Namespace
