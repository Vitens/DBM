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
Imports System.Globalization
Imports System.Math
Imports System.TimeSpan
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMDate


    ' Contains date and time functions.


    Public Shared Function PreviousInterval(timestamp As DateTime) As DateTime

      ' Returns a DateTime for the passed timestamp aligned on the previous
      ' interval.

      Return timestamp.AddTicks(
        -timestamp.Ticks Mod CalculationInterval*TicksPerSecond)

    End Function


    Public Shared Function NextInterval(timestamp As DateTime) As DateTime

      Return PreviousInterval(timestamp.AddSeconds(CalculationInterval))

    End Function


    Public Shared Function IsOnInterval(timestamp As DateTime) As Boolean

      ' Returns true if the end timestamp is exactly on an interval boundary.

      Return timestamp.Equals(PreviousInterval(timestamp))

    End Function


    Public Shared Function IntervalSeconds(numberOfValues As Integer,
      durationSeconds As Double) As Double

      ' Number of values desired. If =0, all intervals will be returned. If >0,
      ' that number of values will be returned. If <0, the negative value minus
      ' 1 number of values will be returned (f.ex. -25 over a 24 hour period
      ' will return an hourly value). Duration is in seconds. The first value
      ' returned will then be the first interval in the time range, and the last
      ' value returned will be the final calculation interval (f.ex. requesting
      ' 6 values for a duration of 60 minutes with a calculation interval of 5
      ' minutes will return an interval of 660 seconds which will then return
      ' values at :00, :11, :22, :33, :44, and :55).

      If numberOfValues < 0 Then numberOfValues = -numberOfValues-1
      If numberOfValues = 1 Then
        Return durationSeconds ' Return a single value
      Else
        Return Max(1, (durationSeconds/CalculationInterval-1)/
          (numberOfValues-1))*CalculationInterval ' Required interval
      End If

    End Function


    Public Shared Function Computus(year As Integer) As DateTime

      ' Computus (Latin for computation) is the calculation of the date of
      ' Easter in the Christian calendar. The computus is used to set the time
      ' for ecclesiastical purposes, in particular to calculate the date of
      ' Easter.

      Dim H, i, L As Integer

      H = (year\100-year\400-(8*(year\100)+13)\25+19*(year Mod 19)+15) Mod 30
      i = H-H\28-(H\28)*(29\H+1)*(21-year Mod 19)\11
      L = i-(year+year\4+i+2-year\100+year\400) Mod 7

      Return New DateTime(year, 3+(L+40)\44, L+28-31*((172+L)\176))

    End Function


    Public Shared Function IsHoliday(timestamp As DateTime,
      culture As CultureInfo) As Boolean

      ' Returns True if the passed date is a holiday.

      Dim DaysSinceEaster As Integer =
        timestamp.Subtract(Computus(timestamp.Year)).Days

      With timestamp
        ' For any culture, consider the following days a holiday:
        '  New Year's Day, Easter, Ascension Day, Pentecost,
        '  Christmas Day, and New Year's Eve.
        IsHoliday = (.Month = 1 And .Day = 1) Or
          {0, 39, 49}.Contains(DaysSinceEaster) Or
          (.Month = 12 And {25, 31}.Contains(.Day))
        If culture.Name.Equals("nl-NL") Then
          ' For the Netherlands, consider the following days a holiday:
          '  2nd day of Easter, Royal day, Liberation Day (lustrum),
          '  2nd day of Pentecost, and Boxing Day
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


    Public Shared Function DaysSinceSunday(timestamp As DateTime) As Integer

      ' Return the number of days between the passed timestamp and the
      ' preceding Sunday.

      Return timestamp.DayOfWeek ' 0=Sunday, 6=Saturday

    End Function


    Public Shared Function PreviousSunday(timestamp As DateTime) As DateTime

      ' Returns a DateTime for the Sunday preceding the timestamp passed.

      Return timestamp.AddDays(-DaysSinceSunday(timestamp))

    End Function


    Public Shared Function OffsetHoliday(timestamp As DateTime,
      culture As CultureInfo) As DateTime

      If UseSundayForHolidays And IsHoliday(timestamp, culture) Then
        Return PreviousSunday(timestamp) ' Offset holidays
      Else
        Return timestamp
      End If

    End Function


    Public Shared Function DataPreparationTimestamp(
      startTimestamp As DateTime) As DateTime

      startTimestamp = PreviousInterval(startTimestamp.AddDays(
        -7*ComparePatterns).AddSeconds(
        -(EMAPreviousPeriods+CorrelationPreviousPeriods)*CalculationInterval))
      If UseSundayForHolidays Then startTimestamp =
        PreviousSunday(startTimestamp)

      Return startTimestamp

    End Function


    Public Shared Function EMATimeOffset(count As Integer) As Integer

      ' The use of an exponential moving average (EMA) time shifts the resulting
      ' calculated data. To compensate for this, an offset should be applied
      ' based on exponentially increasing weighting factors. The returned value
      ' is in seconds and is only shifted in whole intervals.

      ' Floor(n/2.91136) is a fast approximation of Round(n-EMA(1..n)).
      Return CInt(-Floor(count/2.91136)*CalculationInterval)

    End Function


  End Class


End Namespace
