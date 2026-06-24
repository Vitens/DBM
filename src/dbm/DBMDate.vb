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

      ' Returns the timestamp offset for a finite Exponential Moving Average.
      '
      ' A finite EMA of length N uses exponentially decaying weights:
      '   w(k) ∝ r^k
      '
      ' where k = 0 is the newest sample and:
      '   r = 1 - α = (N - 1)/(N + 1)
      '
      ' for the standard EMA definition α = 2/(N+1).
      '
      ' The effective timestamp of the EMA is the center of mass of the
      ' normalized weights:
      '   μ(N) = Σ(k·w(k)) / Σ(w(k))
      '
      ' As N → ∞ (for the finite-length EMA family), this converges to:
      '   lim[N→∞] μ(N)/N
      '     = 1/2 − 1/(e² − 1)
      '     = 0.3434823572503343...
      '
      ' Therefore the EMA effectively represents data approximately 34.35% of
      ' the sample count back from the newest sample.
      '
      ' The returned value is negative because the effective observation time
      ' lies in the past relative to the most recent sample.
      '
      ' Proof (sketch):
      '   μ(N) = Σ(k·r^k) / Σ(r^k)
      '
      ' with:
      '   r = (N - 1)/(N + 1)
      '
      ' Using closed-form sums:
      '   Σ(r^k) = (1 - r^N)/(1 - r)
      '   Σ(k·r^k) = [r - N·r^N + (N-1)·r^(N+1)]/(1-r)^2
      '
      ' Taking the limit N → ∞ gives:
      '   r^N = (1 - 2/(N+1))^N → e^-2
      '
      ' and therefore:
      '   μ(N)/N → 1/2 − 1/(e² − 1)

      Const EMACenterOfMassFraction As Double = 0.343482
      Return CInt(-Floor(count * EMACenterOfMassFraction) * CalculationInterval)

    End Function


  End Class


End Namespace
