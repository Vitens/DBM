Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMConstants

    Public Const CalculationInterval As Integer=            300 ' seconds; 300 = 5 minutes
    Public Const ComparePatterns As Integer=                12 ' weeks
    Public Const EMAPreviousPeriods As Integer=             CInt(0.5*3600/CalculationInterval) ' previous periods; 6 = 35 minutes, current value inclusive
    Public Const ConfidenceInterval As Double=              0.99 ' confidence interval for outlier detection
    Public Const CorrelationPreviousPeriods As Integer=     CInt(2*3600/CalculationInterval-1) ' previous periods; 23 = 2 hours, current value inclusive
    Public Const CorrelationThreshold As Double=            0.83666 ' absolute correlation lower limit for detecting (anti)correlation

End Class
