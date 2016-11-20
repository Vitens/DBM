Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Structure DBMResult

    Public Factor,CurrValue,PredValue,LowContrLimit,UppContrLimit As Double

End Structure
