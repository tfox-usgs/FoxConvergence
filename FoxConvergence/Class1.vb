Public Class Converge
    Private _High As Double = Double.MaxValue
    Private _Low As Double = Double.MinValue
    Private _CurrentValue As Double = 0
    Private _Tolerance As Double = 0
    Private _Converged As Boolean = False
    Private _IterationCount As Long = 0

    Public Enum Hint As Integer
        TooHigh = 1
        TooLow = -1
        Exact = 0
    End Enum

    Public Sub New(Low As Double, High As Double, Tolerance As Double)
        Try
            If High < Low Then
                Dim Temp As Double = High
                High = Low
                Low = Temp
            End If
            _Low = Low
            _High = High
            _Tolerance = Math.Abs(Tolerance)
            _CurrentValue = (_High + _Low) / 2
            _Converged = _High - _Low <= _Tolerance
        Catch ex As Exception
        End Try
    End Sub
    Public ReadOnly Property IterationCount() As Long
        Get
            Return _IterationCount
        End Get
    End Property
    Public ReadOnly Property HighValue() As Double
        Get
            Return _High
        End Get
    End Property

    Public ReadOnly Property LowValue() As Double
        Get
            Return _Low
        End Get
    End Property

    Public ReadOnly Property Tolerance() As Double
        Get
            Return _Tolerance
        End Get
    End Property

    Public ReadOnly Property CurrentValue() As Double
        Get
            Return _CurrentValue
        End Get
    End Property

    Public ReadOnly Property Converged() As Boolean
        Get
            Return _Converged
        End Get
    End Property

    Public Function NewValue(Hint As Converge.Hint) As Double
        Try
            _IterationCount += 1
            Select Case Hint
                Case Hint.TooHigh
                    _High = _CurrentValue
                Case Hint.TooLow
                    _Low = _CurrentValue
                Case Hint.Exact
                    _Low = _CurrentValue
                    _High = _CurrentValue
            End Select
            _CurrentValue = (_High + _Low) / 2
            _Converged = Math.Abs(_High - _CurrentValue) <= _Tolerance
        Catch ex As Exception
        End Try
        Return _CurrentValue
    End Function

End Class
