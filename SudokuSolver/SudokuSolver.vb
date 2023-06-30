Public Class SudokuSolver
    Public Shared Function IsSafe(Grid As Integer()(),
                                  Row As Integer,
                                  Col As Integer,
                                  K As Integer) As Boolean

        'Check Row if Safe
        Dim SafeRow As Boolean = Not Grid(Row).Any(Function(X) X = K)

        'Check Column if Safe
        Dim SafeCol As Boolean = Not Enumerable.Range(0, 9).Any(Function(I) Grid(I)(Col) = K)

        'Check Sub Grid if Safe
        Dim SubGridRowIndex As Integer = Math.Floor(Row / 3) * 3
        Dim SubGridColIndex As Integer = Math.Floor(Col / 3) * 3

        Dim SafeSG As Boolean = Not Enumerable.Range(SubGridRowIndex, 3) _
                                .SelectMany(Function(R) Enumerable.Range(SubGridColIndex, 3) _
                                .Select(Function(C) Grid(R)(C))).Any(Function(X) X = K)

        Return SafeRow AndAlso SafeCol AndAlso SafeSG
    End Function
    Shared Function SolveSudoku(Grid As Integer()(),
                                R As Integer,
                                C As Integer) As Boolean

        If R = 9 Then
            Return True
        ElseIf C = 9 Then
            Return SolveSudoku(Grid, R + 1, 0)
        ElseIf Grid(R)(C) <> 0 Then
            Return SolveSudoku(Grid, R, C + 1)
        Else
            'Check Possible combination
            For K As Integer = 1 To 9
                If IsSafe(Grid, R, C, K) Then
                    Grid(R)(C) = K

                    If SolveSudoku(Grid, R, C + 1) Then
                        Return True
                    End If

                    Grid(R)(C) = 0
                End If
            Next

            Return False
        End If
    End Function
End Class
