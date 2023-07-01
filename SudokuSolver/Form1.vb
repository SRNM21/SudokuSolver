Imports System.Data.Common
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar

Public Class Form1
    Private RowsBtn As New List(Of List(Of Windows.Forms.Button))
    Private CurrBtn As Windows.Forms.Button
    Private GridValue()() As Integer = {
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0}}
    Private Flags()() As Integer = {
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0},
        New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0}}
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Add all Child (Buttons) of t   he Grid Panel into List
        For R As Integer = 0 To 8
            Dim ColsBtn As New List(Of Windows.Forms.Button)

            For C As Integer = 0 To 8
                For Each Cell As Windows.Forms.Button In GridPanel.Controls
                    If Cell.Name = $"GridBtn{R}{C}" Then
                        ColsBtn.Add(Cell)
                    End If
                Next
            Next

            RowsBtn.Add(ColsBtn)
        Next
    End Sub
    Private Sub GridBtn_Click(sender As Object, e As EventArgs) Handles GridBtn88.Click, GridBtn87.Click, GridBtn86.Click, GridBtn85.Click, GridBtn84.Click, GridBtn83.Click, GridBtn82.Click, GridBtn81.Click, GridBtn80.Click, GridBtn78.Click, GridBtn77.Click, GridBtn76.Click, GridBtn75.Click, GridBtn74.Click, GridBtn73.Click, GridBtn72.Click, GridBtn71.Click, GridBtn70.Click, GridBtn68.Click, GridBtn67.Click, GridBtn66.Click, GridBtn65.Click, GridBtn64.Click, GridBtn63.Click, GridBtn62.Click, GridBtn61.Click, GridBtn60.Click, GridBtn58.Click, GridBtn57.Click, GridBtn56.Click, GridBtn55.Click, GridBtn54.Click, GridBtn53.Click, GridBtn52.Click, GridBtn51.Click, GridBtn50.Click, GridBtn48.Click, GridBtn47.Click, GridBtn46.Click, GridBtn45.Click, GridBtn44.Click, GridBtn43.Click, GridBtn42.Click, GridBtn41.Click, GridBtn40.Click, GridBtn38.Click, GridBtn37.Click, GridBtn36.Click, GridBtn35.Click, GridBtn34.Click, GridBtn33.Click, GridBtn32.Click, GridBtn31.Click, GridBtn30.Click, GridBtn28.Click, GridBtn27.Click, GridBtn26.Click, GridBtn25.Click, GridBtn24.Click, GridBtn23.Click, GridBtn22.Click, GridBtn21.Click, GridBtn20.Click, GridBtn18.Click, GridBtn17.Click, GridBtn16.Click, GridBtn15.Click, GridBtn14.Click, GridBtn13.Click, GridBtn12.Click, GridBtn11.Click, GridBtn10.Click, GridBtn08.Click, GridBtn07.Click, GridBtn06.Click, GridBtn05.Click, GridBtn04.Click, GridBtn03.Click, GridBtn02.Click, GridBtn01.Click, GridBtn00.Click
        Dim Cell As Windows.Forms.Button = DirectCast(sender, Windows.Forms.Button)
        CurrBtn = Cell

        'Highlight Cell
        HighlightCell(Cell)
        HighlightDupe()
    End Sub
    Private Sub HighlightCell(Cell As Windows.Forms.Button)
        Dim ValidBC As Color = Color.FromArgb(30, 220, 220, 220)

        'Reset Highlight 
        RowsBtn.SelectMany(Function(Row) Row).ToList().ForEach(Sub(Btn) Btn.BackColor = Color.Transparent)

        'Highlight the selected cell
        If Cell IsNot Nothing Then
            Cell.BackColor = Color.FromArgb(50, 100, 100, 225)
            Cell.ForeColor = Color.FromArgb(220, 220, 220)

            'Get the index of the selected cell
            Dim R As Integer = CurrBtn.Name(7).ToString()
            Dim C As Integer = CurrBtn.Name(8).ToString()

            'Highlight the Row, Column, and Sub Grid of the selected Cell
            HighlightRow(Cell, R, ValidBC)
            HighlightCol(Cell, C, ValidBC)
            HighlightSG(Cell, R, C, ValidBC)
        End If
    End Sub

    Private Sub HighlightDupe()
        Dim InvalidBC As Color = Color.FromArgb(30, 255, 100, 100)
        Dim Cell As Windows.Forms.Button

        HighlightCell(CurrBtn)

        'Check each cell if it has duplicates
        For R As Integer = 0 To 8
            For C As Integer = 0 To 8
                Cell = RowsBtn(R)(C)

                'Highlight the Row with red if it has Duplicates
                If DuplicateInRow(R) Then
                    HighlightRow(Cell, R, InvalidBC)
                End If

                'Highlight the Column with red if it has Duplicates
                If DuplicateInCol(C) Then
                    HighlightCol(Cell, C, InvalidBC)
                End If

                'Highlight the Sub Grid with red if it has Duplicates
                If DuplicateInSG(R, C) Then
                    HighlightSG(Cell, R, C, InvalidBC)
                End If
            Next
        Next
    End Sub
    Private Sub HighlightRow(Cell As Windows.Forms.Button, Row As Integer, BC As Color)
        'Highlight Row
        For C As Integer = 0 To 8
            If Not RowsBtn(Row)(C).Equals(Cell) Then
                RowsBtn(Row)(C).BackColor = BC
            End If
        Next
    End Sub
    Private Sub HighlightCol(Cell As Windows.Forms.Button, Col As Integer, BC As Color)
        'Highlight Column
        For R As Integer = 0 To 8
            If Not RowsBtn(R)(Col).Equals(Cell) Then
                RowsBtn(R)(Col).BackColor = BC
            End If
        Next
    End Sub
    Private Sub HighlightSG(Cell As Windows.Forms.Button, Row As Integer, Col As Integer, BC As Color)
        Dim StartingRow As Integer = (Row \ 3) * 3
        Dim StartingCol As Integer = (Col \ 3) * 3

        'Highlight Sub Grid
        For R As Integer = StartingRow To StartingRow + 2
            For C As Integer = StartingCol To StartingCol + 2
                If Not RowsBtn(R)(C).Equals(Cell) Then
                    RowsBtn(R)(C).BackColor = BC
                End If
            Next
        Next
    End Sub
    Private Function ValidGrid()
        'Check each cell if it has duplicates in Row, Column, and Sub Grid
        For R As Integer = 0 To 8
            For C As Integer = 0 To 8
                If DuplicateInRow(R) OrElse DuplicateInCol(C) OrElse DuplicateInSG(R, C) Then
                    Return False
                End If
            Next
        Next

        Return True
    End Function
    Private Function DuplicateInRow(R As Integer)
        Dim RowValues As New List(Of Integer)()

        'Add all the values of the Row into the List
        For Col As Integer = 0 To 8
            Dim Value As Integer = GridValue(R)(Col)

            If Value <> 0 Then
                RowValues.Add(Value)
            End If
        Next

        'Check for Duplicates
        Dim Duplicates As List(Of Integer) = RowValues.GroupBy(Function(X) X).Where(Function(D) D.Count() > 1).Select(Function(G) G.Key).ToList()

        Return Duplicates.Count > 0
    End Function
    Private Function DuplicateInCol(C As Integer)
        Dim Column As New List(Of Integer)()

        'Add all the values of the Column into the List
        For Row As Integer = 0 To 8
            Dim Value As Integer = GridValue(Row)(C)

            If Value <> 0 Then
                Column.Add(Value)
            End If
        Next

        'Check for Duplicates
        Dim Duplicates As List(Of Integer) = Column.GroupBy(Function(X) X).Where(Function(D) D.Count() > 1).Select(Function(G) G.Key).ToList()

        Return Duplicates.Count > 0
    End Function
    Private Function DuplicateInSG(R As Integer, C As Integer)
        Dim SubGridRowIndex As Integer = Math.Floor(R / 3) * 3
        Dim SubGridColIndex As Integer = Math.Floor(C / 3) * 3
        Dim SubGrid As New List(Of Integer)()

        'Add all the values of the Sub Grid into the List
        For RI As Integer = SubGridRowIndex To SubGridRowIndex + 2
            For CI As Integer = SubGridColIndex To SubGridColIndex + 2
                Dim Value As Integer = GridValue(RI)(CI)

                If Value <> 0 Then
                    SubGrid.Add(Value)
                End If
            Next
        Next

        'Check for Duplicates
        Dim Duplicates As List(Of Integer) = SubGrid.GroupBy(Function(X) X).Where(Function(D) D.Count() > 1).Select(Function(G) G.Key).ToList()

        Return Duplicates.Count > 0
    End Function
    Private Sub DisplayCell()
        'Display the solved Sudoku
        For R As Integer = 0 To 8
            For C As Integer = 0 To 8
                RowsBtn(R)(C).Text = GridValue(R)(C).ToString()

                If Flags(R)(C) = 0 Then
                    RowsBtn(R)(C).ForeColor = Color.LightGreen
                End If
            Next
        Next
    End Sub
    Private Sub FuncBtns_Click(sender As Object, e As EventArgs) Handles FuncBtn1.Click, FuncBtn9.Click, FuncBtn8.Click, FuncBtn7.Click, FuncBtn6.Click, FuncBtn5.Click, FuncBtn4.Click, FuncBtn3.Click, FuncBtn2.Click
        Dim FuncBtn As Windows.Forms.Button = DirectCast(sender, Windows.Forms.Button)

        'Add value and text to the selected cell 
        If CurrBtn IsNot Nothing Then
            CurrBtn.Text = FuncBtn.Text.ToString()

            'Get the index of the selected cell
            Dim R As Integer = CurrBtn.Name(7).ToString()
            Dim C As Integer = CurrBtn.Name(8).ToString()

            'Add Value and Flag
            GridValue(R)(C) = FuncBtn.Text.ToString()
            Flags(R)(C) = 1

            HighlightDupe()
        End If
    End Sub
    Private Sub SolveBtn_Click(sender As Object, e As EventArgs) Handles SolveBtn.Click
        'Check if the Grid is valid
        If ValidGrid() Then
            'Solve Grid and display
            If SudokuSolver.SolveSudoku(GridValue, 0, 0) Then
                DisplayCell()
            End If
        Else
            'Show error if invalid
            MessageBox.Show("Duplicate value on the Grid.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        HighlightDupe()
    End Sub
    Private Sub EraseBtn_Click(sender As Object, e As EventArgs) Handles EraseBtn.Click
        'Remove the value and text of the selected cell
        If CurrBtn IsNot Nothing Then
            CurrBtn.Text = ""

            'Get the index of the selected cell
            Dim R As Integer = CurrBtn.Name(7).ToString()
            Dim C As Integer = CurrBtn.Name(8).ToString()

            'Reset the Value and Flag
            GridValue(R)(C) = 0
            Flags(R)(C) = 0

            HighlightDupe()
        End If
    End Sub
    Private Sub ResetBtn_Click(sender As Object, e As EventArgs) Handles ResetBtn.Click
        'Remove Buttons Text
        RowsBtn.SelectMany(Function(Row) Row).ToList().ForEach(Sub(Btn) Btn.Text = "")

        'Reset the value of grid and flags
        GridValue = New Integer(8)() {}
        Flags = New Integer(8)() {}

        For i As Integer = 0 To 8
            GridValue(i) = New Integer(8) {}
            Flags(i) = New Integer(8) {}
        Next

        HighlightDupe()
    End Sub
    Private Sub ExitBtn_Click(sender As Object, e As EventArgs) Handles ExitBtn.Click
        Dim Confirm As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        'Confirm Exit
        If Confirm = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub
    'Erase Button Hover
    Private Sub EraseBtn_MouseEnter(sender As Object, e As EventArgs) Handles EraseBtn.MouseEnter
        EraseBtn.BackgroundImage = My.Resources.EraseBtnHover
    End Sub
    Private Sub EraseBtn_MouseLeave(sender As Object, e As EventArgs) Handles EraseBtn.MouseLeave
        EraseBtn.BackgroundImage = My.Resources.EraseBtn
    End Sub
    'Reset Button Hover
    Private Sub ResetBtn_MouseEnter(sender As Object, e As EventArgs) Handles ResetBtn.MouseEnter
        ResetBtn.BackgroundImage = My.Resources.ResetBtnHover
    End Sub
    Private Sub ResetBtn_MouseLeave(sender As Object, e As EventArgs) Handles ResetBtn.MouseLeave
        ResetBtn.BackgroundImage = My.Resources.ResetBtn
    End Sub
    'Exit Button Hover
    Private Sub ExitBtn_MouseEnter(sender As Object, e As EventArgs) Handles ExitBtn.MouseEnter
        ExitBtn.BackgroundImage = My.Resources.ExitBtnHover
    End Sub
    Private Sub ExitBtn_MouseLeave(sender As Object, e As EventArgs) Handles ExitBtn.MouseLeave
        ExitBtn.BackgroundImage = My.Resources.ExitBtn
    End Sub
End Class