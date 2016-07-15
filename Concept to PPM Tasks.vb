Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Public Class Form1
    Public xlApp As Excel.Application = New Excel.Application
    Public xlBooks As Excel.Workbooks = Nothing
    Public xlBook As Excel.Workbook = Nothing
    Public xlSheet As Excel.Worksheet = Nothing
    Public xlSheets As Excel.Sheets = Nothing
    Public colnum As Integer = 1
    Dim N As Integer

    Private Sub ButBrowse_Click(sender As Object, e As EventArgs) Handles ButBrowse.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        'Browse for file
        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "H:/"
        fd.Filter = "All Files (*.*)|*.*|CSV (*.csv)|*.csv"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True
        'user presses ok, sets text box to path.
        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Path.Text = strFileName
        End If
    End Sub



    Public Sub ConvertCSVToExcel(Fromcsv As String, Toxlsx As String)
        Dim Exl As New Excel.Application()
        Try
            Dim wb1 As Excel.Workbook = Exl.Workbooks.Open(Fromcsv, Format:=4)
            wb1.SaveAs(Toxlsx, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
            wb1.Close()
            Exl.Quit()
        Catch ex As Exception
            Exl.DisplayAlerts = False
            Exl.Quit()
            MsgBox(ex.Message & " Problem saving as .xlsx. Close all open ""Export"" files.")
        End Try
    End Sub

    Private Sub ButSubmit_Click(sender As Object, e As EventArgs) Handles ButSubmit.Click
        'Set variables
        Dim find(10) As String
        'Apply data to array depending on combobox text.
        If ComboBox1.Text = "Concept" Then
            find(0) = "Task ID"
            find(1) = "Asset Code"
            find(2) = "Asset Description"
            find(3) = "Building"
            find(4) = "Floor"
            find(5) = "Description"
            find(6) = "Room No"
            find(7) = "Key"
            find(8) = "Frequency"
            find(9) = "Reported Date"
            find(10) = "Comments"

        ElseIf ComboBox1.Text = "AXIM" Then

        End If

        If Path.Text <> "" Then
            Dim FileName As String = Nothing
            Dim letter As String = Nothing
            Dim cell As String = Nothing
            Dim CellNumber As String = Nothing
            Dim FilePath As String = Nothing
            Dim proceed As Boolean = Nothing
            Dim count As Integer = 2
            'setup file path
            Dim parts() As String = Split(Path.Text, ".")
            FilePath = parts(0) + "." + parts(1)
            FilePath = FilePath & ".xlsx"
            'run conversion
            Call ConvertCSVToExcel(Path.Text, FilePath)
            FileName = FilePath
            'try to open file
            Try
                xlBooks = xlApp.Workbooks
                xlBook = xlBooks.Open(FileName)
                xlApp.Visible = True
                xlSheet = xlBook.Worksheets(1)
                proceed = True
                'if fail attempt to stop.
            Catch ex As Exception
                MsgBox("Please Close all export files")
                proceed = False
            End Try
            'only if above is success, proceed
            If proceed = True Then
                xlApp.ScreenUpdating = False
                'add new sheet for transfer
                xlBook.Sheets.Add()
                xlBook.Worksheets(1).activate
                'collect the letter
                letter = Filldata.ColSelect(colnum)
                CellNumber = CountCells(letter, 2)

                'Sheet Columns and data fill_________________________________________________
                'finds data depending on coloum headings, if it cant find the coloum, notifies you then skips.

                Try
                    Call Filldata.FillData(find(0), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(0) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(1), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(1) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(2), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(2) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(3), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(3) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(4), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(4) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(5), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(5) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(6), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(6) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(7), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(7) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(8), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(8) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(9), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(9) & " Field")
                End Try

                Try
                    Call Filldata.FillData(find(10), CellNumber)
                Catch ex As Exception
                    MsgBox("Missing " & find(10) & " Field")
                End Try

                'Validation Sheet__________________________________________________
                'runs validation on the sheet (Y/N)

                xlBook.Sheets.Add()
                xlSheet = xlBook.Worksheets(2)
                xlBook.Worksheets(2).activate

                xlSheet.Name = "Data"

                xlSheet.Range("A1").Value = "Y"
                xlSheet.Range("A2").Value = "N"

                xlSheet = xlBook.Worksheets(1)
                xlBook.Worksheets(1).activate

                'Validation Columns and data _______________________________________
                'Adds specific validation to cells.

                CellNumber = CellNumber - 1
                Call FillValidatedColumn("Good Condition", CellNumber)
                Call FillValidatedColumn("Adjusted / Repaired", CellNumber)
                Call FillValidatedColumn("Attention Required", CellNumber)
                Call FillValidatedColumn("Filters Cleaned or Replaced", CellNumber)

                Call FillBlankColumn("Comments / Findings", CellNumber)
                Call FillBlankColumn("Date", CellNumber)
                Call FillBlankColumn("Engineer", CellNumber)

                'Colouring Cells_____________________________________________________
                'colours cells depending on what is editable and what isnt.

                Dim ColCount As Integer = Nothing
                Dim RowCount As Integer = Nothing
                Dim LockLetter As String = Nothing
                Dim RightCell As String = Nothing

                xlSheet.Range("A1").Select()
                'set top rows to dark
                Do Until xlApp.Selection.value = Nothing
                    Call ColorDark()
                    xlApp.Selection.Offset(0, 1).Select
                Loop

                xlSheet.Range("A2").Select()
                'count coloums
                Do Until xlApp.Selection.value = Nothing
                    ColCount = ColCount + 1
                    xlApp.Selection.Offset(0, 1).Select
                Loop

                LockLetter = Filldata.ColSelect(ColCount)
                xlSheet.Range("A2").Select()
                'count rows
                Do Until xlApp.Selection.value = Nothing
                    RowCount = RowCount + 1
                    xlApp.Selection.Offset(1, 0).Select
                Loop

                RightCell = LockLetter & (RowCount + 1)

                With xlSheet.Range("A2:" & RightCell).Interior
                    .Pattern = Excel.XlPattern.xlPatternSolid
                    .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With

                xlBook.Worksheets(1).Rows(count & ":" & (CellNumber + 1)).RowHeight = 40
                count = 0
                xlApp.DisplayAlerts = False
                xlApp.Sheets(3).delete
                xlSheet = xlBook.Worksheets(1)
                xlBook.Worksheets(1).activate
                xlSheet.Name = "Export"
                CellNumber = CellNumber + 3

                'Sheet Bottom_________________________________________________

                xlSheet.Range("A" & CellNumber).Select()
                Dim numrows As Long, numcolumns As Integer
                numrows = xlApp.Selection.Rows.Count
                numcolumns = xlApp.Selection.Columns.Count
                xlApp.Selection.Resize(numrows, numcolumns + 4).Select
                xlApp.Selection.merge
                Call ColorDark()
                xlApp.Selection.value = "Comments / Notes"
                Call borders()

                Do Until count = 10
                    xlApp.Selection.Offset(1, 0).Select
                    xlApp.Selection.Resize(numrows, numcolumns + 4).Select
                    xlApp.Selection.merge
                    xlApp.Selection.locked = False
                    Call borders()
                    count = count + 1
                Loop

                xlSheet.Range("G" & CellNumber).Select()
                Call ColorDark()
                xlApp.Selection.value = "All Filters Required?"
                Call borders()

                xlSheet.Range("H" & CellNumber).Select()
                Call ValidateCell()
                xlApp.Selection.locked = False
                Call borders()

                xlSheet.Range("G" & (CellNumber + 2)).Select()
                xlApp.Selection.Resize(numrows + 5, numcolumns).Select
                xlApp.Selection.merge
                With xlApp.Selection
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                    .WrapText = True
                End With
                Call ColorDark()
                xlApp.Selection.value = "Specific Filters Required - Specify Sizes"
                Call borders()

                xlSheet.Range("H" & (CellNumber + 2)).Select()
                count = 0
                Do Until count = 6
                    xlApp.Selection.Resize(numrows, numcolumns + 1).Select
                    xlApp.Selection.merge
                    xlApp.Selection.locked = False
                    Call borders()
                    xlApp.Selection.Offset(1, 0).Select
                    count = count + 1
                Loop

                xlSheet.Range("K" & CellNumber).Select()
                xlApp.Selection.Resize(numrows, numcolumns + 2).Select
                xlApp.Selection.merge
                Call ColorDark()
                xlApp.Selection.value = "All Belts Required?"
                Call borders()

                xlSheet.Range("N" & CellNumber).Select()
                Call ValidateCell()
                xlApp.Selection.locked = False
                Call borders()

                xlSheet.Range("K" & (CellNumber + 2)).Select()
                xlApp.Selection.WrapText = True
                xlApp.Selection.Resize(numrows + 5, numcolumns + 2).Select
                xlApp.Selection.merge
                With xlApp.Selection
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                    .WrapText = True
                End With
                Call ColorDark()
                xlApp.Selection.value = "Specific Belts Required - Specify Sizes"
                Call borders()

                xlSheet.Range("N" & (CellNumber + 2)).Select()
                count = 0
                Do Until count = 6
                    xlApp.Selection.Resize(numrows, numcolumns + 1).Select
                    xlApp.Selection.merge
                    xlApp.Selection.locked = False
                    Call borders()
                    xlApp.Selection.Offset(1, 0).Select
                    count = count + 1
                Loop

                xlSheet.Rows("2:2").Select
                xlApp.ActiveWindow.FreezePanes = True

                xlSheet.Range("A1").Select()

                colnum = 1

                xlApp.DisplayAlerts = True
                xlApp.ScreenUpdating = True

                With xlApp.ActiveSheet
                    .Protect
                    .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
                    .Protect
                End With

                'Dim fd As SaveFileDialog = New SaveFileDialog()
                'Dim strFileName As String

                'fd.Title = "Open File Dialog"
                'fd.InitialDirectory = "S:\shared Documents\AFM Guernsey\PPMs\"
                'fd.Filter = "Excel Document (*.xlsx)|*.xlsx|Word Document (*.docx)|*.docx"
                'fd.FilterIndex = 1
                'fd.OverwritePrompt = True
                'fd.RestoreDirectory = False

                'If fd.ShowDialog() = DialogResult.OK Then
                '    strFileName = fd.FileName
                'End If

                MsgBox("Macro Complete")
            Else

            End If
        Else

            MsgBox("Browse for file first")

        End If
    End Sub

    Sub ColorDark()
        With xlApp.Selection.Interior
            .Pattern = Excel.XlPattern.xlPatternSolid
            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorLight2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End Sub


    Sub borders()

        xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                             Weight:=Excel.XlBorderWeight.xlThin,
                                             ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)

    End Sub

    Sub FillBlankColumn(ByVal value As String, ByVal cellnumber As Integer)
        Dim letter, cell As String

        letter = Filldata.ColSelect(colnum)
        cell = letter & "1"
        xlSheet.Range(cell).Value = value
        xlSheet.Range(cell).Select()
        Call borders()

        For N = 1 To cellnumber + 1
            Call borders()
            xlApp.Selection.Offset(1, 0).Select
            xlApp.Selection.locked = False
        Next N
        If value = "Date" Then
            xlBook.Worksheets(1).Columns(letter).ColumnWidth = 15
        Else
            xlBook.Worksheets(1).Columns(letter).AutoFit
        End If

        colnum = colnum + 1

    End Sub

    Sub FillValidatedColumn(ByVal value As String, ByVal cellnumber As Integer)
        Dim letter, cell As String

        letter = Filldata.ColSelect(colnum)
        cell = letter & "1"
        xlSheet.Range(cell).Value = value
        xlSheet.Range(cell).Select()
        xlSheet.Range(cell).Font.Size = 8
        Call borders()
        xlSheet.Range(cell).WrapText = True
        xlBook.Worksheets(1).Columns(letter).AutoFit
        Call FillDataAndValidate(letter, cellnumber)
        colnum = colnum + 1

    End Sub

    Function CountCells(ByVal letter As Char, ByVal count As Integer)
        Dim counter As Integer = 1
        Dim cell As String
        Dim pass As Boolean = False

        cell = letter & count
        xlSheet = xlBook.Worksheets(2)
        xlBook.Worksheets(2).activate

        xlSheet.Range(cell).Select()
        Do Until xlApp.Selection.Value = ""
            counter = counter + 1
            xlApp.Selection.Offset(1, 0).Select
        Loop

        Return counter
    End Function

    Sub FillDataAndValidate(ByVal letter As Char, ByVal CellNumber As Integer)
        Dim count As Integer = 2
        Dim cell As String
        Dim pass As Boolean = False

        cell = letter & count

        xlSheet = xlBook.Worksheets(1)
        xlBook.Worksheets(1).activate

        xlSheet.Range(cell).Select()

        For N = 1 To CellNumber
            Call borders()
            Call ValidateCell()
            xlApp.Selection.locked = False
            xlApp.Selection.Offset(1, 0).Select
        Next N
    End Sub
    Sub ValidateCell()
        With xlApp.Selection.Validation
            .Delete
            .Add(Type:=Excel.XlDVType.xlValidateList,
                 AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                 Operator:=Excel.XlFormatConditionOperator.xlBetween,
                 Formula1:="=Data!A1:A2")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Version.Text = String.Format("Version {0}", System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "AXIM" Then
            ButSubmit.Enabled = False
        Else
            ButSubmit.Enabled = True
        End If
    End Sub
End Class
