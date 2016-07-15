Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Public Class Form1
    Dim xlApp As Excel.Application = New Excel.Application
    Dim xlBooks As Excel.Workbooks = Nothing
    Dim xlBook As Excel.Workbook = Nothing
    Dim xlSheet As Excel.Worksheet = Nothing
    Dim xlSheets As Excel.Sheets = Nothing
    Dim colnum As Integer = 1
    Dim N As Integer

    Private Sub ButBrowse_Click(sender As Object, e As EventArgs) Handles ButBrowse.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "H:/"
        fd.Filter = "All Files (*.*)|*.*|CSV (*.csv)|*.csv"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Path.Text = strFileName
        End If
    End Sub

    Function ColSelect(ByVal count)
        Dim letter As String = Nothing

        Select Case count
            Case 1
                letter = "A"
            Case 2
                letter = "B"
            Case 3
                letter = "C"
            Case 4
                letter = "D"
            Case 5
                letter = "E"
            Case 6
                letter = "F"
            Case 7
                letter = "G"
            Case 8
                letter = "H"
            Case 9
                letter = "I"
            Case 10
                letter = "J"
            Case 11
                letter = "K"
            Case 12
                letter = "L"
            Case 13
                letter = "M"
            Case 14
                letter = "N"
            Case 15
                letter = "O"
            Case 16
                letter = "P"
            Case 17
                letter = "Q"
            Case 18
                letter = "R"
            Case 19
                letter = "S"
            Case 20
                letter = "T"
            Case 21
                letter = "U"
            Case 22
                letter = "V"
            Case 23
                letter = "W"
            Case 24
                letter = "X"
            Case 25
                letter = "Y"
            Case 26
                letter = "Z"
            Case 27
                letter = "AA"
            Case 28
                letter = "AB"
            Case 29
                letter = "AC"
            Case 30
                letter = "AD"
            Case 31
                letter = "AE"
            Case 32
                letter = "AF"
            Case 33
                letter = "AG"
            Case 34
                letter = "AH"
            Case 35
                letter = "AI"
            Case 36
                letter = "AJ"
            Case 37
                letter = "AK"
            Case 38
                letter = "AL"
            Case 39
                letter = "AM"
            Case 40
                letter = "AN"
            Case 41
                letter = "AO"
            Case 42
                letter = "AP"
            Case 43
                letter = "AQ"
            Case 44
                letter = "AR"
            Case 45
                letter = "AS"
            Case 46
                letter = "AT"
            Case 47
                letter = "AU"
            Case 48
                letter = "AV"
            Case 49
                letter = "AW"
            Case 50
                letter = "AX"
            Case 51
                letter = "AY"
            Case 52
                letter = "AZ"
        End Select

        Return letter
    End Function

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
        If Path.Text <> "" Then
            Dim FileName As String = Nothing
            Dim letter As String = Nothing
            Dim cell As String = Nothing
            Dim CellNumber As String = Nothing
            Dim FilePath As String = Nothing
            Dim proceed As Boolean
            Dim count As Integer = 0

            Dim parts() As String = Split(Path.Text, ".")
            FilePath = parts(0)

            FilePath = FilePath & ".xlsx"

            Call ConvertCSVToExcel(Path.Text, FilePath)

            FileName = FilePath

            Try
                xlBooks = xlApp.Workbooks
                xlBook = xlBooks.Open(FileName)
                xlApp.Visible = True
                xlSheet = xlBook.Worksheets(1)
                proceed = True
            Catch ex As Exception
                MsgBox("Please Close all export files")
                proceed = False
            End Try

            If proceed = True Then
                xlApp.ScreenUpdating = False

                xlBook.Sheets.Add()
                xlBook.Worksheets(1).activate

                letter = ColSelect(colnum)
                CellNumber = CountCells(letter, 2)

                Try
                    Call FillData("Task ID", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Task ID Field")
                End Try

                Try
                    Call FillData("Asset Code", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Asset Code Field")
                End Try

                Try
                    Call FillData("Asset Description", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Asset Description Field")
                End Try

                Try
                    Call FillData("Building", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Building Field")
                End Try

                Try
                    Call FillData("Floor", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Floor Field")
                End Try

                Try
                    Call FillData("Description", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Description Field")
                End Try

                Try
                    Call FillData("Room No", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Room No Field")
                End Try

                Try
                    Call FillData("Frequency", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Frequency Field")
                End Try

                Try
                    Call FillData("Reported Date", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Reported Date Field")
                End Try

                Try
                    Call FillData("Comments", CellNumber)
                Catch ex As Exception
                    MsgBox("Missing Comments Field")
                End Try

                xlBook.Sheets.Add()
                xlSheet = xlBook.Worksheets(2)
                xlBook.Worksheets(2).activate

                xlSheet.Name = "Data"

                xlSheet.Range("A1").Value = "Yes"
                xlSheet.Range("A2").Value = "No"

                xlSheet = xlBook.Worksheets(1)
                xlBook.Worksheets(1).activate

                CellNumber = CellNumber - 1
                Call FillValidatedColumn("Good Condition", CellNumber)
                Call FillValidatedColumn("Adjusted / Repaired", CellNumber)
                Call FillValidatedColumn("Attention Required", CellNumber)
                Call FillValidatedColumn("Filters Cleaned or Replaced", CellNumber)

                Call FillBlankColumn("Comments / Findings", CellNumber)
                Call FillBlankColumn("Date", CellNumber)
                Call FillBlankColumn("Engineer", CellNumber)

                xlApp.DisplayAlerts = False

                xlApp.Sheets(3).delete

                xlSheet = xlBook.Worksheets(1)
                xlBook.Worksheets(1).activate

                xlSheet.Name = "Export"
                xlApp.ScreenUpdating = True

                xlApp.Selection.Offset(1, -16).Select
                Dim numrows As Long, numcolumns As Integer
                numrows = xlApp.Selection.Rows.Count
                numcolumns = xlApp.Selection.Columns.Count
                xlApp.Selection.Resize(numrows, numcolumns + 4).Select
                xlApp.Selection.merge
                xlApp.Selection.value = "Comments / Notes"
                xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                             Weight:=Excel.XlBorderWeight.xlThin,
                                             ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)

                Do Until count = 10
                    xlApp.Selection.Offset(1, 0).Select
                    xlApp.Selection.Resize(numrows, numcolumns + 4).Select
                    xlApp.Selection.merge
                    xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                             Weight:=Excel.XlBorderWeight.xlThin,
                                             ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
                    count = count + 1
                Loop

                xlApp.DisplayAlerts = True
                xlApp.ScreenUpdating = True

                colnum = 1

                MsgBox("Macro Complete")
            Else

            End If
        Else

            MsgBox("Browse for file first")

        End If
    End Sub

    Sub FillBlankColumn(ByVal value As String, ByVal cellnumber As Integer)
        Dim letter, cell As String

        letter = ColSelect(colnum)
        cell = letter & "1"
        xlSheet.Range(cell).Value = value
        xlSheet.Range(cell).Select()
        xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                             Weight:=Excel.XlBorderWeight.xlThin,
                                             ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
        For N = 1 To cellnumber + 1
            xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                             Weight:=Excel.XlBorderWeight.xlThin,
                                             ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
            xlApp.Selection.Offset(1, 0).Select
        Next N
        xlBook.Worksheets(1).Columns(letter).AutoFit
        colnum = colnum + 1

    End Sub

    Sub FillValidatedColumn(ByVal value As String, ByVal cellnumber As Integer)
        Dim letter, cell As String

        letter = ColSelect(colnum)
        cell = letter & "1"
        xlSheet.Range(cell).Value = value
        xlSheet.Range(cell).Select()
        xlSheet.Range(cell).Font.Size = 8
        xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                                 Weight:=Excel.XlBorderWeight.xlThin,
                                                 ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
        xlApp.ActiveCell.Orientation = 90
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
            xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                                     Weight:=Excel.XlBorderWeight.xlThin,
                                                     ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
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
            xlApp.Selection.Offset(1, 0).Select
        Next N
    End Sub

    Sub FillData(ByVal locate As String, ByVal cellnumber As Integer)
        Dim count As Integer = 1
        Dim letter As String
        Dim cell As String
        Dim endcell As String
        Dim pass As Boolean = False

        xlSheet = xlBook.Worksheets(1)
        xlBook.Worksheets(1).activate

        letter = ColSelect(colnum)
        cell = letter & "1"
        xlSheet.Range(cell).Select()

        xlBook.Worksheets(2).Activate
        xlSheet = xlBook.Worksheets(2)

        Do Until pass = True

            letter = ColSelect(count)
            count = count + 1
            cell = letter & "1"
            endcell = letter & "200"
            xlSheet.Range(cell).Select()

            If xlSheet.Range(cell).Value = locate Then
                pass = True

                For N = 1 To cellnumber
                    If xlApp.Selection.value <> Nothing Then
                        xlApp.Selection.Copy
                        xlSheet = xlBook.Worksheets(1)
                        xlBook.Worksheets(1).activate
                        xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteValues)
                        If locate = "Reported Date" Then
                            xlApp.Selection.numberformat = "dd/mm/yyyy hh:mm"
                        End If
                        xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                                     Weight:=Excel.XlBorderWeight.xlThin,
                                                     ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
                            xlApp.Selection.Offset(1, 0).Select
                            xlBook.Worksheets(2).Activate
                            xlSheet = xlBook.Worksheets(2)
                        ElseIf xlApp.Selection.value = Nothing Then
                            xlSheet = xlBook.Worksheets(1)
                        xlBook.Worksheets(1).activate
                        xlApp.Selection.BorderAround(linestyle:=Excel.XlLineStyle.xlContinuous,
                                                     Weight:=Excel.XlBorderWeight.xlThin,
                                                     ColorIndex:=Excel.XlColorIndex.xlColorIndexAutomatic)
                        xlApp.Selection.Offset(1, 0).Select
                        xlBook.Worksheets(2).Activate
                        xlSheet = xlBook.Worksheets(2)
                    End If
                    xlApp.Selection.Offset(1, 0).Select
                Next N
            End If
        Loop
        xlSheet = xlBook.Worksheets(2)
        xlBook.Worksheets(2).activate

        letter = ColSelect(colnum)
        colnum = colnum + 1

        xlBook.Worksheets(1).Columns(letter).AutoFit
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Version.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
    End Sub
End Class
