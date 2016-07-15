Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Module Filldata
    Dim N As Integer
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



    Sub FillData(ByVal locate As String, ByVal cellnumber As Integer)
        Dim count As Integer = 1
        Dim letter As String
        Dim cell As String
        Dim endcell As String
        Dim pass As Boolean = False

        Form1.xlSheet = Form1.xlBook.Worksheets(1)
        Form1.xlBook.Worksheets(1).activate

        letter = ColSelect(Form1.colnum)
        cell = letter & "1"
        Form1.xlSheet.Range(cell).Select()

        Form1.xlBook.Worksheets(2).Activate
        Form1.xlSheet = Form1.xlBook.Worksheets(2)

        Do Until pass = True

            letter = ColSelect(count)
            count = count + 1
            cell = letter & "1"
            endcell = letter & cellnumber
            Form1.xlSheet.Range(cell).Select()

            If Form1.xlSheet.Range(cell).Value = locate Then
                If locate = "Key" Then
                    letter = ColSelect(count)
                    count = count + 1
                    cell = letter & "1"
                    endcell = letter & cellnumber
                    Form1.xlSheet.Range(cell).Select()
                    Do Until Form1.xlSheet.Range(cell).Value = locate
                        letter = ColSelect(count)
                        count = count + 1
                        cell = letter & "1"
                        endcell = letter & cellnumber
                        Form1.xlSheet.Range(cell).Select()
                    Loop
                End If

                pass = True

                Form1.xlSheet.Range(cell & ":" & endcell).Select()
                Form1.xlApp.Selection.Copy
                Form1.xlSheet = Form1.xlBook.Worksheets(1)
                Form1.xlBook.Worksheets(1).activate
                Form1.xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteValues)
                If locate = "Reported Date" Then
                    Form1.xlApp.Selection.numberformat = "dd/mm/yyyy hh: mm"
                End If
                With Form1.xlApp.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With Form1.xlApp.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With Form1.xlApp.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With Form1.xlApp.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With Form1.xlApp.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With Form1.xlApp.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With

                Form1.xlBook.Worksheets(2).Activate
                Form1.xlSheet = Form1.xlBook.Worksheets(2)


            End If

            Form1.xlApp.Selection.offset(0, 1).select

        Loop
        Form1.xlSheet = Form1.xlBook.Worksheets(2)
        Form1.xlBook.Worksheets(2).activate

        letter = ColSelect(Form1.colnum)
        Form1.colnum = Form1.colnum + 1

        If locate <> "Comments" Then
            Form1.xlBook.Worksheets(1).Columns(letter).AutoFit
        Else
            Form1.xlBook.Worksheets(1).Columns(letter).AutoFit
            Form1.xlBook.Worksheets(1).Columns(letter).WrapText = True
            Form1.xlBook.Worksheets(1).Columns(letter).AutoFit
            Form1.xlBook.Worksheets(1).Columns(letter).ColumnWidth = 40

        End If

    End Sub
End Module
