'Imports System
'Imports System.IO
'Imports Microsoft.Office.Interop

'Module CreateWorkbook

'    Public WB As New Microsoft.Office.Interop.Excel.Application
'    Public WS As New Microsoft.Office.Interop.Excel.Worksheet

'    Sub Main()
'        Dim r As Integer

'        'Create column headers with borders
'        With WS
'            .Range("A1:M500").Select()
'            With WB.Selection.Interior
'                .ColorIndex = 2
'                .Pattern = 1
'                .PatternColorIndex = -4105
'            End With
'            .Cells(r, 1).Value = "Client"
'            .Range("A1").BorderAround(, XlBorderWeight.xlThin, Microsoft.Office.Interop.XlColorIndex.xlColorIndexAutomatic)

'            .Range("A1").BorderAround(microsoft.Office.Interop.Excel.Application

'            .Cells(r, 2).Value = "Category"
'            .Range("B1").BorderAround(, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic)

'            XlColorIndex.xlColorIndexAutomatic)
'            .Cells(r, 3).Value = "Source"
'            .Range("C1").BorderAround(, XlBorderWeight.xlThin, _
'            XlColorIndex.xlColorIndexAutomatic)
'            .Cells(r, 4).Value = "# Files"
'            .Range("D1").BorderAround(, XlBorderWeight.xlThin, _
'            XlColorIndex.xlColorIndexAutomatic)
'            .Cells(r, 5).Value = "Ticket ID"
'            .Range("E1").BorderAround(, XlBorderWeight.xlThin, _
'            XlColorIndex.xlColorIndexAutomatic)
'            .Cells(r, 6).Value = "Assigned To"
'            .Range("F1").BorderAround(, XlBorderWeight.xlThin, _
'            XlColorIndex.xlColorIndexAutomatic)
'            .Cells(r, 7).Value = "Path"
'            .Range("E1").BorderAround(, XlBorderWeight.xlThin, _
'            XlColorIndex.xlColorIndexAutomatic)

'            'Set background color for columns A through F
'            .Range("A" & r & ":G" & r).Select()
'            With WB.Selection
'                .Interior.ColorIndex = 1 'Black Background
'                .Interior.Pattern = 1
'                .Interior.PatternColorIndex = -4105
'                .Font.FontStyle = "Bold"
'                .Font.Size = "8"
'                .Font.ColorIndex = 2 'White Font
'                .HorizontalAlignment = Excel.Constants.xlCenter
'            End With

'            'Set column widths
'            .Range("A1").Select()
'            With WB.Selection
'                .ColumnWidth = 19.71
'            End With
'            .Range("B1").Select()
'            With WB.Selection
'                .ColumnWidth = 14.57
'            End With
'            .Range("C1").Select()
'            With WB.Selection
'                .ColumnWidth = 24
'            End With
'            .Range("D1").Select()
'            With WB.Selection
'                .ColumnWidth = 5.29
'            End With
'            .Range("E1").Select()
'            With WB.Selection
'                .ColumnWidth = 15.29
'            End With
'            .Range("F1").Select()
'            With WB.Selection
'                .ColumnWidth = 21.71
'            End With
'            .Range("G1").Select()
'            With WB.Selection
'                .ColumnWidth = 33.57
'            End With
'        End With

'        'Here is where I run a function to create all of my data
'        'This function loops through a directory structure recursively and does some stuff

'        'Set the borders for row 2 through the end
'        With WS.Range("A2:G" & r)
'            .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlHairline
'            .Borders(XlBordersIndex.xlInsideHorizontal).ColorIndex = 15
'            .Borders(XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlHairline
'            .Borders(XlBordersIndex.xlInsideVertical).ColorIndex = 15
'            .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeLeft).ColorIndex = 1
'            .Borders(XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeRight).ColorIndex = 1
'            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeBottom).ColorIndex = 1
'            .Borders(XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeTop).ColorIndex = 1
'        End With

'        'Format alignment for all except totals section
'        With WS

'            .Range("A1:F" & r).Select()
'            With WB.Selection
'                .HorizontalAlignment = Excel.Constants.xlCenter
'            End With

'            .Range("G2:G" & r).Select()
'            With WB.Selection
'                .HorizontalAlignment = Excel.Constants.xlLeft
'            End With

'        End With
'        'Create arrays for totals section
'        Dim countersary() As Integer = {cwcount, hrcount, hlcount, lbcount, mbcount, _
'        mccount, prcount, pccount, pscount, phcount, _
'        rfcount, rgcount, rpcount, urcount, othercount, _
'        totalfiles}
'        Dim ftypesary() As String = {"xxx1", "xxx2", "xxx3", "xxx4", "xxx5", _
'        "xxx6", "xxx7", "xxx8", _
'        "xxx9", "xxx10", "xxx11", "xxx12", _
'        "xxx13", "xxx14", "Other", ""}

'        'Loop through counters to write totals
'        Dim i As Integer = 0
'        For i = 0 To UBound(countersary)

'            'Go to next row on spreadsheet
'            r += 1

'            With WS

'                'Write each total
'                .Cells(r, 2).Value = "Total " & ftypesary(i) & " Files"
'                .Cells(r, 4).Value = countersary(i)
'                sw.Write("," & "Total " & ftypesary(i) & " Files,,")
'                sw.WriteLine(countersary(i) & ",,,")

'                'Format columns A - C of totals section
'                .Range("A" & r & ":C" & r).Select()
'                With WB.Selection
'                    .Font.FontStyle = "Bold"
'                    .HorizontalAlignment = Excel.Constants.xlLeft
'                End With

'                'Format columns D - G of totals section
'                .Range("D" & r & ":G" & r).Select()
'                With WB.Selection
'                    .Font.FontStyle = "Bold"
'                    .HorizontalAlignment = Excel.Constants.xlRight
'                End With

'            End With

'        Next

'        'Set variables for range of the totals section
'        Dim x As Integer = UBound(ftypesary)
'        Dim y As Integer = r - x

'        'Format borders around the totals section
'        With WS.Range("A" & y & ":G" & r)
'            .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlHairline
'            .Borders(XlBordersIndex.xlInsideHorizontal).ColorIndex = 15
'            '.Borders(XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
'            '.Borders(XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlHairline
'            '.Borders(XlBordersIndex.xlInsideVertical).ColorIndex = 15
'            .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeLeft).ColorIndex = 1
'            .Borders(XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeRight).ColorIndex = 1
'            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeBottom).ColorIndex = 1
'            .Borders(XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
'            .Borders(XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin
'            .Borders(XlBordersIndex.xlEdgeTop).ColorIndex = 1
'        End With

'        'Unselect multiple cells
'        WS.Range("A1").Select()


'        'Set up variables to save the file
'        Dim FileName As String = "Filename_" & mytime & ".xls"
'        Dim SpreadSheet As String = WorkingDir & FileName
'        Dim CSVSave As String = WorkingDir & CSVFileName

'        Try

'            'Save the workbook
'            WS.SaveAs(SpreadSheet)
'            WS.Close()
'        Catch ex As Exception
'            Console.WriteLine(ex.Message)
'            sw.WriteLine("ERROR ENCOUNTERED WITH TICKETCHECK - line 269")
'            'Console.ReadLine()
'        End Try
'    End Sub
'End Module