Imports ClosedXML.Excel
Imports System.Web
Imports System.IO

Public Class ExportToExcel
    Public Sub ExportDatatableToExcel(dt As DataTable, saveAsName As String, resp As HttpResponse)
        Dim wb As XLWorkbook = New XLWorkbook()
        Dim sheetname As String = dt.TableName.ToString
        If String.IsNullOrEmpty(sheetname) Then
            sheetname = "CDK_Sheet1"
        End If
        BuildWorkSheet(wb, sheetname, dt)
        FlushWorkBookToHttpResponse(resp, wb, saveAsName)
    End Sub

    Public Sub ExportDataSetToExcel(ds As DataSet, SaveASName As String, resp As HttpResponse)
        Dim wb As XLWorkbook = New XLWorkbook()
        For tableIndex As Integer = 0 To ds.Tables.Count - 1
            Dim dt As DataTable = ds.Tables(tableIndex)
            Dim sheetname As String = dt.TableName.ToString
            If String.IsNullOrEmpty(sheetname) Then
                sheetname = "CDK_Sheet" + (tableIndex + 1).ToString()
            End If
            BuildWorkSheet(wb, sheetname, dt)
        Next
        FlushWorkBookToHttpResponse(resp, wb, SaveASName)
    End Sub

    Private Sub BuildWorkSheet(ByRef wb As XLWorkbook, ByVal sheetname As String, ByVal dt As DataTable)
        Dim ws = wb.Worksheets.Add(sheetname)
        'Building ExcelSheet Headers from the table column names
        For columnNo = 1 To dt.Columns.Count
            ws.Cell(1, columnNo).Value = dt.Columns(columnNo - 1).ColumnName.ToString()
        Next
        For j = 0 To dt.Rows.Count - 1
            For columnCount = 0 To dt.Columns.Count - 1
                Dim q As Integer = Convert.ToInt32(System.Math.Floor(columnCount / 26))
                If (dt.Rows(j)(columnCount).ToString <> "") Then
                    ws.Cell(j + 2, columnCount + 1).DataType = XLCellValues.Text
                    ws.Cell(j + 2, columnCount + 1).Value = "'" + dt.Rows(j)(columnCount).ToString()
                End If
                If (j Mod 2 = 0) Then
                    ws.Cell(j + 2, columnCount + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(216, 228, 188)
                End If
            Next
        Next
        Dim rngTable = ws.Range(1, 1, dt.Rows.Count, dt.Columns.Count)

        Dim rngHeader = ws.Range(1, 1, 1, dt.Columns.Count)
        rngHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left
        rngHeader.Style.Font.Bold = True
        rngHeader.Style.Fill.BackgroundColor = XLColor.FromArgb(155, 187, 89)

        ws.Columns(1, dt.Columns.Count).AdjustToContents()
        ws.RangeUsed().SetAutoFilter()

    End Sub

    Private Sub FlushWorkBookToHttpResponse(ByRef resp As HttpResponse, ByVal wb As XLWorkbook, ByVal saveAsName As String)
        Dim httpResponse As HttpResponse = resp
        httpResponse.Clear()
        httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        httpResponse.AddHeader("content-disposition", "attachment;filename=" + saveAsName)
        ' Flush the workbook to the Response.OutputStream
        Using memoryStream As New MemoryStream()
            wb.SaveAs(memoryStream)
            memoryStream.WriteTo(httpResponse.OutputStream)
            memoryStream.Close()
        End Using
        httpResponse.Flush()
        httpResponse.Close()
    End Sub
End Class
