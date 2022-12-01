'Fungsi Export Excel dengan penyimpanan otomatis ke Drive
Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ComboBox1.Text = "Tanpa Sparepart" Then

            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()
            'add column to that table
            For i As Integer = 0 To DataGridView1.ColumnCount - 1
                dset.Tables(0).Columns.Add(DataGridView1.Columns(i).HeaderText)
            Next
            'add rows to the table
            Dim dr1 As DataRow
            For i As Integer = 0 To DataGridView1.RowCount - 1
                dr1 = dset.Tables(0).NewRow
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    dr1(j) = DataGridView1.Rows(i).Cells(j).Value
                Next
                dset.Tables(0).Rows.Add(dr1)
            Next

            Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            wBook = excel.Workbooks.Add()
            wSheet = wBook.ActiveSheet()

            Dim dt As System.Data.DataTable = dset.Tables(0)
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                Next
            Next

            wSheet.Columns.AutoFit()
            Dim strFileName As String = "D:\Sumary_barang.xls"
            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try

            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            wBook.SaveAs(strFileName)
            excel.Workbooks.Open(strFileName)
            excel.Visible = True

           
        ElseIf ComboBox1.Text = "Dengan Sparepart" Then
            'verfying the datagridview having data or not
            'If ((dgv_notablmlunas.Columns.Count = 0) Or (dgv_notablmlunas.Rows.Count = 0)) Then
            '    Exit Sub
            'End If

            'Creating dataset to export
            Dim dset As New DataSet
            'add table to dataset
            dset.Tables.Add()
            'add column to that table
            For i As Integer = 0 To DataGridView2.ColumnCount - 1
                dset.Tables(0).Columns.Add(DataGridView2.Columns(i).HeaderText)
            Next
            'add rows to the table
            Dim dr1 As DataRow
            For i As Integer = 0 To DataGridView2.RowCount - 1
                dr1 = dset.Tables(0).NewRow
                For j As Integer = 0 To DataGridView2.Columns.Count - 1
                    dr1(j) = DataGridView2.Rows(i).Cells(j).Value
                Next
                dset.Tables(0).Rows.Add(dr1)
            Next

            Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            wBook = excel.Workbooks.Add()
            wSheet = wBook.ActiveSheet()

            Dim dt As System.Data.DataTable = dset.Tables(0)
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                Next
            Next

            wSheet.Columns.AutoFit()
            Dim strFileName As String = "D:\Sumary_barang_parts.xls"
            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try

            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            wBook.SaveAs(strFileName)
            excel.Workbooks.Open(strFileName)
            excel.Visible = True



        End If
      
    End Sub



'Fungsi Export Excel tanpa penyimpanan otomatis ke Drive
Sub ExportToExcel()
    Dim ExcelApp As Object, ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    Dim j As Integer

    'create object of excel
    ExcelApp = CreateObject("Excel.Application")
    ExcelBook = ExcelApp.WorkBooks.Add
    ExcelSheet = ExcelBook.WorkSheets(1)

    With ExcelSheet
        For Each column As DataGridViewColumn In dgvkendala_penggunaan.Columns
            .cells(1, column.Index + 1) = column.HeaderText
        Next
        For i = 1 To Me.dgvkendala_penggunaan.RowCount
            .cells(i + 1, 1) = Me.dgvkendala_penggunaan.Rows(i - 1).Cells("Column1").Value
            For j = 1 To dgvkendala_penggunaan.Columns.Count - 1
                .cells(i + 1, j + 1) = dgvkendala_penggunaan.Rows(i - 1).Cells(j).Value
            Next
        Next
    End With
    ExcelApp.Visible = True
    '
    ExcelSheet = Nothing
    ExcelBook = Nothing
    ExcelApp = Nothing
End Sub