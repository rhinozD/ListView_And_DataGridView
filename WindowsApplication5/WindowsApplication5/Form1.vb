Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form1
    Public Structure excelData
        Dim ref As String
        Dim cust As String
    End Structure

    Public excelArray() As excelData
    Sub loadExcel()
        Dim con As New OleDbConnection
        Dim Exfilepath As String = "C:\Users\ga-tky-01\Documents\aa.xlsx"
        Dim da As New OleDbDataAdapter("Select * from [Sheet1$]", con)
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim i As Integer = 0
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Exfilepath + ";Extended Properties=""Excel 12.0;HDR=No"""
        ListView1.Items.Clear()
        con.Open()
        ds.Tables.Add(dt)
        da.Fill(dt)
        For Each row In dt.Rows
            ListView1.Items.Add(row.Item(0))
            ListView1.Items(i).SubItems.Add(row.Item(1))
            ListView1.Items(i).SubItems.Add(row.Item(2))
            ListView1.Items(i).SubItems.Add(row.Item(3))
            i = i + 1
        Next
        con.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'ListView1.Items.Add("a")
        'ListView1.Items(0).SubItems.Add("b")
        'ListView1.Items(0).SubItems.Add("c")
        'ListView1.Items(0).SubItems.Add("d")
        'ListView1.Items.Add("1")
        'ListView1.Items(1).SubItems.Add("2")
        'ListView1.Items(1).SubItems.Add("3")
        'ListView1.Items(1).SubItems.Add("4")
        loadExcel()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New OleDbConnection
        Dim Exfilepath As String = "C:\Users\ga-tky-01\Documents\aa.xlsx"
        Dim da As New OleDbDataAdapter("Select * from [Sheet1$]", con)
        Dim dt As New DataTable
        Dim ds As New DataSet
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Exfilepath + ";Extended Properties=""Excel 12.0;HDR=No"""
        con.Open()

        ds.Tables.Add(dt)
        da.Fill(dt)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Columns(0).HeaderText = "a"
        DataGridView1.Columns(1).HeaderText = "b"
        DataGridView1.Columns(2).HeaderText = "c"
        DataGridView1.Columns(3).HeaderText = "d"
        DataGridView1.Refresh()
    End Sub
    Sub exportListView()
        'Try
        '    Me.Cursor = Cursors.WaitCursor
        '    Dim ExcelApp As Object, ExcelBook As Object
        '    Dim ExcelSheet As Object
        '    Dim i As Integer
        '    Dim j As Integer
        '    'create object of excel
        '    ExcelApp = CreateObject("Excel.Application")
        '    ExcelBook = ExcelApp.WorkBooks.Add
        '    ExcelSheet = ExcelBook.WorkSheets(1)
        '    With ExcelSheet
        '        For i = 1 To Me.ListView1.Items.Count
        '            .cells(i, 1) = Me.ListView1.Items(i - 1).Text
        '            For j = 1 To ListView1.Columns.Count - 1
        '                .cells(i, j + 1) = Me.ListView1.Items(i - 1).SubItems(j).Text
        '            Next
        '        Next
        '    End With
        '    ExcelApp.Visible = True
        '    ExcelSheet = Nothing
        '    ExcelBook = Nothing
        '    ExcelApp = Nothing
        '    Me.Cursor = Cursors.Default
        'Catch ex As Exception
        '    Me.Cursor = Cursors.Default
        '    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End Try
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
            For j As Integer = 0 To ListView1.Columns.Count - 1
                xlWorkSheet.Cells(1, col) = ListView1.Columns(j).Text.ToString
                xlWorkSheet.Cells(1, col).EntireRow.Font.Bold = True
            'xlWorkSheet.Range(1, 2).VerticalAlignment = Excel.Constants.xlCenter
                col = col + 1
            Next
            For i = 0 To ListView1.Items.Count - 1
                xlWorkSheet.Cells(i + 2, 1) = ListView1.Items.Item(i).Text.ToString
                xlWorkSheet.Cells(i + 2, 2) = ListView1.Items.Item(i).SubItems(1).Text
                xlWorkSheet.Cells(i + 2, 3) = ListView1.Items.Item(i).SubItems(2).Text
                xlWorkSheet.Cells(i + 2, 4) = ListView1.Items.Item(i).SubItems(3).Text
            Next
            Dim dlg As New SaveFileDialog
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dlg.FilterIndex = 1
            dlg.InitialDirectory = My.Application.Info.DirectoryPath
            dlg.FileName = "test"
            Dim ExcelFile As String = ""
            If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
                ExcelFile = dlg.FileName
                xlWorkSheet.SaveAs(ExcelFile)
        Else
            xlWorkBook.Saved = True
            xlWorkBook.Close()
            End If
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        'Release an automation object
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        exportListView()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                For k As Integer = 1 To DataGridView1.Columns.Count
                    xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
                Next
            Next
        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath
        dlg.FileName = "test"
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        Else
            xlWorkBook.Saved = True
            xlWorkBook.Close()
        End If
        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Sub
End Class
