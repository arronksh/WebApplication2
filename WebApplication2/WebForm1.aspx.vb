Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class WebForm1
    Inherits System.Web.UI.Page
    Dim files As String
    Dim vNoorder As String
    Dim vNama As String
    Dim vTempJmlBarang As Integer
    Dim vItem As String
    Dim vService As String
    Dim vOngkir As Integer
    Dim vAsuransi As Integer
    Dim vHargaMasing As String
    Dim vTotal As Integer
    Dim vGTotal As Integer
    Dim vKeterangan As String
    Dim vStatus As String
    Dim Post As Boolean
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("user") <> "" Then
            Label1.Text = Session("user").ToString
        Else
            Response.Redirect("~/loginForm.aspx")
        End If
    End Sub
    Sub ubahDelimiter(ByVal filename As String)
        Dim lines() As String = IO.File.ReadAllLines(filename)
        Dim cekSep As String = lines(0).ToString.ToUpper
        If Not cekSep.Contains("SEP=") Then
            Dim value As String = File.ReadAllText(filename)
            If value.Substring(0, 20).Contains(";") Then
                value = "SEP=;" & vbCrLf & value
                IO.File.WriteAllText(filename, value)
            Else
                value = "SEP=," & vbCrLf & value
                IO.File.WriteAllText(filename, value)
            End If
        End If
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
            MsgBox(ex.Message)
        Finally
            GC.Collect()
        End Try
    End Sub
    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim excelReader As New 
        interopx()
    End Sub
    Sub interopx()
        files = FileUpload1.PostedFile.FileName
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim range As Excel.Range
        Dim Obj As Object
        files = "C:\Users\RON ROG\Downloads\Tokopedia_Order_20180728.xlsx"
        Try
            Dim extension As String = Path.GetExtension(files)
            'jika .csv, ubah SEP=, / ; dulu
            If extension = ".csv" Then
                ubahDelimiter(files)
            End If
            'execute
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(files)
            xlWorkSheet = xlWorkBook.ActiveSheet
            'display the cells value B2
            'MsgBox(xlWorkSheet.Cells(5, 4).value)
            'edit the cell with new value
            'xlWorkSheet.Cells(3, 4) = "http://vb.net-informations.com"
            range = xlWorkSheet.UsedRange
            Obj = CType(range.Cells(1, 1), Excel.Range)
            'MsgBox("-" & Obj.value & "-")
            If Obj.value = "Nama Toko:" Then
                TextBox1.Text = CType(range.Cells(5, 4), Excel.Range).Value
                'fileToped(range)

            End If
            xlApp.DisplayAlerts = False
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
        Catch ex As Exception
            MsgBox(ex.Message)
            xlApp.DisplayAlerts = False
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
        End Try
    End Sub
    Sub fileToped(ByVal range As Excel.Range)
        'get Tanggal
        For rCnt = 5 To range.Rows.Count
            'MsgBox("masuk FOR")
            If Not (IsNothing(CType(range.Cells(rCnt, 1), Excel.Range).Value)) Then
                vNoorder = CType(range.Cells(rCnt, 2), Excel.Range).Value
                vNoorder = vNoorder.Substring(vNoorder.LastIndexOf("/") + 1, vNoorder.Length - vNoorder.LastIndexOf("/") - 1)
                vNama = CType(range.Cells(rCnt, 10), Excel.Range).Value
                vTempJmlBarang = Convert.ToInt32(CType(range.Cells(rCnt, 5), Excel.Range).Value)
                vItem = CType(range.Cells(rCnt, 3), Excel.Range).Value
                vService = CType(range.Cells(rCnt, 13), Excel.Range).Value
                If CType(range.Cells(rCnt, 14), Excel.Range).Value.ToString.Contains("Rp") Then
                    vOngkir = Convert.ToInt32(CType(range.Cells(rCnt, 14), Excel.Range).Value.ToString.Replace("Rp ", "").Replace(".", ""))
                Else
                    vOngkir = Convert.ToInt32(CType(range.Cells(rCnt, 14), Excel.Range).Value)
                End If
                If CType(range.Cells(rCnt, 15), Excel.Range).Value.ToString.Contains("Rp") Then
                    vAsuransi = Convert.ToInt32(CType(range.Cells(rCnt, 15), Excel.Range).Value.ToString.Replace("Rp ", "").Replace(".", ""))
                Else
                    vAsuransi = Convert.ToInt32(CType(range.Cells(rCnt, 15), Excel.Range).Value)
                End If

                If CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString.Contains("Rp") Then
                    vTotal = Convert.ToInt32(CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString.Replace("Rp ", "").Replace(".", "")) * vTempJmlBarang
                    vHargaMasing = CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString.Replace("Rp ", "").Replace(".", "") & " x " & vTempJmlBarang
                Else
                    vTotal = Convert.ToInt32(CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString) * vTempJmlBarang
                    vHargaMasing = CType(range.Cells(rCnt, 7), Excel.Range).Value & " x " & vTempJmlBarang
                End If

                vKeterangan = CType(range.Cells(rCnt, 6), Excel.Range).Value
                vStatus = CType(range.Cells(rCnt, 19), Excel.Range).Value
            Else
                vTempJmlBarang = Convert.ToInt32(CType(range.Cells(rCnt, 5), Excel.Range).Value)
                vItem &= vbNewLine & CType(range.Cells(rCnt, 3), Excel.Range).Value
                If CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString.Contains("Rp") Then
                    vTotal += Convert.ToInt32(CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString.Replace("Rp ", "").Replace(".", "")) * vTempJmlBarang
                    vHargaMasing &= vbNewLine & CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString.Replace("Rp ", "").Replace(".", "") & " x " & vTempJmlBarang
                Else
                    vTotal += Convert.ToInt32(CType(range.Cells(rCnt, 7), Excel.Range).Value.ToString) * vTempJmlBarang
                    vHargaMasing &= vbNewLine & CType(range.Cells(rCnt, 7), Excel.Range).Value & " x " & vTempJmlBarang
                End If
                vKeterangan &= vbNewLine & CType(range.Cells(rCnt, 6), Excel.Range).Value
            End If
            vGTotal = vTotal + vOngkir + vAsuransi
            '### POST IT OR NOT
            'supaya pengecekan masih dalam rentang
            'MsgBox("for DONE")
            'MsgBox("rcnt=" & rCnt & " dari range count:" & range.Rows.Count)
            If rCnt < range.Rows.Count Then
                'MsgBox(IsNothing(CType(range.Cells(rCnt + 1, 1), Excel.Range).Value))
                If Not (IsNothing(CType(range.Cells(rCnt + 1, 1), Excel.Range).Value)) Then
                    'MsgBox("masukCEKPOST")
                    Post = True
                Else
                    Post = False
                End If
            Else
                'ini baris terakhir, pasti POST
                Post = True
            End If
            If Post Then
                Dim row As String() = New String() {vNoorder, vNama, vItem, vService, vKeterangan, vHargaMasing, vTotal, vOngkir, vAsuransi, vGTotal, vStatus}
                'DGV.Rows.Add(row)
                'DGV.Rows.
            End If
        Next
        'RBOB.Checked = True
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = FileUpload1.PostedFile.FileName
    End Sub

    Protected Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        Session.Remove("user")
        Response.Redirect("~/loginForm.aspx")
    End Sub
End Class