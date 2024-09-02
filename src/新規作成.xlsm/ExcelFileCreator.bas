Attribute VB_Name = "ExcelFileCreator"
Option Explicit

'呼び出しはこちらを
Public Sub run(ByVal num As Byte)
    Dim path As String
    path = chkPath(ThisWorkbook.Worksheets(1).Range("A2").Value)
    Select Case num
    Case 1
        CreateXlsxFile (path)
    Case 2
        CreateXlsmFile (path)
    End Select

End Sub

'事前パスチェック
Private Function chkPath(ByVal path As String) As String
    If path = "" Then
        path = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "", , , vbTextCompare) & "一時保存\"
        If Dir(path, vbDirectory) = "" Then MkDir path
    Else
        If Dir(path, vbDirectory) = "" Then MkDir path
    End If
    chkPath = path
End Function

'xlsx作成
Private Sub CreateXlsxFile(ByVal path As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    If Dir(path & "新規ブック.xlsx") = "" Then
        wb.SaveAs Filename:=path & "新規ブック.xlsx"
    Else
        Dim cnt As Integer
        cnt = 1
        While Dir(path & "新規ブック(" & cnt & ").xlsx", vbNormal) <> ""
            cnt = cnt + 1
        Wend
        wb.SaveAs Filename:=path & "新規ブック(" & cnt & ").xlsx"
    End If
    ThisWorkbook.Close savechanges:=False
End Sub

'xlsm作成
Private Sub CreateXlsmFile(ByVal path As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    If Dir(path & "新規ブック.xlsm") = "" Then
        wb.SaveAs Filename:=path & "新規ブック.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        Dim cnt As Integer
        cnt = 1
        While Dir(path & "新規ブック(" & cnt & ").xlsm", vbNormal) <> ""
            cnt = cnt + 1
        Wend
        wb.SaveAs Filename:=path & "新規ブック(" & cnt & ").xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If
    ThisWorkbook.Close savechanges:=False
End Sub
