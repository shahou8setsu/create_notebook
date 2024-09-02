Attribute VB_Name = "ExcelFileCreator"
Option Explicit

'�Ăяo���͂������
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

'���O�p�X�`�F�b�N
Private Function chkPath(ByVal path As String) As String
    If path = "" Then
        path = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "", , , vbTextCompare) & "�ꎞ�ۑ�\"
        If Dir(path, vbDirectory) = "" Then MkDir path
    Else
        If Dir(path, vbDirectory) = "" Then MkDir path
    End If
    chkPath = path
End Function

'xlsx�쐬
Private Sub CreateXlsxFile(ByVal path As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    If Dir(path & "�V�K�u�b�N.xlsx") = "" Then
        wb.SaveAs Filename:=path & "�V�K�u�b�N.xlsx"
    Else
        Dim cnt As Integer
        cnt = 1
        While Dir(path & "�V�K�u�b�N(" & cnt & ").xlsx", vbNormal) <> ""
            cnt = cnt + 1
        Wend
        wb.SaveAs Filename:=path & "�V�K�u�b�N(" & cnt & ").xlsx"
    End If
    ThisWorkbook.Close savechanges:=False
End Sub

'xlsm�쐬
Private Sub CreateXlsmFile(ByVal path As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    If Dir(path & "�V�K�u�b�N.xlsm") = "" Then
        wb.SaveAs Filename:=path & "�V�K�u�b�N.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        Dim cnt As Integer
        cnt = 1
        While Dir(path & "�V�K�u�b�N(" & cnt & ").xlsm", vbNormal) <> ""
            cnt = cnt + 1
        Wend
        wb.SaveAs Filename:=path & "�V�K�u�b�N(" & cnt & ").xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If
    ThisWorkbook.Close savechanges:=False
End Sub
