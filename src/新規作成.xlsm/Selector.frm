VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Selector 
   Caption         =   "UserForm1"
   ClientHeight    =   2028
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Selector.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Selector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EditButton_Click()
    Unload Me
End Sub

Private Sub ExitButton_Click()
    ThisWorkbook.Close savechanges:=False
    Unload Me
End Sub

Private Sub PathSetButton_Click()
    On Error GoTo catch
    If MsgBox("���݂̕ۑ���͈ȉ��ł��F" & vbCrLf & _
                ThisWorkbook.Worksheets(1).Range("A2").Value & vbCrLf & _
                "�ۑ����ύX���܂����H", vbYesNo + vbQuestion, "�p�X�ύX") = vbNo Then GoTo catch
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.path
        .Title = "�ۑ���t�H���_�I��"
        If .Show = -1 Then
            ThisWorkbook.Worksheets(1).Range("A2").Value = .SelectedItems(1) & "\"
            MsgBox "�ۑ��p�X��ύX���܂���: " & vbCrLf & _
                        ThisWorkbook.Worksheets(1).Range("A2").Value
        End If
    End With
    
    GoTo finally
catch:
    MsgBox "�L�����Z�����܂���"
finally:
End Sub

Private Sub XlsxlButton_Click()
    ExcelFileCreator.run (1)
End Sub

Private Sub XlsmButton_Click()
    ExcelFileCreator.run (2)
End Sub

