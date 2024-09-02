VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Selector 
   Caption         =   "UserForm1"
   ClientHeight    =   2028
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Selector.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    If MsgBox("現在の保存先は以下です：" & vbCrLf & _
                ThisWorkbook.Worksheets(1).Range("A2").Value & vbCrLf & _
                "保存先を変更しますか？", vbYesNo + vbQuestion, "パス変更") = vbNo Then GoTo catch
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.path
        .Title = "保存先フォルダ選択"
        If .Show = -1 Then
            ThisWorkbook.Worksheets(1).Range("A2").Value = .SelectedItems(1) & "\"
            MsgBox "保存パスを変更しました: " & vbCrLf & _
                        ThisWorkbook.Worksheets(1).Range("A2").Value
        End If
    End With
    
    GoTo finally
catch:
    MsgBox "キャンセルしました"
finally:
End Sub

Private Sub XlsxlButton_Click()
    ExcelFileCreator.run (1)
End Sub

Private Sub XlsmButton_Click()
    ExcelFileCreator.run (2)
End Sub

