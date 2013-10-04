VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Public ch_lock As Boolean

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal t As Range)
Dim j
    If ch_lock Then Exit Sub
    ch_lock = True
    ' The first cell in the first sheet is a directory with images
    If "foto_setup" = Sh.CodeName And 1 = t.Row And 1 = t.Column And _
       "" = Trim(t.Text) _
    Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = -1 Then
            For Each j In .SelectedItems
                t.Value2 = j & "\"
                Exit For
            Next
            End If
        End With
    End If
    ch_lock = Not True
End Sub
