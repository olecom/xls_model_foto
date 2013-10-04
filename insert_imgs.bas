Attribute VB_Name = "insert_imgs"
Option Compare Text
Option Explicit

Sub delete_imgs()
Dim fi As Object
On Error Resume Next
    If vbYes <> MsgBox("Удалить графические объекты на листе?", vbCritical + vbYesNo, "SUPro Фото") _
    Then Exit Sub
Application.ScreenUpdating = False
ThisWorkbook.ch_lock = True

    Cells(1, 255).Select
    For Each fi In ActiveWorkbook.ActiveSheet.Shapes
        If 13 = fi.Type Then fi.Delete ' msoPicture
    Next
    Rows(1).Resize(ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).AutoFit
    Cells(1, 1).Select

ThisWorkbook.ch_lock = False
Application.ScreenUpdating = True
On Error GoTo 0
End Sub

Sub insert_imgs_by_kod_filenames()
Dim r As Range, s As String, m As Match, fi As Object, lastr As Long, d As Double
On Error GoTo err_
    If vbYes <> MsgBox("Вставить фото моделей в активный столбец?" + vbLf + _
                       "Условия: коды моделей в начале текста ячейки" + vbLf + _
                       "         ширина столбца определяет ширину картинок!", _
                        vbQuestion + vbYesNo, "SUPro Фото" _
                ) Then GoTo exit_

    Application.ScreenUpdating = False
    ThisWorkbook.ch_lock = True

    s = ThisWorkbook.Sheets(1).Cells(1, 1).Text
    ChDrive Left(s, 1)
    ChDir s

    lastr = Cells.SpecialCells(xlCellTypeLastCell).Row
    Set r = Cells(1, Selection.Column)

    Do While lastr >= r.Row
        If "" = r.Text Then GoTo next_

        Set m = re_get("^[ ]*([\d]*)", r.Text)
        If Not m Is Nothing Then
            s = m.SubMatches(0)
        Else
            GoTo next_
        End If
        s = Dir(s & ".*")

        If s Like "*png" Or s Like "*jpg" Or s Like "*bmp" Or s Like "*gif" Then
            ActiveWorkbook.ActiveSheet.Pictures.Insert s ' png works too
            Set fi = ActiveWorkbook.ActiveSheet.Pictures(ActiveWorkbook.ActiveSheet.Pictures.Count)
            d = r.Columns(1).Width / fi.Width
            fi.Left = r.Left + 1
            fi.Top = r.Top + 1
            fi.Width = r.Columns(1).Width - 2
            fi.Height = fi.Height * d - 2
            r.Rows(1).RowHeight = fi.Height + 11
            r.Interior.ColorIndex = xlNone
            Set fi = Nothing
        Else
            r.Interior.ColorIndex = 15
        End If
next_:
        Set r = r.Offset(1, 0)
    Loop
exit_loop_:
    Set r = Nothing

    GoTo exit_
err_:
MsgBox "[ошибка] При вставке картинок." & vbLf & _
"Обратитесь к разработчикам за помощью." & vbLf & vbLf & _
"код системной ошибки:" & VBA.CStr(Err.Number) & vbLf & _
"описание: " & Err.Description & " " & Err.Source & vbLf & _
"Конец.", 16 + vbMsgBoxHelpButton, "Системная ошибка", Err.HelpFile, Err.HelpContext
exit_:
    ThisWorkbook.ch_lock = False
    Application.ScreenUpdating = Not False
End Sub

' RegExp (from Tools/References/MS VB RE 5.5)
Public Function rematch_no(patrn, strng, Optional no As Long = 0) As Match
Dim regEx, Matches
  Set regEx = New RegExp
  regEx.Pattern = patrn
  regEx.IgnoreCase = True
  regEx.Global = True
  Set Matches = regEx.Execute(strng)
  If Matches.Count = 0 Then
    Set rematch_no = Nothing
  Else
    If no >= 0 Then
        Set rematch_no = Matches(no)
    Else
        Set rematch_no = Matches
    End If
  End If
  Set regEx = Nothing
End Function

Public Function re_get(pat As String, s As String) As Match
Dim m As Match
Set m = rematch_no(pat, s)
If m Is Nothing Then
    Set m = Nothing
    Exit Function
End If
Set re_get = m 'returns Match object
Set m = Nothing
End Function

Public Function rereplace_all(patrn, gde, na4to, Optional glob As Boolean = True) As String
Dim regEx
  Set regEx = New RegExp
  regEx.Pattern = patrn
  regEx.IgnoreCase = True
  regEx.Global = glob
  rereplace_all = regEx.Replace(gde, na4to)
  Set regEx = Nothing
End Function

