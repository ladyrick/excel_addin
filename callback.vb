Public Sub ladyrick_callback(Control As IRibbonControl)
    Select Case Control.ID
        Case "button1"
            call split_file
        Case "button2"
            msgbox "button2"
        Case "button3"
            msgbox "button3"
        Case "button4"
            msgbox "button4"
    End Select
End Sub

Sub split_file()
    Application.ScreenUpdating = False
    Dim clm_d, hh As Integer
    Dim mycell As Range
    Dim nodupes As New Collection
    Dim rngop As Range
    Set shtop = ActiveSheet
    hh = Application.CountA(Range("1:110"))
    clm_d = Application.InputBox(prompt:="请选择作为拆分的列" & Chr(13) & "注意:" & Chr(13) & "1、第一行必须为标题行" & Chr(13) & "2、输处列号（如1或2），用键盘输入", Type:=1)
    If clm_d = False Or clm_d > hh Then Exit Sub
    On Error Resume Next
    For Each mycell In shtop.Range(Cells(4, clm_d), (shtop.Cells(4, clm_d).End(xlDown)))
        nodupes.Add mycell.Value, CStr(mycell.Value)
    Next mycell
    On Error GoTo 0
    Set rngop = Cells.CurrentRegion
    Dim filename As String
    filename = Application.ActiveWorkbook.FullName
    Dim newworkbook As Workbook
    For Each Item In nodupes
        Set newworkbook = Application.Workbooks.Add
        On Error Resume Next
        newworkbook.SaveAs filename:=filename & "." & Item & ".xlsx"
        On Error Resume Next
        Sheets("sheet1").Name = Item
        rngop.AutoFilter Field:=clm_d, Criteria1:=Item
        rngop.Copy
        Sheets(Item).Paste
        newworkbook.Save
        newworkbook.Close
    Next Item
    shtop.AutoFilterMode = False
    shtop.Activate
    Application.ScreenUpdating = True
End Sub
