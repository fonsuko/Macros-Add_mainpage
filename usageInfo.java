Public usageinfo As String

Sub getinfo_Click()

'check exchange
If ActiveSheet.exchange4.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍªØÁÊÒÂ ..."
    GoTo finish
End If

'check manhole
If ActiveSheet.manhole4.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ¢éÍÁÙÅºèÍ_Ñ¡ ..."
    GoTo finish
End If

'check road
If ActiveSheet.manhole4.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍ¶__ ..."
    GoTo finish
End If

'check racking
If ActiveSheet.racking4.Value = "" Then
    MsgBox "Error:  â»Ã´àÅ_Í¡ racking ¢Í§ºèÍ_Ñ¡ ..."
    GoTo finish
End If

'search row
Dim gotorow As Integer
rowid = 4
Do While (Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 3).Value <> "") Or rowid < 10000
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 2).Value & Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 3).Value & Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 8).Value = ActiveSheet.exchange4.Value & ActiveSheet.manhole4.Value & ActiveSheet.road4.Value Then
       gotorow = rowid
       Exit Do
    End If
    rowid = rowid + 1
Loop

'get info
ActiveSheet.type4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 7).Value
ActiveSheet.latt4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 4).Value
ActiveSheet.long4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 5).Value
If ActiveSheet.racking4 = "A" Then
    ActiveSheet.connect4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 11).Value
ElseIf ActiveSheet.racking4 = "B" Then
    ActiveSheet.connect4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 14).Value
ElseIf ActiveSheet.racking4 = "C" Then
    ActiveSheet.connect4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 17).Value
ElseIf ActiveSheet.racking4 = "D" Then
    ActiveSheet.connect4.Value = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 20).Value
Else
ActiveSheet.connect4.Value = "(undefined)"
End If

usageinfo = ""

'clean layout
Dim rackingrow, rackingcol, num As Byte
For rackingrow = 5 To 14
    For rackingcol = 4 To 13
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 2
        ActiveSheet.Cells(rackingrow, rackingcol).Value = ""
    Next rackingcol
Next rackingrow

If ActiveSheet.racking4.Value = "A" Then
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingcol = 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value To 4 Step -1
    For rackingrow = 5 To 4 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value
        If checkductusage(ActiveSheet.exchange4.Value, ActiveSheet.manhole4.Value, _
        ActiveSheet.road4.Value, ActiveSheet.racking4.Value, num) = True Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol


ElseIf ActiveSheet.racking4.Value = "B" Then
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingrow = 4 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value To 5 Step -1
    For rackingcol = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value
        If checkductusage(ActiveSheet.exchange4.Value, ActiveSheet.manhole4.Value, _
        ActiveSheet.road4.Value, ActiveSheet.racking4.Value, num) = True Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow


ElseIf ActiveSheet.racking4.Value = "C" Then
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingcol = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value
    For rackingrow = 5 To 4 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value
        If checkductusage(ActiveSheet.exchange4.Value, ActiveSheet.manhole4.Value, _
        ActiveSheet.road4.Value, ActiveSheet.racking4.Value, num) = True Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol

Else
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingrow = 5 To 4 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value
    For rackingcol = 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value To 4 Step -1
        If checkductusage(ActiveSheet.exchange4.Value, ActiveSheet.manhole4.Value, _
        ActiveSheet.road4.Value, ActiveSheet.racking4.Value, num) = True Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow

End If

ActiveSheet.usage4.Value = usageinfo

finish:
End Sub

Private Function checkductusage(exchange As String, manhole As String, road As String, racking As String, ductno As Byte) As Boolean

Dim gotorow As Integer
gotorow = 0
rowid = 3
Do While (Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 3).Value <> "") And rowid < 10000
    If Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 2).Value _
    & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 3).Value _
    & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 4).Value _
    & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 5).Value _
    & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 6).Value _
    = exchange _
    & manhole _
    & road _
    & racking _
    & CStr(ductno) Then
       gotorow = rowid
       usageinfo = usageinfo + "Duct no. " & ductno & vbNewLine _
       & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 7).Value & " sub-duct | " & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 8).Value & vbNewLine _
       & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 9).Value & " micro-duct | " & Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(rowid, 10).Value & vbNewLine & vbNewLine
       Exit Do
    End If
    rowid = rowid + 1
Loop

If gotorow <> 0 Then
    checkductusage = True
Else
    checkductusage = False
End If


End Function
