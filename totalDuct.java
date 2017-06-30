Public totalduct As Integer //อันนี้เป็นหน้าที่เชื่อมกับหน้าสุดท้าย

Sub getInfra_Click() //link ข้อมูลจากหน้าที่เคยกรอกก่อนหน้านี้
totalduct = 0

'check exchange
If ActiveSheet.exchange3.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍªØÁÊÒÂ ..."
    GoTo finish
End If

'check manhole
If ActiveSheet.manhole3.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ¢éÍÁÙÅºèÍ_Ñ¡ ..."
    GoTo finish
End If

'check road
If ActiveSheet.road3.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍ¶__ ..."
    GoTo finish
End If

'check racking
If ActiveSheet.racking3.Value = "" Then
    MsgBox "Error:  â»Ã´àÅ_Í¡ racking ¢Í§ºèÍ_Ñ¡ ..."
    GoTo finish
End If

'search row
Dim gotorow As Integer
rowid = 4
Do While (Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 3).Value <> "") Or rowid < 10000
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 2).Value & Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 3).Value & Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(rowid, 8).Value = ActiveSheet.exchange3.Value & ActiveSheet.manhole3.Value & ActiveSheet.road3.Value Then
       gotorow = rowid
       Exit Do
    End If
    rowid = rowid + 1
Loop

'clean layout
Dim rackingrow, rackingcol, num As Byte
For rackingrow = 1 To 16
    For rackingcol = 11 To 26
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 2
        ActiveSheet.Cells(rackingrow, rackingcol).Value = ""
    Next rackingcol
Next rackingrow

ActiveSheet.ductno3.Clear
ActiveSheet.ductno3.Value = ""
ActiveSheet.ductno3.Enabled = True

Dim resultductusage As Byte
Dim rowtemp, coltemp As Byte
If ActiveSheet.racking3.Value = "A" Then
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value To 14 Step -1
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 46
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol

    For rackingcol = 13 To 14 - Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 11).Value Step -1
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 46
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol
    totalduct = num - 1


ElseIf ActiveSheet.racking3.Value = "B" Then
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 14).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 14).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingrow = 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value To 4 Step -1
    For rackingcol = 14 To 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 14).Value
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 45
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow

    For rackingrow = 3 To 4 - Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value Step -1
    For rackingcol = 14 To 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 45
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow
    totalduct = num - 1


ElseIf ActiveSheet.racking3.Value = "C" Then
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 17).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 17).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingcol = 14 To 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 17).Value
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 33
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol

    coltemp = rackingcol

    For rackingcol = coltemp To coltemp - 1 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 20).Value
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 33
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol
    totalduct = num - 1


Else 'D
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 21).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 21).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 22).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 22).Value = "" Then
        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
        GoTo finish
    End If
    num = 1
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 21).Value
    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 22).Value To 14 Step -1
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 33
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow

    rowtemp = rackingrow

    For rackingrow = rowtemp To rowtemp - 1 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 23).Value
    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 24).Value To 14 Step -1
        resultductusage = checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num)
        If resultductusage = 3 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
        ElseIf resultductusage = 1 Then
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        Else
            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 33
        End If
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow
    totalduct = num - 1


End If





'If ActiveSheet.racking3.Value = "A" Then
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    num = 1
'    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value To 14 Step -1
'    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value
'        If checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
'        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num) = True Then
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
'        Else
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
'        End If
'        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
'        num = num + 1
'    Next rackingrow
'    Next rackingcol
'    totalduct = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9) * Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10)
'
'
'ElseIf ActiveSheet.racking3.Value = "B" Then
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    num = 1
'    For rackingrow = 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value To 4 Step -1
'    For rackingcol = 14 To 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13).Value
'        If checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
'        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num) = True Then
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
'        Else
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
'        End If
'        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
'        num = num + 1
'    Next rackingcol
'    Next rackingrow
'    totalduct = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12) * Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 13)
'
'
'ElseIf ActiveSheet.racking3.Value = "C" Then
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    num = 1
'    For rackingcol = 14 To 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value
'    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value
'        If checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
'        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num) = True Then
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
'        Else
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
'        End If
'        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
'        num = num + 1
'    Next rackingrow
'    Next rackingcol
'    totalduct = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15) * Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16)
'Else
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value > 10 Or Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value = "" Then
'        MsgBox "Error:  äÁèÁÕ¢éÍÁÙÅ ËÃ_Í ¢_Ò´ racking ãË_èà¡Ô_¡ÇèÒ·Õè¡ÓË_´ (i.e., 10x10) ..."
'        GoTo finish
'    End If
'    num = 1
'    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value
'    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18).Value To 14 Step -1
'        If checkductusage(ActiveSheet.exchange3.Value, ActiveSheet.manhole3.Value, _
'        ActiveSheet.road3.Value, ActiveSheet.racking3.Value, num) = True Then
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 50
'        Else
'            ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
'        End If
'        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
'        num = num + 1
'    Next rackingcol
'    Next rackingrow
'    totalduct = Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 18) * Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19)
'
'
'End If

Sheets("4").ductno3.Clear
Sheets("4").ductno3.Value = ""
Sheets("4").ductno3.Enabled = True
Sheets("4").save3.Enabled = False

finish:
End Sub

Private Function checkductusage(exchange As String, manhole As String, road As String, racking As String, ductno As Byte) As Byte

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
       Exit Do
    End If
    rowid = rowid + 1
Loop

If gotorow <> 0 Then
    If Sheets("¢éÍÁÙÅ¡ÒÃãªé·èÍ").Cells(gotorow, 11).Value <> "" Then
        checkductusage = 2 ' broken
    Else
        checkductusage = 3 ' used
    End If
Else
    checkductusage = 1 'unused
End If


End Function
