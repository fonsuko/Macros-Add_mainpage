
//get duct number by adding rank of duct
Sub getductno_Click()
totalduct = 0

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
If ActiveSheet.road4.Value = "" Then
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

'clean layout
Dim rackingrow, rackingcol, num As Byte
For rackingrow = 1 To 16
    For rackingcol = 11 To 26
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 2
        ActiveSheet.Cells(rackingrow, rackingcol).Value = ""
    Next rackingcol
Next rackingrow

Dim rowtemp, coltemp As Byte //this is extra of duct ranking that happens sometime.
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
    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 9).Value To 14 Step -1
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 10).Value
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol

    For rackingcol = 13 To 14 - Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 11).Value Step -1
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 12).Value
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol


ElseIf ActiveSheet.racking4.Value = "B" Then
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
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow

    For rackingrow = 3 To 4 - Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 15).Value Step -1
    For rackingcol = 14 To 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 16).Value
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow


ElseIf ActiveSheet.racking4.Value = "C" Then
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
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol

    coltemp = rackingcol

    For rackingcol = coltemp To coltemp - 1 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 19).Value
    For rackingrow = 4 To 3 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 20).Value
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingrow
    Next rackingcol


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
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow

    rowtemp = rackingrow

    For rackingrow = rowtemp To rowtemp - 1 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 23).Value
    For rackingcol = 13 + Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(gotorow, 24).Value To 14 Step -1
        ActiveSheet.Cells(rackingrow, rackingcol).Interior.ColorIndex = 48
        ActiveSheet.Cells(rackingrow, rackingcol).Value = num
        num = num + 1
    Next rackingcol
    Next rackingrow


End If

Sheets("3").save4.Enabled = True

finish:
End Sub
