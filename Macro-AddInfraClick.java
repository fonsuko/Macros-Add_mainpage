Sub addInfra_Click()


'finding starting row
Dim startingRow As Integer
Dim findRow As Integer
startingRow = 4
findRow = 3
Do While findRow < 2000
    If Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 3).Value <> "" Then
        findRow = findRow + 1
        startingRow = findRow
    Else
        findRow = 2000
    End If
Loop


'check exchange
If ActiveSheet.exchange2.Value = "" Or ActiveSheet.exchange2.Value = "ªØÁÊÒÂ" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍªØÁÊÒÂ ..."
    GoTo finish
End If
Dim temp As String
temp = Trim(ActiveSheet.exchange2.Value)

'check manhole
If ActiveSheet.manhole2.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍªØÁÊÒÂ manhole ..."
    GoTo finish
End If
temp = temp & Trim(ActiveSheet.manhole2.Value)

'check street
If ActiveSheet.street2.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ_èÒª_èÍ¶__ ..."
    GoTo finish
End If
If Left(ActiveSheet.street2.Value, 3) = "¶__" Then
    temp = temp & Trim(Mid(ActiveSheet.street2.Value, 4, Len(ActiveSheet.street2.Value)))
Else
    temp = temp & Trim(ActiveSheet.street2.Value)
End If


'check dup
Dim i As Integer
i = 3
Do While i < startingRow
    If temp <> Trim(Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(i, 2).Value) & Trim(Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(i, 3).Value) & Trim(Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(i, 8).Value) Then
        i = i + 1
    Else
        MsgBox "Error:  ¢éÍÁÙÅ«éÓ ..."
        GoTo finish
    End If

Loop


'check lat
If ActiveSheet.lat2.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ_èÒ Latitude ..."
    GoTo finish
End If

'check long
If ActiveSheet.long2.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ_èÒ Longtitude ..."
    GoTo finish
End If

'check lat long
If ActiveSheet.lat2.Value > ActiveSheet.long2.Value Then
    MsgBox "Error:  â»Ã´µÃÇ¨ÊÍº_èÒ Latitude áÅ_ Longtitude ÍÕ¡_ÃÑé§ ..."
    GoTo finish
End If

'check manhole or pullbox
Dim dd As DropDown
Set dd = ActiveSheet.Shapes("DropDown2").OLEFormat.Object
If dd.List(dd.Value) = "" Then
    MsgBox "Error:  â»Ã´àÅ_Í¡»Ã_àÀ·ºèÍ_Ñ¡ ..."
    GoTo finish
End If

'check surveyer
If ActiveSheet.surveyer2.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍ_ÙéÊÓÃÇ¨ ..."
    GoTo finish
End If

'check A
If ActiveSheet.checkA.Value = True And (ActiveSheet.rowA.Value = "" Or ActiveSheet.colA.Value = "") Then
    MsgBox "Error:  â»Ã´¡ÃÍ¡¢éÍÁÙÅâ_Ã§ÊÃéÒ§¢Í§ Racking A ãËé_Ãº ..."
    GoTo finish
End If
'check B
If ActiveSheet.checkB.Value = True And (ActiveSheet.rowB.Value = "" Or ActiveSheet.colB.Value = "") Then
    MsgBox "Error:  â»Ã´¡ÃÍ¡¢éÍÁÙÅâ_Ã§ÊÃéÒ§¢Í§ Racking B ãËé_Ãº ..."
    GoTo finish
End If
'check C
If ActiveSheet.checkC.Value = True And (ActiveSheet.rowC.Value = "" Or ActiveSheet.colC.Value = "") Then
    MsgBox "Error:  â»Ã´¡ÃÍ¡¢éÍÁÙÅâ_Ã§ÊÃéÒ§¢Í§ Racking C ãËé_Ãº ..."
    GoTo finish
End If
'check D
If ActiveSheet.checkD.Value = True And (ActiveSheet.rowD.Value = "" Or ActiveSheet.colD.Value = "") Then
    MsgBox "Error:  â»Ã´¡ÃÍ¡¢éÍÁÙÅâ_Ã§ÊÃéÒ§¢Í§ Racking D ãËé_Ãº ..."
    GoTo finish
End If

'check xA
If (ActiveSheet.colxA.Value = "" And ActiveSheet.rowxA.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ col extra ¢Í§ Racking A ..."
    GoTo finish
End If
If (ActiveSheet.rowxA.Value = "" And ActiveSheet.colxA.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ row extra ¢Í§ Racking A ..."
    GoTo finish
End If

'check xB
If (ActiveSheet.colxB.Value = "" And ActiveSheet.rowxb.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ col extra ¢Í§ Racking B ..."
    GoTo finish
End If
If (ActiveSheet.rowxb.Value = "" And ActiveSheet.colxB.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ row extra ¢Í§ Racking B ..."
    GoTo finish
End If

'check xC
If (ActiveSheet.colxC.Value = "" And ActiveSheet.rowxC.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ col extra ¢Í§ Racking C ..."
    GoTo finish
End If
If (ActiveSheet.rowxC.Value = "" And ActiveSheet.colxC.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ row extra ¢Í§ Racking C ..."
    GoTo finish
End If

'check xD
If (ActiveSheet.colxD.Value = "" And ActiveSheet.rowxD.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ col extra ¢Í§ Racking D ..."
    GoTo finish
End If
If (ActiveSheet.rowxD.Value = "" And ActiveSheet.colxD.Value <> "") Then
    MsgBox "Error:  ¡ÃØ_ÒÃ_ºØ_èÒ row extra ¢Í§ Racking D ..."
    GoTo finish
End If

'add info
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 2).Value = Trim(ActiveSheet.exchange2.Value)

Dim rowid As Integer
rowid = 3
Do While (Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(rowid, 2).Value <> ActiveSheet.exchange2.Value) And rowid < 200
    rowid = rowid + 1
Loop
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 1).Value = Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(rowid, 1).Value

Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 3).Value = Trim(ActiveSheet.manhole2.Value)
If Left(ActiveSheet.street2.Value, 3) = "¶__" Then
    Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 8).Value = Mid(ActiveSheet.street2.Value, 4, Len(ActiveSheet.street2.Value))
Else
    Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 8).Value = ActiveSheet.street2.Value
End If
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 4).Value = ActiveSheet.lat2.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 5).Value = ActiveSheet.long2.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 6).Value = dd.List(dd.Value)
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 7).Value = ActiveSheet.type2.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 25).Value = ActiveSheet.surveyer2.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 26).Value = Now

If ActiveSheet.checkA.Value = True Then
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 9).Value = ActiveSheet.rowA.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 10).Value = ActiveSheet.colA.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 11).Value = ActiveSheet.rowxA.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 12).Value = ActiveSheet.colxA.Value
End If

If ActiveSheet.checkB.Value = True Then
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 13).Value = ActiveSheet.rowB.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 14).Value = ActiveSheet.colB.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 15).Value = ActiveSheet.rowxb.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 16).Value = ActiveSheet.colxB.Value
End If

If ActiveSheet.checkC.Value = True Then
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 17).Value = ActiveSheet.rowC.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 18).Value = ActiveSheet.colC.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 19).Value = ActiveSheet.rowxC.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 20).Value = ActiveSheet.colxC.Value
End If

If ActiveSheet.checkD.Value = True Then
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 21).Value = ActiveSheet.rowD.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 22).Value = ActiveSheet.colD.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 23).Value = ActiveSheet.rowxD.Value
Sheets("¢éÍÁÙÅâ_Ã§ÊÃéÒ§·èÍ").Cells(startingRow, 24).Value = ActiveSheet.colxD.Value
End If

MsgBox "Save successfully ..."

finish:

End Sub
