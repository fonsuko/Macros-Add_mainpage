Sub addEx_Click()

'finding starting row
Dim startingRow As Integer
Dim findRow As Integer
startingRow = 3
findRow = 3
Do While findRow < 2000
    If Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 1).Value <> "" Then //thai language doesn't exist !=[]=!!
        findRow = findRow + 1
        startingRow = findRow
    Else
        findRow = 2000
    End If
Loop


'check province
Dim dd As DropDown
Set dd = ActiveSheet.Shapes("province1").OLEFormat.Object
If dd.List(dd.Value) = "" Or Left(dd.List(dd.Value), 2) = "--" Then
    MsgBox "Error:  â»Ã´àÅ_Í¡¨Ñ§ËÇÑ´ ..."
    GoTo finish
End If


'check exchange
If ActiveSheet.exchange1.Value = "" Or ActiveSheet.exchange1.Value = "ªØÁÊÒÂ" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍªØÁÊÒÂ ..."
    GoTo finish
End If

'check dup
Dim i As Integer
i = 3
Do While i < startingRow
    If dd.List(dd.Value) & Trim(ActiveSheet.exchange1.Value) <> Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(i, 1).Value & Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(i, 2).Value Then
        i = i + 1
    Else
        MsgBox "Error:  ¢éÍÁÙÅ«éÓ ..."
        GoTo finish
    End If

Loop

'check lat
If ActiveSheet.lat1.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ_èÒ Latitude ..."
    GoTo finish
End If

'check long
If ActiveSheet.long1.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ_èÒ Longtitude ..."
    GoTo finish
End If

'check lat long
If ActiveSheet.lat1.Value > ActiveSheet.long1.Value Then
    MsgBox "Error:  â»Ã´µÃÇ¨ÊÍº_èÒ Latitude áÅ_ Longtitude ÍÕ¡_ÃÑé§ ..."
    GoTo finish
End If

'check street
If ActiveSheet.street1.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØ_èÒª_èÍ¶__ ..."
    GoTo finish
End If

'check surveyer
If ActiveSheet.surveyer1.Value = "" Then
    MsgBox "Error:  â»Ã´Ã_ºØª_èÍ_ÙéÊÓÃÇ¨ ..."
    GoTo finish
End If

'add
Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 1).Value = dd.List(dd.Value)
Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 2).Value = Trim(ActiveSheet.exchange1.Value)
Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 3).Value = ActiveSheet.lat1.Value
Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 4).Value = ActiveSheet.long1.Value
If Left(ActiveSheet.street1.Value, 3) = "¶__" Then
    Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 5).Value = Mid(ActiveSheet.street1.Value, 4, Len(ActiveSheet.street1.Value))
Else
    Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 5).Value = ActiveSheet.street1.Value
End If
Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 6).Value = ActiveSheet.surveyer1.Value
Sheets("¢éÍÁÙÅªØÁÊÒÂ").Cells(startingRow, 7).Value = Now

MsgBox "Save successfully ..."

Sheets("2").exchange2.Clear
Sheets("2").exchange2.Value = ""

Sheets("3").exchange4.Clear
Sheets("3").exchange4.Value = ""

Sheets("4").exchange3.Clear
Sheets("4").exchange3.Value = ""

finish:

End Sub
