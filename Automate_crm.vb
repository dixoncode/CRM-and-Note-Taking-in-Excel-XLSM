Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogSheet") ' Ensure this sheet exists

    ' Check if the change was made in Column M
    If Not Intersect(Target, Me.Range("M:M")) Is Nothing Then
        Application.EnableEvents = False ' Prevent infinite loops
        On Error GoTo ErrorHandler ' Error handling

        ' Insert a new row at Row 2 (pushing old entries down)
        ws.Rows(2).Insert Shift:=xlDown

        ' Log details in LogSheet
        ws.Cells(2, 1).Value = Now ' Timestamp (Column A)
        ws.Cells(2, 2).Value = Me.Cells(Target.Row, 2).Value ' Client (Column B)
        ws.Cells(2, 3).Value = Target.Value ' Note (Column C)

    End If

ExitHandler:
    Application.EnableEvents = True ' Re-enable events
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
