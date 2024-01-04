Private Sub Worksheet_Change(ByVal Target As Range)

Dim c1 As Range

ActiveSheet.Unprotect Password:="0012645234"

For Each c1 In Target
    If c1.Value <> "" Then
        c1.Locked = True
        ActiveSheet.Protect Password:="0012645234"
    End If
    
Next c1
End Sub
