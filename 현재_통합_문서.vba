Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call unkeyset
End Sub

Private Sub Workbook_Open()
    ActiveSheet.Protect "Tkdlqjrj", False, True ' Tkdlqjrj is protection password
    Call keyset
End Sub
