Attribute VB_Name = "ProjectWorkbookEvents"

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    MsgBox "workbook newsheet 1"
End Sub

Private Sub Workbook_Open()
    MsgBox "workbook open 1"
End Sub

Private Sub Workbook_Close()
    MsgBox "workbook close 1"
End Sub
