Attribute VB_Name = "ProjectAutoEvents"

Private Sub Auto_Open()
    'MsgBox "auto open 1"
End Sub

Private Sub Auto_Close()

    Application.ScreenUpdating = False
    
    Call FormatSheet
    
    Application.ScreenUpdating = True
    
End Sub

' Zoom level = 100
' Go to cell A1
Private Sub FormatSheet()

    Dim book As Workbook
    Set book = activeWorkbook
    
    If book Is Nothing Then
        Exit Sub
    End If
    
    Dim savedSheet As Worksheet
    Set savedSheet = book.ActiveSheet
    
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        sheet.Activate
        ActiveWindow.Zoom = 100
        ActiveWindow.ScrollRow = 1
		ActiveWindow.ScrollColumn = 1
        sheet.range("A1").Select
    Next sheet
    
    savedSheet.Activate
    
    Call book.Save
    
End Sub

