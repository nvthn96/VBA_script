Attribute VB_Name = "CommonFunction"
' Chuyen 1 dong nhieu cau thanh moi dong mot cau
Sub SplitSentenceToLines()
Attribute SplitSentenceToLines.VB_ProcData.VB_Invoke_Func = "S\n14"

    Dim str As String
    str = ActiveCell.Value2
    str = Replace(str, "?.", "?")
    str = Replace(str, "!.", "!")

    Static regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
      .Pattern = "[^\.\?\!]+[\.\?\!]?"
      .Global = True
    End With
    Set matches = regex.Execute(str)
    
    Dim count As Integer
    coutn = 0
    
    For Each mch In matches
        If count > 0 Then
            Application.CutCopyMode = False
            Cells(ActiveCell.Row + count, ActiveCell.Column).EntireRow.Insert
        End If
        Cells(ActiveCell.Row + count, ActiveCell.Column).Value2 = Trim(mch)
        count = count + 1
    Next mch
    
End Sub

' Ghep nhieu dong lai thanh dong dau tien
Sub JoinMultipleLines()
Attribute JoinMultipleLines.VB_ProcData.VB_Invoke_Func = "J\n14"

    Dim merge As String
    Dim rng As range
    
    Set rng = Application.Selection
    merge = Trim(rng.Cells(1, 1).Value2)
    
    For i = 2 To rng.Rows.count
        merge = merge & " " & Trim(rng.Cells(i, 1).Value2)
    Next
    
    rng.ClearContents
    rng.Cells(1, 1).Value2 = merge
    
    For i = 2 To rng.Rows.count
        rng.Cells(i, 1).EntireRow.Delete
    Next
    
End Sub


' Chen truoc doan dang chon
Sub InsertRowTop()
Attribute InsertRowTop.VB_ProcData.VB_Invoke_Func = "T\n14"

    Dim rng As range
    Set rng = Application.Selection
    Application.CutCopyMode = False
    Cells(rng.Row, 1).EntireRow.Insert
    
End Sub

' Chen sau doan dang chon
Sub InsertRowDown()
Attribute InsertRowDown.VB_ProcData.VB_Invoke_Func = "D\n14"
    
    Dim rng As range
    Set rng = Application.Selection
    Application.CutCopyMode = False
    Cells(rng.Row + rng.Rows.count, 1).EntireRow.Insert
    
End Sub

' Chen truoc va sau cua nhung dong dang duoc chon
Sub InsertSpaceTopDown()
Attribute InsertSpaceTopDown.VB_ProcData.VB_Invoke_Func = "P\n14"
    
    Call InsertRowDown
    Call InsertRowTop
    
End Sub

' Di chuyen noi dung moi cell xuong nhung dong trong phia duoi
Sub MoveCellsToNextRow()
    
    Dim rng As range
    Set rng = Application.Selection
    
    Dim count As Integer
    count = 0
    Dim cellValue As String
    
    For i = rng.Column To rng.Column + rng.Columns.count
        cellValue = Trim(Cells(rng.Row, i).Value2)
        If cellValue <> "" Then
            Cells(rng.Row, i).Value2 = ""
            count = count + 1
            
            Call InsertSpaceDown
            Cells(rng.Row + count, rng.Column).Value2 = cellValue
            
        End If
    Next
    
End Sub
