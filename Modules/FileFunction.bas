Attribute VB_Name = "FileFunction"
' Luu danh sach sheet thanh file text trong folder chi dinh
Sub SaveAllTextToFolder()
    
    Application.ScreenUpdating = False

    Dim objFileDialog As Office.FileDialog
    Set objFileDialog = Application.FileDialog(MsoFileDialogType.msoFileDialogFolderPicker)
    
    With objFileDialog
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .Title = "Folder Picker"
        If (.Show > 0) Then
        End If
        If (.SelectedItems.count > 0) Then
            Dim sFolder As String
            sFolder = .SelectedItems(1)
            
            Dim book As Workbook
            Set book = activeWorkbook
            
            If book Is Nothing Then
                Exit Sub
            End If
            
            Dim savedSheet As Worksheet
            Set savedSheet = book.ActiveSheet
            
            Dim txt As String
            Dim txtAll As String
            Dim usedRng As range
            
            Dim sht As Worksheet
            For Each sht In book.Worksheets
                If IsSkipSheet(sht) = False Then
                    sht.Activate
                    Set usedRng = sht.UsedRange
                    ' text cua 1 worksheet
                    txt = ""
                    For Each cel In usedRng.Cells
                        txt = txt & cel.Value2 & vbCrLf
                    Next
                    
                    Dim sFilePath As String
                    sFilePath = sFolder & "\" & sht.Name & ".txt"
                    Call WriteToFile(sFilePath, "utf-8", txt)
                    
                    ' text cua tat ca worksheet
                    If Len(txtAll) <> 0 Then
                        txtAll = txtAll & vbCrLf
                    End If
                    txtAll = txtAll & txt
                    
                End If
            Next
            
            Dim sFileAll As String
            sFileAll = sFolder & "\" & "All Text.txt"
            Call WriteToFile(sFileAll, "utf-8", txtAll)
            
            savedSheet.Activate
            
        End If
    End With
    
    MsgBox "Exported!"
    
    Application.ScreenUpdating = True
    
End Sub

' kiem tra sheet co duoc bo qua khi export ra file khong
Function IsSkipSheet(sht As Worksheet)
    If InStr(LCase(sht.Name), LCase("#skip")) <> 0 Then
        IsSkipSheet = True
    Else
        IsSkipSheet = False
    End If
End Function

' ghi text ra file
Sub WriteToFile(sPath As String, sCharset As String, sData As String)

    Static adoStream As Object
    Set adoStream = CreateObject("ADODB.Stream")
    'Unicode coding
    'adoStream.Charset = "Unicode"
    'adoStream.Charset = "utf-8"
    adoStream.Charset = sCharset
    'open sream
    adoStream.Open
    'write a text
    adoStream.WriteText sData
    'save to file
    adoStream.SaveToFile sPath, 2
    adoStream.Close
    
End Sub

' Remove tat ca cell ma chi co space
Sub RemoveAllSpaceOnly()
    
    Application.ScreenUpdating = False

    Dim book As Workbook
    Set book = activeWorkbook
    
    If book Is Nothing Then
        Exit Sub
    End If
    
    Dim savedSheet As Worksheet
    Set savedSheet = book.ActiveSheet
    
    Dim usedRng As range
    
    Dim sht As Worksheet
    For Each sht In book.Worksheets
        If IsSkipSheet(sht) = False Then
            sht.Activate
            Set usedRng = sht.UsedRange
            For Each cel In usedRng.Cells
                If Len(cel.Value2) <> 0 Then
                    If Len(Trim(cel.Value2)) = 0 Then
                        cel.Value2 = vbNullString
                    End If
                End If
            Next
            
        End If
    Next
    
    savedSheet.Activate
    
    Application.ScreenUpdating = True
    
End Sub
